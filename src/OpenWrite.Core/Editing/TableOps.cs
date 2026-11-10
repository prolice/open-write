using System.Globalization;
using System.Xml.Linq;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;

namespace OpenWrite.Core.Editing;

/// <summary>
/// Opérations sur les tableaux (lot 6), écrites comme Word les écrit : création (« Grille du tableau », w:tblLook 04A0,
/// largeurs en twips), rangées et colonnes, fusion (w:gridSpan, w:vMerge) et fractionnement des cellules, largeurs et
/// hauteurs, mise en forme des cellules (trame, bordures, alignement), style et options de style, ajustement
/// automatique, conversion texte ↔ tableau, tri, formule. Les opérations modifient le tableau en place ; l'appelant les
/// place dans une transaction du journal sur le bloc du tableau.
/// </summary>
public static class TableOps
{
    // ================================================================ carte de la grille

    /// <summary>Cellule d'une rangée : colonne de la grille où elle commence et nombre de colonnes couvertes.</summary>
    public readonly record struct Slot(TableCell Cell, int Col, int Span);

    /// <summary>Tableau vu comme une grille : rangées, cellules et leurs colonnes (w:gridBefore et w:gridSpan compris).</summary>
    public sealed class Map
    {
        public Map(Table table)
        {
            Table = table;
            W = table.Props.Ns;
            Rows = table.Items.OfType<TableRow>().ToList();
            foreach (var row in Rows)
            {
                var slots = new List<Slot>();
                int col = Int(row.Props.Get("gridBefore"), W) ?? 0;
                foreach (var cell in row.Items.OfType<TableCell>())
                {
                    int span = Math.Max(1, Int(cell.Props.Get("gridSpan"), W) ?? 1);
                    slots.Add(new Slot(cell, col, span));
                    col += span;
                }
                Cells.Add(slots);
                Columns = Math.Max(Columns, col + (Int(row.Props.Get("gridAfter"), W) ?? 0));
            }
            Columns = Math.Max(Columns, table.Grid?.Elements(W + "gridCol").Count() ?? 0);
        }

        public Table Table { get; }
        public XNamespace W { get; }
        public List<TableRow> Rows { get; }
        public List<List<Slot>> Cells { get; } = new();
        public int Columns { get; }

        /// <summary>Cellule qui couvre la colonne <paramref name="col"/> de la rangée (null : hors des cellules).</summary>
        public Slot? At(int row, int col) =>
            row < 0 || row >= Cells.Count ? null : Cells[row].Cast<Slot?>().FirstOrDefault(s => s!.Value.Col <= col && col < s.Value.Col + s.Value.Span);

        /// <summary>Rangée et colonne de départ d'une cellule.</summary>
        public (int Row, int Col)? Find(TableCell cell)
        {
            for (int r = 0; r < Cells.Count; r++)
                foreach (var s in Cells[r])
                    if (ReferenceEquals(s.Cell, cell)) return (r, s.Col);
            return null;
        }

        public bool IsContinue(TableCell c) => c.Props.Get("vMerge") is { } v && v.Attribute(W + "val")?.Value is null or "continue";
        public bool IsRestart(TableCell c) => c.Props.Get("vMerge")?.Attribute(W + "val")?.Value == "restart";

        /// <summary>Première rangée d'une fusion verticale qui passe par (row, col).</summary>
        public int MergeTop(int row, int col)
        {
            while (row > 0 && At(row, col) is { } s && IsContinue(s.Cell) && At(row - 1, col) is { } up && up.Col == s.Col) row--;
            return row;
        }

        /// <summary>Dernière rangée d'une fusion verticale qui passe par (row, col).</summary>
        public int MergeBottom(int row, int col)
        {
            var s0 = At(row, col);
            if (s0 is null) return row;
            while (row + 1 < Cells.Count && At(row + 1, col) is { } down && down.Col == s0.Value.Col && IsContinue(down.Cell)) row++;
            return row;
        }

        /// <summary>Largeurs de la grille (twips), complétées si w:tblGrid en a moins que de colonnes.</summary>
        public List<int> GridWidths()
        {
            var g = Table.Grid?.Elements(W + "gridCol").Select(e => Int(e, W, "w") ?? 0).ToList() ?? new List<int>();
            int fallback = g.Count > 0 && g.Any(x => x > 0) ? (int)g.Where(x => x > 0).Average() : 1440;
            while (g.Count < Columns) g.Add(fallback);
            for (int i = 0; i < g.Count; i++) if (g[i] <= 0) g[i] = fallback;
            return g;
        }

        /// <summary>
        /// Rectangle de cellules agrandi pour couvrir entièrement les cellules fusionnées qu'il touche (Word sélectionne
        /// toujours des cellules entières).
        /// </summary>
        public (int R0, int C0, int R1, int C1) Normalize(int r0, int c0, int r1, int c1)
        {
            if (r0 > r1) (r0, r1) = (r1, r0);
            if (c0 > c1) (c0, c1) = (c1, c0);
            for (bool changed = true; changed;)
            {
                changed = false;
                for (int r = r0; r <= r1; r++)
                    foreach (var s in Cells[r])
                    {
                        if (s.Col + s.Span - 1 < c0 || s.Col > c1) continue;
                        if (s.Col < c0) { c0 = s.Col; changed = true; }
                        if (s.Col + s.Span - 1 > c1) { c1 = s.Col + s.Span - 1; changed = true; }
                        int top = MergeTop(r, s.Col), bottom = MergeBottom(r, s.Col);
                        if (top < r0) { r0 = top; changed = true; }
                        if (bottom > r1) { r1 = bottom; changed = true; }
                    }
            }
            return (r0, c0, r1, c1);
        }

        /// <summary>Cellules (premières de leur fusion verticale) du rectangle, rangée par rangée.</summary>
        public IEnumerable<(int Row, Slot Slot)> In(int r0, int c0, int r1, int c1)
        {
            for (int r = r0; r <= r1 && r < Cells.Count; r++)
                foreach (var s in Cells[r])
                    if (s.Col + s.Span - 1 >= c0 && s.Col <= c1) yield return (r, s);
        }
    }

    static int? Int(XElement? e, XNamespace w, string attr = "val") =>
        int.TryParse(e?.Attribute(w + attr)?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : null;

    static string Tw(double v) => ((int)Math.Round(v)).ToString(CultureInfo.InvariantCulture);

    // ================================================================ création

    /// <summary>Paragraphe vide qui reprend les propriétés de <paramref name="like"/> (sans saut de section ni numérotation).</summary>
    public static Paragraph EmptyParagraph(XNamespace w, Paragraph? like, Func<string> paraId)
    {
        var p = new Paragraph { Props = like?.Props.Clone() ?? new PropertyBag("pPr", w) };
        p.Props.Remove("sectPr");
        p.Props.Remove("numPr");
        if (like is not null)
            foreach (var a in like.Attributes)
                p.Attributes.Add(a.Name == FlowEdits.W14 + "paraId" || a.Name == FlowEdits.W14 + "textId" ? new XAttribute(a.Name, paraId()) : new XAttribute(a));
        return p;
    }

    static TableCell NewCell(XNamespace w, PropertyBag props, Paragraph? like, Func<string> paraId)
    {
        var c = new TableCell { Props = props };
        c.Content.Blocks.Add(EmptyParagraph(w, like, paraId));
        return c;
    }

    static XElement Width(XNamespace w, string name, int twips, string type = "dxa") =>
        new(w + name, new XAttribute(w + "w", twips), new XAttribute(w + "type", type));

    /// <summary>w:tblLook écrit comme Word (valeur hexadécimale et attributs nommés).</summary>
    public static XElement LookElement(XNamespace w, TableLook look)
    {
        int val = (look.FirstRow ? 0x20 : 0) | (look.LastRow ? 0x40 : 0) | (look.FirstColumn ? 0x80 : 0) | (look.LastColumn ? 0x100 : 0)
                | (look.NoHBand ? 0x200 : 0) | (look.NoVBand ? 0x400 : 0);
        static string B(bool b) => b ? "1" : "0";
        return new XElement(w + "tblLook", new XAttribute(w + "val", val.ToString("X4", CultureInfo.InvariantCulture)),
            new XAttribute(w + "firstRow", B(look.FirstRow)), new XAttribute(w + "lastRow", B(look.LastRow)),
            new XAttribute(w + "firstColumn", B(look.FirstColumn)), new XAttribute(w + "lastColumn", B(look.LastColumn)),
            new XAttribute(w + "noHBand", B(look.NoHBand)), new XAttribute(w + "noVBand", B(look.NoVBand)));
    }

    /// <summary>
    /// Nouveau tableau comme « Insertion ▸ Tableau » de Word : largeur du texte partagée en colonnes égales (Word retire
    /// 10 twips au total : 3 × 3005 + 1 pour 9026), w:tblW automatique, w:tblLook 04A0 ; <paramref name="autoFit"/> :
    /// largeurs automatiques (ajuster au contenu).
    /// </summary>
    public static Table Create(XNamespace w, int rows, int cols, double textWidthPt, string? styleId, Paragraph? like, Func<string> paraId, bool autoFit = false)
    {
        rows = Math.Clamp(rows, 1, 500);
        cols = Math.Clamp(cols, 1, 63);
        int total = Math.Max(cols * 200, (int)Math.Round(textWidthPt * 20) - 10);
        var widths = Enumerable.Repeat(total / cols, cols).ToArray();
        widths[^1] += total - widths.Sum();
        var t = new Table { Props = new PropertyBag("tblPr", w) };
        if (styleId is not null) t.Props.SetVal("tblStyle", styleId);
        t.Props.Set(Width(w, "tblW", 0, "auto"));
        t.Props.Set(LookElement(w, TableLook.Default));
        t.Grid = new XElement(w + "tblGrid", widths.Select(x => new XElement(w + "gridCol", new XAttribute(w + "w", x))));
        for (int r = 0; r < rows; r++)
        {
            var row = new TableRow();
            for (int c = 0; c < cols; c++)
            {
                var props = new PropertyBag("tcPr", w);
                props.Set(autoFit ? Width(w, "tcW", 0, "auto") : Width(w, "tcW", widths[c]));
                row.Items.Add(NewCell(w, props, like, paraId));
            }
            t.Items.Add(row);
        }
        return t;
    }

    /// <summary>
    /// Insère un tableau au point d'insertion, comme Word : en tête de paragraphe, avant lui ; ailleurs, le paragraphe est
    /// coupé et le tableau placé entre les deux moitiés (un paragraphe suit toujours le tableau). Renvoie le début de la
    /// première cellule. Transaction : bloc <c>pos.BlockIndex</c> du conteneur.
    /// </summary>
    public static DocPos InsertTable(Story root, DocPos pos, Table table, Func<string> paraId)
    {
        var container = DocPos.ResolveContainer(root, pos.ContainerPath);
        int at = pos.BlockIndex;
        if (pos.Offset > 0)
        {
            TextOps.Split(root, pos, paraId);
            at++;
        }
        container.Blocks.Insert(at, table);
        int rowItem = table.Items.FindIndex(i => i is TableRow);
        int cellItem = ((TableRow)table.Items[rowItem]).Items.FindIndex(i => i is TableCell);
        return new DocPos(pos.ContainerPath.Concat(new[] { at, rowItem, cellItem, 0 }).ToArray(), 0);
    }

    // ================================================================ rangées

    /// <summary>
    /// Insère <paramref name="count"/> rangées au-dessus ou en dessous de la rangée <paramref name="row"/>, sur son modèle
    /// (cellules et largeurs, mise en forme du premier paragraphe de chaque cellule, sans le texte) ; une fusion verticale
    /// qui traverse le point d'insertion se prolonge. Renvoie le rang de la première nouvelle rangée.
    /// </summary>
    public static int InsertRows(Table t, int row, bool below, int count, Func<string> paraId)
    {
        var map = new Map(t);
        var w = map.W;
        var src = map.Rows[row];
        var next = below ? (row + 1 < map.Rows.Count ? map.Cells[row + 1] : null) : null;
        int insertAt = t.Items.IndexOf(src) + (below ? 1 : 0);
        for (int k = 0; k < count; k++)
        {
            var nr = new TableRow { Props = src.Props.Clone(), PropsException = src.PropsException?.Clone() };
            foreach (var a in src.Attributes)
                nr.Attributes.Add(a.Name == FlowEdits.W14 + "paraId" || a.Name == FlowEdits.W14 + "textId" ? new XAttribute(a.Name, paraId()) : new XAttribute(a));
            nr.Props.Remove("cnfStyle");
            if (below) nr.Props.Remove("tblHeader");
            foreach (var item in src.Items)
            {
                if (item is not TableCell c) continue;
                var props = c.Props.Clone();
                props.Remove("cnfStyle");
                bool merged = map.IsRestart(c) || map.IsContinue(c);
                var col = map.Find(c)!.Value.Col;
                bool keep = merged && (below
                    ? next?.FirstOrDefault(s => s.Col == col) is { } n && map.IsContinue(n.Cell)
                    : map.IsContinue(c));
                if (keep) props.Set(new XElement(w + "vMerge"));
                else props.Remove("vMerge");
                nr.Items.Add(NewCell(w, props, c.Content.Paragraphs().FirstOrDefault(), paraId));
            }
            t.Items.Insert(insertAt + k, nr);
        }
        return row + (below ? 1 : 0);
    }

    /// <summary>
    /// Supprime les rangées <paramref name="from"/>..<paramref name="to"/> ; une fusion verticale qui commençait dans
    /// une rangée supprimée recommence à la première rangée gardée. Vrai si le tableau n'a plus de rangée.
    /// </summary>
    public static bool DeleteRows(Table t, int from, int to)
    {
        var map = new Map(t);
        if (from > to) (from, to) = (to, from);
        if (to + 1 < map.Rows.Count)
            foreach (var s in map.Cells[to + 1])
                if (map.IsContinue(s.Cell) && map.MergeTop(to + 1, s.Col) is var top && top >= from && top <= to)
                    s.Cell.Props.Set(new XElement(map.W + "vMerge", new XAttribute(map.W + "val", "restart")));
        for (int r = from; r <= to && r < map.Rows.Count; r++) t.Items.Remove(map.Rows[r]);
        return !t.Items.OfType<TableRow>().Any();
    }

    // ================================================================ colonnes

    /// <summary>
    /// Insère <paramref name="count"/> colonnes avant la colonne de grille <paramref name="at"/> (= nombre de colonnes :
    /// à la fin), de la largeur de la colonne voisine. Une cellule fusionnée qui traverse le point d'insertion s'élargit.
    /// <paramref name="fitTwips"/> : largeur disponible ; un tableau qui la dépasse est ramené à cette largeur (Word).
    /// </summary>
    public static void InsertColumns(Table t, int at, int count, Func<string> paraId, int? fitTwips = null)
    {
        var map = new Map(t);
        var w = map.W;
        var grid = map.GridWidths();
        at = Math.Clamp(at, 0, map.Columns);
        int source = Math.Min(at, grid.Count - 1);
        int width = grid.Count > 0 ? grid[Math.Max(0, source)] : 1440;
        for (int r = 0; r < map.Cells.Count; r++)
        {
            var row = map.Rows[r];
            var slots = map.Cells[r];
            var inside = slots.Cast<Slot?>().FirstOrDefault(s => s!.Value.Col < at && at < s.Value.Col + s.Value.Span);
            if (inside is { } wide)
            {
                wide.Cell.Props.SetVal("gridSpan", (wide.Span + count).ToString(CultureInfo.InvariantCulture));
                continue;
            }
            // cellule voisine : celle qui commence au point d'insertion, sinon la dernière avant
            var after = slots.Cast<Slot?>().FirstOrDefault(s => s!.Value.Col >= at);
            var model = after ?? (slots.Count > 0 ? slots[^1] : null);
            if (model is null) continue;
            int index = after is { } a ? row.Items.IndexOf(a.Cell) : row.Items.IndexOf(slots[^1].Cell) + 1;
            for (int k = 0; k < count; k++)
            {
                var props = model.Value.Cell.Props.Clone();
                foreach (var n in new[] { "cnfStyle", "gridSpan", "vMerge", "hMerge" }) props.Remove(n);
                props.Set(Width(w, "tcW", width));
                row.Items.Insert(index + k, NewCell(w, props, model.Value.Cell.Content.Paragraphs().FirstOrDefault(), paraId));
            }
        }
        grid.InsertRange(at, Enumerable.Repeat(width, count));
        if (fitTwips is { } fit && grid.Sum() > fit && grid.Sum() > 0)
        {
            double k = (double)fit / grid.Sum();
            for (int i = 0; i < grid.Count; i++) grid[i] = Math.Max(1, (int)Math.Round(grid[i] * k));
        }
        SetGrid(t, grid);
    }

    /// <summary>Supprime les colonnes de grille <paramref name="c0"/>..<paramref name="c1"/>. Vrai si le tableau est vide.</summary>
    public static bool DeleteColumns(Table t, int c0, int c1)
    {
        var map = new Map(t);
        if (c0 > c1) (c0, c1) = (c1, c0);
        int n = c1 - c0 + 1;
        for (int r = 0; r < map.Cells.Count; r++)
        {
            var row = map.Rows[r];
            foreach (var s in map.Cells[r])
            {
                int end = s.Col + s.Span - 1;
                int overlap = Math.Min(end, c1) - Math.Max(s.Col, c0) + 1;
                if (overlap <= 0) continue;
                if (overlap >= s.Span) row.Items.Remove(s.Cell);
                else s.Cell.Props.SetVal("gridSpan", s.Span - overlap == 1 ? null : (s.Span - overlap).ToString(CultureInfo.InvariantCulture));
            }
            if (!row.Items.OfType<TableCell>().Any()) t.Items.Remove(row);
        }
        var grid = map.GridWidths();
        if (c0 < grid.Count) grid.RemoveRange(c0, Math.Min(n, grid.Count - c0));
        SetGrid(t, grid);
        return !t.Items.OfType<TableRow>().Any();
    }

    /// <summary>Écrit la grille et met à jour les largeurs préférées des cellules en points (w:tcW dxa).</summary>
    public static void SetGrid(Table t, IReadOnlyList<int> grid)
    {
        var w = t.Props.Ns;
        t.Grid = new XElement(w + "tblGrid", grid.Select(x => new XElement(w + "gridCol", new XAttribute(w + "w", x))));
        var map = new Map(t);
        foreach (var slots in map.Cells)
            foreach (var s in slots)
            {
                var tcw = s.Cell.Props.Get("tcW");
                if (tcw?.Attribute(w + "type")?.Value is "pct" or "auto") continue;
                s.Cell.Props.Set(Width(w, "tcW", Enumerable.Range(s.Col, s.Span).Where(c => c < grid.Count).Sum(c => grid[c])));
            }
        if (t.Props.Get("tblW") is { } tw && tw.Attribute(w + "type")?.Value == "dxa") t.Props.Set(Width(w, "tblW", grid.Sum()));
    }

    // ================================================================ fusion et fractionnement

    /// <summary>
    /// Fusionne le rectangle de cellules (agrandi aux fusions qu'il touche) : le texte de toutes les cellules va dans la
    /// première (paragraphes vides ignorés), chaque rangée garde une cellule qui couvre les colonnes (w:gridSpan), plusieurs
    /// rangées forment une fusion verticale (w:vMerge). Renvoie la cellule fusionnée.
    /// </summary>
    public static TableCell Merge(Table t, int r0, int c0, int r1, int c1, Func<string> paraId)
    {
        var map = new Map(t);
        var w = map.W;
        (r0, c0, r1, c1) = map.Normalize(r0, c0, r1, c1);
        var grid = map.GridWidths();
        var first = map.At(r0, c0)!.Value.Cell;
        // contenu : celui de toutes les cellules, dans l'ordre
        var keep = first.Content.Blocks.Where(b => !IsEmptyParagraph(b)).ToList();
        foreach (var (_, s) in map.In(r0, c0, r1, c1))
        {
            if (ReferenceEquals(s.Cell, first)) continue;
            keep.AddRange(s.Cell.Content.Blocks.Where(b => !IsEmptyParagraph(b)));
        }
        if (keep.Count == 0 || keep[^1] is not Paragraph) keep.Add(EmptyParagraph(w, first.Content.Paragraphs().FirstOrDefault(), paraId));
        first.Content.Blocks.Clear();
        first.Content.Blocks.AddRange(keep);
        int span = c1 - c0 + 1;
        int width = Enumerable.Range(c0, span).Where(c => c < grid.Count).Sum(c => grid[c]);
        for (int r = r0; r <= r1; r++)
        {
            var row = map.Rows[r];
            var inRow = map.Cells[r].Where(s => s.Col + s.Span - 1 >= c0 && s.Col <= c1).ToList();
            if (inRow.Count == 0) continue;
            var lead = inRow[0].Cell;
            foreach (var s in inRow.Skip(1)) row.Items.Remove(s.Cell);
            lead.Props.SetVal("gridSpan", span == 1 ? null : span.ToString(CultureInfo.InvariantCulture));
            if (lead.Props.Get("tcW")?.Attribute(w + "type")?.Value is not ("auto" or "pct")) lead.Props.Set(Width(w, "tcW", width));
            if (r1 > r0)
                lead.Props.Set(r == r0 ? new XElement(w + "vMerge", new XAttribute(w + "val", "restart")) : new XElement(w + "vMerge"));
            else lead.Props.Remove("vMerge");
            if (r > r0)
            {
                var like = lead.Content.Paragraphs().FirstOrDefault() ?? first.Content.Paragraphs().FirstOrDefault();
                lead.Content.Blocks.Clear();
                lead.Content.Blocks.Add(EmptyParagraph(w, like, paraId));
            }
        }
        return first;
    }

    static bool IsEmptyParagraph(Block b) => b is Paragraph p && ParagraphAtoms.Length(p) == 0 && p.Props.Get("sectPr") is null;

    /// <summary>
    /// Fractionne la cellule (rangée, colonne de départ) en <paramref name="cols"/> × <paramref name="rows"/> cellules,
    /// comme Word : les colonnes nouvelles sont ajoutées à la grille (les cellules des autres rangées s'élargissent pour
    /// les couvrir), les rangées nouvelles fusionnent verticalement les autres cellules de la rangée. Le texte reste
    /// dans la première cellule. Une cellule fusionnée verticalement est fractionnée sur toutes ses rangées.
    /// </summary>
    public static void Split(Table t, int row, int col, int cols, int rows, Func<string> paraId)
    {
        cols = Math.Clamp(cols, 1, 63);
        rows = Math.Clamp(rows, 1, 100);
        var map = new Map(t);
        var w = map.W;
        if (map.At(row, col) is not { } slot) return;
        int top = map.MergeTop(row, slot.Col), bottom = map.MergeBottom(row, slot.Col);
        if (cols > 1)
        {
            // nouvelles limites de colonnes dans la largeur de la cellule
            var grid = map.GridWidths();
            int x0 = grid.Take(slot.Col).Sum(), width = grid.Skip(slot.Col).Take(slot.Span).Sum();
            var bounds = new List<int>();
            for (int k = 1; k < cols; k++) bounds.Add(x0 + (int)Math.Round(width * k / (double)cols));
            foreach (var b in bounds)
            {
                // limite absente de la grille : la colonne qui la contient est coupée, les cellules qui la couvrent s'élargissent
                var pos = new List<int> { 0 };
                foreach (var g in grid) pos.Add(pos[^1] + g);
                if (pos.Any(p => Math.Abs(p - b) <= 1)) continue;
                int c = pos.FindIndex(p => p > b) - 1;
                int left = b - pos[c], right = grid[c] - left;
                grid[c] = left;
                grid.Insert(c + 1, right);
                var m = new Map(t);
                foreach (var slots in m.Cells)
                    foreach (var s in slots)
                        if (s.Col <= c && c < s.Col + s.Span) s.Cell.Props.SetVal("gridSpan", (s.Span + 1).ToString(CultureInfo.InvariantCulture));
                // w:gridBefore / w:gridAfter des rangées qui sautent cette colonne
                foreach (var r in m.Rows)
                {
                    if (Int(r.Props.Get("gridBefore"), w) is { } gb && gb > c) r.Props.Set(new XElement(w + "gridBefore", new XAttribute(w + "val", gb + 1)));
                }
                SetGridOnly(t, grid);
            }
            // chaque rangée de la cellule : cellules de départ selon les limites
            var map2 = new Map(t);
            var pos2 = new List<int> { 0 };
            foreach (var g in grid) pos2.Add(pos2[^1] + g);
            var edges = new List<int> { x0 };
            edges.AddRange(bounds);
            edges.Add(x0 + width);
            for (int r = top; r <= bottom; r++)
            {
                var s = map2.Cells[r].First(x => pos2[x.Col] >= x0 - 1 && pos2[x.Col] <= x0 + 1);
                var rowObj = map2.Rows[r];
                int index = rowObj.Items.IndexOf(s.Cell);
                var like = s.Cell.Content.Paragraphs().FirstOrDefault();
                for (int k = 0; k < cols; k++)
                {
                    int ca = pos2.FindIndex(p => Math.Abs(p - edges[k]) <= 1), cb = pos2.FindIndex(p => Math.Abs(p - edges[k + 1]) <= 1);
                    int span = Math.Max(1, cb - ca);
                    var cell = k == 0 ? s.Cell : NewCell(w, CloneCellProps(s.Cell), like, paraId);
                    cell.Props.SetVal("gridSpan", span == 1 ? null : span.ToString(CultureInfo.InvariantCulture));
                    if (cell.Props.Get("tcW")?.Attribute(w + "type")?.Value is not ("auto" or "pct")) cell.Props.Set(Width(w, "tcW", edges[k + 1] - edges[k]));
                    if (k > 0) rowObj.Items.Insert(index + k, cell);
                }
            }
            SetGridOnly(t, grid);
        }
        if (rows > 1)
        {
            var m = new Map(t);
            var target = m.Rows[bottom];
            var targetSlots = m.Cells[bottom];
            var splitCols = m.Cells[row].Where(s => s.Col >= slot.Col && s.Col < slot.Col + slot.Span).Select(s => s.Col).ToHashSet();
            if (splitCols.Count == 0) splitCols.Add(slot.Col);
            // autres cellules de la rangée : fusion verticale sur les nouvelles rangées
            foreach (var s in targetSlots)
                if (!splitCols.Contains(s.Col) && !m.IsContinue(s.Cell) && !m.IsRestart(s.Cell))
                    s.Cell.Props.Set(new XElement(w + "vMerge", new XAttribute(w + "val", "restart")));
            int insertAt = t.Items.IndexOf(target) + 1;
            for (int k = 1; k < rows; k++)
            {
                var nr = new TableRow { Props = target.Props.Clone() };
                nr.Props.Remove("cnfStyle");
                nr.Props.Remove("trHeight");
                foreach (var s in targetSlots)
                {
                    var props = CloneCellProps(s.Cell);
                    if (splitCols.Contains(s.Col)) props.Remove("vMerge");
                    else props.Set(new XElement(w + "vMerge"));
                    nr.Items.Add(NewCell(w, props, s.Cell.Content.Paragraphs().FirstOrDefault(), paraId));
                }
                t.Items.Insert(insertAt + k - 1, nr);
            }
        }
    }

    static PropertyBag CloneCellProps(TableCell c)
    {
        var p = c.Props.Clone();
        p.Remove("cnfStyle");
        p.Remove("vMerge");
        return p;
    }

    static void SetGridOnly(Table t, IReadOnlyList<int> grid)
    {
        var w = t.Props.Ns;
        t.Grid = new XElement(w + "tblGrid", grid.Select(x => new XElement(w + "gridCol", new XAttribute(w + "w", x))));
    }

    /// <summary>
    /// Fractionne le tableau avant la rangée <paramref name="row"/> : un paragraphe vide sépare les deux tableaux (Word,
    /// Ctrl+Maj+Entrée). Renvoie la position de ce paragraphe. Transaction : bloc du tableau.
    /// </summary>
    public static DocPos SplitTable(Story root, IReadOnlyList<int> tablePath, int row, Func<string> paraId)
    {
        var container = DocPos.ResolveContainer(root, tablePath.Take(tablePath.Count - 1).ToArray());
        int index = tablePath[^1];
        var t = (Table)container.Blocks[index];
        var map = new Map(t);
        var w = map.W;
        if (row <= 0 || row >= map.Rows.Count) throw new ArgumentException("Le tableau ne se fractionne qu'entre deux rangées.");
        var second = new Table { Props = t.Props.Clone(), Grid = t.Grid is null ? null : new XElement(t.Grid) };
        foreach (var a in t.Attributes) second.Attributes.Add(new XAttribute(a));
        int at = t.Items.IndexOf(map.Rows[row]);
        second.Items.AddRange(t.Items.Skip(at));
        t.Items.RemoveRange(at, t.Items.Count - at);
        // une fusion verticale coupée recommence dans le second tableau
        foreach (var s in new Map(second).Cells.FirstOrDefault() ?? new List<Slot>())
            if (map.IsContinue(s.Cell)) s.Cell.Props.Set(new XElement(w + "vMerge", new XAttribute(w + "val", "restart")));
        var like = map.Cells[row - 1].FirstOrDefault().Cell?.Content.Paragraphs().FirstOrDefault();
        var p = EmptyParagraph(w, like, paraId);
        p.Props.Remove("pStyle");
        p.Props.Remove("jc");
        container.Blocks.Insert(index + 1, p);
        container.Blocks.Insert(index + 2, second);
        return new DocPos(tablePath.Take(tablePath.Count - 1).Append(index + 1).ToArray(), 0);
    }

    /// <summary>Copie d'un rectangle de cellules en tableau autonome (copier des cellules) : rangées et colonnes hors du rectangle retirées.</summary>
    public static Table Extract(Table t, int r0, int c0, int r1, int c1)
    {
        var copy = (Table)t.Clone();
        var map = new Map(copy);
        (r0, c0, r1, c1) = map.Normalize(r0, c0, r1, c1);
        if (r1 + 1 < map.Rows.Count) DeleteRows(copy, r1 + 1, map.Rows.Count - 1);
        if (r0 > 0) DeleteRows(copy, 0, r0 - 1);
        int cols = new Map(copy).Columns;
        if (c1 + 1 < cols) DeleteColumns(copy, c1 + 1, cols - 1);
        if (c0 > 0) DeleteColumns(copy, 0, c0 - 1);
        foreach (var row in copy.Items.OfType<TableRow>()) row.Props.Remove("gridBefore");
        return copy;
    }

    // ================================================================ largeurs et hauteurs

    /// <summary>
    /// Déplace la limite <paramref name="boundary"/> entre les colonnes de grille boundary-1 et boundary (bord droit du
    /// tableau si = nombre de colonnes) de <paramref name="deltaTwips"/> : la colonne voisine compense (le tableau garde sa
    /// largeur), sauf au bord droit. Les largeurs deviennent fixes (w:tcW en points), comme quand on tire une bordure dans Word.
    /// </summary>
    public static void MoveBoundary(Table t, int boundary, int deltaTwips, int minTwips = 200)
    {
        var map = new Map(t);
        var grid = map.GridWidths();
        if (boundary <= 0 || boundary > grid.Count) return;
        int left = boundary - 1;
        if (boundary == grid.Count)
            grid[left] = Math.Max(minTwips, grid[left] + deltaTwips);
        else
        {
            int d = Math.Clamp(deltaTwips, minTwips - grid[left], grid[boundary] - minTwips);
            grid[left] += d;
            grid[boundary] -= d;
        }
        FixWidths(t);
        SetGrid(t, grid);
    }

    /// <summary>Largeurs automatiques ou en pourcentage remplacées par des largeurs en points (avant un redimensionnement).</summary>
    static void FixWidths(Table t)
    {
        var w = t.Props.Ns;
        foreach (var c in t.Items.OfType<TableRow>().SelectMany(r => r.Items.OfType<TableCell>()))
            if (c.Props.Get("tcW")?.Attribute(w + "type")?.Value is "auto" or "pct") c.Props.Set(Width(w, "tcW", 0));
        if (t.Props.Get("tblW")?.Attribute(w + "type")?.Value == "pct") t.Props.Set(Width(w, "tblW", 0, "auto"));
    }

    /// <summary>Largeur des colonnes de grille <paramref name="c0"/>..<paramref name="c1"/> (twips).</summary>
    public static void SetColumnWidth(Table t, int c0, int c1, int twips)
    {
        var grid = new Map(t).GridWidths();
        for (int c = Math.Max(0, c0); c <= c1 && c < grid.Count; c++) grid[c] = Math.Max(100, twips);
        FixWidths(t);
        SetGrid(t, grid);
    }

    /// <summary>Colonnes <paramref name="c0"/>..<paramref name="c1"/> de même largeur (leur largeur totale gardée).</summary>
    public static void DistributeColumns(Table t, int c0, int c1)
    {
        var grid = new Map(t).GridWidths();
        c1 = Math.Min(c1, grid.Count - 1);
        if (c1 <= c0) return;
        int total = grid.Skip(c0).Take(c1 - c0 + 1).Sum(), n = c1 - c0 + 1;
        for (int c = c0; c <= c1; c++) grid[c] = total / n;
        grid[c1] += total - total / n * n;
        FixWidths(t);
        SetGrid(t, grid);
    }

    /// <summary>Hauteur de rangée (w:trHeight) : null = automatique ; <paramref name="rule"/> atLeast ou exact.</summary>
    public static void SetRowHeight(TableRow row, double? points, string rule = "atLeast")
    {
        var w = row.Props.Ns;
        if (points is not { } h || h <= 0) { row.Props.Remove("trHeight"); return; }
        var e = new XElement(w + "trHeight", new XAttribute(w + "val", Tw(h * 20)));
        if (rule == "exact") e.SetAttributeValue(w + "hRule", "exact");
        row.Props.Set(e);
    }

    // ================================================================ mise en forme des cellules et du tableau

    /// <summary>Trame d'une cellule (RRGGBB) ; null : sans couleur (w:fill="auto", masque aussi la trame du style).</summary>
    public static void SetShading(TableCell cell, string? fill)
    {
        var w = cell.Props.Ns;
        cell.Props.Set(new XElement(w + "shd", new XAttribute(w + "val", "clear"), new XAttribute(w + "color", "auto"), new XAttribute(w + "fill", fill ?? "auto")));
    }

    /// <summary>Trait de bordure : style Word (single, double, dotted, dashed, thick), épaisseur en huitièmes de point, couleur.</summary>
    public readonly record struct Line(string Style = "single", int Size = 4, string Color = "auto")
    {
        public XElement Element(XNamespace w, string side) => Style is "none" or "nil"
            ? new XElement(w + side, new XAttribute(w + "val", "nil"))
            : new XElement(w + side, new XAttribute(w + "val", Style), new XAttribute(w + "sz", Size), new XAttribute(w + "space", 0), new XAttribute(w + "color", Color));
    }

    /// <summary>
    /// Bordures d'un rectangle de cellules, comme le menu Bordures de Word : all (toutes), outside (extérieures), inside
    /// (intérieures), insideH, insideV, top, bottom, left, right, none (aucune). Écrites sur chaque cellule (w:tcBorders) ;
    /// un trait intérieur est posé sur les deux cellules qui le partagent.
    /// </summary>
    public static void SetBorders(Table t, int r0, int c0, int r1, int c1, string which, Line line)
    {
        var map = new Map(t);
        var w = map.W;
        (r0, c0, r1, c1) = map.Normalize(r0, c0, r1, c1);
        var none = new Line("nil");
        foreach (var (r, s) in map.In(r0, c0, r1, c1))
        {
            int rowEnd = map.MergeBottom(r, s.Col);
            bool top = r == r0, bottom = rowEnd >= r1, left = s.Col <= c0, right = s.Col + s.Span - 1 >= c1;
            var sides = new Dictionary<string, Line?>
            {
                ["top"] = null, ["bottom"] = null, ["left"] = null, ["right"] = null,
            };
            switch (which)
            {
                case "all": sides["top"] = sides["bottom"] = sides["left"] = sides["right"] = line; break;
                case "none": sides["top"] = sides["bottom"] = sides["left"] = sides["right"] = none; break;
                case "outside":
                    if (top) sides["top"] = line;
                    if (bottom) sides["bottom"] = line;
                    if (left) sides["left"] = line;
                    if (right) sides["right"] = line;
                    break;
                case "inside":
                    if (!top) sides["top"] = line;
                    if (!bottom) sides["bottom"] = line;
                    if (!left) sides["left"] = line;
                    if (!right) sides["right"] = line;
                    break;
                case "insideH":
                    if (!top) sides["top"] = line;
                    if (!bottom) sides["bottom"] = line;
                    break;
                case "insideV":
                    if (!left) sides["left"] = line;
                    if (!right) sides["right"] = line;
                    break;
                case "top": if (top) sides["top"] = line; break;
                case "bottom": if (bottom) sides["bottom"] = line; break;
                case "left": if (left) sides["left"] = line; break;
                case "right": if (right) sides["right"] = line; break;
            }
            if (sides.Values.All(v => v is null)) continue;
            var borders = s.Cell.Props.Get("tcBorders") is { } b ? new XElement(b) : new XElement(w + "tcBorders");
            foreach (var side in new[] { "top", "left", "bottom", "right" })
            {
                if (sides[side] is not { } l) continue;
                borders.Element(w + side)?.Remove();
                if (side == "left") borders.Element(w + "start")?.Remove();
                if (side == "right") borders.Element(w + "end")?.Remove();
                borders.Add(l.Element(w, side));
            }
            // ordre du schéma : top, start/left, bottom, end/right, insideH, insideV, tl2br, tr2bl
            var order = new[] { "top", "start", "left", "bottom", "end", "right", "insideH", "insideV", "tl2br", "tr2bl" };
            var sorted = borders.Elements().OrderBy(e => Array.IndexOf(order, e.Name.LocalName) is var i && i < 0 ? 99 : i).ToList();
            borders.RemoveNodes();
            borders.Add(sorted);
            s.Cell.Props.Set(borders);
        }
    }

    /// <summary>Alignement d'une cellule : vertical (top, center, bottom) et horizontal de ses paragraphes (left, center, right).</summary>
    public static void SetAlignment(TableCell cell, string? vertical, string? horizontal)
    {
        if (vertical is not null) cell.Props.SetVal("vAlign", vertical == "top" ? null : vertical);
        if (horizontal is not null)
            foreach (var p in cell.Content.Paragraphs()) p.Props.SetVal("jc", horizontal);
    }

    /// <summary>Marges de cellule par défaut du tableau (w:tblCellMar, points ; null : inchangée).</summary>
    public static void SetCellMargins(Table t, double? top, double? bottom, double? left, double? right)
    {
        var w = t.Props.Ns;
        var mar = t.Props.Get("tblCellMar") is { } m ? new XElement(m) : new XElement(w + "tblCellMar");
        void Side(string name, double? v)
        {
            if (v is not { } x) return;
            mar.Element(w + name)?.Remove();
            if (name == "left") mar.Element(w + "start")?.Remove();
            if (name == "right") mar.Element(w + "end")?.Remove();
            mar.Add(Width(w, name, (int)Math.Round(x * 20)));
        }
        Side("top", top); Side("left", left); Side("bottom", bottom); Side("right", right);
        var order = new[] { "top", "start", "left", "bottom", "end", "right" };
        var sorted = mar.Elements().OrderBy(e => Array.IndexOf(order, e.Name.LocalName)).ToList();
        mar.RemoveNodes();
        mar.Add(sorted);
        t.Props.Set(mar);
    }

    /// <summary>Style du tableau ; les mises en forme conditionnelles en cache (w:cnfStyle) sont retirées, Word les recalcule.</summary>
    public static void SetStyle(Table t, string? styleId)
    {
        t.Props.SetVal("tblStyle", styleId);
        RemoveCnf(t);
    }

    /// <summary>Options de style du tableau (ligne d'en-tête, bandes…).</summary>
    public static void SetLook(Table t, TableLook look)
    {
        t.Props.Set(LookElement(t.Props.Ns, look));
        RemoveCnf(t);
    }

    static void RemoveCnf(Table t)
    {
        foreach (var row in t.Items.OfType<TableRow>())
        {
            row.Props.Remove("cnfStyle");
            foreach (var c in row.Items.OfType<TableCell>())
            {
                c.Props.Remove("cnfStyle");
                foreach (var p in c.Content.Blocks.OfType<Paragraph>()) p.Props.Remove("cnfStyle");
            }
        }
    }

    /// <summary>Rangée d'en-tête répétée en haut de chaque page (w:tblHeader).</summary>
    public static void SetHeaderRow(TableRow row, bool on) => row.Props.SetFlag("tblHeader", on);

    /// <summary>Autoriser le fractionnement de la rangée sur deux pages (w:cantSplit absent).</summary>
    public static void SetAllowBreak(TableRow row, bool allow) => row.Props.SetFlag("cantSplit", !allow);

    /// <summary>Alignement du tableau (left, center, right) et retrait à gauche (points).</summary>
    public static void SetTableAlignment(Table t, string? jc, double? indent = null)
    {
        var w = t.Props.Ns;
        if (jc is not null) t.Props.SetVal("jc", jc == "left" ? null : jc);
        if (indent is { } ind) t.Props.Set(Width(w, "tblInd", (int)Math.Round(ind * 20)));
        if (jc is "center" or "right") t.Props.Remove("tblInd");
        foreach (var r in t.Items.OfType<TableRow>()) r.Props.Remove("jc");
    }

    /// <summary>
    /// Ajustement automatique, comme Word : « content » (au contenu : largeurs automatiques), « window » (à la fenêtre :
    /// 100 % de la largeur, colonnes en pourcentage), « fixed » (largeur de colonne fixe : grille actuelle en points).
    /// </summary>
    public static void AutoFit(Table t, string mode, IReadOnlyList<int>? currentGrid = null)
    {
        var w = t.Props.Ns;
        var map = new Map(t);
        var grid = currentGrid?.ToList() ?? map.GridWidths();
        int total = Math.Max(1, grid.Sum());
        switch (mode)
        {
            case "content":
                t.Props.Remove("tblLayout");
                t.Props.Set(Width(w, "tblW", 0, "auto"));
                foreach (var s in map.Cells.SelectMany(x => x)) s.Cell.Props.Set(Width(w, "tcW", 0, "auto"));
                break;
            case "window":
                t.Props.Remove("tblLayout");
                t.Props.Set(Width(w, "tblW", 5000, "pct"));
                foreach (var s in map.Cells.SelectMany(x => x))
                    s.Cell.Props.Set(Width(w, "tcW", (int)Math.Round(5000.0 * Enumerable.Range(s.Col, s.Span).Where(c => c < grid.Count).Sum(c => grid[c]) / total), "pct"));
                break;
            default:
                t.Props.Set(new XElement(w + "tblLayout", new XAttribute(w + "type", "fixed")));
                t.Props.Set(Width(w, "tblW", 0, "auto"));
                foreach (var s in map.Cells.SelectMany(x => x)) s.Cell.Props.Set(Width(w, "tcW", 0));
                SetGrid(t, grid);
                break;
        }
    }

    // ================================================================ conversions, tri, formule

    /// <summary>
    /// Tableau → texte : chaque rangée devient un paragraphe, les cellules séparées par <paramref name="separator"/>
    /// (tabulation par défaut) ; les paragraphes intérieurs d'une cellule restent des paragraphes.
    /// </summary>
    public static List<Paragraph> ToText(Table t, string separator, Func<string> paraId)
    {
        var w = t.Props.Ns;
        var result = new List<Paragraph>();
        foreach (var row in t.Items.OfType<TableRow>())
        {
            Paragraph? current = null;
            foreach (var cell in row.Items.OfType<TableCell>())
            {
                var paras = cell.Content.Paragraphs().ToList();
                for (int k = 0; k < paras.Count; k++)
                {
                    var src = paras[k];
                    if (current is null || k > 0)
                    {
                        current = new Paragraph { Props = src.Props.Clone() };
                        current.Props.Remove("cnfStyle");
                        foreach (var a in src.Attributes)
                            current.Attributes.Add(a.Name == FlowEdits.W14 + "paraId" || a.Name == FlowEdits.W14 + "textId" ? new XAttribute(a.Name, paraId()) : new XAttribute(a));
                        result.Add(current);
                    }
                    else
                    {
                        var sep = new Run();
                        if (separator == "\t") sep.Content.Add(new TabContent());
                        else sep.Content.Add(new TextContent { Value = separator });
                        current.Inlines.Add(sep);
                    }
                    current.Inlines.AddRange(((Paragraph)src.Clone()).Inlines);
                }
            }
        }
        return result;
    }

    /// <summary>
    /// Texte → tableau : une rangée par paragraphe, les cellules coupées au séparateur (tabulation, « ; », « , » ou autre
    /// caractère) ; nombre de colonnes = la rangée la plus longue. Chaque cellule garde la mise en forme du texte.
    /// </summary>
    public static Table FromParagraphs(IReadOnlyList<Paragraph> paragraphs, char separator, double textWidthPt, string? styleId, Func<string> paraId)
    {
        var w = paragraphs.Count > 0 ? paragraphs[0].Props.Ns : (XNamespace)PropertyBag.TransitionalNs;
        var rows = new List<List<Paragraph>>();
        foreach (var p in paragraphs)
        {
            var temp = new Story();
            var copy = (Paragraph)p.Clone();
            copy.Props.Remove("sectPr");
            copy.Props.Remove("numPr");
            temp.Blocks.Add(copy);
            // coupures au séparateur, de la fin vers le début (les décalages restent valables)
            var text = AtomText(copy);
            for (int k = text.Count - 1; k >= 0; k--)
            {
                if (text[k] != separator) continue;
                TextOps.Split(temp, new DocPos(new[] { 0 }, k), paraId);
                var next = new DocPos(new[] { 1 }, 0);
                TextOps.Delete(temp, next, next.WithOffset(1));
            }
            rows.Add(temp.Blocks.OfType<Paragraph>().ToList());
        }
        int cols = Math.Max(1, rows.Count == 0 ? 1 : rows.Max(r => r.Count));
        var table = Create(w, Math.Max(1, rows.Count), cols, textWidthPt, styleId, paragraphs.FirstOrDefault(), paraId);
        var map = new Map(table);
        for (int r = 0; r < rows.Count; r++)
            for (int c = 0; c < rows[r].Count; c++)
            {
                var cell = map.Cells[r][c].Cell;
                cell.Content.Blocks.Clear();
                cell.Content.Blocks.Add(rows[r][c]);
            }
        return table;
    }

    /// <summary>Caractère de chaque atome du paragraphe (tabulation = '\t', autres atomes non textuels = '\0').</summary>
    static List<char> AtomText(Paragraph p)
    {
        var res = new List<char>();
        foreach (var a in ParagraphAtoms.Of(p))
        {
            if (a.Content is TextContent t && !t.Deleted) res.AddRange(t.Value);
            else if (a.Content is TabContent) res.Add('\t');
            else for (int k = 0; k < a.Length; k++) res.Add('\0');
        }
        return res;
    }

    /// <summary>Texte d'une cellule (paragraphes séparés par des espaces).</summary>
    public static string CellText(TableCell c) => string.Join(" ", c.Content.Paragraphs().Select(p => p.Text)).Trim();

    /// <summary>
    /// Trie les rangées (sauf les <paramref name="headerRows"/> premières) selon le texte de la colonne de grille
    /// <paramref name="col"/> : nombres si toutes les valeurs en sont, sinon texte (ordre français). Refusé si des
    /// cellules sont fusionnées verticalement.
    /// </summary>
    public static void Sort(Table t, int col, bool ascending, int headerRows)
    {
        var map = new Map(t);
        if (map.Cells.SelectMany(x => x).Any(s => map.IsContinue(s.Cell) || map.IsRestart(s.Cell)))
            throw new NotSupportedException("Le tableau contient des cellules fusionnées verticalement.");
        var body = map.Rows.Skip(headerRows).ToList();
        string Key(TableRow r) => map.At(map.Rows.IndexOf(r), col) is { } s ? CellText(s.Cell) : "";
        bool numeric = body.All(r => Key(r).Length == 0 || TryNumber(Key(r), out _));
        var culture = CultureInfo.GetCultureInfo("fr-FR");
        var sorted = numeric
            ? body.OrderBy(r => TryNumber(Key(r), out var v) ? v : double.MinValue).ToList()
            : body.OrderBy(r => Key(r), StringComparer.Create(culture, CompareOptions.IgnoreCase)).ToList();
        if (!ascending) sorted.Reverse();
        int first = t.Items.IndexOf(body.FirstOrDefault()!);
        if (first < 0) return;
        foreach (var r in body) t.Items.Remove(r);
        t.Items.InsertRange(first, sorted);
    }

    static bool TryNumber(string s, out double v)
    {
        var clean = s.Replace(" ", "").Replace(" ", "").Replace(" ", "").Replace("€", "").Replace("%", "").Trim();
        return double.TryParse(clean.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out v);
    }

    /// <summary>
    /// Formule de tableau de Word (=SUM(ABOVE), AVERAGE, MIN, MAX, COUNT, PRODUCT ; ABOVE, BELOW, LEFT, RIGHT) calculée
    /// pour la cellule (rangée, colonne) : nombres des cellules dans la direction, jusqu'à la première cellule vide après
    /// des nombres (Word). Null si la formule n'est pas comprise.
    /// </summary>
    public static double? Compute(Table t, int row, int col, string formula)
    {
        var m = System.Text.RegularExpressions.Regex.Match(formula.Trim(), @"^=?\s*(SUM|AVERAGE|MIN|MAX|COUNT|PRODUCT)\s*\(\s*(ABOVE|BELOW|LEFT|RIGHT)\s*\)\s*$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (!m.Success) return null;
        var map = new Map(t);
        var values = new List<double>();
        var dir = m.Groups[2].Value.ToUpperInvariant();
        int dr = dir switch { "ABOVE" => -1, "BELOW" => 1, _ => 0 }, dc = dir switch { "LEFT" => -1, "RIGHT" => 1, _ => 0 };
        var start = map.At(row, col);
        int r = row + dr, c = dc < 0 ? (start?.Col ?? col) - 1 : dc > 0 ? (start is { } st ? st.Col + st.Span : col + 1) : col;
        bool seen = false;
        while (r >= 0 && r < map.Rows.Count && c >= 0 && c < map.Columns)
        {
            if (map.At(r, c) is { } s)
            {
                var text = CellText(s.Cell);
                if (TryNumber(text, out var v)) { values.Add(v); seen = true; }
                else if (seen && text.Length == 0 && dc == 0) break;
                if (dc != 0) c = dc > 0 ? s.Col + s.Span : s.Col - 1;
            }
            else if (dc != 0) c += dc;
            r += dr;
        }
        return m.Groups[1].Value.ToUpperInvariant() switch
        {
            "SUM" => values.Sum(),
            "AVERAGE" => values.Count == 0 ? 0 : values.Average(),
            "MIN" => values.Count == 0 ? 0 : values.Min(),
            "MAX" => values.Count == 0 ? 0 : values.Max(),
            "COUNT" => values.Count,
            "PRODUCT" => values.Aggregate(1.0, (a, b) => a * b),
            _ => null,
        };
    }

    /// <summary>Résultat d'une formule écrit comme Word en français (virgule décimale, deux décimales au plus).</summary>
    public static string FormatNumber(double v) => v.ToString(v == Math.Floor(v) ? "0" : "0.##", CultureInfo.GetCultureInfo("fr-FR"));
}
