using System.Globalization;
using System.Xml.Linq;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;

namespace OpenWrite.Layout;

/// <summary>Trait de bordure : épaisseur (points ; 0 = aucun), couleur RRGGBB (null = automatique), style Word (single, double…).</summary>
public readonly record struct BorderLine(double Width, string? Color = null, string Style = "single")
{
    public bool Visible => Width > 0;

    /// <summary>
    /// Place prise par le trait (points) : un trait double fait deux traits et l'espace entre eux (3 fois l'épaisseur
    /// w:sz), un triple 5 fois ; les traits composés « fin-épais » sont comptés comme doubles.
    /// </summary>
    public double Thickness => (Style ?? "") switch
    {
        "double" => 3 * Width,
        "triple" => 5 * Width,
        var s when s.StartsWith("thinThick", StringComparison.Ordinal) || s.StartsWith("thickThin", StringComparison.Ordinal) => 3 * Width,
        _ => Width,
    };
}

// Mise en page des tableaux (lot 6, PLAN.md §8.7) :
// - grille : w:tblGrid, complétée par w:tcW quand elle manque ; tableau « ajusté au contenu » (w:tblW et tous les w:tcW
//   automatiques, disposition non fixe) recalculé d'après le contenu (largeurs minimale et maximale des cellules) ;
//   w:tblW en pourcentage ou en points : grille mise à l'échelle ;
// - position : w:tblInd (Word 2013 et suivants : depuis le bord du tableau ; avant : depuis le texte de la première
//   cellule), w:jc centré / à droite ;
// - cellules : marges (w:tblCellMar, w:tcMar, style compris), alignement vertical, fusion horizontale (w:gridSpan,
//   w:gridBefore / w:gridAfter) et verticale (w:vMerge : la dernière rangée fusionnée grandit si le contenu dépasse),
//   tableaux imbriqués, fond et bordures résolus (cellule > exceptions de rangée > tableau > mise en forme
//   conditionnelle du style > style), régions conditionnelles transmises au résolveur (<see cref="CellStyle"/>).
public sealed partial class DocumentLayout
{
    /// <summary>Cellule mesurée dans sa rangée.</summary>
    sealed class CellPart
    {
        public required TableCell Cell { get; init; }
        public int Index, Col, Span = 1;
        public double X, W;
        public string? Fill;
        public BorderLine Top, Bottom, Left, Right;
        public string VAlign = "top";
        /// <summary>Rangées couvertes par une cellule qui commence une fusion verticale.</summary>
        public int RowSpan = 1;
        /// <summary>Suite d'une fusion verticale : la cellule qui la commence (son contenu n'est pas affiché).</summary>
        public CellPart? Restart;
        /// <summary>Dernière suite d'une fusion verticale (elle porte le trait du bas).</summary>
        public bool EndsMerge;
        /// <summary>Hauteur du contenu, marges haute et basse comprises.</summary>
        public double ContentHeight;
        /// <summary>Lignes des paragraphes : abscisse sur la page, ordonnée depuis le haut de la rangée.</summary>
        public List<(ParagraphLayout L, int Line, double X, double Y)> Lines { get; } = new();
        /// <summary>Cellules des tableaux imbriqués (ordonnée depuis le haut de la rangée).</summary>
        public List<CellBox> Boxes { get; } = new();
        public bool Merged => RowSpan > 1;
    }

    /// <summary>Rangée mesurée : hauteur du contenu, bordures horizontales comptées dans la hauteur, cellules.</summary>
    sealed class RowBox
    {
        public required TableRow Row { get; init; }
        public required Table Table { get; init; }
        public int Index;
        public double Content, TopBorder, BottomBorder;
        public double Height => Content + TopBorder + BottomBorder;
        public List<CellPart> Cells { get; } = new();
    }

    sealed class TableBox
    {
        public required Table Table { get; init; }
        public List<RowBox> Rows { get; } = new();
        public double X, Width;
    }

    const double DefaultCellMargin = 5.4;     // marges gauche et droite des cellules de « Tableau Normal » (108 twips)

    static double Twips(string? v) => double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? d / 20 : 0;

    static int? IntVal(XElement? e, XNamespace W, string attr = "val") =>
        int.TryParse(e?.Attribute(W + attr)?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : null;

    /// <summary>Largeur préférée (w:tblW, w:tcW) en points ; pourcentage rapporté à <paramref name="whole"/> ; null = automatique.</summary>
    static double? Preferred(XElement? e, XNamespace W, double whole)
    {
        if (e is null) return null;
        var type = e.Attribute(W + "type")?.Value ?? "dxa";
        var v = e.Attribute(W + "w")?.Value;
        if (v is null || type is "auto" or "nil") return null;
        if (type == "pct")
        {
            // fiftieths of a percent, ou « 50% » (format strict)
            if (v.EndsWith('%') && double.TryParse(v[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out var pc)) return whole * pc / 100;
            return double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out var f) && f > 0 ? whole * f / 5000 : null;
        }
        double w = Twips(v);
        return w > 0 ? w : null;
    }

    static XElement? SideOf(XElement? borders, XNamespace W, string side) =>
        borders?.Element(W + side) ?? side switch
        {
            "left" => borders?.Element(W + "start"),
            "right" => borders?.Element(W + "end"),
            _ => null,
        };

    static BorderLine LineOf(XElement? b, XNamespace W)
    {
        if (b is null || b.Attribute(W + "val")?.Value is null or "none" or "nil") return default;
        double w = double.TryParse(b.Attribute(W + "sz")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? Math.Max(0.25, v / 8) : 0.5;
        var color = b.Attribute(W + "color")?.Value;
        return new BorderLine(w, color is { Length: 6 } && !color.Equals("auto", StringComparison.OrdinalIgnoreCase) ? color.ToUpperInvariant() : null, b.Attribute(W + "val")!.Value);
    }

    static string? FillOf(XElement? shd, XNamespace W)
    {
        if (shd is null) return null;
        var val = shd.Attribute(W + "val")?.Value;
        var fill = val is "solid" ? shd.Attribute(W + "color")?.Value : shd.Attribute(W + "fill")?.Value;
        return fill is { Length: 6 } && !fill.Equals("auto", StringComparison.OrdinalIgnoreCase) ? fill.ToUpperInvariant() : null;
    }

    /// <summary>Fusionne deux jeux de côtés (w:tblBorders, w:tblCellMar…) : ceux de <paramref name="over"/> l'emportent.</summary>
    static XElement? MergeSides(XElement? under, XElement? over)
    {
        if (under is null) return over;
        if (over is null) return under;
        var acc = new XElement(under);
        foreach (var e in over.Elements())
        {
            acc.Element(e.Name)?.Remove();
            acc.Add(new XElement(e));
        }
        return acc;
    }

    /// <summary>Mesure un tableau posé à <paramref name="left"/> dans une colonne de largeur <paramref name="width"/>.</summary>
    TableBox MeasureTable(Table t, double left, double width)
    {
        var W = t.Props.Ns;
        var styles = _lines.Styles;
        // pas de w:tblStyle : style de tableau par défaut (« Tableau Normal »)
        var styleId = styles.TableStyle(t.Props.Val("tblStyle") ?? "");
        var stylePr = styles.TableStyleProps(styleId);
        XElement? Prop(string name) => t.Props.Get(name) ?? stylePr?.Element(W + name);
        var look = TableStyles.Look(t.Props.Get("tblLook"));
        int rowBand = IntVal(Prop("tblStyleRowBandSize"), W) ?? 1, colBand = IntVal(Prop("tblStyleColBandSize"), W) ?? 1;
        var tableMargins = MergeSides(stylePr?.Element(W + "tblCellMar"), t.Props.Get("tblCellMar"));
        var rows = t.Items.OfType<TableRow>().ToList();
        var box = new TableBox { Table = t };

        // ---- cellules : colonne de la grille, régions, propriétés du style
        var cellsOf = new List<List<CellPart>>();
        var levels = new Dictionary<CellPart, (CellStyle Style, XElement? TcPr)>();
        int gridCount = 0;
        for (int r = 0; r < rows.Count; r++)
        {
            var row = rows[r];
            var cells = row.Items.OfType<TableCell>().ToList();
            var parts = new List<CellPart>();
            int col = IntVal(row.Props.Get("gridBefore"), W) ?? 0;
            for (int ci = 0; ci < cells.Count; ci++)
            {
                int span = Math.Max(1, IntVal(cells[ci].Props.Get("gridSpan"), W) ?? 1);
                var part = new CellPart { Cell = cells[ci], Index = ci, Col = col, Span = span };
                var cs = new CellStyle(styleId, TableStyles.Regions(look, r, rows.Count, ci, cells.Count, rowBand, colBand));
                levels[part] = (cs, styles.TableLevel(cs).TcPr);
                parts.Add(part);
                col += span;
            }
            gridCount = Math.Max(gridCount, col + (IntVal(row.Props.Get("gridAfter"), W) ?? 0));
            cellsOf.Add(parts);
        }

        double Margin(CellPart p, string side, double def)
        {
            var own = p.Cell.Props.Get("tcMar");
            var cond = levels[p].TcPr?.Element(W + "tcMar");
            var e = SideOf(own, W, side) ?? SideOf(cond, W, side) ?? SideOf(tableMargins, W, side);
            if (e is null) return def;
            return e.Attribute(W + "type")?.Value == "nil" ? 0 : Twips(e.Attribute(W + "w")?.Value);
        }

        // ---- grille
        var grid = t.Grid?.Elements(W + "gridCol").Select(g => Twips(g.Attribute(W + "w")?.Value)).ToList() ?? new List<double>();
        while (grid.Count < gridCount) grid.Add(0);
        if (grid.Count == 0) grid.Add(width);
        if (grid.Any(g => g <= 0))
        {
            // colonnes sans largeur : w:tcW des cellules d'une seule colonne, sinon la place restante partagée
            foreach (var parts in cellsOf)
                foreach (var p in parts)
                    if (p.Span == 1 && p.Col < grid.Count && grid[p.Col] <= 0 && Preferred(p.Cell.Props.Get("tcW"), W, width) is { } pw)
                        grid[p.Col] = pw;
            int missing = grid.Count(g => g <= 0);
            if (missing > 0)
            {
                double rest = Math.Max(missing * 12, width - grid.Where(g => g > 0).Sum());
                for (int k = 0; k < grid.Count; k++) if (grid[k] <= 0) grid[k] = rest / missing;
            }
        }
        double? tablePref = Preferred(Prop("tblW"), W, width);
        bool fixedLayout = Prop("tblLayout")?.Attribute(W + "type")?.Value == "fixed";
        // « ajuster au contenu » de Word : w:tcW explicitement automatiques (une grille seule, sans w:tcW, est gardée)
        bool allAuto = cellsOf.SelectMany(x => x).All(p => p.Cell.Props.Get("tcW") is { } tcw && Preferred(tcw, W, width) is null);
        if (!fixedLayout && tablePref is null && allAuto && cellsOf.Count > 0)
            grid = AutoFit(cellsOf, levels, grid, width, Margin);
        else if (tablePref is { } tw && grid.Sum() > 0 && Math.Abs(grid.Sum() - tw) > 1)
        {
            double k = tw / grid.Sum();
            for (int i = 0; i < grid.Count; i++) grid[i] *= k;
        }
        var colX = new double[grid.Count + 1];
        for (int i = 0; i < grid.Count; i++) colX[i + 1] = colX[i] + grid[i];
        double tableWidth = colX[^1];

        // ---- position du tableau
        double indent = Twips(Prop("tblInd")?.Attribute(W + "w")?.Value);
        var jc = Prop("jc")?.Attribute(W + "val")?.Value ?? rows.FirstOrDefault()?.Props.Get("jc")?.Attribute(W + "val")?.Value ?? "left";
        double firstMargin = cellsOf.FirstOrDefault()?.FirstOrDefault() is { } fc ? Margin(fc, "left", DefaultCellMargin) : DefaultCellMargin;
        double x0 = jc switch
        {
            "center" => left + (width - tableWidth) / 2,
            "right" or "end" => left + width - tableWidth,
            // Word 2013 et suivants : retrait mesuré jusqu'au bord du tableau ; avant, jusqu'au texte de la première cellule
            _ => left + indent - (styles.Sheet.CompatibilityMode >= 15 ? 0 : firstMargin),
        };
        box.X = x0;
        box.Width = tableWidth;

        // ---- fusions verticales : cellule de départ et nombre de rangées
        for (int r = 0; r < rows.Count; r++)
            foreach (var p in cellsOf[r])
            {
                var vm = p.Cell.Props.Get("vMerge");
                if (vm is null || vm.Attribute(W + "val")?.Value != "restart") continue;
                int span = 1;
                CellPart? lastPart = null;
                for (int k = r + 1; k < rows.Count; k++)
                {
                    var next = cellsOf[k].FirstOrDefault(q => q.Col == p.Col);
                    if (next?.Cell.Props.Get("vMerge") is not { } nv || nv.Attribute(W + "val")?.Value is not (null or "continue")) break;
                    next.Restart = p;
                    lastPart = next;
                    span++;
                }
                if (lastPart is not null) lastPart.EndsMerge = true;
                p.RowSpan = span;
            }

        // ---- géométrie, fond, bordures et contenu des cellules
        for (int r = 0; r < rows.Count; r++)
        {
            var row = rows[r];
            var parts = cellsOf[r];
            var rb = new RowBox { Row = row, Table = t, Index = r };
            // marge du haut : la plus grande des cellules de la rangée, pour toutes (Word aligne le début du texte des
            // cellules d'une rangée ; mesuré : w:tcMar haut de 14 pt sur une cellule, texte des autres descendu aussi)
            double rowTop = parts.Where(p => p.Restart is null).Select(p => Margin(p, "top", 0)).DefaultIfEmpty(0).Max();
            for (int ci = 0; ci < parts.Count; ci++)
            {
                var p = parts[ci];
                var (cs, tcLevel) = levels[p];
                int c1 = Math.Min(p.Col, grid.Count), c2 = Math.Min(p.Col + p.Span, grid.Count);
                p.X = x0 + colX[c1];
                p.W = colX[c2] - colX[c1];
                int lastRow = r + p.RowSpan - 1;
                var own = p.Cell.Props.Get("tcBorders");
                var cond = tcLevel?.Element(W + "tcBorders");
                BorderLine Side(string side, string tableSide, bool edge)
                {
                    // cellule > exceptions de rangée > tableau > mise en forme conditionnelle du style > style ; côté du
                    // bord du tableau ou intérieur (w:insideH / w:insideV). Dans une région conditionnelle, les côtés
                    // haut et bas sont ceux de la région (une rangée d'en-tête a son propre trait du bas)
                    var condSide = side is "top" or "bottom" ? SideOf(cond, W, side) ?? SideOf(cond, W, "insideH")
                        : edge ? SideOf(cond, W, side) : SideOf(cond, W, "insideV") ?? SideOf(cond, W, side);
                    var e = SideOf(own, W, side)
                        ?? SideOf(row.PropsException?.Get("tblBorders"), W, tableSide)
                        ?? SideOf(t.Props.Get("tblBorders"), W, tableSide)
                        ?? condSide
                        ?? SideOf(stylePr?.Element(W + "tblBorders"), W, tableSide);
                    return LineOf(e, W);
                }
                p.Top = Side("top", r == 0 ? "top" : "insideH", r == 0);
                p.Bottom = Side("bottom", lastRow >= rows.Count - 1 ? "bottom" : "insideH", lastRow >= rows.Count - 1);
                p.Left = Side("left", ci == 0 ? "left" : "insideV", ci == 0);
                p.Right = Side("right", ci == parts.Count - 1 ? "right" : "insideV", ci == parts.Count - 1);
                p.Fill = FillOf(p.Cell.Props.Get("shd"), W) ?? FillOf(tcLevel?.Element(W + "shd"), W);
                p.VAlign = p.Cell.Props.Val("vAlign") ?? tcLevel?.Element(W + "vAlign")?.Attribute(W + "val")?.Value ?? "top";
                if (p.Restart is not null) { rb.Cells.Add(p); continue; }
                double ml = Margin(p, "left", DefaultCellMargin), mr = Margin(p, "right", DefaultCellMargin);
                double mb = Margin(p, "bottom", 0);
                double bottom = StackCell(p, p.Cell.Content, p.X + ml, Math.Max(1, p.W - ml - mr), cs, rowTop);
                p.ContentHeight = bottom + mb;
                rb.Cells.Add(p);
            }
            box.Rows.Add(rb);
        }

        // ---- hauteurs : contenu des cellules non fusionnées, w:trHeight, puis fusions verticales
        for (int r = 0; r < box.Rows.Count; r++)
        {
            var rb = box.Rows[r];
            double h = 0;
            foreach (var p in rb.Cells)
                if (p.Restart is null && !p.Merged) h = Math.Max(h, p.ContentHeight);
            if (rb.Row.Props.Get("trHeight") is { } th)
            {
                double v = Twips(th.Attribute(W + "val")?.Value);
                h = th.Attribute(W + "hRule")?.Value == "exact" ? v : Math.Max(h, v);
            }
            rb.Content = h;
            // bordures horizontales comptées dans la hauteur : celle du bas de chaque rangée, et pour une bordure partagée
            // entre deux rangées, la plus épaisse des deux (Syllabus : rangée à bordure basse de 3 pt plus haute de
            // 2,8 pt dans Word, filet de 0,5 pt : 0,7 pt)
            double top = rb.Cells.Count == 0 ? 0 : rb.Cells.Max(p => p.Restart is null ? p.Top.Thickness : 0);
            rb.TopBorder = r == 0 ? top : Math.Max(0, top - box.Rows[r - 1].BottomBorder);
            rb.BottomBorder = rb.Cells.Count == 0 ? 0 : rb.Cells.Max(p => p.Merged ? 0 : p.Bottom.Thickness);
        }
        foreach (var rb in box.Rows)
            foreach (var p in rb.Cells.Where(p => p.Merged))
            {
                var spanned = box.Rows.Skip(rb.Index).Take(p.RowSpan).ToList();
                double have = spanned.Sum(x => x.Height) - spanned[0].TopBorder - spanned[^1].BottomBorder;
                var last = spanned[^1];
                bool exact = last.Row.Props.Get("trHeight")?.Attribute(W + "hRule")?.Value == "exact";
                if (p.ContentHeight > have + 0.01 && !exact) last.Content += p.ContentHeight - have;
            }
        // ---- alignement vertical
        foreach (var rb in box.Rows)
            foreach (var p in rb.Cells)
            {
                if (p.Restart is not null || p.VAlign is not ("center" or "bottom")) continue;
                var spanned = box.Rows.Skip(rb.Index).Take(p.RowSpan).ToList();
                double avail = spanned.Sum(x => x.Height) - spanned[0].TopBorder - spanned[^1].BottomBorder;
                double dy = avail - p.ContentHeight;
                if (p.VAlign == "center") dy /= 2;
                if (dy <= 0.01) continue;
                for (int k = 0; k < p.Lines.Count; k++) p.Lines[k] = p.Lines[k] with { Y = p.Lines[k].Y + dy };
                for (int k = 0; k < p.Boxes.Count; k++) p.Boxes[k] = p.Boxes[k] with { Y = p.Boxes[k].Y + dy };
            }
        return box;
    }

    /// <summary>
    /// Tableau ajusté au contenu : chaque colonne prend la largeur de sa plus longue ligne ; si le tout dépasse la place,
    /// les colonnes qui tiennent dans leur part prennent leur largeur et les autres se partagent le reste (au prorata de
    /// leur largeur maximale, sans descendre sous le plus long mot). Calibré sur des tableaux « Ajuster au contenu » de Word.
    /// </summary>
    List<double> AutoFit(List<List<CellPart>> cellsOf, Dictionary<CellPart, (CellStyle Style, XElement? TcPr)> levels, List<double> grid, double avail,
        Func<CellPart, string, double, double> margin)
    {
        int n = grid.Count;
        var max = new double[n];
        var min = new double[n];
        foreach (var parts in cellsOf)
            foreach (var p in parts)
            {
                if (p.Span != 1 || p.Col >= n) continue;
                double pad = margin(p, "left", DefaultCellMargin) + margin(p, "right", DefaultCellMargin);
                double mi = 0, ma = 0;
                foreach (var para in p.Cell.Content.Paragraphs())
                {
                    var (a, b) = _lines.ContentWidths(para, levels[p].Style);
                    mi = Math.Max(mi, a);
                    ma = Math.Max(ma, b);
                }
                min[p.Col] = Math.Max(min[p.Col], mi + pad);
                max[p.Col] = Math.Max(max[p.Col], ma + pad);
            }
        for (int k = 0; k < n; k++)
        {
            if (max[k] <= 0) max[k] = grid[k] > 0 ? grid[k] : 2 * DefaultCellMargin;
            min[k] = Math.Min(min[k] <= 0 ? max[k] : min[k], max[k]);
        }
        if (max.Sum() <= avail) return max.ToList();
        var res = new double[n];
        var open = Enumerable.Range(0, n).ToList();
        double remaining = avail;
        while (open.Count > 0)
        {
            double fair = remaining / open.Count;
            var small = open.Where(k => max[k] <= fair).ToList();
            if (small.Count == 0) break;
            foreach (var k in small) { res[k] = max[k]; remaining -= max[k]; open.Remove(k); }
        }
        if (open.Count > 0)
        {
            double total = open.Sum(k => max[k]);
            foreach (var k in open) res[k] = Math.Max(min[k], remaining * max[k] / total);
        }
        return res.ToList();
    }

    /// <summary>Empile le contenu d'une cellule (paragraphes, tableaux imbriqués) à partir de <paramref name="top"/> ; renvoie le bas.</summary>
    double StackCell(CellPart part, Story content, double x, double width, CellStyle cs, double top)
    {
        double cy = top;
        ParagraphLayout? prev = null;
        for (int bi = 0; bi < content.Blocks.Count; bi++)
        {
            var b = content.Blocks[bi];
            switch (b)
            {
                case Paragraph para:
                    var pl = LayoutOf(para, width, cs);
                    if (pl.Content.Hidden) break;
                    // marque de fin de cellule vide après un tableau imbriqué : Word ne lui donne pas de hauteur
                    // (mesuré : le paragraphe suivant le tableau extérieur vient juste sous le tableau imbriqué)
                    if (bi == content.Blocks.Count - 1 && bi > 0 && content.Blocks[bi - 1] is Table && pl.Content.Text.Length == 0)
                    {
                        for (int l = 0; l < pl.Lines.Count; l++) part.Lines.Add((pl, l, x, cy));
                        break;
                    }
                    double before = pl.SpaceBefore;
                    if (prev is not null && pl.Format.ContextualSpacing && prev.Format.StyleId == pl.Format.StyleId) before = 0;
                    cy += before;
                    for (int l = 0; l < pl.Lines.Count; l++)
                    {
                        part.Lines.Add((pl, l, x, cy));
                        cy += pl.Lines[l].Height;
                    }
                    cy += pl.SpaceAfter;
                    prev = pl;
                    if (Trace && pl.Lines.Count > 1)
                        Console.WriteLine($"    cellule {width:0.0} pt : « {pl.Content.Text[..Math.Min(40, pl.Content.Text.Length)]} » {pl.Lines.Count} lignes : "
                            + string.Join(" | ", pl.Lines.Select(ln => pl.Content.Text[ln.Start..ln.End])));
                    break;
                case Table nested:
                    var nb = MeasureTable(nested, x, width);
                    foreach (var row in nb.Rows)
                    {
                        EmitRow(row, cy, row.Height, (l, i, lx, ly) => part.Lines.Add((l, i, lx, ly)), cb => part.Boxes.Add(cb with { Nested = true }));
                        cy += row.Height;
                    }
                    prev = null;
                    break;
                case BlockContentControl sdt:
                    cy = StackCell(part, sdt.Content, x, width, cs, cy);
                    prev = null;
                    break;
            }
        }
        return cy;
    }

    /// <summary>Cellule affichée d'une rangée posée à <paramref name="y"/> sur une hauteur <paramref name="h"/>.</summary>
    static CellBox BoxOf(RowBox row, CellPart p, double y, double h)
    {
        // fusion verticale : la cellule de départ et ses suites forment un seul rectangle (fond et côtés de la
        // cellule de départ, pas de trait entre elles)
        var look = p.Restart ?? p;
        var bottom = p.Merged || p.Restart is not null && !p.EndsMerge ? default : p.Bottom;
        return new CellBox(p.X, y, p.W, h, look.Fill, p.Restart is not null ? default : p.Top, bottom, look.Left, look.Right)
        {
            Cell = (p.Restart ?? p).Cell, Table = row.Table, Row = row.Index, Column = p.Col,
        };
    }

    /// <summary>Pose une rangée entière : cellules (et celles des tableaux imbriqués) puis lignes des paragraphes.</summary>
    static void EmitRow(RowBox row, double y, double h, Action<ParagraphLayout, int, double, double> line, Action<CellBox> cell)
    {
        foreach (var p in row.Cells)
        {
            cell(BoxOf(row, p, y, h));
            foreach (var b in p.Boxes) cell(b with { Y = y + b.Y });
        }
        foreach (var p in row.Cells)
            foreach (var (l, i, lx, ly) in p.Lines) line(l, i, lx, y + ly);
    }

    static void PlaceRowAt(RowBox row, double y, Page page) =>
        EmitRow(row, y, row.Height, (l, i, x, ly) => page.Lines.Add(new PlacedLine(l, i, x, ly)), page.Cells.Add);

    /// <summary>Tableau : rangées entières, une rangée trop haute pour la place restante est coupée entre ses lignes ou passe à la page suivante.</summary>
    int? PlaceTable(Table t, int i, int startRow, double startOffset, Cursor c, PageBreaker newPage)
    {
        var box = MeasureTable(t, c.ColX, c.ColW);
        var rows = box.Rows;
        // tableau flottant (w:tblpPr) : il ne prend pas de place dans le flux ; ancré en bas de la marge, il réduit
        // la hauteur du corps de la page (Syllabus, page de garde) ; les autres positions viennent avec le lot 7
        if (t.Props.Get("tblpPr") is { } pos)
        {
            var W = t.Props.Ns;
            double height = rows.Sum(r => r.Height);
            if (pos.Attribute(W + "tblpYSpec")?.Value == "bottom" && (pos.Attribute(W + "vertAnchor")?.Value ?? "margin") == "margin")
            {
                double y = c.Page.ContentBottom - height;
                foreach (var r in rows)
                {
                    PlaceRowAt(r, y, c.Page);
                    y += r.Height;
                }
                c.Page.ReservedBottom += height + Twips(pos.Attribute(W + "bottomFromText")?.Value);
            }
            return null;
        }
        // rangées d'en-tête (w:tblHeader en tête du tableau) : répétées en haut de chaque page où le tableau continue
        int headers = 0;
        while (headers < rows.Count && rows[headers].Row.Props.Get("tblHeader") is { } th && th.Attribute(t.Props.Ns + "val")?.Value is null or "1" or "true" or "on") headers++;
        void RepeatHeaders()
        {
            for (int h = 0; h < headers && h < rows.Count; h++)
            {
                PlaceRowAt(rows[h], c.Y, c.Page);
                c.Y += rows[h].Height;
            }
        }
        if (startRow > 0 || startOffset > 0)
            if (startRow >= headers) RepeatHeaders();
        double offset = startOffset;
        for (int r = startRow; r < rows.Count; r++, offset = 0)
        {
            var box1 = rows[r];
            while (true)
            {
                double remaining = RowPartHeight(box1, offset, null);
                if (c.Y + remaining <= c.Page.ContentBottom + 0.01 || c.Empty && !CanSplit(box1.Row, t))
                {
                    PlaceRowPart(box1, offset, null, c.Y, c.Page);
                    c.Y += remaining;
                    c.Empty = false;
                    break;
                }
                // rangée trop haute pour la place restante : coupée entre les lignes de ses cellules (sauf w:cantSplit),
                // sinon renvoyée entière à la page suivante
                double cut = AdjustCut(box1, offset, offset + (c.Page.ContentBottom - c.Y));
                if (CanSplit(box1.Row, t) && r >= headers && HasLinesBetween(box1, offset, cut))
                {
                    double placed = RowPartHeight(box1, offset, cut);
                    PlaceRowPart(box1, offset, cut, c.Y, c.Page);
                    c.Y += placed;
                    c.Empty = false;
                    if (newPage(i, r, fromBreak: false, cut) is { } k) return k;
                    RepeatHeaders();
                    offset = cut;
                    continue;
                }
                if (c.Empty)
                {
                    // ni coupable ni entière : placée telle quelle (elle dépasse le bas de la page)
                    PlaceRowPart(box1, offset, null, c.Y, c.Page);
                    c.Y += remaining;
                    c.Empty = false;
                    break;
                }
                if (newPage(i, r, fromBreak: false, offset) is { } k2) return k2;
                RepeatHeaders();
            }
        }
        c.Prev = null;
        return null;
    }

    static bool CanSplit(TableRow row, Table t) =>
        (row.Props.Get("cantSplit") is not { } cs || cs.Attribute(t.Props.Ns + "val")?.Value is "0" or "false" or "off")
        // une rangée à hauteur minimale (w:trHeight) ne se coupe pas (Syllabus : rangée de 146 pt envoyée entière à la
        // page suivante par Word alors que 120 pt restaient ; écart ramené de -137 à -3 pt sur la page suivante)
        && row.Props.Get("trHeight") is null;

    /// <summary>
    /// Coupure d'une rangée entre deux pages, veuves et orphelines comprises dans les paragraphes des cellules (Word :
    /// une rangée dont il ne resterait qu'une ligne d'un paragraphe en bas de page passe entière à la page suivante).
    /// </summary>
    static double AdjustCut(RowBox box, double from, double cut)
    {
        for (int pass = 0; pass < 8; pass++)
        {
            double before = cut;
            foreach (var p in box.Cells)
            {
                if (p.Merged || p.Restart is not null) continue;
                foreach (var g in p.Lines.GroupBy(l => l.L))
                {
                    if (!g.Key.Format.WidowControl) continue;
                    var rest = g.Where(l => l.Y + l.L.Lines[l.Line].Height > from + 0.01).ToList();
                    int n = rest.Count;
                    if (n < 2) continue;
                    int placed = rest.Count(l => l.Y + l.L.Lines[l.Line].Height <= cut + 0.01);
                    if (placed == 0 || placed == n) continue;
                    if (n - placed == 1) cut = Math.Min(cut, rest[placed - 1].Y);                // veuve : une ligne de plus descend
                    else if (placed == 1 && rest[0].Line == 0) cut = Math.Min(cut, rest[0].Y);   // orpheline : tout descend
                }
            }
            if (cut >= before - 0.001) break;
        }
        return cut;
    }

    /// <summary>Lignes des cellules d'une rangée qui comptent pour sa coupure (hors cellules fusionnées verticalement).</summary>
    static IEnumerable<(int Cell, ParagraphLayout L, int Line, double X, double Top, double Bottom)> RowLines(RowBox box)
    {
        for (int ci = 0; ci < box.Cells.Count; ci++)
        {
            var p = box.Cells[ci];
            if (p.Merged) continue;
            foreach (var (l, i, x, y) in p.Lines) yield return (ci, l, i, x, y, y + l.Lines[i].Height);
        }
    }

    /// <summary>Haut de la première ligne restante de chaque cellule (lignes dont le bas dépasse <paramref name="from"/>).</summary>
    static double[] CellShift(RowBox box, double from)
    {
        var shift = Enumerable.Repeat(double.NaN, box.Cells.Count).ToArray();
        foreach (var x in RowLines(box))
            if (x.Bottom > from + 0.01 && double.IsNaN(shift[x.Cell])) shift[x.Cell] = x.Top;
        return shift;
    }

    static bool HasLinesBetween(RowBox box, double from, double to) =>
        RowLines(box).Any(x => x.Bottom > from + 0.01 && x.Bottom <= to + 0.01);

    /// <summary>Hauteur de la partie [from, to[ de la rangée (to null : jusqu'au bout, bordure et hauteur minimale comprises).</summary>
    static double RowPartHeight(RowBox box, double from, double? to)
    {
        if (from <= 0 && to is null) return box.Height;
        var shift = CellShift(box, from);
        double h = 0;
        foreach (var x in RowLines(box))
        {
            if (x.Bottom <= from + 0.01 || to is { } t && x.Bottom > t + 0.01) continue;
            h = Math.Max(h, x.Bottom - shift[x.Cell]);
        }
        if (to is null) h += box.Height - box.Content;     // bordures de la rangée, à la fin de la dernière partie
        return h;
    }

    static void PlaceRowPart(RowBox box, double from, double? to, double y, Page page)
    {
        if (from <= 0 && to is null) { PlaceRowAt(box, y, page); return; }
        var shift = CellShift(box, from);
        double h = RowPartHeight(box, from, to);
        for (int ci = 0; ci < box.Cells.Count; ci++)
        {
            var p = box.Cells[ci];
            page.Cells.Add(BoxOf(box, p, y, h));
            double s = double.IsNaN(shift[ci]) ? from : shift[ci];
            foreach (var b in p.Boxes)
                if (b.Y >= from - 0.01 && (to is not { } tb || b.Y < tb)) page.Cells.Add(b with { Y = y + b.Y - s });
            // cellule fusionnée verticalement : son contenu est posé avec la première partie de la rangée
            if (p.Merged)
            {
                if (from <= 0) foreach (var (l, i, x, ly) in p.Lines) page.Lines.Add(new PlacedLine(l, i, x, y + ly));
                continue;
            }
        }
        foreach (var x in RowLines(box))
        {
            if (x.Bottom <= from + 0.01 || to is { } t && x.Bottom > t + 0.01) continue;
            page.Lines.Add(new PlacedLine(x.L, x.Line, x.X, y + x.Top - shift[x.Cell]));
        }
    }
}
