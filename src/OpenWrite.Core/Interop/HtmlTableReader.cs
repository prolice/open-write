using System.Globalization;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;

namespace OpenWrite.Core.Interop;

/// <summary>
/// Tableaux du « HTML Format » du presse-papiers (Excel, pages web, autres traitements de texte) en vrais tableaux,
/// comme Word quand il colle depuis Excel (calibré sur son collage par COM) : largeurs des colonnes (&lt;col&gt;),
/// hauteurs des rangées (au moins), bordures par côté de cellule (0,5 pt → w:sz 4, 1 pt → 8 ; windowtext → auto),
/// trame, marges des cellules (padding, cellpadding), alignements horizontal et vertical, pas de retour à la ligne,
/// fusions (colspan, rowspan), mise en forme des caractères. Le CSS vient des règles &lt;style&gt; (td, th, table, tr,
/// .classe, td.classe) puis des styles en ligne, dans l'ordre de la cascade.
/// Excel place la balise &lt;table&gt; avant &lt;!--StartFragment--&gt; : le tableau est cherché dans tout le HTML.
/// </summary>
public static partial class HtmlTableReader
{
    /// <summary>
    /// Blocs collés (paragraphes et tableaux) si la partie sélectionnée contient au moins un tableau, null sinon (le
    /// lecteur HTML ordinaire s'en charge). <paramref name="textWidthPt"/> : largeur répartie entre les colonnes quand
    /// la source n'en donne pas.
    /// </summary>
    public static Fragment? Read(string cfHtml, XNamespace w, double textWidthPt, Func<string> paraId)
    {
        var css = Css(cfHtml);
        var fragment = HtmlReader.Fragment(cfHtml);
        var blocks = new List<Block>();
        if (!Regex.IsMatch(fragment, @"<table\b", RegexOptions.IgnoreCase))
        {
            // fragment entièrement dans un tableau ouvert avant lui (Excel) : on reprend ce tableau
            if (!Regex.IsMatch(fragment, @"<t[dr]\b", RegexOptions.IgnoreCase)) return null;
            int s = cfHtml.IndexOf("<!--StartFragment", StringComparison.OrdinalIgnoreCase);
            int open = s < 0 ? -1 : LastTableBefore(cfHtml, s);
            if (open < 0) return null;
            if (ParseTable(cfHtml, open, css, w, textWidthPt, paraId, out _) is not { } only) return null;
            return new Fragment(new List<Block> { only }, false);
        }

        // paragraphs et tableaux dans l'ordre du fragment : rien de ce qui est autour d'un tableau n'est perdu
        int pos = 0;
        bool lastMark = false;
        while (pos < fragment.Length)
        {
            var m = Regex.Match(fragment[pos..], @"<table\b", RegexOptions.IgnoreCase);
            int start = m.Success ? pos + m.Index : fragment.Length;
            if (start > pos)
            {
                var paras = HtmlReader.Read(fragment[pos..start], out var mark);
                var built = Fragments.Build(paras, mark, w);
                blocks.AddRange(built.Blocks);
                lastMark = mark;
            }
            if (!m.Success) break;
            if (ParseTable(fragment, start, css, w, textWidthPt, paraId, out int end) is { } table)
            {
                blocks.Add(table);
                lastMark = false;
            }
            pos = Math.Max(end, start + 6);
        }
        return blocks.Any(b => b is Table) ? new Fragment(blocks, lastMark) : null;
    }

    /// <summary>Dernière balise &lt;table&gt; ouverte avant <paramref name="index"/> et pas encore fermée.</summary>
    static int LastTableBefore(string html, int index)
    {
        var stack = new Stack<int>();
        foreach (Match m in Regex.Matches(html[..index], @"<(/?)table\b", RegexOptions.IgnoreCase))
        {
            if (m.Groups[1].Value == "/") { if (stack.Count > 0) stack.Pop(); }
            else stack.Push(m.Index);
        }
        return stack.Count > 0 ? stack.Peek() : -1;
    }

    // ================================================================ CSS

    /// <summary>Règles CSS des balises &lt;style&gt; : sélecteur simple (en minuscules) → déclarations, dans l'ordre.</summary>
    static Dictionary<string, List<(string Prop, string Value)>> Css(string html)
    {
        var rules = new Dictionary<string, List<(string, string)>>();
        foreach (Match st in Regex.Matches(html, @"<style\b[^>]*>(.*?)</style>", RegexOptions.Singleline | RegexOptions.IgnoreCase))
        {
            var text = Regex.Replace(st.Groups[1].Value, @"/\*.*?\*/|<!--|-->", " ", RegexOptions.Singleline);
            foreach (Match r in Regex.Matches(text, @"([^{}]+)\{([^{}]*)\}"))
            {
                var decls = Declarations(r.Groups[2].Value);
                foreach (var sel in r.Groups[1].Value.Split(','))
                {
                    var key = Regex.Replace(sel.Trim(), @"\s+", " ").ToLowerInvariant();
                    if (key.Length == 0 || key.StartsWith('@')) continue;
                    if (!rules.TryGetValue(key, out var list)) rules[key] = list = new();
                    list.AddRange(decls);
                }
            }
        }
        return rules;
    }

    static List<(string Prop, string Value)> Declarations(string style)
    {
        var list = new List<(string, string)>();
        foreach (var decl in style.Split(';'))
        {
            int k = decl.IndexOf(':');
            if (k <= 0) continue;
            var prop = decl[..k].Trim().ToLowerInvariant();
            var value = WebUtility.HtmlDecode(decl[(k + 1)..].Replace("!important", "").Trim());
            if (prop.Length > 0) list.Add((prop, value));
        }
        return list;
    }

    /// <summary>Style calculé d'un élément : règle de la balise, puis de ses classes, puis style en ligne.</summary>
    static Dictionary<string, string> Computed(Dictionary<string, List<(string Prop, string Value)>> css, string tag, string attrs, Dictionary<string, string>? inherited = null)
    {
        var style = inherited is null ? new Dictionary<string, string>() : new Dictionary<string, string>(inherited);
        void Apply(IEnumerable<(string Prop, string Value)> decls)
        {
            foreach (var (prop, value) in decls)
            {
                if (prop is "border" or "border-top" or "border-right" or "border-bottom" or "border-left")
                {
                    var sides = prop == "border" ? new[] { "top", "right", "bottom", "left" } : new[] { prop[7..] };
                    foreach (var side in sides) style["border-" + side] = value;
                }
                else if (prop is "padding")
                {
                    var parts = value.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length == 0) continue;
                    string P(int i) => parts[i < parts.Length ? i : i == 2 ? 0 : i == 3 ? (parts.Length > 1 ? 1 : 0) : 0];
                    style["padding-top"] = P(0); style["padding-right"] = P(1); style["padding-bottom"] = P(2); style["padding-left"] = P(3);
                }
                else style[prop] = value;
            }
        }
        if (css.TryGetValue(tag, out var t)) Apply(t);
        if (Attr(attrs, "class") is { } classes)
            foreach (var cls in classes.Split(' ', StringSplitOptions.RemoveEmptyEntries))
            {
                if (css.TryGetValue("." + cls.ToLowerInvariant(), out var c)) Apply(c);
                if (css.TryGetValue(tag + "." + cls.ToLowerInvariant(), out var tc)) Apply(tc);
            }
        if (Attr(attrs, "style") is { } inline) Apply(Declarations(inline));
        return style;
    }

    [GeneratedRegex(@"([a-zA-Z-]+)\s*=\s*(""[^""]*""|'[^']*'|[^\s>]+)")]
    private static partial Regex AttributeRegex();

    static string? Attr(string attrs, string name)
    {
        foreach (Match a in AttributeRegex().Matches(attrs))
            if (a.Groups[1].Value.Equals(name, StringComparison.OrdinalIgnoreCase))
                return WebUtility.HtmlDecode(a.Groups[2].Value.Trim('"', '\''));
        return null;
    }

    /// <summary>Longueur CSS ou attribut HTML en points (px = 0,75 pt ; un nombre seul = des pixels).</summary>
    static double? Length(string? v)
    {
        if (v is null) return null;
        var m = Regex.Match(v.Trim().ToLowerInvariant(), @"^(-?[\d.]+)\s*(pt|px|in|cm|mm|pc)?$");
        if (!m.Success || !double.TryParse(m.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var n)) return null;
        return m.Groups[2].Value switch
        {
            "pt" => n, "in" => n * 72, "cm" => n * 72 / 2.54, "mm" => n * 72 / 25.4, "pc" => n * 12, _ => n * 0.75,
        };
    }

    /// <summary>Bordure CSS (« .5pt solid windowtext ») en trait Word ; null si la cellule n'en dit rien.</summary>
    internal static TableOps.Line? Border(string? v)
    {
        if (v is null) return null;
        string? style = null, color = null;
        double? width = null;
        foreach (var tok in Regex.Matches(v, @"rgba?\([^)]*\)|\S+").Select(m => m.Value))
        {
            var t = tok.ToLowerInvariant();
            if (t is "none" or "hidden" or "solid" or "double" or "dotted" or "dashed" or "groove" or "ridge" or "inset" or "outset") style = t;
            else if (t is "thin") width = 0.75;
            else if (t is "medium") width = 2.25;
            else if (t is "thick") width = 3.75;
            else if (Length(t) is { } len) width = len;
            else color = t;
        }
        if (style is null or "none" or "hidden" || width is <= 0) return new TableOps.Line("nil");
        double pt = width ?? 0.75;
        var wordStyle = style switch
        {
            "double" => "double", "dotted" => "dotted", "dashed" => "dashed", "groove" => "threeDEngrave",
            "ridge" => "threeDEmboss", "inset" => "inset", "outset" => "outset", _ => "single",
        };
        // double : l'épaisseur CSS couvre les deux traits et l'espace ; Word donne celle d'un trait
        int size = (int)Math.Round(wordStyle == "double" ? pt * 8 / 3 : pt * 8);
        var col = color is null or "windowtext" or "black" or "auto" or "currentcolor" or "#000000" or "#000" ? "auto" : HtmlReader.Color(color) ?? "auto";
        return new TableOps.Line(wordStyle, Math.Clamp(size, 2, 96), col);
    }

    // ================================================================ tableau

    sealed class CellSpec
    {
        public required int Col, ColSpan, RowSpan;
        public required string Tag, Attrs, Html;
        public required Dictionary<string, string> Style;
    }

    /// <summary>Lit le tableau qui commence à <paramref name="start"/> (« &lt;table ») ; <paramref name="end"/> : après sa fin.</summary>
    static Table? ParseTable(string html, int start, Dictionary<string, List<(string Prop, string Value)>> css, XNamespace w,
        double textWidthPt, Func<string> paraId, out int end)
    {
        end = html.Length;
        var tokens = Regex.Matches(html[start..], @"<!--.*?-->|<(/?)([a-zA-Z][a-zA-Z0-9:]*)([^>]*)>", RegexOptions.Singleline);
        string tableAttrs = "";
        var tableStyle = new Dictionary<string, string>();
        var cols = new List<double?>();
        var rows = new List<(string Attrs, Dictionary<string, string> Style, List<CellSpec> Cells)>();
        var occupied = new List<HashSet<int>>();           // colonnes prises par des rowspan venus d'en haut
        int depth = 0;
        CellSpec? cell = null;
        int cellStart = 0;
        Dictionary<string, string>? rowStyle = null;

        void CloseCell(int at)
        {
            if (cell is null) return;
            cell.Html = html.Substring(start + cellStart, Math.Max(0, at - cellStart));
            cell = null;
        }
        HashSet<int> Occ(int r)
        {
            while (occupied.Count <= r) occupied.Add(new HashSet<int>());
            return occupied[r];
        }
        void OpenRow(string attrs)
        {
            rowStyle = Computed(css, "tr", attrs);
            rows.Add((attrs, rowStyle, new List<CellSpec>()));
        }

        foreach (Match t in tokens)
        {
            if (!t.Groups[2].Success) continue;
            bool closing = t.Groups[1].Value == "/";
            var tag = t.Groups[2].Value.ToLowerInvariant();
            var attrs = t.Groups[3].Value;
            if (tag == "table")
            {
                if (!closing)
                {
                    depth++;
                    if (depth == 1) { tableAttrs = attrs; tableStyle = Computed(css, "table", attrs); }
                    continue;
                }
                depth--;
                if (depth == 0)
                {
                    CloseCell(t.Index);
                    end = start + t.Index + t.Length;
                    break;
                }
                continue;
            }
            if (depth != 1) continue;
            switch (tag)
            {
                case "col" when !closing:
                {
                    int span = Math.Clamp(int.TryParse(Attr(attrs, "span"), out var sp) ? sp : 1, 1, 63);
                    var st = Computed(css, "col", attrs);
                    double? width = st.TryGetValue("width", out var wv) ? Length(wv) : null;
                    width ??= Length(Attr(attrs, "width"));
                    for (int i = 0; i < span; i++) cols.Add(width);
                    break;
                }
                case "tr" when !closing:
                    CloseCell(t.Index);
                    OpenRow(attrs);
                    break;
                case "tr":
                    CloseCell(t.Index);
                    rowStyle = null;
                    break;
                case "td" or "th" when !closing:
                {
                    CloseCell(t.Index);
                    if (rowStyle is null) OpenRow("");
                    int r = rows.Count - 1;
                    var row = rows[r];
                    int colSpan = Math.Clamp(int.TryParse(Attr(attrs, "colspan"), out var cs) ? cs : 1, 1, 63);
                    int rowSpan = Math.Clamp(int.TryParse(Attr(attrs, "rowspan"), out var rs) ? rs : 1, 1, 500);
                    int col = row.Cells.Count > 0 ? row.Cells[^1].Col + row.Cells[^1].ColSpan : 0;
                    while (Occ(r).Contains(col)) col++;
                    for (int k = 1; k < rowSpan; k++)
                        for (int c = col; c < col + colSpan; c++) Occ(r + k).Add(c);
                    cell = new CellSpec
                    {
                        Col = col, ColSpan = colSpan, RowSpan = rowSpan, Tag = tag, Attrs = attrs, Html = "",
                        Style = Computed(css, tag, attrs),
                    };
                    row.Cells.Add(cell);
                    cellStart = t.Index + t.Length;
                    break;
                }
                case "td" or "th":
                    CloseCell(t.Index);
                    break;
            }
        }
        CloseCell(Math.Max(cellStart, end - start));
        if (rows.Count == 0 || rows.All(r => r.Cells.Count == 0)) return null;
        // rowspan qui dépassent la dernière rangée : ramenés au tableau
        for (int r = 0; r < rows.Count; r++)
            foreach (var c in rows[r].Cells) c.RowSpan = Math.Min(c.RowSpan, rows.Count - r);

        int columns = Math.Min(63, Math.Max(rows.Max(r => r.Cells.Count == 0 ? 0 : r.Cells.Max(c => c.Col + c.ColSpan)),
            occupied.Take(rows.Count).Select(o => o.Count == 0 ? 0 : o.Max() + 1).DefaultIfEmpty(0).Max()));
        if (columns == 0) return null;

        // ---- largeurs : <col>, sinon cellules d'une seule colonne, sinon la largeur du texte partagée
        var widths = new double?[columns];
        for (int c = 0; c < columns && c < cols.Count; c++) widths[c] = cols[c];
        foreach (var r in rows)
            foreach (var c in r.Cells)
                if (c.ColSpan == 1 && c.Col < columns && widths[c.Col] is null)
                    widths[c.Col] = (c.Style.TryGetValue("width", out var wv) ? Length(wv) : null) ?? Length(Attr(c.Attrs, "width"));
        bool fromSource = widths.Any(x => x is > 0);
        double fallback = fromSource ? widths.Where(x => x is > 0).Average(x => x!.Value) : Math.Max(10, (textWidthPt - 0.5) / columns);
        var grid = widths.Select(x => (int)Math.Round((x is > 0 ? x.Value : fallback) * 20)).Select(x => Math.Max(x, 100)).ToArray();

        // ---- tableau
        var table = new Table { Props = new PropertyBag("tblPr", w) };
        table.Props.Set(new XElement(w + "tblW", new XAttribute(w + "w", fromSource ? grid.Sum() : 0), new XAttribute(w + "type", fromSource ? "dxa" : "auto")));
        if (tableStyle.TryGetValue("margin-left", out var ml) && Length(ml) is { } indent && indent > 0)
            table.Props.Set(new XElement(w + "tblInd", new XAttribute(w + "w", (int)Math.Round(indent * 20)), new XAttribute(w + "type", "dxa")));
        bool anyCssBorder = rows.Any(r => r.Cells.Any(c => c.Style.Keys.Any(k => k.StartsWith("border-", StringComparison.Ordinal))));
        // pages web sans CSS de bordure : la grille simple de Word, sauf border="0"
        if (!anyCssBorder && Attr(tableAttrs, "border") is not "0")
        {
            var line = new TableOps.Line("single", 4, "auto");
            table.Props.Set(new XElement(w + "tblBorders", new[] { "top", "left", "bottom", "right", "insideH", "insideV" }.Select(side => line.Element(w, side))));
        }
        if (Length(Attr(tableAttrs, "cellpadding")) is { } pad)
        {
            int p = (int)Math.Round(pad * 20);
            table.Props.Set(new XElement(w + "tblCellMar",
                new XElement(w + "left", new XAttribute(w + "w", p), new XAttribute(w + "type", "dxa")),
                new XElement(w + "right", new XAttribute(w + "w", p), new XAttribute(w + "type", "dxa"))));
        }
        table.Props.Set(TableOps.LookElement(w, TableLook.Default));
        table.Grid = new XElement(w + "tblGrid", grid.Select(x => new XElement(w + "gridCol", new XAttribute(w + "w", x))));

        var origins = new Dictionary<(int Row, int Col), TableCell>();
        for (int r = 0; r < rows.Count; r++)
        {
            var (rowAttrs, rStyle, cells) = rows[r];
            var row = new TableRow { Props = new PropertyBag("trPr", w) };
            double? height = (rStyle.TryGetValue("height", out var hv) ? Length(hv) : null) ?? Length(Attr(rowAttrs, "height"));
            height ??= cells.Select(c => (c.Style.TryGetValue("height", out var ch) ? Length(ch) : null) ?? Length(Attr(c.Attrs, "height"))).Where(x => x is > 0).DefaultIfEmpty(null).Max();
            if (height is > 0) row.Props.Set(new XElement(w + "trHeight", new XAttribute(w + "val", (int)Math.Round(height.Value * 20))));
            if (cells.Count > 0 && cells.All(c => c.Tag == "th") && r == 0) row.Props.SetFlag("tblHeader", true);

            // cellules de la rangée et suites de fusions verticales, dans l'ordre des colonnes
            var items = new List<(int Col, int Span, TableCell Cell)>();
            foreach (var c in cells)
            {
                if (c.Col >= columns) continue;
                int span = Math.Min(c.ColSpan, columns - c.Col);
                var tc = BuildCell(c, span, grid, w, paraId, css);
                if (c.RowSpan > 1) tc.Props.Set(new XElement(w + "vMerge", new XAttribute(w + "val", "restart")));
                items.Add((c.Col, span, tc));
                for (int k = 1; k < c.RowSpan; k++) origins[(r + k, c.Col)] = tc;
            }
            foreach (var ((rr, col), origin) in origins.Where(o => o.Key.Row == r).ToList())
            {
                var cont = new TableCell { Props = origin.Props.Clone() };
                cont.Props.Set(new XElement(w + "vMerge"));
                var mark = origin.Content.Paragraphs().FirstOrDefault();
                cont.Content.Blocks.Add(TableOps.EmptyParagraph(w, mark, paraId));
                int span = Math.Max(1, int.TryParse(origin.Props.Val("gridSpan"), out var gs) ? gs : 1);
                items.Add((col, span, cont));
            }
            items.Sort((a, b) => a.Col.CompareTo(b.Col));
            // trous (rangée plus courte que la grille) : cellules vides sans bordure
            int next = 0;
            foreach (var (col, span, tc) in items)
            {
                for (; next < col; next++) row.Items.Add(EmptyCell(w, grid[next], paraId));
                if (col < next) continue;                       // chevauchement (HTML incohérent) : ignoré
                row.Items.Add(tc);
                next = col + span;
            }
            for (; next < columns; next++) row.Items.Add(EmptyCell(w, grid[next], paraId));
            table.Items.Add(row);
        }
        return table;
    }

    static TableCell EmptyCell(XNamespace w, int width, Func<string> paraId)
    {
        var c = new TableCell { Props = new PropertyBag("tcPr", w) };
        c.Props.Set(new XElement(w + "tcW", new XAttribute(w + "w", width), new XAttribute(w + "type", "dxa")));
        c.Content.Blocks.Add(TableOps.EmptyParagraph(w, null, paraId));
        return c;
    }

    static readonly string[] TextProps = { "color", "font-size", "font-weight", "font-style", "font-family", "text-decoration", "vertical-align" };

    static TableCell BuildCell(CellSpec c, int span, int[] grid, XNamespace w, Func<string> paraId, Dictionary<string, List<(string Prop, string Value)>> css)
    {
        var tc = new TableCell { Props = new PropertyBag("tcPr", w) };
        var st = c.Style;
        tc.Props.Set(new XElement(w + "tcW", new XAttribute(w + "w", grid.Skip(c.Col).Take(span).Sum()), new XAttribute(w + "type", "dxa")));
        if (span > 1) tc.Props.SetVal("gridSpan", span.ToString(CultureInfo.InvariantCulture));

        // bordures : seuls les côtés dont le CSS parle
        var sides = new[] { ("top", "top"), ("left", "left"), ("bottom", "bottom"), ("right", "right") }
            .Select(s => (Side: s.Item1, Line: st.TryGetValue("border-" + s.Item2, out var bv) ? Border(bv) : null))
            .Where(s => s.Line is not null).ToList();
        if (sides.Count > 0) tc.Props.Set(new XElement(w + "tcBorders", sides.Select(s => s.Line!.Value.Element(w, s.Side))));

        var fill = (st.TryGetValue("background-color", out var bg) ? bg : st.TryGetValue("background", out var bg2) ? bg2 : null) ?? Attr(c.Attrs, "bgcolor");
        if (fill is not null && HtmlReader.Color(fill.Split(' ')[0]) is { } hex)
            tc.Props.Set(new XElement(w + "shd", new XAttribute(w + "val", "clear"), new XAttribute(w + "color", "auto"), new XAttribute(w + "fill", hex)));
        if (st.TryGetValue("white-space", out var ws) && ws.Trim().Equals("nowrap", StringComparison.OrdinalIgnoreCase) || Attr(c.Attrs, "nowrap") is not null
            || Regex.IsMatch(c.Attrs, @"\bnowrap\b", RegexOptions.IgnoreCase))
            tc.Props.SetFlag("noWrap", true);
        var pads = new[] { "top", "left", "bottom", "right" }.Select(s => (s, v: st.TryGetValue("padding-" + s, out var pv) ? Length(pv) : null)).ToList();
        if (pads.Any(p => p.v is not null))
            tc.Props.Set(new XElement(w + "tcMar", pads.Select(p => new XElement(w + p.s,
                new XAttribute(w + "w", (int)Math.Round(Math.Max(0, p.v ?? 0) * 20)), new XAttribute(w + "type", "dxa")))));
        var valign = (st.TryGetValue("vertical-align", out var va) ? va : Attr(c.Attrs, "valign"))?.Trim().ToLowerInvariant();
        if (valign is "middle" or "center") tc.Props.SetVal("vAlign", "center");
        else if (valign is "bottom") tc.Props.SetVal("vAlign", "bottom");

        // contenu : mise en forme de la cellule passée au lecteur HTML par un span englobant
        var style = new StringBuilder();
        foreach (var p in TextProps)
            if (st.TryGetValue(p, out var v) && !(p == "vertical-align")) style.Append(p).Append(':').Append(v.Replace('"', '\'')).Append(';');
        if (c.Tag == "th" && !st.ContainsKey("font-weight")) style.Append("font-weight:bold;");
        var inner = ExpandClasses(c.Html, css);
        var paras = HtmlReader.Read($"<span style=\"{style}\">{inner}</span>", out _);
        var built = Fragments.Build(paras, true, w).Blocks.OfType<Paragraph>().ToList();
        // marque de paragraphe : mise en forme de la cellule (Word), pour que la frappe la garde
        var markSpec = new Fragments.ParaSpec();
        markSpec.Runs.Add(("x", HtmlReader.Read($"<span style=\"{style}\">x</span>", out _).FirstOrDefault()?.Runs.FirstOrDefault().Format ?? new Fragments.CharFormat()));
        var markRun = Fragments.Build(new[] { markSpec }, true, w).Blocks.OfType<Paragraph>().First().Inlines.OfType<Run>().First();
        if (built.Count == 0) built.Add(TableOps.EmptyParagraph(w, null, paraId));
        var align = (st.TryGetValue("text-align", out var ta) ? ta : Attr(c.Attrs, "align"))?.Trim().ToLowerInvariant() switch
        {
            "center" => "center", "right" or "end" => "right", "justify" => "both", _ => null,
        };
        foreach (var p in built)
        {
            if (align is not null && p.Props.Val("jc") is null) p.Props.SetVal("jc", align);
            if (markRun.Props.Element is { } rpr && rpr.HasElements) p.Props.Set(new XElement(rpr));
            tc.Content.Blocks.Add(p);
        }
        return tc;
    }

    /// <summary>
    /// Balises internes avec une classe (Excel : &lt;font class="font5"&gt; pour le texte enrichi d'une cellule) : le CSS
    /// de la classe est recopié en style en ligne, que le lecteur HTML comprend.
    /// </summary>
    static string ExpandClasses(string html, Dictionary<string, List<(string Prop, string Value)>> css) =>
        Regex.Replace(html, @"<([a-zA-Z][a-zA-Z0-9]*)(\s[^>]*)>", m =>
        {
            var tag = m.Groups[1].Value.ToLowerInvariant();
            var attrs = m.Groups[2].Value;
            if (Attr(attrs, "class") is not { } classes) return m.Value;
            var decls = classes.Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .SelectMany(cls => (css.TryGetValue("." + cls.ToLowerInvariant(), out var d) ? d : new()).Concat(css.TryGetValue(tag + "." + cls.ToLowerInvariant(), out var td) ? td : new()))
                .Where(d => TextProps.Contains(d.Prop) && d.Prop != "vertical-align")
                .Select(d => $"{d.Prop}:{d.Value.Replace('"', '\'')}").ToList();
            if (decls.Count == 0) return m.Value;
            if (Attr(attrs, "style") is { } inline) decls.Add(inline.Replace('"', '\''));
            var rest = Regex.Replace(attrs, @"\b(style|class)\s*=\s*(""[^""]*""|'[^']*'|[^\s>]+)", "", RegexOptions.IgnoreCase);
            return $"<{m.Groups[1].Value} style=\"{string.Join(";", decls)}\"{rest}>";
        });
}
