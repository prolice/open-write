using System.Globalization;
using System.Windows;
using System.Windows.Media;
using System.Xml.Linq;
using OpenWrite.Layout;

namespace OpenWrite.App.Layout;

/// <summary>
/// Dessin d'une page mise en page (PLAN.md §3.3), en points : le texte est posé glyphe par glyphe avec exactement les
/// avances utilisées par la mise en page (GlyphRun), ce qui garantit que l'écran montre la coupure calculée.
/// Fond et bordures des cellules, puces et numéros, soulignés et barrés, surlignage, exposant et indice, images en
/// ligne, en-tête et pied (numéros de page calculés), notes de bas de page.
/// </summary>
public sealed class PageRenderer
{
    readonly Func<XElement, ImageSource?> _image;
    static readonly GlyphTypeface? Fallback = LoadFallback();

    /// <param name="image">Image d'un objet en ligne (w:drawing), ou null (un cadre gris est dessiné).</param>
    public PageRenderer(Func<XElement, ImageSource?> image) => _image = image;

    /// <summary>Marques de mise en forme (Ctrl+Maj+8) : ¶, espaces, tabulations, retours à la ligne, sauts de page.</summary>
    public bool ShowMarks { get; set; }

    static readonly Brush MarkInk = Freeze(new SolidColorBrush(Color.FromRgb(0x5B, 0x74, 0xA8)));

    static GlyphTypeface? LoadFallback()
    {
        foreach (var name in new[] { "Segoe UI Symbol", "Segoe UI", "Arial" })
            if (new Typeface(new FontFamily(name), FontStyles.Normal, FontWeights.Normal, FontStretches.Normal).TryGetGlyphTypeface(out var gt)) return gt;
        return null;
    }

    /// <summary>Dessine le contenu de la page (sans le fond blanc), en points depuis le coin haut gauche.</summary>
    public void Draw(DrawingContext dc, Page page, DocumentLayout layout, int number, int count)
    {
        if (layout.HeaderFooter(page, count) is { } hf) DrawContent(dc, hf, header: true);
        DrawContent(dc, page, header: false);
        if (layout.Notes(page) is { } notes) DrawContent(dc, notes, header: false);
    }

    static readonly Pen RulePen = Freeze(new Pen(Brushes.Black, 0.5));

    static readonly Brush HeaderInk = Freeze(new SolidColorBrush(Color.FromArgb(0xFF, 0x60, 0x60, 0x60)));

    void DrawContent(DrawingContext dc, Page page, bool header)
    {
        foreach (var cell in page.Cells) DrawCell(dc, cell);
        // traits entre colonnes (w:cols w:sep)
        foreach (var (x, top, bottom) in page.ColumnRules) dc.DrawLine(RulePen, new Point(x, top), new Point(x, bottom));
        DrawParagraphBoxes(dc, page);
        foreach (var line in page.Lines)
        {
            DrawLine(dc, line, header ? HeaderInk : null);
            if (ShowMarks) DrawMarks(dc, line);
        }
    }

    /// <summary>Marques de mise en forme d'une ligne, aux positions exactes de la mise en page.</summary>
    static void DrawMarks(DrawingContext dc, PlacedLine placed)
    {
        var pl = placed.Paragraph;
        var line = pl.Lines[placed.Line];
        var c = pl.Content;
        double baseline = placed.Y + line.Baseline;
        double Size(int k) => k < c.Text.Length ? c.Styles[c.StyleOf[k]].Size : c.Mark.Size;
        for (int k = line.Start; k < line.End && k < c.Text.Length; k++)
        {
            char ch = c.Text[k];
            if (ch is not (' ' or '\u00A0' or Atom.Tab)) continue;
            double x0 = placed.X + CaretGeometry.X(pl, placed.Line, k), x1 = placed.X + CaretGeometry.X(pl, placed.Line, k + 1);
            double size = Size(k);
            if (ch == Atom.Tab) DrawMark(dc, "→", Math.Max(x0, (x0 + x1) / 2 - size * 0.35), baseline, size);
            else DrawMark(dc, ch == ' ' ? "·" : "°", (x0 + x1) / 2 - size * 0.14, baseline, size);
        }
        double xe = placed.X + CaretGeometry.X(pl, placed.Line, line.End);
        switch (line.EndKind)
        {
            case LineEnd.Paragraph: DrawMark(dc, "¶", xe + 0.5, baseline, c.Mark.Size); break;
            case LineEnd.LineBreak: DrawMark(dc, "↵", xe + 0.5, baseline, Size(line.End)); break;
            case LineEnd.PageBreak:
            case LineEnd.ColumnBreak:
                var label = line.EndKind == LineEnd.PageBreak ? "Saut de page" : "Saut de colonne";
                var pen = Freeze(new Pen(MarkInk, 0.6) { DashStyle = DashStyles.Dot });
                double right = placed.X + pl.Width, y = baseline - c.Mark.Size * 0.3;
                double mid = (xe + right) / 2;
                dc.DrawLine(pen, new Point(xe + 2, y), new Point(mid - 34, y));
                DrawMark(dc, label, mid - 30, y + 2.5, 7);
                dc.DrawLine(pen, new Point(mid + 34, y), new Point(right, y));
                if (line.End >= c.Text.Length - 1) DrawMark(dc, "¶", right + 1, baseline, c.Mark.Size);
                break;
        }
    }

    static void DrawMark(DrawingContext dc, string text, double x, double baseline, double size)
    {
        if (Fallback is not { } fb) return;
        var glyphs = new List<ushort>();
        var adv = new List<double>();
        foreach (var ch in text)
            if (fb.CharacterToGlyphMap.TryGetValue(ch, out var g)) { glyphs.Add(g); adv.Add(fb.AdvanceWidths[g] * size); }
        if (glyphs.Count == 0) return;
        dc.DrawGlyphRun(MarkInk, new GlyphRun(fb, 0, false, size, 1.0f, glyphs.ToArray(), new Point(x, baseline), adv.ToArray(), null, null, null, null, null, null));
    }

    /// <summary>
    /// Bordures et trame des paragraphes (w:pBdr, w:shd) : les paragraphes consécutifs aux bordures identiques forment
    /// une seule boîte (bordure du haut au premier, du bas au dernier), comme dans Word et comme la mise en page.
    /// </summary>
    static void DrawParagraphBoxes(DrawingContext dc, Page page)
    {
        // étendue verticale de chaque paragraphe sur la page, dans l'ordre
        var paras = new List<(ParagraphLayout L, double X, double Top, double Bottom)>();
        foreach (var l in page.Lines)
        {
            double bottom = l.Y + l.Paragraph.Lines[l.Line].Height;
            if (paras.Count > 0 && ReferenceEquals(paras[^1].L, l.Paragraph) && Math.Abs(paras[^1].X - l.X) < 0.01)
                paras[^1] = (l.Paragraph, l.X, Math.Min(paras[^1].Top, l.Y), Math.Max(paras[^1].Bottom, bottom));
            else paras.Add((l.Paragraph, l.X, l.Y, bottom));
        }
        for (int i = 0; i < paras.Count; i++)
        {
            var (pl, x, top, bottom) = paras[i];
            var eff = pl.Format.Effective;
            var W = eff.Name.Namespace;
            var bdr = eff.Element(W + "pBdr");
            var fill = eff.Element(W + "shd")?.Attribute(W + "fill")?.Value;
            if (bdr is null && (fill is null or "auto")) continue;
            // groupe : paragraphes suivants aux mêmes bordures et à la même trame
            int j = i;
            while (j + 1 < paras.Count && Same(paras[j + 1].L, pl) && Math.Abs(paras[j + 1].X - x) < 0.01) j++;
            double gTop = top, gBottom = paras[j].Bottom;
            (double Width, double Space) Side(string side)
            {
                var b = bdr?.Element(W + side);
                if (b is null || b.Attribute(W + "val")?.Value is null or "none" or "nil") return (0, 0);
                double sz = double.TryParse(b.Attribute(W + "sz")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v / 8 : 0.5;
                double sp = double.TryParse(b.Attribute(W + "space")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var s2) ? s2 : 0;
                return (sz, sp);
            }
            var (tw, ts) = Side("top");
            var (bw, bs) = Side("bottom");
            var (lw, ls) = Side("left");
            var (rw, rs) = Side("right");
            double left = x + pl.Format.IndentLeft / 20.0 - ls - lw / 2, right = x + pl.Width - pl.Format.IndentRight / 20.0 + rs + rw / 2;
            double y1 = gTop - ts - tw / 2, y2 = gBottom + bs + bw / 2;
            if (fill is { Length: 6 } && ParseColor(fill) is { } fc)
                dc.DrawRectangle(Freeze(new SolidColorBrush(fc)), null, new Rect(left, y1, Math.Max(0, right - left), Math.Max(0, y2 - y1)));
            else if (eff.Element(W + "shd")?.Attribute(W + "val")?.Value is { } pat && pat.StartsWith("pct") && int.TryParse(pat[3..], out var pct))
                dc.DrawRectangle(Freeze(new SolidColorBrush(System.Windows.Media.Color.FromArgb((byte)(pct * 255 / 100), 0, 0, 0))), null,
                    new Rect(left, y1, Math.Max(0, right - left), Math.Max(0, y2 - y1)));
            void Edge(double t, double x1, double ya, double x2, double yb)
            {
                if (t > 0) dc.DrawLine(new Pen(Brushes.Black, t), new Point(x1, ya), new Point(x2, yb));
            }
            Edge(tw, left, y1, right, y1);
            Edge(bw, left, y2, right, y2);
            Edge(lw, left, y1, left, y2);
            Edge(rw, right, y1, right, y2);
            i = j;
        }
    }

    static bool Same(ParagraphLayout a, ParagraphLayout b)
    {
        var W = a.Format.Effective.Name.Namespace;
        var ba = a.Format.Effective.Element(W + "pBdr");
        var bb = b.Format.Effective.Element(W + "pBdr");
        var sa = a.Format.Effective.Element(W + "shd");
        var sb = b.Format.Effective.Element(W + "shd");
        return (ba is null ? bb is null : bb is not null && System.Xml.Linq.XNode.DeepEquals(ba, bb))
            && (sa is null ? sb is null : sb is not null && System.Xml.Linq.XNode.DeepEquals(sa, sb));
    }

    static void DrawCell(DrawingContext dc, CellBox c)
    {
        if (c.Fill is { } fill && ParseColor(fill) is { } fc) dc.DrawRectangle(Freeze(new SolidColorBrush(fc)), null, new Rect(c.X, c.Y, Math.Max(0, c.W), Math.Max(0, c.H)));
        void Edge(double t, double x1, double y1, double x2, double y2)
        {
            if (t <= 0) return;
            dc.DrawLine(new Pen(Brushes.Black, t), new Point(x1, y1), new Point(x2, y2));
        }
        Edge(c.Top, c.X, c.Y, c.X + c.W, c.Y);
        Edge(c.Bottom, c.X, c.Y + c.H, c.X + c.W, c.Y + c.H);
        Edge(c.Left, c.X, c.Y, c.X, c.Y + c.H);
        Edge(c.Right, c.X + c.W, c.Y, c.X + c.W, c.Y + c.H);
    }

    void DrawLine(DrawingContext dc, PlacedLine placed, Brush? ink)
    {
        var pl = placed.Paragraph;
        var line = pl.Lines[placed.Line];
        var c = pl.Content;
        double baseline = placed.Y + line.Baseline;
        if (placed.Line == 0 && pl.Marker is { } m && m.Text.Length > 0)
            DrawText(dc, m.Text, m.Style, placed.X + m.X, baseline, 0, ink);
        // justification : chaque espace s'élargit ; les morceaux suivants se décalent d'autant
        double shift = 0;
        foreach (var piece in line.Pieces)
        {
            var st = c.Styles[piece.Style];
            double x = placed.X + piece.X + shift;
            var text = c.Text.AsSpan(piece.Start, piece.Length);
            int spaces = 0;
            foreach (var ch in text) if (ch == ' ') spaces++;
            double extra = line.ExtraPerSpace;
            double width = piece.Width + spaces * extra;
            var rf = st.Format;
            if (rf.Highlight is { } hl && Highlight(hl) is { } hc)
                dc.DrawRectangle(Freeze(new SolidColorBrush(hc)), null, new Rect(x, placed.Y, width, line.Height));
            double dy = rf.VerticalAlign switch { "superscript" => -rf.Size * 0.33, "subscript" => rf.Size * 0.14, _ => 0 };
            double cx = x;
            int k = piece.Start;
            var run = new System.Text.StringBuilder();
            void Flush()
            {
                if (run.Length == 0) return;
                cx = DrawText(dc, run.ToString(), st, cx, baseline + dy, extra, ink);
                run.Clear();
            }
            for (; k < piece.Start + piece.Length; k++)
            {
                char ch = c.Text[k];
                if (ch == Atom.Object)
                {
                    Flush();
                    if (c.Objects.TryGetValue(k, out var size))
                    {
                        var rect = new Rect(cx, baseline - size.H, Math.Max(0.1, size.W), Math.Max(0.1, size.H));
                        if (c.ObjectXml.TryGetValue(k, out var xml) && _image(xml) is { } img) dc.DrawImage(img, rect);
                        else dc.DrawRectangle(Freeze(new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xEE, 0xF1, 0xF6))), new Pen(Brushes.LightGray, 0.5), rect);
                        cx += size.W;
                    }
                    continue;
                }
                if (ch is Atom.Tab or Atom.LineBreak or Atom.PageBreak or Atom.ColumnBreak or '­') continue;
                run.Append(ch);
            }
            Flush();
            var brush = ink ?? Ink(rf.Color);
            double size0 = Math.Max(1, rf.Size);
            if (rf.Underline is not null and not "none")
            {
                double uy = baseline + size0 * 0.12;
                dc.DrawLine(new Pen(brush, Math.Max(0.5, size0 / 18)), new Point(x, uy), new Point(x + width - (text.EndsWith(" ") ? st.Face.Advance(' ') * st.Size + extra : 0), uy));
                if (rf.Underline == "double") dc.DrawLine(new Pen(brush, Math.Max(0.5, size0 / 18)), new Point(x, uy + size0 * 0.1), new Point(x + width, uy + size0 * 0.1));
            }
            if (rf.Strike || rf.DoubleStrike)
            {
                double sy = baseline - size0 * 0.28;
                dc.DrawLine(new Pen(brush, Math.Max(0.5, size0 / 18)), new Point(x, sy), new Point(x + width, sy));
                if (rf.DoubleStrike) dc.DrawLine(new Pen(brush, Math.Max(0.5, size0 / 18)), new Point(x, sy - size0 * 0.1), new Point(x + width, sy - size0 * 0.1));
            }
            shift += spaces * extra;
        }
    }

    /// <summary>Pose un texte d'un même style ; renvoie l'abscisse de fin.</summary>
    static double DrawText(DrawingContext dc, string text, TextStyle st, double x, double baseline, double extraPerSpace, Brush? ink)
    {
        var brush = ink ?? Ink(st.Format.Color);
        var face = st.Face as WpfFontProvider.GlyphFace;
        var glyphs = new List<ushort>();
        var advances = new List<double>();
        double runStart = x;
        void Flush()
        {
            if (glyphs.Count > 0 && face is not null)
                dc.DrawGlyphRun(brush, new GlyphRun(face.Typeface, 0, false, st.Size, 1.0f, glyphs.ToArray(), new Point(runStart, baseline),
                    advances.ToArray(), null, null, null, null, null, null));
            glyphs.Clear();
            advances.Clear();
        }
        foreach (var ch in text)
        {
            double adv = st.Face.Advance(ch) * st.Size + st.Tracking + (ch == ' ' ? extraPerSpace : 0);
            if (face?.Glyph(ch) is { } g)
            {
                if (glyphs.Count == 0) runStart = x;
                glyphs.Add(g);
                advances.Add(adv);
            }
            else
            {
                // caractère absent de la police : police de secours (☐, ☒, symboles…)
                Flush();
                if (Fallback is { } fb && fb.CharacterToGlyphMap.TryGetValue(ch, out var fg))
                    dc.DrawGlyphRun(brush, new GlyphRun(fb, 0, false, st.Size, 1.0f, new[] { fg }, new Point(x, baseline), new[] { adv }, null, null, null, null, null, null));
            }
            x += adv;
        }
        Flush();
        return x;
    }

    static readonly Dictionary<string, Brush> Inks = new();

    static Brush Ink(string? rgb)
    {
        if (rgb is null) return Brushes.Black;
        if (!Inks.TryGetValue(rgb, out var b)) Inks[rgb] = b = ParseColor(rgb) is { } c ? Freeze(new SolidColorBrush(c)) : Brushes.Black;
        return b;
    }

    static Color? ParseColor(string rgb) =>
        rgb.Length == 6 && int.TryParse(rgb, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var v)
            ? System.Windows.Media.Color.FromRgb((byte)(v >> 16), (byte)(v >> 8), (byte)v) : null;

    static Color? Highlight(string name) => name switch
    {
        "yellow" => Colors.Yellow, "green" => Colors.Lime, "cyan" => Colors.Cyan, "magenta" => Colors.Magenta, "blue" => Colors.Blue,
        "red" => Colors.Red, "darkBlue" => Colors.DarkBlue, "darkCyan" => Colors.DarkCyan, "darkGreen" => Colors.DarkGreen,
        "darkMagenta" => Colors.DarkMagenta, "darkRed" => Colors.DarkRed, "darkYellow" => Colors.Olive, "darkGray" => Colors.DarkGray,
        "lightGray" => Colors.LightGray, "black" => Colors.Black, "white" => Colors.White, _ => null,
    };

    static T Freeze<T>(T f) where T : Freezable { f.Freeze(); return f; }
}
