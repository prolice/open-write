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
    readonly Func<string?, string, ImageSource?> _image;
    static readonly GlyphTypeface? Fallback = LoadFallback();

    /// <param name="image">
    /// Image d'une relation (r:embed) d'une partie : null = corps du document, sinon r:id de l'en-tête ou du pied ; null
    /// si elle manque ou n'est pas lisible (un cadre gris est dessiné).
    /// </param>
    public PageRenderer(Func<string?, string, ImageSource?> image) => _image = image;

    /// <summary>Partie (en-tête, pied) des lignes en cours de dessin : ses images ; null = corps.</summary>
    string? _part;

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
    /// <summary>Barres de modification dans la marge de gauche (null : aucune) : rouge en marques simples, grise avec toutes les marques.</summary>
    public Brush? RevisionBar { get; set; }

    /// <summary>Abscisse de la barre de modification de la page en cours de dessin.</summary>
    double _barX;

    public void Draw(DrawingContext dc, Page page, DocumentLayout layout, int number, int count)
    {
        _barX = Math.Max(4, page.Geometry.Left - 14);
        // filigrane (derrière le texte), bordure de page, numéros de ligne
        if (layout.Watermark(page) is { } wm) DrawWatermark(dc, page, wm);
        DrawPageBorders(dc, page);
        DrawLineNumbers(dc, page, layout);
        // en-têtes et pieds sous le corps (objets derrière le texte, texte, autres objets), puis le corps de même
        DrawFloats(dc, page, f => f.PartId is not null && f.Object.BehindText);
        if (layout.HeaderFooter(page, count) is { } hf)
        {
            foreach (var (first, n, part) in hf.Parts)
            {
                _part = part;
                DrawContent(dc, hf, header: !HeaderMode && !ForOutput, first, n);
            }
            _part = null;
        }
        DrawFloats(dc, page, f => f.PartId is not null && !f.Object.BehindText);
        // édition d'un en-tête : le corps est grisé, les limites de l'en-tête et du pied sont tracées (comme Word)
        if (HeaderMode) dc.PushOpacity(0.4);
        DrawFloats(dc, page, f => f.PartId is null && f.Object.BehindText);
        DrawContent(dc, page, header: false);
        if (layout.Notes(page) is { } notes) DrawContent(dc, notes, header: false);
        DrawFloats(dc, page, f => f.PartId is null && !f.Object.BehindText);
        if (HeaderMode)
        {
            dc.Pop();
            var g = page.Geometry;
            double top = page.ContentTop, bottom = g.PageHeight - (page.FirstOfSection ? g.FirstBottom : g.Bottom);
            dc.DrawLine(BoundaryPen, new Point(0, top), new Point(g.PageWidth, top));
            dc.DrawLine(BoundaryPen, new Point(0, bottom), new Point(g.PageWidth, bottom));
            DrawLabel(dc, layout.HeaderType(page) switch { "first" => "En-tête de la première page", "even" => "En-tête de page paire", _ => "En-tête" }, g.Left, top + 1);
            DrawLabel(dc, layout.HeaderType(page) switch { "first" => "Pied de la première page", "even" => "Pied de page paire", _ => "Pied de page" }, g.Left, bottom - 11);
        }
    }

    static readonly Dictionary<string, Color> VmlColors = new(StringComparer.OrdinalIgnoreCase)
    {
        ["silver"] = Color.FromRgb(0xC0, 0xC0, 0xC0), ["gray"] = Colors.Gray, ["black"] = Colors.Black, ["red"] = Colors.Red, ["blue"] = Colors.Blue,
    };

    static Color VmlColor(string c) => VmlColors.TryGetValue(c, out var k) ? k
        : c.TrimStart('#') is { Length: 6 } hex && ParseColor(hex) is { } p ? p : Color.FromRgb(0xC0, 0xC0, 0xC0);

    /// <summary>Filigrane : texte WordArt centré dans la marge, tourné, étiré à la taille de la forme.</summary>
    static void DrawWatermark(DrawingContext dc, Page page, (string Text, string Font, string Color, double Rotation, double Width, double Height, double Opacity) wm)
    {
        if (wm.Text.Length == 0) return;
        var g = page.Geometry;
        double cx = g.Left + g.ColumnWidth / 2, cy = (page.ContentTop + g.PageHeight - g.Bottom) / 2;
        var brush = new SolidColorBrush(VmlColor(wm.Color)) { Opacity = Math.Clamp(wm.Opacity, 0.05, 1) };
        brush.Freeze();
        var ft = new FormattedText(wm.Text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, new Typeface(wm.Font), 100, brush, 1.0);
        double sx = wm.Width > 0 ? wm.Width / Math.Max(1, ft.Width) : 1, sy = wm.Height > 0 ? wm.Height / Math.Max(1, ft.Height) : sx;
        dc.PushTransform(new TranslateTransform(cx, cy));
        dc.PushTransform(new RotateTransform(wm.Rotation));
        dc.PushTransform(new ScaleTransform(sx, sy));
        dc.DrawText(ft, new Point(-ft.Width / 2, -ft.Height / 2));
        dc.Pop(); dc.Pop(); dc.Pop();
    }

    /// <summary>Bordure de page (w:pgBorders) : distance au bord de la page ou au texte, première page / autres.</summary>
    static void DrawPageBorders(DrawingContext dc, Page page)
    {
        var g = page.Geometry;
        if (g.SectPr?.Element(g.SectPr.Name.Namespace + "pgBorders") is not { } b) return;
        var w = b.Name.Namespace;
        var display = b.Attribute(w + "display")?.Value ?? "allPages";
        if (display == "firstPage" && !page.FirstOfSection || display == "notFirstPage" && page.FirstOfSection) return;
        bool fromPage = b.Attribute(w + "offsetFrom")?.Value == "page";
        double bottomText = g.PageHeight - (page.FirstOfSection ? g.FirstBottom : g.Bottom);
        foreach (var side in new[] { "top", "left", "bottom", "right" })
        {
            if (b.Element(w + side) is not { } e || e.Attribute(w + "val")?.Value is null or "none" or "nil") continue;
            double Num(string a, double d) => double.TryParse(e.Attribute(w + a)?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : d;
            double width = Math.Max(0.25, Num("sz", 4) / 8), space = Num("space", fromPage ? 24 : 4);
            var val = e.Attribute(w + "val")!.Value;
            var color = e.Attribute(w + "color")?.Value is { Length: 6 } c && ParseColor(c) is { } col ? col : Colors.Black;
            var pen = new Pen(new SolidColorBrush(color), width) { DashStyle = val switch { "dotted" => DashStyles.Dot, "dashed" or "dashSmallGap" => DashStyles.Dash, "dotDash" => DashStyles.DashDot, _ => DashStyles.Solid } };
            pen.Freeze();
            // rectangle de la bordure
            double l = fromPage ? space : g.Left - space, r = fromPage ? g.PageWidth - space : g.PageWidth - g.Right + space;
            double t = fromPage ? space : page.ContentTop - space, bt = fromPage ? g.PageHeight - space : bottomText + space;
            (Point a, Point z) = side switch
            {
                "top" => (new Point(l, t), new Point(r, t)),
                "bottom" => (new Point(l, bt), new Point(r, bt)),
                "left" => (new Point(l, t), new Point(l, bt)),
                _ => (new Point(r, t), new Point(r, bt)),
            };
            dc.DrawLine(pen, a, z);
            if (val is "double" or "thinThickSmallGap" or "thickThinSmallGap")
            {
                double o = width * 2 + 1;
                var (dx, dy) = side switch { "top" => (0.0, o), "bottom" => (0.0, -o), "left" => (o, 0.0), _ => (-o, 0.0) };
                dc.DrawLine(pen, new Point(a.X + dx, a.Y + dy), new Point(z.X + dx, z.Y + dy));
            }
        }
    }

    /// <summary>Numéros de ligne dans la marge de gauche (w:lnNumType : pas, distance).</summary>
    static void DrawLineNumbers(DrawingContext dc, Page page, DocumentLayout layout)
    {
        if (page.FirstLineNumber <= 0) return;
        var g = page.Geometry;
        var ln = g.SectPr!.Element(g.SectPr.Name.Namespace + "lnNumType")!;
        var w = ln.Name.Namespace;
        int countBy = int.TryParse(ln.Attribute(w + "countBy")?.Value, out var cb) && cb > 0 ? cb : 1;
        double distance = double.TryParse(ln.Attribute(w + "distance")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? d / 20 : 18;
        int n = page.FirstLineNumber;
        foreach (var l in layout.NumberedLines(page))
        {
            if (n % countBy == 0)
            {
                var box = l.Paragraph.Lines[l.Line];
                var rf = layout.DefaultRunFormat;
                var ft = new FormattedText(n.ToString(CultureInfo.InvariantCulture), CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
                    new Typeface(rf.FontAscii ?? "Calibri"), Math.Max(6, rf.Size), Brushes.Black, 1.0);
                dc.DrawText(ft, new Point(g.Left - distance - ft.Width, l.Y + box.Baseline - ft.Baseline));
            }
            n++;
        }
    }

    /// <summary>Édition des en-têtes et pieds : en-têtes à l'encre normale, corps grisé.</summary>
    public bool HeaderMode { get; set; }

    /// <summary>
    /// Fautes d'orthographe d'un paragraphe (index affichés, longueur) à souligner d'un trait ondulé rouge ; null = aucune
    /// vérification (impression, PDF, vérification désactivée).
    /// </summary>
    public Func<ParagraphContent, IReadOnlyList<(int Start, int Length)>>? Misspellings { get; set; }

    static readonly Pen SpellPen = Freeze(new Pen(Freeze(new SolidColorBrush(Color.FromRgb(0xE0, 0x1B, 0x24))), 0.75) { LineJoin = PenLineJoin.Round });

    static bool Revised(byte[]? revs, int a, int b)
    {
        if (revs is null) return false;
        for (int i = Math.Max(0, a); i < b && i < revs.Length; i++) if (revs[i] != 0) return true;
        return false;
    }

    /// <summary>Trait ondulé de x1 à x2 (points), comme Word sous une faute.</summary>
    static void DrawSquiggle(DrawingContext dc, double x1, double x2, double y)
    {
        if (x2 - x1 < 1) return;
        var g = new StreamGeometry();
        using (var ctx = g.Open())
        {
            const double Step = 1.5, Amp = 0.9;
            ctx.BeginFigure(new Point(x1, y), false, false);
            bool up = true;
            for (double x = x1 + Step; x < x2 + Step / 2; x += Step, up = !up)
                ctx.LineTo(new Point(Math.Min(x, x2), up ? y - Amp / 2 : y + Amp / 2), true, true);
        }
        g.Freeze();
        dc.DrawGeometry(null, SpellPen, g);
    }

    /// <summary>Impression, PDF : en-têtes et pieds dans leurs vraies couleurs (à l'écran, ils sont estompés).</summary>
    public bool ForOutput { get; set; }

    static readonly Pen BoundaryPen = Freeze(new Pen(Freeze(new SolidColorBrush(Color.FromRgb(0x08, 0x7C, 0xFF))), 0.6) { DashStyle = DashStyles.Dash });
    static readonly Brush LabelBack = Freeze(new SolidColorBrush(Color.FromRgb(0xE1, 0xE9, 0xFF)));

    static void DrawLabel(DrawingContext dc, string text, double x, double y)
    {
        var ft = new FormattedText(text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, new Typeface("Segoe UI"), 7.5, MarkInk, 1.0);
        dc.DrawRectangle(LabelBack, null, new Rect(x, y, ft.Width + 6, ft.Height + 1));
        dc.DrawText(ft, new Point(x + 3, y + 0.5));
    }

    static readonly Pen RulePen = Freeze(new Pen(Brushes.Black, 0.5));

    static readonly Brush HeaderInk = Freeze(new SolidColorBrush(Color.FromArgb(0xFF, 0x60, 0x60, 0x60)));

    void DrawContent(DrawingContext dc, Page page, bool header, int first = 0, int count = -1)
    {
        if (count < 0)
        {
            foreach (var cell in page.Cells) DrawCell(dc, cell);
            // traits entre colonnes (w:cols w:sep)
            foreach (var (x, top, bottom) in page.ColumnRules) dc.DrawLine(RulePen, new Point(x, top), new Point(x, bottom));
            DrawParagraphBoxes(dc, page);
        }
        else if (first == 0) foreach (var cell in page.Cells) DrawCell(dc, cell);
        int end = count < 0 ? page.Lines.Count : Math.Min(page.Lines.Count, first + count);
        for (int k = first; k < end; k++)
        {
            var line = page.Lines[k];
            DrawLine(dc, line, header ? HeaderInk : null);
            if (ShowMarks) DrawMarks(dc, line);
        }
    }

    /// <summary>Objets flottants de la page choisis par <paramref name="which"/>, dans l'ordre d'empilement, avec le texte de leurs zones.</summary>
    void DrawFloats(DrawingContext dc, Page page, Func<FloatBox, bool> which)
    {
        foreach (var f in page.Floats.Where(which).OrderBy(f => f.Object.Z))
        {
            var part = f.PartId;
            bool cellsDrawn = false;
            DrawingRenderer.Draw(dc, f.Object, f.X, f.Y, id => _image(part, id), box =>
            {
                if (!cellsDrawn) { foreach (var cell in f.TextCells) DrawCell(dc, cell); cellsDrawn = true; }
                var saved = _part;
                _part = part;
                foreach (var (b, line) in f.TextLines)
                    if (b == box)
                    {
                        DrawLine(dc, line, null);
                        if (ShowMarks) DrawMarks(dc, line);
                    }
                _part = saved;
            });
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

    /// <summary>
    /// Quadrillage des tableaux (Disposition ▸ Quadrillage, actif par défaut comme dans Word) : côtés de cellule sans
    /// bordure tracés en tirets bleu pâle, à l'écran seulement.
    /// </summary>
    public static bool ShowGridlines { get; set; } = true;

    static readonly Pen GridlinePen = Freeze(new Pen(Freeze(new SolidColorBrush(Color.FromRgb(0x9D, 0xC3, 0xE6))), 0.5) { DashStyle = new DashStyle(new double[] { 2, 2 }, 0) });

    static void DrawCell(DrawingContext dc, CellBox c)
    {
        if (c.Fill is { } fill && ParseColor(fill) is { } fc) dc.DrawRectangle(Freeze(new SolidColorBrush(fc)), null, new Rect(c.X, c.Y, Math.Max(0, c.W), Math.Max(0, c.H)));
        if (ShowGridlines && c.Cell is not null)
        {
            if (!c.Top.Visible && !c.JoinTop) dc.DrawLine(GridlinePen, new Point(c.X, c.Y), new Point(c.X + c.W, c.Y));
            if (!c.Bottom.Visible && !c.JoinBottom) dc.DrawLine(GridlinePen, new Point(c.X, c.Y + c.H), new Point(c.X + c.W, c.Y + c.H));
            if (!c.Left.Visible) dc.DrawLine(GridlinePen, new Point(c.X, c.Y), new Point(c.X, c.Y + c.H));
            if (!c.Right.Visible) dc.DrawLine(GridlinePen, new Point(c.X + c.W, c.Y), new Point(c.X + c.W, c.Y + c.H));
        }
        DrawBorder(dc, c.Top, c.X, c.Y, c.X + c.W, c.Y);
        DrawBorder(dc, c.Bottom, c.X, c.Y + c.H, c.X + c.W, c.Y + c.H);
        DrawBorder(dc, c.Left, c.X, c.Y, c.X, c.Y + c.H);
        DrawBorder(dc, c.Right, c.X + c.W, c.Y, c.X + c.W, c.Y + c.H);
    }

    static readonly System.Collections.Concurrent.ConcurrentDictionary<(double, string?, string), Pen> BorderPens = new();

    /// <summary>
    /// Trait de bordure de cellule : couleur, épaisseur et style de Word (simple, épais, double, pointillés, tirets ;
    /// les styles composés et artistiques sont tracés simples).
    /// </summary>
    static void DrawBorder(DrawingContext dc, BorderLine b, double x1, double y1, double x2, double y2)
    {
        if (!b.Visible) return;
        if (b.Style is "double" or "triple" || b.Style.StartsWith("thinThick", StringComparison.Ordinal) || b.Style.StartsWith("thickThin", StringComparison.Ordinal))
        {
            // deux traits parallèles écartés de l'épaisseur du trait
            var pen = BorderPen(b with { Style = "single" });
            double d = Math.Max(0.5, b.Width);
            bool horizontal = Math.Abs(y2 - y1) < Math.Abs(x2 - x1);
            var (dx, dy) = horizontal ? (0.0, d) : (d, 0.0);
            dc.DrawLine(pen, new Point(x1 - dx, y1 - dy), new Point(x2 - dx, y2 - dy));
            dc.DrawLine(pen, new Point(x1 + dx, y1 + dy), new Point(x2 + dx, y2 + dy));
            return;
        }
        dc.DrawLine(BorderPen(b), new Point(x1, y1), new Point(x2, y2));
    }

    static Pen BorderPen(BorderLine b)
    {
        var key = (b.Width, b.Color, b.Style);
        if (BorderPens.TryGetValue(key, out var pen)) return pen;
        Brush brush = b.Color is { } col && ParseColor(col) is { } c ? Freeze(new SolidColorBrush(c)) : Brushes.Black;
        pen = new Pen(brush, b.Width);
        pen.DashStyle = b.Style switch
        {
            "dotted" => new DashStyle(new double[] { 1, 1 }, 0),
            "dashed" or "dashSmallGap" => new DashStyle(new double[] { 3, 2 }, 0),
            "dotDash" => new DashStyle(new double[] { 3, 1.5, 1, 1.5 }, 0),
            "dotDotDash" => new DashStyle(new double[] { 3, 1.5, 1, 1.5, 1, 1.5 }, 0),
            _ => DashStyles.Solid,
        };
        pen.Freeze();
        return BorderPens[key] = pen;
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
        // espaces de fin de ligne : jamais soulignés (ceux entre les mots le sont, comme dans Word)
        int lastInk = line.End - 1;
        while (lastInk >= line.Start && lastInk < c.Text.Length && c.Text[lastInk] is ' ' or Atom.LineBreak or Atom.PageBreak or Atom.ColumnBreak) lastInk--;
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
                        if (c.Drawings.TryGetValue(k, out var hr) && hr.HorizontalRule is not null)
                        {
                            // filet horizontal : trait plein de la couleur du rectangle VML
                            var fillRgb = (hr.Root as Core.Drawing.ShapeVisual)?.Fill?.Color ?? "A0A0A0";
                            dc.DrawRectangle(Ink(fillRgb), null, new Rect(cx, baseline - size.H, Math.Max(0.5, size.W), Math.Max(0.5, size.H)));
                        }
                        else if (c.Drawings.TryGetValue(k, out var obj))
                        {
                            // dessin en ligne posé sur la ligne de base ; marges d'effet autour du cadre
                            var part = _part;
                            DrawingRenderer.Draw(dc, obj, cx + obj.Effect.L, baseline - size.H + obj.Effect.T, id => _image(part, id));
                        }
                        else dc.DrawRectangle(Freeze(new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xEE, 0xF1, 0xF6))), new Pen(Brushes.LightGray, 0.5),
                            new Rect(cx, baseline - size.H, Math.Max(0.1, size.W), Math.Max(0.1, size.H)));
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
                int trailing = 0;
                for (int t = piece.Start + piece.Length - 1; t > lastInk && t >= piece.Start; t--) if (c.Text[t] == ' ') trailing++;
                dc.DrawLine(new Pen(brush, Math.Max(0.5, size0 / 18)), new Point(x, uy), new Point(x + width - trailing * (st.Width(' ') + extra), uy));
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
        // révisions : barre dans la marge de gauche, à la hauteur de la ligne (comme Word)
        if (RevisionBar is { } bar && (Revised(c.Revisions, line.Start, line.End) || (c.MarkRevised && line.End >= c.Text.Length)))
            dc.DrawRectangle(bar, null, new Rect(_barX, placed.Y, 1.2, line.Height));
        // fautes d'orthographe : sous les caractères de la ligne, sans les espaces qui la terminent
        if (Misspellings?.Invoke(c) is { Count: > 0 } errors)
        {
            double y = baseline + Math.Clamp((line.Height - line.Baseline) * 0.45, 1.6, 3.5);
            foreach (var (st, len) in errors)
            {
                int a = Math.Max(st, line.Start), b = Math.Min(st + len, Math.Min(line.End, lastInk + 1));
                if (b <= a) continue;
                DrawSquiggle(dc, placed.X + CaretGeometry.X(pl, placed.Line, a), placed.X + CaretGeometry.X(pl, placed.Line, b), y);
            }
        }
    }

    /// <summary>Pose un texte d'un même style ; renvoie l'abscisse de fin.</summary>
    static double DrawText(DrawingContext dc, string text, TextStyle st, double x, double baseline, double extraPerSpace, Brush? ink)
    {
        var brush = ink ?? Ink(st.Format.Color);
        var face = st.Face as WpfFontProvider.GlyphFace;
        var glyphs = new List<ushort>();
        var advances = new List<double>();
        // caractères de chaque glyphe (un pour un) : l'export PDF en a besoin pour écrire le texte (copier, rechercher)
        var chars = new List<char>();
        double runStart = x;
        void Flush()
        {
            if (glyphs.Count > 0 && face is not null)
            {
                // échelle horizontale (w:w) : glyphes étirés autour du début du morceau, avances ramenées d'autant
                bool scaled = Math.Abs(st.Scale - 1) > 1e-6;
                if (scaled) dc.PushTransform(new ScaleTransform(st.Scale, 1, runStart, 0));
                dc.DrawGlyphRun(brush, new GlyphRun(face.Typeface, 0, false, st.Size, 1.0f, glyphs.ToArray(), new Point(runStart, baseline),
                    scaled ? advances.Select(a => a / st.Scale).ToArray() : advances.ToArray(), null, chars.ToArray(), null, null, null, null));
                if (scaled) dc.Pop();
            }
            glyphs.Clear();
            advances.Clear();
            chars.Clear();
        }
        foreach (var ch in text)
        {
            double adv = st.Width(ch) + (ch == ' ' ? extraPerSpace : 0);
            if (face?.Glyph(ch) is { } g)
            {
                if (glyphs.Count == 0) runStart = x;
                glyphs.Add(g);
                advances.Add(adv);
                chars.Add(ch);
            }
            else
            {
                // caractère absent de la police : police de secours (☐, ☒, symboles…)
                Flush();
                if (Fallback is { } fb && fb.CharacterToGlyphMap.TryGetValue(ch, out var fg))
                    dc.DrawGlyphRun(brush, new GlyphRun(fb, 0, false, st.Size, 1.0f, new[] { fg }, new Point(x, baseline), new[] { adv }, null, new[] { ch }, null, null, null, null));
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

    static Color? Highlight(string name) => name.Length == 7 && name[0] == '#' ? ParseColor(name[1..]) : name switch
    {
        "yellow" => Colors.Yellow, "green" => Colors.Lime, "cyan" => Colors.Cyan, "magenta" => Colors.Magenta, "blue" => Colors.Blue,
        "red" => Colors.Red, "darkBlue" => Colors.DarkBlue, "darkCyan" => Colors.DarkCyan, "darkGreen" => Colors.DarkGreen,
        "darkMagenta" => Colors.DarkMagenta, "darkRed" => Colors.DarkRed, "darkYellow" => Colors.Olive, "darkGray" => Colors.DarkGray,
        "lightGray" => Colors.LightGray, "black" => Colors.Black, "white" => Colors.White, _ => null,
    };

    static T Freeze<T>(T f) where T : Freezable { f.Freeze(); return f; }
}
