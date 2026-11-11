// Repris d'OpenPub 2.1.4 (OpenPub.App/Output/DrawingToPdf.cs). OpenWrite : compteur de liens, octets d'origine des images
// décodées depuis le paquet (PdfImageCache.Original).
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using PdfSharp.Drawing;
using PdfSharp.Events;
using PdfSharp.Fonts;
using PdfSharp.Pdf;

namespace OpenWrite.App.Output;

/// <summary>Compteurs d'un export (diagnostic, auto-test).</summary>
public sealed class PdfStats
{
    public int Paths, TextRuns, TextAsOutlines, Images, Rasterized, Links, OriginalImages;
    public override string ToString() =>
        $"{Paths} traces, {TextRuns} morceaux de texte ({TextAsOutlines} en contours), {Images} images, {Rasterized} effets en image, {Links} liens, {OriginalImages} image(s) d'origine";
}

/// <summary>
/// Traduit un arbre <see cref="Drawing"/> WPF (le rendu de <see cref="Layout.PageRenderer"/>) en appels PDFsharp :
/// traces et remplissages vectoriels, texte avec la vraie police (glyphes et positions du GlyphRun), images
/// d'origine. Ce qui n'a pas d'equivalent vectoriel (masque d'opacite, degrades, effets) est rasterise localement.
/// Tout ce qu'on ajoutera plus tard au rendu passe ainsi dans le PDF sans code supplementaire.
/// </summary>
public sealed class DrawingToPdf
{
    readonly XGraphics _g;
    readonly PdfFontCache _fonts;
    readonly PdfImageCache _images;
    public PdfStats Stats { get; }
    /// <summary>Resolution des parties rasterisees (points par pouce).</summary>
    public double RasterDpi { get; init; } = 300;
    /// <summary>Decoupe de la page a remplacer (fond perdu : la page n'est plus coupee a son bord).</summary>
    public (Rect Page, Rect Replacement)? PageClip { get; init; }

    ushort[]? _pendingGlyphs;

    public DrawingToPdf(XGraphics g, PdfDocument doc, PdfFontCache fonts, PdfImageCache images, PdfStats stats)
    {
        _g = g; _fonts = fonts; _images = images; Stats = stats;
        doc.RenderEvents.RenderTextEvent += OnRenderText;
    }

    public void Detach(PdfDocument doc) => doc.RenderEvents.RenderTextEvent -= OnRenderText;

    /// <summary>PDFsharp choisit les glyphes d'apres les caracteres : on impose ceux de WPF (memes formes qu'a l'ecran).</summary>
    void OnRenderText(object? sender, RenderTextEventArgs e)
    {
        if (_pendingGlyphs is not { } glyphs) return;
        var pairs = e.CodePointGlyphIndexPairs;
        if (pairs.Length != glyphs.Length) return;
        for (int i = 0; i < pairs.Length; i++)
            if (glyphs[i] != 0) pairs[i] = new CodePointGlyphIndexPair(pairs[i].CodePoint, glyphs[i]);
        e.CodePointGlyphIndexPairs = pairs;
    }

    public void Draw(Drawing d) => Draw(d, 1, 0);

    void Draw(Drawing d, double opacity, int depth)
    {
        switch (d)
        {
            case DrawingGroup g: DrawGroup(g, opacity, depth); break;
            case GeometryDrawing gd: DrawGeometry(gd, opacity); break;
            case GlyphRunDrawing gr: DrawGlyphs(gr, opacity); break;
            case ImageDrawing id: DrawImage(id, opacity); break;
            default: Rasterize(d, opacity); break;
        }
    }

    void DrawGroup(DrawingGroup grp, double opacity, int depth)
    {
        // masque d'opacite (degrade du reflet), effet bitmap : en image
#pragma warning disable CS0618 // BitmapEffect est obsolete mais peut encore etre utilise
        if (grp.OpacityMask is not null || grp.BitmapEffect is not null) { Rasterize(grp, opacity); return; }
#pragma warning restore CS0618
        double op = opacity * grp.Opacity;
        if (op <= 0.002 || grp.Children.Count == 0) return;
        var transform = grp.Transform is { } t && !t.Value.IsIdentity ? t.Value : (Matrix?)null;
        var clip = grp.ClipGeometry;
        if (clip is RectangleGeometry rg && PageClip is { } pc && depth <= 1 && rg.Transform.Value.IsIdentity && Near(rg.Rect, pc.Page))
            clip = new RectangleGeometry(pc.Replacement);
        XGraphicsState? state = null;
        if (transform is not null || clip is not null) state = _g.Save();
        if (transform is { } m) _g.MultiplyTransform(new XMatrix(m.M11, m.M12, m.M21, m.M22, m.OffsetX, m.OffsetY), XMatrixOrder.Prepend);
        if (clip is not null)
        {
            if (clip.Bounds.IsEmpty) { _g.Restore(state!); return; }
            _g.IntersectClip(ToPath(clip));
        }
        // une opacite de groupe s'applique a chaque enfant (exact pour un seul enfant, ce que produit le rendu)
        foreach (var child in grp.Children) Draw(child, op, depth + 1);
        if (state is not null) _g.Restore(state);
    }

    static bool Near(Rect a, Rect b) =>
        Math.Abs(a.X - b.X) < 0.01 && Math.Abs(a.Y - b.Y) < 0.01 && Math.Abs(a.Width - b.Width) < 0.01 && Math.Abs(a.Height - b.Height) < 0.01;

    // ------------------------------------------------------------------ traces

    static XColor? ColorOf(Brush? b, double opacity)
    {
        if (b is not SolidColorBrush sb) return null;
        var c = sb.Color;
        double a = c.A / 255.0 * sb.Opacity * opacity;
        return XColor.FromArgb((int)Math.Round(Math.Clamp(a, 0, 1) * 255), c.R, c.G, c.B);
    }

    void DrawGeometry(GeometryDrawing gd, double opacity)
    {
        if (gd.Geometry is null || gd.Geometry.Bounds.IsEmpty) return;
        bool hasFill = gd.Brush is not null, hasPen = gd.Pen is { Thickness: > 0, Brush: not null };
        if (!hasFill && !hasPen) return;
        // degrades, motifs : en image
        if ((hasFill && gd.Brush is not SolidColorBrush) || (hasPen && gd.Pen!.Brush is not SolidColorBrush)) { Rasterize(gd, opacity); return; }
        var path = ToPath(gd.Geometry);
        XBrush? brush = hasFill && ColorOf(gd.Brush, opacity) is { A: > 0 } fc ? new XSolidBrush(fc) : null;
        XPen? pen = hasPen && ColorOf(gd.Pen!.Brush, opacity) is { A: > 0 } pcol ? PenOf(gd.Pen!, pcol) : null;
        if (brush is null && pen is null) return;
        Stats.Paths++;
        if (pen is not null && brush is not null) _g.DrawPath(pen, brush, path);
        else if (pen is not null) _g.DrawPath(pen, path);
        else _g.DrawPath(brush!, path);
    }

    static XPen PenOf(Pen p, XColor color)
    {
        var pen = new XPen(color, p.Thickness)
        {
            LineCap = p.StartLineCap switch { PenLineCap.Round or PenLineCap.Triangle => XLineCap.Round, PenLineCap.Square => XLineCap.Square, _ => XLineCap.Flat },
            LineJoin = p.LineJoin switch { PenLineJoin.Round => XLineJoin.Round, PenLineJoin.Bevel => XLineJoin.Bevel, _ => XLineJoin.Miter },
            MiterLimit = p.MiterLimit,
        };
        if (p.DashStyle is { Dashes.Count: > 0 } ds)
        {
            // les tirets WPF sont exprimes en multiples de l'epaisseur, comme ceux de PDFsharp
            pen.DashStyle = XDashStyle.Custom;
            pen.DashPattern = ds.Dashes.Select(x => Math.Max(x, 0.001)).ToArray();
            pen.DashOffset = ds.Offset;
        }
        return pen;
    }

    /// <summary>Geometrie WPF -> trace PDFsharp (arcs convertis en courbes de Bezier, transformation appliquee aux points).</summary>
    public static XGraphicsPath ToPath(Geometry geo)
    {
        var pg = PathGeometry.CreateFromGeometry(geo);
        var m = pg.Transform?.Value ?? Matrix.Identity;
        var path = new XGraphicsPath { FillMode = pg.FillRule == FillRule.Nonzero ? XFillMode.Winding : XFillMode.Alternate };
        XPoint T(Point p) { var q = m.Transform(p); return new XPoint(q.X, q.Y); }
        foreach (var fig in pg.Figures)
        {
            if (fig.Segments.Count == 0) continue;
            path.StartFigure();
            var cur = fig.StartPoint;
            void Line(Point to) { path.AddLine(T(cur), T(to)); cur = to; }
            void Cubic(Point c1, Point c2, Point to) { path.AddBezier(T(cur), T(c1), T(c2), T(to)); cur = to; }
            void Quad(Point q, Point to) => Cubic(cur + (q - cur) * (2.0 / 3), to + (q - to) * (2.0 / 3), to);
            foreach (var seg in fig.Segments)
            {
                switch (seg)
                {
                    case LineSegment ls: Line(ls.Point); break;
                    case PolyLineSegment pl: foreach (var p in pl.Points) Line(p); break;
                    case BezierSegment bz: Cubic(bz.Point1, bz.Point2, bz.Point3); break;
                    case PolyBezierSegment pb:
                        for (int i = 0; i + 2 < pb.Points.Count; i += 3) Cubic(pb.Points[i], pb.Points[i + 1], pb.Points[i + 2]);
                        break;
                    case QuadraticBezierSegment qb: Quad(qb.Point1, qb.Point2); break;
                    case PolyQuadraticBezierSegment pq:
                        for (int i = 0; i + 1 < pq.Points.Count; i += 2) Quad(pq.Points[i], pq.Points[i + 1]);
                        break;
                    case ArcSegment arc:
                        foreach (var (c1, c2, to) in ArcToBeziers(cur, arc)) Cubic(c1, c2, to);
                        cur = arc.Point;
                        break;
                }
            }
            if (fig.IsClosed) path.CloseFigure();
        }
        return path;
    }

    /// <summary>Arc elliptique (forme WPF/SVG : extremites, rayons, rotation, drapeaux) en courbes de Bezier d'au plus 90 degres.</summary>
    public static List<(Point C1, Point C2, Point To)> ArcToBeziers(Point from, ArcSegment arc)
    {
        var res = new List<(Point, Point, Point)>();
        var to = arc.Point;
        double rx = Math.Abs(arc.Size.Width), ry = Math.Abs(arc.Size.Height);
        if ((from - to).Length < 1e-9) return res;
        if (rx < 1e-9 || ry < 1e-9) { res.Add((from, to, to)); return res; }
        double phi = arc.RotationAngle * Math.PI / 180, cos = Math.Cos(phi), sin = Math.Sin(phi);
        double dx = (from.X - to.X) / 2, dy = (from.Y - to.Y) / 2;
        double x1 = cos * dx + sin * dy, y1 = -sin * dx + cos * dy;
        double lambda = x1 * x1 / (rx * rx) + y1 * y1 / (ry * ry);
        if (lambda > 1) { var s = Math.Sqrt(lambda); rx *= s; ry *= s; }
        double num = rx * rx * ry * ry - rx * rx * y1 * y1 - ry * ry * x1 * x1;
        double den = rx * rx * y1 * y1 + ry * ry * x1 * x1;
        double coef = Math.Sqrt(Math.Max(0, num / den));
        bool sweep = arc.SweepDirection == SweepDirection.Clockwise;
        if (arc.IsLargeArc == sweep) coef = -coef;
        double cxp = coef * rx * y1 / ry, cyp = -coef * ry * x1 / rx;
        double cx = cos * cxp - sin * cyp + (from.X + to.X) / 2, cy = sin * cxp + cos * cyp + (from.Y + to.Y) / 2;
        static double Angle(double ux, double uy, double vx, double vy)
        {
            double a = Math.Atan2(ux * vy - uy * vx, ux * vx + uy * vy);
            return a;
        }
        double t1 = Angle(1, 0, (x1 - cxp) / rx, (y1 - cyp) / ry);
        double dt = Angle((x1 - cxp) / rx, (y1 - cyp) / ry, (-x1 - cxp) / rx, (-y1 - cyp) / ry);
        if (!sweep && dt > 0) dt -= 2 * Math.PI;
        else if (sweep && dt < 0) dt += 2 * Math.PI;
        int n = Math.Max(1, (int)Math.Ceiling(Math.Abs(dt) / (Math.PI / 2) - 1e-9));
        double step = dt / n, k = 4.0 / 3 * Math.Tan(step / 4);
        Point P(double t) => new(cx + rx * Math.Cos(t) * cos - ry * Math.Sin(t) * sin, cy + rx * Math.Cos(t) * sin + ry * Math.Sin(t) * cos);
        Vector D(double t) => new(-rx * Math.Sin(t) * cos - ry * Math.Cos(t) * sin, -rx * Math.Sin(t) * sin + ry * Math.Cos(t) * cos);
        double a0 = t1;
        for (int i = 0; i < n; i++)
        {
            double a1 = a0 + step;
            var p0 = P(a0); var p1 = i == n - 1 ? to : P(a1);
            res.Add((p0 + D(a0) * k, p1 - D(a1) * k, p1));
            a0 = a1;
        }
        return res;
    }

    // ------------------------------------------------------------------ texte

    void DrawGlyphs(GlyphRunDrawing d, double opacity)
    {
        var run = d.GlyphRun;
        if (run is null || run.GlyphIndices.Count == 0) return;
        if (ColorOf(d.ForegroundBrush, opacity) is not { } color) { Rasterize(d, opacity); return; }
        if (color.A == 0) return;
        var gt = run.GlyphTypeface;
        bool boldSim = (gt.StyleSimulations & StyleSimulations.BoldSimulation) != 0;
        string? family = boldSim || run.BidiLevel % 2 == 1 || run.IsSideways ? null : PdfFontResolver.Instance.Register(gt);
        if (family is null)
        {
            // gras simule, police non incorporable : contours vectoriels (meme dessin, mais plus du texte)
            Stats.TextAsOutlines++;
            var geo = run.BuildGeometry();
            if (!geo.Bounds.IsEmpty) _g.DrawPath(new XSolidBrush(color), ToPath(geo));
            return;
        }
        Stats.TextRuns++;
        var font = _fonts.Get(family, run.FontRenderingEmSize);
        var brush = new XSolidBrush(color);
        double size = run.FontRenderingEmSize;
        var glyphs = run.GlyphIndices;
        var adv = run.AdvanceWidths;
        var offs = run.GlyphOffsets;
        var chars = run.Characters;
        var cmap = run.ClusterMap;
        double x0 = run.BaselineOrigin.X, y0 = run.BaselineOrigin.Y;

        // grappes caractere(s) -> glyphe(s) ; chaque grappe 1:1 rejoint le segment courant si PDFsharp la placerait
        // au meme endroit que WPF (avance identique), sinon un nouveau segment commence a la position exacte
        var segText = new System.Text.StringBuilder();
        var segGlyphs = new List<ushort>();
        double segX = 0, expected = 0;
        void Flush()
        {
            if (segText.Length > 0 && segText.ToString().Any(c => !char.IsWhiteSpace(c)))
            {
                _pendingGlyphs = segGlyphs.Count == segText.Length ? segGlyphs.ToArray() : null;
                try { _g.DrawString(segText.ToString(), font, brush, segX, y0, XStringFormats.BaseLineLeft); }
                finally { _pendingGlyphs = null; }
            }
            segText.Clear(); segGlyphs.Clear();
        }
        int nChars = chars?.Count ?? 0;
        double x = x0;
        int c = 0, gi = 0;
        while (gi < glyphs.Count)
        {
            // limites de la grappe
            int c1 = c, g1 = gi + 1;
            if (cmap is not null && nChars > 0)
            {
                c1 = c + 1;
                while (c1 < nChars && cmap[c1] == cmap[c]) c1++;
                g1 = c1 < nChars ? cmap[c1] : glyphs.Count;
                if (g1 <= gi) g1 = gi + 1;
            }
            else c1 = Math.Min(c + 1, nChars);
            double clusterAdv = 0;
            for (int k = gi; k < g1; k++) clusterAdv += adv[k];
            bool offset = offs is not null && Enumerable.Range(gi, g1 - gi).Any(k => Math.Abs(offs[k].X) > 1e-4 || Math.Abs(offs[k].Y) > 1e-4);
            string text = nChars > 0 && c1 > c ? string.Concat(Enumerable.Range(c, c1 - c).Select(i => chars![i])) : "";
            bool simple = text.Length == 1 && g1 - gi == 1 && !offset;
            if (simple)
            {
                ushort glyph = glyphs[gi];
                double natural = Math.Round(gt.AdvanceWidths.TryGetValue(glyph, out var aw) ? aw * 1000 : 0) / 1000 * size;
                if (segText.Length == 0 || Math.Abs(x - expected) > 0.02) { Flush(); segX = x; expected = x; }
                segText.Append(text);
                segGlyphs.Add(glyph);
                expected += natural;
            }
            else
            {
                // ligature, caractere hors plan de base, decalage : grappe seule, a sa position exacte
                Flush();
                if (text.Length > 0)
                {
                    double ox = offset ? offs![gi].X : 0, oy = offset ? offs![gi].Y : 0;
                    segX = x + ox;
                    segText.Append(text);
                    if (text.Length == g1 - gi) segGlyphs.AddRange(Enumerable.Range(gi, g1 - gi).Select(k => glyphs[k]));
                    double save = y0;
                    y0 -= oy;
                    Flush();
                    y0 = save;
                }
            }
            x += clusterAdv;
            gi = g1;
            c = c1;
        }
        Flush();
    }

    // ------------------------------------------------------------------ images

    void DrawImage(ImageDrawing id, double opacity)
    {
        if (id.ImageSource is not BitmapSource bmp || id.Rect.IsEmpty) { if (id.ImageSource is not null) Rasterize(id, opacity); return; }
        var img = _images.Get(bmp, opacity);
        if (img is null) { Rasterize(id, opacity); return; }
        Stats.Images++;
        var r = id.Rect;
        _g.DrawImage(img, r.X, r.Y, r.Width, r.Height);
    }

    /// <summary>Rasterise une partie du dessin (300 ppp) et la place a sa position.</summary>
    void Rasterize(Drawing d, double opacity)
    {
        var b = d.Bounds;
        if (b.IsEmpty || b.Width <= 0 || b.Height <= 0) return;
        if (PageClip is { } pc) b.Intersect(pc.Replacement);
        if (b.IsEmpty || b.Width <= 0 || b.Height <= 0) return;
        double scale = RasterDpi / 72;
        const int MaxSide = 7000;
        scale = Math.Min(scale, MaxSide / Math.Max(b.Width, b.Height));
        int w = Math.Max(1, (int)Math.Ceiling(b.Width * scale)), h = Math.Max(1, (int)Math.Ceiling(b.Height * scale));
        var dv = new DrawingVisual();
        using (var dc = dv.RenderOpen())
        {
            dc.PushTransform(new ScaleTransform(scale, scale));
            dc.PushTransform(new TranslateTransform(-b.X, -b.Y));
            if (opacity < 1) dc.PushOpacity(opacity);
            dc.DrawDrawing(d);
        }
        RenderOptions.SetBitmapScalingMode(dv, BitmapScalingMode.HighQuality);
        var rtb = new RenderTargetBitmap(w, h, 96, 96, PixelFormats.Pbgra32);
        rtb.Render(dv);
        rtb.Freeze();
        var img = _images.FromRendered(rtb);
        Stats.Rasterized++;
        _g.DrawImage(img, b.X, b.Y, w / scale, h / scale);
    }
}

/// <summary>
/// Images du PDF : le fichier d'origine quand c'est possible (JPEG recopie tel quel, sans recompression), sinon
/// l'image decodee en PNG sans perte. Une meme image utilisee plusieurs fois n'est incorporee qu'une fois.
/// </summary>
public sealed class PdfImageCache
{
    readonly Dictionary<(BitmapSource, int), XImage?> _cache = new();
    public int Originals { get; private set; }

    /// <summary>
    /// Octets d'origine d'une image (OpenWrite : images lues dans le paquet .docx, sans fichier) : un JPEG ou un PNG
    /// est incorporé tel quel plutôt que décodé puis réencodé.
    /// </summary>
    public Func<BitmapSource, byte[]?>? Original { get; init; }

    public XImage? Get(BitmapSource bmp, double opacity)
    {
        int op = (int)Math.Round(Math.Clamp(opacity, 0, 1) * 1000);
        if (_cache.TryGetValue((bmp, op), out var x)) return x;
        try
        {
            if (op >= 1000 && bmp is BitmapImage { UriSource: { IsFile: true } uri } && File.Exists(uri.LocalPath))
            {
                var ext = Path.GetExtension(uri.LocalPath).ToLowerInvariant();
                if (ext is ".jpg" or ".jpeg" or ".png")
                {
                    try
                    {
                        x = XImage.FromStream(new MemoryStream(File.ReadAllBytes(uri.LocalPath)));
                        Originals++;
                    }
                    catch (Exception) { x = null; }
                }
            }
            if (x is null && op >= 1000 && Original?.Invoke(bmp) is { Length: > 8 } bytes && IsJpegOrPng(bytes))
            {
                try
                {
                    x = XImage.FromStream(new MemoryStream(bytes));
                    Originals++;
                }
                catch (Exception) { x = null; }
            }
            x ??= FromBitmap(op >= 1000 ? bmp : WithOpacity(bmp, op / 1000.0));
        }
        catch (Exception) { x = null; }
        _cache[(bmp, op)] = x;
        return x;
    }

    static bool IsJpegOrPng(byte[] b) =>
        (b[0] == 0xFF && b[1] == 0xD8) || (b[0] == 0x89 && b[1] == (byte)'P' && b[2] == (byte)'N' && b[3] == (byte)'G');

    public XImage FromRendered(BitmapSource bmp) => FromBitmap(bmp);

    static XImage FromBitmap(BitmapSource bmp)
    {
        // PNG 8 bits par composante (avec alpha si l'image en a), lisible par PDFsharp
        BitmapSource src = bmp.Format == PixelFormats.Bgra32 || bmp.Format == PixelFormats.Bgr24 ? bmp
            : HasAlpha(bmp.Format) ? new FormatConvertedBitmap(bmp, PixelFormats.Bgra32, null, 0) : new FormatConvertedBitmap(bmp, PixelFormats.Bgr24, null, 0);
        var enc = new PngBitmapEncoder();
        enc.Frames.Add(BitmapFrame.Create(src));
        var ms = new MemoryStream();
        enc.Save(ms);
        ms.Position = 0;
        return XImage.FromStream(ms);
    }

    static bool HasAlpha(PixelFormat f) =>
        f == PixelFormats.Bgra32 || f == PixelFormats.Pbgra32 || f == PixelFormats.Prgba64 || f == PixelFormats.Rgba64 ||
        f == PixelFormats.Prgba128Float || f == PixelFormats.Rgba128Float || f == PixelFormats.Indexed8 || f == PixelFormats.Indexed4 ||
        f == PixelFormats.Indexed2 || f == PixelFormats.Indexed1;

    /// <summary>Transparence d'une image : alpha multiplie dans les pixels (aucun reechantillonnage).</summary>
    static BitmapSource WithOpacity(BitmapSource bmp, double opacity)
    {
        var conv = new FormatConvertedBitmap(bmp, PixelFormats.Bgra32, null, 0);
        int w = conv.PixelWidth, h = conv.PixelHeight, stride = w * 4;
        var px = new byte[h * stride];
        conv.CopyPixels(px, stride, 0);
        for (int i = 3; i < px.Length; i += 4) px[i] = (byte)Math.Round(px[i] * opacity);
        var res = BitmapSource.Create(w, h, conv.DpiX, conv.DpiY, PixelFormats.Bgra32, null, px, stride);
        res.Freeze();
        return res;
    }
}
