using System.Globalization;
using System.Windows;
using System.Windows.Media;
using OpenWrite.Core.Drawing;
using PathSegment = OpenWrite.Core.Drawing.PathSegment;
using Visual = OpenWrite.Core.Drawing.Visual;

namespace OpenWrite.App.Layout;

/// <summary>
/// Dessin des objets graphiques (lot 7) : images (rognage, forme de découpe, niveaux de gris, transparence), formes
/// (prédéfinies DrawingML, géométries libres, chemins VML ; remplissage uni, dégradé ou image ; trait, tirets, trait
/// double, flèches), groupes et zones de dessin, objets de repli (aperçu). Rotation et retournements autour du centre de
/// chaque élément ; le texte des zones de texte est dessiné dans le même repère (il tourne avec la forme).
/// </summary>
public static class DrawingRenderer
{
    /// <param name="image">Image d'un identifiant de relation (r:embed) de la partie du dessin.</param>
    /// <param name="text">Dessine le texte d'une zone de texte (indice dans DrawingContent.TextBoxes), dans le repère courant.</param>
    public static void Draw(DrawingContext dc, DrawingObject o, double x, double y, Func<string, ImageSource?> image, Action<int>? text = null)
    {
        DrawVisual(dc, o.Root, x, y, image, text);
    }

    static void DrawVisual(DrawingContext dc, Visual v, double ox, double oy, Func<string, ImageSource?> image, Action<int>? text)
    {
        var rect = new Rect(ox + v.X, oy + v.Y, Math.Max(0, v.W), Math.Max(0, v.H));
        int pushed = 0;
        var center = new Point(rect.X + rect.Width / 2, rect.Y + rect.Height / 2);
        if (v.Rotation != 0) { dc.PushTransform(new RotateTransform(v.Rotation, center.X, center.Y)); pushed++; }
        if (v.FlipH || v.FlipV) { dc.PushTransform(new ScaleTransform(v.FlipH ? -1 : 1, v.FlipV ? -1 : 1, center.X, center.Y)); pushed++; }
        try
        {
            switch (v)
            {
                case PictureVisual p: DrawPicture(dc, p, rect, image); break;
                case ShapeVisual s: DrawShape(dc, s, rect, image, text); break;
                case GroupVisual g:
                    if (g.Fill is { } gf && Brush(gf, rect, image) is { } gb) dc.DrawRectangle(gb, null, rect);
                    foreach (var c in g.Children) DrawVisual(dc, c, rect.X, rect.Y, image, text);
                    break;
                case FallbackVisual f: DrawFallback(dc, f, rect, image); break;
            }
        }
        finally { for (int i = 0; i < pushed; i++) dc.Pop(); }
    }

    // ---------------------------------------------------------------- images

    static void DrawPicture(DrawingContext dc, PictureVisual p, Rect rect, Func<string, ImageSource?> image)
    {
        if (rect.Width <= 0 || rect.Height <= 0) return;
        var img = p.ImageId is { } id ? image(id) : null;
        var clip = p.Geometry != "rect" ? Geometry(p.Geometry, null, null, null, rect) : null;
        if (clip is not null) dc.PushClip(clip);
        else dc.PushClip(new RectangleGeometry(rect));
        if (p.Alpha < 1) dc.PushOpacity(p.Alpha);
        if (img is not null)
        {
            if (p.Grayscale && img is System.Windows.Media.Imaging.BitmapSource bs)
                img = Freeze(new System.Windows.Media.Imaging.FormatConvertedBitmap(bs, PixelFormats.Gray8, null, 0));
            // rognage : l'image entière est posée plus grande que le cadre, le cadre la découpe
            var (l, t, r, b) = p.Crop;
            double fw = 1 - l - r, fh = 1 - t - b;
            if (fw <= 0.001) fw = 1;
            if (fh <= 0.001) fh = 1;
            double tw = rect.Width / fw, th = rect.Height / fh;
            dc.DrawImage(img, new Rect(rect.X - l * tw, rect.Y - t * th, tw, th));
            if (p.Brightness != 0 || p.Contrast != 0)
            {
                // luminosité : voile blanc ou noir (approximation de la correction de Word)
                double a = Math.Clamp(Math.Abs(p.Brightness), 0, 1);
                if (a > 0) dc.DrawRectangle(Freeze(new SolidColorBrush(Color.FromArgb((byte)(a * 200), (byte)(p.Brightness > 0 ? 255 : 0), (byte)(p.Brightness > 0 ? 255 : 0), (byte)(p.Brightness > 0 ? 255 : 0)))), null, rect);
            }
        }
        else dc.DrawRectangle(Placeholder, PlaceholderPen, rect);
        if (p.Alpha < 1) dc.Pop();
        dc.Pop();
        if (p.Line is { } line && Pen(line) is { } pen)
            dc.DrawGeometry(null, pen, clip ?? new RectangleGeometry(rect));
    }

    static readonly Brush Placeholder = Freeze(new SolidColorBrush(Color.FromRgb(0xEE, 0xF1, 0xF6)));
    static readonly Pen PlaceholderPen = Freeze(new Pen(Brushes.LightGray, 0.5));

    static void DrawFallback(DrawingContext dc, FallbackVisual f, Rect rect, Func<string, ImageSource?> image)
    {
        if (f.ImageId is { } id && image(id) is { } img) { dc.DrawImage(img, rect); return; }
        dc.DrawRectangle(Placeholder, PlaceholderPen, rect);
        var label = f.Kind switch { "chart" => "Graphique", "smartart" => "SmartArt", "ole" => "Objet", _ => "Objet" };
        var ft = new FormattedText(label, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, new Typeface("Segoe UI"), 9, Brushes.Gray, 1.0);
        if (ft.Width < rect.Width && ft.Height < rect.Height)
            dc.DrawText(ft, new Point(rect.X + (rect.Width - ft.Width) / 2, rect.Y + (rect.Height - ft.Height) / 2));
    }

    // ---------------------------------------------------------------- formes

    static void DrawShape(DrawingContext dc, ShapeVisual s, Rect rect, Func<string, ImageSource?> image, Action<int>? text)
    {
        List<ShapePath> paths;
        if (s.Paths is not null) paths = s.Paths;
        else if (s.Custom is not null) paths = ShapeGeometry.Evaluate(s.Custom, rect.Width, rect.Height).Paths;
        else paths = ShapeGeometry.PresetPaths(s.Preset ?? "rect", rect.Width, rect.Height, s.Adjust);
        var fill = s.Fill is { } f ? Brush(f, rect, image) : null;
        var pen = s.Line is { } l ? Pen(l) : null;
        foreach (var path in paths)
        {
            var geo = Build(path.Segments, rect.X, rect.Y, path.Fill != "none");
            if (geo is null) continue;
            Brush? b = path.Fill switch
            {
                "none" => null,
                "darken" => Modulate(fill, 0.6),
                "darkenLess" => Modulate(fill, 0.8),
                "lighten" => Modulate(fill, 1.4),
                "lightenLess" => Modulate(fill, 1.2),
                _ => fill,
            };
            if (b is not null) dc.DrawGeometry(b, null, geo);
            if (pen is not null && path.Stroke)
            {
                if (s.Line!.Compound is "dbl" or "thickThin" or "thinThick" or "tri")
                {
                    // trait composé : le trait entier, évidé au milieu (deux traits fins de part et d'autre)
                    dc.DrawGeometry(null, pen, geo);
                    dc.DrawGeometry(null, Pen(s.Line with { Width = s.Line.Width / 3, Color = "FFFFFF" })!, geo);
                }
                else dc.DrawGeometry(null, pen, geo);
                if (s.Line.HeadEnd is not null || s.Line.TailEnd is not null) DrawArrows(dc, path.Segments, rect, s.Line, pen);
            }
        }
        if (s.WordArt is { } wa && wa.Text.Length > 0 && s.Fill is not { Kind: "none" })
        {
            var brush = fill ?? Brushes.Gray;
            var ft = new FormattedText(wa.Text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, new Typeface(wa.Font), 100, brush, 1.0);
            dc.PushTransform(new TranslateTransform(rect.X, rect.Y));
            dc.PushTransform(new ScaleTransform(rect.Width / Math.Max(1, ft.Width), rect.Height / Math.Max(1, ft.Height)));
            dc.DrawText(ft, new Point(0, 0));
            dc.Pop(); dc.Pop();
        }
        if (s.TextBox >= 0 && text is not null)
        {
            dc.PushClip(new RectangleGeometry(new Rect(rect.X - 1, rect.Y - 1, rect.Width + 2, rect.Height + 2)));
            text(s.TextBox);
            dc.Pop();
        }
    }

    /// <summary>Géométrie WPF d'une forme prédéfinie (découpe d'une image).</summary>
    static System.Windows.Media.Geometry? Geometry(string preset, Dictionary<string, string>? adjust, object? _, object? __, Rect rect)
    {
        var group = new GeometryGroup { FillRule = FillRule.Nonzero };
        foreach (var p in ShapeGeometry.PresetPaths(preset, rect.Width, rect.Height, adjust))
            if (Build(p.Segments, rect.X, rect.Y, true) is { } g) group.Children.Add(g);
        group.Freeze();
        return group.Children.Count > 0 ? group : null;
    }

    static StreamGeometry? Build(List<PathSegment> segs, double ox, double oy, bool filled) =>
        segs.Count == 0 ? null : Closed(segs, ox, oy, filled);

    static StreamGeometry Closed(List<PathSegment> segs, double ox, double oy, bool filled)
    {
        // découpe en figures (M … Z) et ferme celles qui finissent par Z
        var geo = new StreamGeometry { FillRule = FillRule.EvenOdd };
        using (var ctx = geo.Open())
        {
            var figure = new List<PathSegment>();
            void Emit(bool closed)
            {
                if (figure.Count == 0) return;
                Point P(double x, double y) => new(ox + x, oy + y);
                var first = figure[0];
                var start = first.Op switch { 'C' => P(first.Points[4], first.Points[5]), 'Q' => P(first.Points[2], first.Points[3]), _ => P(first.Points[0], first.Points[1]) };
                ctx.BeginFigure(start, filled, closed);
                foreach (var s in figure.Skip(first.Op == 'M' ? 1 : 0))
                {
                    var a = s.Points;
                    switch (s.Op)
                    {
                        case 'L': ctx.LineTo(P(a[0], a[1]), true, true); break;
                        case 'C': ctx.BezierTo(P(a[0], a[1]), P(a[2], a[3]), P(a[4], a[5]), true, true); break;
                        case 'Q': ctx.QuadraticBezierTo(P(a[0], a[1]), P(a[2], a[3]), true, true); break;
                    }
                }
                figure.Clear();
            }
            foreach (var s in segs)
            {
                if (s.Op == 'Z') { Emit(true); continue; }
                if (s.Op == 'M' && figure.Count > 0) Emit(false);
                figure.Add(s);
            }
            Emit(false);
        }
        geo.Freeze();
        return geo;
    }

    /// <summary>Pointes de flèche (triangle, flèche ouverte, losange, ovale) aux extrémités d'un trait.</summary>
    static void DrawArrows(DrawingContext dc, List<PathSegment> segs, Rect rect, LineSpec line, Pen pen)
    {
        var pts = new List<Point>();
        foreach (var s in segs)
        {
            var a = s.Points;
            for (int i = 0; i + 1 < a.Length; i += 2) pts.Add(new Point(rect.X + a[i], rect.Y + a[i + 1]));
        }
        if (pts.Count < 2) return;
        var brush = pen.Brush;
        double size = Math.Max(3, line.Width * 3);
        void Head(Point tip, Point from, string kind)
        {
            var d = tip - from;
            if (d.Length < 1e-6) return;
            d.Normalize();
            var n = new Vector(-d.Y, d.X);
            var basePt = tip - d * size;
            var g = new StreamGeometry();
            using (var ctx = g.Open())
            {
                if (kind is "oval")
                {
                    dc.DrawEllipse(brush, null, tip, size / 2, size / 2);
                    return;
                }
                if (kind is "diamond")
                {
                    ctx.BeginFigure(tip, true, true);
                    ctx.LineTo(tip - d * size / 2 + n * size / 2, true, true);
                    ctx.LineTo(tip - d * size, true, true);
                    ctx.LineTo(tip - d * size / 2 - n * size / 2, true, true);
                }
                else
                {
                    bool open = kind == "arrow";
                    ctx.BeginFigure(basePt + n * size / 2, !open, !open);
                    ctx.LineTo(tip, true, true);
                    ctx.LineTo(basePt - n * size / 2, true, true);
                }
            }
            g.Freeze();
            if (kind == "arrow") dc.DrawGeometry(null, pen, g);
            else dc.DrawGeometry(brush, null, g);
        }
        if (line.HeadEnd is { } h) Head(pts[0], pts[1], h);
        if (line.TailEnd is { } t) Head(pts[^1], pts[^2], t);
    }

    // ---------------------------------------------------------------- pinceaux et traits

    static Brush? Brush(FillSpec f, Rect rect, Func<string, ImageSource?> image)
    {
        switch (f.Kind)
        {
            case "none": return null;
            case "solid": return f.Color is { } c ? Freeze(new SolidColorBrush(ColorOf(c, f.Alpha))) : null;
            case "picture":
                return f.ImageId is { } id && image(id) is { } img ? Freeze(new ImageBrush(img) { Stretch = Stretch.Fill }) : null;
            case "gradient" when f.Stops is { Count: > 0 } stops:
            {
                var gs = new GradientStopCollection(stops.Select(s => new GradientStop(ColorOf(s.Color, s.Alpha), s.Pos)));
                if (f.Radial) return Freeze(new RadialGradientBrush(gs));
                // angle DrawingML : 0 = de gauche à droite, sens horaire
                double a = f.Angle * Math.PI / 180;
                var dir = new Vector(Math.Cos(a), Math.Sin(a));
                return Freeze(new LinearGradientBrush(gs, new Point(0.5 - dir.X / 2, 0.5 - dir.Y / 2), new Point(0.5 + dir.X / 2, 0.5 + dir.Y / 2)));
            }
        }
        return f.Color is { } col ? Freeze(new SolidColorBrush(ColorOf(col, f.Alpha))) : null;
    }

    static Brush? Modulate(Brush? b, double factor)
    {
        if (b is not SolidColorBrush s) return b;
        var c = s.Color;
        byte M(byte v) => (byte)Math.Clamp(factor > 1 ? v + (255 - v) * (factor - 1) : v * factor, 0, 255);
        return Freeze(new SolidColorBrush(Color.FromArgb(c.A, M(c.R), M(c.G), M(c.B))));
    }

    static readonly Dictionary<LineSpec, Pen?> Pens = new();

    static Pen? Pen(LineSpec l)
    {
        if (Pens.TryGetValue(l, out var cached)) return cached;
        var pen = new Pen(Freeze(new SolidColorBrush(ColorOf(l.Color, l.Alpha))), Math.Max(0.25, l.Width))
        {
            LineJoin = PenLineJoin.Round,
            DashStyle = l.Dash switch
            {
                "sysDot" => new DashStyle(new double[] { 1, 1 }, 0),
                "dot" => new DashStyle(new double[] { 1, 3 }, 0),
                "dash" => new DashStyle(new double[] { 4, 3 }, 0),
                "sysDash" => new DashStyle(new double[] { 3, 1 }, 0),
                "lgDash" => new DashStyle(new double[] { 8, 3 }, 0),
                "dashDot" => new DashStyle(new double[] { 4, 3, 1, 3 }, 0),
                "sysDashDot" => new DashStyle(new double[] { 3, 1, 1, 1 }, 0),
                "lgDashDot" => new DashStyle(new double[] { 8, 3, 1, 3 }, 0),
                "lgDashDotDot" or "sysDashDotDot" => new DashStyle(new double[] { 8, 3, 1, 3, 1, 3 }, 0),
                _ => DashStyles.Solid,
            },
        };
        pen.Freeze();
        Pens[l] = pen;
        return pen;
    }

    static Color ColorOf(string rgb, double alpha)
    {
        int v = int.TryParse(rgb, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var n) ? n : 0;
        return Color.FromArgb((byte)Math.Round(Math.Clamp(alpha, 0, 1) * 255), (byte)(v >> 16), (byte)(v >> 8), (byte)v);
    }

    static T Freeze<T>(T f) where T : Freezable { f.Freeze(); return f; }
}
