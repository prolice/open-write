using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace OpenWrite.Core.Drawing;

// VML (w:pict, w:object, mc:Fallback) : documents enregistrés par Word 2007 et avant, ou par d'autres programmes
// (Horaires : traits v:line, groupes v:group en ligne). Le style CSS de l'élément porte la position et la taille.
public static partial class DrawingReader
{
    static readonly HashSet<string> VmlShapes = new() { "shape", "rect", "roundrect", "oval", "line", "polyline", "image", "group", "arc", "curve" };

    static DrawingObject? ReadVml(XElement pict, Ctx ctx)
    {
        var types = pict.Elements(V + "shapetype").Where(t => t.Attribute("id") is not null)
            .GroupBy(t => t.Attribute("id")!.Value).ToDictionary(g => g.Key, g => g.First());
        var shape = pict.Elements().FirstOrDefault(e => e.Name.Namespace == V && VmlShapes.Contains(e.Name.LocalName));
        if (shape is null) return null;
        var style = Css(shape.Attribute("style")?.Value);
        bool absolute = style.GetValueOrDefault("position") == "absolute";
        bool ole = pict.Name.LocalName == "object";

        // taille et position (points) ; un trait se mesure par ses extrémités
        double w = Len(style.GetValueOrDefault("width")), h = Len(style.GetValueOrDefault("height"));
        double left = Len(style.GetValueOrDefault("margin-left")) + Len(style.GetValueOrDefault("left"));
        double top = Len(style.GetValueOrDefault("margin-top")) + Len(style.GetValueOrDefault("top"));
        if (shape.Name.LocalName == "line")
        {
            var (x1, y1) = Point(shape.Attribute("from")?.Value);
            var (x2, y2) = Point(shape.Attribute("to")?.Value);
            left += Math.Min(x1, x2);
            top += Math.Min(y1, y2);
            w = Math.Abs(x2 - x1);
            h = Math.Abs(y2 - y1);
        }
        var root = ole ? OleVisual(shape, w, h) : VmlVisual(shape, types, 0, 0, w, h, null, ctx);
        if (root is null) return null;

        var wrapEl = shape.Element(W10 + "wrap");
        bool watermark = shape.Element(V + "textpath") is not null
            && ((shape.Attribute("id")?.Value ?? "") + (shape.Attribute(O + "spid")?.Value ?? "")).Contains("WaterMark", StringComparison.OrdinalIgnoreCase);
        double z = D(style.GetValueOrDefault("z-index"));
        var distance = (Len(style.GetValueOrDefault("mso-wrap-distance-top")), Len(style.GetValueOrDefault("mso-wrap-distance-bottom")),
                        Len(style.GetValueOrDefault("mso-wrap-distance-left"), 9), Len(style.GetValueOrDefault("mso-wrap-distance-right"), 9));
        if (!absolute)
            return new DrawingObject
            {
                Source = ctx.Source, Inline = true, Width = w, Height = h, Root = root, Vml = true,
                Name = shape.Attribute("id")?.Value, Description = shape.Attribute("alt")?.Value,
                Hidden = style.GetValueOrDefault("visibility") == "hidden",
            };

        var wrap = (wrapEl?.Attribute("type")?.Value) switch
        {
            "square" => WrapKind.Square,
            "tight" => WrapKind.Tight,
            "through" => WrapKind.Through,
            "topAndBottom" => WrapKind.TopAndBottom,
            _ => WrapKind.None,
        };
        var side = wrapEl?.Attribute("side")?.Value switch { "left" => "left", "right" => "right", "largest" => "largest", _ => "bothSides" };
        string HFrom(string? v) => v switch
        {
            "page" => "page", "margin" => "margin", "char" => "character", "left-margin-area" => "leftMargin",
            "right-margin-area" => "rightMargin", "inner-margin-area" => "insideMargin", "outer-margin-area" => "outsideMargin",
            _ => "column",
        };
        string VFrom(string? v) => v switch
        {
            "page" => "page", "margin" => "margin", "line" => "line", "top-margin-area" => "topMargin",
            "bottom-margin-area" => "bottomMargin", "inner-margin-area" => "insideMargin", "outer-margin-area" => "outsideMargin",
            _ => "paragraph",
        };
        string? Align(string? v) => v is null or "absolute" ? null : v;
        return new DrawingObject
        {
            Source = ctx.Source, Inline = false, Width = w, Height = h, Root = root, Vml = true, Watermark = watermark,
            H = new AnchorAxis(HFrom(style.GetValueOrDefault("mso-position-horizontal-relative")), Align(style.GetValueOrDefault("mso-position-horizontal")), left),
            V = new AnchorAxis(VFrom(style.GetValueOrDefault("mso-position-vertical-relative")), Align(style.GetValueOrDefault("mso-position-vertical")), top),
            Wrap = wrap, WrapSide = side, Distance = distance,
            BehindText = wrap == WrapKind.None && z < 0,
            Z = (long)z,
            AllowOverlap = shape.Attribute(O + "allowoverlap")?.Value is not ("f" or "false"),
            LayoutInCell = shape.Attribute(O + "allowincell")?.Value is not ("f" or "false"),
            Name = shape.Attribute("id")?.Value, Description = shape.Attribute("alt")?.Value,
            Hidden = style.GetValueOrDefault("visibility") == "hidden",
        };
    }

    /// <summary>Objet OLE (w:object) : son aperçu (v:imagedata).</summary>
    static Visual OleVisual(XElement shape, double w, double h) => new FallbackVisual
    {
        Kind = "ole", W = w, H = h,
        ImageId = shape.Element(V + "imagedata")?.Attribute(R + "id")?.Value,
    };

    /// <summary>
    /// Élément VML dans le cadre (x, y, w, h) du parent ; <paramref name="group"/> : repère du groupe parent
    /// (origine et taille des coordonnées, taille réelle) pour les enfants d'un v:group.
    /// </summary>
    static Visual? VmlVisual(XElement e, Dictionary<string, XElement> types, double x, double y, double w, double h,
        (double Ox, double Oy, double Cw, double Ch, double W, double H)? group, Ctx ctx)
    {
        var style = Css(e.Attribute("style")?.Value);
        if (group is { } g)
        {
            // coordonnées du groupe (sans unité) → points
            double sx = g.Cw != 0 ? g.W / g.Cw : 1, sy = g.Ch != 0 ? g.H / g.Ch : 1;
            double gl = D(Num(style.GetValueOrDefault("left"))), gt = D(Num(style.GetValueOrDefault("top")));
            double gw = D(Num(style.GetValueOrDefault("width"))), gh = D(Num(style.GetValueOrDefault("height")));
            if (e.Name.LocalName == "line")
            {
                var (x1, y1) = PointRaw(e.Attribute("from")?.Value);
                var (x2, y2) = PointRaw(e.Attribute("to")?.Value);
                gl = Math.Min(x1, x2); gt = Math.Min(y1, y2); gw = Math.Abs(x2 - x1); gh = Math.Abs(y2 - y1);
            }
            (x, y, w, h) = ((gl - g.Ox) * sx, (gt - g.Oy) * sy, gw * sx, gh * sy);
        }
        var rotText = style.GetValueOrDefault("rotation");
        double rot = rotText is null ? 0 : rotText.EndsWith("fd") ? D(rotText[..^2]) / 65536 : D(Num(rotText));
        var flip = style.GetValueOrDefault("flip") ?? "";

        if (e.Name.LocalName == "group")
        {
            var gv = new GroupVisual { X = x, Y = y, W = w, H = h, Rotation = rot, FlipH = flip.Contains('x'), FlipV = flip.Contains('y') };
            var (ox, oy) = PointRaw(e.Attribute("coordorigin")?.Value);
            var (cw, ch) = PointRaw(e.Attribute("coordsize")?.Value ?? "1000,1000");
            foreach (var child in e.Elements().Where(c => c.Name.Namespace == V && VmlShapes.Contains(c.Name.LocalName)))
                if (VmlVisual(child, types, 0, 0, 0, 0, (ox, oy, cw, ch, w, h), ctx) is { } v) gv.Children.Add(v);
            return gv;
        }

        // image (v:image, ou v:shape de type « cadre d'image » avec v:imagedata)
        var img = e.Element(V + "imagedata");
        var imageId = e.Name.LocalName == "image" ? e.Attribute(R + "id")?.Value ?? e.Attribute("src")?.Value : img?.Attribute(R + "id")?.Value ?? img?.Attribute(O + "relid")?.Value;
        var typeEl = e.Attribute("type")?.Value is { } tref && types.TryGetValue(tref.TrimStart('#'), out var te) ? te : null;
        string spt = e.Attribute(O + "spt")?.Value ?? typeEl?.Attribute(O + "spt")?.Value ?? "";
        if (imageId is not null)
        {
            var src = e.Name.LocalName == "image" ? e : img!;
            return new PictureVisual
            {
                ImageId = imageId, X = x, Y = y, W = w, H = h, Rotation = rot, FlipH = flip.Contains('x'), FlipV = flip.Contains('y'),
                Crop = (Crop(src.Attribute("cropleft")?.Value), Crop(src.Attribute("croptop")?.Value), Crop(src.Attribute("cropright")?.Value), Crop(src.Attribute("cropbottom")?.Value)),
                Line = Stroke(e, typeEl, defaultOn: false),
                Grayscale = src.Attribute("grayscale")?.Value is "t" or "true",
            };
        }

        // géométrie
        string? preset = null;
        Dictionary<string, string>? adjust = null;
        List<ShapePath>? paths = null;
        bool isLine = false;
        switch (e.Name.LocalName)
        {
            case "rect": preset = "rect"; break;
            case "oval": preset = "ellipse"; break;
            case "roundrect":
                preset = "roundRect";
                var arc = e.Attribute("arcsize")?.Value;
                double frac = arc is null ? 0.2 : arc.EndsWith('f') ? D(arc[..^1]) / 65536 : D(arc);
                adjust = new() { ["adj"] = $"val {(int)Math.Round(frac * 100000)}" };
                break;
            case "line":
            {
                isLine = true;
                // extrémités dans le cadre : le trait va d'un coin à l'autre selon le sens d'origine
                var (x1, y1) = group is null ? Point(e.Attribute("from")?.Value) : PointRaw(e.Attribute("from")?.Value);
                var (x2, y2) = group is null ? Point(e.Attribute("to")?.Value) : PointRaw(e.Attribute("to")?.Value);
                double ax = x1 <= x2 ? 0 : w, bx = x1 <= x2 ? w : 0, ay = y1 <= y2 ? 0 : h, by = y1 <= y2 ? h : 0;
                paths = new() { new ShapePath(new() { new('M', new[] { ax, ay }), new('L', new[] { bx, by }) }, "none", true) };
                break;
            }
            case "polyline":
            {
                var pts = (e.Attribute("points")?.Value ?? "").Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries).Select(s => Len(s)).ToList();
                if (pts.Count >= 4)
                {
                    double minX = Enumerable.Range(0, pts.Count / 2).Min(i => pts[2 * i]), minY = Enumerable.Range(0, pts.Count / 2).Min(i => pts[2 * i + 1]);
                    var segs = new List<PathSegment>();
                    for (int i = 0; i + 1 < pts.Count; i += 2) segs.Add(new(i == 0 ? 'M' : 'L', new[] { pts[i] - minX, pts[i + 1] - minY }));
                    paths = new() { new ShapePath(segs, On(e.Attribute("filled")?.Value) ? "norm" : "none", true) };
                }
                isLine = true;
                break;
            }
            default:
                preset = spt switch
                {
                    "1" or "75" or "202" or "" => "rect",
                    "2" => "roundRect", "3" => "ellipse", "4" => "diamond", "5" => "triangle", "6" => "rtTriangle",
                    "7" => "parallelogram", "8" => "trapezoid", "9" => "hexagon", "10" => "octagon", "11" => "plus",
                    "12" => "star5", "13" => "rightArrow", "15" => "homePlate", "16" => "cube", "20" => "line",
                    "32" => "straightConnector1", "55" => "chevron", "56" => "flowChartTerminator", "66" => "leftArrow",
                    "67" => "downArrow", "68" => "upArrow", "69" => "leftRightArrow", "70" => "upDownArrow",
                    "109" => "flowChartProcess", "110" => "flowChartDecision", "116" => "flowChartPredefinedProcess",
                    "136" => "rect", "183" => "sun", "184" => "moon", "96" => "smileyFace", "74" => "heart",
                    _ => null,
                };
                isLine = spt is "20" or "32";
                if (preset is null || spt is "" && (e.Attribute("path") ?? typeEl?.Attribute("path")) is not null)
                {
                    // chemin VML (sans formules) dans le repère coordsize
                    var path = e.Attribute("path")?.Value ?? typeEl?.Attribute("path")?.Value;
                    var (cw, ch) = PointRaw(e.Attribute("coordsize")?.Value ?? typeEl?.Attribute("coordsize")?.Value ?? "21600,21600");
                    if (path is not null && !path.Contains('@') && VmlPath(path, w / Math.Max(1, cw), h / Math.Max(1, ch)) is { Count: > 0 } vp)
                    {
                        paths = vp;
                        preset = null;
                    }
                    preset ??= paths is null ? "rect" : null;
                }
                break;
        }
        var textbox = e.Element(V + "textbox");
        var txbx = textbox?.Elements().FirstOrDefault(c => c.Name.LocalName == "txbxContent");
        var inset = (textbox?.Attribute("inset")?.Value ?? "").Split(',');
        double Inset(int i, double def) => i < inset.Length && inset[i].Trim().Length > 0 ? Len(inset[i].Trim()) : def;
        var tbStyle = Css(textbox?.Attribute("style")?.Value);
        var textpath = e.Element(V + "textpath");
        (string, string)? wordArt = textpath is null ? null
            : (textpath.Attribute("string")?.Value ?? "", Regex.Match(textpath.Attribute("style")?.Value ?? "", "font-family:\\s*\"?([^\";]+)").Groups[1].Value is { Length: > 0 } f ? f : "Calibri");
        return new ShapeVisual
        {
            X = x, Y = y, W = w, H = h, Rotation = rot, FlipH = flip.Contains('x'), FlipV = flip.Contains('y'),
            Preset = preset, Adjust = adjust, Paths = paths,
            Fill = isLine ? null : VmlFill(e, typeEl),
            Line = Stroke(e, typeEl, defaultOn: true),
            TextBox = ctx.BoxIndex(txbx),
            Insets = (Inset(0, 7.2), Inset(1, 3.6), Inset(2, 7.2), Inset(3, 3.6)),
            TextAnchor = style.GetValueOrDefault("v-text-anchor") switch { "middle" or "middle-center" => "ctr", "bottom" or "bottom-center" => "b", _ => "t" },
            TextDirection = style.GetValueOrDefault("layout-flow") is "vertical" ? "vert" : null,
            AutoFit = tbStyle.GetValueOrDefault("mso-fit-shape-to-text") is "t" or "true",
            WordArt = wordArt,
        };
    }

    static FillSpec? VmlFill(XElement e, XElement? type)
    {
        var fill = e.Element(V + "fill") ?? type?.Element(V + "fill");
        var filled = e.Attribute("filled")?.Value ?? type?.Attribute("filled")?.Value;
        if (filled is "f" or "false" || fill?.Attribute("on")?.Value is "f" or "false") return new FillSpec("none");
        var color = VmlColor(fill?.Attribute("color")?.Value ?? e.Attribute("fillcolor")?.Value ?? type?.Attribute("fillcolor")?.Value) ?? "FFFFFF";
        double alpha = Opacity(fill?.Attribute("opacity")?.Value);
        switch (fill?.Attribute("type")?.Value)
        {
            case "gradient" or "gradientRadial":
                var c2 = VmlColor(fill.Attribute("color2")?.Value) ?? "FFFFFF";
                double a2 = Opacity(fill.Attribute(O + "opacity2")?.Value);
                return new FillSpec("gradient", color, alpha)
                {
                    Stops = new() { (0, color, alpha), (1, c2, a2) },
                    Angle = (90 - D(fill.Attribute("angle")?.Value) + 360) % 360,
                    Radial = fill.Attribute("type")?.Value == "gradientRadial",
                };
            case "frame" or "tile":
                return new FillSpec("picture") { ImageId = fill.Attribute(R + "id")?.Value ?? fill.Attribute(O + "relid")?.Value };
        }
        return new FillSpec("solid", color, alpha);
    }

    static LineSpec? Stroke(XElement e, XElement? type, bool defaultOn)
    {
        var st = e.Element(V + "stroke") ?? type?.Element(V + "stroke");
        var stroked = e.Attribute("stroked")?.Value ?? type?.Attribute("stroked")?.Value;
        bool on = stroked is null ? defaultOn : stroked is not ("f" or "false");
        if (st?.Attribute("on")?.Value is "f" or "false") on = false;
        else if (st?.Attribute("on")?.Value is "t" or "true") on = true;
        if (!on) return null;
        var color = VmlColor(st?.Attribute("color")?.Value ?? e.Attribute("strokecolor")?.Value ?? type?.Attribute("strokecolor")?.Value) ?? "000000";
        double weight = Len(st?.Attribute("weight")?.Value ?? e.Attribute("strokeweight")?.Value ?? type?.Attribute("strokeweight")?.Value, 0.75);
        var dash = st?.Attribute("dashstyle")?.Value switch
        {
            null or "solid" => "solid", "dash" => "dash", "shortdash" => "sysDash", "dot" => "dot", "shortdot" => "sysDot",
            "dashdot" => "dashDot", "shortdashdot" => "sysDashDot",
            "longdash" => "lgDash", "longdashdot" => "lgDashDot", "longdashdotdot" => "lgDashDotDot", var other => other,
        };
        string? Arrow(string? v) => v is null or "none" ? null : v switch { "block" => "triangle", "classic" => "stealth", "open" => "arrow", var o => o };
        return new LineSpec(color, weight, dash, Opacity(st?.Attribute("opacity")?.Value))
        {
            HeadEnd = Arrow(st?.Attribute("startarrow")?.Value), TailEnd = Arrow(st?.Attribute("endarrow")?.Value),
            Compound = st?.Attribute("linestyle")?.Value switch { "thinThin" => "dbl", "thinThick" => "thinThick", "thickThin" => "thickThin", "thickBetweenThin" => "tri", _ => null },
        };
    }

    static double Opacity(string? v)
    {
        if (v is null) return 1;
        if (v.EndsWith('f')) return Math.Clamp(D(v[..^1]) / 65536, 0, 1);
        if (v.StartsWith('.')) v = "0" + v;
        return Math.Clamp(D(v, 1), 0, 1);
    }

    static double Crop(string? v) => v is null ? 0 : v.EndsWith('f') ? D(v[..^1]) / 65536 : D(v);

    /// <summary>Couleur VML : #rrggbb, #rgb, nom, « black [3213] », window / windowText.</summary>
    internal static string? VmlColor(string? v)
    {
        if (v is null) return null;
        v = Regex.Replace(v, @"\s*\[.*\]", "").Trim();
        if (v.StartsWith('#'))
        {
            var hex = v[1..];
            if (hex.Length == 3) hex = string.Concat(hex.Select(c => $"{c}{c}"));
            return hex.Length == 6 ? hex.ToUpperInvariant() : null;
        }
        if (v.Equals("windowText", StringComparison.OrdinalIgnoreCase)) return "000000";
        if (v.Equals("window", StringComparison.OrdinalIgnoreCase)) return "FFFFFF";
        return Presets.TryGetValue(v, out var c) ? c : null;
    }

    /// <summary>Style CSS d'un élément VML (propriétés en minuscules).</summary>
    internal static Dictionary<string, string> Css(string? style)
    {
        var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var decl in (style ?? "").Split(';'))
        {
            int k = decl.IndexOf(':');
            if (k <= 0) continue;
            d[decl[..k].Trim().ToLowerInvariant()] = decl[(k + 1)..].Trim();
        }
        return d;
    }

    static string? Num(string? v) => v is null ? null : Regex.Match(v, @"-?[\d.]+").Value;

    /// <summary>Longueur CSS en points (pt, in, cm, mm, pc, px, emu ; nombre seul = pixels).</summary>
    internal static double Len(string? v, double def = 0)
    {
        if (string.IsNullOrWhiteSpace(v)) return def;
        var m = Regex.Match(v.Trim().ToLowerInvariant(), @"^(-?[\d.]+)\s*(pt|px|in|cm|mm|pc|emu)?");
        if (!m.Success || !double.TryParse(m.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var n)) return def;
        return m.Groups[2].Value switch
        {
            "pt" => n, "in" => n * 72, "cm" => n * 72 / 2.54, "mm" => n * 72 / 25.4, "pc" => n * 12, "emu" => n / Emu, _ => n * 0.75,
        };
    }

    static (double, double) Point(string? v)
    {
        var p = (v ?? "0,0").Split(',');
        return (Len(p[0]), p.Length > 1 ? Len(p[1]) : 0);
    }

    static (double, double) PointRaw(string? v)
    {
        var p = (v ?? "0,0").Split(',');
        return (D(Num(p[0])), p.Length > 1 ? D(Num(p[1])) : 0);
    }

    /// <summary>
    /// Chemin VML (m, l, c, x, e, r, v, t, nf, ns, qx, qy) sans formules, mis à l'échelle ; null si une commande n'est
    /// pas comprise (la forme reste un rectangle).
    /// </summary>
    static List<ShapePath>? VmlPath(string path, double sx, double sy)
    {
        var tokens = Regex.Matches(path, @"[a-zA-Z]+|-?\d+(\.\d+)?").Select(m => m.Value).ToList();
        var result = new List<ShapePath>();
        var segs = new List<PathSegment>();
        bool fill = true, stroke = true;
        double cx = 0, cy = 0;
        int i = 0;
        string cmd = "";
        double Next() => i < tokens.Count && !char.IsLetter(tokens[i][0]) ? D(tokens[i++]) : double.NaN;
        bool HasNum() => i < tokens.Count && !char.IsLetter(tokens[i][0]);
        void Flush()
        {
            if (segs.Count > 0) result.Add(new ShapePath(segs, fill ? "norm" : "none", stroke));
            segs = new();
        }
        while (i < tokens.Count)
        {
            if (char.IsLetter(tokens[i][0])) cmd = tokens[i++];
            switch (cmd)
            {
                case "m": { var x = Next(); var y = Next(); if (double.IsNaN(y)) return null; (cx, cy) = (x * sx, y * sy); segs.Add(new('M', new[] { cx, cy })); break; }
                case "t": { var x = Next(); var y = Next(); if (double.IsNaN(y)) return null; (cx, cy) = (cx + x * sx, cy + y * sy); segs.Add(new('M', new[] { cx, cy })); break; }
                case "l": { var x = Next(); var y = Next(); if (double.IsNaN(y)) return null; (cx, cy) = (x * sx, y * sy); segs.Add(new('L', new[] { cx, cy })); break; }
                case "r": { var x = Next(); var y = Next(); if (double.IsNaN(y)) return null; (cx, cy) = (cx + x * sx, cy + y * sy); segs.Add(new('L', new[] { cx, cy })); break; }
                case "c":
                {
                    var p = Enumerable.Range(0, 6).Select(_ => Next()).ToArray();
                    if (p.Any(double.IsNaN)) return null;
                    segs.Add(new('C', new[] { p[0] * sx, p[1] * sy, p[2] * sx, p[3] * sy, p[4] * sx, p[5] * sy }));
                    (cx, cy) = (p[4] * sx, p[5] * sy);
                    break;
                }
                case "v":
                {
                    var p = Enumerable.Range(0, 6).Select(_ => Next()).ToArray();
                    if (p.Any(double.IsNaN)) return null;
                    segs.Add(new('C', new[] { cx + p[0] * sx, cy + p[1] * sy, cx + p[2] * sx, cy + p[3] * sy, cx + p[4] * sx, cy + p[5] * sy }));
                    (cx, cy) = (cx + p[4] * sx, cy + p[5] * sy);
                    break;
                }
                case "x": segs.Add(new('Z', Array.Empty<double>())); break;
                case "e": Flush(); fill = stroke = true; break;
                case "nf": fill = false; break;
                case "ns": stroke = false; break;
                default: return null;
            }
            // commandes répétées : « l 10,0 20,0 » — les nombres suivants reprennent la commande
            if (cmd is "x" or "e" or "nf" or "ns" && HasNum()) return null;
        }
        Flush();
        return result;
    }
}
