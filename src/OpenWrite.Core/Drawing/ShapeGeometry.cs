using System.Globalization;
using System.Reflection;
using System.Xml.Linq;

namespace OpenWrite.Core.Drawing;

/// <summary>Segment d'un contour : M (aller à), L (ligne), C (Bézier cubique : 3 points), Q (quadratique : 2 points), Z (fermer).</summary>
public readonly record struct PathSegment(char Op, double[] Points);

/// <summary>
/// Chemin d'une forme, en points dans le cadre de la forme (0,0 en haut à gauche). <see cref="Fill"/> : mode de
/// remplissage DrawingML (norm, none, lighten, lightenLess, darken, darkenLess) ; <see cref="Stroke"/> : tracé du contour.
/// </summary>
public sealed record ShapePath(List<PathSegment> Segments, string Fill, bool Stroke);

/// <summary>
/// Géométrie des formes DrawingML : formes prédéfinies (a:prstGeom, définitions officielles de la norme,
/// presetShapeDefinitions.xml) et géométries libres (a:custGeom), évaluées pour une taille donnée avec le langage des
/// guides (*/, +-, at2, pin…). Les arcs (arcTo) sont convertis en Béziers cubiques. Rectangle du texte (a:rect)
/// compris.
/// </summary>
public static class ShapeGeometry
{
    public static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    static readonly Lazy<Dictionary<string, XElement>> Presets = new(() =>
    {
        var asm = typeof(ShapeGeometry).Assembly;
        var name = asm.GetManifestResourceNames().First(n => n.EndsWith("presetShapeDefinitions.xml", StringComparison.Ordinal));
        using var s = asm.GetManifestResourceStream(name)!;
        var root = XDocument.Load(s).Root!;
        return root.Elements().ToDictionary(e => e.Name.LocalName, StringComparer.Ordinal);
    });

    /// <summary>Noms des formes prédéfinies connues.</summary>
    public static IEnumerable<string> PresetNames => Presets.Value.Keys;

    public static bool IsPreset(string name) => Presets.Value.ContainsKey(name);

    /// <summary>Définition (avLst, gdLst, rect, pathLst) d'une forme prédéfinie ; null si inconnue.</summary>
    public static XElement? Preset(string name) => Presets.Value.TryGetValue(name, out var e) ? e : null;

    /// <summary>
    /// Chemins d'une forme prédéfinie de taille <paramref name="w"/> × <paramref name="h"/> points ;
    /// <paramref name="adjust"/> : réglages du fichier (a:avLst/a:gd, « val n »). Forme inconnue : rectangle.
    /// </summary>
    public static List<ShapePath> PresetPaths(string name, double w, double h, IReadOnlyDictionary<string, string>? adjust = null)
    {
        var def = Preset(name) ?? Preset("rect")!;
        return Evaluate(def, w, h, adjust).Paths;
    }

    /// <summary>Résultat de l'évaluation : chemins et rectangle du texte (gauche, haut, droite, bas ; points).</summary>
    public sealed record Result(List<ShapePath> Paths, (double L, double T, double R, double B) TextRect);

    /// <summary>
    /// Évalue une géométrie (définition prédéfinie ou a:custGeom : avLst, gdLst, rect, pathLst) pour une taille en points.
    /// </summary>
    public static Result Evaluate(XElement geom, double w, double h, IReadOnlyDictionary<string, string>? adjust = null)
    {
        // les formules travaillent en EMU comme Word (arrondis des « */ » sans effet sensible en points)
        var v = new Dictionary<string, double>(StringComparer.Ordinal);
        Builtins(v, w, h);
        foreach (var gd in geom.Element(A + "avLst")?.Elements(A + "gd") ?? Enumerable.Empty<XElement>())
        {
            var name = gd.Attribute("name")?.Value;
            if (name is null) continue;
            var f = adjust is not null && adjust.TryGetValue(name, out var a) ? a : gd.Attribute("fmla")?.Value ?? "val 0";
            v[name] = Formula(f, v);
        }
        if (adjust is not null)
            foreach (var (k, f) in adjust)
                if (!v.ContainsKey(k)) v[k] = Formula(f, v);
        foreach (var gd in geom.Element(A + "gdLst")?.Elements(A + "gd") ?? Enumerable.Empty<XElement>())
            if (gd.Attribute("name")?.Value is { } name) v[name] = Formula(gd.Attribute("fmla")?.Value ?? "val 0", v);

        double Val(string? s) => s is null ? 0 : double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var n) ? n : v.GetValueOrDefault(s);

        var paths = new List<ShapePath>();
        foreach (var path in geom.Element(A + "pathLst")?.Elements(A + "path") ?? Enumerable.Empty<XElement>())
        {
            // chemin de taille propre (w, h) : ramené à la taille de la forme
            double pw = Val(path.Attribute("w")?.Value), ph = Val(path.Attribute("h")?.Value);
            double sx = pw > 0 ? w / pw : 1, sy = ph > 0 ? h / ph : 1;
            var segs = new List<PathSegment>();
            double cx = 0, cy = 0, startX = 0, startY = 0;
            (double, double) Pt(XElement pt) => (Val(pt.Attribute("x")?.Value) * sx, Val(pt.Attribute("y")?.Value) * sy);
            foreach (var cmd in path.Elements())
            {
                switch (cmd.Name.LocalName)
                {
                    case "moveTo":
                    {
                        var (x, y) = Pt(cmd.Element(A + "pt")!);
                        segs.Add(new PathSegment('M', new[] { x, y }));
                        (cx, cy, startX, startY) = (x, y, x, y);
                        break;
                    }
                    case "lnTo":
                    {
                        var (x, y) = Pt(cmd.Element(A + "pt")!);
                        segs.Add(new PathSegment('L', new[] { x, y }));
                        (cx, cy) = (x, y);
                        break;
                    }
                    case "cubicBezTo":
                    {
                        var pts = cmd.Elements(A + "pt").Select(Pt).ToList();
                        if (pts.Count < 3) break;
                        segs.Add(new PathSegment('C', new[] { pts[0].Item1, pts[0].Item2, pts[1].Item1, pts[1].Item2, pts[2].Item1, pts[2].Item2 }));
                        (cx, cy) = pts[2];
                        break;
                    }
                    case "quadBezTo":
                    {
                        var pts = cmd.Elements(A + "pt").Select(Pt).ToList();
                        if (pts.Count < 2) break;
                        segs.Add(new PathSegment('Q', new[] { pts[0].Item1, pts[0].Item2, pts[1].Item1, pts[1].Item2 }));
                        (cx, cy) = pts[1];
                        break;
                    }
                    case "arcTo":
                    {
                        double wr = Val(cmd.Attribute("wR")?.Value) * sx, hr = Val(cmd.Attribute("hR")?.Value) * sy;
                        double st = Val(cmd.Attribute("stAng")?.Value) / 60000.0, sw = Val(cmd.Attribute("swAng")?.Value) / 60000.0;
                        if (segs.Count == 0) segs.Add(new PathSegment('M', new[] { cx, cy }));
                        (cx, cy) = Arc(segs, cx, cy, wr, hr, st, sw);
                        break;
                    }
                    case "close":
                        segs.Add(new PathSegment('Z', Array.Empty<double>()));
                        (cx, cy) = (startX, startY);
                        break;
                }
            }
            var fill = path.Attribute("fill")?.Value ?? "norm";
            bool stroke = path.Attribute("stroke")?.Value is not ("0" or "false");
            paths.Add(new ShapePath(segs, fill, stroke));
        }
        var rect = geom.Element(A + "rect");
        var text = rect is null ? (0.0, 0.0, w, h)
            : (Val(rect.Attribute("l")?.Value), Val(rect.Attribute("t")?.Value), Val(rect.Attribute("r")?.Value), Val(rect.Attribute("b")?.Value));
        return new Result(paths, text);
    }

    /// <summary>
    /// Arc d'ellipse DrawingML depuis le point courant : rayons, angle de départ et balayage en degrés (angles « visuels »
    /// convertis en paramètres de l'ellipse), découpé en Béziers d'au plus 90°.
    /// </summary>
    static (double, double) Arc(List<PathSegment> segs, double x, double y, double wr, double hr, double stDeg, double swDeg)
    {
        if (wr <= 0 || hr <= 0 || swDeg == 0) return (x, y);
        static double Param(double deg, double wr, double hr)
        {
            double a = deg * Math.PI / 180;
            return Math.Atan2(wr * Math.Sin(a), hr * Math.Cos(a));
        }
        double t0 = Param(stDeg, wr, hr);
        double t1 = Param(stDeg + swDeg, wr, hr);
        // garder le sens et l'amplitude du balayage
        double sweep = t1 - t0;
        if (swDeg > 0) { while (sweep <= 0) sweep += 2 * Math.PI; if (swDeg >= 360) sweep = 2 * Math.PI; while (sweep > 2 * Math.PI + 1e-9) sweep -= 2 * Math.PI; }
        else { while (sweep >= 0) sweep -= 2 * Math.PI; if (swDeg <= -360) sweep = -2 * Math.PI; while (sweep < -2 * Math.PI - 1e-9) sweep += 2 * Math.PI; }
        double cx = x - wr * Math.Cos(t0), cy = y - hr * Math.Sin(t0);
        int n = Math.Max(1, (int)Math.Ceiling(Math.Abs(sweep) / (Math.PI / 2) - 1e-9));
        double step = sweep / n;
        double k = 4.0 / 3 * Math.Tan(step / 4);
        double t = t0;
        double px = x, py = y;
        for (int i = 0; i < n; i++)
        {
            double ta = t, tb = t + step;
            double ax = cx + wr * Math.Cos(ta), ay = cy + hr * Math.Sin(ta);
            double bx = cx + wr * Math.Cos(tb), by = cy + hr * Math.Sin(tb);
            double c1x = ax - k * wr * Math.Sin(ta), c1y = ay + k * hr * Math.Cos(ta);
            double c2x = bx + k * wr * Math.Sin(tb), c2y = by - k * hr * Math.Cos(tb);
            segs.Add(new PathSegment('C', new[] { c1x, c1y, c2x, c2y, bx, by }));
            (px, py) = (bx, by);
            t = tb;
        }
        return (px, py);
    }

    static void Builtins(Dictionary<string, double> v, double w, double h)
    {
        v["w"] = w; v["h"] = h; v["l"] = 0; v["t"] = 0; v["r"] = w; v["b"] = h;
        v["hc"] = w / 2; v["vc"] = h / 2;
        double ss = Math.Min(w, h), ls = Math.Max(w, h);
        v["ss"] = ss; v["ls"] = ls;
        foreach (var d in new[] { 2, 3, 4, 5, 6, 8, 10, 12, 32 })
        {
            v["wd" + d] = w / d;
            v["hd" + d] = h / d;
            v["ssd" + d] = ss / d;
        }
        v["ssd16"] = ss / 16;
        v["cd2"] = 10800000; v["cd4"] = 5400000; v["cd8"] = 2700000;
        v["3cd4"] = 16200000; v["3cd8"] = 8100000; v["5cd8"] = 13500000; v["7cd8"] = 18900000;
    }

    /// <summary>Formule de guide DrawingML ; angles en 60 000es de degré.</summary>
    public static double Formula(string fmla, IReadOnlyDictionary<string, double> v)
    {
        var parts = fmla.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0) return 0;
        double A(int i) => i < parts.Length
            ? double.TryParse(parts[i], NumberStyles.Float, CultureInfo.InvariantCulture, out var n) ? n : v.TryGetValue(parts[i], out var x) ? x : 0
            : 0;
        const double Deg = Math.PI / 180 / 60000;
        return parts[0] switch
        {
            "val" => A(1),
            "*/" => A(3) == 0 ? 0 : A(1) * A(2) / A(3),
            "+-" => A(1) + A(2) - A(3),
            "+/" => A(3) == 0 ? 0 : (A(1) + A(2)) / A(3),
            "?:" => A(1) > 0 ? A(2) : A(3),
            "abs" => Math.Abs(A(1)),
            "at2" => Math.Atan2(A(2), A(1)) / Deg,
            "cat2" => A(1) * Math.Cos(Math.Atan2(A(3), A(2))),
            "sat2" => A(1) * Math.Sin(Math.Atan2(A(3), A(2))),
            "cos" => A(1) * Math.Cos(A(2) * Deg),
            "sin" => A(1) * Math.Sin(A(2) * Deg),
            "tan" => A(1) * Math.Tan(A(2) * Deg),
            "max" => Math.Max(A(1), A(2)),
            "min" => Math.Min(A(1), A(2)),
            "mod" => Math.Sqrt(A(1) * A(1) + A(2) * A(2) + A(3) * A(3)),
            "pin" => A(2) < A(1) ? A(1) : A(2) > A(3) ? A(3) : A(2),
            "sqrt" => Math.Sqrt(Math.Max(0, A(1))),
            _ => 0,
        };
    }
}
