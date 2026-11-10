using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;

namespace OpenWrite.Core.Drawing;

/// <summary>
/// Lecture d'un dessin (<see cref="DrawingContent"/>) en <see cref="DrawingObject"/> : DrawingML (wp:inline, wp:anchor ;
/// images, formes wps, groupes wpg, zones de dessin wpc ; graphiques et SmartArt en objets de repli) et VML ancien
/// (w:pict : v:shape, v:rect, v:roundrect, v:oval, v:line, v:polyline, v:image, v:group ; w10:wrap ; v:textbox).
/// mc:AlternateContent : la variante DrawingML (Choice) si elle est comprise, sinon le VML (Fallback). Couleurs du
/// thème et styles de forme (wps:style) résolus avec la feuille de styles.
/// </summary>
public static partial class DrawingReader
{
    public static readonly XNamespace A = ShapeGeometry.A;
    public static readonly XNamespace Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    public static readonly XNamespace Wp14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
    public static readonly XNamespace Pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    public static readonly XNamespace Wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
    public static readonly XNamespace Wpg = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
    public static readonly XNamespace Wpc = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";
    public static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    public static readonly XNamespace V = "urn:schemas-microsoft-com:vml";
    public static readonly XNamespace O = "urn:schemas-microsoft-com:office:office";
    public static readonly XNamespace W10 = "urn:schemas-microsoft-com:office:word";
    const double Emu = 12700;

    /// <summary>Contenus de zone de texte (w:txbxContent) hors zones imbriquées, dans l'ordre (comme StoryMapping).</summary>
    public static List<XElement> TextBoxContents(XElement drawing)
    {
        var res = new List<XElement>();
        void Walk(XElement e)
        {
            foreach (var ch in e.Elements())
            {
                if (ch.Name.LocalName == "txbxContent") res.Add(ch);
                else Walk(ch);
            }
        }
        Walk(drawing);
        return res;
    }

    /// <summary>Lit le dessin ; null s'il n'a rien d'affichable (élément inconnu, forme vide).</summary>
    public static DrawingObject? Read(DrawingContent d, StyleSheet? sheet = null)
    {
        var ctx = new Ctx(d, sheet);
        var x = d.Xml;
        if (x.Name == Mc + "AlternateContent")
        {
            // Choice DrawingML d'abord (wps, wpg, wpc, wp14… : compris), sinon le VML de repli
            foreach (var choice in x.Elements(Mc + "Choice"))
                if (choice.Elements().FirstOrDefault(e => e.Name.LocalName == "drawing") is { } dr && ReadDrawingMl(dr, ctx) is { } o) return o;
            if (x.Element(Mc + "Fallback") is { } fb)
                foreach (var e in fb.Elements())
                    if (ReadElement(e, ctx) is { } o) return o;
            return null;
        }
        return ReadElement(x, ctx);
    }

    static DrawingObject? ReadElement(XElement x, Ctx ctx) => x.Name.LocalName switch
    {
        "drawing" => ReadDrawingMl(x, ctx),
        "pict" or "object" => ReadVml(x, ctx),
        _ => null,
    };

    sealed class Ctx(DrawingContent d, StyleSheet? sheet)
    {
        public DrawingContent Source { get; } = d;
        public StyleSheet? Sheet { get; } = sheet;
        List<XElement>? _boxes;
        /// <summary>Indice d'un w:txbxContent dans les zones de texte du dessin.</summary>
        public int BoxIndex(XElement? txbx)
        {
            if (txbx is null) return -1;
            _boxes ??= TextBoxContents(Source.Xml);
            return _boxes.IndexOf(txbx);
        }
    }

    static double D(string? s, double def = 0) =>
        s is not null && double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : def;

    static bool On(string? s, bool def = false) => s is null ? def : s is "1" or "true" or "t" or "on";

    // ================================================================ DrawingML

    static DrawingObject? ReadDrawingMl(XElement drawing, Ctx ctx)
    {
        var frame = drawing.Element(Wp + "inline") ?? drawing.Element(Wp + "anchor");
        if (frame is null) return null;
        bool inline = frame.Name.LocalName == "inline";
        var ext = frame.Element(Wp + "extent");
        double w = D(ext?.Attribute("cx")?.Value) / Emu, h = D(ext?.Attribute("cy")?.Value) / Emu;
        var eff = frame.Element(Wp + "effectExtent");
        var effect = (D(eff?.Attribute("l")?.Value) / Emu, D(eff?.Attribute("t")?.Value) / Emu, D(eff?.Attribute("r")?.Value) / Emu, D(eff?.Attribute("b")?.Value) / Emu);
        var docPr = frame.Element(Wp + "docPr");
        var gd = frame.Element(A + "graphic")?.Element(A + "graphicData");
        var root = gd is null ? null : ReadGraphic(gd, 0, 0, w, h, ctx);
        if (root is null) return null;

        if (inline)
            return new DrawingObject
            {
                Source = ctx.Source, Inline = true, Width = w, Height = h, Effect = effect, Root = root,
                Name = docPr?.Attribute("name")?.Value, Description = docPr?.Attribute("descr")?.Value, Hidden = On(docPr?.Attribute("hidden")?.Value),
                Distance = (D(frame.Attribute("distT")?.Value) / Emu, D(frame.Attribute("distB")?.Value) / Emu, D(frame.Attribute("distL")?.Value) / Emu, D(frame.Attribute("distR")?.Value) / Emu),
            };

        AnchorAxis Axis(XElement? pos, string def)
        {
            if (pos is null) return new AnchorAxis(def, null, 0);
            var from = pos.Attribute("relativeFrom")?.Value ?? def;
            if (pos.Element(Wp + "align") is { } al) return new AnchorAxis(from, al.Value.Trim(), 0);
            if (pos.Element(Wp14 + "pctPosHOffset") is { } ph) return new AnchorAxis(from, null, 0, D(ph.Value) / 100000);
            if (pos.Element(Wp14 + "pctPosVOffset") is { } pv) return new AnchorAxis(from, null, 0, D(pv.Value) / 100000);
            return new AnchorAxis(from, null, D(pos.Element(Wp + "posOffset")?.Value) / Emu);
        }
        var simple = On(frame.Attribute("simplePos")?.Value);
        var hAxis = Axis(frame.Element(Wp + "positionH"), "column");
        var vAxis = Axis(frame.Element(Wp + "positionV"), "paragraph");
        if (simple && frame.Element(Wp + "simplePos") is { } sp)
        {
            hAxis = new AnchorAxis("page", null, D(sp.Attribute("x")?.Value) / Emu);
            vAxis = new AnchorAxis("page", null, D(sp.Attribute("y")?.Value) / Emu);
        }
        var (wrap, side, poly) = ReadWrap(frame);
        (double, string)? Rel(XName name, XName pct)
        {
            var e = frame.Element(name);
            var p = e?.Element(pct);
            return p is null || D(p.Value) <= 0 ? null : (D(p.Value) / 100000, e!.Attribute("relativeFrom")?.Value ?? "margin");
        }
        return new DrawingObject
        {
            Source = ctx.Source, Inline = false, Width = w, Height = h, Effect = effect, Root = root,
            H = hAxis, V = vAxis, Wrap = wrap, WrapSide = side, WrapPolygon = poly,
            Distance = (D(frame.Attribute("distT")?.Value) / Emu, D(frame.Attribute("distB")?.Value) / Emu, D(frame.Attribute("distL")?.Value) / Emu, D(frame.Attribute("distR")?.Value) / Emu),
            BehindText = On(frame.Attribute("behindDoc")?.Value),
            Z = (long)D(frame.Attribute("relativeHeight")?.Value),
            AllowOverlap = On(frame.Attribute("allowOverlap")?.Value, true),
            LayoutInCell = On(frame.Attribute("layoutInCell")?.Value, true),
            Locked = On(frame.Attribute("locked")?.Value),
            RelWidth = Rel(Wp14 + "sizeRelH", Wp14 + "pctWidth"),
            RelHeight = Rel(Wp14 + "sizeRelV", Wp14 + "pctHeight"),
            Name = docPr?.Attribute("name")?.Value, Description = docPr?.Attribute("descr")?.Value, Hidden = On(docPr?.Attribute("hidden")?.Value),
        };
    }

    static (WrapKind, string, List<(double, double)>?) ReadWrap(XElement anchor)
    {
        foreach (var e in anchor.Elements())
        {
            if (e.Name.Namespace != Wp) continue;
            var side = e.Attribute("wrapText")?.Value ?? "bothSides";
            switch (e.Name.LocalName)
            {
                case "wrapNone": return (WrapKind.None, side, null);
                case "wrapSquare": return (WrapKind.Square, side, null);
                case "wrapTopAndBottom": return (WrapKind.TopAndBottom, side, null);
                case "wrapTight":
                case "wrapThrough":
                {
                    // polygone en 21 600es du cadre
                    var poly = e.Element(Wp + "wrapPolygon")?.Elements().Where(p => p.Name.LocalName is "start" or "lineTo")
                        .Select(p => (D(p.Attribute("x")?.Value) / 21600, D(p.Attribute("y")?.Value) / 21600)).ToList();
                    return (e.Name.LocalName == "wrapTight" ? WrapKind.Tight : WrapKind.Through, side, poly is { Count: >= 3 } ? poly : null);
                }
            }
        }
        return (WrapKind.None, "bothSides", null);
    }

    /// <summary>Contenu d'un a:graphicData placé dans le cadre (x, y, w, h).</summary>
    static Visual? ReadGraphic(XElement gd, double x, double y, double w, double h, Ctx ctx)
    {
        var uri = gd.Attribute("uri")?.Value ?? "";
        var body = gd.Elements().FirstOrDefault();
        if (body is null) return null;
        if (body.Name == Pic + "pic") return ReadPicture(body, x, y, w, h, ctx, useXfrm: false);
        if (body.Name == Wps + "wsp") return ReadShape(body, x, y, w, h, ctx, useXfrm: false);
        if (body.Name == Wpg + "wgp") return ReadGroup(body, x, y, w, h, ctx, isRoot: true);
        if (body.Name == Wpc + "wpc") return ReadCanvas(body, x, y, w, h, ctx);
        if (uri.EndsWith("/chart", StringComparison.Ordinal) || body.Name.LocalName == "chart")
            return new FallbackVisual { Kind = "chart", X = x, Y = y, W = w, H = h };
        if (uri.EndsWith("/diagram", StringComparison.Ordinal) || body.Name.LocalName == "relIds")
        {
            // SmartArt : la forme calculée est dans la partie dessin désignée par l'extension dsp (dataModelExt/relId)
            return new FallbackVisual { Kind = "smartart", X = x, Y = y, W = w, H = h, DrawingId = body.Attribute(R + "dm")?.Value };
        }
        return new FallbackVisual { Kind = "unknown", X = x, Y = y, W = w, H = h };
    }

    /// <summary>a:xfrm : position, taille, rotation, retournements ; null sans xfrm.</summary>
    static (double X, double Y, double W, double H, double Rot, bool FlipH, bool FlipV)? Xfrm(XElement? xfrm)
    {
        if (xfrm is null) return null;
        var off = xfrm.Element(A + "off");
        var ext = xfrm.Element(A + "ext");
        return (D(off?.Attribute("x")?.Value) / Emu, D(off?.Attribute("y")?.Value) / Emu, D(ext?.Attribute("cx")?.Value) / Emu, D(ext?.Attribute("cy")?.Value) / Emu,
                D(xfrm.Attribute("rot")?.Value) / 60000, On(xfrm.Attribute("flipH")?.Value), On(xfrm.Attribute("flipV")?.Value));
    }

    static void Place(Visual v, double x, double y, double w, double h, XElement? xfrm, bool useXfrm)
    {
        var xf = Xfrm(xfrm);
        if (useXfrm && xf is { } f) { (v.X, v.Y, v.W, v.H) = (f.X, f.Y, f.W, f.H); }
        else { (v.X, v.Y, v.W, v.H) = (x, y, w, h); }
        if (xf is { } r) { v.Rotation = r.Rot; v.FlipH = r.FlipH; v.FlipV = r.FlipV; }
    }

    static Visual ReadPicture(XElement pic, double x, double y, double w, double h, Ctx ctx, bool useXfrm)
    {
        var fill = pic.Element(Pic + "blipFill");
        var blip = fill?.Element(A + "blip");
        var src = fill?.Element(A + "srcRect");
        var spPr = pic.Element(Pic + "spPr");
        double brightness = 0, contrast = 0, alpha = 1;
        bool gray = false;
        if (blip is not null)
        {
            if (blip.Element(A + "lum") is { } lum) { brightness = D(lum.Attribute("bright")?.Value) / 100000; contrast = D(lum.Attribute("contrast")?.Value) / 100000; }
            if (blip.Descendants().FirstOrDefault(e => e.Name.LocalName == "brightnessContrast") is { } bc)
            { brightness = D(bc.Attribute("bright")?.Value) / 100000; contrast = D(bc.Attribute("contrast")?.Value) / 100000; }
            gray = blip.Element(A + "grayscl") is not null;
            if (blip.Element(A + "alphaModFix") is { } am) alpha = D(am.Attribute("amt")?.Value, 100000) / 100000;
        }
        var v = new PictureVisual
        {
            ImageId = blip?.Attribute(R + "embed")?.Value,
            LinkId = blip?.Attribute(R + "link")?.Value,
            Crop = (D(src?.Attribute("l")?.Value) / 100000, D(src?.Attribute("t")?.Value) / 100000, D(src?.Attribute("r")?.Value) / 100000, D(src?.Attribute("b")?.Value) / 100000),
            Geometry = spPr?.Element(A + "prstGeom")?.Attribute("prst")?.Value ?? "rect",
            Line = ReadLine(spPr?.Element(A + "ln"), null, ctx),
            Brightness = brightness, Contrast = contrast, Grayscale = gray, Alpha = alpha,
        };
        Place(v, x, y, w, h, spPr?.Element(A + "xfrm"), useXfrm);
        return v;
    }

    static Visual ReadShape(XElement wsp, double x, double y, double w, double h, Ctx ctx, bool useXfrm)
    {
        var spPr = wsp.Element(Wps + "spPr");
        var style = wsp.Element(Wps + "style");
        var body = wsp.Element(Wps + "bodyPr");
        var prst = spPr?.Element(A + "prstGeom");
        var cust = spPr?.Element(A + "custGeom");
        Dictionary<string, string>? adjust = null;
        if (prst?.Element(A + "avLst") is { HasElements: true } av)
            adjust = av.Elements(A + "gd").Where(g => g.Attribute("name") is not null).ToDictionary(g => g.Attribute("name")!.Value, g => g.Attribute("fmla")?.Value ?? "val 0");
        // zone de texte : wps:txbx/w:txbxContent
        var txbx = wsp.Element(Wps + "txbx")?.Elements().FirstOrDefault(e => e.Name.LocalName == "txbxContent");
        var fill = ReadFill(spPr, style?.Element(A + "fillRef"), ctx);
        var line = ReadLine(spPr?.Element(A + "ln"), style?.Element(A + "lnRef"), ctx);
        var v = new ShapeVisual
        {
            Preset = cust is not null ? null : prst?.Attribute("prst")?.Value ?? "rect",
            Adjust = adjust,
            Custom = cust,
            Fill = fill,
            Line = line,
            TextBox = ctx.BoxIndex(txbx),
            Insets = (D(body?.Attribute("lIns")?.Value, 91440) / Emu, D(body?.Attribute("tIns")?.Value, 45720) / Emu,
                      D(body?.Attribute("rIns")?.Value, 91440) / Emu, D(body?.Attribute("bIns")?.Value, 45720) / Emu),
            TextAnchor = body?.Attribute("anchor")?.Value ?? "t",
            TextDirection = body?.Attribute("vert")?.Value is { } vert && vert != "horz" ? vert : null,
            AutoFit = body?.Element(A + "spAutoFit") is not null,
        };
        Place(v, x, y, w, h, spPr?.Element(A + "xfrm"), useXfrm);
        return v;
    }

    /// <summary>Groupe wpg : enfants ramenés du repère enfant (chOff, chExt) au cadre du groupe.</summary>
    static Visual ReadGroup(XElement grp, double x, double y, double w, double h, Ctx ctx, bool isRoot)
    {
        var grpSpPr = grp.Element(Wpg + "grpSpPr");
        var xfrm = grpSpPr?.Element(A + "xfrm");
        var g = new GroupVisual();
        Place(g, x, y, w, h, xfrm, useXfrm: !isRoot);
        var chOff = xfrm?.Element(A + "chOff");
        var chExt = xfrm?.Element(A + "chExt");
        double cx = D(chOff?.Attribute("x")?.Value) / Emu, cy = D(chOff?.Attribute("y")?.Value) / Emu;
        double cw = D(chExt?.Attribute("cx")?.Value) / Emu, ch = D(chExt?.Attribute("cy")?.Value) / Emu;
        double sx = cw > 0 ? g.W / cw : 1, sy = ch > 0 ? g.H / ch : 1;
        foreach (var child in grp.Elements())
        {
            Visual? v = child.Name.LocalName switch
            {
                "wsp" => ReadShape(child, 0, 0, 0, 0, ctx, useXfrm: true),
                "pic" => ReadPicture(child, 0, 0, 0, 0, ctx, useXfrm: true),
                "grpSp" => ReadGroup(child, 0, 0, 0, 0, ctx, isRoot: false),
                "graphicFrame" => child.Element(A + "graphic")?.Element(A + "graphicData") is { } gd2 && Xfrm(child.Element(Wpg + "xfrm")) is { } fx
                    ? ReadGraphic(gd2, fx.X, fx.Y, fx.W, fx.H, ctx) : null,
                _ => null,
            };
            if (v is null) continue;
            // repère enfant → repère du groupe (points, origine du groupe)
            v.X = (v.X - cx) * sx;
            v.Y = (v.Y - cy) * sy;
            v.W *= sx;
            v.H *= sy;
            g.Children.Add(v);
        }
        return g;
    }

    /// <summary>Zone de dessin wpc : fond, puis formes et images en coordonnées absolues du cadre.</summary>
    static Visual ReadCanvas(XElement wpc, double x, double y, double w, double h, Ctx ctx)
    {
        var g = new GroupVisual { Fill = ReadFill(wpc.Element(Wpc + "bg"), null, ctx), X = x, Y = y, W = w, H = h };
        foreach (var child in wpc.Elements())
        {
            Visual? v = child.Name.LocalName switch
            {
                "wsp" => ReadShape(child, 0, 0, 0, 0, ctx, useXfrm: true),
                "pic" => ReadPicture(child, 0, 0, 0, 0, ctx, useXfrm: true),
                "wgp" => ReadGroup(child, 0, 0, 0, 0, ctx, isRoot: false),
                _ => null,
            };
            if (v is not null) g.Children.Add(v);
        }
        return g;
    }

    // ---------------------------------------------------------------- couleurs, remplissages, traits

    /// <summary>Remplissage de spPr ; sinon celui du style (fillRef), sinon null (aucun).</summary>
    static FillSpec? ReadFill(XElement? spPr, XElement? fillRef, Ctx ctx)
    {
        if (spPr is not null)
            foreach (var e in spPr.Elements())
            {
                switch (e.Name.LocalName)
                {
                    case "noFill": return new FillSpec("none");
                    case "solidFill":
                        return Color(e, ctx, null) is { } c ? new FillSpec("solid", c.Rgb, c.Alpha) : null;
                    case "gradFill": return Gradient(e, ctx, null);
                    case "blipFill": return new FillSpec("picture") { ImageId = e.Element(A + "blip")?.Attribute(R + "embed")?.Value };
                    case "pattFill":
                        return Color(e.Element(A + "fgClr"), ctx, null) is { } pc ? new FillSpec("solid", pc.Rgb, pc.Alpha) : null;
                }
            }
        if (fillRef is null) return null;
        int idx = (int)D(fillRef.Attribute("idx")?.Value);
        var phClr = Color(fillRef, ctx, null);
        if (idx == 0) return new FillSpec("none");
        // style du thème : fillStyleLst (1..3) ou bgFillStyleLst (1001..)
        var fmt = ctx.Sheet?.Theme?.Element(A + "themeElements")?.Element(A + "fmtScheme");
        var list = idx >= 1000 ? fmt?.Element(A + "bgFillStyleLst") : fmt?.Element(A + "fillStyleLst");
        var item = list?.Elements().ElementAtOrDefault((idx >= 1000 ? idx - 1000 : idx) - 1);
        if (item is not null)
        {
            switch (item.Name.LocalName)
            {
                case "solidFill": return Color(item, ctx, phClr) is { } c ? new FillSpec("solid", c.Rgb, c.Alpha) : null;
                case "gradFill": return Gradient(item, ctx, phClr);
                case "noFill": return new FillSpec("none");
            }
        }
        return phClr is { } pcl ? new FillSpec("solid", pcl.Rgb, pcl.Alpha) : null;
    }

    static FillSpec? Gradient(XElement grad, Ctx ctx, (string Rgb, double Alpha)? ph)
    {
        var stops = grad.Element(A + "gsLst")?.Elements(A + "gs")
            .Select(gs => (Pos: D(gs.Attribute("pos")?.Value) / 100000, C: Color(gs, ctx, ph)))
            .Where(s => s.C is not null).Select(s => (s.Pos, s.C!.Value.Rgb, s.C.Value.Alpha)).OrderBy(s => s.Pos).ToList();
        if (stops is null || stops.Count == 0) return null;
        double angle = D(grad.Element(A + "lin")?.Attribute("ang")?.Value) / 60000;
        return new FillSpec("gradient", stops[0].Rgb, stops[0].Alpha) { Stops = stops, Angle = angle, Radial = grad.Element(A + "path") is not null };
    }

    /// <summary>Trait de a:ln ; sinon celui du style (lnRef) ; null : pas de trait.</summary>
    static LineSpec? ReadLine(XElement? ln, XElement? lnRef, Ctx ctx)
    {
        XElement? themeLn = null;
        (string Rgb, double Alpha)? ph = null;
        if (lnRef is not null)
        {
            int idx = (int)D(lnRef.Attribute("idx")?.Value);
            ph = Color(lnRef, ctx, null);
            if (idx > 0) themeLn = ctx.Sheet?.Theme?.Element(A + "themeElements")?.Element(A + "fmtScheme")?.Element(A + "lnStyleLst")?.Elements(A + "ln").ElementAtOrDefault(idx - 1);
            else if (ln is null) return null;
        }
        if (ln is null && themeLn is null) return null;
        if (ln?.Element(A + "noFill") is not null) return null;
        double width = D(ln?.Attribute("w")?.Value ?? themeLn?.Attribute("w")?.Value, 9525) / Emu;
        (string Rgb, double Alpha)? color = null;
        var solid = ln?.Element(A + "solidFill") ?? ln?.Element(A + "gradFill")?.Descendants(A + "gs").FirstOrDefault();
        if (solid is not null) color = Color(solid, ctx, ph);
        else if (themeLn?.Element(A + "solidFill") is { } ts) color = Color(ts, ctx, ph);
        else if (themeLn?.Element(A + "noFill") is not null && ln is null) return null;
        color ??= ph;
        if (color is null) return null;
        var dash = ln?.Element(A + "prstDash")?.Attribute("val")?.Value ?? themeLn?.Element(A + "prstDash")?.Attribute("val")?.Value ?? "solid";
        string? End(XElement? e) => e?.Attribute("type")?.Value is { } t && t != "none" ? t : null;
        return new LineSpec(color.Value.Rgb, width, dash, color.Value.Alpha)
        {
            HeadEnd = End(ln?.Element(A + "headEnd")), TailEnd = End(ln?.Element(A + "tailEnd")),
            Compound = ln?.Attribute("cmpd")?.Value is { } cmpd && cmpd != "sng" ? cmpd : null,
        };
    }

    static readonly Dictionary<string, string> Presets = new(StringComparer.OrdinalIgnoreCase)
    {
        ["black"] = "000000", ["white"] = "FFFFFF", ["red"] = "FF0000", ["green"] = "008000", ["blue"] = "0000FF",
        ["yellow"] = "FFFF00", ["gray"] = "808080", ["grey"] = "808080", ["silver"] = "C0C0C0", ["navy"] = "000080",
        ["maroon"] = "800000", ["purple"] = "800080", ["teal"] = "008080", ["olive"] = "808000", ["lime"] = "00FF00",
        ["aqua"] = "00FFFF", ["cyan"] = "00FFFF", ["fuchsia"] = "FF00FF", ["magenta"] = "FF00FF", ["orange"] = "FFA500",
        ["ltGray"] = "C0C0C0", ["dkGray"] = "808080",
    };

    /// <summary>
    /// Couleur DrawingML portée par <paramref name="holder"/> (srgbClr, schemeClr, sysClr, prstClr, scrgbClr) avec ses
    /// modificateurs (lumMod, lumOff, tint, shade, alpha) ; phClr = couleur du style.
    /// </summary>
    public static (string Rgb, double Alpha)? Color(XElement? holder, StyleSheet? sheet, (string Rgb, double Alpha)? ph = null) =>
        Color(holder, new Ctx(new DrawingContent(new XElement("x")), sheet), ph);

    static (string Rgb, double Alpha)? Color(XElement? holder, Ctx ctx, (string Rgb, double Alpha)? ph)
    {
        if (holder is null) return null;
        var c = holder.Elements().FirstOrDefault(e => e.Name.Namespace == A && e.Name.LocalName is "srgbClr" or "schemeClr" or "sysClr" or "prstClr" or "scrgbClr" or "hslClr");
        if (c is null) return null;
        string? rgb = c.Name.LocalName switch
        {
            "srgbClr" => c.Attribute("val")?.Value,
            "sysClr" => c.Attribute("lastClr")?.Value ?? (c.Attribute("val")?.Value == "window" ? "FFFFFF" : "000000"),
            "prstClr" => Presets.GetValueOrDefault(c.Attribute("val")?.Value ?? "", "000000"),
            "scrgbClr" => $"{Lin(c.Attribute("r")?.Value):X2}{Lin(c.Attribute("g")?.Value):X2}{Lin(c.Attribute("b")?.Value):X2}",
            "schemeClr" => Scheme(c.Attribute("val")?.Value, ctx, ph),
            _ => null,
        };
        if (rgb is null || rgb.Length != 6) return null;
        double alpha = c.Name.LocalName == "schemeClr" && c.Attribute("val")?.Value == "phClr" && ph is { } p ? p.Alpha : 1;
        foreach (var m in c.Elements())
        {
            double val = D(m.Attribute("val")?.Value) / 100000;
            switch (m.Name.LocalName)
            {
                case "lumMod": rgb = Hsl(rgb, l => l * val); break;
                case "lumOff": rgb = Hsl(rgb, l => l + val); break;
                case "tint": rgb = Hsl(rgb, l => l * val + (1 - val)); break;
                case "shade": rgb = Hsl(rgb, l => l * val); break;
                case "alpha": alpha = val; break;
            }
        }
        return (rgb.ToUpperInvariant(), alpha);

        static int Lin(string? pct) => (int)Math.Round(Math.Clamp(Math.Pow(D(pct) / 100000, 1 / 2.2), 0, 1) * 255);
    }

    static string? Scheme(string? val, Ctx ctx, (string Rgb, double Alpha)? ph)
    {
        if (val == "phClr") return ph?.Rgb;
        var name = val switch
        {
            "tx1" => "text1", "bg1" => "background1", "tx2" => "text2", "bg2" => "background2",
            "dk1" => "dark1", "lt1" => "light1", "dk2" => "dark2", "lt2" => "light2",
            "hlink" => "hyperlink", "folHlink" => "followedHyperlink",
            _ => val,
        };
        if (name is null) return null;
        return ctx.Sheet?.ThemeColor(name) ?? name switch
        {
            "text1" or "dark1" => "000000",
            "background1" or "light1" => "FFFFFF",
            "text2" or "dark2" => "44546A",
            "background2" or "light2" => "E7E6E6",
            "accent1" => "4472C4", "accent2" => "ED7D31", "accent3" => "A5A5A5", "accent4" => "FFC000", "accent5" => "5B9BD5", "accent6" => "70AD47",
            _ => null,
        };
    }

    /// <summary>Applique une fonction à la luminance (TSL) d'une couleur RRGGBB.</summary>
    internal static string Hsl(string rgb, Func<double, double> f)
    {
        double r = Convert.ToInt32(rgb[..2], 16) / 255.0, g = Convert.ToInt32(rgb[2..4], 16) / 255.0, b = Convert.ToInt32(rgb[4..], 16) / 255.0;
        double max = Math.Max(r, Math.Max(g, b)), min = Math.Min(r, Math.Min(g, b)), l = (max + min) / 2, hh = 0, s = 0;
        if (max != min)
        {
            double d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
            hh = max == r ? (g - b) / d + (g < b ? 6 : 0) : max == g ? (b - r) / d + 2 : (r - g) / d + 4;
            hh /= 6;
        }
        l = Math.Clamp(f(l), 0, 1);
        static double Hue(double p, double q, double t)
        {
            if (t < 0) t += 1;
            if (t > 1) t -= 1;
            return t < 1 / 6.0 ? p + (q - p) * 6 * t : t < 0.5 ? q : t < 2 / 3.0 ? p + (q - p) * (2 / 3.0 - t) * 6 : p;
        }
        if (s == 0) r = g = b = l;
        else
        {
            double q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
            r = Hue(p, q, hh + 1 / 3.0);
            g = Hue(p, q, hh);
            b = Hue(p, q, hh - 1 / 3.0);
        }
        return $"{(int)Math.Round(r * 255):X2}{(int)Math.Round(g * 255):X2}{(int)Math.Round(b * 255):X2}";
    }
}
