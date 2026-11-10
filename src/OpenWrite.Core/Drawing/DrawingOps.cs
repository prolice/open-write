using System.Globalization;
using System.Xml.Linq;
using OpenWrite.Core.Model;

namespace OpenWrite.Core.Drawing;

/// <summary>
/// Modifications des dessins (lot 7), écrites dans le XML comme Word les écrirait : taille, position, habillage, passage
/// en ligne ↔ flottant, rotation, retournement, ordre d'empilement, rognage, texte de remplacement, remplissage et
/// contour, corrections d'image ; création d'images, de formes et de zones de texte. Chaque modification remplace
/// <see cref="DrawingContent.Xml"/> par un nouvel élément (la mise en page met en cache par élément) ; elle se fait dans
/// une transaction du journal qui a photographié le paragraphe.
/// VML (documents anciens) : taille, position, habillage seulement.
/// </summary>
public static class DrawingOps
{
    static readonly XNamespace A = DrawingReader.A, Wp = DrawingReader.Wp, Wp14 = DrawingReader.Wp14, Pic = DrawingReader.Pic,
        Wps = DrawingReader.Wps, Wpg = DrawingReader.Wpg, R = DrawingReader.R, Mc = DrawingReader.Mc, V = DrawingReader.V, W10 = DrawingReader.W10;
    const double Emu = 12700;

    static string E(double pt) => ((long)Math.Round(pt * Emu)).ToString(CultureInfo.InvariantCulture);
    static string N(double v) => v.ToString("0.###", CultureInfo.InvariantCulture);

    /// <summary>Applique <paramref name="edit"/> à une copie du XML (et à la variante DrawingML / VML d'un mc:AlternateContent).</summary>
    static void Edit(DrawingContent d, Action<XElement> edit)
    {
        var x = new XElement(d.Xml);
        edit(x);
        d.Xml = x;
    }

    /// <summary>Cadre DrawingML (wp:inline / wp:anchor) de chaque variante du dessin.</summary>
    static IEnumerable<XElement> Frames(XElement x) => x.Descendants().Where(e => e.Name == Wp + "inline" || e.Name == Wp + "anchor");

    /// <summary>Forme VML principale de chaque variante (w:pict, mc:Fallback).</summary>
    static IEnumerable<XElement> VmlShapes(XElement x) => x.Descendants().Where(e => e.Name.Namespace == V
        && e.Name.LocalName is "shape" or "rect" or "roundrect" or "oval" or "line" or "group" or "image" or "polyline"
        && e.Parent?.Name.Namespace != V);

    // ================================================================ taille

    /// <summary>Taille du cadre (points) ; les formes, images et groupes suivent (a:xfrm/a:ext), le style VML aussi.</summary>
    public static void SetSize(DrawingContent d, double w, double h) => Edit(d, x =>
    {
        w = Math.Max(1, w);
        h = Math.Max(1, h);
        foreach (var f in Frames(x))
        {
            if (f.Element(Wp + "extent") is { } ext) { ext.SetAttributeValue("cx", E(w)); ext.SetAttributeValue("cy", E(h)); }
            // élément graphique racine : son xfrm suit le cadre
            var body = f.Element(A + "graphic")?.Element(A + "graphicData")?.Elements().FirstOrDefault();
            var xfrm = body?.Descendants(A + "xfrm").FirstOrDefault();
            if (xfrm?.Element(A + "ext") is { } aext) { aext.SetAttributeValue("cx", E(w)); aext.SetAttributeValue("cy", E(h)); }
        }
        foreach (var s in VmlShapes(x))
        {
            var css = DrawingReader.Css(s.Attribute("style")?.Value);
            css["width"] = N(w) + "pt";
            css["height"] = N(h) + "pt";
            s.SetAttributeValue("style", Css(css));
        }
    });

    static string Css(Dictionary<string, string> css) => string.Join(";", css.Select(kv => $"{kv.Key}:{kv.Value}"));

    // ================================================================ position et habillage

    /// <summary>Position d'un objet flottant : décalages en points depuis les repères donnés (relativeFrom).</summary>
    public static void SetPosition(DrawingContent d, string hFrom, double x, string vFrom, double y) => Edit(d, xml =>
    {
        foreach (var a in Frames(xml).Where(f => f.Name == Wp + "anchor"))
        {
            a.SetAttributeValue("simplePos", "0");
            Replace(a, Wp + "positionH", new XElement(Wp + "positionH", new XAttribute("relativeFrom", hFrom), new XElement(Wp + "posOffset", E(x))));
            Replace(a, Wp + "positionV", new XElement(Wp + "positionV", new XAttribute("relativeFrom", vFrom), new XElement(Wp + "posOffset", E(y))));
        }
        foreach (var s in VmlShapes(xml))
        {
            var css = DrawingReader.Css(s.Attribute("style")?.Value);
            if (css.GetValueOrDefault("position") != "absolute") continue;
            css["margin-left"] = N(x) + "pt";
            css["margin-top"] = N(y) + "pt";
            css.Remove("left");
            css.Remove("top");
            css["mso-position-horizontal"] = "absolute";
            css["mso-position-vertical"] = "absolute";
            css["mso-position-horizontal-relative"] = hFrom switch { "page" => "page", "margin" => "margin", "character" => "char", _ => "text" };
            css["mso-position-vertical-relative"] = vFrom switch { "page" => "page", "margin" => "margin", "line" => "line", _ => "text" };
            s.SetAttributeValue("style", Css(css));
        }
    });

    /// <summary>Alignement d'un objet flottant (left, center, right ; top, center, bottom) dans un repère.</summary>
    public static void SetAlignment(DrawingContent d, string hFrom, string hAlign, string vFrom, string vAlign) => Edit(d, xml =>
    {
        foreach (var a in Frames(xml).Where(f => f.Name == Wp + "anchor"))
        {
            Replace(a, Wp + "positionH", new XElement(Wp + "positionH", new XAttribute("relativeFrom", hFrom), new XElement(Wp + "align", hAlign)));
            Replace(a, Wp + "positionV", new XElement(Wp + "positionV", new XAttribute("relativeFrom", vFrom), new XElement(Wp + "align", vAlign)));
        }
    });

    /// <summary>Remplace (ou ajoute à sa place dans l'ordre du schéma) un enfant de wp:anchor.</summary>
    static void Replace(XElement anchor, XName name, XElement value)
    {
        if (anchor.Element(name) is { } old) { old.ReplaceWith(value); return; }
        var order = new[] { "simplePos", "positionH", "positionV", "extent", "effectExtent", "wrapNone", "wrapSquare", "wrapTight", "wrapThrough", "wrapTopAndBottom", "docPr", "cNvGraphicFramePr", "graphic" };
        int idx = Array.IndexOf(order, name.LocalName);
        var after = anchor.Elements().LastOrDefault(e => e.Name.Namespace == Wp && Array.IndexOf(order, e.Name.LocalName) is var k && k >= 0 && k < idx);
        if (after is not null) after.AddAfterSelf(value); else anchor.AddFirst(value);
    }

    static readonly string[] WrapNames = { "wrapNone", "wrapSquare", "wrapTight", "wrapThrough", "wrapTopAndBottom" };

    /// <summary>
    /// Habillage d'un objet flottant : carré, rapproché, au travers, haut et bas, ou aucun (devant ou derrière le texte).
    /// Rapproché / au travers : polygone du cadre (rectangle) s'il n'y en a pas.
    /// </summary>
    public static void SetWrap(DrawingContent d, WrapKind wrap, bool behind = false, string side = "bothSides") => Edit(d, xml =>
    {
        foreach (var a in Frames(xml).Where(f => f.Name == Wp + "anchor"))
        {
            var oldPoly = a.Elements().FirstOrDefault(e => e.Name.Namespace == Wp && e.Name.LocalName is "wrapTight" or "wrapThrough")?.Element(Wp + "wrapPolygon");
            foreach (var n in WrapNames) a.Element(Wp + n)?.Remove();
            XElement el = wrap switch
            {
                WrapKind.Square => new XElement(Wp + "wrapSquare", new XAttribute("wrapText", side)),
                WrapKind.Tight or WrapKind.Through => new XElement(Wp + (wrap == WrapKind.Tight ? "wrapTight" : "wrapThrough"), new XAttribute("wrapText", side),
                    oldPoly is not null ? new XElement(oldPoly) : new XElement(Wp + "wrapPolygon", new XAttribute("edited", "0"),
                        new XElement(Wp + "start", new XAttribute("x", 0), new XAttribute("y", 0)),
                        new XElement(Wp + "lineTo", new XAttribute("x", 0), new XAttribute("y", 21600)),
                        new XElement(Wp + "lineTo", new XAttribute("x", 21600), new XAttribute("y", 21600)),
                        new XElement(Wp + "lineTo", new XAttribute("x", 21600), new XAttribute("y", 0)),
                        new XElement(Wp + "lineTo", new XAttribute("x", 0), new XAttribute("y", 0)))),
                WrapKind.TopAndBottom => new XElement(Wp + "wrapTopAndBottom"),
                _ => new XElement(Wp + "wrapNone"),
            };
            Replace(a, el.Name, el);
            a.SetAttributeValue("behindDoc", behind && wrap == WrapKind.None ? "1" : "0");
            // distances par défaut de Word : 0,125 pouce à gauche et à droite pour carré / rapproché
            if (wrap is WrapKind.Square or WrapKind.Tight or WrapKind.Through)
            {
                if (a.Attribute("distL")?.Value is null or "0") a.SetAttributeValue("distL", "114300");
                if (a.Attribute("distR")?.Value is null or "0") a.SetAttributeValue("distR", "114300");
            }
        }
        foreach (var s in VmlShapes(xml))
        {
            s.Element(W10 + "wrap")?.Remove();
            if (wrap != WrapKind.None)
                s.Add(new XElement(W10 + "wrap", new XAttribute("type", wrap switch
                {
                    WrapKind.Square => "square", WrapKind.Tight => "tight", WrapKind.Through => "through", _ => "topAndBottom",
                })));
            var css = DrawingReader.Css(s.Attribute("style")?.Value);
            if (css.TryGetValue("z-index", out var z) && double.TryParse(z, NumberStyles.Float, CultureInfo.InvariantCulture, out var zv))
                css["z-index"] = ((long)(behind ? -Math.Abs(zv) : Math.Abs(zv))).ToString(CultureInfo.InvariantCulture);
            s.SetAttributeValue("style", Css(css));
        }
    });

    /// <summary>
    /// Objet en ligne → flottant, à la position donnée (points depuis la colonne et le paragraphe) : wp:inline devient
    /// wp:anchor comme Word (simplePos, positions, habillage, z au-dessus des autres).
    /// </summary>
    public static void ToAnchor(DrawingContent d, double x, double y, WrapKind wrap, bool behind, long z)
    {
        Edit(d, xml =>
        {
            foreach (var inl in Frames(xml).Where(f => f.Name == Wp + "inline").ToList())
            {
                var a = new XElement(Wp + "anchor",
                    new XAttribute("distT", inl.Attribute("distT")?.Value ?? "0"), new XAttribute("distB", inl.Attribute("distB")?.Value ?? "0"),
                    new XAttribute("distL", "114300"), new XAttribute("distR", "114300"),
                    new XAttribute("simplePos", "0"), new XAttribute("relativeHeight", z), new XAttribute("behindDoc", behind ? "1" : "0"),
                    new XAttribute("locked", "0"), new XAttribute("layoutInCell", "1"), new XAttribute("allowOverlap", "1"),
                    inl.Attributes().Where(at => at.Name.Namespace == Wp14).Select(at => new XAttribute(at)),
                    new XElement(Wp + "simplePos", new XAttribute("x", 0), new XAttribute("y", 0)),
                    new XElement(Wp + "positionH", new XAttribute("relativeFrom", "column"), new XElement(Wp + "posOffset", E(x))),
                    new XElement(Wp + "positionV", new XAttribute("relativeFrom", "paragraph"), new XElement(Wp + "posOffset", E(y))),
                    inl.Elements().Select(e => new XElement(e)));
                inl.ReplaceWith(a);
            }
        });
        SetWrap(d, wrap, behind);
    }

    /// <summary>Objet flottant → en ligne avec le texte (wp:anchor devient wp:inline ; VML : position statique).</summary>
    public static void ToInline(DrawingContent d) => Edit(d, xml =>
    {
        foreach (var a in Frames(xml).Where(f => f.Name == Wp + "anchor").ToList())
        {
            var keep = new[] { "extent", "effectExtent", "docPr", "cNvGraphicFramePr", "graphic" };
            var inl = new XElement(Wp + "inline",
                new XAttribute("distT", "0"), new XAttribute("distB", "0"), new XAttribute("distL", "0"), new XAttribute("distR", "0"),
                a.Attributes().Where(at => at.Name.Namespace == Wp14).Select(at => new XAttribute(at)),
                a.Elements().Where(e => e.Name.Namespace == Wp && keep.Contains(e.Name.LocalName) || e.Name == A + "graphic").Select(e => new XElement(e)));
            if (inl.Element(Wp + "effectExtent") is null) inl.Element(Wp + "extent")?.AddAfterSelf(new XElement(Wp + "effectExtent", new XAttribute("l", 0), new XAttribute("t", 0), new XAttribute("r", 0), new XAttribute("b", 0)));
            a.ReplaceWith(inl);
        }
        foreach (var s in VmlShapes(xml))
        {
            var css = DrawingReader.Css(s.Attribute("style")?.Value);
            foreach (var k in new[] { "position", "margin-left", "margin-top", "left", "top", "z-index", "mso-position-horizontal", "mso-position-vertical",
                         "mso-position-horizontal-relative", "mso-position-vertical-relative" }) css.Remove(k);
            s.SetAttributeValue("style", Css(css));
            s.Element(W10 + "wrap")?.Remove();
        }
    });

    /// <summary>Ordre d'empilement (wp:anchor/@relativeHeight, z-index VML).</summary>
    public static void SetZ(DrawingContent d, long z) => Edit(d, xml =>
    {
        foreach (var a in Frames(xml).Where(f => f.Name == Wp + "anchor")) a.SetAttributeValue("relativeHeight", Math.Max(0, z));
        foreach (var s in VmlShapes(xml))
        {
            var css = DrawingReader.Css(s.Attribute("style")?.Value);
            bool behind = css.TryGetValue("z-index", out var old) && old.StartsWith('-');
            css["z-index"] = (behind ? -Math.Max(1, z) : Math.Max(1, z)).ToString(CultureInfo.InvariantCulture);
            s.SetAttributeValue("style", Css(css));
        }
    });

    // ================================================================ apparence

    /// <summary>Élément racine du graphique DrawingML (pic:pic, wps:wsp, wpg:wgp…) de chaque variante.</summary>
    static IEnumerable<XElement> Graphics(XElement x) =>
        Frames(x).Select(f => f.Element(A + "graphic")?.Element(A + "graphicData")?.Elements().FirstOrDefault()).Where(e => e is not null)!;

    static XElement? SpPr(XElement g) => g.Element(Pic + "spPr") ?? g.Element(Wps + "spPr");

    static XElement EnsureXfrm(XElement spPr)
    {
        if (spPr.Element(A + "xfrm") is { } x) return x;
        x = new XElement(A + "xfrm", new XElement(A + "off", new XAttribute("x", 0), new XAttribute("y", 0)), new XElement(A + "ext", new XAttribute("cx", 0), new XAttribute("cy", 0)));
        spPr.AddFirst(x);
        return x;
    }

    /// <summary>Rotation en degrés (sens horaire) ; VML : style rotation.</summary>
    public static void SetRotation(DrawingContent d, double degrees) => Edit(d, xml =>
    {
        degrees = ((degrees % 360) + 360) % 360;
        foreach (var g in Graphics(xml))
        {
            var spPr = SpPr(g) ?? g.Element(Wpg + "grpSpPr");
            if (spPr is null) continue;
            var xfrm = EnsureXfrm(spPr);
            if (degrees == 0) xfrm.Attribute("rot")?.Remove();
            else xfrm.SetAttributeValue("rot", ((long)Math.Round(degrees * 60000)).ToString(CultureInfo.InvariantCulture));
        }
        foreach (var s in VmlShapes(xml))
        {
            var css = DrawingReader.Css(s.Attribute("style")?.Value);
            if (degrees == 0) css.Remove("rotation"); else css["rotation"] = N(degrees);
            s.SetAttributeValue("style", Css(css));
        }
    });

    /// <summary>Retournement horizontal ou vertical (bascule).</summary>
    public static void Flip(DrawingContent d, bool horizontal) => Edit(d, xml =>
    {
        var attr = horizontal ? "flipH" : "flipV";
        foreach (var g in Graphics(xml))
        {
            var spPr = SpPr(g) ?? g.Element(Wpg + "grpSpPr");
            if (spPr is null) continue;
            var xfrm = EnsureXfrm(spPr);
            if (xfrm.Attribute(attr)?.Value is "1" or "true") xfrm.Attribute(attr)!.Remove();
            else xfrm.SetAttributeValue(attr, "1");
        }
    });

    /// <summary>Rognage d'une image (fractions de chaque bord, 0..1).</summary>
    public static void SetCrop(DrawingContent d, double l, double t, double r, double b) => Edit(d, xml =>
    {
        foreach (var g in Graphics(xml).Where(g => g.Name == Pic + "pic"))
        {
            var fill = g.Element(Pic + "blipFill");
            if (fill is null) continue;
            fill.Element(A + "srcRect")?.Remove();
            if (l != 0 || t != 0 || r != 0 || b != 0)
            {
                string P(double v) => ((long)Math.Round(v * 100000)).ToString(CultureInfo.InvariantCulture);
                var rect = new XElement(A + "srcRect");
                if (l != 0) rect.SetAttributeValue("l", P(l));
                if (t != 0) rect.SetAttributeValue("t", P(t));
                if (r != 0) rect.SetAttributeValue("r", P(r));
                if (b != 0) rect.SetAttributeValue("b", P(b));
                (fill.Element(A + "blip") is { } blip ? (Action<XElement>)blip.AddAfterSelf : fill.AddFirst)(rect);
            }
        }
    });

    /// <summary>Texte de remplacement (wp:docPr/@descr) et titre.</summary>
    public static void SetAltText(DrawingContent d, string? description, string? title = null) => Edit(d, xml =>
    {
        foreach (var docPr in xml.Descendants(Wp + "docPr"))
        {
            if (string.IsNullOrEmpty(description)) docPr.Attribute("descr")?.Remove(); else docPr.SetAttributeValue("descr", description);
            if (title is not null) { if (title.Length == 0) docPr.Attribute("title")?.Remove(); else docPr.SetAttributeValue("title", title); }
        }
        foreach (var s in VmlShapes(xml))
            if (string.IsNullOrEmpty(description)) s.Attribute("alt")?.Remove(); else s.SetAttributeValue("alt", description);
    });

    static readonly string[] FillNames = { "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill" };

    /// <summary>Remplissage uni RRGGBB d'une forme (null : sans remplissage).</summary>
    public static void SetFill(DrawingContent d, string? rgb) => Edit(d, xml =>
    {
        foreach (var g in Graphics(xml).Where(g => g.Name == Wps + "wsp"))
        {
            var spPr = g.Element(Wps + "spPr")!;
            foreach (var n in FillNames) spPr.Element(A + n)?.Remove();
            var fill = rgb is null ? new XElement(A + "noFill") : new XElement(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", rgb)));
            // ordre du schéma : xfrm, géométrie, remplissage, ln…
            var geom = spPr.Element(A + "prstGeom") ?? spPr.Element(A + "custGeom") ?? spPr.Element(A + "xfrm");
            if (geom is not null) geom.AddAfterSelf(fill); else spPr.AddFirst(fill);
        }
        foreach (var s in VmlShapes(xml))
        {
            if (rgb is null) s.SetAttributeValue("filled", "f");
            else { s.SetAttributeValue("filled", "t"); s.SetAttributeValue("fillcolor", "#" + rgb); }
        }
    });

    /// <summary>Contour : couleur RRGGBB (null : aucun), épaisseur en points, tirets (prstDash).</summary>
    public static void SetLine(DrawingContent d, string? rgb, double width = 0.75, string dash = "solid") => Edit(d, xml =>
    {
        foreach (var g in Graphics(xml).Where(g => g.Name == Wps + "wsp" || g.Name == Pic + "pic"))
        {
            var spPr = SpPr(g);
            if (spPr is null) continue;
            spPr.Element(A + "ln")?.Remove();
            var ln = rgb is null ? new XElement(A + "ln", new XElement(A + "noFill"))
                : new XElement(A + "ln", new XAttribute("w", E(width)), new XElement(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", rgb))),
                    dash != "solid" ? new XElement(A + "prstDash", new XAttribute("val", dash)) : null);
            var before = spPr.Elements().LastOrDefault(e => e.Name.Namespace == A && (FillNames.Contains(e.Name.LocalName) || e.Name.LocalName is "xfrm" or "prstGeom" or "custGeom"));
            if (before is not null) before.AddAfterSelf(ln); else spPr.Add(ln);
        }
        foreach (var s in VmlShapes(xml))
        {
            if (rgb is null) s.SetAttributeValue("stroked", "f");
            else { s.SetAttributeValue("stroked", "t"); s.SetAttributeValue("strokecolor", "#" + rgb); s.SetAttributeValue("strokeweight", N(width) + "pt"); }
        }
    });

    /// <summary>Corrections d'une image : luminosité et contraste (-1..1), niveaux de gris.</summary>
    public static void SetCorrections(DrawingContent d, double brightness, double contrast, bool grayscale) => Edit(d, xml =>
    {
        foreach (var blip in Graphics(xml).Where(g => g.Name == Pic + "pic").Select(g => g.Element(Pic + "blipFill")?.Element(A + "blip")).Where(b => b is not null))
        {
            blip!.Element(A + "lum")?.Remove();
            blip.Element(A + "grayscl")?.Remove();
            if (grayscale) blip.AddFirst(new XElement(A + "grayscl"));
            if (brightness != 0 || contrast != 0)
            {
                var lum = new XElement(A + "lum");
                if (brightness != 0) lum.SetAttributeValue("bright", ((long)Math.Round(brightness * 100000)).ToString(CultureInfo.InvariantCulture));
                if (contrast != 0) lum.SetAttributeValue("contrast", ((long)Math.Round(contrast * 100000)).ToString(CultureInfo.InvariantCulture));
                var ext = blip.Element(A + "extLst");
                if (ext is not null) ext.AddBeforeSelf(lum); else blip.Add(lum);
            }
        }
    });

    /// <summary>Forme prédéfinie d'une forme ou du cadre d'une image (rect, ellipse, roundRect…).</summary>
    public static void SetGeometry(DrawingContent d, string preset) => Edit(d, xml =>
    {
        foreach (var g in Graphics(xml))
        {
            var spPr = SpPr(g);
            if (spPr is null) continue;
            spPr.Element(A + "custGeom")?.Remove();
            var geom = new XElement(A + "prstGeom", new XAttribute("prst", preset), new XElement(A + "avLst"));
            if (spPr.Element(A + "prstGeom") is { } old) old.ReplaceWith(geom);
            else if (spPr.Element(A + "xfrm") is { } xf) xf.AddAfterSelf(geom);
            else spPr.AddFirst(geom);
        }
    });

    // ================================================================ identifiants

    /// <summary>Plus grand wp:docPr/@id des dessins d'histoires (corps, en-têtes…), zones de texte comprises.</summary>
    public static long MaxDocPrId(IEnumerable<Story> stories)
    {
        long max = 0;
        foreach (var d in stories.SelectMany(AllDrawings))
            foreach (var p in d.Xml.Descendants(Wp + "docPr"))
                if (long.TryParse(p.Attribute("id")?.Value, out var v)) max = Math.Max(max, v);
        return max;
    }

    /// <summary>Plus grand ordre d'empilement des objets flottants (pour placer un nouvel objet devant les autres).</summary>
    public static long MaxZ(IEnumerable<Story> stories)
    {
        long max = 251658240;                    // valeur de départ de Word
        foreach (var d in stories.SelectMany(AllDrawings))
            foreach (var a in d.Xml.Descendants(Wp + "anchor"))
                if (long.TryParse(a.Attribute("relativeHeight")?.Value, out var v)) max = Math.Max(max, v);
        return max;
    }

    /// <summary>Dessins d'une histoire, et des zones de texte qu'ils contiennent.</summary>
    public static IEnumerable<DrawingContent> AllDrawings(Story story)
    {
        foreach (var p in story.Paragraphs())
            foreach (var r in p.Runs())
                foreach (var d in r.Content.OfType<DrawingContent>())
                {
                    yield return d;
                    foreach (var box in d.TextBoxes)
                        foreach (var inner in AllDrawings(box)) yield return inner;
                }
    }

    /// <summary>Renumérote les wp:docPr d'un dessin (copie collée) à partir de <paramref name="next"/>.</summary>
    public static void RenumberIds(DrawingContent d, Func<long> next) => Edit(d, xml =>
    {
        foreach (var docPr in xml.Descendants(Wp + "docPr")) docPr.SetAttributeValue("id", next());
        foreach (var c in xml.Descendants().Where(e => e.Name.LocalName == "cNvPr" && e.Attribute("id") is not null)) c.SetAttributeValue("id", next());
    });

    // ================================================================ création

    /// <summary>Image en ligne (w:drawing/wp:inline/pic:pic) comme l'insère Word.</summary>
    public static XElement CreatePicture(XNamespace w, long id, string name, string relId, double width, double height, string? description = null) =>
        new(w + "drawing",
            new XElement(Wp + "inline", new XAttribute("distT", 0), new XAttribute("distB", 0), new XAttribute("distL", 0), new XAttribute("distR", 0),
                new XElement(Wp + "extent", new XAttribute("cx", E(width)), new XAttribute("cy", E(height))),
                new XElement(Wp + "effectExtent", new XAttribute("l", 0), new XAttribute("t", 0), new XAttribute("r", 0), new XAttribute("b", 0)),
                new XElement(Wp + "docPr", new XAttribute("id", id), new XAttribute("name", name), description is null ? null : new XAttribute("descr", description)),
                new XElement(Wp + "cNvGraphicFramePr", new XElement(A + "graphicFrameLocks", new XAttribute(XNamespace.Xmlns + "a", A.NamespaceName), new XAttribute("noChangeAspect", 1))),
                new XElement(A + "graphic", new XAttribute(XNamespace.Xmlns + "a", A.NamespaceName),
                    new XElement(A + "graphicData", new XAttribute("uri", Pic.NamespaceName),
                        new XElement(Pic + "pic", new XAttribute(XNamespace.Xmlns + "pic", Pic.NamespaceName),
                            new XElement(Pic + "nvPicPr", new XElement(Pic + "cNvPr", new XAttribute("id", 0), new XAttribute("name", name)), new XElement(Pic + "cNvPicPr")),
                            new XElement(Pic + "blipFill", new XElement(A + "blip", new XAttribute(R + "embed", relId)), new XElement(A + "stretch", new XElement(A + "fillRect"))),
                            new XElement(Pic + "spPr",
                                new XElement(A + "xfrm", new XElement(A + "off", new XAttribute("x", 0), new XAttribute("y", 0)),
                                    new XElement(A + "ext", new XAttribute("cx", E(width)), new XAttribute("cy", E(height)))),
                                new XElement(A + "prstGeom", new XAttribute("prst", "rect"), new XElement(A + "avLst"))))))));

    /// <summary>
    /// Forme flottante (wps:wsp) comme Word en dessine : devant le texte, remplissage et contour du thème (accent 1),
    /// texte centré en blanc ; zone de texte (<paramref name="textBox"/>) : fond blanc, trait noir fin, habillage carré,
    /// paragraphe vide à écrire. Position : décalages en points depuis la colonne et le paragraphe.
    /// </summary>
    public static XElement CreateShape(XNamespace w, long id, string preset, double x, double y, double width, double height, long z, bool textBox = false)
    {
        bool line = preset is "line" or "straightConnector1";
        var name = textBox ? $"Zone de texte {id}" : $"{FrenchName(preset)} {id}";
        var spPr = new XElement(Wps + "spPr",
            new XElement(A + "xfrm", new XElement(A + "off", new XAttribute("x", 0), new XAttribute("y", 0)), new XElement(A + "ext", new XAttribute("cx", E(width)), new XAttribute("cy", E(height)))),
            new XElement(A + "prstGeom", new XAttribute("prst", preset), new XElement(A + "avLst")));
        if (textBox)
        {
            spPr.Add(new XElement(A + "solidFill", new XElement(A + "schemeClr", new XAttribute("val", "lt1"))),
                     new XElement(A + "ln", new XAttribute("w", 6350), new XElement(A + "solidFill", new XElement(A + "prstClr", new XAttribute("val", "black")))));
        }
        var wsp = new XElement(Wps + "wsp",
            new XElement(Wps + "cNvSpPr", textBox ? new XAttribute("txBox", "1") : null),
            spPr,
            textBox ? null : new XElement(Wps + "style",
                new XElement(A + "lnRef", new XAttribute("idx", 2), new XElement(A + "schemeClr", new XAttribute("val", "accent1"), new XElement(A + "shade", new XAttribute("val", 50000)))),
                new XElement(A + "fillRef", new XAttribute("idx", line ? 0 : 1), new XElement(A + "schemeClr", new XAttribute("val", "accent1"))),
                new XElement(A + "effectRef", new XAttribute("idx", 0), new XElement(A + "schemeClr", new XAttribute("val", "accent1"))),
                new XElement(A + "fontRef", new XAttribute("idx", "minor"), new XElement(A + "schemeClr", new XAttribute("val", "lt1")))),
            textBox ? new XElement(Wps + "txbx", new XElement(w + "txbxContent", new XElement(w + "p"))) : null,
            new XElement(Wps + "bodyPr", new XAttribute("rot", 0), new XAttribute("vert", "horz"), new XAttribute("wrap", "square"),
                new XAttribute("lIns", 91440), new XAttribute("tIns", 45720), new XAttribute("rIns", 91440), new XAttribute("bIns", 45720),
                new XAttribute("anchor", textBox ? "t" : "ctr"), new XAttribute("anchorCtr", 0), new XElement(A + "noAutofit")));
        var anchor = new XElement(Wp + "anchor",
            new XAttribute("distT", textBox ? 45720 : 0), new XAttribute("distB", textBox ? 45720 : 0), new XAttribute("distL", 114300), new XAttribute("distR", 114300),
            new XAttribute("simplePos", 0), new XAttribute("relativeHeight", z), new XAttribute("behindDoc", 0), new XAttribute("locked", 0),
            new XAttribute("layoutInCell", 1), new XAttribute("allowOverlap", 1),
            new XElement(Wp + "simplePos", new XAttribute("x", 0), new XAttribute("y", 0)),
            new XElement(Wp + "positionH", new XAttribute("relativeFrom", "column"), new XElement(Wp + "posOffset", E(x))),
            new XElement(Wp + "positionV", new XAttribute("relativeFrom", "paragraph"), new XElement(Wp + "posOffset", E(y))),
            new XElement(Wp + "extent", new XAttribute("cx", E(width)), new XAttribute("cy", E(height))),
            new XElement(Wp + "effectExtent", new XAttribute("l", 0), new XAttribute("t", 0), new XAttribute("r", 0), new XAttribute("b", 0)),
            textBox ? new XElement(Wp + "wrapSquare", new XAttribute("wrapText", "bothSides")) : new XElement(Wp + "wrapNone"),
            new XElement(Wp + "docPr", new XAttribute("id", id), new XAttribute("name", name)),
            new XElement(Wp + "cNvGraphicFramePr"),
            new XElement(A + "graphic", new XAttribute(XNamespace.Xmlns + "a", A.NamespaceName),
                new XElement(A + "graphicData", new XAttribute("uri", Wps.NamespaceName), wsp)));
        // mc:AlternateContent comme Word : la forme wps n'est lue que par Word 2010 et suivants
        return new XElement(Mc + "AlternateContent",
            new XElement(Mc + "Choice", new XAttribute("Requires", "wps"), new XElement(w + "drawing", anchor)));
    }

    /// <summary>Nom français d'une forme (nom de l'objet dans le volet Sélection de Word).</summary>
    public static string FrenchName(string preset) => preset switch
    {
        "rect" => "Rectangle", "roundRect" => "Rectangle : coins arrondis", "ellipse" => "Ellipse", "triangle" => "Triangle isocèle",
        "rtTriangle" => "Triangle rectangle", "diamond" => "Losange", "pentagon" => "Pentagone", "hexagon" => "Hexagone",
        "octagon" => "Octogone", "line" or "straightConnector1" => "Connecteur droit", "rightArrow" => "Flèche : droite",
        "leftArrow" => "Flèche : gauche", "upArrow" => "Flèche : haut", "downArrow" => "Flèche : bas", "star5" => "Étoile : 5 branches",
        "heart" => "Cœur", "smileyFace" => "Émoticône", "cloud" => "Nuage", "wedgeRectCallout" => "Bulle narrative : rectangle",
        _ => "Forme",
    };
}
