using System.Globalization;
using System.Xml.Linq;
using OpenWrite.Core.Model;

namespace OpenWrite.Core.Editing;

/// <summary>
/// Sections (lot 5) : la section d'un paragraphe (son w:sectPr, porté par le dernier paragraphe de la section ou par le
/// corps pour la dernière), taille, orientation, marges, colonnes, numérotation des pages, sauts de section.
/// Le w:sectPr se modifie en place (ses enfants inconnus restent).
/// </summary>
public static class SectionOps
{
    static string Tw(double pt) => Math.Round(pt * 20).ToString(CultureInfo.InvariantCulture);

    /// <summary>
    /// w:sectPr de la section qui contient le bloc de premier niveau <paramref name="topIndex"/> : celui du premier paragraphe
    /// à partir de ce bloc qui en porte un, sinon celui du corps (créé s'il manque).
    /// </summary>
    public static XElement SectPrFor(Story body, int topIndex, XNamespace w)
    {
        for (int i = Math.Max(0, topIndex); i < body.Blocks.Count; i++)
            if (body.Blocks[i] is Paragraph p && p.Props.Get("sectPr") is { } sp) return sp;
        if (body.FinalSection?.Element is { } final) return final;
        var created = new XElement(w + "sectPr",
            new XElement(w + "pgSz", new XAttribute(w + "w", 11906), new XAttribute(w + "h", 16838)),
            new XElement(w + "pgMar", new XAttribute(w + "top", 1417), new XAttribute(w + "right", 1417), new XAttribute(w + "bottom", 1417),
                new XAttribute(w + "left", 1417), new XAttribute(w + "header", 708), new XAttribute(w + "footer", 708), new XAttribute(w + "gutter", 0)));
        body.FinalSection = new PropertyBag("sectPr", w, created);
        return created;
    }

    /// <summary>Tous les w:sectPr du corps, dans l'ordre (le dernier : celui du corps).</summary>
    public static List<XElement> All(Story body, XNamespace w)
    {
        var list = body.Blocks.OfType<Paragraph>().Select(p => p.Props.Get("sectPr")).Where(s => s is not null).Select(s => s!).ToList();
        list.Add(SectPrFor(body, body.Blocks.Count, w));
        return list;
    }

    /// <summary>Ordre des enfants de w:sectPr (schéma) : références d'en-têtes et de pieds, puis propriétés de page.</summary>
    static readonly string[] Order =
    {
        "headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType",
        "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange",
    };

    /// <summary>Élément enfant (créé à sa place s'il manque).</summary>
    public static XElement Child(XElement sectPr, string local)
    {
        var w = sectPr.Name.Namespace;
        if (sectPr.Element(w + local) is { } e) return e;
        e = new XElement(w + local);
        int rank = Array.IndexOf(Order, local);
        var after = sectPr.Elements().FirstOrDefault(x => x.Name.Namespace == w && Array.IndexOf(Order, x.Name.LocalName) > rank);
        if (after is null) sectPr.Add(e); else after.AddBeforeSelf(e);
        return e;
    }

    static double Get(XElement? e, XNamespace w, string attr, double defTwips) =>
        e?.Attribute(w + attr)?.Value is { } v && double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? d / 20 : defTwips / 20;

    /// <summary>Taille de page (points) et orientation actuelles.</summary>
    public static (double W, double H, bool Landscape) PageSize(XElement sectPr)
    {
        var w = sectPr.Name.Namespace;
        var sz = sectPr.Element(w + "pgSz");
        double pw = Get(sz, w, "w", 11906), ph = Get(sz, w, "h", 16838);
        return (pw, ph, sz?.Attribute(w + "orient")?.Value == "landscape" || pw > ph);
    }

    /// <summary>Taille du papier (points, dans le sens portrait) ; l'orientation actuelle est gardée.</summary>
    public static void SetPaper(XElement sectPr, double widthPt, double heightPt)
    {
        var w = sectPr.Name.Namespace;
        bool landscape = PageSize(sectPr).Landscape;
        var sz = Child(sectPr, "pgSz");
        double small = Math.Min(widthPt, heightPt), big = Math.Max(widthPt, heightPt);
        sz.SetAttributeValue(w + "w", Tw(landscape ? big : small));
        sz.SetAttributeValue(w + "h", Tw(landscape ? small : big));
    }

    /// <summary>Portrait ou paysage : largeur et hauteur échangées, marges tournées comme Word (haut → gauche…).</summary>
    public static void SetOrientation(XElement sectPr, bool landscape)
    {
        var w = sectPr.Name.Namespace;
        var (pw, ph, now) = PageSize(sectPr);
        if (now == landscape) return;
        var sz = Child(sectPr, "pgSz");
        sz.SetAttributeValue(w + "w", Tw(ph));
        sz.SetAttributeValue(w + "h", Tw(pw));
        sz.SetAttributeValue(w + "orient", landscape ? "landscape" : null);
        if (sectPr.Element(w + "pgMar") is { } mar)
        {
            string? t = mar.Attribute(w + "top")?.Value, b = mar.Attribute(w + "bottom")?.Value, l = mar.Attribute(w + "left")?.Value, r = mar.Attribute(w + "right")?.Value;
            if (landscape) { mar.SetAttributeValue(w + "top", l); mar.SetAttributeValue(w + "bottom", r); mar.SetAttributeValue(w + "left", b); mar.SetAttributeValue(w + "right", t); }
            else { mar.SetAttributeValue(w + "left", t); mar.SetAttributeValue(w + "right", b); mar.SetAttributeValue(w + "bottom", l); mar.SetAttributeValue(w + "top", r); }
        }
    }

    /// <summary>Marges (points ; null : inchangée), en-tête et pied compris.</summary>
    public static void SetMargins(XElement sectPr, double? top = null, double? bottom = null, double? left = null, double? right = null, double? header = null, double? footer = null)
    {
        var w = sectPr.Name.Namespace;
        var mar = Child(sectPr, "pgMar");
        void Set(string a, double? v) { if (v is { } x) mar.SetAttributeValue(w + a, Tw(x)); }
        Set("top", top); Set("bottom", bottom); Set("left", left); Set("right", right); Set("header", header); Set("footer", footer);
        if (mar.Attribute(w + "gutter") is null) mar.SetAttributeValue(w + "gutter", 0);
    }

    public static (double Top, double Bottom, double Left, double Right, double Header, double Footer) Margins(XElement sectPr)
    {
        var w = sectPr.Name.Namespace;
        var m = sectPr.Element(w + "pgMar");
        return (Get(m, w, "top", 1417), Get(m, w, "bottom", 1417), Get(m, w, "left", 1417), Get(m, w, "right", 1417), Get(m, w, "header", 708), Get(m, w, "footer", 708));
    }

    /// <summary>
    /// Colonnes : nombre égal (espacement en points), ou deux colonnes inégales (« gauche » / « droite » de Word : la
    /// petite fait la moitié de la grande) ; trait entre les colonnes.
    /// </summary>
    public static void SetColumns(XElement sectPr, int count, double spacePt = 36, string? uneven = null, bool separator = false)
    {
        var w = sectPr.Name.Namespace;
        var cols = Child(sectPr, "cols");
        cols.RemoveAll();
        cols.SetAttributeValue(w + "space", Tw(spacePt));
        if (separator) cols.SetAttributeValue(w + "sep", "1");
        if (uneven is null || count != 2)
        {
            if (count > 1) cols.SetAttributeValue(w + "num", count);
            return;
        }
        var (pw, _, _) = PageSize(sectPr);
        var (_, _, l, r, _, _) = Margins(sectPr);
        double text = pw - l - r - spacePt;
        double small = text / 3, big = text - small;
        cols.SetAttributeValue(w + "num", 2);
        cols.SetAttributeValue(w + "equalWidth", "0");
        var (first, second) = uneven == "left" ? (small, big) : (big, small);
        cols.Add(new XElement(w + "col", new XAttribute(w + "w", Tw(first)), new XAttribute(w + "space", Tw(spacePt))),
                 new XElement(w + "col", new XAttribute(w + "w", Tw(second))));
    }

    /// <summary>Numérotation des pages de la section : départ (null : suite de la précédente) et format (decimal, lowerRoman…).</summary>
    public static void SetPageNumbering(XElement sectPr, int? start, string? format)
    {
        var w = sectPr.Name.Namespace;
        var pn = Child(sectPr, "pgNumType");
        pn.SetAttributeValue(w + "start", start?.ToString(CultureInfo.InvariantCulture));
        pn.SetAttributeValue(w + "fmt", format is null or "decimal" ? null : format);
        if (!pn.HasAttributes) pn.Remove();
    }

    /// <summary>
    /// Saut de section au point d'insertion (paragraphe coupé là) : la section avant le saut garde les propriétés et les
    /// en-têtes actuels ; celle d'après commence selon <paramref name="type"/> (nextPage, continuous, evenPage, oddPage)
    /// et hérite des en-têtes (comme un nouveau saut dans Word). Renvoie le début du paragraphe d'après.
    /// </summary>
    public static DocPos InsertBreak(Story body, DocPos pos, string type, Func<string> newParaId)
    {
        if (pos.Path.Count != 1) throw new NotSupportedException("Un saut de section ne se place pas dans un tableau.");
        var w = pos.Paragraph(body).Props.Ns;
        var before = new XElement(SectPrFor(body, pos.BlockIndex, w));
        var after = TextOps.Split(body, pos, newParaId);
        // section d'après (le paragraphe coupé a pu emporter son w:sectPr dans la seconde moitié) : ses références
        // d'en-têtes sont retirées (héritées de la section d'avant) et son type change
        var current = SectPrFor(body, pos.BlockIndex + 1, w);
        current.Elements().Where(e => e.Name.LocalName is "headerReference" or "footerReference").Remove();
        var t = Child(current, "type");
        t.SetAttributeValue(w + "val", type);
        if (type == "nextPage") t.Remove();
        ((Paragraph)body.Blocks[pos.BlockIndex]).Props.Set(before);
        return after;
    }
}
