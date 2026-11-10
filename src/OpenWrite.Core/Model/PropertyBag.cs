using System.Xml.Linq;

namespace OpenWrite.Core.Model;

/// <summary>
/// Propriétés d'un nœud (w:pPr, w:rPr, w:tblPr, w:trPr, w:tcPr, w:sectPr…) : l'élément XML d'origine est conservé
/// entier — ce qu'OpenWrite ne comprend pas (extensions w14/w15, propriétés rares) repart donc tel quel.
/// Les modifications insèrent l'élément à sa place dans l'ordre du schéma (Word refuse un ordre incorrect).
/// </summary>
public sealed class PropertyBag
{
    public const string TransitionalNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public PropertyBag(string localName, XNamespace? ns = null, XElement? element = null)
    {
        LocalName = localName;
        Ns = ns ?? TransitionalNs;
        Element = element;
    }

    public string LocalName { get; }
    /// <summary>Espace de noms WordprocessingML du document (transitionnel ou strict ISO 29500).</summary>
    public XNamespace Ns { get; }
    /// <summary>Élément d'origine ; null si le nœud n'a pas de propriétés.</summary>
    public XElement? Element { get; private set; }

    public bool IsEmpty => Element is null || !Element.HasElements && !Element.HasAttributes;

    public XElement? Get(string local) => Element?.Element(Ns + local);

    /// <summary>Valeur de l'attribut w:val d'une propriété (ex. Val("pStyle")).</summary>
    public string? Val(string local) => Get(local)?.Attribute(Ns + "val")?.Value;

    /// <summary>Propriété booléenne (w:b, w:i…) : présente sans w:val, ou w:val vrai.</summary>
    public bool Flag(string local) => Get(local) is { } e && (e.Attribute(Ns + "val")?.Value is null or "1" or "true" or "on");

    /// <summary>Remplace (ou ajoute) une propriété, à sa place dans l'ordre du schéma.</summary>
    public void Set(XElement child)
    {
        Element ??= new XElement(Ns + LocalName);
        var existing = Element.Element(child.Name);
        if (existing is not null) { existing.ReplaceWith(child); return; }
        double rank = Rank(child.Name);
        var after = Element.Elements().FirstOrDefault(e => Rank(e.Name) > rank);
        if (after is null) Element.Add(child); else after.AddBeforeSelf(child);
    }

    public void SetFlag(string local, bool on)
    {
        if (on) Set(new XElement(Ns + local));
        else Remove(local);
    }

    public void SetVal(string local, string? value)
    {
        if (value is null) Remove(local);
        else Set(new XElement(Ns + local, new XAttribute(Ns + "val", value)));
    }

    public void Remove(string local) => Get(local)?.Remove();

    public PropertyBag Clone() => new(LocalName, Ns, Element is null ? null : new XElement(Element));

    /// <summary>
    /// Rang d'un élément dans la séquence du schéma (ECMA-376 partie 1, CT_PPrBase, CT_RPr…). Les éléments
    /// inconnus (extensions d'autres espaces de noms) se placent après les éléments connus mais avant
    /// l'historique des modifications (pPrChange, rPrChange…), qui est toujours dernier.
    /// </summary>
    double Rank(XName name)
    {
        if (!Orders.TryGetValue(LocalName, out var order)) return 0;    // conteneur sans ordre connu : à la fin
        if (name.Namespace == Ns)
        {
            int i = Array.IndexOf(order, name.LocalName);
            if (i >= 0) return i;
        }
        return order.Length - 1.5;
    }

    static readonly Dictionary<string, string[]> Orders = new()
    {
        ["pPr"] = new[]
        {
            "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers",
            "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE",
            "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
            "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle",
            "rPr", "sectPr", "pPrChange",
        },
        ["rPr"] = new[]
        {
            "ins", "del", "moveFrom", "moveTo",
            "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss",
            "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs",
            "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
            "specVanish", "oMath", "rPrChange",
        },
        ["tblPr"] = new[]
        {
            "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc",
            "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription",
            "tblPrChange",
        },
        ["trPr"] = new[]
        {
            "cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit", "trHeight", "tblHeader",
            "tblCellSpacing", "jc", "hidden", "ins", "del", "trPrChange",
        },
        ["tcPr"] = new[]
        {
            "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar", "textDirection",
            "tcFitText", "vAlign", "hideMark", "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange",
        },
    };
}
