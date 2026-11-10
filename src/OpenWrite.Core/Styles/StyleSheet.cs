using System.Xml.Linq;

namespace OpenWrite.Core.Styles;

/// <summary>
/// Parties qui définissent la mise en forme héritée : styles (w:styles), numérotation (w:numbering), thème (a:theme)
/// et paramètres (w:settings, pour w:clrSchemeMapping). Lecture seule ; chaque partie peut manquer.
/// </summary>
public sealed class StyleSheet
{
    public StyleSheet(XElement? styles, XElement? numbering = null, XElement? theme = null, XElement? settings = null, XNamespace? w = null)
    {
        W = w ?? styles?.Name.Namespace ?? "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        Styles = styles;
        Numbering = numbering;
        Theme = theme;
        Settings = settings;
        if (styles is not null)
            foreach (var s in styles.Elements(W + "style"))
                if (s.Attribute(W + "styleId")?.Value is { } id) _byId.TryAdd(id, s);
    }

    public XNamespace W { get; }
    public XElement? Styles { get; }
    public XElement? Numbering { get; }
    public XElement? Theme { get; }
    public XElement? Settings { get; }

    readonly Dictionary<string, XElement> _byId = new(StringComparer.Ordinal);

    public XElement? Style(string? id) => id is not null && _byId.TryGetValue(id, out var s) ? s : null;

    /// <summary>Type du style : paragraph, character, table, numbering.</summary>
    public string? TypeOf(XElement style) => style.Attribute(W + "type")?.Value;

    /// <summary>Style par défaut d'un type (w:default="1") : Normal, Police par défaut, Tableau normal.</summary>
    public XElement? DefaultStyle(string type) => Styles?.Elements(W + "style")
        .FirstOrDefault(s => TypeOf(s) == type && s.Attribute(W + "default")?.Value is "1" or "true" or "on");

    /// <summary>Nom du style (w:name) : nom anglais intégré (« heading 1 ») ou nom personnalisé.</summary>
    public string? NameOf(string id) => Style(id)?.Element(W + "name")?.Attribute(W + "val")?.Value;

    /// <summary>Chaîne basedOn, de l'ancêtre le plus lointain au style lui-même (protégée contre les boucles).</summary>
    public List<XElement> Chain(string? id)
    {
        var res = new List<XElement>();
        var seen = new HashSet<string>();
        for (var s = Style(id); s is not null && seen.Add(s.Attribute(W + "styleId")!.Value); s = Style(s.Element(W + "basedOn")?.Attribute(W + "val")?.Value))
            res.Insert(0, s);
        return res;
    }

    public XElement? DefaultRunProps => Styles?.Element(W + "docDefaults")?.Element(W + "rPrDefault")?.Element(W + "rPr");
    public XElement? DefaultParagraphProps => Styles?.Element(W + "docDefaults")?.Element(W + "pPrDefault")?.Element(W + "pPr");

    // ---------------------------------------------------------------- numérotation

    /// <summary>
    /// Niveau de liste (w:lvl) pour un numéro et un niveau : remplacement de niveau du w:num s'il existe, sinon le niveau
    /// de l'abstractNum ; un abstractNum lié à un style de numérotation (w:numStyleLink) est suivi une fois.
    /// </summary>
    public XElement? ListLevel(int numId, int ilvl)
    {
        if (Numbering is null || numId <= 0) return null;
        var num = Numbering.Elements(W + "num").FirstOrDefault(n => n.Attribute(W + "numId")?.Value == numId.ToString());
        if (num is null) return null;
        var over = num.Elements(W + "lvlOverride").FirstOrDefault(o => o.Attribute(W + "ilvl")?.Value == ilvl.ToString())?.Element(W + "lvl");
        if (over is not null) return over;
        var abs = AbstractNum(num.Element(W + "abstractNumId")?.Attribute(W + "val")?.Value);
        if (abs?.Element(W + "numStyleLink")?.Attribute(W + "val")?.Value is { } link
            && Style(link)?.Element(W + "pPr")?.Element(W + "numPr")?.Element(W + "numId")?.Attribute(W + "val")?.Value is { } linkedNum
            && int.TryParse(linkedNum, out var ln) && ln != numId)
        {
            var num2 = Numbering.Elements(W + "num").FirstOrDefault(n => n.Attribute(W + "numId")?.Value == ln.ToString());
            abs = AbstractNum(num2?.Element(W + "abstractNumId")?.Attribute(W + "val")?.Value) ?? abs;
        }
        return abs?.Elements(W + "lvl").FirstOrDefault(l => l.Attribute(W + "ilvl")?.Value == ilvl.ToString());
    }

    XElement? AbstractNum(string? id) => id is null ? null
        : Numbering!.Elements(W + "abstractNum").FirstOrDefault(a => a.Attribute(W + "abstractNumId")?.Value == id);

    // ---------------------------------------------------------------- thème

    static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>Police du thème : majorHAnsi, minorEastAsia, majorBidi… (latin pour Ascii/HAnsi, ea, cs).</summary>
    public string? ThemeFont(string themeFont)
    {
        var scheme = Theme?.Element(A + "themeElements")?.Element(A + "fontScheme");
        if (scheme is null) return null;
        var group = scheme.Element(A + (themeFont.StartsWith("major") ? "majorFont" : "minorFont"));
        var slot = themeFont.EndsWith("EastAsia") ? "ea" : themeFont.EndsWith("Bidi") ? "cs" : "latin";
        var face = group?.Element(A + slot)?.Attribute("typeface")?.Value;
        return string.IsNullOrEmpty(face) ? null : face;
    }

    /// <summary>Couleur du thème (RRGGBB) pour une valeur w:themeColor (text1, accent1…), via w:clrSchemeMapping.</summary>
    public string? ThemeColor(string themeColor)
    {
        var map = Settings?.Element(W + "clrSchemeMapping");
        string Mapped(string attr, string def) => map?.Attribute(W + attr)?.Value ?? def;
        var slot = themeColor switch
        {
            "text1" => Mapped("t1", "dark1"),
            "background1" => Mapped("bg1", "light1"),
            "text2" => Mapped("t2", "dark2"),
            "background2" => Mapped("bg2", "light2"),
            _ => themeColor,
        };
        var schemeName = slot switch
        {
            "dark1" => "dk1", "light1" => "lt1", "dark2" => "dk2", "light2" => "lt2",
            "hyperlink" => "hlink", "followedHyperlink" => "folHlink",
            _ => slot,                                                   // accent1…accent6
        };
        var c = Theme?.Element(A + "themeElements")?.Element(A + "clrScheme")?.Element(A + schemeName);
        var srgb = c?.Element(A + "srgbClr")?.Attribute("val")?.Value;
        var sys = c?.Element(A + "sysClr")?.Attribute("lastClr")?.Value;
        return (srgb ?? sys)?.ToUpperInvariant();
    }
}
