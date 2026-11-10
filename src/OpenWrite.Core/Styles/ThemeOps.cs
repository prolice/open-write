using System.Xml.Linq;

namespace OpenWrite.Core.Styles;

/// <summary>
/// Thème du document (theme1.xml, lot 4) : polices des titres et du corps, palette de couleurs. Les styles qui utilisent
/// les polices et couleurs du thème (titres, citations…) suivent ; le texte en couleurs ou polices fixes ne change pas.
/// </summary>
public static class ThemeOps
{
    static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    /// <summary>Paires de polices (titres, corps) proposées par Création ▸ Polices.</summary>
    public static readonly (string Name, string Major, string Minor)[] FontPairs =
    {
        ("Office", "Aptos Display", "Aptos"), ("Office 2013-2022", "Calibri Light", "Calibri"), ("Office 2007-2010", "Cambria", "Calibri"),
        ("Arial", "Arial", "Arial"), ("Times New Roman", "Times New Roman", "Times New Roman"), ("Georgia", "Georgia", "Georgia"),
        ("Garamond", "Garamond", "Garamond"), ("Century Gothic", "Century Gothic", "Century Gothic"), ("Segoe UI", "Segoe UI Light", "Segoe UI"),
        ("Corbel", "Corbel", "Corbel"), ("Candara", "Candara", "Candara"), ("Franklin Gothic", "Franklin Gothic Medium", "Franklin Gothic Book"),
        ("Trebuchet MS", "Trebuchet MS", "Trebuchet MS"), ("Verdana", "Verdana", "Verdana"),
    };

    /// <summary>Palettes : texte/arrière-plan foncé 2 et clair 2, accents 1 à 6, lien, lien visité.</summary>
    public static readonly (string Name, string[] Colors)[] Palettes =
    {
        ("Office", new[] { "0E2841", "E8E8E8", "156082", "E97132", "196B24", "0F9ED5", "A02B93", "4EA72E", "467886", "96607D" }),
        ("Office 2013-2022", new[] { "44546A", "E7E6E6", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72" }),
        ("Office 2007-2010", new[] { "1F497D", "EEECE1", "4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646", "0000FF", "800080" }),
        ("Bleu", new[] { "17406D", "DBEFF9", "0F6FC6", "009DD9", "0BD0D9", "10CF9B", "7CCA62", "A5C249", "F49100", "85DFD0" }),
        ("Vert", new[] { "455F51", "E3DED1", "549E39", "8AB833", "C0CF3A", "029676", "4AB5C4", "0989B1", "6B9F25", "BA6906" }),
        ("Rouge orangé", new[] { "505046", "EEECE1", "E84C22", "FFBD47", "B64926", "FF8427", "CC9900", "B22600", "CC9900", "666699" }),
        ("Violet", new[] { "373545", "DCD8DC", "AD84C6", "8784C7", "5D739A", "6997AF", "84ACB6", "6F8183", "69A020", "8C8C8C" }),
        ("Nuances de gris", new[] { "000000", "F8F8F8", "DDDDDD", "B2B2B2", "969696", "808080", "5F5F5F", "4D4D4D", "5F5F5F", "919191" }),
    };

    static readonly string[] Slots = { "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink" };

    /// <summary>Polices du thème actuelles (titres, corps).</summary>
    public static (string? Major, string? Minor) Fonts(XElement theme)
    {
        var fs = theme.Element(A + "themeElements")?.Element(A + "fontScheme");
        return (fs?.Element(A + "majorFont")?.Element(A + "latin")?.Attribute("typeface")?.Value,
                fs?.Element(A + "minorFont")?.Element(A + "latin")?.Attribute("typeface")?.Value);
    }

    /// <summary>Change les polices latines des titres et du corps (les polices d'autres écritures restent).</summary>
    public static bool SetFonts(XElement theme, string name, string major, string minor)
    {
        var fs = theme.Element(A + "themeElements")?.Element(A + "fontScheme");
        if (fs is null) return false;
        fs.SetAttributeValue("name", name);
        foreach (var (group, face) in new[] { ("majorFont", major), ("minorFont", minor) })
            if (fs.Element(A + group)?.Element(A + "latin") is { } latin) latin.SetAttributeValue("typeface", face);
        return true;
    }

    /// <summary>Remplace la palette (foncé 2, clair 2, accents, liens) ; texte 1 et arrière-plan 1 restent.</summary>
    public static bool SetColors(XElement theme, string name, string[] colors)
    {
        var cs = theme.Element(A + "themeElements")?.Element(A + "clrScheme");
        if (cs is null || colors.Length != Slots.Length) return false;
        cs.SetAttributeValue("name", name);
        for (int i = 0; i < Slots.Length; i++)
            if (cs.Element(A + Slots[i]) is { } slot) slot.ReplaceNodes(new XElement(A + "srgbClr", new XAttribute("val", colors[i])));
        return true;
    }

    /// <summary>Accents du thème (pour les sélecteurs de couleur) : nom, valeur.</summary>
    public static List<(string Name, string Rgb)> Accents(XElement theme)
    {
        var cs = theme.Element(A + "themeElements")?.Element(A + "clrScheme");
        var res = new List<(string, string)>();
        for (int i = 1; i <= 6; i++)
            if (cs?.Element(A + $"accent{i}")?.Element(A + "srgbClr")?.Attribute("val")?.Value is { } v) res.Add(($"Accent {i}", v.ToUpperInvariant()));
        return res;
    }
}
