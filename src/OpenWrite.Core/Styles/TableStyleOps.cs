using System.Xml.Linq;

namespace OpenWrite.Core.Styles;

/// <summary>
/// Styles de tableau intégrés (lot 6) : « Grille du tableau » et les familles de la galerie de Word (Grille 1 Clair,
/// Grille 4, Grille 5 Foncé, Liste 3, Liste 4 ; sans accent et accents 1 à 6), écrits dans styles.xml au premier usage
/// comme le fait Word (avec « Tableau Normal », leur parent). Reconnus par leur nom intégré anglais (w:name).
/// </summary>
public static class TableStyleOps
{
    /// <summary>Couleurs du thème Office par défaut, quand le document n'a pas de thème.</summary>
    static readonly Dictionary<string, string> OfficeColors = new()
    {
        ["text1"] = "000000", ["background1"] = "FFFFFF", ["text2"] = "44546A", ["background2"] = "E7E6E6",
        ["accent1"] = "4472C4", ["accent2"] = "ED7D31", ["accent3"] = "A5A5A5", ["accent4"] = "FFC000", ["accent5"] = "5B9BD5", ["accent6"] = "70AD47",
    };

    /// <summary>
    /// Nuances des couleurs du thème Office telles que Word les écrit dans ses styles de tableau (couleur:tint|shade:valeur
    /// → RRGGBB) : des constantes d'Office, que la formule ne retrouve pas toutes à une unité près (6 sur 33).
    /// </summary>
    static readonly Dictionary<string, string> OfficeShades = new()
    {
        ["accent1:shade:BF"] = "2F5496", ["accent1:tint:33"] = "D9E2F3", ["accent1:tint:66"] = "B4C6E7", ["accent1:tint:99"] = "8EAADB",
        ["accent2:tint:33"] = "FBE4D5", ["accent2:tint:66"] = "F7CAAC", ["accent2:tint:99"] = "F4B083", ["accent3:tint:33"] = "EDEDED",
        ["accent3:tint:66"] = "DBDBDB", ["accent3:tint:99"] = "C9C9C9", ["accent4:tint:33"] = "FFF2CC", ["accent4:tint:66"] = "FFE599",
        ["accent4:tint:99"] = "FFD966", ["accent5:tint:33"] = "DEEAF6", ["accent5:tint:66"] = "BDD6EE", ["accent5:tint:99"] = "9CC2E5",
        ["accent6:tint:33"] = "E2EFD9", ["accent6:tint:66"] = "C5E0B3", ["accent6:tint:99"] = "A8D08D", ["text1:tint:33"] = "CCCCCC",
        ["text1:tint:66"] = "999999", ["text1:tint:99"] = "666666", ["text1:tint:A6"] = "595959", ["text1:tint:BF"] = "404040",
        ["text1:tint:D8"] = "272727",
    };

    /// <summary>Styles de la galerie, dans l'ordre de Word : nom intégré, nom français.</summary>
    public static IEnumerable<(string Name, string French)> Gallery()
    {
        yield return ("Table Grid", "Grille du tableau");
        foreach (var f in TableStyleDefs.Families)
        {
            yield return (f.Name, f.French);
            for (int a = 1; a <= 6; a++) yield return ($"{f.Name} Accent {a}", $"{f.French} - Accentuation {a}");
        }
    }

    /// <summary>Nom français d'un style de tableau intégré (null s'il n'en est pas un).</summary>
    public static string? French(string? name) =>
        name is null ? null : name.Equals("Normal Table", StringComparison.OrdinalIgnoreCase) ? "Tableau Normal"
        : Gallery().FirstOrDefault(g => g.Name.Equals(name, StringComparison.OrdinalIgnoreCase)).French;

    /// <summary>Accent (1…6) d'une variante, 0 sans accent ; famille.</summary>
    static ((string Name, string French, string Id, string Neutral, string Accent) Family, int Accent)? Parse(string name)
    {
        foreach (var f in TableStyleDefs.Families)
        {
            if (name.Equals(f.Name, StringComparison.OrdinalIgnoreCase)) return (f, 0);
            for (int a = 1; a <= 6; a++)
                if (name.Equals($"{f.Name} Accent {a}", StringComparison.OrdinalIgnoreCase)) return (f, a);
        }
        return null;
    }

    public static bool IsBuiltIn(string name) =>
        name.Equals("Table Grid", StringComparison.OrdinalIgnoreCase) || name.Equals("Normal Table", StringComparison.OrdinalIgnoreCase) || Parse(name) is not null;

    /// <summary>
    /// Identifiant du style de tableau intégré <paramref name="name"/>, ajouté à styles.xml s'il n'y est pas (avec
    /// « Tableau Normal »). Null si le nom n'est pas un style de la galerie.
    /// </summary>
    public static string? Ensure(StyleSheet sheet, XElement stylesRoot, string name)
    {
        var w = sheet.W;
        if (Find(stylesRoot, w, name) is { } existing) return existing.Attribute(w + "styleId")?.Value;
        string body, id;
        bool isDefault = false;
        if (name.Equals("Normal Table", StringComparison.OrdinalIgnoreCase))
        {
            (body, id, isDefault) = (TableStyleDefs.NormalTable, "TableNormal", !stylesRoot.Elements(w + "style").Any(s => s.Attribute(w + "type")?.Value == "table" && s.Attribute(w + "default")?.Value is "1" or "true"));
            name = "Normal Table";
        }
        else if (name.Equals("Table Grid", StringComparison.OrdinalIgnoreCase)) (body, id, name) = (TableStyleDefs.TableGrid, "TableGrid", "Table Grid");
        else if (Parse(name) is { } p)
        {
            body = p.Accent == 0 ? p.Family.Neutral : p.Family.Accent.Replace("{A}", $"accent{p.Accent}");
            id = p.Accent == 0 ? p.Family.Id : $"{p.Family.Id}-Accent{p.Accent}";
            name = p.Accent == 0 ? p.Family.Name : $"{p.Family.Name} Accent {p.Accent}";
        }
        else return null;
        // parent : le style de tableau par défaut du document, sinon « Tableau Normal » créé
        string? basedOn = null;
        if (name != "Normal Table")
            basedOn = stylesRoot.Elements(w + "style").FirstOrDefault(s => s.Attribute(w + "type")?.Value == "table" && s.Attribute(w + "default")?.Value is "1" or "true")
                ?.Attribute(w + "styleId")?.Value ?? Ensure(sheet, stylesRoot, "Normal Table");
        id = UniqueId(stylesRoot, w, id);
        var xml = $"<w:style xmlns:w=\"{w.NamespaceName}\" w:type=\"table\"{(isDefault ? " w:default=\"1\"" : "")} w:styleId=\"{id}\"><w:name w:val=\"{name}\"/>"
                + (basedOn is not null ? $"<w:basedOn w:val=\"{basedOn}\"/>" : "") + body + "</w:style>";
        var style = XElement.Parse(xml);
        FillThemeColors(sheet, style);
        foreach (var x in style.DescendantsAndSelf())
            x.Attributes().Where(a => a.IsNamespaceDeclaration && a.Value == w.NamespaceName).Remove();
        stylesRoot.Add(style);
        return id;
    }

    /// <summary>
    /// Valeurs RRGGBB des couleurs du thème (w:color, w:val d'une couleur de texte, w:fill d'une trame) pour le thème du
    /// document. Nuance en luminance TSL comme le résolveur (claire : L × t + 1 − t ; foncée : L × s), mais les
    /// définitions de styles de tableau de Word tronquent les composantes au lieu de les arrondir (accent1 4472C4,
    /// w:themeTint="99" → 8EAADB et non 8FAADC) ; pour le thème Office, les valeurs exactes de Word (OfficeShades) —
    /// vérifié sur les 36 styles de la galerie (OpenWrite.Fidelity --table-styles).
    /// </summary>
    public static void FillThemeColors(StyleSheet sheet, XElement root)
    {
        var w = sheet.W;
        string? Hex(string? theme, string? tint, string? shade)
        {
            if (theme is null) return null;
            var rgb = sheet.ThemeColor(theme) ?? OfficeColors.GetValueOrDefault(theme);
            if (rgb is not { Length: 6 }) return null;
            double? t = Byte(tint), sh = Byte(shade);
            if (t is null && sh is null) return rgb.ToUpperInvariant();
            // thème Office : valeurs de Word
            if (OfficeColors.TryGetValue(theme, out var office) && rgb.Equals(office, StringComparison.OrdinalIgnoreCase)
                && OfficeShades.TryGetValue($"{theme}:{(tint is not null ? "tint" : "shade")}:{(tint ?? shade)!.ToUpperInvariant()}", out var known))
                return known;
            return Luminance(rgb, l => t is { } tv ? l * tv + (1 - tv) : l * sh!.Value);
        }
        foreach (var e in root.Descendants())
        {
            if (e.Attribute(w + "themeColor")?.Value is { } tc && Hex(tc, e.Attribute(w + "themeTint")?.Value, e.Attribute(w + "themeShade")?.Value) is { } hex)
            {
                // couleur de texte : w:val ; trait de bordure : w:color, placé avant les attributs de thème (ordre de Word)
                var attr = e.Name.LocalName == "color" ? w + "val" : w + "color";
                SetBefore(e, attr, hex, w + "themeColor");
            }
            if (e.Attribute(w + "themeFill")?.Value is { } tf && Hex(tf, e.Attribute(w + "themeFillTint")?.Value, e.Attribute(w + "themeFillShade")?.Value) is { } fill)
                SetBefore(e, w + "fill", fill, w + "themeFill");
        }
    }

    /// <summary>Luminance TSL modifiée, composantes tronquées (définitions de styles de tableau de Word).</summary>
    static string Luminance(string rgb, Func<double, double> f)
    {
        double r = Convert.ToInt32(rgb[..2], 16) / 255.0, g = Convert.ToInt32(rgb[2..4], 16) / 255.0, b = Convert.ToInt32(rgb[4..], 16) / 255.0;
        double max = Math.Max(r, Math.Max(g, b)), min = Math.Min(r, Math.Min(g, b)), l = (max + min) / 2, h = 0, s = 0;
        if (max != min)
        {
            double d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
            h = max == r ? (g - b) / d + (g < b ? 6 : 0) : max == g ? (b - r) / d + 2 : (r - g) / d + 4;
            h /= 6;
        }
        l = Math.Clamp(f(l), 0, 1);
        double Hue(double p, double q, double t)
        {
            if (t < 0) t += 1;
            if (t > 1) t -= 1;
            return t < 1 / 6.0 ? p + (q - p) * 6 * t : t < 0.5 ? q : t < 2 / 3.0 ? p + (q - p) * (2 / 3.0 - t) * 6 : p;
        }
        if (s == 0) r = g = b = l;
        else
        {
            double q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
            r = Hue(p, q, h + 1 / 3.0);
            g = Hue(p, q, h);
            b = Hue(p, q, h - 1 / 3.0);
        }
        static int C(double v) => Math.Clamp((int)Math.Floor(v * 255 + 1e-6), 0, 255);
        return $"{C(r):X2}{C(g):X2}{C(b):X2}";
    }

    static double? Byte(string? hex) =>
        hex is not null && int.TryParse(hex, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out var v) ? v / 255.0 : null;

    static void SetBefore(XElement e, XName name, string value, XName before)
    {
        if (e.Attribute(name) is { } a) { a.Value = value; return; }
        var attrs = e.Attributes().ToList();
        int i = attrs.FindIndex(x => x.Name == before);
        attrs.Insert(i < 0 ? attrs.Count : i, new XAttribute(name, value));
        e.ReplaceAttributes(attrs);
    }

    static XElement? Find(XElement root, XNamespace w, string name) => root.Elements(w + "style").FirstOrDefault(s =>
        s.Attribute(w + "type")?.Value == "table" && string.Equals(s.Element(w + "name")?.Attribute(w + "val")?.Value, name, StringComparison.OrdinalIgnoreCase));

    static string UniqueId(XElement root, XNamespace w, string wanted)
    {
        var ids = root.Elements(w + "style").Select(s => s.Attribute(w + "styleId")?.Value).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var id = wanted;
        for (int k = 2; ids.Contains(id); k++) id = wanted + k;
        return id;
    }
}
