using System.Xml.Linq;
using OpenWrite.Core.Model;

namespace OpenWrite.Core.Styles;

/// <summary>
/// Styles (lot 4) : noms affichés en français, styles intégrés de Word ajoutés à styles.xml au premier usage (comme Word,
/// qui ne les écrit que lorsqu'ils servent), création et modification de styles. Les styles sont reconnus par leur nom
/// intégré (w:name, « heading 1 »…), jamais par leur identifiant, qui dépend de la langue de Word (« Titre1 », « Heading1 »).
/// </summary>
public static class StyleOps
{
    /// <summary>Styles intégrés : nom Word → nom français affiché (celui de Word en français).</summary>
    public static readonly IReadOnlyDictionary<string, string> FrenchNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        ["Normal"] = "Normal", ["heading 1"] = "Titre 1", ["heading 2"] = "Titre 2", ["heading 3"] = "Titre 3", ["heading 4"] = "Titre 4",
        ["heading 5"] = "Titre 5", ["heading 6"] = "Titre 6", ["heading 7"] = "Titre 7", ["heading 8"] = "Titre 8", ["heading 9"] = "Titre 9",
        ["Title"] = "Titre", ["Subtitle"] = "Sous-titre", ["Quote"] = "Citation", ["Intense Quote"] = "Citation intense",
        ["List Paragraph"] = "Paragraphe de liste", ["No Spacing"] = "Sans interligne", ["Strong"] = "Élevé", ["Emphasis"] = "Accentuation",
        ["Subtle Emphasis"] = "Accentuation légère", ["Intense Emphasis"] = "Accentuation intense", ["Hyperlink"] = "Lien hypertexte",
        ["FollowedHyperlink"] = "Lien hypertexte visité", ["Default Paragraph Font"] = "Police par défaut", ["Table Grid"] = "Grille du tableau",
        ["Normal Table"] = "Tableau normal", ["No List"] = "Aucune liste", ["caption"] = "Légende", ["header"] = "En-tête", ["footer"] = "Pied de page",
        ["footnote text"] = "Note de bas de page", ["footnote reference"] = "Appel de note de bas de page", ["endnote text"] = "Note de fin",
        ["endnote reference"] = "Appel de note de fin", ["annotation text"] = "Commentaire", ["toc 1"] = "TM 1", ["toc 2"] = "TM 2", ["toc 3"] = "TM 3",
        ["TOC Heading"] = "En-tête de table des matières", ["List Bullet"] = "Liste à puces", ["List Number"] = "Liste numérotée",
        ["Balloon Text"] = "Texte de bulles", ["Book Title"] = "Titre du livre", ["Subtle Reference"] = "Référence légère",
        ["Intense Reference"] = "Référence intense", ["Body Text"] = "Corps de texte", ["Plain Text"] = "Texte brut",
    };

    /// <summary>Nom affiché d'un style : nom français du style intégré, sinon le nom du document, sinon l'identifiant.</summary>
    public static string DisplayName(StyleSheet sheet, string? id)
    {
        if (id is null) return "Normal";
        var name = sheet.NameOf(id);
        return name is null ? id : FrenchNames.TryGetValue(name, out var fr) ? fr : name;
    }

    /// <summary>Style du document qui porte ce nom intégré (« heading 1 »), quel que soit son identifiant.</summary>
    public static XElement? ByName(StyleSheet sheet, string builtInName)
    {
        var w = sheet.W;
        return sheet.Styles?.Elements(w + "style").FirstOrDefault(s =>
            string.Equals(s.Element(w + "name")?.Attribute(w + "val")?.Value, builtInName, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>Galerie des styles rapides de Word (dans l'ordre de Word), complétée par les styles qFormat du document.</summary>
    public static readonly string[] QuickStyles =
        { "Normal", "No Spacing", "heading 1", "heading 2", "heading 3", "Title", "Subtitle", "Subtle Emphasis", "Emphasis", "Intense Emphasis",
          "Strong", "Quote", "Intense Quote", "List Paragraph" };

    // ------------------------------------------------------------------ définitions des styles intégrés

    sealed record BuiltIn(string Type, string Id, string? BasedOn, string? Next, int Priority, string PPr, string RPr);

    /// <summary>
    /// Définitions de Word 2013 et suivants (modèle Normal, thème Office) pour les styles de la galerie, écrites dans
    /// styles.xml quand un style est appliqué pour la première fois. Polices et couleurs du thème du document.
    /// </summary>
    static readonly Dictionary<string, BuiltIn> BuiltIns = new(StringComparer.OrdinalIgnoreCase)
    {
        ["heading 1"] = new("paragraph", "Titre1", "Normal", "Normal", 9,
            "<w:keepNext/><w:keepLines/><w:spacing w:before=\"240\" w:after=\"0\"/><w:outlineLvl w:val=\"0\"/>",
            "<w:rFonts w:asciiTheme=\"majorHAnsi\" w:eastAsiaTheme=\"majorEastAsia\" w:hAnsiTheme=\"majorHAnsi\" w:cstheme=\"majorBidi\"/><w:color w:val=\"2F5496\" w:themeColor=\"accent1\" w:themeShade=\"BF\"/><w:sz w:val=\"32\"/><w:szCs w:val=\"32\"/>"),
        ["heading 2"] = new("paragraph", "Titre2", "Normal", "Normal", 9,
            "<w:keepNext/><w:keepLines/><w:spacing w:before=\"40\" w:after=\"0\"/><w:outlineLvl w:val=\"1\"/>",
            "<w:rFonts w:asciiTheme=\"majorHAnsi\" w:eastAsiaTheme=\"majorEastAsia\" w:hAnsiTheme=\"majorHAnsi\" w:cstheme=\"majorBidi\"/><w:color w:val=\"2F5496\" w:themeColor=\"accent1\" w:themeShade=\"BF\"/><w:sz w:val=\"26\"/><w:szCs w:val=\"26\"/>"),
        ["heading 3"] = new("paragraph", "Titre3", "Normal", "Normal", 9,
            "<w:keepNext/><w:keepLines/><w:spacing w:before=\"40\" w:after=\"0\"/><w:outlineLvl w:val=\"2\"/>",
            "<w:rFonts w:asciiTheme=\"majorHAnsi\" w:eastAsiaTheme=\"majorEastAsia\" w:hAnsiTheme=\"majorHAnsi\" w:cstheme=\"majorBidi\"/><w:color w:val=\"1F3763\" w:themeColor=\"accent1\" w:themeShade=\"7F\"/><w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/>"),
        ["heading 4"] = new("paragraph", "Titre4", "Normal", "Normal", 9,
            "<w:keepNext/><w:keepLines/><w:spacing w:before=\"40\" w:after=\"0\"/><w:outlineLvl w:val=\"3\"/>",
            "<w:rFonts w:asciiTheme=\"majorHAnsi\" w:eastAsiaTheme=\"majorEastAsia\" w:hAnsiTheme=\"majorHAnsi\" w:cstheme=\"majorBidi\"/><w:i/><w:iCs/><w:color w:val=\"2F5496\" w:themeColor=\"accent1\" w:themeShade=\"BF\"/>"),
        ["Title"] = new("paragraph", "Titre", "Normal", "Normal", 10,
            "<w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/><w:contextualSpacing/>",
            "<w:rFonts w:asciiTheme=\"majorHAnsi\" w:eastAsiaTheme=\"majorEastAsia\" w:hAnsiTheme=\"majorHAnsi\" w:cstheme=\"majorBidi\"/><w:spacing w:val=\"-10\"/><w:kern w:val=\"28\"/><w:sz w:val=\"56\"/><w:szCs w:val=\"56\"/>"),
        ["Subtitle"] = new("paragraph", "Sous-titre", "Normal", "Normal", 11,
            "<w:spacing w:after=\"160\"/>",
            "<w:rFonts w:eastAsiaTheme=\"minorEastAsia\"/><w:color w:val=\"5A5A5A\" w:themeColor=\"text1\" w:themeTint=\"A5\"/><w:spacing w:val=\"15\"/><w:sz w:val=\"22\"/><w:szCs w:val=\"22\"/>"),
        ["Quote"] = new("paragraph", "Citation", "Normal", "Normal", 29,
            "<w:spacing w:before=\"200\" w:after=\"160\"/><w:ind w:left=\"864\" w:right=\"864\"/><w:jc w:val=\"center\"/>",
            "<w:i/><w:iCs/><w:color w:val=\"404040\" w:themeColor=\"text1\" w:themeTint=\"BF\"/>"),
        ["Intense Quote"] = new("paragraph", "Citationintense", "Normal", "Normal", 30,
            "<w:pBdr><w:top w:val=\"single\" w:sz=\"4\" w:space=\"10\" w:color=\"4472C4\" w:themeColor=\"accent1\"/><w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"10\" w:color=\"4472C4\" w:themeColor=\"accent1\"/></w:pBdr><w:spacing w:before=\"360\" w:after=\"360\"/><w:ind w:left=\"864\" w:right=\"864\"/><w:jc w:val=\"center\"/>",
            "<w:i/><w:iCs/><w:color w:val=\"4472C4\" w:themeColor=\"accent1\"/>"),
        ["List Paragraph"] = new("paragraph", "Paragraphedeliste", "Normal", null, 34, "<w:ind w:left=\"720\"/><w:contextualSpacing/>", ""),
        ["No Spacing"] = new("paragraph", "Sansinterligne", null, null, 1, "<w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/>", ""),
        ["Strong"] = new("character", "lev", "Policepardfaut", null, 22, "", "<w:b/><w:bCs/>"),
        ["Emphasis"] = new("character", "Accentuation", "Policepardfaut", null, 20, "", "<w:i/><w:iCs/>"),
        ["Subtle Emphasis"] = new("character", "Accentuationlgre", "Policepardfaut", null, 19, "", "<w:i/><w:iCs/><w:color w:val=\"404040\" w:themeColor=\"text1\" w:themeTint=\"BF\"/>"),
        ["Intense Emphasis"] = new("character", "Accentuationintense", "Policepardfaut", null, 21, "", "<w:i/><w:iCs/><w:color w:val=\"4472C4\" w:themeColor=\"accent1\"/>"),
        ["Hyperlink"] = new("character", "Lienhypertexte", "Policepardfaut", null, 99, "", "<w:color w:val=\"0563C1\" w:themeColor=\"hyperlink\"/><w:u w:val=\"single\"/>"),
        ["Default Paragraph Font"] = new("character", "Policepardfaut", null, null, 1, "", ""),
        ["annotation text"] = new("paragraph", "Commentaire", "Normal", null, 99, "<w:spacing w:line=\"240\" w:lineRule=\"auto\"/>", "<w:sz w:val=\"20\"/><w:szCs w:val=\"20\"/>"),
        ["annotation reference"] = new("character", "Marquedecommentaire", "Policepardfaut", null, 99, "", "<w:sz w:val=\"16\"/><w:szCs w:val=\"16\"/>"),
        ["header"] = new("paragraph", "En-tte", "Normal", null, 99, "<w:tabs><w:tab w:val=\"center\" w:pos=\"4536\"/><w:tab w:val=\"right\" w:pos=\"9072\"/></w:tabs><w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/>", ""),
        ["footer"] = new("paragraph", "Pieddepage", "Normal", null, 99, "<w:tabs><w:tab w:val=\"center\" w:pos=\"4536\"/><w:tab w:val=\"right\" w:pos=\"9072\"/></w:tabs><w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/>", ""),
    };

    /// <summary>Type (paragraph / character) d'un style intégré connu, null s'il ne l'est pas.</summary>
    public static string? BuiltInType(string name) => BuiltIns.TryGetValue(name, out var b) ? b.Type : null;

    public static bool IsKnownBuiltIn(string name) => BuiltIns.ContainsKey(name);

    /// <summary>
    /// Identifiant du style qui porte ce nom intégré, en l'ajoutant à styles.xml (avec ses parents) s'il n'y est pas encore.
    /// Null si le style est inconnu. <paramref name="stylesRoot"/> est la racine vivante de styles.xml (modifiée en place).
    /// </summary>
    public static string? Ensure(StyleSheet sheet, XElement stylesRoot, string builtInName)
    {
        var w = sheet.W;
        if (FindIn(stylesRoot, w, builtInName) is { } existing) return existing.Attribute(w + "styleId")?.Value;
        if (!BuiltIns.TryGetValue(builtInName, out var def)) return null;
        // parent d'abord (style de caractère : « Police par défaut » ; paragraphe : Normal, toujours présent)
        string? basedOn = def.BasedOn;
        if (def.Type == "character" && basedOn is not null) basedOn = Ensure(sheet, stylesRoot, "Default Paragraph Font");
        else if (basedOn == "Normal") basedOn = FindIn(stylesRoot, w, "Normal")?.Attribute(w + "styleId")?.Value ?? "Normal";
        string? next = def.Next == "Normal" ? FindIn(stylesRoot, w, "Normal")?.Attribute(w + "styleId")?.Value ?? "Normal" : def.Next;
        var id = UniqueId(stylesRoot, w, def.Id);
        string ns = $"xmlns:w=\"{w.NamespaceName}\"";
        var xml = $"<w:style {ns} w:type=\"{def.Type}\"{(builtInName == "Default Paragraph Font" ? " w:default=\"1\"" : "")} w:styleId=\"{id}\">" +
                  $"<w:name w:val=\"{builtInName}\"/>" + (basedOn is not null ? $"<w:basedOn w:val=\"{basedOn}\"/>" : "") +
                  (next is not null ? $"<w:next w:val=\"{next}\"/>" : "") + $"<w:uiPriority w:val=\"{def.Priority}\"/>" +
                  (builtInName == "Default Paragraph Font" || builtInName.StartsWith("annotation ", StringComparison.Ordinal) ? "<w:semiHidden/><w:unhideWhenUsed/>" : "<w:qFormat/>") +
                  (def.PPr.Length > 0 ? $"<w:pPr>{def.PPr}</w:pPr>" : "") + (def.RPr.Length > 0 ? $"<w:rPr>{def.RPr}</w:rPr>" : "") + "</w:style>";
        var style = XElement.Parse(xml);
        StripNamespaceDeclarations(style, w);
        stylesRoot.Add(style);
        return id;
    }

    static XElement? FindIn(XElement root, XNamespace w, string name) => root.Elements(w + "style").FirstOrDefault(s =>
        string.Equals(s.Element(w + "name")?.Attribute(w + "val")?.Value, name, StringComparison.OrdinalIgnoreCase));

    static string UniqueId(XElement root, XNamespace w, string wanted)
    {
        var ids = root.Elements(w + "style").Select(s => s.Attribute(w + "styleId")?.Value).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var id = wanted;
        for (int k = 2; ids.Contains(id); k++) id = wanted + k;
        return id;
    }

    /// <summary>Retire la déclaration xmlns:w ajoutée par l'analyse (la racine de styles.xml la porte déjà).</summary>
    static void StripNamespaceDeclarations(XElement e, XNamespace w)
    {
        foreach (var x in e.DescendantsAndSelf())
            x.Attributes().Where(a => a.IsNamespaceDeclaration && a.Value == w.NamespaceName).Remove();
    }

    // ------------------------------------------------------------------ nouveaux styles

    /// <summary>
    /// Nouveau style de paragraphe (ou de caractère) : nom affiché, basé sur un style, propriétés données (copiées).
    /// Renvoie son identifiant.
    /// </summary>
    public static string Create(StyleSheet sheet, XElement stylesRoot, string type, string name, string? basedOn, string? next, XElement? pPr, XElement? rPr)
    {
        var w = sheet.W;
        var id = UniqueId(stylesRoot, w, new string(name.Where(char.IsLetterOrDigit).ToArray()) is { Length: > 0 } clean ? clean : "Style");
        var style = new XElement(w + "style", new XAttribute(w + "type", type), new XAttribute(w + "customStyle", "1"), new XAttribute(w + "styleId", id),
            new XElement(w + "name", new XAttribute(w + "val", name)));
        if (basedOn is not null) style.Add(new XElement(w + "basedOn", new XAttribute(w + "val", basedOn)));
        if (next is not null && type == "paragraph") style.Add(new XElement(w + "next", new XAttribute(w + "val", next)));
        style.Add(new XElement(w + "qFormat"));
        if (type == "paragraph" && pPr is not null && pPr.HasElements) style.Add(new XElement(w + "pPr", pPr.Elements().Select(e => new XElement(e))));
        if (rPr is not null && rPr.HasElements) style.Add(new XElement(w + "rPr", rPr.Elements().Select(e => new XElement(e))));
        stylesRoot.Add(style);
        return id;
    }

    /// <summary>
    /// Propriétés de paragraphe et de caractère d'un style, prêtes à modifier en place : les éléments w:pPr / w:rPr du
    /// style (créés et placés dans l'ordre du schéma au besoin).
    /// </summary>
    public static (Paragraph Para, Run Run) Editable(StyleSheet sheet, XElement style)
    {
        var w = sheet.W;
        var ppr = style.Element(w + "pPr");
        if (ppr is null)
        {
            ppr = new XElement(w + "pPr");
            if (style.Element(w + "rPr") is { } r0) r0.AddBeforeSelf(ppr);
            else if (style.Elements().FirstOrDefault(e => e.Name.LocalName is "tblPr" or "trPr" or "tcPr" or "tblStylePr") is { } t0) t0.AddBeforeSelf(ppr);
            else style.Add(ppr);
        }
        var rpr = style.Element(w + "rPr");
        if (rpr is null)
        {
            rpr = new XElement(w + "rPr");
            if (style.Elements().FirstOrDefault(e => e.Name.LocalName is "tblPr" or "trPr" or "tcPr" or "tblStylePr") is { } t1) t1.AddBeforeSelf(rpr);
            else style.Add(rpr);
        }
        return (new Paragraph { Props = new PropertyBag("pPr", w, ppr) }, new Run { Props = new PropertyBag("rPr", w, rpr) });
    }

    /// <summary>Après une modification : retire les w:pPr / w:rPr restés vides.</summary>
    public static void Tidy(StyleSheet sheet, XElement style)
    {
        var w = sheet.W;
        foreach (var name in new[] { "pPr", "rPr" })
            if (style.Element(w + name) is { } e && !e.HasElements && !e.HasAttributes) e.Remove();
    }

    static readonly string[] StyleOrder =
    {
        "name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden", "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat",
        "locked", "personal", "personalCompose", "personalReply", "rsid", "pPr", "rPr", "tblPr", "trPr", "tcPr", "tblStylePr",
    };

    /// <summary>Pose (ou retire, <paramref name="value"/> null) une propriété à w:val d'un style, à sa place dans l'ordre du schéma.</summary>
    public static void SetStyleValue(XElement style, XNamespace w, string local, string? value)
    {
        style.Element(w + local)?.Remove();
        if (value is null) return;
        var el = new XElement(w + local, new XAttribute(w + "val", value));
        int rank = Array.IndexOf(StyleOrder, local);
        var after = style.Elements().FirstOrDefault(e => Array.IndexOf(StyleOrder, e.Name.LocalName) > rank);
        if (after is null) style.Add(el); else after.AddBeforeSelf(el);
    }

    static readonly HashSet<string> NotCapturedP = new() { "pStyle", "numPr", "sectPr", "rPr", "pPrChange", "framePr", "cnfStyle", "divId" };
    static readonly HashSet<string> NotCapturedR = new() { "rStyle", "ins", "del", "moveFrom", "moveTo", "rPrChange", "lang", "noProof" };

    /// <summary>Mise en forme directe d'un paragraphe et de son premier segment de texte, à reprendre dans un style.</summary>
    public static (XElement? PPr, XElement? RPr) Capture(Paragraph p)
    {
        XElement? ppr = p.Props.Element is { } pe ? new XElement(pe.Name, pe.Elements().Where(e => !NotCapturedP.Contains(e.Name.LocalName)).Select(e => new XElement(e))) : null;
        var run = p.Runs().FirstOrDefault(r => r.Content.Any(c => c is TextContent));
        XElement? rpr = run?.Props.Element is { } re ? new XElement(re.Name, re.Elements().Where(e => !NotCapturedR.Contains(e.Name.LocalName)).Select(e => new XElement(e))) : null;
        return (ppr is { HasElements: true } ? ppr : null, rpr is { HasElements: true } ? rpr : null);
    }

    /// <summary>Copie des propriétés dans celles d'un style (chaque élément remplace celui du même nom).</summary>
    public static void Merge(StyleSheet sheet, XElement style, XElement? pPr, XElement? rPr)
    {
        var (para, run) = Editable(sheet, style);
        foreach (var e in pPr?.Elements() ?? Enumerable.Empty<XElement>()) para.Props.Set(new XElement(e));
        foreach (var e in rPr?.Elements() ?? Enumerable.Empty<XElement>()) run.Props.Set(new XElement(e));
        Tidy(sheet, style);
    }

    /// <summary>Retire d'un paragraphe la mise en forme directe reprise par son style (après <see cref="Capture"/>).</summary>
    public static void Absorb(Paragraph p, XElement? pPr, XElement? rPr)
    {
        foreach (var e in pPr?.Elements() ?? Enumerable.Empty<XElement>()) p.Props.Remove(e.Name.LocalName);
        foreach (var r in p.Runs())
            foreach (var e in rPr?.Elements() ?? Enumerable.Empty<XElement>())
                if (r.Props.Get(e.Name.LocalName) is { } own && XNode.DeepEquals(own, e)) own.Remove();
    }
}
