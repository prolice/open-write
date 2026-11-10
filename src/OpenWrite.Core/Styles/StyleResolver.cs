using System.Globalization;
using System.Xml.Linq;
using OpenWrite.Core.Model;

namespace OpenWrite.Core.Styles;

/// <summary>Mise en forme de paragraphe résolue (longueurs en twips, comme WordprocessingML).</summary>
public sealed record ParagraphFormat
{
    public string? StyleId { get; init; }
    /// <summary>w:jc normalisé : left, center, right, both, distribute… (start → left, end → right).</summary>
    public string Alignment { get; init; } = "left";
    public int IndentLeft { get; init; }
    public int IndentRight { get; init; }
    /// <summary>Retrait de première ligne : positif (w:firstLine) ou négatif (w:hanging).</summary>
    public int IndentFirstLine { get; init; }
    public int SpacingBefore { get; init; }
    public int SpacingAfter { get; init; }
    public bool BeforeAutospacing { get; init; }
    public bool AfterAutospacing { get; init; }
    /// <summary>Interligne : w:line et w:lineRule (auto = multiple de 240, atLeast / exact = twips).</summary>
    public int Line { get; init; } = 240;
    public string LineRule { get; init; } = "auto";
    public bool KeepNext { get; init; }
    public bool KeepLines { get; init; }
    public bool PageBreakBefore { get; init; }
    public bool WidowControl { get; init; }
    public bool ContextualSpacing { get; init; }
    public int? OutlineLevel { get; init; }
    public int NumId { get; init; }
    public int ListLevel { get; init; }
    /// <summary>Propriétés fusionnées (w:pPr), pour tout ce qui n'a pas d'accès typé (bordures, trame, tabulations…).</summary>
    public required XElement Effective { get; init; }
}

/// <summary>Mise en forme de caractère résolue (tailles en points).</summary>
public sealed record RunFormat
{
    public string? StyleId { get; init; }
    public string? FontAscii { get; init; }
    public string? FontHAnsi { get; init; }
    public string? FontEastAsia { get; init; }
    public string? FontComplex { get; init; }
    public double Size { get; init; } = 10;
    public double SizeComplex { get; init; } = 10;
    public bool Bold { get; init; }
    public bool Italic { get; init; }
    /// <summary>Type de soulignement (single, double, words…) ; null si aucun.</summary>
    public string? Underline { get; init; }
    public bool Strike { get; init; }
    public bool DoubleStrike { get; init; }
    public bool Caps { get; init; }
    public bool SmallCaps { get; init; }
    public bool Hidden { get; init; }
    /// <summary>Couleur RRGGBB (thème résolu, nuance appliquée) ; null = automatique.</summary>
    public string? Color { get; init; }
    public string? Highlight { get; init; }
    /// <summary>superscript, subscript ou null.</summary>
    public string? VerticalAlign { get; init; }
    /// <summary>Espacement des caractères (w:spacing), en twips.</summary>
    public int CharacterSpacing { get; init; }
    /// <summary>Échelle horizontale des caractères (w:w), en pour cent.</summary>
    public int Scale { get; init; } = 100;
    public string? Language { get; init; }
    public required XElement Effective { get; init; }
}

/// <summary>
/// Résolution des propriétés héritées (PLAN.md §3.1), dans l'ordre que Word applique réellement (calibré par COM sur
/// le corpus, PLAN.md §10.5) : valeurs par défaut du document → style de tableau → style de paragraphe (chaîne basedOn)
/// → niveau de liste (la norme, ECMA-376 §17.7.2, le place avant le style de paragraphe ; Word non) → style de caractère
/// (chaîne basedOn) → mise en forme directe.
/// <list type="bullet">
/// <item>Fusion par propriété ; pour w:rFonts, w:spacing, w:ind, w:lang, par attribut (w:firstLine et w:hanging s'excluent,
///   une police explicite remplace la police de thème du même emplacement) ; bordures et w:numPr, par côté / par enfant ;
///   tabulations cumulées (w:val="clear" retire un taquet hérité).</item>
/// <item>Propriétés bascule (gras, italique, majuscules, barré, masqué…) : entre niveaux de styles elles s'inversent
///   (OU exclusif : gras dans le style de paragraphe ET dans le style de caractère = pas gras) ; la mise en forme directe impose sa valeur.</item>
/// <item>Polices et couleurs du thème résolues (w:clrSchemeMapping, w:themeTint / w:themeShade en luminance TSL).</item>
/// </list>
/// Styles de tableau (lot 6) : le niveau « tableau » comprend les mises en forme conditionnelles (ligne d'en-tête, bandes…)
/// des régions de la cellule (<see cref="CellStyle"/>, <see cref="TableLevel"/>).
/// Le résolveur met en cache les fusions de styles ; il est à recréer quand les styles changent (lot 4).
/// </summary>
public sealed class StyleResolver
{
    readonly StyleSheet _s;
    readonly XNamespace W;
    readonly Dictionary<string, (XElement? PPr, XElement? RPr)> _styleCache = new();

    public StyleResolver(StyleSheet sheet)
    {
        _s = sheet;
        W = sheet.W;
    }

    public StyleSheet Sheet => _s;

    static readonly HashSet<string> Toggles = new()
    {
        "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "vanish",
    };

    // ================================================================ paragraphe

    /// <summary>
    /// Propriétés de paragraphe résolues ; <paramref name="cell"/> = style du tableau qui contient le paragraphe et régions
    /// de sa cellule (mise en forme conditionnelle ; une simple chaîne = identifiant du style, sans régions).
    /// </summary>
    public ParagraphFormat Paragraph(Paragraph p, CellStyle cell = default)
    {
        var styleId = ParagraphStyleId(p);
        var direct = p.Props.Element;
        // numérotation : w:numId et w:ilvl fusionnés (style puis mise en forme directe) ; w:numId="0" retire la liste
        var styleNum = StyleProps(styleId).PPr?.Element(W + "numPr");
        var directNum = direct?.Element(W + "numPr");
        int? styleNumId = Int(styleNum?.Element(W + "numId"), "val"), directNumId = Int(directNum?.Element(W + "numId"), "val");
        int ilvl = Int(directNum?.Element(W + "ilvl"), "val") ?? Int(styleNum?.Element(W + "ilvl"), "val") ?? 0;
        int numId = directNumId ?? styleNumId ?? 0;
        // Calibration Word (PLAN.md §10.5), mesurée par COM sur le corpus :
        // - liste active : les propriétés du niveau de liste s'appliquent APRÈS le style de paragraphe (la norme les place
        //   avant). Annexe : « Paragraphe de liste » w:firstLine=709 + liste directe w:hanging=360 → Word -18 pt ;
        // - liste du style retirée par le paragraphe (w:numId="0") : ni le retrait de la liste NI celui du style ne
        //   s'appliquent. FrameView : style « Legal2_L1 » (w:firstLine=720, liste) → Word 0 ; Syllabus : style « coupe »
        //   (liste niveau 2, w:left=1788) → Word 0.
        var lvlPPr = numId > 0 ? _s.ListLevel(numId, ilvl)?.Element(W + "pPr") : null;
        bool dropStyleIndent = directNumId == 0 && styleNumId is > 0;
        var acc = MergeParagraph(styleId, cell, direct, lvlPPr, dropStyleIndent);
        return ToParagraphFormat(acc, styleId, numId, ilvl);
    }

    string? ParagraphStyleId(Paragraph p) =>
        p.Props.Val("pStyle") is { } id && _s.Style(id) is not null ? id : _s.DefaultStyle("paragraph")?.Attribute(W + "styleId")?.Value;

    XElement MergeParagraph(string? styleId, CellStyle cell, XElement? direct, XElement? numbering, bool dropStyleIndent = false)
    {
        var acc = new XElement(W + "pPr");
        Merge(acc, _s.DefaultParagraphProps);
        Merge(acc, TableLevel(cell).PPr);
        var style = StyleProps(styleId).PPr;
        if (dropStyleIndent && style is not null)
        {
            style = new XElement(style);
            style.Element(W + "ind")?.Remove();
        }
        Merge(acc, style);
        Merge(acc, numbering);
        if (direct is not null)
        {
            var d = new XElement(direct);
            d.Elements().Where(e => e.Name == W + "rPr" || e.Name == W + "sectPr" || e.Name == W + "pPrChange").Remove();
            Merge(acc, d);
        }
        return acc;
    }

    /// <summary>Style de tableau effectif : celui demandé s'il existe, sinon le style de tableau par défaut (« Tableau Normal »).</summary>
    public string? TableStyle(string? tableStyleId) =>
        tableStyleId is not null && _s.Style(tableStyleId) is not null ? tableStyleId
        : tableStyleId is null ? null : _s.DefaultStyle("table")?.Attribute(W + "styleId")?.Value;

    readonly Dictionary<(string, TableRegion), (XElement? PPr, XElement? RPr, XElement? TcPr)> _tableCache = new();

    /// <summary>
    /// Niveau « style de tableau » d'un paragraphe de cellule : pPr / rPr / tcPr du style (chaîne basedOn), puis ceux de
    /// ses mises en forme conditionnelles (w:tblStylePr) pour les régions de la cellule, dans l'ordre de
    /// <see cref="TableStyles.Order"/> (la suivante l'emporte ; à l'intérieur du niveau, pas d'OU exclusif).
    /// </summary>
    public (XElement? PPr, XElement? RPr, XElement? TcPr) TableLevel(CellStyle cell)
    {
        var id = TableStyle(cell.TableStyleId);
        if (id is null) return (null, null, null);
        if (_tableCache.TryGetValue((id, cell.Regions), out var cached)) return cached;
        var chain = _s.Chain(id);
        XElement? pPr = null, rPr = null, tcPr = null;
        void Layer(XElement? holder)
        {
            if (holder is null) return;
            if (holder.Element(W + "pPr") is { } sp) Merge(pPr ??= new XElement(W + "pPr"), sp);
            if (holder.Element(W + "rPr") is { } sr) MergeRun(rPr ??= new XElement(W + "rPr"), sr);
            if (holder.Element(W + "tcPr") is { } tc) Merge(tcPr ??= new XElement(W + "tcPr"), tc);
        }
        foreach (var st in chain)
        {
            Layer(st);
            Layer(st.Elements(W + "tblStylePr").FirstOrDefault(x => x.Attribute(W + "type")?.Value == "wholeTable"));
        }
        foreach (var (region, type) in TableStyles.Order)
            if ((cell.Regions & region) != 0)
                foreach (var st in chain)
                    Layer(st.Elements(W + "tblStylePr").FirstOrDefault(x => x.Attribute(W + "type")?.Value == type));
        return _tableCache[(id, cell.Regions)] = (pPr, rPr, tcPr);
    }

    /// <summary>
    /// Propriétés de tableau (w:tblPr) du style et de sa chaîne basedOn, sans celles du tableau lui-même : bordures,
    /// marges de cellule, retrait, alignement, taille des bandes.
    /// </summary>
    public XElement? TableStyleProps(string? tableStyleId)
    {
        var id = TableStyle(tableStyleId);
        if (id is null) return null;
        XElement? acc = null;
        foreach (var st in _s.Chain(id))
            if (st.Element(W + "tblPr") is { } tp) Merge(acc ??= new XElement(W + "tblPr"), tp);
        return acc;
    }

    ParagraphFormat ToParagraphFormat(XElement pPr, string? styleId, int numId, int ilvl)
    {
        var ind = pPr.Element(W + "ind");
        var sp = pPr.Element(W + "spacing");
        var jc = pPr.Element(W + "jc")?.Attribute(W + "val")?.Value ?? "left";
        return new ParagraphFormat
        {
            StyleId = styleId,
            Alignment = jc switch { "start" => "left", "end" => "right", _ => jc },
            IndentLeft = Int(ind, "start") ?? Int(ind, "left") ?? 0,
            IndentRight = Int(ind, "end") ?? Int(ind, "right") ?? 0,
            IndentFirstLine = Int(ind, "hanging") is { } h ? -h : Int(ind, "firstLine") ?? 0,
            SpacingBefore = Int(sp, "before") ?? 0,
            SpacingAfter = Int(sp, "after") ?? 0,
            BeforeAutospacing = OnOff(sp?.Attribute(W + "beforeAutospacing")),
            AfterAutospacing = OnOff(sp?.Attribute(W + "afterAutospacing")),
            Line = Int(sp, "line") ?? 240,
            LineRule = sp?.Attribute(W + "lineRule")?.Value ?? "auto",
            KeepNext = Flag(pPr, "keepNext"),
            KeepLines = Flag(pPr, "keepLines"),
            PageBreakBefore = Flag(pPr, "pageBreakBefore"),
            // absent : actif (Word ; ECMA-376 dit l'inverse) — vérifié par COM, ParagraphFormat.WidowControl vrai pour tous les
            // paragraphes sans w:widowControl (Annexe 286/286, tableau long : une rangée coupée à une ligne passe à la page suivante)
            WidowControl = pPr.Element(W + "widowControl") is not { } wc || IsOn(wc),
            ContextualSpacing = Flag(pPr, "contextualSpacing"),
            OutlineLevel = Int(pPr.Element(W + "outlineLvl"), "val") is { } ol and < 9 ? ol : null,
            NumId = numId,
            ListLevel = ilvl,
            Effective = pPr,
        };
    }

    // ================================================================ caractère

    /// <summary>Propriétés d'un segment du paragraphe.</summary>
    public RunFormat Run(Paragraph p, Run r, CellStyle cell = default) => ResolveRun(p, r.Props.Element, r.Props.Val("rStyle"), cell);

    /// <summary>Propriétés de la marque de paragraphe (pPr/rPr) : taille de la ligne d'un paragraphe vide.</summary>
    public RunFormat ParagraphMark(Paragraph p, CellStyle cell = default)
    {
        var mark = p.Props.Get("rPr");
        return ResolveRun(p, mark, mark?.Element(W + "rStyle")?.Attribute(W + "val")?.Value, cell);
    }

    /// <summary>Propriétés du numéro ou de la puce d'un paragraphe de liste (w:lvl/w:rPr par-dessus la marque de paragraphe).</summary>
    public RunFormat? ListMarker(Paragraph p, CellStyle cell = default)
    {
        var pf = Paragraph(p, cell);
        var lvl = _s.ListLevel(pf.NumId, pf.ListLevel);
        if (lvl is null) return null;
        var mark = ParagraphMark(p, cell).Effective;
        var acc = new XElement(mark);
        MergeRun(acc, lvl.Element(W + "rPr"));
        return ToRunFormat(acc, null);
    }

    RunFormat ResolveRun(Paragraph p, XElement? direct, string? charStyleId, CellStyle cell)
    {
        var pStyle = ParagraphStyleId(p);
        if (charStyleId is not null && _s.Style(charStyleId) is null) charStyleId = null;
        var tbl = TableLevel(cell).RPr;
        var para = StyleProps(pStyle).RPr;
        var chr = StyleProps(charStyleId).RPr;

        var acc = new XElement(W + "rPr");
        MergeRun(acc, _s.DefaultRunProps);
        MergeRun(acc, tbl);
        MergeRun(acc, para);
        MergeRun(acc, chr);
        // propriétés bascule : OU exclusif entre les niveaux de styles qui les définissent (sinon valeur par défaut)
        foreach (var t in Toggles)
        {
            var levels = new[] { tbl, para, chr }.Select(x => x?.Element(W + t)).Where(x => x is not null).ToList();
            if (levels.Count == 0) continue;
            bool v = levels.Aggregate(false, (cur, e) => cur ^ IsOn(e!));
            acc.Element(W + t)?.Remove();
            acc.Add(v ? new XElement(W + t) : new XElement(W + t, new XAttribute(W + "val", "0")));
        }
        if (direct is not null)
        {
            var d = new XElement(direct);
            d.Elements().Where(e => e.Name.LocalName is "ins" or "del" or "moveFrom" or "moveTo" or "rPrChange" or "rStyle").Remove();
            MergeRun(acc, d);
        }
        return ToRunFormat(acc, charStyleId);
    }

    RunFormat ToRunFormat(XElement rPr, string? charStyleId)
    {
        var fonts = rPr.Element(W + "rFonts");
        string? Font(string slot)
        {
            var theme = fonts?.Attribute(W + slot + "Theme")?.Value;
            if (theme is not null && _s.ThemeFont(theme) is { } tf) return tf;
            return fonts?.Attribute(W + slot)?.Value;
        }
        var u = rPr.Element(W + "u")?.Attribute(W + "val")?.Value;
        return new RunFormat
        {
            StyleId = charStyleId,
            FontAscii = Font("ascii"),
            FontHAnsi = Font("hAnsi"),
            FontEastAsia = Font("eastAsia"),
            FontComplex = Font("cs"),
            Size = (HalfPoints(rPr.Element(W + "sz")) ?? 20) / 2.0,
            SizeComplex = (HalfPoints(rPr.Element(W + "szCs")) ?? HalfPoints(rPr.Element(W + "sz")) ?? 20) / 2.0,
            Bold = Flag(rPr, "b"),
            Italic = Flag(rPr, "i"),
            Underline = rPr.Element(W + "u") is null || u == "none" ? null : u ?? "single",
            Strike = Flag(rPr, "strike"),
            DoubleStrike = Flag(rPr, "dstrike"),
            Caps = Flag(rPr, "caps"),
            SmallCaps = Flag(rPr, "smallCaps"),
            Hidden = Flag(rPr, "vanish"),
            Color = ColorOf(rPr.Element(W + "color")),
            Highlight = rPr.Element(W + "highlight")?.Attribute(W + "val")?.Value is { } hl and not "none" ? hl : null,
            VerticalAlign = rPr.Element(W + "vertAlign")?.Attribute(W + "val")?.Value is { } va and not "baseline" ? va : null,
            CharacterSpacing = Int(rPr.Element(W + "spacing"), "val") ?? 0,
            Scale = Int(rPr.Element(W + "w"), "val") is { } sc and > 0 and <= 600 ? sc : 100,
            Language = rPr.Element(W + "lang")?.Attribute(W + "val")?.Value,
            Effective = rPr,
        };
    }

    /// <summary>Couleur d'un élément w:color (thème, nuance) en RRGGBB ; null = automatique.</summary>
    public string? ResolveColor(XElement? color) => ColorOf(color);

    string? ColorOf(XElement? color)
    {
        if (color is null) return null;
        string? rgb = null;
        if (color.Attribute(W + "themeColor")?.Value is { } tc) rgb = _s.ThemeColor(tc);
        rgb ??= color.Attribute(W + "val")?.Value;
        if (rgb is null || rgb.Equals("auto", StringComparison.OrdinalIgnoreCase) || rgb.Length != 6) return null;
        rgb = rgb.ToUpperInvariant();
        if (color.Attribute(W + "themeColor") is not null)
        {
            if (Hex(color.Attribute(W + "themeTint")?.Value) is { } tint) rgb = Luminance(rgb, l => l * tint + (1 - tint));
            if (Hex(color.Attribute(W + "themeShade")?.Value) is { } shade) rgb = Luminance(rgb, l => l * shade);
        }
        return rgb;
    }

    static double? Hex(string? v) => v is not null && int.TryParse(v, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var n) ? n / 255.0 : null;

    /// <summary>Applique une fonction à la luminance (TSL) d'une couleur RRGGBB.</summary>
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
        return $"{(int)Math.Round(r * 255):X2}{(int)Math.Round(g * 255):X2}{(int)Math.Round(b * 255):X2}";
    }

    // ================================================================ styles

    /// <summary>pPr et rPr fusionnés le long de la chaîne basedOn d'un style (sans les valeurs par défaut du document).</summary>
    (XElement? PPr, XElement? RPr) StyleProps(string? id)
    {
        if (id is null) return (null, null);
        if (_styleCache.TryGetValue(id, out var cached)) return cached;
        XElement? pPr = null, rPr = null;
        foreach (var s in _s.Chain(id))
        {
            if (s.Element(W + "pPr") is { } sp) Merge(pPr ??= new XElement(W + "pPr"), sp);
            if (s.Element(W + "rPr") is { } sr) MergeRun(rPr ??= new XElement(W + "rPr"), sr);
        }
        return _styleCache[id] = (pPr, rPr);
    }

    // ================================================================ fusion

    static readonly HashSet<string> ByChild = new() { "pBdr", "numPr", "tblBorders", "tcBorders", "tblCellMar", "tcMar" };

    void MergeRun(XElement acc, XElement? layer) => Merge(acc, layer);

    /// <summary>Fusionne une couche de propriétés dans l'accumulateur (règles : voir la classe).</summary>
    void Merge(XElement acc, XElement? layer)
    {
        if (layer is null) return;
        foreach (var c in layer.Elements())
        {
            var existing = acc.Element(c.Name);
            if (c.Name.Namespace == W)
            {
                switch (c.Name.LocalName)
                {
                    case "tabs": MergeTabs(acc, c); continue;
                    case "rFonts" when existing is not null: MergeFonts(existing, c); continue;
                    case "ind" when existing is not null: MergeIndent(existing, c); continue;
                    case "spacing" when existing is not null && acc.Name.LocalName == "pPr": MergeSpacing(existing, c); continue;
                    case "lang" or "eastAsianLayout" when existing is not null:
                        foreach (var a in c.Attributes()) existing.SetAttributeValue(a.Name, a.Value);
                        continue;
                }
                if (existing is not null && ByChild.Contains(c.Name.LocalName))
                {
                    foreach (var a in c.Attributes()) existing.SetAttributeValue(a.Name, a.Value);
                    foreach (var child in c.Elements())
                    {
                        existing.Element(child.Name)?.Remove();
                        existing.Add(new XElement(child));
                    }
                    continue;
                }
            }
            existing?.Remove();
            acc.Add(new XElement(c));
        }
    }

    void MergeFonts(XElement acc, XElement layer)
    {
        foreach (var slot in new[] { "ascii", "hAnsi", "eastAsia", "cs" })
        {
            var v = layer.Attribute(W + slot);
            var t = layer.Attribute(W + slot + "Theme");
            if (v is null && t is null) continue;
            // une police explicite remplace la police de thème héritée du même emplacement, et inversement
            acc.Attribute(W + slot)?.Remove();
            acc.Attribute(W + slot + "Theme")?.Remove();
            if (v is not null) acc.SetAttributeValue(v.Name, v.Value);
            if (t is not null) acc.SetAttributeValue(t.Name, t.Value);
        }
        if (layer.Attribute(W + "hint") is { } hint) acc.SetAttributeValue(hint.Name, hint.Value);
    }

    void MergeIndent(XElement acc, XElement layer)
    {
        foreach (var a in layer.Attributes())
        {
            var local = a.Name.LocalName;
            // w:firstLine et w:hanging s'excluent ; left/start et right/end sont synonymes
            if (local is "firstLine" or "hanging") { acc.Attribute(W + "firstLine")?.Remove(); acc.Attribute(W + "hanging")?.Remove(); }
            if (local is "firstLineChars" or "hangingChars") { acc.Attribute(W + "firstLineChars")?.Remove(); acc.Attribute(W + "hangingChars")?.Remove(); }
            if (local is "left" or "start") { acc.Attribute(W + "left")?.Remove(); acc.Attribute(W + "start")?.Remove(); }
            if (local is "right" or "end") { acc.Attribute(W + "right")?.Remove(); acc.Attribute(W + "end")?.Remove(); }
            acc.SetAttributeValue(a.Name, a.Value);
        }
    }

    void MergeSpacing(XElement acc, XElement layer)
    {
        foreach (var a in layer.Attributes()) acc.SetAttributeValue(a.Name, a.Value);
        // w:line sans w:lineRule : règle « auto » (valeur par défaut du schéma), pas celle héritée
        if (layer.Attribute(W + "line") is not null && layer.Attribute(W + "lineRule") is null) acc.Attribute(W + "lineRule")?.Remove();
    }

    void MergeTabs(XElement acc, XElement layer)
    {
        var tabs = acc.Element(W + "tabs") ?? new XElement(W + "tabs");
        if (tabs.Parent is null) acc.Add(tabs);
        foreach (var tab in layer.Elements(W + "tab"))
        {
            var pos = tab.Attribute(W + "pos")?.Value;
            tabs.Elements(W + "tab").Where(t => t.Attribute(W + "pos")?.Value == pos).Remove();
            if (tab.Attribute(W + "val")?.Value != "clear") tabs.Add(new XElement(tab));
        }
        var sorted = tabs.Elements(W + "tab").OrderBy(t => Int(t, "pos") ?? 0).ToList();
        tabs.RemoveNodes();
        tabs.Add(sorted);
    }

    // ================================================================ lecture

    /// <summary>Longueur en twips (ou nombre entier). Le format strict ISO 29500 admet les mesures avec unité :
    /// « 8pt », « 12.95pt », « 2.5cm »… (calibration corpus : strict-format.docx, w:after="8pt" → Word 8 pt).</summary>
    int? Int(XElement? e, string attr) => Measure(e?.Attribute(W + attr)?.Value, 20);

    /// <summary>Taille en demi-points (w:sz) : entier, ou mesure avec unité en format strict.</summary>
    int? HalfPoints(XElement? e) => Measure(e?.Attribute(W + "val")?.Value, 2);

    /// <summary>Valeur entière telle quelle, ou mesure avec unité convertie : points × <paramref name="perPoint"/>.</summary>
    static int? Measure(string? v, double perPoint)
    {
        if (v is null) return null;
        if (int.TryParse(v, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n)) return n;
        double? pts = null;
        foreach (var (unit, factor) in Units)
            if (v.EndsWith(unit, StringComparison.Ordinal) && double.TryParse(v[..^unit.Length], NumberStyles.Float, CultureInfo.InvariantCulture, out var d))
            { pts = d * factor; break; }
        return pts is { } p ? (int)Math.Round(p * perPoint) : null;
    }

    static readonly (string Unit, double PointsPer)[] Units = { ("pt", 1), ("in", 72), ("cm", 72 / 2.54), ("mm", 72 / 25.4), ("pc", 12), ("pi", 12) };

    bool Flag(XElement parent, string local) => parent.Element(W + local) is { } e && IsOn(e);

    bool IsOn(XElement e) => OnOff(e.Attribute(W + "val")) || e.Attribute(W + "val") is null;

    static bool OnOff(XAttribute? a) => a?.Value is "1" or "true" or "on";
}
