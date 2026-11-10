using System.Xml.Linq;
using OpenWrite.Core.Model;

namespace OpenWrite.Core.Styles;

public enum ListKind { Bullet, Numbered, Headings }

/// <summary>
/// Listes (lot 4) : définitions de numérotation (w:abstractNum + w:num de numbering.xml), écrites comme Word les écrit
/// pour ses puces et numéros par défaut ; application aux paragraphes (w:numPr), niveaux, nouvelle numérotation.
/// <paramref name="root"/> désigne toujours la racine vivante de numbering.xml (modifiée en place).
/// </summary>
public static class ListOps
{
    /// <summary>Puces de Word par niveau : caractère et police (Symbol, Courier New, Wingdings).</summary>
    static readonly (string Text, string Font)[] Bullets = { ("", "Symbol"), ("o", "Courier New"), ("", "Wingdings") };
    static readonly (string Fmt, string Jc)[] Numbers = { ("decimal", "left"), ("lowerLetter", "left"), ("lowerRoman", "right") };

    static int Max(XElement root, XNamespace w, string element, string attr) =>
        root.Elements(w + element).Select(e => int.TryParse(e.Attribute(w + attr)?.Value, out var v) ? v : -1).DefaultIfEmpty(-1).Max();

    /// <summary>Nouvelle liste (définition et instance) ; renvoie son numId. <paramref name="headingStyles"/> : styles Titre 1…3 liés (liste de titres).</summary>
    public static int Create(XElement root, XNamespace w, ListKind kind, IReadOnlyList<string>? headingStyles = null)
    {
        int abstractId = Max(root, w, "abstractNum", "abstractNumId") + 1;
        int numId = Math.Max(Max(root, w, "num", "numId") + 1, 1);
        var abs = new XElement(w + "abstractNum", new XAttribute(w + "abstractNumId", abstractId),
            new XElement(w + "multiLevelType", new XAttribute(w + "val", kind == ListKind.Headings ? "multilevel" : "hybridMultilevel")));
        for (int l = 0; l < 9; l++)
        {
            var lvl = new XElement(w + "lvl", new XAttribute(w + "ilvl", l), new XElement(w + "start", new XAttribute(w + "val", 1)));
            switch (kind)
            {
                case ListKind.Bullet:
                    var (text, font) = Bullets[l % 3];
                    lvl.Add(new XElement(w + "numFmt", new XAttribute(w + "val", "bullet")),
                        new XElement(w + "lvlText", new XAttribute(w + "val", text)), new XElement(w + "lvlJc", new XAttribute(w + "val", "left")),
                        new XElement(w + "pPr", new XElement(w + "ind", new XAttribute(w + "left", 720 * (l + 1)), new XAttribute(w + "hanging", 360))),
                        new XElement(w + "rPr", new XElement(w + "rFonts", new XAttribute(w + "ascii", font), new XAttribute(w + "hAnsi", font), new XAttribute(w + "hint", "default"))));
                    break;
                case ListKind.Numbered:
                    var (fmt, jc) = Numbers[l % 3];
                    lvl.Add(new XElement(w + "numFmt", new XAttribute(w + "val", fmt)),
                        new XElement(w + "lvlText", new XAttribute(w + "val", $"%{l + 1}.")), new XElement(w + "lvlJc", new XAttribute(w + "val", jc)),
                        new XElement(w + "pPr", new XElement(w + "ind", new XAttribute(w + "left", 720 * (l + 1)), new XAttribute(w + "hanging", jc == "right" ? 180 : 360))));
                    break;
                default:
                    // 1. / 1.1. / 1.1.1. … retraits de Word pour sa liste « Titres »
                    string lvlText = string.Concat(Enumerable.Range(1, l + 1).Select(k => $"%{k}.")).TrimEnd('.') + (l == 0 ? "." : "");
                    if (headingStyles is not null && l < headingStyles.Count) lvl.Add(new XElement(w + "pStyle", new XAttribute(w + "val", headingStyles[l])));
                    lvl.Add(new XElement(w + "numFmt", new XAttribute(w + "val", "decimal")),
                        new XElement(w + "lvlText", new XAttribute(w + "val", lvlText)), new XElement(w + "lvlJc", new XAttribute(w + "val", "left")),
                        new XElement(w + "pPr", new XElement(w + "ind", new XAttribute(w + "left", 432 + 144 * l), new XAttribute(w + "hanging", 432 + 144 * l))));
                    // l'élément pStyle se place après lvlRestart / start (ordre du schéma : start, numFmt, lvlRestart, pStyle…)
                    if (lvl.Element(w + "pStyle") is { } ps) { ps.Remove(); lvl.Element(w + "numFmt")!.AddAfterSelf(ps); }
                    break;
            }
            abs.Add(lvl);
        }
        // ordre du schéma : numPicBullet*, abstractNum*, num*, numIdMacAtCleanup
        var lastAbstract = root.Elements(w + "abstractNum").LastOrDefault() ?? root.Elements(w + "numPicBullet").LastOrDefault();
        if (lastAbstract is not null) lastAbstract.AddAfterSelf(abs);
        else root.AddFirst(abs);
        AddNum(root, w, numId, abstractId);
        return numId;
    }

    static XElement AddNum(XElement root, XNamespace w, int numId, int abstractId)
    {
        var num = new XElement(w + "num", new XAttribute(w + "numId", numId), new XElement(w + "abstractNumId", new XAttribute(w + "val", abstractId)));
        if (root.Element(w + "numIdMacAtCleanup") is { } mac) mac.AddBeforeSelf(num); else root.Add(num);
        return num;
    }

    /// <summary>
    /// Nouvelle instance de la même liste qui recommence à sa valeur de départ au niveau <paramref name="ilvl"/> (« Recommencer à 1 »).
    /// </summary>
    public static int Restart(XElement root, XNamespace w, int numId, int ilvl)
    {
        var num = root.Elements(w + "num").First(n => n.Attribute(w + "numId")?.Value == numId.ToString());
        int abstractId = int.Parse(num.Element(w + "abstractNumId")!.Attribute(w + "val")!.Value);
        var abs = root.Elements(w + "abstractNum").FirstOrDefault(a => a.Attribute(w + "abstractNumId")?.Value == abstractId.ToString());
        int start = int.TryParse(abs?.Elements(w + "lvl").FirstOrDefault(l => l.Attribute(w + "ilvl")?.Value == ilvl.ToString())
                                   ?.Element(w + "start")?.Attribute(w + "val")?.Value, out var st) ? st : 1;
        int newId = Max(root, w, "num", "numId") + 1;
        var created = AddNum(root, w, newId, abstractId);
        created.Add(new XElement(w + "lvlOverride", new XAttribute(w + "ilvl", ilvl), new XElement(w + "startOverride", new XAttribute(w + "val", start))));
        return newId;
    }

    /// <summary>Définition (w:abstractNum) d'une liste, pour comparer deux listes (« Continuer la numérotation »).</summary>
    public static int? AbstractOf(XElement root, XNamespace w, int numId) =>
        root.Elements(w + "num").FirstOrDefault(n => n.Attribute(w + "numId")?.Value == numId.ToString())?.Element(w + "abstractNumId")?.Attribute(w + "val")?.Value
            is { } v && int.TryParse(v, out var a) ? a : null;

    /// <summary>Liste à puces (format du premier niveau) ?</summary>
    public static bool IsBullet(StyleSheet sheet, int numId) =>
        sheet.ListLevel(numId, 0)?.Element(sheet.W + "numFmt")?.Attribute(sheet.W + "val")?.Value == "bullet";

    /// <summary>Met un paragraphe dans une liste (w:numPr : niveau, liste).</summary>
    public static void Set(Paragraph p, int numId, int ilvl)
    {
        var w = p.Props.Ns;
        p.Props.Set(new XElement(w + "numPr", new XElement(w + "ilvl", new XAttribute(w + "val", ilvl)), new XElement(w + "numId", new XAttribute(w + "val", numId))));
    }

    /// <summary>Sort un paragraphe de sa liste ; si son style en donne une, numId 0 l'annule (comme Word).</summary>
    public static void Remove(Paragraph p, bool styleHasList)
    {
        var w = p.Props.Ns;
        p.Props.Remove("numPr");
        if (styleHasList) p.Props.Set(new XElement(w + "numPr", new XElement(w + "numId", new XAttribute(w + "val", 0))));
    }

    /// <summary>Lie un style de paragraphe à un niveau de liste (liste de titres : Titre 1 → niveau 0…).</summary>
    public static void LinkStyle(StyleSheet sheet, XElement style, int numId, int ilvl)
    {
        var w = sheet.W;
        var (para, _) = StyleOps.Editable(sheet, style);
        var numPr = new XElement(w + "numPr");
        if (ilvl > 0) numPr.Add(new XElement(w + "ilvl", new XAttribute(w + "val", ilvl)));
        numPr.Add(new XElement(w + "numId", new XAttribute(w + "val", numId)));
        para.Props.Set(numPr);
        StyleOps.Tidy(sheet, style);
    }
}
