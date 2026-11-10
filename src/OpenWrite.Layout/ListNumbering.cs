using System.Globalization;
using System.Text;
using System.Xml.Linq;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;

namespace OpenWrite.Layout;

/// <summary>
/// Numéros des listes (premier pas du lot 4) : compteurs par définition de liste (w:abstractNum) et par niveau, dans
/// l'ordre du document ; valeur de départ (w:start, w:startOverride du w:num), remise à zéro des niveaux inférieurs,
/// texte du niveau (w:lvlText, « %1.%2. ») dans le format de chaque niveau (décimal, lettres, chiffres romains).
/// </summary>
public static class ListNumbering
{
    public static Dictionary<Paragraph, string> Compute(Story body, StyleResolver styles)
    {
        var sheet = styles.Sheet;
        var W = sheet.W;
        var result = new Dictionary<Paragraph, string>(ReferenceEqualityComparer.Instance);
        if (sheet.Numbering is not { } numbering) return result;
        var counters = new Dictionary<string, int[]>();
        var started = new HashSet<string>();
        var pending = new Dictionary<string, int>();
        foreach (var p in body.Paragraphs())
        {
            var pf = styles.Paragraph(p);
            if (pf.NumId <= 0 || sheet.ListLevel(pf.NumId, pf.ListLevel) is not { } lvl) continue;
            var num = numbering.Elements(W + "num").FirstOrDefault(n => n.Attribute(W + "numId")?.Value == pf.NumId.ToString(CultureInfo.InvariantCulture));
            var key = num?.Element(W + "abstractNumId")?.Attribute(W + "val")?.Value ?? "n" + pf.NumId;
            if (!counters.TryGetValue(key, out var c)) counters[key] = c = new int[9];
            int level = Math.Clamp(pf.ListLevel, 0, 8);
            // premier paragraphe d'un w:num : ses w:startOverride relancent les niveaux concernés
            if (num is not null && started.Add("num" + pf.NumId))
                foreach (var o in num.Elements(W + "lvlOverride"))
                    if (Int(o.Element(W + "startOverride")) is { } so && int.TryParse(o.Attribute(W + "ilvl")?.Value, out var ol) && ol is >= 0 and < 9)
                    {
                        pending[$"{key}/{ol}"] = so;
                        started.Remove($"{key}/{ol}");
                    }
            // premier élément du niveau depuis sa remise à zéro : valeur de départ ; sinon, suivant
            if (started.Add($"{key}/{level}"))
                c[level] = pending.Remove($"{key}/{level}", out var so2) ? so2 : Int(lvl.Element(W + "start")) ?? 1;
            else c[level]++;
            // les niveaux inférieurs repartent de leur valeur de départ
            for (int k = level + 1; k < 9; k++) started.Remove($"{key}/{k}");
            var fmt = lvl.Element(W + "numFmt")?.Attribute(W + "val")?.Value ?? "decimal";
            if (fmt is "bullet" or "none") continue;
            var text = lvl.Element(W + "lvlText")?.Attribute(W + "val")?.Value ?? "";
            var sb = new StringBuilder();
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == '%' && i + 1 < text.Length && text[i + 1] is >= '1' and <= '9')
                {
                    int l = text[i + 1] - '1';
                    var f = sheet.ListLevel(pf.NumId, l)?.Element(W + "numFmt")?.Attribute(W + "val")?.Value ?? "decimal";
                    sb.Append(Format(Math.Max(1, c[l]), f));
                    i++;
                }
                else sb.Append(text[i]);
            }
            result[p] = sb.ToString();
        }
        return result;

        int? Int(XElement? e) => e?.Attribute(W + "val")?.Value is { } v && int.TryParse(v, out var n) ? n : null;
    }

    public static string Format(int n, string fmt) => fmt switch
    {
        "lowerLetter" => Letters(n).ToLowerInvariant(),
        "upperLetter" => Letters(n),
        "lowerRoman" => Roman(n).ToLowerInvariant(),
        "upperRoman" => Roman(n),
        "decimalZero" => n.ToString("00", CultureInfo.InvariantCulture),
        _ => n.ToString(CultureInfo.InvariantCulture),
    };

    // Word : a…z, puis aa…zz, aaa…
    static string Letters(int n)
    {
        if (n <= 0) return "";
        char ch = (char)('A' + (n - 1) % 26);
        return new string(ch, (n - 1) / 26 + 1);
    }

    static string Roman(int n)
    {
        if (n <= 0 || n >= 4000) return n.ToString(CultureInfo.InvariantCulture);
        var map = new (int, string)[] { (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I") };
        var sb = new StringBuilder();
        foreach (var (v, s) in map) while (n >= v) { sb.Append(s); n -= v; }
        return sb.ToString();
    }
}
