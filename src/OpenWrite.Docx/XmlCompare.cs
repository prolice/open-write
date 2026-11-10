using System.Text;
using System.Xml.Linq;

namespace OpenWrite.Docx;

/// <summary>
/// Comparaison XML normalisée (banc d'aller-retour, PLAN.md §5.1) : ordre des attributs, déclarations et préfixes
/// d'espaces de noms et blancs entre éléments ne comptent pas ; tout le reste compte (noms, valeurs, ordre des
/// éléments, texte). Les rsid et w14:paraId comptent aussi : OpenWrite doit les conserver.
/// </summary>
public static class XmlCompare
{
    /// <summary>Différences entre deux éléments (au plus <paramref name="max"/>), chacune avec son chemin.</summary>
    public static List<string> Diff(XElement a, XElement b, int max = 20)
    {
        var res = new List<string>();
        Walk(a, b, "/" + Short(a.Name), res, max);
        return res;
    }

    static void Walk(XElement a, XElement b, string path, List<string> res, int max)
    {
        if (res.Count >= max) return;
        if (a.Name != b.Name) { res.Add($"{path} : élément {Short(a.Name)} ≠ {Short(b.Name)}"); return; }
        var aa = Attrs(a);
        var ba = Attrs(b);
        foreach (var (k, v) in aa)
            if (!ba.TryGetValue(k, out var v2)) res.Add($"{path} : attribut {k} perdu");
            else if (v != v2) res.Add($"{path} : attribut {k} « {v} » ≠ « {v2} »");
        foreach (var k in ba.Keys.Where(k => !aa.ContainsKey(k))) res.Add($"{path} : attribut {k} ajouté");
        var an = Children(a);
        var bn = Children(b);
        int n = Math.Min(an.Count, bn.Count);
        for (int i = 0; i < n && res.Count < max; i++)
        {
            switch (an[i], bn[i])
            {
                case (XElement x, XElement y): Walk(x, y, $"{path}/{Short(x.Name)}[{i}]", res, max); break;
                case (XText x, XText y) when x.Value != y.Value: res.Add($"{path} : texte « {Cut(x.Value)} » ≠ « {Cut(y.Value)} »"); break;
                case (XText, XText): break;
                default: res.Add($"{path}[{i}] : nature du nœud différente"); break;
            }
        }
        if (an.Count != bn.Count && res.Count < max)
            res.Add($"{path} : {an.Count} enfant(s) ≠ {bn.Count}" + (an.Count > n ? $", perdu : {Describe(an[n])}" : $", ajouté : {Describe(bn[n])}"));
    }

    static Dictionary<string, string> Attrs(XElement e) =>
        e.Attributes().Where(x => !x.IsNamespaceDeclaration).ToDictionary(x => Short(x.Name), x => x.Value);

    /// <summary>Nœuds significatifs : éléments et texte ; les blancs entre éléments et les commentaires XML sont ignorés.</summary>
    static List<XNode> Children(XElement e)
    {
        bool mixed = e.HasElements;
        return e.Nodes().Where(n => n is XElement || (n is XText t && !(mixed && string.IsNullOrWhiteSpace(t.Value)))).ToList();
    }

    static string Describe(XNode n) => n is XElement e ? Short(e.Name) : $"texte « {Cut(((XText)n).Value)} »";

    static string Cut(string s) => s.Length <= 40 ? s : s[..40] + "…";

    static readonly Dictionary<string, string> Prefixes = new()
    {
        ["http://schemas.openxmlformats.org/wordprocessingml/2006/main"] = "w",
        ["http://purl.oclc.org/ooxml/wordprocessingml/main"] = "w",
        ["http://schemas.microsoft.com/office/word/2010/wordml"] = "w14",
        ["http://schemas.microsoft.com/office/word/2012/wordml"] = "w15",
        ["http://schemas.openxmlformats.org/markup-compatibility/2006"] = "mc",
        ["http://schemas.openxmlformats.org/officeDocument/2006/relationships"] = "r",
        ["http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"] = "wp",
        ["http://schemas.openxmlformats.org/drawingml/2006/main"] = "a",
        ["http://schemas.openxmlformats.org/officeDocument/2006/math"] = "m",
        ["urn:schemas-microsoft-com:vml"] = "v",
        ["urn:schemas-microsoft-com:office:office"] = "o",
        ["http://schemas.microsoft.com/office/word/2010/wordprocessingShape"] = "wps",
        ["http://www.w3.org/XML/1998/namespace"] = "xml",
    };

    /// <summary>Nom court lisible : préfixe usuel + nom local.</summary>
    public static string Short(XName n) =>
        n.Namespace == XNamespace.None ? n.LocalName : (Prefixes.TryGetValue(n.NamespaceName, out var p) ? p : "{" + n.NamespaceName + "}") + ":" + n.LocalName;

    /// <summary>Forme canonique d'un élément (pour comparer des listes de blocs).</summary>
    public static string Canonical(XElement e)
    {
        var sb = new StringBuilder();
        Canon(e, sb);
        return sb.ToString();
    }

    static void Canon(XElement e, StringBuilder sb)
    {
        sb.Append('<').Append(Short(e.Name));
        foreach (var (k, v) in Attrs(e).OrderBy(x => x.Key, StringComparer.Ordinal)) sb.Append(' ').Append(k).Append("=\"").Append(v).Append('"');
        sb.Append('>');
        foreach (var n in Children(e))
            if (n is XElement x) Canon(x, sb); else sb.Append(((XText)n).Value);
        sb.Append("</>");
    }

    /// <summary>
    /// Étendue des différences entre deux suites d'éléments (blocs du corps) : nombre d'éléments différents de part
    /// et d'autre une fois retirés le début et la fin communs. Une modification locale doit donner (1, 1).
    /// </summary>
    public static (int Removed, int Added) ChangedSpan(IReadOnlyList<XElement> before, IReadOnlyList<XElement> after)
    {
        var a = before.Select(Canonical).ToList();
        var b = after.Select(Canonical).ToList();
        int pre = 0;
        while (pre < a.Count && pre < b.Count && a[pre] == b[pre]) pre++;
        int suf = 0;
        while (suf < a.Count - pre && suf < b.Count - pre && a[^(suf + 1)] == b[^(suf + 1)]) suf++;
        return (a.Count - pre - suf, b.Count - pre - suf);
    }
}
