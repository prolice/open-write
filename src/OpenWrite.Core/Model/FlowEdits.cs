using System.Xml.Linq;

namespace OpenWrite.Core.Model;

/// <summary>
/// Premières opérations sur le modèle en flux (prototype du lot 1) : elles vérifient qu'une modification locale ne
/// touche que le paragraphe concerné et produit un XML que Word accepte. Les opérations inversibles et le journal
/// d'annulation viennent avec le modèle définitif.
/// </summary>
public static class FlowEdits
{
    public static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

    /// <summary>Insère du texte dans un segment de texte.</summary>
    public static void InsertText(TextContent t, int offset, string text) => t.Value = t.Value.Insert(offset, text);

    /// <summary>
    /// Scinde un paragraphe au caractère <paramref name="offset"/> du texte <paramref name="t"/> du segment <paramref name="run"/>
    /// (segment placé directement dans le paragraphe). Le second paragraphe reçoit une copie des propriétés et de nouveaux
    /// identifiants w14:paraId/w14:textId (uniques dans le document, comme Word l'exige) ; un saut de section éventuel reste
    /// sur le second paragraphe, comme dans Word. Renvoie le nouveau paragraphe, ou null si la position n'est pas prise en charge.
    /// </summary>
    public static Paragraph? SplitParagraph(Story story, Paragraph p, Run run, TextContent t, int offset, Func<string> newId)
    {
        int bi = story.Blocks.IndexOf(p), ri = p.Inlines.IndexOf(run), ci = run.Content.IndexOf(t);
        if (bi < 0 || ri < 0 || ci < 0 || offset < 0 || offset > t.Value.Length) return null;

        var p2 = new Paragraph { Props = p.Props.Clone() };
        foreach (var a in p.Attributes)
        {
            if (a.Name == W14 + "paraId" || a.Name == W14 + "textId") p2.Attributes.Add(new XAttribute(a.Name, newId()));
            else p2.Attributes.Add(new XAttribute(a));
        }
        p.Props.Remove("sectPr");            // le saut de section suit la fin du paragraphe d'origine

        // segment coupé en deux : même mise en forme
        var r2 = new Run { Props = run.Props.Clone() };
        r2.Attributes.AddRange(run.Attributes.Select(a => new XAttribute(a)));
        var tail = new TextContent { Value = t.Value[offset..], Deleted = t.Deleted };
        tail.Attributes.AddRange(t.Attributes.Select(a => new XAttribute(a)));
        t.Value = t.Value[..offset];
        if (tail.Value.Length > 0) r2.Content.Add(tail);
        r2.Content.AddRange(run.Content.Skip(ci + 1));
        run.Content.RemoveRange(ci + 1, run.Content.Count - ci - 1);
        if (t.Value.Length == 0) run.Content.Remove(t);

        if (r2.Content.Count > 0) p2.Inlines.Add(r2);
        p2.Inlines.AddRange(p.Inlines.Skip(ri + 1));
        p.Inlines.RemoveRange(ri + 1, p.Inlines.Count - ri - 1);
        if (run.Content.Count == 0) p.Inlines.Remove(run);

        story.Blocks.Insert(bi + 1, p2);
        return p2;
    }

    /// <summary>Générateur d'identifiants w14:paraId : 8 chiffres hexadécimaux, inférieurs à 0x80000000, absents du document.</summary>
    public static Func<string> IdGenerator(IEnumerable<string> existing, int seed = 1)
    {
        var used = new HashSet<string>(existing, StringComparer.OrdinalIgnoreCase);
        var rnd = new Random(seed);
        return () =>
        {
            string id;
            do id = rnd.Next(1, int.MaxValue).ToString("X8"); while (!used.Add(id));
            return id;
        };
    }
}
