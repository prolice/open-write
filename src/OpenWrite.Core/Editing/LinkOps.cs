using System.Xml.Linq;
using OpenWrite.Core.Model;

namespace OpenWrite.Core.Editing;

/// <summary>
/// Liens hypertextes (w:hyperlink, lot 3) : trouver le lien d'une position, envelopper une plage d'un paragraphe dans un
/// nouveau lien, retirer un lien en gardant son texte. L'adresse externe est une relation de la partie (r:id), un lien
/// interne vise un signet (w:anchor).
/// </summary>
public static class LinkOps
{
    public static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    static int Length(Inline i) => i switch
    {
        Run r => r.Content.Sum(ParagraphAtoms.LengthOf),
        InlineContainer b => b.Children.Sum(Length),
        InlineContentControl s => s.Content.Sum(Length),
        _ => ParagraphAtoms.LengthOf(i),
    };

    /// <summary>Lien hypertexte (le plus intérieur) qui contient le caractère à la position <paramref name="offset"/>, et sa plage.</summary>
    public static (InlineContainer Link, int Start, int End)? LinkAt(Paragraph p, int offset)
    {
        (InlineContainer, int, int)? found = null;
        void Walk(List<Inline> list, int pos)
        {
            foreach (var i in list)
            {
                int len = Length(i);
                if (offset >= pos && offset < pos + len)
                {
                    if (i is InlineContainer box)
                    {
                        if (box.Name.LocalName == "hyperlink") found = (box, pos, pos + len);
                        Walk(box.Children, pos);
                    }
                    else if (i is InlineContentControl s) Walk(s.Content, pos);
                    return;
                }
                pos += len;
            }
        }
        Walk(p.Inlines, 0);
        return found;
    }

    /// <summary>Nouveau lien : adresse externe (r:id d'une relation) ou signet (w:anchor).</summary>
    public static InlineContainer NewLink(XNamespace w, string? relId, string? anchor)
    {
        var link = new InlineContainer(w + "hyperlink");
        if (relId is not null) link.Attributes.Add(new XAttribute(R + "id", relId));
        if (anchor is not null) link.Attributes.Add(new XAttribute(w + "anchor", anchor));
        link.Attributes.Add(new XAttribute(w + "history", "1"));
        return link;
    }

    /// <summary>
    /// Enveloppe la plage [<paramref name="a"/>, <paramref name="b"/>[ du paragraphe dans <paramref name="link"/> : les segments
    /// sont coupés aux bornes ; la plage doit être faite d'éléments de premier niveau du paragraphe (pas à cheval sur un
    /// autre lien ou une révision).
    /// </summary>
    public static void Wrap(Paragraph p, int a, int b, InlineContainer link)
    {
        if (b <= a) throw new ArgumentException("Plage vide.");
        FormatOps.SplitRunAt(p, b);
        FormatOps.SplitRunAt(p, a);
        int pos = 0, first = -1, last = -1;
        for (int k = 0; k < p.Inlines.Count; k++)
        {
            int len = Length(p.Inlines[k]);
            int s = pos, e = pos + len;
            pos = e;
            if (len == 0) { if (s > a && s < b) last = k; continue; }
            if (e <= a || s >= b) continue;
            if (s < a || e > b) throw new NotSupportedException("Le texte choisi est à cheval sur un autre lien ou une révision.");
            if (first < 0) first = k;
            last = k;
        }
        if (first < 0) throw new ArgumentException("Plage sans texte.");
        var moved = p.Inlines.GetRange(first, last - first + 1);
        p.Inlines.RemoveRange(first, last - first + 1);
        link.Children.AddRange(moved);
        p.Inlines.Insert(first, link);
    }

    /// <summary>Retire un lien et garde son texte à sa place.</summary>
    public static bool Unwrap(Paragraph p, InlineContainer link) => Unwrap(p.Inlines, link);

    static bool Unwrap(List<Inline> list, InlineContainer link)
    {
        int k = list.IndexOf(link);
        if (k >= 0)
        {
            list.RemoveAt(k);
            list.InsertRange(k, link.Children);
            return true;
        }
        foreach (var i in list)
            if (i is InlineContainer box && Unwrap(box.Children, link) || i is InlineContentControl s && Unwrap(s.Content, link)) return true;
        return false;
    }

    /// <summary>Change la cible d'un lien existant.</summary>
    public static void Retarget(InlineContainer link, XNamespace w, string? relId, string? anchor)
    {
        link.Attributes.RemoveAll(x => x.Name == R + "id" || x.Name == w + "anchor");
        if (relId is not null) link.Attributes.Insert(0, new XAttribute(R + "id", relId));
        if (anchor is not null) link.Attributes.Insert(0, new XAttribute(w + "anchor", anchor));
    }
}
