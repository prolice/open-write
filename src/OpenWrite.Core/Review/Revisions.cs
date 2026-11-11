using System.Globalization;
using System.Xml.Linq;
using OpenWrite.Core.Model;

namespace OpenWrite.Core.Review;

/// <summary>Sorte de révision (suivi des modifications de Word).</summary>
public enum RevisionKind
{
    /// <summary>Texte inséré (w:ins).</summary>
    Inserted,
    /// <summary>Texte supprimé (w:del).</summary>
    Deleted,
    /// <summary>Texte déplacé, à son ancienne place (w:moveFrom) : traité comme une suppression.</summary>
    MovedFrom,
    /// <summary>Texte déplacé, à sa nouvelle place (w:moveTo) : traité comme une insertion.</summary>
    MovedTo,
    /// <summary>Marque de paragraphe insérée (w:pPr/w:rPr/w:ins) : refusée, le paragraphe rejoint le suivant.</summary>
    ParagraphInserted,
    /// <summary>Marque de paragraphe supprimée (w:pPr/w:rPr/w:del) : acceptée, le paragraphe rejoint le suivant.</summary>
    ParagraphDeleted,
    /// <summary>Mise en forme de caractère modifiée (w:rPrChange).</summary>
    Formatted,
    /// <summary>Mise en forme de paragraphe modifiée (w:pPrChange).</summary>
    ParagraphFormatted,
    /// <summary>Ligne de tableau insérée (w:trPr/w:ins).</summary>
    RowInserted,
    /// <summary>Ligne de tableau supprimée (w:trPr/w:del).</summary>
    RowDeleted,
    /// <summary>Autre propriété modifiée (tableau, ligne, cellule, section…).</summary>
    PropertyChanged,
}

/// <summary>
/// Une révision trouvée dans le modèle. <see cref="Node"/> est l'objet vivant concerné (conteneur w:ins/w:del, élément
/// w:rPrChange…, ligne de tableau) ; <see cref="Paragraph"/> et la plage [<see cref="Start"/>, <see cref="End"/>[ (décalages
/// du modèle) situent la révision dans le texte pour la sélectionner.
/// </summary>
public sealed record Revision(RevisionKind Kind, string? Author, DateTime? Date, object Node, Paragraph? Paragraph, int Start, int End)
{
    public string Label => Kind switch
    {
        RevisionKind.Inserted => "Insertion",
        RevisionKind.Deleted => "Suppression",
        RevisionKind.MovedFrom => "Déplacé (origine)",
        RevisionKind.MovedTo => "Déplacé (destination)",
        RevisionKind.ParagraphInserted => "Paragraphe inséré",
        RevisionKind.ParagraphDeleted => "Paragraphes fusionnés",
        RevisionKind.Formatted => "Mise en forme",
        RevisionKind.ParagraphFormatted => "Mise en forme du paragraphe",
        RevisionKind.RowInserted => "Ligne insérée",
        RevisionKind.RowDeleted => "Ligne supprimée",
        _ => "Propriétés modifiées",
    };
}

/// <summary>
/// Révisions d'une suite de blocs : liste dans l'ordre du document, accepter ou refuser une révision ou toutes
/// (éventuellement d'un seul auteur), comme Révision ▸ Accepter / Refuser de Word. À exécuter dans une transaction du
/// journal pour être annulable.
/// </summary>
public static class Revisions
{
    static readonly HashSet<string> InlineKinds = new() { "ins", "del", "moveFrom", "moveTo" };
    static readonly HashSet<string> MoveMarkers = new() { "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd" };
    static readonly HashSet<string> PropertyChanges = new() { "tblPrChange", "trPrChange", "tcPrChange", "sectPrChange", "tblGridChange", "tblPrExChange", "cellIns", "cellDel", "cellMerge", "numberingChange" };

    static string? Attr(IEnumerable<XAttribute> attrs, string local) => attrs.FirstOrDefault(a => a.Name.LocalName == local)?.Value;

    static DateTime? DateOf(string? s) =>
        s is not null && DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out var d) ? d : null;

    static Revision Of(RevisionKind kind, IEnumerable<XAttribute> attrs, object node, Paragraph? p, int start, int end) =>
        new(kind, Attr(attrs, "author"), DateOf(Attr(attrs, "date")), node, p, start, end);

    // ================================================================ liste

    /// <summary>Toutes les révisions de la suite de blocs (tableaux et contrôles de contenu compris), dans l'ordre du document.</summary>
    public static List<Revision> List(Story root)
    {
        var res = new List<Revision>();
        Blocks(root, res);
        if (root.FinalSection?.Element is { } fs)
            foreach (var e in fs.Elements().Where(x => PropertyChanges.Contains(x.Name.LocalName))) res.Add(Of(RevisionKind.PropertyChanged, e.Attributes(), e, null, 0, 0));
        return res;
    }

    static void Blocks(Story s, List<Revision> res)
    {
        foreach (var b in s.Blocks)
        {
            switch (b)
            {
                case Paragraph p: Paragraph(p, res); break;
                case Table t:
                    foreach (var e in Changes(t.Props.Element)) res.Add(Of(RevisionKind.PropertyChanged, e.Attributes(), e, null, 0, 0));
                    if (t.Grid?.Elements().FirstOrDefault(x => x.Name.LocalName == "tblGridChange") is { } g) res.Add(Of(RevisionKind.PropertyChanged, g.Attributes(), g, null, 0, 0));
                    foreach (var row in t.Items.OfType<TableRow>())
                    {
                        var w = row.Props.Ns;
                        if (row.Props.Get("ins") is { } ri) res.Add(Of(RevisionKind.RowInserted, ri.Attributes(), row, FirstParagraph(row), 0, 0));
                        if (row.Props.Get("del") is { } rd) res.Add(Of(RevisionKind.RowDeleted, rd.Attributes(), row, FirstParagraph(row), 0, 0));
                        foreach (var e in Changes(row.Props.Element).Concat(Changes(row.PropsException?.Element)))
                            res.Add(Of(RevisionKind.PropertyChanged, e.Attributes(), e, FirstParagraph(row), 0, 0));
                        foreach (var cell in row.Items.OfType<TableCell>())
                        {
                            foreach (var e in Changes(cell.Props.Element)) res.Add(Of(RevisionKind.PropertyChanged, e.Attributes(), e, cell.Content.Paragraphs().FirstOrDefault(), 0, 0));
                            Blocks(cell.Content, res);
                        }
                    }
                    break;
                case BlockContentControl c: Blocks(c.Content, res); break;
            }
        }
    }

    static Paragraph? FirstParagraph(TableRow row) => row.Items.OfType<TableCell>().SelectMany(c => c.Content.Paragraphs()).FirstOrDefault();

    static IEnumerable<XElement> Changes(XElement? props) => props?.Elements().Where(x => PropertyChanges.Contains(x.Name.LocalName)) ?? Enumerable.Empty<XElement>();

    static void Paragraph(Paragraph p, List<Revision> res)
    {
        var w = p.Props.Ns;
        int length = ParagraphAtoms.Length(p);
        if (p.Props.Get("pPrChange") is { } pc) res.Add(Of(RevisionKind.ParagraphFormatted, pc.Attributes(), pc, p, 0, length));
        Inlines(p, p.Inlines, 0, res);
        var mark = p.Props.Get("rPr");
        if (mark?.Element(w + "ins") is { } mi) res.Add(Of(RevisionKind.ParagraphInserted, mi.Attributes(), mi, p, length, length));
        if (mark?.Element(w + "del") is { } md) res.Add(Of(RevisionKind.ParagraphDeleted, md.Attributes(), md, p, length, length));
        if (mark?.Element(w + "rPrChange") is { } mc) res.Add(Of(RevisionKind.Formatted, mc.Attributes(), mc, p, length, length));
    }

    static int Inlines(Paragraph p, List<Inline> list, int pos, List<Revision> res)
    {
        foreach (var i in list)
        {
            switch (i)
            {
                case Run r:
                {
                    int len = r.Content.Sum(ParagraphAtoms.LengthOf);
                    if (r.Props.Get("rPrChange") is { } rc) res.Add(Of(RevisionKind.Formatted, rc.Attributes(), rc, p, pos, pos + len));
                    pos += len;
                    break;
                }
                case InlineContainer box when InlineKinds.Contains(box.Name.LocalName):
                {
                    int start = pos;
                    var kind = box.Name.LocalName switch { "ins" => RevisionKind.Inserted, "del" => RevisionKind.Deleted, "moveFrom" => RevisionKind.MovedFrom, _ => RevisionKind.MovedTo };
                    var rev = Of(kind, box.Attributes, box, p, start, start);
                    int at = res.Count;
                    res.Add(rev);
                    pos = Inlines(p, box.Children, pos, res);
                    res[at] = rev with { End = pos };
                    break;
                }
                case InlineContainer box: pos = Inlines(p, box.Children, pos, res); break;
                case InlineContentControl sdt: pos = Inlines(p, sdt.Content, pos, res); break;
                default: pos += ParagraphAtoms.LengthOf(i); break;
            }
        }
        return pos;
    }

    /// <summary>Identifiants w:id des révisions et marques de commentaire d'une suite de blocs (identifiants uniques).</summary>
    public static IEnumerable<int> AnnotationIds(Story s)
    {
        foreach (var r in List(s))
        {
            var id = r.Node switch
            {
                InlineContainer c => Attr(c.Attributes, "id"),
                XElement e => (string?)e.Attributes().FirstOrDefault(a => a.Name.LocalName == "id"),
                TableRow row => (string?)(row.Props.Get("ins") ?? row.Props.Get("del"))?.Attributes().FirstOrDefault(a => a.Name.LocalName == "id"),
                _ => null,
            };
            if (int.TryParse(id, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n)) yield return n;
        }
        foreach (var p in s.Paragraphs())
            foreach (var o in OpaqueInlines(p.Inlines))
                if (o.Xml.Name.LocalName is "commentRangeStart" or "commentRangeEnd" && int.TryParse((string?)o.Xml.Attributes().FirstOrDefault(a => a.Name.LocalName == "id"), out var n))
                    yield return n;
    }

    static IEnumerable<OpaqueInline> OpaqueInlines(List<Inline> list)
    {
        foreach (var i in list)
            switch (i)
            {
                case OpaqueInline o: yield return o; break;
                case InlineContainer b: foreach (var x in OpaqueInlines(b.Children)) yield return x; break;
                case InlineContentControl c: foreach (var x in OpaqueInlines(c.Content)) yield return x; break;
            }
    }

    // ================================================================ accepter / refuser

    /// <summary>Accepte une révision (le document garde la modification).</summary>
    public static void Accept(Story root, Revision r) => Resolve(root, r, accept: true);

    /// <summary>Refuse une révision (le document revient à l'état d'avant).</summary>
    public static void Reject(Story root, Revision r) => Resolve(root, r, accept: false);

    /// <summary>
    /// Accepte (ou refuse) toutes les révisions choisies par <paramref name="filter"/> (toutes par défaut) ; renvoie leur
    /// nombre. Traitées de la dernière à la première : les fusions de paragraphes ne déplacent pas celles d'avant.
    /// </summary>
    public static int ResolveAll(Story root, bool accept, Func<Revision, bool>? filter = null)
    {
        var list = List(root).Where(r => filter?.Invoke(r) ?? true).ToList();
        for (int k = list.Count - 1; k >= 0; k--) Resolve(root, list[k], accept);
        if (!List(root).Any(r => r.Kind is RevisionKind.MovedFrom or RevisionKind.MovedTo))
            foreach (var p in root.Paragraphs()) RemoveMoveMarkers(p.Inlines);
        return list.Count;
    }

    static void RemoveMoveMarkers(List<Inline> list)
    {
        list.RemoveAll(i => i is OpaqueInline o && MoveMarkers.Contains(o.Xml.Name.LocalName));
        foreach (var b in list.OfType<InlineContainer>()) RemoveMoveMarkers(b.Children);
    }

    static void Resolve(Story root, Revision r, bool accept)
    {
        switch (r.Kind)
        {
            case RevisionKind.Inserted or RevisionKind.MovedTo or RevisionKind.Deleted or RevisionKind.MovedFrom:
            {
                var box = (InlineContainer)r.Node;
                if (FindOwner(root, box) is not { } found) return;
                bool keep = r.Kind is RevisionKind.Inserted or RevisionKind.MovedTo ? accept : !accept;
                int i = found.Owner.IndexOf(box);
                found.Owner.RemoveAt(i);
                if (keep)
                {
                    if (!accept) Undelete(box.Children);
                    found.Owner.InsertRange(i, box.Children);
                }
                RemoveEmptyBoxes(found.Paragraph.Inlines);
                break;
            }
            case RevisionKind.Formatted:
            {
                var change = (XElement)r.Node;
                var rpr = change.Parent;
                if (rpr is null) return;
                if (!accept && change.Elements().FirstOrDefault(e => e.Name.LocalName == "rPr") is { } old)
                {
                    // ancienne mise en forme ; les révisions de la marque (w:ins, w:del) restent
                    var kept = rpr.Elements().Where(e => e.Name.LocalName is "ins" or "del" or "moveFrom" or "moveTo").Select(e => new XElement(e)).ToList();
                    rpr.ReplaceNodes(kept.Concat(old.Elements().Select(e => new XElement(e))));
                }
                else change.Remove();
                break;
            }
            case RevisionKind.ParagraphFormatted:
            {
                var change = (XElement)r.Node;
                var ppr = change.Parent;
                if (ppr is null) return;
                if (!accept && change.Elements().FirstOrDefault(e => e.Name.LocalName == "pPr") is { } old)
                {
                    var p = r.Paragraph!;
                    var rpr = p.Props.Get("rPr");
                    var sect = p.Props.Get("sectPr");
                    var bag = new PropertyBag("pPr", p.Props.Ns, new XElement(ppr.Name, ppr.Attributes()));
                    foreach (var e in old.Elements()) bag.Set(new XElement(e));
                    if (rpr is not null) bag.Set(new XElement(rpr));
                    if (sect is not null) bag.Set(new XElement(sect));
                    ppr.ReplaceNodes(bag.Element!.Nodes());
                }
                else change.Remove();
                break;
            }
            case RevisionKind.ParagraphInserted or RevisionKind.ParagraphDeleted:
            {
                var marker = (XElement)r.Node;
                var p = r.Paragraph!;
                var rpr = marker.Parent;
                marker.Remove();
                if (rpr is not null && !rpr.HasElements && !rpr.HasAttributes) p.Props.Remove("rPr");
                bool merge = r.Kind == RevisionKind.ParagraphInserted ? !accept : accept;
                if (merge && FindBlock(root, p) is { } at && at.Index + 1 < at.Container.Blocks.Count && at.Container.Blocks[at.Index + 1] is Paragraph)
                    MergeWithNext(at.Container, at.Index);
                break;
            }
            case RevisionKind.RowInserted or RevisionKind.RowDeleted:
            {
                var row = (TableRow)r.Node;
                var marker = row.Props.Get(r.Kind == RevisionKind.RowInserted ? "ins" : "del");
                bool remove = r.Kind == RevisionKind.RowInserted ? !accept : accept;
                if (!remove) { marker?.Remove(); break; }
                foreach (var t in Tables(root))
                {
                    if (!t.Table.Items.Remove(row)) continue;
                    if (!t.Table.Items.OfType<TableRow>().Any()) t.Container.Blocks.Remove(t.Table);
                    break;
                }
                break;
            }
            default:
            {
                var change = (XElement)r.Node;
                var props = change.Parent;
                if (props is null) return;
                // tableau, ligne, cellule : l'ancienne valeur est l'unique enfant ; section, grille : approximation, la marque seule part
                if (!accept && change.Name.LocalName is "tblPrChange" or "trPrChange" or "tcPrChange" && change.Elements().FirstOrDefault() is { } old)
                {
                    var keep = props.Elements().Where(e => e.Name.LocalName is "ins" or "del" or "cellIns" or "cellDel" or "cellMerge").Select(e => new XElement(e)).ToList();
                    props.ReplaceNodes(old.Elements().Select(e => new XElement(e)).Concat(keep));
                }
                else change.Remove();
                break;
            }
        }
    }

    /// <summary>Texte supprimé qui revient (refus d'une suppression) : w:delText → w:t, w:delInstrText → w:instrText.</summary>
    static void Undelete(List<Inline> list)
    {
        foreach (var i in list)
            switch (i)
            {
                case Run r:
                    for (int k = 0; k < r.Content.Count; k++)
                    {
                        RunContent? back = r.Content[k] switch
                        {
                            TextContent { Deleted: true } t => new TextContent { Value = t.Value },
                            FieldCode { Deleted: true } f => new FieldCode { Value = f.Value },
                            _ => null,
                        };
                        if (back is null) continue;
                        back.Attributes.AddRange(r.Content[k].Attributes.Where(a => a.Name.LocalName != "space").Select(a => new XAttribute(a)));
                        r.Content[k] = back;
                    }
                    break;
                case InlineContainer b when b.Name.LocalName is not ("del" or "moveFrom"): Undelete(b.Children); break;
                case InlineContentControl c: Undelete(c.Content); break;
            }
    }

    /// <summary>Conteneur et liste propriétaire d'un élément en ligne dans une suite de blocs.</summary>
    static (Paragraph Paragraph, List<Inline> Owner)? FindOwner(Story root, Inline target)
    {
        foreach (var p in root.Paragraphs())
        {
            var path = new List<(List<Inline>, InlineContainer)>();
            if (TrackContext.PathTo(p.Inlines, target, path, out var owner) && owner is not null) return (p, owner);
        }
        return null;
    }

    static (Story Container, int Index)? FindBlock(Story s, Block target)
    {
        int i = s.Blocks.IndexOf(target);
        if (i >= 0) return (s, i);
        foreach (var b in s.Blocks)
        {
            var inner = b switch
            {
                Table t => t.Items.OfType<TableRow>().SelectMany(r => r.Items.OfType<TableCell>()).Select(c => FindBlock(c.Content, target)).FirstOrDefault(x => x is not null),
                BlockContentControl c => FindBlock(c.Content, target),
                _ => null,
            };
            if (inner is not null) return inner;
        }
        return null;
    }

    static IEnumerable<(Story Container, Table Table)> Tables(Story s)
    {
        foreach (var b in s.Blocks.ToList())
        {
            if (b is Table t)
            {
                yield return (s, t);
                foreach (var cell in t.Items.OfType<TableRow>().SelectMany(r => r.Items.OfType<TableCell>()))
                    foreach (var x in Tables(cell.Content)) yield return x;
            }
            else if (b is BlockContentControl c) foreach (var x in Tables(c.Content)) yield return x;
        }
    }

    /// <summary>
    /// Le paragraphe <paramref name="index"/> perd sa marque : son contenu passe au début du suivant, qui garde ses
    /// propriétés (c'est sa marque qui reste, comme dans Word).
    /// </summary>
    public static void MergeWithNext(Story container, int index)
    {
        var p = (Paragraph)container.Blocks[index];
        var next = (Paragraph)container.Blocks[index + 1];
        next.Inlines.InsertRange(0, p.Inlines);
        container.Blocks.RemoveAt(index);
    }

    /// <summary>Révisions et liens devenus vides (contenu supprimé, révision résolue).</summary>
    public static void RemoveEmptyBoxes(List<Inline> list)
    {
        foreach (var box in list.OfType<InlineContainer>().ToList())
        {
            RemoveEmptyBoxes(box.Children);
            if (box.Children.Count == 0 && box.Name.LocalName is "ins" or "del" or "moveFrom" or "moveTo" or "hyperlink") list.Remove(box);
        }
    }
}
