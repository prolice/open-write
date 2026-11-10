namespace OpenWrite.Core.Model;

public enum FlowEventKind { ParagraphStart, ParagraphEnd, Content, Marker, ContainerStart, ContainerEnd }

/// <summary>
/// Étape du parcours : <see cref="Item"/> est un <see cref="RunContent"/> (Content), un marqueur de signet (Marker)
/// ou un <see cref="InlineContainer"/> (ContainerStart/End). <see cref="Paragraph"/> est null pour un marqueur placé
/// entre deux blocs, deux lignes ou deux cellules.
/// </summary>
public sealed record FlowEvent(FlowEventKind Kind, Paragraph? Paragraph, Run? Run = null, object? Item = null);

/// <summary>Parcours d'une suite de blocs dans l'ordre du document (cellules, contrôles de contenu, liens et révisions compris).</summary>
public static class FlowWalker
{
    public static IEnumerable<FlowEvent> Walk(Story story)
    {
        foreach (var b in story.Blocks)
            foreach (var e in Walk(b)) yield return e;
    }

    static IEnumerable<FlowEvent> Walk(Block b)
    {
        switch (b)
        {
            case Paragraph p:
                yield return new FlowEvent(FlowEventKind.ParagraphStart, p);
                foreach (var e in Walk(p, p.Inlines)) yield return e;
                yield return new FlowEvent(FlowEventKind.ParagraphEnd, p);
                break;
            case Table t:
                foreach (var item in t.Items)
                {
                    if (item is Inline m) { yield return new FlowEvent(FlowEventKind.Marker, null, Item: m); continue; }
                    if (item is not TableRow row) continue;
                    foreach (var ci in row.Items)
                    {
                        if (ci is Inline rm) yield return new FlowEvent(FlowEventKind.Marker, null, Item: rm);
                        else if (ci is TableCell cell)
                            foreach (var e in Walk(cell.Content)) yield return e;
                    }
                }
                break;
            case BlockContentControl s:
                foreach (var e in Walk(s.Content)) yield return e;
                break;
            case InlineMarkerBlock mb:
                yield return new FlowEvent(FlowEventKind.Marker, null, Item: mb.Marker);
                break;
        }
    }

    static IEnumerable<FlowEvent> Walk(Paragraph p, IEnumerable<Inline> inlines)
    {
        foreach (var i in inlines)
            switch (i)
            {
                case Run r:
                    foreach (var c in r.Content) yield return new FlowEvent(FlowEventKind.Content, p, r, c);
                    break;
                case InlineContainer box:
                    yield return new FlowEvent(FlowEventKind.ContainerStart, p, Item: box);
                    foreach (var e in Walk(p, box.Children)) yield return e;
                    yield return new FlowEvent(FlowEventKind.ContainerEnd, p, Item: box);
                    break;
                case InlineContentControl s:
                    foreach (var e in Walk(p, s.Content)) yield return e;
                    break;
                case BookmarkStart or BookmarkEnd:
                    yield return new FlowEvent(FlowEventKind.Marker, p, Item: i);
                    break;
            }
    }
}
