using System.Text;

namespace OpenWrite.Core.Model;

/// <summary>Signet trouvé dans le texte : marqueurs de début et de fin appariés par w:id.</summary>
public sealed class Bookmark
{
    public required BookmarkStart Start { get; init; }
    public BookmarkEnd? End { get; internal set; }
    public Paragraph? StartParagraph { get; internal init; }
    public Paragraph? EndParagraph { get; internal set; }
    public string Name => Start.Name;
    public string Id => Start.Id;
    /// <summary>Signet masqué (créé par Word : _GoBack, _Toc…, _Ref…) : absent de la boîte Signet, sauf demande.</summary>
    public bool Hidden => Name.StartsWith('_');
    /// <summary>Signet de colonnes de tableau (w:colFirst / w:colLast).</summary>
    public bool IsColumnRange => Start.Get("colFirst") is not null;
    internal readonly StringBuilder TextBuilder = new();
    /// <summary>Texte couvert par le signet (« \r » entre deux paragraphes).</summary>
    public string Text => TextBuilder.ToString();
    public override string ToString() => $"{Name} (#{Id}) « {Text} »";
}

/// <summary>Index des signets d'une suite de blocs.</summary>
public sealed class BookmarkIndex
{
    public List<Bookmark> All { get; } = new();
    public List<string> Problems { get; } = new();

    /// <summary>Signet par nom (insensible à la casse, comme Word) ; le premier en cas de doublon.</summary>
    public Bookmark? Find(string name) => All.FirstOrDefault(b => string.Equals(b.Name, name, StringComparison.OrdinalIgnoreCase));

    public static BookmarkIndex Build(Story story)
    {
        var idx = new BookmarkIndex();
        var open = new Dictionary<string, Bookmark>();
        foreach (var e in FlowWalker.Walk(story))
        {
            switch (e.Kind)
            {
                case FlowEventKind.Marker when e.Item is BookmarkStart s:
                    var b = new Bookmark { Start = s, StartParagraph = e.Paragraph };
                    if (open.ContainsKey(s.Id)) idx.Problems.Add($"identifiant de signet en double : {s.Id}");
                    open[s.Id] = b;
                    idx.All.Add(b);
                    break;
                case FlowEventKind.Marker when e.Item is BookmarkEnd end:
                    if (open.Remove(end.Id, out var started)) { started.End = end; started.EndParagraph = e.Paragraph; }
                    else idx.Problems.Add($"fin de signet sans début : {end.Id}");
                    break;
                case FlowEventKind.Content when e.Item is TextContent { Deleted: false } t:
                    foreach (var o in open.Values) o.TextBuilder.Append(t.Value);
                    break;
                case FlowEventKind.Content when e.Item is TabContent:
                    foreach (var o in open.Values) o.TextBuilder.Append('\t');
                    break;
                case FlowEventKind.ParagraphEnd:
                    foreach (var o in open.Values) o.TextBuilder.Append('\r');
                    break;
            }
        }
        foreach (var b in open.Values) idx.Problems.Add($"signet non fermé : {b.Name}");
        foreach (var g in idx.All.GroupBy(b => b.Name, StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1))
            idx.Problems.Add($"nom de signet en double : {g.Key}");
        return idx;
    }
}
