using System.Globalization;

namespace OpenWrite.Core.Model;

/// <summary>
/// Opérations sur les champs et les signets (lot 1). Comme FlowEdits, elles modifient le modèle en place ;
/// l'appelant marque la partie modifiée (StoryPart.Dirty) et le journal d'annulation les enveloppera.
/// </summary>
public static class MarkerOps
{
    // ================================================================ champs

    /// <summary>
    /// Remplace le résultat affiché d'un champ (ce que fera la mise à jour des champs PAGE, DATE, REF…) : le texte va dans
    /// le premier segment du résultat (sa mise en forme est gardée), les autres segments de texte du résultat sont vidés.
    /// Sans segment de résultat, un segment est créé après le séparateur avec la mise en forme du séparateur.
    /// Renvoie faux pour un champ simple, sans séparateur ou dont le résultat contient un champ imbriqué ou s'étend sur
    /// plusieurs paragraphes (cas de la table des matières, régénérée entièrement au lot 9).
    /// </summary>
    public static bool SetResultText(Field f, string text)
    {
        if (f.Simple is not null || f.Separate is null || f.End is null || f.BeginParagraph != f.EndParagraph || f.BeginParagraph is null) return false;
        if (f.Children.Any(c => !f.CodeParts.Contains(c))) return false;     // champ imbriqué dans le résultat
        var p = f.BeginParagraph;
        if (f.ResultRuns.Count > 0)
        {
            f.ResultRuns[0].Text.Value = text;
            foreach (var (run, t) in f.ResultRuns.Skip(1)) run.Content.Remove(t);
            foreach (var (run, _) in f.ResultRuns.Skip(1).Where(r => r.Run.Content.Count == 0)) RemoveRun(p.Inlines, run);
            return true;
        }
        var sepRun = FindRunOf(p.Inlines, f.Separate);
        if (sepRun is null) return false;
        var owner = OwnerOf(p.Inlines, sepRun)!;
        var r = new Run { Props = sepRun.Props.Clone() };
        r.Content.Add(new TextContent { Value = text });
        owner.Insert(owner.IndexOf(sepRun) + 1, r);
        return true;
    }

    static Run? FindRunOf(List<Inline> inlines, RunContent c)
    {
        foreach (var i in inlines)
        {
            if (i is Run r && r.Content.Contains(c)) return r;
            if (Children(i) is { } ch && FindRunOf(ch, c) is { } found) return found;
        }
        return null;
    }

    static List<Inline>? Children(Inline i) => i switch { InlineContainer b => b.Children, InlineContentControl s => s.Content, _ => null };

    static List<Inline>? OwnerOf(List<Inline> inlines, Inline target)
    {
        if (inlines.Contains(target)) return inlines;
        foreach (var i in inlines)
            if (Children(i) is { } ch && OwnerOf(ch, target) is { } o) return o;
        return null;
    }

    static void RemoveRun(List<Inline> inlines, Run run) => OwnerOf(inlines, run)?.Remove(run);

    // ================================================================ signets

    /// <summary>Nom de signet accepté par Word : 1 à 40 caractères, une lettre d'abord, puis lettres, chiffres ou « _ ».</summary>
    public static bool IsValidName(string name) =>
        name.Length is >= 1 and <= 40 && char.IsLetter(name[0]) && name.All(c => char.IsLetterOrDigit(c) || c == '_');

    /// <summary>Identifiant libre : un de plus que le plus grand identifiant de signet des parties données.</summary>
    public static string NextId(IEnumerable<Story> stories)
    {
        int max = -1;
        foreach (var b in stories.SelectMany(s => BookmarkIndex.Build(s).All))
            if (int.TryParse(b.Id, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n)) max = Math.Max(max, n);
        return (max + 1).ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>Signet autour de tout le contenu d'un paragraphe (comme Insertion ▸ Signet sur un paragraphe sélectionné).</summary>
    public static Bookmark AddAroundParagraph(Paragraph p, string name, string id)
    {
        if (!IsValidName(name)) throw new ArgumentException($"Nom de signet incorrect : « {name} ».", nameof(name));
        var ns = p.Props.Ns;
        var start = new BookmarkStart(ns);
        start.Id = id;
        start.Name = name;
        var end = new BookmarkEnd(ns) { Id = id };
        p.Inlines.Insert(0, start);
        p.Inlines.Add(end);
        return new Bookmark { Start = start, End = end, StartParagraph = p, EndParagraph = p };
    }

    /// <summary>Signet autour d'un segment.</summary>
    public static Bookmark AddAroundRun(Paragraph p, Run run, string name, string id)
    {
        if (!IsValidName(name)) throw new ArgumentException($"Nom de signet incorrect : « {name} ».", nameof(name));
        var owner = OwnerOf(p.Inlines, run) ?? throw new ArgumentException("Le segment n'appartient pas au paragraphe.", nameof(run));
        var ns = p.Props.Ns;
        var start = new BookmarkStart(ns);
        start.Id = id;
        start.Name = name;
        var end = new BookmarkEnd(ns) { Id = id };
        owner.Insert(owner.IndexOf(run), start);
        owner.Insert(owner.IndexOf(run) + 1, end);
        return new Bookmark { Start = start, End = end, StartParagraph = p, EndParagraph = p };
    }

    public static void Rename(Bookmark b, string name)
    {
        if (!IsValidName(name)) throw new ArgumentException($"Nom de signet incorrect : « {name} ».", nameof(name));
        b.Start.Name = name;
    }

    /// <summary>Supprime les deux marqueurs d'un signet (le texte couvert reste).</summary>
    public static void Remove(Story story, Bookmark b)
    {
        RemoveMarker(story, b.Start);
        if (b.End is not null) RemoveMarker(story, b.End);
    }

    static bool RemoveMarker(Story story, Inline marker)
    {
        for (int i = 0; i < story.Blocks.Count; i++)
        {
            switch (story.Blocks[i])
            {
                case InlineMarkerBlock mb when mb.Marker == marker:
                    story.Blocks.RemoveAt(i);
                    return true;
                case Paragraph p when OwnerOf(p.Inlines, marker) is { } owner:
                    owner.Remove(marker);
                    return true;
                case Table t:
                    if (t.Items.Remove(marker)) return true;
                    foreach (var row in t.Items.OfType<TableRow>())
                    {
                        if (row.Items.Remove(marker)) return true;
                        foreach (var cell in row.Items.OfType<TableCell>())
                            if (RemoveMarker(cell.Content, marker)) return true;
                    }
                    break;
                case BlockContentControl s when RemoveMarker(s.Content, marker):
                    return true;
            }
        }
        return false;
    }
}
