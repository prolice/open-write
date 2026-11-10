namespace OpenWrite.Core.Model;

/// <summary>
/// Position dans une suite de blocs racine (corps, en-tête, note…) : chemin du paragraphe + décalage en caractères.
/// <para>
/// Chemin : index des blocs du conteneur, de la racine au paragraphe. Un tableau consomme trois index (bloc du tableau,
/// élément de <see cref="Table.Items"/>, élément de <see cref="TableRow.Items"/>), un contrôle de contenu un seul, puis on
/// continue dans la cellule ou le contrôle. Exemple : [4, 1, 2, 0] = 1er paragraphe de la 3e cellule (élément 2) de
/// la ligne d'élément 1 du tableau qui est le 5e bloc.
/// </para>
/// <para>
/// Décalage : caractères du paragraphe au sens d'OpenWrite (<see cref="ParagraphAtoms"/>) — texte (y compris supprimé
/// avec suivi des modifications), code de champ, et 1 pour une tabulation, un retour à la ligne, un caractère de champ
/// ou un objet (dessin, appel de note) ; 0 pour les repères (signets, commentaires, erreurs d'orthographe).
/// </para>
/// Les positions ne référencent aucun objet : elles restent valables quand l'annulation remplace des blocs par leurs copies.
/// </summary>
public sealed class DocPos : IEquatable<DocPos>, IComparable<DocPos>
{
    public DocPos(IReadOnlyList<int> path, int offset)
    {
        if (path.Count == 0) throw new ArgumentException("Chemin vide.", nameof(path));
        Path = path.ToArray();
        Offset = offset;
    }

    public IReadOnlyList<int> Path { get; }
    public int Offset { get; }

    /// <summary>Index du paragraphe dans son conteneur.</summary>
    public int BlockIndex => Path[^1];
    /// <summary>Chemin du conteneur (suite de blocs qui contient le paragraphe).</summary>
    public IReadOnlyList<int> ContainerPath => Path.Take(Path.Count - 1).ToArray();

    public DocPos WithOffset(int offset) => new(Path, offset);

    public bool Equals(DocPos? o) => o is not null && Offset == o.Offset && Path.SequenceEqual(o.Path);
    public override bool Equals(object? obj) => Equals(obj as DocPos);
    public override int GetHashCode() => Path.Aggregate(Offset, HashCode.Combine);

    /// <summary>Ordre du document (les index croissent dans l'ordre de lecture).</summary>
    public int CompareTo(DocPos? o)
    {
        if (o is null) return 1;
        for (int i = 0; i < Math.Min(Path.Count, o.Path.Count); i++)
            if (Path[i] != o.Path[i]) return Path[i].CompareTo(o.Path[i]);
        return Path.Count != o.Path.Count ? Path.Count.CompareTo(o.Path.Count) : Offset.CompareTo(o.Offset);
    }

    public static bool operator <(DocPos a, DocPos b) => a.CompareTo(b) < 0;
    public static bool operator >(DocPos a, DocPos b) => a.CompareTo(b) > 0;

    public override string ToString() => $"[{string.Join(",", Path)}]:{Offset}";

    // ---------------------------------------------------------------- résolution

    /// <summary>Suite de blocs désignée par un chemin de conteneur (racine si vide).</summary>
    public static Story ResolveContainer(Story root, IReadOnlyList<int> containerPath)
    {
        var story = root;
        int i = 0;
        while (i < containerPath.Count)
        {
            var block = story.Blocks[containerPath[i]];
            switch (block)
            {
                case Table t when i + 2 < containerPath.Count:
                    var row = t.Items[containerPath[i + 1]] as TableRow ?? throw new ArgumentException($"Chemin {Show(containerPath)} : pas une ligne de tableau.");
                    var cell = row.Items[containerPath[i + 2]] as TableCell ?? throw new ArgumentException($"Chemin {Show(containerPath)} : pas une cellule.");
                    story = cell.Content;
                    i += 3;
                    break;
                case BlockContentControl s:
                    story = s.Content;
                    i += 1;
                    break;
                default:
                    throw new ArgumentException($"Chemin {Show(containerPath)} : le bloc {containerPath[i]} ne contient pas de blocs.");
            }
        }
        return story;
    }

    static string Show(IReadOnlyList<int> p) => "[" + string.Join(",", p) + "]";

    public Paragraph Paragraph(Story root) =>
        ResolveContainer(root, ContainerPath).Blocks[BlockIndex] as Paragraph
        ?? throw new ArgumentException($"La position {this} ne désigne pas un paragraphe.");

    /// <summary>Position d'un paragraphe (recherche dans la racine, ses tableaux et ses contrôles de contenu).</summary>
    public static DocPos Of(Story root, Paragraph p, int offset = 0) =>
        Find(root, p, new List<int>()) is { } path ? new DocPos(path, offset) : throw new ArgumentException("Paragraphe absent de cette suite de blocs.");

    static List<int>? Find(Story s, Paragraph p, List<int> prefix)
    {
        for (int i = 0; i < s.Blocks.Count; i++)
        {
            switch (s.Blocks[i])
            {
                case Paragraph q when q == p:
                    return new List<int>(prefix) { i };
                case Table t:
                    for (int r = 0; r < t.Items.Count; r++)
                        if (t.Items[r] is TableRow row)
                            for (int c = 0; c < row.Items.Count; c++)
                                if (row.Items[c] is TableCell cell && Find(cell.Content, p, new List<int>(prefix) { i, r, c }) is { } found)
                                    return found;
                    break;
                case BlockContentControl sdt when Find(sdt.Content, p, new List<int>(prefix) { i }) is { } inSdt:
                    return inSdt;
            }
        }
        return null;
    }
}

/// <summary>
/// Élément positionnable d'un paragraphe : contenu de segment (<see cref="Content"/> non null) ou élément de ligne
/// (repère, îlot). <see cref="Owner"/> est la liste qui contient <see cref="Inline"/> (paragraphe ou conteneur),
/// <see cref="Container"/> le conteneur qui l'entoure le plus près (lien, révision…), s'il y en a un.
/// </summary>
public sealed record Atom(List<Inline> Owner, Inline Inline, InlineContainer? Container, RunContent? Content, int Start, int Length)
{
    public int End => Start + Length;
    public Run? Run => Inline as Run;
    /// <summary>Atome dans une suppression suivie (w:del, w:moveFrom) : on n'y insère jamais de texte.</summary>
    public bool InDeletion => Container?.Name.LocalName is "del" or "moveFrom" || Content is TextContent { Deleted: true };
}

/// <summary>Découpage d'un paragraphe en atomes, dans l'ordre (voir <see cref="DocPos"/> pour la longueur de chacun).</summary>
public static class ParagraphAtoms
{
    /// <summary>Éléments sans largeur dans le texte.</summary>
    static readonly HashSet<string> ZeroWidth = new()
    {
        "proofErr", "commentRangeStart", "commentRangeEnd", "permStart", "permEnd", "moveFromRangeStart", "moveFromRangeEnd",
        "moveToRangeStart", "moveToRangeEnd", "customXmlInsRangeStart", "customXmlInsRangeEnd", "customXmlDelRangeStart",
        "customXmlDelRangeEnd", "lastRenderedPageBreak",
    };

    public static int LengthOf(RunContent c) => c switch
    {
        TextContent t => t.Value.Length,
        FieldCode f => f.Value.Length,
        OpaqueRunContent o when ZeroWidth.Contains(o.Xml.Name.LocalName) => 0,
        _ => 1,
    };

    public static int LengthOf(Inline i) => i switch
    {
        OpaqueInline o => ZeroWidth.Contains(o.Xml.Name.LocalName) ? 0 : 1,
        _ => 0,
    };

    public static List<Atom> Of(Paragraph p)
    {
        var res = new List<Atom>();
        int pos = 0;
        Walk(p.Inlines, null, res, ref pos);
        return res;
    }

    public static int Length(Paragraph p) => Of(p).Sum(a => a.Length);

    static void Walk(List<Inline> list, InlineContainer? container, List<Atom> res, ref int pos)
    {
        foreach (var i in list)
        {
            switch (i)
            {
                case Run r:
                    if (r.Content.Count == 0) res.Add(new Atom(list, r, container, null, pos, 0));
                    foreach (var c in r.Content)
                    {
                        int len = LengthOf(c);
                        res.Add(new Atom(list, r, container, c, pos, len));
                        pos += len;
                    }
                    break;
                case InlineContainer box:
                    Walk(box.Children, box, res, ref pos);
                    break;
                case InlineContentControl s:
                    Walk(s.Content, container, res, ref pos);
                    break;
                default:
                    int l = LengthOf(i);
                    res.Add(new Atom(list, i, container, null, pos, l));
                    pos += l;
                    break;
            }
        }
    }
}
