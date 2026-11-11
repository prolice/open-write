using System.Globalization;
using System.Xml.Linq;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;

namespace OpenWrite.Core.Review;

/// <summary>
/// Suivi des modifications en cours (lot 8) : quand un contexte est actif (<see cref="Use"/>), les opérations de texte
/// (<see cref="TextOps"/>, <see cref="Fragments"/>) et de mise en forme (<see cref="FormatOps"/>) enregistrent leurs
/// changements comme Word au lieu de les appliquer : texte inséré dans un w:ins, texte supprimé dans un w:del (w:delText),
/// marque de paragraphe insérée ou supprimée (w:pPr/w:rPr/w:ins|w:del), mise en forme d'avant gardée (w:rPrChange,
/// w:pPrChange). Supprimer son propre ajout le retire vraiment (Word fait de même).
/// </summary>
public sealed class TrackContext
{
    [ThreadStatic] static TrackContext? _current;

    /// <summary>Contexte du fil courant (null : pas de suivi).</summary>
    public static TrackContext? Current => _current;

    public required string Author { get; init; }
    /// <summary>Date fixe des révisions (tests) ; null : l'heure de chaque opération, comme Word.</summary>
    public DateTime? FixedDate { get; init; }
    public DateTime Date => FixedDate ?? DateTime.Now;

    /// <summary>
    /// Fin de la dernière suppression suivie (après le texte barré) : la touche Suppr y place le point d'insertion,
    /// sinon elle buterait sur le texte qu'elle vient de barrer.
    /// </summary>
    public DocPos? LastDeletionEnd { get; set; }
    readonly Func<int> _nextId;

    public TrackContext(Func<int> nextId) => _nextId = nextId;

    /// <summary>Active ce contexte jusqu'à la fin du bloc using (null : désactive le suivi).</summary>
    public static IDisposable Use(TrackContext? ctx)
    {
        var previous = _current;
        _current = ctx;
        return new Restore(previous);
    }

    sealed class Restore(TrackContext? previous) : IDisposable
    {
        public void Dispose() => _current = previous;
    }

    public string DateText => Date.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);

    /// <summary>Attributs w:id, w:author, w:date d'une nouvelle révision.</summary>
    public XAttribute[] Attributes(XNamespace w) =>
        new[] { new XAttribute(w + "id", _nextId().ToString(CultureInfo.InvariantCulture)), new XAttribute(w + "author", Author), new XAttribute(w + "date", DateText) };

    public bool IsMine(XElement e) => (string?)e.Attribute(e.Name.Namespace + "author") == Author;
    public bool IsMine(InlineContainer c) => c.Attributes.FirstOrDefault(a => a.Name.LocalName == "author")?.Value == Author;

    /// <summary>
    /// Générateur d'identifiants de révision : un de plus que le plus grand w:id des révisions et commentaires
    /// existants (Word demande des identifiants uniques parmi les annotations).
    /// </summary>
    public static Func<int> IdsAfter(IEnumerable<Story> stories)
    {
        int max = 0;
        foreach (var s in stories)
            foreach (var id in Revisions.AnnotationIds(s))
                if (id > max) max = id;
        int next = max;
        return () => ++next;
    }

    // ================================================================ enregistrement

    static readonly HashSet<string> RevisionBoxes = new() { "ins", "del", "moveFrom", "moveTo" };

    /// <summary>Chaîne des conteneurs (liens, révisions…) qui mènent à <paramref name="target"/>, du plus extérieur au plus proche.</summary>
    internal static bool PathTo(List<Inline> list, Inline target, List<(List<Inline> Owner, InlineContainer Box)> path, out List<Inline>? owner)
    {
        foreach (var i in list)
        {
            if (ReferenceEquals(i, target)) { owner = list; return true; }
            if (i is InlineContainer box)
            {
                path.Add((list, box));
                if (PathTo(box.Children, target, path, out owner)) return true;
                path.RemoveAt(path.Count - 1);
            }
            else if (i is InlineContentControl sdt && PathTo(sdt.Content, target, path, out owner)) return true;
        }
        owner = null;
        return false;
    }

    /// <summary>Segments entièrement compris dans [a, b[ (après coupe aux bornes), avec leur liste propriétaire.</summary>
    static List<(Run Run, List<Inline> Owner, List<InlineContainer> Boxes)> RunsIn(Paragraph p, int a, int b)
    {
        FormatOps.SplitRunAt(p, b);
        FormatOps.SplitRunAt(p, a);
        var res = new List<(Run, List<Inline>, List<InlineContainer>)>();
        var seen = new HashSet<Run>();
        foreach (var atom in ParagraphAtoms.Of(p))
        {
            if (atom.Run is not { } run || atom.Length == 0 || atom.Start < a || atom.End > b || !seen.Add(run)) continue;
            var path = new List<(List<Inline>, InlineContainer)>();
            if (!PathTo(p.Inlines, run, path, out var owner) || owner is null) continue;
            res.Add((run, owner, path.Select(x => x.Item2).ToList()));
        }
        return res;
    }

    /// <summary>Marque comme inséré le texte [a, b[ d'un paragraphe (déjà dans le modèle) : w:ins, fusionné avec un w:ins voisin du même auteur.</summary>
    public void MarkInserted(Paragraph p, int a, int b)
    {
        if (b <= a) return;
        var w = p.Props.Ns;
        foreach (var (run, owner, boxes) in RunsIn(p, a, b))
        {
            if (boxes.Any(x => x.Name.LocalName is "ins" or "moveTo" && IsMine(x))) continue;
            int i = owner.IndexOf(run);
            if (i > 0 && owner[i - 1] is InlineContainer prev && prev.Name.LocalName == "ins" && IsMine(prev))
            {
                owner.RemoveAt(i);
                prev.Children.Add(run);
                // le w:ins suivant du même auteur rejoint aussi celui-ci
                if (i < owner.Count && owner[i] is InlineContainer next && next.Name.LocalName == "ins" && IsMine(next))
                {
                    owner.RemoveAt(i);
                    prev.Children.AddRange(next.Children);
                }
                continue;
            }
            if (i + 1 < owner.Count && owner[i + 1] is InlineContainer nextBox && nextBox.Name.LocalName == "ins" && IsMine(nextBox))
            {
                owner.RemoveAt(i);
                nextBox.Children.Insert(0, run);
                continue;
            }
            var ins = new InlineContainer(w + "ins");
            ins.Attributes.AddRange(Attributes(w));
            ins.Children.Add(run);
            owner[i] = ins;
        }
    }

    /// <summary>
    /// Supprime avec suivi le texte [a, b[ d'un paragraphe : son propre ajout est retiré vraiment, le reste passe dans un
    /// w:del (texte en w:delText). Renvoie le nombre de caractères réellement retirés (les positions suivantes reculent d'autant).
    /// </summary>
    public int MarkDeleted(Paragraph p, int a, int b)
    {
        if (b <= a) return 0;
        var w = p.Props.Ns;
        int removed = 0;
        foreach (var (run, owner, boxes) in RunsIn(p, a, b))
        {
            if (boxes.Any(x => x.Name.LocalName is "del" or "moveFrom")) continue;           // déjà supprimé
            if (run.Content.Any(c => c is TextContent { Deleted: true } or FieldCode { Deleted: true })) continue;
            var nearestIns = boxes.LastOrDefault(x => x.Name.LocalName is "ins" or "moveTo");
            if (nearestIns is not null && IsMine(nearestIns))
            {
                removed += run.Content.Sum(ParagraphAtoms.LengthOf);
                owner.Remove(run);
                continue;
            }
            // texte → w:delText, code de champ → w:delInstrText
            for (int k = 0; k < run.Content.Count; k++)
            {
                RunContent? repl = run.Content[k] switch
                {
                    TextContent t => new TextContent { Value = t.Value, Deleted = true },
                    FieldCode f => new FieldCode { Value = f.Value, Deleted = true },
                    _ => null,
                };
                if (repl is null) continue;
                repl.Attributes.AddRange(run.Content[k].Attributes.Where(x => x.Name.LocalName != "space").Select(x => new XAttribute(x)));
                run.Content[k] = repl;
            }
            int i = owner.IndexOf(run);
            if (i > 0 && owner[i - 1] is InlineContainer prev && prev.Name.LocalName == "del" && IsMine(prev))
            {
                owner.RemoveAt(i);
                prev.Children.Add(run);
                continue;
            }
            if (i + 1 < owner.Count && owner[i + 1] is InlineContainer next && next.Name.LocalName == "del" && IsMine(next))
            {
                owner.RemoveAt(i);
                next.Children.Insert(0, run);
                continue;
            }
            var del = new InlineContainer(w + "del");
            del.Attributes.AddRange(Attributes(w));
            del.Children.Add(run);
            owner[i] = del;
        }
        Revisions.RemoveEmptyBoxes(p.Inlines);
        return removed;
    }

    /// <summary>Propriétés de la marque de paragraphe (w:pPr/w:rPr) vues comme un sac ordonné.</summary>
    static PropertyBag MarkBag(Paragraph p)
    {
        var w = p.Props.Ns;
        if (p.Props.Get("rPr") is not { } rpr)
        {
            rpr = new XElement(w + "rPr");
            p.Props.Set(rpr);
        }
        return new PropertyBag("rPr", w, rpr);
    }

    /// <summary>Marque de paragraphe insérée (Entrée avec suivi).</summary>
    public void MarkParagraphInserted(Paragraph p)
    {
        var w = p.Props.Ns;
        if (p.Props.Get("rPr")?.Element(w + "ins") is { } existing && IsMine(existing)) return;
        MarkBag(p).Set(new XElement(w + "ins", Attributes(w)));
    }

    /// <summary>
    /// Marque de paragraphe supprimée (fusion avec le suivant, avec suivi). Faux si c'était sa propre marque insérée :
    /// l'appelant fusionne alors vraiment les deux paragraphes.
    /// </summary>
    public bool MarkParagraphDeleted(Paragraph p)
    {
        var w = p.Props.Ns;
        var rpr = p.Props.Get("rPr");
        if (rpr?.Element(w + "ins") is { } ins && IsMine(ins))
        {
            ins.Remove();
            if (!rpr.HasElements && !rpr.HasAttributes) rpr.Remove();
            return false;
        }
        if (rpr?.Element(w + "del") is not null) return true;
        MarkBag(p).Set(new XElement(w + "del", Attributes(w)));
        return true;
    }

    /// <summary>Segment dans un ajout de cet auteur : sa mise en forme n'est pas suivie (comme Word).</summary>
    public bool InMyInsertion(Paragraph p, Run run)
    {
        var path = new List<(List<Inline>, InlineContainer)>();
        return PathTo(p.Inlines, run, path, out _) && path.Any(x => x.Item2.Name.LocalName is "ins" or "moveTo" && IsMine(x.Item2));
    }

    /// <summary>Mise en forme de caractère modifiée : l'ancienne est gardée dans w:rPrChange (la première fois seulement).</summary>
    public void RecordRunFormat(Run run, XElement? before)
    {
        var w = run.Props.Ns;
        var now = run.Props.Element;
        if (Same(before, now, "rPrChange") || now?.Element(w + "rPrChange") is not null) return;
        var old = new XElement(w + "rPr", before?.Elements().Where(e => e.Name.LocalName is not ("rPrChange" or "ins" or "del" or "moveFrom" or "moveTo")).Select(e => new XElement(e)));
        run.Props.Set(new XElement(w + "rPrChange", Attributes(w), old));
    }

    /// <summary>Mise en forme de paragraphe modifiée : l'ancienne est gardée dans w:pPrChange (la première fois seulement).</summary>
    public void RecordParagraphFormat(Paragraph p, XElement? before)
    {
        var w = p.Props.Ns;
        var now = p.Props.Element;
        static XElement? Strip(XElement? e) => e is null ? null
            : new XElement(e.Name, e.Elements().Where(x => x.Name.LocalName is not ("rPr" or "sectPr" or "pPrChange")).Select(x => new XElement(x)));
        if (XNode.DeepEquals(Strip(before) ?? new XElement(w + "pPr"), Strip(now) ?? new XElement(w + "pPr")) || now?.Element(w + "pPrChange") is not null) return;
        var old = new XElement(w + "pPr", Strip(before)?.Elements());
        p.Props.Set(new XElement(w + "pPrChange", Attributes(w), old));
    }

    static bool Same(XElement? a, XElement? b, string ignore)
    {
        static XElement N(XElement? e) => new("x", e?.Elements().Where(x => x.Name.LocalName != "rPrChange").Select(x => new XElement(x)));
        return XNode.DeepEquals(N(a), N(b));
    }

    // ================================================================ plages (plusieurs paragraphes)

    /// <summary>Marque comme inséré tout ce qui est entre deux positions d'un même conteneur (texte collé, Entrée…).</summary>
    public void MarkInsertedRange(Story root, DocPos from, DocPos to)
    {
        if (!from.ContainerPath.SequenceEqual(to.ContainerPath)) return;
        var container = DocPos.ResolveContainer(root, from.ContainerPath);
        for (int b = from.BlockIndex; b <= to.BlockIndex && b < container.Blocks.Count; b++)
        {
            if (container.Blocks[b] is not Paragraph p) continue;
            int a = b == from.BlockIndex ? from.Offset : 0;
            int e = b == to.BlockIndex ? to.Offset : ParagraphAtoms.Length(p);
            MarkInserted(p, a, e);
            if (b < to.BlockIndex) MarkParagraphInserted(p);
        }
    }

    /// <summary>
    /// Suppression suivie entre deux positions d'un même conteneur : texte de chaque paragraphe, marques des paragraphes
    /// fusionnés (ou vraie fusion si c'était sa propre marque insérée). Renvoie la position de fin après l'opération (pour
    /// placer le point d'insertion après le texte barré, touche Suppr).
    /// </summary>
    public DocPos DeleteRange(Story root, DocPos from, DocPos to)
    {
        if (!from.ContainerPath.SequenceEqual(to.ContainerPath))
            throw new NotSupportedException("Suppression entre des cellules ou à travers un tableau avec le suivi des modifications : pas encore prise en charge.");
        var container = DocPos.ResolveContainer(root, from.ContainerPath);
        for (int b = from.BlockIndex; b <= to.BlockIndex; b++)
            if (container.Blocks[b] is not Paragraph)
                throw new NotSupportedException("Suppression d'un tableau avec le suivi des modifications : pas encore prise en charge.");
        var end = to;
        LastDeletionEnd = null;
        // du dernier paragraphe au premier : les positions des précédents restent valables
        for (int b = to.BlockIndex; b >= from.BlockIndex; b--)
        {
            var p = (Paragraph)container.Blocks[b];
            int a = b == from.BlockIndex ? from.Offset : 0;
            int e = b == to.BlockIndex ? to.Offset : ParagraphAtoms.Length(p);
            int removed = MarkDeleted(p, a, e);
            if (b == to.BlockIndex) end = to.WithOffset(to.Offset - removed);
            if (b < to.BlockIndex && !MarkParagraphDeleted(p))
            {
                // sa propre marque insérée : les deux paragraphes se rejoignent vraiment (le texte de p passe devant le suivant)
                int len = ParagraphAtoms.Length(p);
                Revisions.MergeWithNext(container, b);
                if (end.BlockIndex == b + 1) end = new DocPos(end.ContainerPath.Append(b).ToArray(), end.Offset + len);
                else if (end.BlockIndex > b + 1) end = new DocPos(end.ContainerPath.Append(end.BlockIndex - 1).ToArray(), end.Offset);
            }
        }
        LastDeletionEnd = end;
        return end;
    }
}
