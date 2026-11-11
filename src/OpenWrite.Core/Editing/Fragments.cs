using System.Xml.Linq;
using OpenWrite.Core.Model;

namespace OpenWrite.Core.Editing;

/// <summary>
/// Morceau de document copié (presse-papiers interne, ou construit depuis du RTF / HTML) : des blocs, dont le premier
/// et le dernier paragraphes peuvent être partiels. <see cref="EndsWithParagraph"/> : la marque du dernier paragraphe
/// fait partie du morceau (sélection jusqu'au début du paragraphe suivant) ; sinon, le dernier paragraphe collé se
/// fond dans le paragraphe d'arrivée, comme dans Word.
/// </summary>
public sealed record Fragment(List<Block> Blocks, bool EndsWithParagraph)
{
    public bool IsEmpty => Blocks.Count == 0 || Blocks.Count == 1 && Blocks[0] is Paragraph p && p.Inlines.Count == 0 && !EndsWithParagraph;
}

public static class Fragments
{
    /// <summary>
    /// Copie de la plage [<paramref name="from"/>, <paramref name="to"/>[ (même conteneur) : blocs clonés, coupés aux bornes,
    /// sans ce qui ne peut pas être dupliqué tel quel (signets, commentaires, appels de notes, sauts de section, identifiants).
    /// </summary>
    public static Fragment Extract(Story root, DocPos from, DocPos to)
    {
        if (to < from) (from, to) = (to, from);
        bool endInBlock = to.Path.Count > 1 && !from.ContainerPath.SequenceEqual(to.ContainerPath);
        (from, to) = FormatOps.Normalize(from, to);
        var container = DocPos.ResolveContainer(root, from.ContainerPath);
        int last = to.BlockIndex;
        bool endsWithMark = false;
        // fin dans un tableau : le tableau entier est copié (il se termine par une marque de paragraphe implicite)
        if (endInBlock) endsWithMark = true;
        else if (to.Offset == 0 && to.BlockIndex > from.BlockIndex && container.Blocks[to.BlockIndex] is Paragraph) { last--; endsWithMark = true; }
        var tmp = new Story();
        tmp.Blocks.AddRange(container.Blocks.GetRange(from.BlockIndex, last - from.BlockIndex + 1).Select(b => b.Clone()));
        int n = tmp.Blocks.Count;
        if (!endsWithMark && tmp.Blocks[n - 1] is Paragraph lp)
            TextOps.Delete(tmp, new DocPos(new[] { n - 1 }, to.Offset), new DocPos(new[] { n - 1 }, ParagraphAtoms.Length(lp)));
        if (tmp.Blocks[0] is Paragraph && from.Offset > 0)
            TextOps.Delete(tmp, new DocPos(new[] { 0 }, 0), new DocPos(new[] { 0 }, from.Offset));
        Clean(tmp);
        return new Fragment(tmp.Blocks.ToList(), endsWithMark);
    }

    static readonly HashSet<string> DroppedInlines = new() { "commentRangeStart", "commentRangeEnd", "permStart", "permEnd", "proofErr" };
    static readonly HashSet<string> DroppedContents = new() { "commentReference", "footnoteReference", "endnoteReference", "annotationRef", "lastRenderedPageBreak" };

    /// <summary>Retire signets, commentaires, appels de notes, sauts de section et identifiants de paragraphe.</summary>
    static void Clean(Story s)
    {
        s.Blocks.RemoveAll(b => b is InlineMarkerBlock);
        foreach (var p in s.Paragraphs())
        {
            p.Props.Remove("sectPr");
            p.Attributes.RemoveAll(a => a.Name == FlowEdits.W14 + "paraId" || a.Name == FlowEdits.W14 + "textId");
            CleanInlines(p.Inlines);
        }
        foreach (var t in s.Blocks.OfType<Table>())
            foreach (var row in t.Items.OfType<TableRow>())
                row.Attributes.RemoveAll(a => a.Name == FlowEdits.W14 + "paraId" || a.Name == FlowEdits.W14 + "textId");
    }

    static void CleanInlines(List<Inline> list)
    {
        list.RemoveAll(i => i is BookmarkStart or BookmarkEnd || i is OpaqueInline o && DroppedInlines.Contains(o.Xml.Name.LocalName));
        foreach (var i in list)
            switch (i)
            {
                case Run r: r.Content.RemoveAll(c => c is OpaqueRunContent o && DroppedContents.Contains(o.Xml.Name.LocalName)); break;
                case InlineContainer box: CleanInlines(box.Children); break;
                case InlineContentControl sdt: CleanInlines(sdt.Content); break;
            }
        list.RemoveAll(i => i is Run { Content.Count: 0 });
    }

    /// <summary>
    /// Colle un morceau à une position : le paragraphe est coupé, les blocs insérés entre les deux moitiés ; le premier
    /// paragraphe collé se fond dans la première moitié (ou la remplace si elle est vide), le dernier dans la seconde s'il
    /// est partiel. Nouveaux identifiants de paragraphe. Renvoie la position après le texte collé.
    /// </summary>
    public static DocPos Insert(Story root, DocPos pos, Fragment fragment, Func<string> newParaId)
    {
        // suivi des modifications : le collage se fait normalement, puis tout ce qu'il a ajouté est marqué inséré
        var track = Review.TrackContext.Current;
        if (track is null) return InsertCore(root, pos, fragment, newParaId);
        DocPos end;
        using (Review.TrackContext.Use(null)) end = InsertCore(root, pos, fragment, newParaId);
        track.MarkInsertedRange(root, pos, end);
        return end;
    }

    static DocPos InsertCore(Story root, DocPos pos, Fragment fragment, Func<string> newParaId)
    {
        var blocks = fragment.Blocks.Select(b => b.Clone()).ToList();
        if (blocks.Count == 0) return pos;
        var tmp = new Story();
        tmp.Blocks.AddRange(blocks);
        foreach (var p in tmp.Paragraphs())
            foreach (var name in new[] { "paraId", "textId" })
                if (p.Attributes.FirstOrDefault(a => a.Name == FlowEdits.W14 + name) is { } attr) attr.Value = newParaId();

        var container = DocPos.ResolveContainer(root, pos.ContainerPath);
        var prefix = pos.ContainerPath.ToArray();
        DocPos At(int index, int offset) => new(prefix.Append(index).ToArray(), offset);
        int i = pos.BlockIndex;
        bool leftEmpty = pos.Offset == 0;
        TextOps.Split(root, pos, newParaId);                 // L en i, R en i + 1
        var left = (Paragraph)container.Blocks[i];
        container.Blocks.InsertRange(i + 1, blocks);
        int rightIndex = i + 1 + blocks.Count;

        // premier paragraphe collé : remplace la moitié gauche vide, sinon s'y ajoute
        if (blocks[0] is Paragraph first)
        {
            if (leftEmpty && ParagraphAtoms.Length(left) == 0 && (blocks.Count > 1 || fragment.EndsWithParagraph))
            {
                // la marque de paragraphe d'arrivée (saut de section éventuel) n'est jamais à gauche : rien à garder
                container.Blocks.RemoveAt(i);
            }
            else
            {
                left.Inlines.AddRange(first.Inlines);
                container.Blocks.RemoveAt(i + 1);
            }
            rightIndex--;
        }
        if (fragment.EndsWithParagraph) return At(rightIndex, 0);
        // dernier paragraphe partiel : la moitié droite s'y ajoute ; la marque gardée est celle du paragraphe d'arrivée
        // (Word) : propriétés de la moitié droite, saut de section compris
        var rightProps = ((Paragraph)container.Blocks[rightIndex]).Props.Clone();
        var lastIndex = rightIndex - 1;
        if (container.Blocks[lastIndex] is not Paragraph lastP) return At(rightIndex, 0);
        int offset = ParagraphAtoms.Length(lastP);
        TextOps.Delete(root, At(lastIndex, offset), At(rightIndex, 0));
        if (!ReferenceEquals(lastP, left)) lastP.Props = rightProps;
        return At(lastIndex, offset);
    }

    /// <summary>Remplace une plage (éventuellement vide) par un morceau, en une seule opération annulable.</summary>
    public static DocPos Paste(UndoJournal journal, Story root, string label, DocPos from, DocPos to, Fragment fragment, Func<string> newParaId)
    {
        if (to < from) (from, to) = (to, from);
        if (!from.ContainerPath.SequenceEqual(to.ContainerPath) && TextOps.AcrossTopLevel(from, to) is var (s, _, e, _))
            return TextOps.Edit(journal, root, label, Array.Empty<int>(), s, e - s + 1, to,
                () => Insert(root, TextOps.DeleteAcross(root, from, to), fragment, newParaId), p => p);
        return TextOps.Edit(journal, root, label, from.ContainerPath, from.BlockIndex, to.BlockIndex - from.BlockIndex + 1, to,
            () => Insert(root, from.Equals(to) ? from : TextOps.Delete(root, from, to), fragment, newParaId), p => p);
    }

    // ------------------------------------------------------------------ construction (RTF, HTML)

    /// <summary>Mise en forme de caractère lue dans un format d'échange.</summary>
    public sealed record CharFormat
    {
        public bool Bold { get; init; }
        public bool Italic { get; init; }
        public bool Underline { get; init; }
        public bool Strike { get; init; }
        public double? Size { get; init; }
        public string? Font { get; init; }
        /// <summary>RRGGBB.</summary>
        public string? Color { get; init; }
        public string? VertAlign { get; init; }
    }

    /// <summary>Paragraphe lu dans un format d'échange : alignement et segments (\t et \v : tabulation et retour à la ligne).</summary>
    public sealed class ParaSpec
    {
        public string? Align { get; set; }
        public List<(string Text, CharFormat Format)> Runs { get; } = new();
    }

    /// <summary>Morceau construit depuis des paragraphes lus (le dernier est partiel sauf <paramref name="lastHasMark"/>).</summary>
    public static Fragment Build(IReadOnlyList<ParaSpec> paras, bool lastHasMark, XNamespace? w = null)
    {
        w ??= PropertyBag.TransitionalNs;
        var blocks = new List<Block>();
        foreach (var spec in paras)
        {
            var p = new Paragraph { Props = new PropertyBag("pPr", w) };
            if (spec.Align is { } jc && jc != "left") p.Props.SetVal("jc", jc);
            foreach (var (text, f) in spec.Runs)
            {
                if (text.Length == 0) continue;
                var run = new Run { Props = new PropertyBag("rPr", w) };
                if (f.Font is { } font) FormatOps.SetFont(run, font);
                if (f.Bold) FormatOps.SetToggle(run, "b", true);
                if (f.Italic) FormatOps.SetToggle(run, "i", true);
                if (f.Strike) FormatOps.SetToggle(run, "strike", true);
                if (f.Underline) FormatOps.SetUnderline(run, true);
                if (f.Color is { } col) FormatOps.SetValue(run, "color", col);
                if (f.Size is { } sz) FormatOps.SetValue(run, "sz", Math.Round(sz * 2).ToString(System.Globalization.CultureInfo.InvariantCulture));
                if (f.VertAlign is { } va) FormatOps.SetValue(run, "vertAlign", va);
                int start = 0;
                for (int k = 0; k <= text.Length; k++)
                {
                    if (k < text.Length && text[k] is not ('\t' or '\v')) continue;
                    if (k > start) run.Content.Add(new TextContent { Value = text[start..k] });
                    if (k < text.Length) run.Content.Add(text[k] == '\t' ? new TabContent() : new BreakContent());
                    start = k + 1;
                }
                if (run.Content.Count > 0) p.Inlines.Add(run);
            }
            blocks.Add(p);
        }
        return new Fragment(blocks, lastHasMark);
    }
}
