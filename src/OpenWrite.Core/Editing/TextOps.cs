using System.Xml.Linq;
using OpenWrite.Core.Model;
using OpenWrite.Core.Review;

namespace OpenWrite.Core.Editing;

/// <summary>
/// Opérations de texte exprimées en positions (<see cref="DocPos"/>). Elles modifient le modèle en place ;
/// pour les rendre annulables, on les exécute dans une transaction du <see cref="UndoJournal"/> (voir <see cref="Edit"/>).
/// </summary>
public static class TextOps
{
    // ================================================================ insérer du texte

    /// <summary>
    /// Insère du texte (sans tabulation ni retour : voir <see cref="Split"/>) à une position. Comme Word, le texte prend
    /// la mise en forme du caractère qui précède ; au début d'un paragraphe, celle du caractère qui suit ; dans un
    /// paragraphe vide, celle de la marque de paragraphe. Jamais dans une suppression suivie. Renvoie la position après le texte.
    /// </summary>
    public static DocPos Insert(Story root, DocPos pos, string text)
    {
        var end = InsertCore(root, pos, text);
        // suivi des modifications : le texte tapé passe dans un w:ins
        if (TrackContext.Current is { } t && text.Length > 0) t.MarkInserted(pos.Paragraph(root), pos.Offset, end.Offset);
        return end;
    }

    static DocPos InsertCore(Story root, DocPos pos, string text)
    {
        if (text.Length == 0) return pos;
        if (text.Any(c => c is '\t' or '\r' or '\n' or '\v')) throw new ArgumentException("Tabulations et retours : opérations dédiées.", nameof(text));
        var p = pos.Paragraph(root);
        var atoms = ParagraphAtoms.Of(p);
        int o = pos.Offset;
        if (o < 0 || o > atoms.Sum(a => a.Length)) throw new ArgumentOutOfRangeException(nameof(pos), $"Décalage {o} hors du paragraphe.");

        // 1. au milieu d'un texte
        var inside = atoms.FirstOrDefault(a => a.Content is TextContent && !a.InDeletion && a.Start < o && o < a.End);
        if (inside is not null) { var t = (TextContent)inside.Content!; t.Value = t.Value.Insert(o - inside.Start, text); return pos.WithOffset(o + text.Length); }

        // 2. après le caractère qui précède (sa mise en forme)
        var before = atoms.LastOrDefault(a => a.Length > 0 && a.End == o && a.Run is not null && !a.InDeletion);
        if (before is not null)
        {
            if (before.Content is TextContent bt) bt.Value += text;
            else InsertTextAfter(before.Run!, before.Content!, text);
            return pos.WithOffset(o + text.Length);
        }
        // 3. avant le caractère qui suit
        var after = atoms.FirstOrDefault(a => a.Length > 0 && a.Start == o && a.Run is not null && !a.InDeletion);
        if (after is not null)
        {
            if (after.Content is TextContent at) at.Value = text + at.Value;
            else after.Run!.Content.Insert(after.Run.Content.IndexOf(after.Content!), new TextContent { Value = text });
            return pos.WithOffset(o + text.Length);
        }
        // 4. nouveau segment avec la mise en forme de la marque de paragraphe, après le dernier repère placé avant o
        var run = new Run { Props = MarkProps(p) };
        run.Content.Add(new TextContent { Value = text });
        var anchor = atoms.LastOrDefault(a => a.End <= o && !a.InDeletion);
        if (anchor is null) p.Inlines.Insert(0, run);
        else anchor.Owner.Insert(anchor.Owner.IndexOf(anchor.Inline) + 1, run);
        return pos.WithOffset(o + text.Length);
    }

    /// <summary>
    /// Insère un caractère spécial (tabulation, retour à la ligne, saut de page) à une position, avec les mêmes règles de
    /// mise en forme que <see cref="Insert"/> ; un texte coupé en deux reste dans le même segment. Renvoie la position après.
    /// </summary>
    public static DocPos InsertSpecial(Story root, DocPos pos, SpecialChar kind)
    {
        var end = InsertSpecialCore(root, pos, kind);
        if (TrackContext.Current is { } t) t.MarkInserted(pos.Paragraph(root), pos.Offset, end.Offset);
        return end;
    }

    static DocPos InsertSpecialCore(Story root, DocPos pos, SpecialChar kind)
    {
        var p = pos.Paragraph(root);
        var ns = p.Props.Ns;
        RunContent content = kind switch
        {
            SpecialChar.Tab => new TabContent(),
            SpecialChar.LineBreak => new BreakContent(),
            SpecialChar.PageBreak => new BreakContent { Attributes = { new XAttribute(ns + "type", "page") } },
            SpecialChar.ColumnBreak => new BreakContent { Attributes = { new XAttribute(ns + "type", "column") } },
            _ => throw new ArgumentOutOfRangeException(nameof(kind)),
        };
        var atoms = ParagraphAtoms.Of(p);
        int o = pos.Offset;
        if (o < 0 || o > atoms.Sum(a => a.Length)) throw new ArgumentOutOfRangeException(nameof(pos), $"Décalage {o} hors du paragraphe.");
        var result = pos.WithOffset(o + 1);
        // au milieu d'un texte : coupé en deux dans le même segment
        var inside = atoms.FirstOrDefault(a => a.Content is TextContent && !a.InDeletion && a.Start < o && o < a.End);
        if (inside is not null)
        {
            var t = (TextContent)inside.Content!;
            var tail = new TextContent { Value = t.Value[(o - inside.Start)..] };
            t.Value = t.Value[..(o - inside.Start)];
            int i = inside.Run!.Content.IndexOf(t);
            inside.Run.Content.InsertRange(i + 1, new[] { content, tail });
            return result;
        }
        var before = atoms.LastOrDefault(a => a.Length > 0 && a.End == o && a.Run is not null && !a.InDeletion);
        if (before is not null) { before.Run!.Content.Insert(before.Run.Content.IndexOf(before.Content!) + 1, content); return result; }
        var after = atoms.FirstOrDefault(a => a.Length > 0 && a.Start == o && a.Run is not null && !a.InDeletion);
        if (after is not null) { after.Run!.Content.Insert(after.Run.Content.IndexOf(after.Content!), content); return result; }
        var run = new Run { Props = MarkProps(p) };
        run.Content.Add(content);
        var anchor = atoms.LastOrDefault(a => a.End <= o && !a.InDeletion);
        if (anchor is null) p.Inlines.Insert(0, run);
        else anchor.Owner.Insert(anchor.Owner.IndexOf(anchor.Inline) + 1, run);
        return result;
    }

    /// <summary>
    /// Insère un texte de plusieurs lignes (collage) : chaque fin de ligne crée un paragraphe, les tabulations et les
    /// retours à la ligne (\v, comme Word en texte brut) deviennent des caractères spéciaux.
    /// </summary>
    public static DocPos InsertMultiline(Story root, DocPos pos, string text, Func<string> newParaId)
    {
        var lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        for (int k = 0; k < lines.Length; k++)
        {
            if (k > 0) pos = Split(root, pos, newParaId);
            var line = lines[k];
            int start = 0;
            for (int i = 0; i <= line.Length; i++)
            {
                if (i < line.Length && line[i] is not ('\t' or '\v' or '\f')) continue;
                // caractères de contrôle restants (hors tabulation et sauts) : ignorés
                var chunk = new string(line[start..i].Where(ch => !char.IsControl(ch)).ToArray());
                pos = Insert(root, pos, chunk);
                if (i < line.Length) pos = InsertSpecial(root, pos, line[i] switch { '\t' => SpecialChar.Tab, '\f' => SpecialChar.PageBreak, _ => SpecialChar.LineBreak });
                start = i + 1;
            }
        }
        return pos;
    }

    /// <summary>
    /// Remplace le texte d'une plage d'un même paragraphe (Rechercher / Remplacer) : le nouveau texte prend la mise en
    /// forme du <strong>premier caractère remplacé</strong>, comme Word (pas celle du caractère qui précède).
    /// Renvoie la position après le nouveau texte.
    /// </summary>
    public static DocPos ReplaceText(Story root, DocPos from, DocPos to, string text)
    {
        if (to < from) (from, to) = (to, from);
        // suivi des modifications : l'ancien texte est barré, le nouveau inséré
        if (TrackContext.Current is not null) return Insert(root, Delete(root, from, to), text);
        bool firstIsText = from.Path.SequenceEqual(to.Path)
            && ParagraphAtoms.Of(from.Paragraph(root)).Any(a => a.Content is TextContent && !a.InDeletion && a.Start <= from.Offset && from.Offset < a.End);
        if (!firstIsText || to.Offset == from.Offset || text.Length == 0)
            return Insert(root, Delete(root, from, to), text);
        int n = to.Offset - from.Offset;
        // inséré après le 1er caractère (il en prend la mise en forme), puis l'ancien texte retiré de part et d'autre
        Insert(root, from.WithOffset(from.Offset + 1), text);
        Delete(root, from, from.WithOffset(from.Offset + 1));
        Delete(root, from.WithOffset(from.Offset + text.Length), from.WithOffset(from.Offset + text.Length + n - 1));
        return from.WithOffset(from.Offset + text.Length);
    }

    static void InsertTextAfter(Run run, RunContent c, string text)
    {
        int i = run.Content.IndexOf(c);
        if (i + 1 < run.Content.Count && run.Content[i + 1] is TextContent next) next.Value = text + next.Value;
        else run.Content.Insert(i + 1, new TextContent { Value = text });
    }

    /// <summary>Mise en forme d'un nouveau segment dans un paragraphe vide : celle de la marque (pPr/rPr) sans ses révisions.</summary>
    static PropertyBag MarkProps(Paragraph p)
    {
        var ns = p.Props.Ns;
        var mark = p.Props.Get("rPr");
        if (mark is null) return new PropertyBag("rPr", ns);
        var copy = new XElement(mark);
        copy.Elements().Where(e => e.Name.Namespace == ns && e.Name.LocalName is "ins" or "del" or "moveFrom" or "moveTo" or "rPrChange").Remove();
        return new PropertyBag("rPr", ns, copy.HasElements ? copy : null);
    }

    // ================================================================ supprimer

    /// <summary>
    /// Supprime le texte entre deux positions de la même suite de blocs. Sur plusieurs paragraphes, les blocs intermédiaires
    /// (paragraphes, tableaux) disparaissent et le premier paragraphe reçoit la fin du dernier ; il garde ses propriétés,
    /// sauf un saut de section du dernier, qui lui est transmis (la fin de section suit la fin du texte).
    /// Les repères (signets…) compris dans la plage sont gardés, regroupés au point de suppression.
    /// </summary>
    public static DocPos Delete(Story root, DocPos from, DocPos to)
    {
        if (to < from) (from, to) = (to, from);
        // suivi des modifications : texte barré, marques de paragraphe supprimées (voir TrackContext.DeleteRange)
        if (TrackContext.Current is { } track)
        {
            track.DeleteRange(root, from, to);
            return from;
        }
        if (!from.ContainerPath.SequenceEqual(to.ContainerPath))
            throw new NotSupportedException("Suppression entre des conteneurs différents (cellules, contrôles) : pas encore prise en charge.");
        var container = DocPos.ResolveContainer(root, from.ContainerPath);
        var first = from.Paragraph(root);
        if (from.BlockIndex == to.BlockIndex)
        {
            DeleteInParagraph(first, from.Offset, to.Offset);
            return from;
        }
        var last = to.Paragraph(root);
        DeleteInParagraph(first, from.Offset, ParagraphAtoms.Length(first));
        DeleteInParagraph(last, 0, to.Offset);
        first.Inlines.AddRange(last.Inlines);
        if (last.Props.Get("sectPr") is { } sect) first.Props.Set(new XElement(sect));
        else first.Props.Remove("sectPr");
        container.Blocks.RemoveRange(from.BlockIndex + 1, to.BlockIndex - from.BlockIndex);
        return from;
    }

    /// <summary>
    /// Plage qui sort d'un tableau (ou d'un contrôle de contenu) au niveau du corps : bornes ramenées au premier niveau.
    /// Une borne dans un bloc composé désigne ce bloc <strong>entier</strong> (offset null). Null si la plage reste à
    /// l'intérieur d'un même bloc composé (deux cellules d'un tableau : pas encore pris en charge).
    /// </summary>
    public static (int Start, int? StartOffset, int End, int? EndOffset)? AcrossTopLevel(DocPos from, DocPos to)
    {
        if (to < from) (from, to) = (to, from);
        if (from.Path[0] == to.Path[0]) return null;
        return (from.Path[0], from.Path.Count == 1 ? from.Offset : null, to.Path[0], to.Path.Count == 1 ? to.Offset : null);
    }

    /// <summary>
    /// Suppression d'une plage qui traverse des tableaux au niveau du corps : les blocs composés touchés partent entiers
    /// (comme les rangées sélectionnées dans Word), le texte des paragraphes de bord est coupé ; les paragraphes de part et
    /// d'autre d'un tableau supprimé ne sont pas fusionnés. Renvoie la position du point d'insertion après.
    /// </summary>
    public static DocPos DeleteAcross(Story root, DocPos from, DocPos to)
    {
        if (to < from) (from, to) = (to, from);
        if (TrackContext.Current is not null && !from.ContainerPath.SequenceEqual(to.ContainerPath))
            throw new NotSupportedException("Suppression d'un tableau avec le suivi des modifications : pas encore prise en charge.");
        if (from.ContainerPath.SequenceEqual(to.ContainerPath)) return Delete(root, from, to);
        if (AcrossTopLevel(from, to) is not var (s, so, e, eo))
            throw new NotSupportedException("Sélection entre deux cellules d'un tableau : pas encore prise en charge.");
        var blocks = root.Blocks;
        if (so is { } a)
        {
            // début dans un paragraphe : fin de ce paragraphe supprimée, puis les blocs jusqu'au tableau compris
            var p = (Paragraph)blocks[s];
            DeleteInParagraph(p, a, ParagraphAtoms.Length(p));
            int removeEnd = eo is null ? e : e - 1;
            if (eo is { } b && blocks[e] is Paragraph last) DeleteInParagraph(last, 0, b);
            blocks.RemoveRange(s + 1, removeEnd - s);
            return new DocPos(new[] { s }, a);
        }
        if (eo is { } bo)
        {
            // début dans un tableau, fin dans un paragraphe : blocs jusqu'à ce paragraphe, début du paragraphe supprimés
            if (blocks[e] is Paragraph last) DeleteInParagraph(last, 0, bo);
            blocks.RemoveRange(s, e - s);
            return new DocPos(new[] { s }, 0);
        }
        // deux tableaux (ou contrôles) : tout part ; un paragraphe vide reçoit le point d'insertion s'il n'y en a pas
        blocks.RemoveRange(s, e - s + 1);
        if (s >= blocks.Count || blocks[s] is not Paragraph)
        {
            var ns = root.Paragraphs().FirstOrDefault()?.Props.Ns ?? PropertyBag.TransitionalNs;
            blocks.Insert(s, new Paragraph { Props = new PropertyBag("pPr", ns) });
        }
        return new DocPos(new[] { s }, 0);
    }

    static void DeleteInParagraph(Paragraph p, int a, int b)
    {
        if (b <= a) return;
        var emptied = new HashSet<Run>();
        foreach (var atom in ParagraphAtoms.Of(p))
        {
            if (atom.Length == 0 || atom.End <= a || atom.Start >= b) continue;
            int s = Math.Max(a, atom.Start) - atom.Start, e = Math.Min(b, atom.End) - atom.Start;
            if (atom.Content is TextContent t && (s > 0 || e < t.Value.Length)) { t.Value = t.Value.Remove(s, e - s); continue; }
            if (atom.Content is FieldCode fc && (s > 0 || e < fc.Value.Length)) { fc.Value = fc.Value.Remove(s, e - s); continue; }
            if (atom.Content is not null)
            {
                atom.Run!.Content.Remove(atom.Content);
                emptied.Add(atom.Run);
            }
            else atom.Owner.Remove(atom.Inline);          // îlot de largeur 1 (équation…)
        }
        // segments vidés par la suppression (un segment déjà vide à l'origine reste)
        foreach (var run in emptied.Where(r => r.Content.Count == 0)) RemoveInline(p.Inlines, run);
        RemoveEmptyContainers(p.Inlines);
    }

    static bool RemoveInline(List<Inline> list, Inline target)
    {
        if (list.Remove(target)) return true;
        foreach (var i in list)
            if (i is InlineContainer box && RemoveInline(box.Children, target)) return true;
            else if (i is InlineContentControl s && RemoveInline(s.Content, target)) return true;
        return false;
    }

    /// <summary>Liens et révisions devenus vides après une suppression (un lien sans texte n'a plus de sens).</summary>
    static void RemoveEmptyContainers(List<Inline> list)
    {
        foreach (var box in list.OfType<InlineContainer>().ToList())
        {
            RemoveEmptyContainers(box.Children);
            if (box.Children.Count == 0 && box.Name.LocalName is "hyperlink" or "ins" or "del" or "moveFrom" or "moveTo" or "smartTag") list.Remove(box);
        }
    }

    // ================================================================ scinder (touche Entrée)

    /// <summary>
    /// Scinde le paragraphe à une position (Entrée). Les deux paragraphes gardent les propriétés ; le second reçoit de
    /// nouveaux w14:paraId/w14:textId ; un saut de section reste sur le second ; un lien coupé en deux est dupliqué.
    /// Renvoie le début du second paragraphe.
    /// <para>
    /// En fin de paragraphe (Entrée après le dernier caractère), comme Word : le nouveau paragraphe prend le style
    /// suivant (<paramref name="nextStyle"/> : w:next du style, « Titre 1 » → « Normal ») ; s'il garde le même style,
    /// sa marque reprend la mise en forme du dernier caractère, et la frappe continue en gras après un texte en gras.
    /// </para>
    /// </summary>
    public static DocPos Split(Story root, DocPos pos, Func<string> newParaId, Func<string?, string?>? nextStyle = null)
    {
        var container = DocPos.ResolveContainer(root, pos.ContainerPath);
        var p = pos.Paragraph(root);
        var p2 = new Paragraph { Props = p.Props.Clone() };
        foreach (var a in p.Attributes)
            p2.Attributes.Add(a.Name == FlowEdits.W14 + "paraId" || a.Name == FlowEdits.W14 + "textId" ? new XAttribute(a.Name, newParaId()) : new XAttribute(a));
        p.Props.Remove("sectPr");
        int o = pos.Offset;
        var (left, right) = SplitInlines(p.Inlines, ref o);
        p.Inlines.Clear();
        p.Inlines.AddRange(left);
        p2.Inlines.AddRange(right);
        if (pos.Offset >= ParagraphAtoms.Length(p) && ParagraphAtoms.Length(p2) == 0)
        {
            var style = p.Props.Val("pStyle");
            var next = nextStyle?.Invoke(style);
            if (next is not null && next != style)
            {
                p2.Props.SetVal("pStyle", next);
                p2.Props.Remove("rPr");
            }
            else if (LastRun(left) is { } last)
            {
                var mark = last.Props.Element is { } rpr ? new XElement(rpr) : new XElement(p.Props.Ns + "rPr");
                mark.Elements().Where(e => e.Name.LocalName is "ins" or "del" or "moveFrom" or "moveTo" or "rPrChange").Remove();
                p2.Props.Remove("rPr");
                if (mark.HasElements) p2.Props.Set(mark);
            }
        }
        container.Blocks.Insert(pos.BlockIndex + 1, p2);
        // suivi des modifications : la marque qui termine le premier paragraphe est nouvelle
        TrackContext.Current?.MarkParagraphInserted(p);
        return new DocPos(pos.ContainerPath.Append(pos.BlockIndex + 1).ToArray(), 0);
    }

    /// <summary>
    /// Coupe une liste d'éléments au décalage <paramref name="o"/>, consommé au fil de la liste (-1 une fois la coupe faite).
    /// Règle unique : un élément va à gauche si la coupe est à sa fin ou au-delà (o ≥ longueur) ; un repère sans largeur
    /// placé au point de coupe reste donc à gauche, un caractère qui commence au point de coupe part à droite.
    /// </summary>
    static (List<Inline> Left, List<Inline> Right) SplitInlines(List<Inline> list, ref int o)
    {
        var left = new List<Inline>();
        var right = new List<Inline>();
        foreach (var i in list)
        {
            int len = InlineLength(i);
            if (o < 0) { right.Add(i); continue; }
            if (o >= len) { left.Add(i); o -= len; continue; }
            switch (i)
            {
                case Run r:
                {
                    var keep = new List<RunContent>();
                    var moved = new List<RunContent>();
                    foreach (var c in r.Content)
                    {
                        int cl = ParagraphAtoms.LengthOf(c);
                        if (o < 0) moved.Add(c);
                        else if (o >= cl) { keep.Add(c); o -= cl; }
                        else if (o == 0) { moved.Add(c); o = -1; }
                        else
                        {
                            // au milieu d'un texte ou d'un code de champ (les autres contenus font 1 caractère)
                            var (head, tail) = CutContent(c, o);
                            keep.Add(head);
                            moved.Add(tail);
                            o = -1;
                        }
                    }
                    if (keep.Count == 0) { right.Add(r); break; }
                    var r2 = new Run { Props = r.Props.Clone() };
                    r2.Attributes.AddRange(r.Attributes.Select(a => new XAttribute(a)));
                    r.Content.Clear();
                    r.Content.AddRange(keep);
                    r2.Content.AddRange(moved);
                    left.Add(r);
                    right.Add(r2);
                    break;
                }
                case InlineContainer box:
                {
                    var (l, rr) = SplitInlines(box.Children, ref o);
                    if (l.Count == 0) { right.Add(box); break; }
                    var copy = new InlineContainer(box.Name);
                    copy.Attributes.AddRange(box.Attributes.Select(a => new XAttribute(a)));
                    copy.Head.AddRange(box.Head.Select(h => new XElement(h)));
                    box.Children.Clear();
                    box.Children.AddRange(l);
                    copy.Children.AddRange(rr);
                    left.Add(box);
                    right.Add(copy);
                    break;
                }
                default:
                    // contrôle de contenu ou îlot coupé au milieu : il reste entier, à droite
                    right.Add(i);
                    o = -1;
                    break;
            }
            o = -1;
        }
        return (left, right);
    }

    /// <summary>Dernier segment qui porte du texte (ou une tabulation…), en descendant dans les liens et révisions.</summary>
    static Run? LastRun(List<Inline> list)
    {
        for (int i = list.Count - 1; i >= 0; i--)
        {
            switch (list[i])
            {
                case Run r when r.Content.Any(c => c is TextContent or TabContent or BreakContent): return r;
                case InlineContainer box when LastRun(box.Children) is { } inBox: return inBox;
                case InlineContentControl sdt when LastRun(sdt.Content) is { } inSdt: return inSdt;
            }
        }
        return null;
    }

    static (RunContent Head, RunContent Tail) CutContent(RunContent c, int at)
    {
        RunContent tail;
        switch (c)
        {
            case TextContent t:
                tail = new TextContent { Value = t.Value[at..], Deleted = t.Deleted };
                t.Value = t.Value[..at];
                break;
            case FieldCode f:
                tail = new FieldCode { Value = f.Value[at..], Deleted = f.Deleted };
                f.Value = f.Value[..at];
                break;
            default: throw new InvalidOperationException(c.GetType().Name);
        }
        tail.Attributes.AddRange(c.Attributes.Select(a => new XAttribute(a)));
        return (c, tail);
    }

    static int InlineLength(Inline i) => i switch
    {
        Run r => r.Content.Sum(ParagraphAtoms.LengthOf),
        InlineContainer b => b.Children.Sum(InlineLength),
        InlineContentControl s => s.Content.Sum(InlineLength),
        _ => ParagraphAtoms.LengthOf(i),
    };

    // ================================================================ annulables

    /// <summary>
    /// Exécute une opération dans une transaction du journal : les blocs <paramref name="index"/>..+<paramref name="count"/>
    /// du conteneur sont copiés avant et après. Si l'opération échoue, le modèle est remis dans son état d'avant.
    /// </summary>
    public static T Edit<T>(UndoJournal journal, Story root, string label, IReadOnlyList<int> containerPath, int index, int count,
        DocPos? caretBefore, Func<T> op, Func<T, DocPos?> caretAfter, string? mergeKey = null)
    {
        using var tx = journal.Begin(label, root, containerPath, index, count, caretBefore, mergeKey);
        var result = op();
        tx.Commit(caretAfter(result));
        return result;
    }

    /// <summary>Frappe annulable ; la frappe continue est regroupée par mot, comme dans Word.</summary>
    public static DocPos Type(UndoJournal journal, Story root, DocPos pos, string text) =>
        Edit(journal, root, "Frappe", pos.ContainerPath, pos.BlockIndex, 1, pos, () => Insert(root, pos, text), p => p,
             mergeKey: journal.TypingKey(pos, text));

    public static DocPos DeleteRange(UndoJournal journal, Story root, DocPos from, DocPos to)
    {
        if (to < from) (from, to) = (to, from);
        return Edit(journal, root, "Supprimer", from.ContainerPath, from.BlockIndex, to.BlockIndex - from.BlockIndex + 1, to,
                    () => Delete(root, from, to), p => p);
    }

    public static DocPos SplitParagraph(UndoJournal journal, Story root, DocPos pos, Func<string> newParaId, Func<string?, string?>? nextStyle = null) =>
        Edit(journal, root, "Nouveau paragraphe", pos.ContainerPath, pos.BlockIndex, 1, pos, () => Split(root, pos, newParaId, nextStyle), p => p);

    /// <summary>
    /// Remplace la plage [<paramref name="from"/>, <paramref name="to"/>] (éventuellement vide) par un texte de
    /// plusieurs lignes, en une seule opération annulable (frappe sur une sélection, collage).
    /// </summary>
    public static DocPos Replace(UndoJournal journal, Story root, string label, DocPos from, DocPos to, string text, Func<string> newParaId)
    {
        if (to < from) (from, to) = (to, from);
        if (!from.ContainerPath.SequenceEqual(to.ContainerPath) && AcrossTopLevel(from, to) is var (s, _, e, _))
            return Edit(journal, root, label, Array.Empty<int>(), s, e - s + 1, to,
                () => InsertMultiline(root, DeleteAcross(root, from, to), text, newParaId), p => p);
        return Edit(journal, root, label, from.ContainerPath, from.BlockIndex, to.BlockIndex - from.BlockIndex + 1, to,
            () => InsertMultiline(root, from.Equals(to) ? from : Delete(root, from, to), text, newParaId), p => p);
    }

    public static DocPos InsertSpecial(UndoJournal journal, Story root, DocPos pos, SpecialChar kind) =>
        Edit(journal, root, kind switch { SpecialChar.Tab => "Tabulation", SpecialChar.PageBreak => "Saut de page", SpecialChar.ColumnBreak => "Saut de colonne", _ => "Saut de ligne" },
             pos.ContainerPath, pos.BlockIndex, 1, pos, () => InsertSpecial(root, pos, kind), p => p);
}

/// <summary>Caractères spéciaux insérés au clavier : Tab, Maj+Entrée, Ctrl+Entrée, Ctrl+Maj+Entrée.</summary>
public enum SpecialChar { Tab, LineBreak, PageBreak, ColumnBreak }
