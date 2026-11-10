using System.Xml.Linq;
using OpenWrite.Core.Model;

namespace OpenWrite.Core.Editing;

/// <summary>
/// Mise en forme directe (lot 3) : propriétés de caractère (w:rPr des segments) sur une plage, propriétés de paragraphe
/// (w:pPr) sur les paragraphes d'une plage. Les segments sont coupés aux bornes de la plage pour que la propriété ne
/// s'applique qu'au texte visé ; les attributs et le reste de la mise en forme de chaque segment sont gardés.
/// Comme les autres opérations, elles s'exécutent dans une transaction du journal pour être annulables (<see cref="Apply"/>).
/// </summary>
public static class FormatOps
{
    /// <summary>
    /// Applique <paramref name="apply"/> aux propriétés de chaque segment compris dans [<paramref name="from"/>,
    /// <paramref name="to"/>[ (même conteneur ; les tableaux traversés sont pris en entier). Renvoie le nombre de segments touchés.
    /// </summary>
    public static int ApplyToRuns(Story root, DocPos from, DocPos to, Action<Paragraph, Run> apply)
    {
        if (to < from) (from, to) = (to, from);
        (from, to) = Normalize(from, to);
        if (!from.ContainerPath.SequenceEqual(to.ContainerPath))
            throw new NotSupportedException("Plage entre des conteneurs différents (cellules, contrôles) : pas encore prise en charge.");
        var container = DocPos.ResolveContainer(root, from.ContainerPath);
        int count = 0;
        for (int b = from.BlockIndex; b <= to.BlockIndex; b++)
        {
            var block = container.Blocks[b];
            if (block is not Paragraph p)
            {
                // tableau ou contrôle de contenu traversé : tout son texte
                foreach (var inner in new Story { Blocks = { block } }.Paragraphs())
                    foreach (var r in RunsIn(inner.Inlines)) { apply(inner, r); count++; }
                continue;
            }
            int a = b == from.BlockIndex ? from.Offset : 0;
            int e = b == to.BlockIndex ? to.Offset : ParagraphAtoms.Length(p);
            if (e <= a) continue;
            SplitRunAt(p, e);
            SplitRunAt(p, a);
            foreach (var atom in ParagraphAtoms.Of(p))
            {
                if (atom.Run is not { } run || atom.Length == 0 || atom.Start < a || atom.End > e) continue;
                if (atom.Content != run.Content.FirstOrDefault(c => ParagraphAtoms.LengthOf(c) > 0)) continue;   // une fois par segment
                apply(p, run);
                count++;
            }
        }
        return count;
    }

    /// <summary>Applique <paramref name="apply"/> aux propriétés de chaque paragraphe de la plage (tableaux traversés compris).</summary>
    public static void ApplyToParagraphs(Story root, DocPos from, DocPos to, Action<Paragraph> apply)
    {
        if (to < from) (from, to) = (to, from);
        (from, to) = Normalize(from, to);
        if (!from.ContainerPath.SequenceEqual(to.ContainerPath))
            throw new NotSupportedException("Plage entre des conteneurs différents (cellules, contrôles) : pas encore prise en charge.");
        var container = DocPos.ResolveContainer(root, from.ContainerPath);
        for (int b = from.BlockIndex; b <= to.BlockIndex; b++)
            foreach (var p in new Story { Blocks = { container.Blocks[b] } }.Paragraphs()) apply(p);
    }

    /// <summary>Exécute une mise en forme dans une transaction du journal (blocs de la plage), en gardant la sélection.</summary>
    public static void Apply(UndoJournal journal, Story root, string label, DocPos from, DocPos to, Action op)
    {
        if (to < from) (from, to) = (to, from);
        var original = to;
        (from, to) = Normalize(from, to);
        TextOps.Edit(journal, root, label, from.ContainerPath, from.BlockIndex, to.BlockIndex - from.BlockIndex + 1, original,
            () => { op(); return 0; }, _ => original);
    }

    /// <summary>
    /// Plage qui traverse un tableau : bornes ramenées au premier niveau (un tableau touché est pris en entier, voir
    /// <see cref="TextOps.AcrossTopLevel"/>) ; entre deux cellules d'un même tableau (sélection de cellules, lot 6) : le
    /// bloc de premier niveau qui le contient ; inchangée dans un même conteneur.
    /// </summary>
    public static (DocPos From, DocPos To) Normalize(DocPos from, DocPos to)
    {
        if (from.ContainerPath.SequenceEqual(to.ContainerPath)) return (from, to);
        if (TextOps.AcrossTopLevel(from, to) is not var (s, so, e, eo))
        {
            var top = new DocPos(new[] { from.Path[0] }, 0);
            return (top, top);
        }
        return (new DocPos(new[] { s }, so ?? 0), new DocPos(new[] { e }, eo ?? 0));
    }

    // ------------------------------------------------------------------ propriétés courantes

    /// <summary>
    /// Active ou retire une propriété booléenne (b, i, strike…) et sa variante « script complexe » (bCs, iCs). Retirée
    /// mais toujours active par le style (<paramref name="inherited"/>) : posée à « 0 », comme Word.
    /// </summary>
    public static void SetToggle(Run run, string local, bool on, Func<Run, bool>? inherited = null)
    {
        var props = run.Props;
        string? cs = local switch { "b" => "bCs", "i" => "iCs", _ => null };
        if (on)
        {
            props.Set(new XElement(props.Ns + local));
            if (cs is not null) props.Set(new XElement(props.Ns + cs));
            return;
        }
        props.Remove(local);
        if (cs is not null) props.Remove(cs);
        if (inherited?.Invoke(run) == true)
        {
            props.Set(new XElement(props.Ns + local, new XAttribute(props.Ns + "val", "0")));
            if (cs is not null) props.Set(new XElement(props.Ns + cs, new XAttribute(props.Ns + "val", "0")));
        }
    }

    /// <summary>Soulignement simple ou aucun (w:u).</summary>
    public static void SetUnderline(Run run, bool on, Func<Run, bool>? inherited = null)
    {
        var props = run.Props;
        if (on) { props.SetVal("u", "single"); return; }
        props.Remove("u");
        if (inherited?.Invoke(run) == true) props.SetVal("u", "none");
    }

    /// <summary>Valeur d'une propriété à attribut w:val (sz, color, highlight, vertAlign…) ; null la retire.</summary>
    public static void SetValue(Run run, string local, string? value)
    {
        run.Props.SetVal(local, value);
        if (local == "sz") run.Props.SetVal("szCs", value);
    }

    /// <summary>Police du thème (titres : major, corps : minor) pour l'alphabet latin et les autres écritures.</summary>
    public static void SetThemeFont(Run run, bool major)
    {
        var ns = run.Props.Ns;
        var kind = major ? "major" : "minor";
        var fonts = run.Props.Get("rFonts") is { } existing ? new XElement(existing) : new XElement(ns + "rFonts");
        foreach (var a in new[] { "ascii", "hAnsi", "cs", "eastAsia" }) fonts.Attribute(ns + a)?.Remove();
        fonts.SetAttributeValue(ns + "asciiTheme", kind + "HAnsi");
        fonts.SetAttributeValue(ns + "hAnsiTheme", kind + "HAnsi");
        fonts.SetAttributeValue(ns + "eastAsiaTheme", kind + "EastAsia");
        fonts.SetAttributeValue(ns + "cstheme", kind + "Bidi");
        run.Props.Set(fonts);
    }

    /// <summary>Police (w:rFonts : ascii, hAnsi, cs ; les polices de thème cèdent la place).</summary>
    public static void SetFont(Run run, string family)
    {
        var ns = run.Props.Ns;
        var fonts = run.Props.Get("rFonts") is { } existing ? new XElement(existing) : new XElement(ns + "rFonts");
        foreach (var a in new[] { "asciiTheme", "hAnsiTheme", "cstheme" }) fonts.Attribute(ns + a)?.Remove();
        foreach (var a in new[] { "ascii", "hAnsi", "cs" }) fonts.SetAttributeValue(ns + a, family);
        run.Props.Set(fonts);
    }

    /// <summary>
    /// Ctrl+Espace : retire la mise en forme directe des caractères (le style de caractère, la langue et l'historique
    /// des révisions restent).
    /// </summary>
    public static void ClearCharacter(Run run)
    {
        if (run.Props.Element is not { } rpr) return;
        rpr.Elements().Where(e => e.Name.Namespace != run.Props.Ns
                                  || e.Name.LocalName is not ("rStyle" or "lang" or "noProof" or "ins" or "del" or "moveFrom" or "moveTo" or "rPrChange"))
           .Remove();
    }

    /// <summary>
    /// Ctrl+Q : retire la mise en forme directe du paragraphe (style, liste, saut de section, marque de paragraphe,
    /// cadre et historique des révisions restent).
    /// </summary>
    public static void ClearParagraph(Paragraph p)
    {
        if (p.Props.Element is not { } ppr) return;
        ppr.Elements().Where(e => e.Name.Namespace != p.Props.Ns
                                  || e.Name.LocalName is not ("pStyle" or "numPr" or "sectPr" or "rPr" or "pPrChange" or "framePr" or "cnfStyle" or "divId"))
           .Remove();
    }

    /// <summary>Retraits (twips ; null = inchangé). <paramref name="firstLine"/> négatif = retrait suspendu (w:hanging).</summary>
    public static void SetIndent(Paragraph p, int? left = null, int? right = null, int? firstLine = null)
    {
        var ns = p.Props.Ns;
        var ind = p.Props.Get("ind") is { } existing ? new XElement(existing) : new XElement(ns + "ind");
        void Set(string name, string alt, int value)
        {
            ind.Attribute(ns + alt)?.Remove();
            ind.Attribute(ns + (alt + "Chars"))?.Remove();
            ind.Attribute(ns + (name + "Chars"))?.Remove();
            ind.SetAttributeValue(ns + name, value);
        }
        if (left is { } l) Set("left", "start", l);
        if (right is { } r) Set("right", "end", r);
        if (firstLine is { } f)
        {
            foreach (var a in new[] { "firstLine", "hanging", "firstLineChars", "hangingChars" }) ind.Attribute(ns + a)?.Remove();
            ind.SetAttributeValue(ns + (f < 0 ? "hanging" : "firstLine"), Math.Abs(f));
        }
        p.Props.Set(ind);
    }

    /// <summary>Espacements avant / après (twips) et interligne (w:line, w:lineRule) ; null = inchangé.</summary>
    public static void SetSpacing(Paragraph p, int? before = null, int? after = null, int? line = null, string? lineRule = null)
    {
        var ns = p.Props.Ns;
        var sp = p.Props.Get("spacing") is { } existing ? new XElement(existing) : new XElement(ns + "spacing");
        if (before is { } b)
        {
            foreach (var a in new[] { "beforeLines", "beforeAutospacing" }) sp.Attribute(ns + a)?.Remove();
            sp.SetAttributeValue(ns + "before", b);
        }
        if (after is { } af)
        {
            foreach (var a in new[] { "afterLines", "afterAutospacing" }) sp.Attribute(ns + a)?.Remove();
            sp.SetAttributeValue(ns + "after", af);
        }
        if (line is { } ln)
        {
            sp.SetAttributeValue(ns + "line", ln);
            sp.SetAttributeValue(ns + "lineRule", lineRule ?? "auto");
        }
        p.Props.Set(sp);
    }

    /// <summary>Option de paragraphe (keepNext, keepLines, pageBreakBefore, widowControl…) ; « 0 » si retirée mais héritée.</summary>
    public static void SetParagraphFlag(Paragraph p, string local, bool on, bool inherited)
    {
        if (on) { p.Props.Set(new XElement(p.Props.Ns + local)); return; }
        p.Props.Remove(local);
        if (inherited) p.Props.Set(new XElement(p.Props.Ns + local, new XAttribute(p.Props.Ns + "val", "0")));
    }

    // ------------------------------------------------------------------ segments

    static IEnumerable<Run> RunsIn(List<Inline> list)
    {
        foreach (var i in list)
            switch (i)
            {
                case Run r: yield return r; break;
                case InlineContainer box: foreach (var r in RunsIn(box.Children)) yield return r; break;
                case InlineContentControl s: foreach (var r in RunsIn(s.Content)) yield return r; break;
            }
    }

    /// <summary>
    /// Coupe le segment qui contient la position <paramref name="offset"/> en son milieu (rien si la position est déjà
    /// entre deux segments) : deux segments de même mise en forme, le texte coupé en deux.
    /// </summary>
    public static void SplitRunAt(Paragraph p, int offset)
    {
        foreach (var atom in ParagraphAtoms.Of(p))
        {
            if (atom.Run is not { } run || atom.Length == 0) continue;
            // décalage du début du segment
            int runStart = atom.Start - run.Content.TakeWhile(c => c != atom.Content).Sum(ParagraphAtoms.LengthOf);
            int runEnd = runStart + run.Content.Sum(ParagraphAtoms.LengthOf);
            if (offset <= runStart || offset >= runEnd) continue;
            var left = new List<RunContent>();
            var right = new List<RunContent>();
            int pos = runStart;
            foreach (var c in run.Content)
            {
                int len = ParagraphAtoms.LengthOf(c);
                if (pos + len <= offset) left.Add(c);
                else if (pos >= offset) right.Add(c);
                else
                {
                    // au milieu d'un texte (ou d'un code de champ)
                    int cut = offset - pos;
                    switch (c)
                    {
                        case TextContent t:
                            var tail = new TextContent { Value = t.Value[cut..], Deleted = t.Deleted };
                            tail.Attributes.AddRange(t.Attributes.Select(a => new XAttribute(a)));
                            t.Value = t.Value[..cut];
                            left.Add(t);
                            right.Add(tail);
                            break;
                        case FieldCode f:
                            var ftail = new FieldCode { Value = f.Value[cut..], Deleted = f.Deleted };
                            ftail.Attributes.AddRange(f.Attributes.Select(a => new XAttribute(a)));
                            f.Value = f.Value[..cut];
                            left.Add(f);
                            right.Add(ftail);
                            break;
                        default:
                            right.Add(c);
                            break;
                    }
                }
                pos += len;
            }
            var r2 = new Run { Props = run.Props.Clone() };
            r2.Attributes.AddRange(run.Attributes.Select(a => new XAttribute(a)));
            r2.Content.AddRange(right);
            run.Content.Clear();
            run.Content.AddRange(left);
            atom.Owner.Insert(atom.Owner.IndexOf(run) + 1, r2);
            return;
        }
    }
}
