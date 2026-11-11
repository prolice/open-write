using System.Globalization;
using System.Xml.Linq;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;

namespace OpenWrite.Docx;

/// <summary>Un commentaire de comments.xml, avec ce qu'en dit commentsExtended.xml (réponse, résolu).</summary>
public sealed record CommentInfo(string Id, string Author, string? Initials, DateTime? Date, StoryHolder Holder, string? ParaId, string? ParentParaId, bool Done)
{
    /// <summary>Texte du commentaire (paragraphes séparés par un retour à la ligne).</summary>
    public string Text => string.Join("\n", Holder.Story.Paragraphs().Select(p =>
        string.Concat(p.Runs().SelectMany(r => r.Content).Select(c => c switch { TextContent { Deleted: false } t => t.Value, TabContent => "\t", BreakContent => "\n", _ => "" }))));

    /// <summary>Le commentaire n'a que du texte simple (une modification du texte ne perd aucune mise en forme).</summary>
    public bool PlainText => Holder.Story.Blocks.All(b => b is Paragraph) && Holder.Story.Paragraphs().SelectMany(p => p.Runs())
        .All(r => r.Content.All(c => c is TextContent or TabContent or BreakContent || c is OpaqueRunContent o && o.Xml.Name.LocalName == "annotationRef"));
}

/// <summary>Place d'un commentaire dans le texte : de sa marque de début à sa marque de fin (décalages du modèle).</summary>
public sealed record CommentAnchor(Paragraph StartParagraph, int Start, Paragraph EndParagraph, int End);

/// <summary>
/// Commentaires du document (lot 8) : lecture (auteur, date, texte, réponses et état résolu de Word 2013+), place dans le
/// texte, ajout, réponse, résolution, modification du texte et suppression — chaque opération dans une transaction du
/// journal (marques du corps, liste des commentaires, texte, commentsExtended.xml), donc annulable. Les commentaires sont
/// ancrés dans le corps (Word les permet aussi dans les notes et les zones de texte : lus, pas encore ajoutés là).
/// </summary>
public sealed class CommentStore
{
    readonly DocxPackage _pkg;

    public CommentStore(DocxPackage pkg) => _pkg = pkg;

    static readonly XNamespace W15 = DocxPackage.W15;
    static readonly XNamespace W14 = FlowEdits.W14;

    StoryPart? Part => _pkg.Stories.Values.FirstOrDefault(s => s.Kind == "comments");

    static string? Attr(StoryHolder h, string local) => h.Attributes.FirstOrDefault(a => a.Name.LocalName == local)?.Value;

    static string? LastParaId(StoryHolder h) =>
        h.Story.Paragraphs().LastOrDefault()?.Attributes.FirstOrDefault(a => a.Name == W14 + "paraId")?.Value;

    XElement? Extended(bool create = false) => _pkg.XmlPartOf("commentsExtended", create)?.Doc.Root;

    /// <summary>Tous les commentaires, dans l'ordre de comments.xml.</summary>
    public List<CommentInfo> All()
    {
        var res = new List<CommentInfo>();
        if (Part is not { } part) return res;
        var ext = Extended()?.Elements(W15 + "commentEx").ToDictionary(e => (string?)e.Attribute(W15 + "paraId") ?? "", e => e) ?? new();
        foreach (var h in part.Items.OfType<StoryHolder>())
        {
            var id = Attr(h, "id");
            if (id is null) continue;
            var paraId = LastParaId(h);
            ext.TryGetValue(paraId ?? "-", out var e);
            DateTime? date = DateTime.TryParse(Attr(h, "date"), CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out var d) ? d : null;
            res.Add(new CommentInfo(id, Attr(h, "author") ?? "", Attr(h, "initials"), date, h, paraId, (string?)e?.Attribute(W15 + "paraIdParent"),
                (string?)e?.Attribute(W15 + "done") is "1" or "true"));
        }
        return res;
    }

    // ================================================================ place dans le texte

    static bool Is(Inline i, string local, string id) =>
        i is OpaqueInline o && o.Xml.Name.LocalName == local && (string?)o.Xml.Attributes().FirstOrDefault(a => a.Name.LocalName == "id") == id;

    static bool IsReference(RunContent c, string id) =>
        c is OpaqueRunContent o && o.Xml.Name.LocalName == "commentReference" && (string?)o.Xml.Attributes().FirstOrDefault(a => a.Name.LocalName == "id") == id;

    /// <summary>Marques d'un commentaire dans une suite de blocs ; sans marques de plage, l'appel du commentaire.</summary>
    public static CommentAnchor? Anchor(Story story, string id)
    {
        Paragraph? sp = null, ep = null, rp = null;
        int s = 0, e = 0, r = 0;
        foreach (var p in story.Paragraphs())
        {
            int pos = 0;
            void Walk(List<Inline> list)
            {
                foreach (var i in list)
                {
                    if (Is(i, "commentRangeStart", id)) { sp = p; s = pos; }
                    else if (Is(i, "commentRangeEnd", id)) { ep = p; e = pos; }
                    switch (i)
                    {
                        case Run run:
                            foreach (var c in run.Content)
                            {
                                if (IsReference(c, id)) { rp = p; r = pos; }
                                pos += ParagraphAtoms.LengthOf(c);
                            }
                            break;
                        case InlineContainer b: Walk(b.Children); break;
                        case InlineContentControl sdt: Walk(sdt.Content); break;
                        default: pos += ParagraphAtoms.LengthOf(i); break;
                    }
                }
            }
            Walk(p.Inlines);
        }
        if (sp is not null && ep is not null) return new CommentAnchor(sp, s, ep, e);
        if (sp is not null && rp is not null) return new CommentAnchor(sp, s, rp, r);
        if (rp is not null) return new CommentAnchor(rp, r, rp, r);
        return null;
    }

    /// <summary>Insère un élément sans largeur à un décalage : avant le caractère qui y commence (début), après celui qui y finit (fin).</summary>
    static (List<Inline> Owner, int Index) Place(Paragraph p, int offset, bool start)
    {
        FormatOps.SplitRunAt(p, offset);
        var atoms = ParagraphAtoms.Of(p);
        if (start)
        {
            var a = atoms.FirstOrDefault(x => x.Length > 0 && x.Start == offset);
            if (a is not null) return (a.Owner, a.Owner.IndexOf(a.Inline));
            return (p.Inlines, p.Inlines.Count);
        }
        var b = atoms.LastOrDefault(x => x.Length > 0 && x.End == offset);
        if (b is not null) return (b.Owner, b.Owner.IndexOf(b.Inline) + 1);
        return (p.Inlines, 0);
    }

    /// <summary>Position juste après un élément en ligne (marque de fin, appel) dans sa liste propriétaire.</summary>
    static (List<Inline> Owner, int Index)? After(Story story, Func<Inline, bool> match)
    {
        foreach (var p in story.Paragraphs())
            if (Find(p.Inlines, match) is { } f) return (f.Owner, f.Index + 1);
        return null;

        static (List<Inline> Owner, int Index)? Find(List<Inline> list, Func<Inline, bool> match)
        {
            for (int i = 0; i < list.Count; i++)
            {
                if (match(list[i])) return (list, i);
                var inner = list[i] switch { InlineContainer b => Find(b.Children, match), InlineContentControl s => Find(s.Content, match), _ => null };
                if (inner is not null) return inner;
            }
            return null;
        }
    }

    // ================================================================ créer

    /// <summary>Styles intégrés « Commentaire » et « Marque de commentaire » (identifiants), fournis par l'appelant (styles.xml).</summary>
    public Func<string, string?>? EnsureStyle { get; set; }

    /// <summary>Prochain identifiant de commentaire : un de plus que le plus grand des commentaires et des révisions.</summary>
    int NextId()
    {
        int max = -1;
        foreach (var c in All()) if (int.TryParse(c.Id, out var n) && n > max) max = n;
        foreach (var s in _pkg.Stories.Values.SelectMany(p => p.Stories()))
            foreach (var n in Core.Review.Revisions.AnnotationIds(s)) if (n > max) max = n;
        return max + 1;
    }

    static string Now() => DateTime.Now.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);

    /// <summary>Contenu d'un commentaire : une marque de commentaire, puis le texte (une ligne par paragraphe).</summary>
    List<Block> Body(XNamespace w, string text, Func<string> newParaId, string? lastParaId = null)
    {
        var textStyle = EnsureStyle?.Invoke("annotation text");
        var refStyle = EnsureStyle?.Invoke("annotation reference");
        var lines = text.Replace("\r\n", "\n").Split('\n');
        var blocks = new List<Block>();
        for (int k = 0; k < lines.Length; k++)
        {
            var p = new Paragraph { Props = new PropertyBag("pPr", w) };
            p.Attributes.Add(new XAttribute(W14 + "paraId", k == lines.Length - 1 && lastParaId is not null ? lastParaId : newParaId()));
            p.Attributes.Add(new XAttribute(W14 + "textId", "77777777"));
            if (textStyle is not null) p.Props.SetVal("pStyle", textStyle);
            if (k == 0)
            {
                var mark = new Run { Props = new PropertyBag("rPr", w) };
                if (refStyle is not null) mark.Props.SetVal("rStyle", refStyle);
                mark.Content.Add(new OpaqueRunContent(new XElement(w + "annotationRef")));
                p.Inlines.Add(mark);
            }
            if (lines[k].Length > 0)
            {
                var run = new Run { Props = new PropertyBag("rPr", w) };
                run.Content.Add(new TextContent { Value = lines[k] });
                p.Inlines.Add(run);
            }
            blocks.Add(p);
        }
        return blocks;
    }

    StoryHolder NewHolder(XNamespace w, string id, string author, string initials, string text, Func<string> newParaId)
    {
        var h = new StoryHolder { Name = w + "comment" };
        h.Attributes.Add(new XAttribute(w + "id", id));
        h.Attributes.Add(new XAttribute(w + "author", author));
        h.Attributes.Add(new XAttribute(w + "date", Now()));
        if (initials.Length > 0) h.Attributes.Add(new XAttribute(w + "initials", initials));
        h.Story.Blocks.AddRange(Body(w, text, newParaId));
        return h;
    }

    void EnsureW14(StoryPart part)
    {
        if (!part.RootAttributes.Any(a => a.IsNamespaceDeclaration && a.Value == W14.NamespaceName))
            part.RootAttributes.Add(new XAttribute(XNamespace.Xmlns + "w14", W14.NamespaceName));
    }

    /// <summary>
    /// Nouveau commentaire sur la plage [<paramref name="from"/>, <paramref name="to"/>] d'un même conteneur du corps
    /// (marques de début et de fin, appel juste après la fin). Renvoie son identifiant.
    /// </summary>
    public string Add(UndoJournal journal, Story root, DocPos from, DocPos to, string author, string initials, string text, Func<string> newParaId,
        Action<EditTransaction>? alsoTrack = null)
    {
        if (to < from) (from, to) = (to, from);
        if (!from.ContainerPath.SequenceEqual(to.ContainerPath)) to = from.WithOffset(ParagraphAtoms.Length(from.Paragraph(root)));
        var part = _pkg.CommentsPart();
        EnsureW14(part);
        var ext = Extended(create: true)!;
        var w = part.W;
        var id = NextId().ToString(CultureInfo.InvariantCulture);
        using var tx = journal.Begin("Nouveau commentaire", root, from.ContainerPath, from.BlockIndex, to.BlockIndex - from.BlockIndex + 1, to);
        tx.TrackList(part.Items);
        tx.TrackXml(ext);
        alsoTrack?.Invoke(tx);
        var holder = NewHolder(w, id, author, initials, text, newParaId);
        part.Items.Add(holder);
        ext.Add(new XElement(W15 + "commentEx", new XAttribute(W15 + "paraId", LastParaId(holder)!), new XAttribute(W15 + "done", "0")));
        // fin d'abord : les positions du début restent valables
        var endP = to.Paragraph(root);
        var (eo, ei) = Place(endP, to.Offset, start: false);
        var refRun = ReferenceRun(w, id);
        eo.Insert(ei, Marker(w, "commentRangeEnd", id));
        eo.Insert(ei + 1, refRun);
        var (so, si) = Place(from.Paragraph(root), from.Offset, start: true);
        so.Insert(si, Marker(w, "commentRangeStart", id));
        tx.Commit(to);
        return id;
    }

    static OpaqueInline Marker(XNamespace w, string local, string id) => new(new XElement(w + local, new XAttribute(w + "id", id)));

    Run ReferenceRun(XNamespace w, string id)
    {
        var run = new Run { Props = new PropertyBag("rPr", w) };
        if (EnsureStyle?.Invoke("annotation reference") is { } st) run.Props.SetVal("rStyle", st);
        run.Content.Add(new OpaqueRunContent(new XElement(w + "commentReference", new XAttribute(w + "id", id))));
        return run;
    }

    /// <summary>Réponse à un commentaire : même plage que lui, reliée par commentsExtended (w15:paraIdParent).</summary>
    public string Reply(UndoJournal journal, Story root, string parentId, string author, string initials, string text, Func<string> newParaId,
        Action<EditTransaction>? alsoTrack = null)
    {
        var part = _pkg.CommentsPart();
        EnsureW14(part);
        var parent = All().FirstOrDefault(c => c.Id == parentId) ?? throw new ArgumentException("commentaire inconnu", nameof(parentId));
        // une réponse à une réponse se rattache au fil (le premier commentaire), comme dans Word
        var top = parent;
        while (top.ParentParaId is { } pp && All().FirstOrDefault(c => c.ParaId == pp) is { } up) top = up;
        var ext = Extended(create: true)!;
        var w = part.W;
        var id = NextId().ToString(CultureInfo.InvariantCulture);
        using var tx = journal.Begin("Répondre au commentaire", root, Array.Empty<int>(), 0, root.Blocks.Count);
        tx.TrackList(part.Items);
        tx.TrackXml(ext);
        tx.TrackStory(top.Holder.Story);
        alsoTrack?.Invoke(tx);
        // le commentaire parent reçoit un w14:paraId s'il n'en a pas (lien de la réponse)
        var parentPara = top.ParaId;
        if (parentPara is null && top.Holder.Story.Paragraphs().LastOrDefault() is { } lp)
        {
            parentPara = newParaId();
            lp.Attributes.Add(new XAttribute(W14 + "paraId", parentPara));
        }
        var holder = NewHolder(w, id, author, initials, text, newParaId);
        int at = part.Items.IndexOf(top.Holder);
        // après le fil existant (le parent et ses réponses)
        var thread = All().Where(c => c.ParentParaId == parentPara).Select(c => part.Items.IndexOf(c.Holder)).DefaultIfEmpty(at).Max();
        part.Items.Insert(thread + 1, holder);
        ext.Add(new XElement(W15 + "commentEx", new XAttribute(W15 + "paraId", LastParaId(holder)!), new XAttribute(W15 + "paraIdParent", parentPara!), new XAttribute(W15 + "done", "0")));
        // marques : juste après celles du parent
        if (After(root, i => Is(i, "commentRangeStart", top.Id)) is { } s) s.Owner.Insert(s.Index, Marker(w, "commentRangeStart", id));
        if (After(root, i => Is(i, "commentRangeEnd", top.Id)) is { } e) e.Owner.Insert(e.Index, Marker(w, "commentRangeEnd", id));
        if (After(root, i => i is Run r && r.Content.Any(c => IsReference(c, top.Id))) is { } rr) rr.Owner.Insert(rr.Index, ReferenceRun(w, id));
        tx.Commit();
        return id;
    }

    // ================================================================ modifier

    /// <summary>Marque un commentaire (et son fil) comme résolu ou le rouvre (w15:done).</summary>
    public void SetDone(UndoJournal journal, Story root, string id, bool done)
    {
        var c = All().FirstOrDefault(x => x.Id == id) ?? throw new ArgumentException("commentaire inconnu", nameof(id));
        var ext = Extended(create: true)!;
        using var tx = journal.Begin(done ? "Résoudre le commentaire" : "Rouvrir le commentaire", root, Array.Empty<int>(), 0, 0);
        tx.TrackXml(ext);
        var paraId = c.ParaId;
        if (paraId is null) { tx.Dispose(); return; }
        var entry = ext.Elements(W15 + "commentEx").FirstOrDefault(e => (string?)e.Attribute(W15 + "paraId") == paraId);
        if (entry is null) ext.Add(entry = new XElement(W15 + "commentEx", new XAttribute(W15 + "paraId", paraId)));
        entry.SetAttributeValue(W15 + "done", done ? "1" : "0");
        tx.Commit();
    }

    /// <summary>Remplace le texte d'un commentaire (le lien des réponses, w14:paraId du dernier paragraphe, est gardé).</summary>
    public void Edit(UndoJournal journal, Story root, string id, string text, Func<string> newParaId)
    {
        var c = All().FirstOrDefault(x => x.Id == id) ?? throw new ArgumentException("commentaire inconnu", nameof(id));
        var w = c.Holder.Name.Namespace;
        using var tx = journal.Begin("Modifier le commentaire", root, Array.Empty<int>(), 0, 0);
        tx.TrackStory(c.Holder.Story);
        var first = c.Holder.Story.Paragraphs().FirstOrDefault();
        var blocks = Body(w, text, newParaId, c.ParaId);
        // mise en forme du paragraphe d'origine gardée
        if (first is not null) foreach (var p in blocks.OfType<Paragraph>()) p.Props = first.Props.Clone();
        c.Holder.Story.Blocks.Clear();
        c.Holder.Story.Blocks.AddRange(blocks);
        tx.Commit();
    }

    /// <summary>Supprime un commentaire, ses réponses et leurs marques dans le texte.</summary>
    public void Delete(UndoJournal journal, Story root, string id, IEnumerable<Story>? otherStories = null)
    {
        if (Part is not { } part) return;
        var all = All();
        var c = all.FirstOrDefault(x => x.Id == id) ?? throw new ArgumentException("commentaire inconnu", nameof(id));
        var doomed = new List<CommentInfo> { c };
        for (int k = 0; k < doomed.Count; k++)
            doomed.AddRange(all.Where(x => x.ParentParaId is not null && x.ParentParaId == doomed[k].ParaId && !doomed.Contains(x)));
        var ext = Extended();
        var ids = _pkg.XmlPartOf("commentsIds")?.Doc.Root;
        var cex = _pkg.XmlPartOf("commentsExtensible")?.Doc.Root;
        using var tx = journal.Begin(doomed.Count > 1 ? "Supprimer le fil de commentaires" : "Supprimer le commentaire", root, Array.Empty<int>(), 0, root.Blocks.Count);
        tx.TrackList(part.Items);
        if (ext is not null) tx.TrackXml(ext);
        if (ids is not null) tx.TrackXml(ids);
        if (cex is not null) tx.TrackXml(cex);
        var others = (otherStories ?? Enumerable.Empty<Story>()).Where(s => !ReferenceEquals(s, root)).ToList();
        foreach (var s in others) tx.TrackStory(s);
        foreach (var d in doomed)
        {
            foreach (var s in others.Prepend(root)) RemoveMarks(s, d.Id);
            part.Items.Remove(d.Holder);
            if (d.ParaId is not { } pid) continue;
            ext?.Elements(W15 + "commentEx").Where(e => (string?)e.Attribute(W15 + "paraId") == pid).Remove();
            if (ids is not null)
            {
                var ns = DocxPackage.W16Cid;
                foreach (var e in ids.Elements(ns + "commentId").Where(e => (string?)e.Attribute(ns + "paraId") == pid).ToList())
                {
                    var durable = (string?)e.Attribute(ns + "durableId");
                    e.Remove();
                    cex?.Elements(DocxPackage.W16Cex + "commentExtensible").Where(x => (string?)x.Attribute(DocxPackage.W16Cex + "durableId") == durable).Remove();
                }
            }
        }
        tx.Commit();
    }

    static void RemoveMarks(Story story, string id)
    {
        foreach (var p in story.Paragraphs()) Remove(p.Inlines);

        void Remove(List<Inline> list)
        {
            list.RemoveAll(i => Is(i, "commentRangeStart", id) || Is(i, "commentRangeEnd", id));
            foreach (var run in list.OfType<Run>().ToList())
            {
                if (run.Content.RemoveAll(c => IsReference(c, id)) > 0 && run.Content.Count == 0) list.Remove(run);
            }
            foreach (var b in list.OfType<InlineContainer>()) Remove(b.Children);
            foreach (var s in list.OfType<InlineContentControl>()) Remove(s.Content);
        }
    }
}
