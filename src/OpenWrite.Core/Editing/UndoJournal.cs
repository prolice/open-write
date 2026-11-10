using System.Xml.Linq;
using OpenWrite.Core.Model;

namespace OpenWrite.Core.Editing;

/// <summary>
/// Annuler / rétablir (lot 1, PLAN.md §10.4). Chaque entrée remplace une plage de blocs d'un conteneur : elle garde une
/// copie des blocs AVANT et APRÈS l'opération. Annuler remet la copie d'avant, rétablir celle d'après.
/// <list type="bullet">
/// <item>Exact par construction : les copies gardent tout (attributs, XML conservé, valeurs lues) — annuler redonne le XML d'origine.</item>
/// <item>Coût local : seuls les blocs touchés sont copiés (un paragraphe pour une frappe), pas le document (OpenPub sérialisait tout).</item>
/// <item>Toute fonction d'édition devient annulable sans écrire son inverse ; une édition qui échoue est défaite (transaction).</item>
/// <item>Les entrées désignent leurs blocs par chemin (<see cref="DocPos"/>), jamais par objet : elles restent valables
///   après d'autres annulations qui ont remplacé des objets par leurs copies.</item>
/// </list>
/// Le journal ne connaît pas les parties du package : <see cref="Changed"/> indique quelle racine a changé.
/// </summary>
public sealed class UndoJournal
{
    sealed class Entry
    {
        public required string Label { get; init; }
        public required Story Root { get; init; }
        public required int[] Container { get; init; }
        public required int Index { get; init; }
        public required List<Block> Before { get; init; }
        public required List<Block> After { get; set; }
        public DocPos? CaretBefore { get; init; }
        public DocPos? CaretAfter { get; set; }
        public string? MergeKey { get; init; }
        public DateTime Time { get; set; }
        /// <summary>Parties XML modifiées (styles, numérotation) : racine vivante, copies avant / après.</summary>
        public List<(XElement Root, XElement Before, XElement After)> Xml { get; init; } = new();
        /// <summary>Autres suites de blocs modifiées (en-têtes d'un filigrane…) : copies avant / après de tous leurs blocs.</summary>
        public List<(Story Root, List<Block> Before, List<Block> After)> Stories { get; init; } = new();
    }

    readonly List<Entry> _undo = new();
    readonly List<Entry> _redo = new();
    EditTransaction? _open;
    /// <summary>Dernière entrée appliquée au moment de l'enregistrement (null : aucune) ; _savedLost : état enregistré inatteignable.</summary>
    Entry? _savedAt;
    bool _savedLost;

    /// <summary>Nombre maximal d'entrées gardées (les plus anciennes sont oubliées).</summary>
    public int MaxEntries { get; set; } = 1000;

    /// <summary>Une racine a été modifiée (opération, annulation ou rétablissement) : sa partie est à régénérer.</summary>
    public event Action<Story>? Changed;

    /// <summary>Une partie XML suivie (styles, numérotation) a changé : à régénérer, résolveur de styles à refaire.</summary>
    public event Action<XElement>? XmlChanged;

    /// <summary>Parties XML touchées par la dernière opération, annulation ou rétablissement (vide le plus souvent).</summary>
    public IReadOnlyList<XElement> LastXml { get; private set; } = Array.Empty<XElement>();

    /// <summary>
    /// Zone touchée par la dernière modification, annulation ou rétablissement : suite de blocs, conteneur, premier bloc
    /// et nombre de blocs qui l'occupent maintenant (pour la remise en page incrémentale).
    /// </summary>
    public (Story Root, IReadOnlyList<int> Container, int Index, int Count)? LastChange { get; private set; }

    public bool CanUndo => _undo.Count > 0;
    public bool CanRedo => _redo.Count > 0;
    public string? UndoLabel => _undo.Count > 0 ? _undo[^1].Label : null;
    public string? RedoLabel => _redo.Count > 0 ? _redo[^1].Label : null;
    public int Count => _undo.Count;

    /// <summary>Le document diffère-t-il de son dernier enregistrement ?</summary>
    public bool IsModified => _savedLost || (_undo.Count > 0 ? _undo[^1] : null) != _savedAt;

    public void MarkSaved()
    {
        _savedAt = _undo.Count > 0 ? _undo[^1] : null;
        _savedLost = false;
    }

    // ---------------------------------------------------------------- transactions

    /// <summary>
    /// Ouvre une transaction sur les blocs <paramref name="index"/>..+<paramref name="count"/> du conteneur : l'opération
    /// ne doit modifier que ces blocs (elle peut en ajouter ou en retirer à cet endroit). Sans <see cref="EditTransaction.Commit"/>,
    /// la fin de la transaction (Dispose) remet les blocs d'origine.
    /// </summary>
    public EditTransaction Begin(string label, Story root, IReadOnlyList<int> containerPath, int index, int count,
        DocPos? caretBefore = null, string? mergeKey = null)
    {
        if (_open is not null) throw new InvalidOperationException("Une transaction est déjà ouverte.");
        var container = DocPos.ResolveContainer(root, containerPath);
        if (index < 0 || count < 0 || index + count > container.Blocks.Count) throw new ArgumentOutOfRangeException(nameof(index));
        var before = container.Blocks.GetRange(index, count).Select(b => b.Clone()).ToList();
        _open = new EditTransaction(this, label, root, containerPath.ToArray(), index, container.Blocks.Count, before, caretBefore, mergeKey);
        return _open;
    }

    internal void Complete(EditTransaction tx, bool commit, DocPos? caretAfter)
    {
        _open = null;
        var container = DocPos.ResolveContainer(tx.Root, tx.Container);
        int newCount = tx.Before.Count + container.Blocks.Count - tx.ContainerCount;
        if (!commit)
        {
            // échec : on remet les blocs d'avant (l'opération a pu les modifier en partie), et les parties XML
            Replace(container, tx.Index, Math.Max(0, newCount), tx.Before.Select(b => b.Clone()));
            foreach (var (root, before) in tx.Xml) Restore(root, before);
            foreach (var (story, before) in tx.Stories) Replace(story, 0, story.Blocks.Count, before.Select(b => b.Clone()));
            return;
        }
        var stories = tx.Stories.Select(x => (x.Root, x.Before, x.Root.Blocks.Select(b => b.Clone()).ToList())).ToList();
        var xml = tx.Xml.Select(x => (x.Root, x.Before, new XElement(x.Root))).ToList();
        LastXml = xml.Select(x => x.Root).ToList();
        var after = container.Blocks.GetRange(tx.Index, newCount).Select(b => b.Clone()).ToList();
        LastChange = (tx.Root, tx.Container, tx.Index, newCount);
        if (_savedAt is not null && _redo.Contains(_savedAt)) _savedLost = true;      // l'état enregistré n'est plus atteignable
        _redo.Clear();
        var last = _undo.Count > 0 ? _undo[^1] : null;
        if (tx.MergeKey is not null && xml.Count == 0 && stories.Count == 0 && last is not null && last.Xml.Count == 0 && last.Stories.Count == 0 && last.MergeKey == tx.MergeKey && last != _savedAt && last.Root == tx.Root
            && last.Container.SequenceEqual(tx.Container) && last.Index == tx.Index && last.After.Count == tx.Before.Count
            && Equals(last.CaretAfter, tx.CaretBefore))
        {
            // frappe continue : une seule entrée (avant = celui de la première frappe, après = celui-ci)
            last.After = after;
            last.CaretAfter = caretAfter;
            last.Time = DateTime.UtcNow;
        }
        else
        {
            _undo.Add(new Entry
            {
                Label = tx.Label, Root = tx.Root, Container = tx.Container, Index = tx.Index, Before = tx.Before, After = after,
                CaretBefore = tx.CaretBefore, CaretAfter = caretAfter, MergeKey = tx.MergeKey, Time = DateTime.UtcNow, Xml = xml, Stories = stories,
            });
            if (_undo.Count > MaxEntries)
            {
                if (_undo[0] == _savedAt) _savedLost = true;
                _undo.RemoveAt(0);
            }
        }
        Changed?.Invoke(tx.Root);
        foreach (var x in stories) Changed?.Invoke(x.Root);
        foreach (var x in xml) XmlChanged?.Invoke(x.Root);
    }

    /// <summary>Remet une racine XML dans l'état d'une copie (même objet racine : les références restent valables).</summary>
    static void Restore(XElement root, XElement state)
    {
        root.ReplaceAttributes(state.Attributes().Select(a => new XAttribute(a)));
        root.ReplaceNodes(state.Nodes());
    }

    // ---------------------------------------------------------------- annuler / rétablir

    /// <summary>Annule la dernière opération ; renvoie la position du point d'insertion d'avant (ou null).</summary>
    public DocPos? Undo()
    {
        if (_undo.Count == 0) return null;
        var e = _undo[^1];
        _undo.RemoveAt(_undo.Count - 1);
        Replace(DocPos.ResolveContainer(e.Root, e.Container), e.Index, e.After.Count, e.Before.Select(b => b.Clone()));
        LastChange = (e.Root, e.Container, e.Index, e.Before.Count);
        foreach (var (root, before, _) in e.Xml) Restore(root, before);
        foreach (var (story, before, _) in e.Stories) Replace(story, 0, story.Blocks.Count, before.Select(b => b.Clone()));
        LastXml = e.Xml.Select(x => x.Root).ToList();
        _redo.Add(e);
        Changed?.Invoke(e.Root);
        foreach (var x in e.Stories) Changed?.Invoke(x.Root);
        foreach (var x in e.Xml) XmlChanged?.Invoke(x.Root);
        return e.CaretBefore;
    }

    /// <summary>Rétablit la dernière opération annulée ; renvoie la position du point d'insertion d'après (ou null).</summary>
    public DocPos? Redo()
    {
        if (_redo.Count == 0) return null;
        var e = _redo[^1];
        _redo.RemoveAt(_redo.Count - 1);
        Replace(DocPos.ResolveContainer(e.Root, e.Container), e.Index, e.Before.Count, e.After.Select(b => b.Clone()));
        LastChange = (e.Root, e.Container, e.Index, e.After.Count);
        foreach (var (root, _, after) in e.Xml) Restore(root, after);
        foreach (var (story, _, after) in e.Stories) Replace(story, 0, story.Blocks.Count, after.Select(b => b.Clone()));
        LastXml = e.Xml.Select(x => x.Root).ToList();
        _undo.Add(e);
        Changed?.Invoke(e.Root);
        foreach (var x in e.Stories) Changed?.Invoke(x.Root);
        foreach (var x in e.Xml) XmlChanged?.Invoke(x.Root);
        return e.CaretAfter;
    }

    /// <summary>Oublie tout (nouveau document, rechargement).</summary>
    public void Clear()
    {
        _undo.Clear();
        _redo.Clear();
        LastChange = null;
        _savedAt = null;
        _savedLost = false;
    }

    static void Replace(Story container, int index, int count, IEnumerable<Block> blocks)
    {
        container.Blocks.RemoveRange(index, count);
        container.Blocks.InsertRange(index, blocks);
    }

    // ---------------------------------------------------------------- regroupement de la frappe

    string? _typingWord;

    /// <summary>
    /// Clé de regroupement de la frappe : les caractères d'un même mot (et l'espace qui le termine) forment une seule
    /// entrée ; une lettre tapée après une espace commence une nouvelle entrée, comme dans Word.
    /// </summary>
    public string TypingKey(DocPos pos, string text)
    {
        var last = _undo.Count > 0 ? _undo[^1] : null;
        bool continues = last?.MergeKey is { } k && k == _typingWord && Equals(last.CaretAfter, pos);
        bool startsWord = text.Length > 0 && !char.IsWhiteSpace(text[0]);
        bool previousEndedWithSpace = continues && _lastTypedSpace;
        _lastTypedSpace = text.Length > 0 && char.IsWhiteSpace(text[^1]);
        if (continues && !(startsWord && previousEndedWithSpace)) return _typingWord!;
        _typingWord = "frappe-" + Guid.NewGuid().ToString("N")[..8];
        return _typingWord;
    }

    bool _lastTypedSpace;
}

/// <summary>Transaction ouverte par <see cref="UndoJournal.Begin"/> ; à valider par <see cref="Commit"/>, sinon défaite.</summary>
public sealed class EditTransaction : IDisposable
{
    readonly UndoJournal _journal;
    bool _done;

    internal EditTransaction(UndoJournal journal, string label, Story root, int[] container, int index, int containerCount,
        List<Block> before, DocPos? caretBefore, string? mergeKey)
    {
        _journal = journal;
        Label = label;
        Root = root;
        Container = container;
        Index = index;
        ContainerCount = containerCount;
        Before = before;
        CaretBefore = caretBefore;
        MergeKey = mergeKey;
    }

    internal string Label { get; }
    internal Story Root { get; }
    internal int[] Container { get; }
    internal int Index { get; }
    internal int ContainerCount { get; }
    internal List<Block> Before { get; }
    internal DocPos? CaretBefore { get; }
    internal string? MergeKey { get; }
    internal List<(XElement Root, XElement Before)> Xml { get; } = new();
    internal List<(Story Root, List<Block> Before)> Stories { get; } = new();

    /// <summary>Suit une autre suite de blocs (un en-tête…) dans cette transaction : tous ses blocs sont copiés maintenant.</summary>
    public void TrackStory(Story root)
    {
        if (ReferenceEquals(root, Root) || Stories.Any(x => ReferenceEquals(x.Root, root))) return;
        Stories.Add((root, root.Blocks.Select(b => b.Clone()).ToList()));
    }

    /// <summary>
    /// Suit une partie XML (styles.xml, numbering.xml) dans cette transaction : copie prise maintenant, restaurée à
    /// l'annulation ou si la transaction échoue.
    /// </summary>
    public void TrackXml(XElement root)
    {
        if (Xml.Any(x => ReferenceEquals(x.Root, root))) return;
        Xml.Add((root, new XElement(root)));
    }

    public void Commit(DocPos? caretAfter = null)
    {
        if (_done) return;
        _done = true;
        _journal.Complete(this, commit: true, caretAfter);
    }

    public void Dispose()
    {
        if (_done) return;
        _done = true;
        _journal.Complete(this, commit: false, null);
    }
}
