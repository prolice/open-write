using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Xml.Linq;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Core.Review;
using OpenWrite.Docx;
using OpenWrite.Layout;

namespace OpenWrite.App;

// Lot 8 (version 1.1), révision : suivi des modifications (w:trackRevisions dans settings.xml, Ctrl+Maj+E ; frappe,
// suppressions, Entrée, collage et mise en forme enregistrés par Core/Review/TrackContext), affichage des révisions
// (marques simples, toutes les marques, aucune marque, original), accepter / refuser (une, toutes, par auteur,
// toutes et arrêter le suivi), modification précédente / suivante ; commentaires (Docx/CommentStore) : nouveau
// (Ctrl+Alt+M), volet Commentaires (fils, réponses, résolu, modifier, supprimer), précédent / suivant.
public partial class MainWindow
{
    static readonly XNamespace Wm = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    RevisionView _view = Enum.TryParse<RevisionView>(AppSettings.Load().RevisionView, out var v) ? v : RevisionView.Simple;
    IDisposable? _trackScope;
    DocxPackage? _reviewPkg;
    CommentStore? _comments;

    /// <summary>Le suivi des modifications est actif pour ce document.</summary>
    internal bool Tracking => _trackScope is not null && TrackContext.Current is not null;
    internal RevisionView View => _view;
    internal CommentStore? Comments => _comments;

    internal static string UserName => AppSettings.Load().UserName is { Length: > 0 } n ? n : Environment.UserName;
    internal static string UserInitials => AppSettings.Load().UserInitials is { Length: > 0 } i ? i
        : string.Concat(UserName.Split(new[] { ' ', '.', '-', '_' }, StringSplitOptions.RemoveEmptyEntries).Take(3).Select(x => char.ToUpperInvariant(x[0])));

    static readonly Brush SimpleBar = Freeze(new SolidColorBrush(Color.FromRgb(0xD1, 0x34, 0x38)));
    static readonly Brush AllBar = Freeze(new SolidColorBrush(Color.FromRgb(0x59, 0x59, 0x59)));
    static T Freeze<T>(T f) where T : Freezable { f.Freeze(); return f; }

    /// <summary>Barre de modification de l'affichage courant (aussi à l'impression et dans le PDF, comme Word).</summary>
    internal Brush? RevisionBarBrush => _view switch { RevisionView.Simple => SimpleBar, RevisionView.All => AllBar, _ => null };

    void InitReview()
    {
        PreviewKeyDown += (_, e) =>
        {
            if (_layout is null) return;
            var mods = Keyboard.Modifiers;
            var key = e.Key == Key.System ? e.SystemKey : e.Key;
            if (key == Key.E && mods == (ModifierKeys.Control | ModifierKeys.Shift)) { ToggleTracking(); e.Handled = true; }
            else if (key == Key.M && mods == (ModifierKeys.Control | ModifierKeys.Alt)) { NewComment(); e.Handled = true; }
        };
    }

    /// <summary>Nouveau document affiché : suivi selon settings.xml, commentaires, volet.</summary>
    void ReadReviewSettings()
    {
        if (ReferenceEquals(_reviewPkg, _pkg)) return;
        _reviewPkg = _pkg;
        StopTracking();
        _comments = _pkg is null ? null : new CommentStore(_pkg) { EnsureStyle = EnsureCommentStyle };
        if (DocTracking) StartTracking();
        UpdateTrackingButton();
        RefreshCommentsPane();
    }

    /// <summary>Styles « Commentaire » et « Marque de commentaire » ajoutés à styles.xml (styles.xml suivi par la transaction en cours).</summary>
    string? EnsureCommentStyle(string builtInName)
    {
        if (StylesRoot(create: true) is not { } root || _styles?.Sheet is not { } sheet) return null;
        return Core.Styles.StyleOps.Ensure(sheet, root, builtInName);
    }

    XElement? SettingsRoot => _pkg?.XmlPartOf("settings")?.Doc.Root;
    internal XElement? SettingsForTest => SettingsRoot;

    bool DocTracking => SettingsRoot?.Element(Wm + "trackRevisions") is { } t && (string?)t.Attribute(Wm + "val") is null or "1" or "true" or "on";

    void StartTracking()
    {
        StopTracking();
        _trackScope = TrackContext.Use(new TrackContext(TrackContext.IdsAfter(AllStories())) { Author = UserName });
    }

    void StopTracking()
    {
        _trackScope?.Dispose();
        _trackScope = null;
    }

    /// <summary>Après une annulation : le suivi suit l'état enregistré dans settings.xml.</summary>
    void SyncTracking()
    {
        if (SettingsRoot is null) return;
        if (DocTracking && !Tracking) StartTracking();
        else if (!DocTracking && Tracking) StopTracking();
        UpdateTrackingButton();
    }

    /// <summary>Éléments de settings.xml qui suivent w:trackRevisions dans l'ordre du schéma.</summary>
    static readonly HashSet<string> AfterTrackRevisions = new()
    {
        "doNotTrackMoves", "doNotTrackFormatting", "documentProtection", "autoFormatOverride", "styleLockTheme", "styleLockQFSet",
        "defaultTabStop", "autoHyphenation", "consecutiveHyphenLimit", "hyphenationZone", "doNotHyphenateCaps", "showEnvelope",
        "summaryLength", "clickAndTypeStyle", "defaultTableStyle", "evenAndOddHeaders", "bookFoldRevPrinting", "bookFoldPrinting",
        "bookFoldPrintingSheets", "drawingGridHorizontalSpacing", "drawingGridVerticalSpacing", "displayHorizontalDrawingGridEvery",
        "displayVerticalDrawingGridEvery", "doNotUseMarginsForDrawingGridOrigin", "drawingGridHorizontalOrigin", "drawingGridVerticalOrigin",
        "doNotShadeFormData", "noPunctuationKerning", "characterSpacingControl", "printTwoOnOne", "strictFirstAndLastChars",
        "noLineBreaksAfter", "noLineBreaksBefore", "savePreviewPicture", "doNotValidateAgainstSchema", "saveInvalidXml",
        "ignoreMixedContent", "alwaysShowPlaceholderText", "doNotDemarcateInvalidXml", "saveXmlDataOnly", "useXSLTWhenSaving",
        "saveThroughXslt", "showXMLTags", "alwaysMergeEmptyNamespace", "updateFields", "hdrShapeDefaults", "footnotePr", "endnotePr",
        "compat", "docVars", "rsids", "mathPr", "attachedSchema", "themeFontLang", "clrSchemeMapping", "doNotIncludeSubdocsInStats",
        "doNotAutoCompressPictures", "forceUpgrade", "captions", "readModeInkLockDown", "smartTagType", "schemaLibrary",
        "shapeDefaults", "doNotEmbedSmartTags", "decimalSymbol", "listSeparator",
    };

    /// <summary>Active ou arrête le suivi des modifications (enregistré dans le document, comme Word).</summary>
    internal void ToggleTracking()
    {
        if (_layout is null || Body is not { } body) return;
        bool on = !Tracking;
        if (SettingsRoot is { } settings)
        {
            using var tx = _journal.Begin(on ? "Activer le suivi des modifications" : "Arrêter le suivi des modifications", body, Array.Empty<int>(), 0, 0);
            tx.TrackXml(settings);
            settings.Elements(Wm + "trackRevisions").Remove();
            if (on)
            {
                var el = new XElement(settings.Name.Namespace + "trackRevisions");
                var after = settings.Elements().FirstOrDefault(e => e.Name.Namespace != settings.Name.Namespace || AfterTrackRevisions.Contains(e.Name.LocalName));
                if (after is null) settings.Add(el); else after.AddBeforeSelf(el);
            }
            tx.Commit();
        }
        if (on) StartTracking(); else StopTracking();
        UpdateTrackingButton();
        UpdateModifiedMarker();
        StatusText.Text = on ? $"Suivi des modifications activé : vos modifications sont enregistrées au nom de « {UserName} » (Ctrl+Maj+E pour l'arrêter)."
                             : "Suivi des modifications arrêté (les révisions déjà faites restent : Révision ▸ Accepter ou Refuser).";
        _pagesView?.Focus();
    }

    void UpdateTrackingButton()
    {
        foreach (var b in _liveButtons.Where(b => Equals(b.Content, "Suivi des modifications")))
            b.Background = Tracking ? R("Selected") : Brushes.Transparent;
    }

    // ================================================================ affichage des révisions

    internal void ShowViewMenu(FrameworkElement anchor)
    {
        var menu = new ContextMenu { PlacementTarget = anchor, Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom };
        LastMenu = menu;
        foreach (var (view, label, tip) in new[]
        {
            (RevisionView.Simple, "Marques simples", "Texte final, une barre rouge dans la marge signale les modifications"),
            (RevisionView.All, "Toutes les marques", "Insertions soulignées et suppressions barrées, en couleur par auteur"),
            (RevisionView.Final, "Aucune marque", "Le texte final, sans aucune marque"),
            (RevisionView.Original, "Original", "Le texte d'avant les modifications"),
        })
        {
            var mi = new MenuItem { Header = label, IsCheckable = true, IsChecked = _view == view, ToolTip = tip };
            mi.Click += (_, _) => SetRevisionView(view);
            menu.Items.Add(mi);
        }
        menu.Items.Add(new Separator());
        var name = new MenuItem { Header = "Changer le nom d'utilisateur…" };
        name.Click += (_, _) => ChangeUserName();
        menu.Items.Add(name);
        menu.IsOpen = true;
    }

    internal void SetRevisionView(RevisionView view, bool persist = true)
    {
        _view = view;
        if (persist) AppSettings.Update(x => x.RevisionView = view.ToString());
        if (_layout is null) return;
        ShowDocument();
        StatusText.Text = view switch
        {
            RevisionView.All => "Toutes les marques : insertions soulignées, suppressions barrées, en couleur par auteur.",
            RevisionView.Final => "Aucune marque : le texte final.",
            RevisionView.Original => "Original : le texte d'avant les modifications (pour modifier, revenez à un autre affichage).",
            _ => "Marques simples : une barre rouge dans la marge signale les lignes modifiées.",
        };
    }

    void ChangeUserName()
    {
        var box = new TextBox { Text = UserName, Width = 260 };
        var initials = new TextBox { Text = UserInitials, Width = 80 };
        var dlg = new Window();
        DialogKit.Setup(dlg, "Nom d'utilisateur");
        var root = new StackPanel { Margin = new Thickness(16) };
        root.Children.Add(new TextBlock { Text = "Les commentaires et les modifications suivies sont enregistrés à ce nom.", TextWrapping = TextWrapping.Wrap, MaxWidth = 360, Margin = new Thickness(0, 0, 0, 10) });
        root.Children.Add(DialogKit.Row("Nom", box, 90));
        root.Children.Add(DialogKit.Row("Initiales", initials, 90));
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 12, 0, 0) };
        var cancel = new Button { Content = "Annuler", IsCancel = true, MinWidth = 90, Margin = new Thickness(0, 0, 8, 0) };
        var ok = new Button { Content = "OK", IsDefault = true, MinWidth = 90, Style = (Style)FindResource("PillButton") };
        ok.Click += (_, _) => dlg.DialogResult = true;
        buttons.Children.Add(cancel);
        buttons.Children.Add(ok);
        root.Children.Add(buttons);
        dlg.Content = root;
        dlg.Owner = this;
        if (dlg.ShowDialog() != true) return;
        AppSettings.Update(x => { x.UserName = box.Text.Trim(); x.UserInitials = initials.Text.Trim(); });
        if (Tracking) StartTracking();
        StatusText.Text = $"Nom d'utilisateur : {UserName} ({UserInitials}).";
    }

    // ================================================================ accepter / refuser

    /// <summary>Ordre des paragraphes du corps (pour comparer des positions).</summary>
    Dictionary<Paragraph, int> ParagraphOrder()
    {
        var order = new Dictionary<Paragraph, int>(ReferenceEqualityComparer.Instance);
        if (Body is { } body) foreach (var p in body.Paragraphs()) order[p] = order.Count;
        return order;
    }

    (int Para, int Offset)? PosKey(TextPoint t, Dictionary<Paragraph, int> order) =>
        order.TryGetValue(t.Paragraph, out var k) && LayoutOf(t.Paragraph) is { } pl ? (k, CaretGeometry.ModelOffset(pl.Content, t.Index)) : null;

    static int Cmp((int, int) a, (int, int) b) => a.Item1 != b.Item1 ? a.Item1.CompareTo(b.Item1) : a.Item2.CompareTo(b.Item2);

    /// <summary>Révisions du corps sous la sélection (ou au point d'insertion).</summary>
    List<Revision> RevisionsAtSelection()
    {
        if (Body is not { } body || _caret is not { } c) return new();
        var order = ParagraphOrder();
        var (from, to) = HasSelection && Ordered() is { } sel ? (sel.From, sel.To) : (c, c);
        if (PosKey(from, order) is not { } a || PosKey(to, order) is not { } b) return new();
        return Revisions.List(body).Where(r => r.Paragraph is { } p && order.TryGetValue(p, out var k)
            && Cmp((k, r.Start), b) <= 0 && Cmp((k, r.End), a) >= 0).ToList();
    }

    /// <summary>Accepte ou refuse les révisions de la sélection (ou au point d'insertion) ; puis va à la suivante si demandé.</summary>
    internal void ResolveCurrent(bool accept, bool moveNext = false)
    {
        if (Body is not { } body) return;
        var targets = RevisionsAtSelection();
        if (targets.Count == 0)
        {
            if (!GoToRevision(next: true)) StatusText.Text = "Le document ne contient aucune modification suivie.";
            return;
        }
        var caret = _caret is { } c ? ToDocPos(c) : null;
        using (TrackContext.Use(null))
        using (var tx = _journal.Begin(accept ? "Accepter la modification" : "Refuser la modification", body, Array.Empty<int>(), 0, body.Blocks.Count, caret))
        {
            for (int k = targets.Count - 1; k >= 0; k--)
                if (accept) Revisions.Accept(body, targets[k]); else Revisions.Reject(body, targets[k]);
            tx.Commit(caret);
        }
        AfterEdit($"{targets.Count} modification(s) {(accept ? "acceptée(s)" : "refusée(s)")}.", caret);
        if (moveNext) GoToRevision(next: true);
    }

    /// <summary>Accepte ou refuse toutes les révisions (de tout le document : corps, en-têtes, pieds, notes, commentaires), d'un auteur ou de tous.</summary>
    internal int ResolveAll(bool accept, string? author = null, bool stopTracking = false)
    {
        if (Body is not { } body) return 0;
        var caret = _caret is { } c ? ToDocPos(c) : null;
        int n = 0;
        Func<Revision, bool>? filter = author is null ? null : r => r.Author == author;
        using (TrackContext.Use(null))
        using (var tx = _journal.Begin(accept ? "Accepter toutes les modifications" : "Refuser toutes les modifications", body, Array.Empty<int>(), 0, body.Blocks.Count, caret))
        {
            foreach (var s in AllStories().Where(s => !ReferenceEquals(s, body) && Revisions.List(s).Any(r => filter?.Invoke(r) ?? true)))
            {
                tx.TrackStory(s);
                n += Revisions.ResolveAll(s, accept, filter);
            }
            n += Revisions.ResolveAll(body, accept, filter);
            tx.Commit(caret);
        }
        if (stopTracking && Tracking) ToggleTracking();
        AfterEdit(n == 0 ? "Aucune modification suivie." : $"{n} modification(s) {(accept ? "acceptée(s)" : "refusée(s)")}{(author is not null ? $" de {author}" : "")}.", caret);
        return n;
    }

    /// <summary>Sélectionne la modification suivante (ou précédente) du corps ; faux s'il n'y en a pas.</summary>
    internal bool GoToRevision(bool next)
    {
        if (Body is not { } body) return false;
        var order = ParagraphOrder();
        var list = Revisions.List(body).Where(r => r.Paragraph is { } p && order.ContainsKey(p)).ToList();
        if (list.Count == 0) { StatusText.Text = "Le document ne contient aucune modification suivie."; return false; }
        var here = HasSelection && Ordered() is { } sel ? (next ? sel.To : sel.From) : _caret;
        var key = here is { } h ? PosKey(h, order) ?? (0, 0) : (0, 0);
        int At(Revision r) => Cmp((order[r.Paragraph!], r.Start), key);
        Revision? target = next ? list.FirstOrDefault(r => At(r) > 0 || (At(r) == 0 && !HasSelection)) : list.LastOrDefault(r => At(r) < 0);
        // reprise au début (ou à la fin), comme Word
        target ??= next ? list[0] : list[^1];
        SelectModelRange(target.Paragraph!, target.Start, target.End);
        StatusText.Text = $"{target.Label}{(target.Author is { } a ? $" — {a}" : "")}{(target.Date is { } d ? $", {d.ToString("d MMMM yyyy HH:mm", CultureInfo.GetCultureInfo("fr-FR"))}" : "")}.";
        return true;
    }

    /// <summary>Sélectionne une plage du modèle d'un paragraphe (texte affiché correspondant) et y fait défiler la vue.</summary>
    void SelectModelRange(Paragraph p, int start, int end)
    {
        if (LayoutOf(p) is not { } pl) return;
        int a = CaretGeometry.DisplayIndex(pl.Content, start), b = CaretGeometry.DisplayIndex(pl.Content, end);
        _anchor = new TextPoint(p, a);
        _caret = new TextPoint(p, Math.Max(a, b));
        _pendingFormat.Clear();
        UpdateCaret();
    }

    IEnumerable<string> RevisionAuthors() =>
        AllStories().SelectMany(s => Revisions.List(s)).Select(r => r.Author ?? "").Distinct().OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase);

    internal void ShowResolveMenu(FrameworkElement anchor, bool accept)
    {
        var menu = new ContextMenu { PlacementTarget = anchor, Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom };
        LastMenu = menu;
        void Item(string header, Action act)
        {
            var mi = new MenuItem { Header = header };
            mi.Click += (_, _) => { act(); _pagesView?.Focus(); };
            menu.Items.Add(mi);
        }
        var verb = accept ? "Accepter" : "Refuser";
        Item($"{verb} et passer à la suivante", () => ResolveCurrent(accept, moveNext: true));
        Item($"{verb} cette modification", () => ResolveCurrent(accept));
        Item($"{verb} toutes les modifications", () => ResolveAll(accept));
        Item($"{verb} toutes les modifications et arrêter le suivi", () => ResolveAll(accept, stopTracking: true));
        var authors = RevisionAuthors().ToList();
        if (authors.Count > 1)
        {
            var sub = new MenuItem { Header = $"{verb} toutes les modifications de" };
            foreach (var a in authors)
            {
                var mi = new MenuItem { Header = a.Length > 0 ? a : "(auteur inconnu)" };
                mi.Click += (_, _) => { ResolveAll(accept, a); _pagesView?.Focus(); };
                sub.Items.Add(mi);
            }
            menu.Items.Add(sub);
        }
        menu.IsOpen = true;
    }

    // ================================================================ commentaires

    /// <summary>Plage du nouveau commentaire : la sélection, sinon le mot du point d'insertion, sinon le point d'insertion.</summary>
    (DocPos From, DocPos To)? _newCommentRange;

    /// <summary>Nouveau commentaire (Ctrl+Alt+M) : le volet s'ouvre sur une carte à remplir.</summary>
    internal void NewComment()
    {
        if (Body is null || _hf is not null || _box is not null)
        {
            StatusText.Text = "Les commentaires s'ajoutent dans le texte du document (pas dans un en-tête, un pied de page ou une zone de texte).";
            return;
        }
        var target = CharacterTarget() ?? (_caret is { } c ? (c, c) : default((TextPoint, TextPoint)?));
        if (target is not { } t || ToDocPos(t.Item1) is not { } from || ToDocPos(t.Item2) is not { } to) return;
        _newCommentRange = (from, to);
        if (t.Item1 != t.Item2) { _anchor = t.Item1; _caret = t.Item2; UpdateCaret(scroll: false); }
        ShowCommentsPane(true);
        _commentsPane?.BeginNew();
    }

    /// <summary>Publie le nouveau commentaire préparé par <see cref="NewComment"/>.</summary>
    internal string? PublishNewComment(string text)
    {
        if (_comments is null || Body is not { } body || _newCommentRange is not { } r || text.Trim().Length == 0) return null;
        string id;
        try
        {
            using (TrackContext.Use(null))
                id = _comments.Add(_journal, body, r.From, r.To, UserName, UserInitials, text.Trim(), NewParaId(), TrackStyles);
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or NotSupportedException)
        {
            StatusText.Text = "Commentaire impossible ici : " + ex.Message;
            return null;
        }
        _newCommentRange = null;
        AfterEdit("Commentaire ajouté.", _caret is { } c ? ToDocPos(c) : null);
        return id;
    }

    /// <summary>styles.xml suivi par la transaction (les styles des commentaires y sont ajoutés la première fois).</summary>
    void TrackStyles(EditTransaction tx)
    {
        if (StylesRoot(create: true) is { } root) tx.TrackXml(root);
    }

    internal void ReplyToComment(string id, string text)
    {
        if (_comments is null || Body is not { } body || text.Trim().Length == 0) return;
        using (TrackContext.Use(null)) _comments.Reply(_journal, body, id, UserName, UserInitials, text.Trim(), NewParaId(), TrackStyles);
        AfterEdit("Réponse ajoutée.", _caret is { } c ? ToDocPos(c) : null);
    }

    internal void ResolveComment(string id, bool done)
    {
        if (_comments is null || Body is not { } body) return;
        _comments.SetDone(_journal, body, id, done);
        AfterEdit(done ? "Commentaire résolu." : "Commentaire rouvert.", _caret is { } c ? ToDocPos(c) : null);
    }

    internal void EditComment(string id, string text)
    {
        if (_comments is null || Body is not { } body || text.Trim().Length == 0) return;
        _comments.Edit(_journal, body, id, text.Trim(), NewParaId());
        AfterEdit("Commentaire modifié.", _caret is { } c ? ToDocPos(c) : null);
    }

    internal void DeleteComment(string id)
    {
        if (_comments is null || Body is not { } body) return;
        var caret = _caret is { } c ? ToDocPos(c) : null;
        _comments.Delete(_journal, body, id, AllStories().Where(s => !ReferenceEquals(s, body) && _pkg!.Stories.Values.First(p => p.Stories().Contains(s)).Kind != "comments"));
        AfterEdit("Commentaire supprimé.", caret);
    }

    /// <summary>Sélectionne le texte commenté.</summary>
    internal void GoToComment(string id)
    {
        if (Body is not { } body || CommentStore.Anchor(body, id) is not { } a) { StatusText.Text = "Ce commentaire n'est pas dans le texte du document."; return; }
        if (LayoutOf(a.StartParagraph) is not { } sp || LayoutOf(a.EndParagraph) is not { } ep) return;
        _anchor = new TextPoint(a.StartParagraph, CaretGeometry.DisplayIndex(sp.Content, a.Start));
        _caret = new TextPoint(a.EndParagraph, CaretGeometry.DisplayIndex(ep.Content, a.End));
        _pendingFormat.Clear();
        UpdateCaret();
        _pagesView?.Focus();
    }

    /// <summary>Commentaires (premiers de leur fil) dans l'ordre du texte.</summary>
    internal List<CommentInfo> CommentThreads()
    {
        if (_comments is null || Body is not { } body) return new();
        var all = _comments.All();
        var order = ParagraphOrder();
        (int, int) Pos(CommentInfo c) => CommentStore.Anchor(body, c.Id) is { } a && order.TryGetValue(a.StartParagraph, out var k) ? (k, a.Start) : (int.MaxValue, 0);
        var paraIds = all.Where(c => c.ParaId is not null).Select(c => c.ParaId!).ToHashSet();
        return all.Where(c => c.ParentParaId is null || !paraIds.Contains(c.ParentParaId)).OrderBy(Pos, Comparer<(int, int)>.Create(Cmp)).ToList();
    }

    /// <summary>Commentaire suivant ou précédent dans le texte.</summary>
    internal bool GoToNextComment(bool next)
    {
        if (Body is not { } body) return false;
        var threads = CommentThreads();
        if (threads.Count == 0) { StatusText.Text = "Le document ne contient aucun commentaire."; return false; }
        var order = ParagraphOrder();
        var here = HasSelection && Ordered() is { } sel ? (next ? sel.To : sel.From) : _caret;
        var key = here is { } h ? PosKey(h, order) ?? (0, 0) : (0, 0);
        var positioned = threads.Select(c => (c, a: CommentStore.Anchor(body, c.Id))).Where(x => x.a is not null && order.ContainsKey(x.a.StartParagraph)).ToList();
        if (positioned.Count == 0) return false;
        var target = next
            ? positioned.FirstOrDefault(x => Cmp((order[x.a!.StartParagraph], x.a.Start), key) > 0 || (Cmp((order[x.a.StartParagraph], x.a.Start), key) == 0 && !HasSelection))
            : positioned.LastOrDefault(x => Cmp((order[x.a!.StartParagraph], x.a.Start), key) < 0);
        if (target.c is null) target = next ? positioned[0] : positioned[^1];
        GoToComment(target.c.Id);
        _commentsPane?.Highlight(target.c.Id);
        return true;
    }

    /// <summary>Commentaire dont le texte commenté contient le point d'insertion.</summary>
    internal string? CommentAtCaret()
    {
        if (Body is not { } body || _caret is not { } c) return null;
        var order = ParagraphOrder();
        if (PosKey(c, order) is not { } key) return null;
        foreach (var t in CommentThreads())
            if (CommentStore.Anchor(body, t.Id) is { } a && order.TryGetValue(a.StartParagraph, out var s) && order.TryGetValue(a.EndParagraph, out var e)
                && Cmp((s, a.Start), key) <= 0 && Cmp(key, (e, a.End)) <= 0)
                return t.Id;
        return null;
    }

    /// <summary>Clic droit : nouveau commentaire, supprimer le commentaire du point d'insertion, accepter / refuser la modification.</summary>
    void AddReviewItems(ContextMenu menu)
    {
        MenuItem Item(string header, Action act, string gesture = "")
        {
            var mi = new MenuItem { Header = header, InputGestureText = gesture };
            mi.Click += (_, _) => { act(); };
            menu.Items.Add(mi);
            return mi;
        }
        if (RevisionsAtSelection().Count > 0)
        {
            Item("Accepter la modification", () => ResolveCurrent(true));
            Item("Refuser la modification", () => ResolveCurrent(false));
        }
        Item("Nouveau commentaire", NewComment, "Ctrl+Alt+M");
        if (CommentAtCaret() is { } id)
        {
            Item("Répondre au commentaire", () => { ShowCommentsPane(true); _commentsPane?.BeginReply(id); });
            Item("Supprimer le commentaire", () => DeleteComment(id));
        }
        menu.Items.Add(new Separator());
    }

    /// <summary>Écrit le document tel qu'il serait enregistré, dans un flux (auto-test : copie à faire ouvrir par Word).</summary>
    internal void SavePackageCopy(System.IO.Stream target) => _pkg?.Save(target);

    // ================================================================ volet

    CommentsPane? _commentsPane;

    internal CommentsPane? CommentsPaneForTest => _commentsPane;

    internal bool CommentsPaneShown => _commentsPane is { Visibility: Visibility.Visible };

    internal void ShowCommentsPane(bool show)
    {
        if (show && _commentsPane is null)
        {
            _commentsPane = new CommentsPane(this);
            var grid = (Grid)FormatPane.Parent;
            Grid.SetColumn(_commentsPane, 2);
            grid.Children.Add(_commentsPane);
        }
        if (_commentsPane is null) return;
        _commentsPane.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
        if (show)
        {
            FormatPane.Visibility = Visibility.Collapsed;
            FormatColumn.Width = new GridLength(300);
            _commentsPane.Refresh();
        }
        else
        {
            bool format = FormatPaneToggle?.IsChecked != false;
            FormatPane.Visibility = format ? Visibility.Visible : Visibility.Collapsed;
            FormatColumn.Width = format ? new GridLength(268) : new GridLength(0);
        }
        foreach (var b in _liveButtons.Where(b => Equals(b.Content, "Afficher")))
            b.Background = show ? R("Selected") : Brushes.Transparent;
    }

    void RefreshCommentsPane()
    {
        if (CommentsPaneShown) _commentsPane!.Refresh();
    }
}
