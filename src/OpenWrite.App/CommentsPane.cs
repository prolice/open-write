using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using OpenWrite.Docx;

namespace OpenWrite.App;

/// <summary>
/// Volet Commentaires (lot 8, comme la liste des commentaires de Word 365) : un fil par commentaire, dans l'ordre du
/// texte (auteur en couleur, date, texte, réponses), résolu grisé ; Répondre, Résoudre / Rouvrir, Modifier, Supprimer ;
/// un clic sur une carte sélectionne le texte commenté. Nouveau commentaire : une carte à remplir en haut du volet.
/// </summary>
public sealed class CommentsPane : Border
{
    readonly MainWindow _main;
    readonly StackPanel _list = new();
    readonly ScrollViewer _scroll;
    string? _highlighted;
    static readonly CultureInfo Fr = CultureInfo.GetCultureInfo("fr-FR");

    static Brush Res(string key) => (Brush)Application.Current.FindResource(key);

    public CommentsPane(MainWindow main)
    {
        _main = main;
        Background = Brushes.White;
        BorderBrush = Res("Line");
        BorderThickness = new Thickness(1, 0, 0, 0);
        var root = new DockPanel { Margin = new Thickness(14, 6, 14, 10) };
        var head = new DockPanel { Margin = new Thickness(0, 8, 0, 8) };
        var close = new Button { Style = (Style)Application.Current.FindResource("RSmall"), Tag = "", ToolTip = "Fermer le volet (Révision ▸ Afficher pour le rouvrir)", Padding = new Thickness(4, 2, 4, 2) };
        close.Click += (_, _) => _main.ShowCommentsPane(false);
        DockPanel.SetDock(close, Dock.Right);
        head.Children.Add(close);
        head.Children.Add(new TextBlock { Text = "Commentaires", FontFamily = (FontFamily)Application.Current.FindResource("TitleFont"), FontSize = 16, FontWeight = FontWeights.ExtraBold, VerticalAlignment = VerticalAlignment.Center });
        DockPanel.SetDock(head, Dock.Top);
        root.Children.Add(head);
        var add = new Button { Content = "Nouveau commentaire", Style = (Style)Application.Current.FindResource("PillButton"), HorizontalAlignment = HorizontalAlignment.Left, Margin = new Thickness(0, 0, 0, 10), ToolTip = "Commenter la sélection (Ctrl+Alt+M)" };
        add.Click += (_, _) => _main.NewComment();
        DockPanel.SetDock(add, Dock.Top);
        root.Children.Add(add);
        _scroll = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Content = _list };
        root.Children.Add(_scroll);
        Child = root;
    }

    /// <summary>Cartes affichées (auto-test).</summary>
    internal int Cards => _list.Children.OfType<Border>().Count(b => b.Tag is string);

    /// <summary>Carte en cours de saisie (nouveau commentaire, réponse ou modification) : zone de texte et bouton.</summary>
    internal (TextBox Box, Button Ok)? Editor { get; private set; }

    static string Date(CommentInfo c) => c.Date is { } d ? d.ToString("d MMM yyyy, HH:mm", Fr) : "";

    Border Card(UIElement content, string? id, Brush accent, bool dim = false) => new()
    {
        Child = content, Tag = id, Background = dim ? new SolidColorBrush(Color.FromRgb(0xF6, 0xF6, 0xF8)) : Brushes.White,
        BorderBrush = id is not null && id == _highlighted ? Res("Accent") : Res("Line"), BorderThickness = new Thickness(id is not null && id == _highlighted ? 2 : 1),
        CornerRadius = new CornerRadius(8), Padding = new Thickness(10, 8, 10, 8), Margin = new Thickness(0, 0, 0, 10),
        Effect = new System.Windows.Media.Effects.DropShadowEffect { BlurRadius = 6, ShadowDepth = 1, Opacity = 0.10 },
    };

    static Brush AuthorBrush(string author)
    {
        // même couleur que les révisions de cet auteur quand le document en a
        var hex = new[] { "B5082E", "2E74B5", "538135", "BF8F00", "7030A0", "C55A11", "1F7A7A", "A5276B" }[(int)((uint)StableHash(author) % 8)];
        return new SolidColorBrush(Color.FromRgb(Convert.ToByte(hex[..2], 16), Convert.ToByte(hex[2..4], 16), Convert.ToByte(hex[4..], 16)));
    }

    static int StableHash(string s) { unchecked { int h = 23; foreach (var ch in s) h = h * 31 + ch; return h; } }

    static UIElement Header(CommentInfo c)
    {
        var sp = new StackPanel { Orientation = Orientation.Horizontal };
        var initials = c.Initials is { Length: > 0 } i ? i : c.Author.Length > 0 ? c.Author[..1].ToUpperInvariant() : "?";
        sp.Children.Add(new Border
        {
            Width = 26, Height = 26, CornerRadius = new CornerRadius(13), Background = AuthorBrush(c.Author), Margin = new Thickness(0, 0, 8, 0),
            Child = new TextBlock { Text = initials, Foreground = Brushes.White, FontSize = 10.5, FontWeight = FontWeights.Bold, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center },
        });
        var names = new StackPanel();
        names.Children.Add(new TextBlock { Text = c.Author.Length > 0 ? c.Author : "(auteur inconnu)", FontWeight = FontWeights.SemiBold, FontSize = 12.5 });
        names.Children.Add(new TextBlock { Text = Date(c), FontSize = 11, Foreground = Res("Muted") });
        sp.Children.Add(names);
        return sp;
    }

    static TextBlock Body(string text) => new() { Text = text, TextWrapping = TextWrapping.Wrap, FontSize = 12.5, Margin = new Thickness(0, 6, 0, 0) };

    Button Link(string text, Action act, string? tip = null)
    {
        var b = new Button { Content = text, Margin = new Thickness(0, 0, 10, 0), Padding = new Thickness(0), Background = Brushes.Transparent, BorderThickness = new Thickness(0), Foreground = Res("Accent"), Cursor = Cursors.Hand, FontSize = 12, ToolTip = tip };
        b.Click += (_, e) => { e.Handled = true; act(); };
        return b;
    }

    (StackPanel Panel, TextBox Box, Button Ok) EditorPanel(string initial, string okText, Action<string> ok, Action cancel)
    {
        var box = new TextBox { Text = initial, AcceptsReturn = true, TextWrapping = TextWrapping.Wrap, MinHeight = 54, MaxHeight = 160, VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Margin = new Thickness(0, 6, 0, 6), Padding = new Thickness(4) };
        var row = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        var no = new Button { Content = "Annuler", MinWidth = 70, Margin = new Thickness(0, 0, 6, 0) };
        var yes = new Button { Content = okText, MinWidth = 80, Style = (Style)Application.Current.FindResource("PillButton") };
        no.Click += (_, _) => { Editor = null; cancel(); };
        yes.Click += (_, _) => { Editor = null; ok(box.Text); };
        box.KeyDown += (_, e) => { if (e.Key == Key.Enter && Keyboard.Modifiers == ModifierKeys.Control) { e.Handled = true; Editor = null; ok(box.Text); } else if (e.Key == Key.Escape) { e.Handled = true; Editor = null; cancel(); } };
        row.Children.Add(no);
        row.Children.Add(yes);
        var panel = new StackPanel();
        panel.Children.Add(box);
        panel.Children.Add(row);
        Editor = (box, yes);
        Dispatcher.BeginInvoke(() => { box.Focus(); box.CaretIndex = box.Text.Length; }, System.Windows.Threading.DispatcherPriority.Input);
        return (panel, box, yes);
    }

    /// <summary>Carte vide en haut du volet pour un nouveau commentaire.</summary>
    public void BeginNew()
    {
        Refresh();
        var sp = new StackPanel();
        sp.Children.Add(Header(new CommentInfo("", MainWindow.UserName, MainWindow.UserInitials, null, new StoryHolder { Name = "comment" }, null, null, false)));
        var (panel, _, _) = EditorPanel("", "Publier", text => { _main.PublishNewComment(text); Refresh(); }, Refresh);
        sp.Children.Add(panel);
        _list.Children.Insert(0, Card(sp, null, Res("Accent")));
        _scroll.ScrollToTop();
    }

    string? _replyTo, _editing;

    public void BeginReply(string id) { _replyTo = id; _editing = null; Refresh(); }

    public void Highlight(string id) { _highlighted = id; Refresh(); }

    /// <summary>Refait la liste à partir du document.</summary>
    public void Refresh()
    {
        Editor = null;
        _list.Children.Clear();
        var threads = _main.CommentThreads();
        var all = _main.Comments?.All() ?? new List<CommentInfo>();
        if (threads.Count == 0)
        {
            _list.Children.Add(new TextBlock { Text = "Aucun commentaire. Sélectionnez du texte puis « Nouveau commentaire » (Ctrl+Alt+M).", TextWrapping = TextWrapping.Wrap, Foreground = Res("Muted"), FontSize = 12 });
            return;
        }
        foreach (var t in threads)
        {
            var replies = new List<CommentInfo>();
            void Collect(CommentInfo c)
            {
                foreach (var r in all.Where(x => x.ParaId is not null && c.ParaId is not null && x.ParentParaId == c.ParaId)) { replies.Add(r); Collect(r); }
            }
            Collect(t);
            var sp = new StackPanel();
            sp.Children.Add(Header(t));
            if (t.Done) sp.Children.Add(new TextBlock { Text = "✓ Résolu", FontSize = 11, Foreground = Res("Muted"), Margin = new Thickness(0, 4, 0, 0) });
            void Show(CommentInfo c, bool reply)
            {
                if (_editing == c.Id)
                {
                    sp.Children.Add(EditorPanel(c.Text, "Enregistrer", text => { _editing = null; _main.EditComment(c.Id, text); }, () => { _editing = null; Refresh(); }).Panel);
                    return;
                }
                if (reply)
                {
                    var head = Header(c);
                    ((FrameworkElement)head).Margin = new Thickness(0, 10, 0, 0);
                    sp.Children.Add(head);
                }
                sp.Children.Add(Body(c.Text));
            }
            Show(t, false);
            foreach (var r in replies) Show(r, true);
            if (_replyTo == t.Id || replies.Any(r => r.Id == _replyTo))
                sp.Children.Add(EditorPanel("", "Répondre", text => { _replyTo = null; _main.ReplyToComment(t.Id, text); }, () => { _replyTo = null; Refresh(); }).Panel);
            var actions = new WrapPanel { Margin = new Thickness(0, 8, 0, 0) };
            actions.Children.Add(Link("Répondre", () => BeginReply(t.Id)));
            actions.Children.Add(Link(t.Done ? "Rouvrir" : "Résoudre", () => _main.ResolveComment(t.Id, !t.Done), t.Done ? "Rouvrir le fil" : "Marquer le fil comme résolu"));
            if (t.PlainText) actions.Children.Add(Link("Modifier", () => { _editing = t.Id; _replyTo = null; Refresh(); }));
            actions.Children.Add(Link("Supprimer", () => _main.DeleteComment(t.Id), replies.Count > 0 ? "Supprimer le fil (commentaire et réponses)" : "Supprimer le commentaire"));
            sp.Children.Add(actions);
            var card = Card(sp, t.Id, AuthorBrush(t.Author), dim: t.Done);
            card.Opacity = t.Done ? 0.75 : 1;
            card.Cursor = Cursors.Hand;
            card.MouseLeftButtonUp += (_, e) =>
            {
                if (e.OriginalSource is DependencyObject d && FindParent<TextBox>(d) is not null) return;
                _highlighted = t.Id;
                _main.GoToComment(t.Id);
                foreach (var b in _list.Children.OfType<Border>())
                {
                    bool on = b.Tag as string == t.Id;
                    b.BorderBrush = on ? Res("Accent") : Res("Line");
                    b.BorderThickness = new Thickness(on ? 2 : 1);
                }
            };
            _list.Children.Add(card);
        }
    }

    static T? FindParent<T>(DependencyObject d) where T : DependencyObject
    {
        for (var x = d; x is not null; x = VisualTreeHelper.GetParent(x)) if (x is T t) return t;
        return null;
    }
}
