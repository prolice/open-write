using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using static OpenWrite.App.DialogKit;

namespace OpenWrite.App;

/// <summary>
/// Révision ▸ Langue ▸ Définir la langue de vérification : langue du texte sélectionné (w:lang), dictionnaire installé
/// ou non, « Ne pas vérifier l'orthographe » (w:noProof), comme la boîte Langue de Word.
/// </summary>
public sealed class LanguageWindow : Window
{
    readonly ListBox _list = new() { Width = 300, Height = 260 };
    readonly CheckBox _noProof = new() { Content = "Ne pas vérifier l'orthographe", Margin = new Thickness(0, 10, 0, 0) };

    public string? ChosenTag { get; private set; }
    public bool NoProof => _noProof.IsChecked == true;

    public LanguageWindow(IReadOnlyList<(string Tag, string Name, bool Checked)> choices, string current, bool noProof)
    {
        Setup(this, "Langue");
        var root = new StackPanel { Margin = new Thickness(16) };
        root.Children.Add(new TextBlock { Text = "Marquer le texte sélectionné comme :", Margin = new Thickness(0, 0, 0, 6) });
        foreach (var (tag, name, ok) in choices)
            _list.Items.Add(new ListBoxItem { Content = (ok ? "✓  " : "     ") + name, Tag = tag, ToolTip = ok ? $"{tag} : dictionnaire installé" : $"{tag} : pas de dictionnaire dans Windows, le texte ne sera pas vérifié" });
        _list.SelectedItem = _list.Items.OfType<ListBoxItem>().FirstOrDefault(i => string.Equals((string)i.Tag, current, StringComparison.OrdinalIgnoreCase))
                             ?? _list.Items.OfType<ListBoxItem>().FirstOrDefault(i => ((string)i.Tag).Split('-')[0].Equals(current.Split('-')[0], StringComparison.OrdinalIgnoreCase));
        _list.Loaded += (_, _) => { if (_list.SelectedItem is not null) _list.ScrollIntoView(_list.SelectedItem); };
        root.Children.Add(_list);
        _noProof.IsChecked = noProof;
        root.Children.Add(_noProof);
        var note = new TextBlock
        {
            Text = "✓ : dictionnaire présent dans Windows. Pour en ajouter : Paramètres ▸ Heure et langue ▸ Langue.",
            TextWrapping = TextWrapping.Wrap, MaxWidth = 300, Margin = new Thickness(0, 8, 0, 0),
            Foreground = (System.Windows.Media.Brush)Application.Current.FindResource("Muted"),
        };
        root.Children.Add(note);
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 14, 0, 0) };
        var cancel = new Button { Content = "Annuler", IsCancel = true, MinWidth = 90, Margin = new Thickness(0, 0, 8, 0) };
        var ok2 = new Button { Content = "OK", IsDefault = true, MinWidth = 90, Style = (Style)Application.Current.FindResource("PillButton") };
        ok2.Click += (_, _) => Accept();
        _list.MouseDoubleClick += (_, _) => Accept();
        buttons.Children.Add(cancel);
        buttons.Children.Add(ok2);
        root.Children.Add(buttons);
        Content = root;
        PreviewKeyDown += (_, e) => { if (e.Key == Key.F1) { Help.HelpService.Open(Help.HelpService.Spelling, this); e.Handled = true; } };
    }

    void Accept()
    {
        if (_list.SelectedItem is not ListBoxItem it) return;
        ChosenTag = (string)it.Tag;
        DialogResult = true;
    }
}
