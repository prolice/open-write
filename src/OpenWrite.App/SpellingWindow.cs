// Repris d'OpenPub 2.1.5 (src/OpenPub.App/SpellingWindow.cs). OpenWrite : langue du mot affichée, Remplacer tout.
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using static OpenWrite.App.DialogKit;

namespace OpenWrite.App;

/// <summary>
/// Vérification de l'orthographe de tout le document (F7, non modale) : chaque mot inconnu est sélectionné dans le
/// texte, avec des suggestions ; Remplacer, Remplacer tout, Ignorer, Ignorer tout, Ajouter au dictionnaire.
/// </summary>
public sealed class SpellingWindow : Window
{
    static readonly Brush Fault = new SolidColorBrush(Color.FromRgb(0xD1, 0x34, 0x38));
    readonly MainWindow _main;
    readonly TextBlock _word = new() { FontSize = 16, FontWeight = FontWeights.Bold, Margin = new Thickness(0, 2, 0, 2) };
    readonly TextBlock _lang = new() { Margin = new Thickness(0, 0, 0, 8) };
    readonly ListBox _suggestions = new() { Height = 150, Width = 260 };
    readonly TextBox _custom = new() { Width = 260, Margin = new Thickness(0, 6, 0, 0) };
    readonly Button _replace, _replaceAll, _ignore, _ignoreAll, _add;
    (string Word, string Lang, Core.Model.Paragraph P, global::OpenWrite.Layout.ParagraphContent C, MainWindow.SpellError E)? _current;

    public SpellingWindow(MainWindow main)
    {
        _main = main;
        Setup(this, "Orthographe");
        _lang.Foreground = (Brush)Application.Current.FindResource("Muted");
        var grid = new Grid { Margin = new Thickness(16) };
        grid.ColumnDefinitions.Add(new ColumnDefinition());
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        var left = new StackPanel();
        left.Children.Add(new TextBlock { Text = "Absent du dictionnaire :", Foreground = (Brush)Application.Current.FindResource("Muted") });
        left.Children.Add(_word);
        left.Children.Add(_lang);
        left.Children.Add(new TextBlock { Text = "Suggestions :", Margin = new Thickness(0, 0, 0, 4) });
        left.Children.Add(_suggestions);
        left.Children.Add(new TextBlock { Text = "Ou remplacer par :", Margin = new Thickness(0, 8, 0, 0) });
        left.Children.Add(_custom);
        grid.Children.Add(left);
        var right = new StackPanel { Margin = new Thickness(14, 20, 0, 0) };
        Button B(string text, bool primary = false)
        {
            var b = new Button { Content = text, MinWidth = 170, Margin = new Thickness(0, 0, 0, 6) };
            if (primary) b.Style = (Style)Application.Current.FindResource("PillButton");
            right.Children.Add(b);
            return b;
        }
        _replace = B("Remplacer", true);
        _replaceAll = B("Remplacer tout");
        _ignore = B("Ignorer");
        _ignoreAll = B("Ignorer tout");
        _add = B("Ajouter au dictionnaire");
        var close = B("Fermer");
        close.IsCancel = true;
        Grid.SetColumn(right, 1);
        grid.Children.Add(right);
        Content = grid;

        _suggestions.SelectionChanged += (_, _) => { if (_suggestions.SelectedItem is string s) _custom.Text = s; };
        _suggestions.MouseDoubleClick += (_, _) => Replace();
        _replace.Click += (_, _) => Replace();
        _replaceAll.Click += (_, _) =>
        {
            if (_current is { } c && _custom.Text.Length > 0 && _custom.Text != c.Word) _main.ReplaceAll(c.Word, _custom.Text, matchCase: true, wholeWord: true);
            Next();
        };
        _ignore.Click += (_, _) => Next();
        _ignoreAll.Click += (_, _) => { if (_current is { } c) { MainWindow.Spell?.Ignore(c.Word); _main.RefreshSpelling(); } Next(); };
        _add.Click += (_, _) => { if (_current is { } c) { MainWindow.Spell?.Add(c.Word, c.Lang); _main.RefreshSpelling(); } Next(); };
        close.Click += (_, _) => Close();
        PreviewKeyDown += (_, e) => { if (e.Key == Key.F1) { Help.HelpService.Open(Help.HelpService.Spelling, this); e.Handled = true; } };
    }

    /// <summary>Mot proposé (auto-test).</summary>
    internal string? CurrentWord => _current?.Word;
    internal string Replacement { get => _custom.Text; set => _custom.Text = value; }

    void Replace()
    {
        if (_current is { } c && _custom.Text.Length > 0 && _custom.Text != c.Word) _main.ReplaceSpellingError(c.P, c.E, c.Word, _custom.Text);
        Next();
    }

    /// <summary>Passe à la faute suivante, ou annonce la fin de la vérification.</summary>
    public void Next()
    {
        _current = _main.NextSpellingError();
        _suggestions.Items.Clear();
        if (_current is not { } c)
        {
            _word.Text = "Vérification terminée";
            _word.Foreground = (Brush)Application.Current.FindResource("Ink");
            _lang.Text = "";
            _custom.Text = "";
            _replace.IsEnabled = _replaceAll.IsEnabled = _ignore.IsEnabled = _ignoreAll.IsEnabled = _add.IsEnabled = false;
            return;
        }
        _word.Text = c.Word;
        _word.Foreground = Fault;
        _lang.Text = "Langue : " + (MainWindow.LanguageChoices().FirstOrDefault(x => x.Tag.Equals(c.Lang, StringComparison.OrdinalIgnoreCase)).Name ?? c.Lang);
        _replace.IsEnabled = _replaceAll.IsEnabled = _ignore.IsEnabled = _ignoreAll.IsEnabled = _add.IsEnabled = true;
        foreach (var s in MainWindow.Spell?.Suggest(c.Word, c.Lang) ?? new List<string>()) _suggestions.Items.Add(s);
        if (_suggestions.Items.Count > 0) _suggestions.SelectedIndex = 0;
        else _custom.Text = c.Word;
    }
}
