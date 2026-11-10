// Repris d'OpenPub 2.1.3 (src/OpenPub.App/FindReplaceWindow.cs).
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using static OpenWrite.App.DialogKit;

namespace OpenWrite.App;

/// <summary>Fenêtre « Rechercher et remplacer » (non modale, Ctrl+F / Ctrl+H) : parcourt le texte du document.</summary>
public sealed class FindReplaceWindow : Window
{
    readonly MainWindow _main;
    readonly TextBox _query = Box(280), _replace = Box(280);
    readonly CheckBox _case = new() { Content = "Respecter la casse et les accents", Margin = new Thickness(0, 6, 0, 0) };
    readonly CheckBox _word = new() { Content = "Mot entier", Margin = new Thickness(0, 4, 0, 0) };
    readonly StackPanel _replaceRow;
    readonly Button _replaceOne, _replaceAll;
    readonly TextBlock _info = new() { Margin = new Thickness(0, 8, 0, 0), TextWrapping = TextWrapping.Wrap, MaxWidth = 400 };

    public FindReplaceWindow(MainWindow main)
    {
        _main = main;
        Setup(this, "Rechercher et remplacer");
        _info.Foreground = (System.Windows.Media.Brush)Application.Current.FindResource("Muted");
        var root = new StackPanel { Margin = new Thickness(16) };
        root.Children.Add(Row("Rechercher", _query, 110));
        _replaceRow = Row("Remplacer par", _replace, 110);
        root.Children.Add(_replaceRow);
        root.Children.Add(_case);
        root.Children.Add(_word);
        root.Children.Add(_info);
        var buttons = new WrapPanel { Margin = new Thickness(0, 12, 0, 0), HorizontalAlignment = HorizontalAlignment.Right };
        var next = new Button { Content = "Suivant", IsDefault = true, MinWidth = 90, Style = (Style)Application.Current.FindResource("PillButton") };
        _replaceOne = new Button { Content = "Remplacer", MinWidth = 90, Margin = new Thickness(6, 0, 0, 0) };
        _replaceAll = new Button { Content = "Remplacer tout", MinWidth = 110, Margin = new Thickness(6, 0, 0, 0) };
        var close = new Button { Content = "Fermer", IsCancel = true, MinWidth = 80, Margin = new Thickness(6, 0, 0, 0) };
        buttons.Children.Add(next); buttons.Children.Add(_replaceOne); buttons.Children.Add(_replaceAll); buttons.Children.Add(close);
        root.Children.Add(buttons);
        Content = root;

        next.Click += (_, _) => _info.Text = _main.FindNext(_query.Text, MatchCase, Whole) ? "" : $"« {_query.Text} » est introuvable.";
        _replaceOne.Click += (_, _) => { _main.ReplaceCurrent(_query.Text, _replace.Text, MatchCase, Whole); _info.Text = ""; };
        _replaceAll.Click += (_, _) =>
        {
            int n = _main.ReplaceAll(_query.Text, _replace.Text, MatchCase, Whole);
            _info.Text = n == 0 ? $"« {_query.Text} » est introuvable." : $"{n} remplacement(s) effectué(s). Ctrl+Z pour annuler.";
        };
        close.Click += (_, _) => Close();
        PreviewKeyDown += (_, e) => { if (e.Key == Key.F1) { Help.HelpService.Open(Help.HelpService.FindReplace, this); e.Handled = true; } };
    }

    bool MatchCase => _case.IsChecked == true;
    bool Whole => _word.IsChecked == true;

    public void Prepare(bool replace, string? initial)
    {
        Title = replace ? "Rechercher et remplacer" : "Rechercher";
        _replaceRow.Visibility = _replaceOne.Visibility = _replaceAll.Visibility = replace ? Visibility.Visible : Visibility.Collapsed;
        if (initial is not null) _query.Text = initial;
        _info.Text = "";
        Dispatcher.BeginInvoke(() => { _query.Focus(); _query.SelectAll(); });
    }
}
