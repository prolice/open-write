// Repris d'OpenPub 2.1.4 (OpenPub.App/Output/OutputUi.cs).
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using OpenWrite.Core.Output;

namespace OpenWrite.App.Output;

/// <summary>Elements communs des boites de sortie (impression, PDF, images), avec les ressources de la charte.</summary>
static class OutputUi
{
    public static T Res<T>(string key) => (T)Application.Current.FindResource(key);

    public static void Style(Window w, string title, string helpTopic)
    {
        // F1 : aide en ligne de la boite
        w.KeyDown += (_, e) => { if (e.Key == System.Windows.Input.Key.F1) { Help.HelpService.Open(helpTopic); e.Handled = true; } };
        w.Title = title;
        w.WindowStartupLocation = WindowStartupLocation.CenterOwner;
        w.ShowInTaskbar = false;
        w.FontFamily = Res<FontFamily>("UiFont");
        w.Foreground = Res<Brush>("Ink");
        w.Background = Res<Brush>("Chrome");
    }

    public static TextBlock Section(string text) => new()
    {
        Text = text, FontWeight = FontWeights.Bold, FontSize = 12.5, Margin = new Thickness(0, 12, 0, 6),
        Foreground = Res<Brush>("SelectedText"),
    };

    public static TextBlock Note(string text) => new()
    {
        Text = text, FontSize = 11.5, TextWrapping = TextWrapping.Wrap, Foreground = Res<Brush>("Muted"), Margin = new Thickness(0, 4, 0, 0),
    };

    public static Grid Row(string label, UIElement control, double labelWidth = 130)
    {
        var g = new Grid { Margin = new Thickness(0, 2, 0, 2) };
        g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(labelWidth) });
        g.ColumnDefinitions.Add(new ColumnDefinition());
        g.Children.Add(new TextBlock { Text = label, VerticalAlignment = VerticalAlignment.Center });
        Grid.SetColumn(control, 1);
        if (control is FrameworkElement fe) fe.HorizontalAlignment = HorizontalAlignment.Left;
        g.Children.Add(control);
        return g;
    }

    public static TextBox Box(string text, double width = 70) => new() { Text = text, Width = width, VerticalContentAlignment = VerticalAlignment.Center };

    public static ComboBox Combo(params (string Label, object Tag)[] items)
    {
        var c = new ComboBox { MinWidth = 200 };
        foreach (var (l, t) in items) c.Items.Add(new ComboBoxItem { Content = l, Tag = t });
        c.SelectedIndex = 0;
        return c;
    }

    public static object? Tag(ComboBox c) => (c.SelectedItem as ComboBoxItem)?.Tag;

    /// <summary>Boutons du bas : Annuler et l'action principale (bouton pilule).</summary>
    public static StackPanel Buttons(Window w, string okText, Func<bool> onOk)
    {
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 16, 0, 0) };
        var cancel = new Button { Content = "Annuler", IsCancel = true, MinWidth = 90, Margin = new Thickness(0, 0, 8, 0) };
        var ok = new Button { Content = okText, IsDefault = true, MinWidth = 110, Style = Res<Style>("PillButton") };
        ok.Click += (_, _) => { if (onOk()) w.DialogResult = true; };
        buttons.Children.Add(cancel);
        buttons.Children.Add(ok);
        return buttons;
    }

    public static double? Number(string text)
    {
        var cu = CultureInfo.CurrentCulture;
        var sep = cu.NumberFormat.NumberDecimalSeparator;
        return double.TryParse(text.Trim().Replace(".", sep).Replace(",", sep), NumberStyles.Float, cu, out var v) ? v : null;
    }
}

/// <summary>Choix des pages : toutes, page active, ou liste « 1-3; 5 ».</summary>
sealed class PageChooser : StackPanel
{
    readonly RadioButton _all = new() { Content = "Toutes les pages", IsChecked = true, Margin = new Thickness(0, 2, 0, 2) };
    readonly RadioButton _active;
    readonly RadioButton _range = new() { Content = "Pages :", VerticalAlignment = VerticalAlignment.Center };
    readonly TextBox _text = OutputUi.Box("", 150);
    readonly int _count, _activeNumber;
    public event Action? Changed;

    public PageChooser(int pageCount, int activeNumber)
    {
        _count = pageCount;
        _activeNumber = Math.Clamp(activeNumber, 1, Math.Max(1, pageCount));
        _active = new RadioButton { Content = $"Page active ({_activeNumber})", Margin = new Thickness(0, 2, 0, 2) };
        Children.Add(_all);
        Children.Add(_active);
        var row = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 2, 0, 2) };
        _text.Margin = new Thickness(8, 0, 0, 0);
        _text.ToolTip = "Exemple : 1-3; 5";
        row.Children.Add(_range);
        row.Children.Add(_text);
        Children.Add(row);
        Children.Add(OutputUi.Note($"Le document a {pageCount} page(s). Exemple : 1-3; 5"));
        _text.GotKeyboardFocus += (_, _) => _range.IsChecked = true;
        foreach (var r in new[] { _all, _active, _range }) r.Checked += (_, _) => Changed?.Invoke();
        _text.TextChanged += (_, _) => { if (_range.IsChecked == true) Changed?.Invoke(); };
    }

    public bool ActiveOnly => _active.IsChecked == true;

    public void SelectActive() => _active.IsChecked = true;

    public void SelectRange(string text) { _text.Text = text; _range.IsChecked = true; }

    public List<int>? Pages(out string? error)
    {
        error = null;
        if (_all.IsChecked == true) return Enumerable.Range(1, _count).ToList();
        if (_active.IsChecked == true) return new List<int> { _activeNumber };
        return OpenWrite.Core.Output.PageRange.Parse(_text.Text, _count, out error);
    }
}
