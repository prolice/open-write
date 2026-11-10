// Repris d'OpenPub 2.1.3 (src/OpenPub.App/ParagraphWindow.cs, FindReplaceWindow.cs : mise en forme des boîtes).
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace OpenWrite.App;

/// <summary>Briques communes des boîtes de dialogue (charte d'OpenWrite) : fenêtre, sections, lignes, listes, nombres.</summary>
static class DialogKit
{
    public static readonly CultureInfo Cu = CultureInfo.CurrentCulture;

    public static void Setup(Window w, string title)
    {
        w.Title = title;
        w.SizeToContent = SizeToContent.WidthAndHeight;
        w.ResizeMode = ResizeMode.NoResize;
        w.WindowStartupLocation = WindowStartupLocation.CenterOwner;
        w.ShowInTaskbar = false;
        w.FontFamily = (FontFamily)Application.Current.FindResource("UiFont");
        w.Foreground = (Brush)Application.Current.FindResource("Ink");
        w.Background = (Brush)Application.Current.FindResource("Chrome");
    }

    public static TextBlock Section(string text) => new()
    {
        Text = text.ToUpper(Cu), FontWeight = FontWeights.Bold, FontSize = 11.5, Margin = new Thickness(0, 12, 0, 4),
        Foreground = (Brush)Application.Current.FindResource("Muted"),
    };

    public static StackPanel Row(string label, UIElement control, double labelWidth = 130)
    {
        var sp = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 2, 0, 4) };
        sp.Children.Add(new TextBlock { Text = label, Width = labelWidth, VerticalAlignment = VerticalAlignment.Center });
        sp.Children.Add(control);
        return sp;
    }

    public static TextBox Box(double width = 70) => new() { Width = width, Padding = new Thickness(4, 2, 4, 2), VerticalContentAlignment = VerticalAlignment.Center };

    public static ComboBox Combo(double width, params (string Label, string Tag)[] items)
    {
        var c = new ComboBox { Width = width };
        foreach (var (label, tag) in items) c.Items.Add(new ComboBoxItem { Content = label, Tag = tag });
        c.SelectedIndex = 0;
        return c;
    }

    public static string Tag(ComboBox c) => (string)((ComboBoxItem)c.SelectedItem).Tag;

    public static void Select(ComboBox c, string tag)
    {
        foreach (ComboBoxItem i in c.Items)
            if ((string)i.Tag == tag) { c.SelectedItem = i; return; }
    }

    /// <summary>Nombre saisi (virgule ou point) ; null si illisible.</summary>
    public static double? Parse(string text)
    {
        text = text.Trim().Replace("cm", "").Replace("pt", "").Trim().Replace('.', ',');
        return double.TryParse(text, NumberStyles.Float, CultureInfo.GetCultureInfo("fr-BE"), out var v) ? v : null;
    }

    public static StackPanel Buttons(Window w, Func<bool> ok)
    {
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 16, 0, 0) };
        var cancel = new Button { Content = "Annuler", IsCancel = true, MinWidth = 90, Margin = new Thickness(0, 0, 8, 0) };
        var okButton = new Button { Content = "OK", IsDefault = true, MinWidth = 90, Style = (Style)Application.Current.FindResource("PillButton") };
        okButton.Click += (_, _) => { if (ok()) w.DialogResult = true; };
        buttons.Children.Add(cancel);
        buttons.Children.Add(okButton);
        return buttons;
    }
}
