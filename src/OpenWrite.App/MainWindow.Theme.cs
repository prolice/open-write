using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using OpenWrite.Core.Styles;

namespace OpenWrite.App;

// Lot 4 (version 0.5) : Création ▸ Thèmes, Polices, Couleurs. Les polices et couleurs du thème (theme1.xml) changent,
// et avec elles les titres, citations… qui les utilisent ; annulable (instantané XML du journal).
public partial class MainWindow
{
    System.Xml.Linq.XElement? ThemeRoot() => _pkg?.XmlPartOf("theme")?.Doc.Root;

    /// <summary>Thèmes : combinaison de polices et de couleurs.</summary>
    static readonly (string Name, int Fonts, int Colors)[] Themes =
    {
        ("Office (Aptos)", 0, 0), ("Office 2013-2022 (Calibri)", 1, 1), ("Office 2007-2010 (Cambria)", 2, 2),
        ("Classique (Georgia, bleu)", 5, 3), ("Nature (Candara, vert)", 10, 4), ("Chaleureux (Garamond, rouge orangé)", 6, 5),
        ("Élégant (Century Gothic, violet)", 7, 6), ("Sobre (Arial, gris)", 3, 7),
    };

    void ChangeTheme(string label, Action<System.Xml.Linq.XElement> change)
    {
        if (Body is not { } body || ThemeRoot() is not { } theme)
        {
            StatusText.Text = "Ce document n'a pas de thème : polices et couleurs du thème indisponibles.";
            return;
        }
        var caret = _caret is { } c ? ToDocPos(c) : null;
        var (anchor, caretPos) = SelectionDocPos();
        using (var tx = _journal.Begin(label, body, Array.Empty<int>(), 0, 0, caret))
        {
            tx.TrackXml(theme);
            change(theme);
            tx.Commit(caret);
        }
        AfterFormat(label, anchor, caretPos);
    }

    internal void SetThemeFonts(int pair)
    {
        var (name, major, minor) = ThemeOps.FontPairs[pair];
        ChangeTheme($"Polices du thème : {name}", t => ThemeOps.SetFonts(t, name, major, minor));
    }

    internal void SetThemeColors(int palette)
    {
        var (name, colors) = ThemeOps.Palettes[palette];
        ChangeTheme($"Couleurs du thème : {name}", t => ThemeOps.SetColors(t, name, colors));
    }

    /// <summary>Menu sous un bouton du ruban.</summary>
    void ShowMenu(FrameworkElement anchor, IEnumerable<(string Header, UIElement? Icon, Action Act)> items)
    {
        var menu = new ContextMenu { PlacementTarget = anchor, Placement = PlacementMode.Bottom };
        foreach (var (header, icon, act) in items)
        {
            var mi = new MenuItem { Header = header, Icon = icon };
            mi.Click += (_, _) => { act(); _pagesView?.Focus(); };
            menu.Items.Add(mi);
        }
        menu.IsOpen = true;
    }

    static UIElement Swatches(string[] colors)
    {
        var sp = new StackPanel { Orientation = Orientation.Horizontal };
        foreach (var c in colors.Skip(2).Take(6))
            sp.Children.Add(new Border { Width = 10, Height = 14, Margin = new Thickness(0, 0, 1, 0), Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#" + c)) });
        return sp;
    }

    internal void ShowThemesMenu(FrameworkElement anchor) => ShowMenu(anchor, Themes.Select(t => (t.Name, (UIElement?)Swatches(ThemeOps.Palettes[t.Colors].Colors), (Action)(() =>
    {
        var (fname, major, minor) = ThemeOps.FontPairs[t.Fonts];
        var (cname, colors) = ThemeOps.Palettes[t.Colors];
        ChangeTheme($"Thème : {t.Name}", th => { ThemeOps.SetFonts(th, fname, major, minor); ThemeOps.SetColors(th, cname, colors); });
    }))));

    internal void ShowThemeFontsMenu(FrameworkElement anchor) => ShowMenu(anchor, ThemeOps.FontPairs.Select((f, i) =>
        ($"{f.Name} — {f.Major} / {f.Minor}", (UIElement?)null, (Action)(() => SetThemeFonts(i)))));

    internal void ShowThemeColorsMenu(FrameworkElement anchor) => ShowMenu(anchor, ThemeOps.Palettes.Select((p, i) =>
        (p.Name, (UIElement?)Swatches(p.Colors), (Action)(() => SetThemeColors(i)))));
}
