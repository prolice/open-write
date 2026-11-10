using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using static OpenWrite.App.DialogKit;

namespace OpenWrite.App;

/// <summary>
/// Fenêtre « Symbole » (Insertion ▸ Symbole) : caractères courants par famille, insérés au point d'insertion ; la
/// fenêtre reste ouverte (comme Word) pour en insérer plusieurs. Le nom Unicode du caractère survolé s'affiche.
/// </summary>
public sealed class SymbolWindow : Window
{
    static readonly (string Family, string Chars)[] Families =
    {
        ("Courants", "€£$¥©®™§¶°±×÷µ‰†‡•…–—«»‹›“”‘’„№℃"),
        ("Espaces et tirets", "  ‑­–—"),
        ("Lettres accentuées", "ÀÂÄÆÇÉÈÊËÎÏÔŒÙÛÜŸàâäæçéèêëîïôœùûüÿÁÍÓÚÑáíóúñßØøÅåĳ"),
        ("Flèches", "←→↑↓↔↕⇐⇒⇔↵➔"),
        ("Mathématiques", "≈≠≤≥∞√∑∏∫∂∆∇∈∉∩∪⊂⊃∀∃¬∧∨½¼¾⅓⅔¹²³ⁿ"),
        ("Grec", "αβγδεζηθικλμνξοπρστυφχψωΓΔΘΛΞΠΣΦΨΩ"),
        ("Cases et puces", "☐☑☒✓✔✗✘■□▪▫●○◆◇★☆♦♣♠♥"),
    };

    static readonly Dictionary<char, string> Names = new()
    {
        [' '] = "Espace insécable (Ctrl+Maj+Espace)", [' '] = "Espace fine insécable (avant ; : ! ?)",
        ['‑'] = "Trait d'union insécable (Ctrl+Maj+-)", ['­'] = "Trait d'union conditionnel", ['–'] = "Tiret demi-cadratin",
        ['—'] = "Tiret cadratin", ['€'] = "Euro", ['©'] = "Copyright", ['®'] = "Marque déposée", ['™'] = "Marque commerciale",
        ['§'] = "Paragraphe (section)", ['¶'] = "Pied-de-mouche", ['°'] = "Degré", ['…'] = "Points de suspension",
        ['«'] = "Guillemet ouvrant", ['»'] = "Guillemet fermant", ['№'] = "Numéro",
    };

    readonly TextBlock _info = new() { Margin = new Thickness(0, 8, 0, 0), Height = 18 };

    public SymbolWindow(Action<string> insert)
    {
        Setup(this, "Symbole");
        _info.Foreground = (Brush)Application.Current.FindResource("Muted");
        var root = new StackPanel { Margin = new Thickness(16), Width = 520 };
        foreach (var (family, chars) in Families)
        {
            root.Children.Add(Section(family));
            var grid = new UniformGrid { Columns = 16 };
            foreach (var ch in chars)
            {
                var label = ch switch { ' ' => "°̲", ' ' => "·", '­' => "¬", '‑' => "‑", _ => ch.ToString() };
                var b = new Button
                {
                    Content = label, FontSize = 16, Height = 30, Margin = new Thickness(1), Padding = new Thickness(0),
                    FontFamily = new FontFamily("Segoe UI, Segoe UI Symbol"), Focusable = false,
                };
                var name = Names.TryGetValue(ch, out var n) ? n : $"U+{(int)ch:X4}";
                b.ToolTip = name;
                b.MouseEnter += (_, _) => _info.Text = $"{name}  —  U+{(int)ch:X4}";
                b.Click += (_, _) => insert(ch.ToString());
                grid.Children.Add(b);
            }
            root.Children.Add(grid);
        }
        root.Children.Add(_info);
        var close = new Button { Content = "Fermer", IsCancel = true, MinWidth = 90, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 10, 0, 0) };
        close.Click += (_, _) => Close();
        root.Children.Add(close);
        Content = root;
    }
}
