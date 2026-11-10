using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;
using static OpenWrite.App.DialogKit;

namespace OpenWrite.App;

/// <summary>
/// Boîte « Police » (Ctrl+D, comme Word) : police, style, taille, soulignement, couleur, effets. Initialisée depuis la
/// mise en forme du premier caractère visé ; seuls les réglages <strong>changés</strong> sont appliqués (une sélection
/// de plusieurs polices garde ses polices si l'on ne change que la couleur).
/// </summary>
public sealed class FontWindow : Window
{
    readonly ComboBox _family = new() { Width = 220, IsEditable = true, IsTextSearchEnabled = true };
    readonly ComboBox _style = Combo(140, ("Normal", "0"), ("Italique", "i"), ("Gras", "b"), ("Gras italique", "bi"));
    readonly ComboBox _size = new() { Width = 70, IsEditable = true };
    readonly ComboBox _underline = Combo(160, ("(aucun)", "none"), ("Simple", "single"), ("Mots seuls", "words"), ("Double", "double"),
        ("Épais", "thick"), ("Pointillé", "dotted"), ("Tirets", "dash"), ("Ondulé", "wave"));
    readonly ComboBox _color = Combo(160, ("Automatique", "auto"), ("Noir", "000000"), ("Gris foncé", "595959"), ("Rouge foncé", "C00000"),
        ("Rouge", "FF0000"), ("Orange", "FFC000"), ("Jaune", "FFFF00"), ("Vert clair", "92D050"), ("Vert", "00B050"),
        ("Bleu clair", "00B0F0"), ("Bleu", "0070C0"), ("Bleu foncé", "002060"), ("Violet", "7030A0"));
    readonly CheckBox _strike = new() { Content = "Barré" }, _dstrike = new() { Content = "Barré double" },
        _sup = new() { Content = "Exposant" }, _sub = new() { Content = "Indice" },
        _small = new() { Content = "Petites majuscules" }, _caps = new() { Content = "Majuscules" },
        _hidden = new() { Content = "Masqué" };
    readonly TextBlock _preview = new() { Text = "Aperçu du texte", FontSize = 20, Margin = new Thickness(12), HorizontalAlignment = HorizontalAlignment.Center };
    readonly RunFormat _initial;
    readonly Dictionary<Control, string> _start = new();

    /// <summary>Modification à appliquer à chaque segment visé (null si annulé ou inchangé).</summary>
    public Action<Paragraph, Run>? Result { get; private set; }

    /// <param name="resolve">Mise en forme résolue d'un segment (pour poser « 0 » quand un style donne encore l'attribut retiré).</param>
    public FontWindow(RunFormat initial, Func<Paragraph, Run, RunFormat>? resolve = null)
    {
        Setup(this, "Police");
        _initial = initial;
        _family.ItemsSource = Fonts.SystemFontFamilies.Select(f => f.Source).Distinct().OrderBy(n => n, StringComparer.CurrentCultureIgnoreCase).ToList();
        _size.ItemsSource = new[] { 8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 }.Select(s => s.ToString(Cu)).ToList();

        var root = new StackPanel { Margin = new Thickness(18), Width = 480 };
        root.Children.Add(Section("Police"));
        root.Children.Add(Row("Police", _family));
        root.Children.Add(Row("Style", _style));
        root.Children.Add(Row("Taille (pt)", _size));
        root.Children.Add(Row("Soulignement", _underline));
        root.Children.Add(Row("Couleur", _color));
        root.Children.Add(Section("Effets"));
        var effects = new UniformGrid2();
        foreach (var cb in new[] { _strike, _dstrike, _sup, _sub, _small, _caps, _hidden }) { cb.Margin = new Thickness(0, 3, 12, 3); effects.Add(cb); }
        root.Children.Add(effects.Panel);
        root.Children.Add(Section("Aperçu"));
        root.Children.Add(new Border
        {
            Background = Brushes.White, BorderBrush = (Brush)Application.Current.FindResource("Line"), BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(8), Height = 70, Child = _preview,
        });
        root.Children.Add(Buttons(this, Build));
        Content = root;

        _family.Text = initial.FontAscii ?? initial.FontHAnsi ?? "";
        Select(_style, (initial.Bold ? "b" : "") + (initial.Italic ? "i" : "") is { Length: > 0 } st ? st : "0");
        _size.Text = initial.Size.ToString("0.#", Cu);
        Select(_underline, initial.Underline is null or "none" ? "none" : initial.Underline);
        if (_underline.SelectedIndex < 0) _underline.SelectedIndex = 1;
        Select(_color, initial.Color ?? "auto");
        if (_color.SelectedIndex < 0)
        {
            _color.Items.Add(new ComboBoxItem { Content = $"Actuelle (#{initial.Color})", Tag = initial.Color });
            _color.SelectedIndex = _color.Items.Count - 1;
        }
        _strike.IsChecked = initial.Strike;
        _dstrike.IsChecked = initial.DoubleStrike;
        _sup.IsChecked = initial.VerticalAlign == "superscript";
        _sub.IsChecked = initial.VerticalAlign == "subscript";
        _small.IsChecked = initial.SmallCaps;
        _caps.IsChecked = initial.Caps;
        _hidden.IsChecked = initial.Hidden;
        _sup.Checked += (_, _) => _sub.IsChecked = false;
        _sub.Checked += (_, _) => _sup.IsChecked = false;
        _strike.Checked += (_, _) => _dstrike.IsChecked = false;
        _dstrike.Checked += (_, _) => _strike.IsChecked = false;

        foreach (var c in new Control[] { _family, _style, _size, _underline, _color, _strike, _dstrike, _sup, _sub, _small, _caps, _hidden }) _start[c] = State(c);
        foreach (var c in _start.Keys)
            switch (c)
            {
                case ComboBox cb: cb.SelectionChanged += (_, _) => Dispatcher.BeginInvoke(Preview); cb.LostKeyboardFocus += (_, _) => Preview(); break;
                case CheckBox ch: ch.Click += (_, _) => Preview(); break;
            }
        _resolve = resolve;
        Preview();
    }

    readonly Func<Paragraph, Run, RunFormat>? _resolve;

    static string State(Control c) => c switch
    {
        ComboBox { IsEditable: true } cb => cb.Text,
        ComboBox cb => Tag(cb),
        CheckBox ch => ch.IsChecked == true ? "1" : "0",
        _ => "",
    };

    bool Changed(Control c) => State(c) != _start[c];

    void Preview()
    {
        try { _preview.FontFamily = new FontFamily(_family.Text.Length > 0 ? _family.Text : "Segoe UI"); } catch (ArgumentException) { }
        var st = Tag(_style);
        _preview.FontWeight = st.Contains('b') ? FontWeights.Bold : FontWeights.Normal;
        _preview.FontStyle = st.Contains('i') ? FontStyles.Italic : FontStyles.Normal;
        if (Parse(_size.Text) is { } pt && pt is >= 1 and <= 200) _preview.FontSize = Math.Min(40, pt * 96 / 72);
        var deco = new TextDecorationCollection();
        if (Tag(_underline) != "none") deco.Add(TextDecorations.Underline);
        if (_strike.IsChecked == true || _dstrike.IsChecked == true) deco.Add(TextDecorations.Strikethrough);
        _preview.TextDecorations = deco;
        _preview.Foreground = Tag(_color) is "auto" ? Brushes.Black : new SolidColorBrush((Color)ColorConverter.ConvertFromString("#" + Tag(_color)));
        _preview.Text = _caps.IsChecked == true || _small.IsChecked == true ? "APERÇU DU TEXTE" : "Aperçu du texte";
        _preview.Opacity = _hidden.IsChecked == true ? 0.4 : 1;
    }

    bool Build()
    {
        var ops = new List<Action<Paragraph, Run>>();
        if (Changed(_family) && _family.Text.Trim() is { Length: > 0 } family) ops.Add((_, r) => FormatOps.SetFont(r, family));
        if (Changed(_style))
        {
            var st = Tag(_style);
            ops.Add((p, r) => FormatOps.SetToggle(r, "b", st.Contains('b'), run => _resolve?.Invoke(p, run).Bold == true));
            ops.Add((p, r) => FormatOps.SetToggle(r, "i", st.Contains('i'), run => _resolve?.Invoke(p, run).Italic == true));
        }
        if (Changed(_size))
        {
            if (Parse(_size.Text) is not { } pt || pt < 1 || pt > 1638)
            {
                MessageBox.Show(this, "La taille doit être un nombre entre 1 et 1638.", "Police", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
            var half = Math.Round(pt * 2).ToString(CultureInfo.InvariantCulture);
            ops.Add((_, r) => FormatOps.SetValue(r, "sz", half));
        }
        if (Changed(_underline))
        {
            var u = Tag(_underline);
            ops.Add((_, r) => FormatOps.SetValue(r, "u", u));
        }
        if (Changed(_color))
        {
            var col = Tag(_color);
            ops.Add((_, r) => FormatOps.SetValue(r, "color", col));
        }
        void Toggle(CheckBox cb, string local, Func<RunFormat, bool> has)
        {
            if (!Changed(cb)) return;
            bool on = cb.IsChecked == true;
            ops.Add((p, r) => FormatOps.SetToggle(r, local, on, run => _resolve is not null && has(_resolve(p, run))));
        }
        Toggle(_strike, "strike", f => f.Strike);
        Toggle(_dstrike, "dstrike", f => f.DoubleStrike);
        Toggle(_small, "smallCaps", f => f.SmallCaps);
        Toggle(_caps, "caps", f => f.Caps);
        Toggle(_hidden, "vanish", f => f.Hidden);
        if (Changed(_sup) || Changed(_sub))
        {
            var va = _sup.IsChecked == true ? "superscript" : _sub.IsChecked == true ? "subscript" : null;
            ops.Add((p, r) =>
            {
                FormatOps.SetValue(r, "vertAlign", va);
                if (va is null && _resolve?.Invoke(p, r).VerticalAlign is not null) FormatOps.SetValue(r, "vertAlign", "baseline");
            });
        }
        if (ops.Count > 0) Result = (p, r) => { foreach (var op in ops) op(p, r); };
        return true;
    }

    /// <summary>Deux colonnes de cases à cocher.</summary>
    sealed class UniformGrid2
    {
        public readonly System.Windows.Controls.Primitives.UniformGrid Panel = new() { Columns = 2 };
        public void Add(UIElement e) => Panel.Children.Add(e);
    }
}
