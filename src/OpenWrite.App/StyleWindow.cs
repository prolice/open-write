using System.Windows;
using System.Windows.Controls;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;
using static OpenWrite.App.DialogKit;

namespace OpenWrite.App;

/// <summary>
/// Boîte « Modifier le style » / « Nouveau style » (lot 4) : nom, style de base, style suivant, et les boîtes Police et
/// Paragraphe pour la mise en forme du style. Le résultat s'applique aux propriétés du style (styles.xml) : tout le
/// document qui l'utilise change.
/// </summary>
public sealed class StyleWindow : Window
{
    readonly TextBox _name = Box(260);
    readonly ComboBox _basedOn = new() { Width = 260 }, _next = new() { Width = 260 };
    readonly TextBlock _summary = new() { TextWrapping = TextWrapping.Wrap, MaxWidth = 400, Margin = new Thickness(0, 8, 0, 0) };
    readonly RunFormat _font;
    readonly ParagraphFormat _para;
    readonly List<Action<Paragraph, Run>> _fontOps = new();
    readonly List<Action<Paragraph, ParagraphFormat>> _paraOps = new();

    public string StyleName => _name.Text.Trim();
    public string? BasedOn => _basedOn.SelectedItem is ComboBoxItem { Tag: string t } && t.Length > 0 ? t : null;
    public string? Next => _next.SelectedItem is ComboBoxItem { Tag: string t } && t.Length > 0 ? t : null;
    public bool BasedOnChanged { get; private set; }
    public bool NextChanged { get; private set; }

    /// <summary>Modifications de la mise en forme du style (sur ses w:pPr / w:rPr) ; null si inchangée.</summary>
    public Action<Paragraph, Run>? Apply { get; private set; }

    /// <param name="choices">Styles de paragraphe du document (identifiant, nom affiché).</param>
    public StyleWindow(string title, string name, bool nameEditable, IReadOnlyList<(string Id, string Display)> choices, string? basedOn, string? next,
        RunFormat font, ParagraphFormat para)
    {
        Setup(this, title);
        _font = font;
        _para = para;
        var root = new StackPanel { Margin = new Thickness(18), Width = 430 };
        root.Children.Add(Section("Propriétés"));
        root.Children.Add(Row("Nom", _name, 140));
        root.Children.Add(Row("Basé sur", _basedOn, 140));
        root.Children.Add(Row("Style suivant", _next, 140));
        _basedOn.Items.Add(new ComboBoxItem { Content = "(aucun)", Tag = "" });
        _next.Items.Add(new ComboBoxItem { Content = "(le même)", Tag = "" });
        foreach (var (id, display) in choices)
        {
            _basedOn.Items.Add(new ComboBoxItem { Content = display, Tag = id });
            _next.Items.Add(new ComboBoxItem { Content = display, Tag = id });
        }
        Select(_basedOn, basedOn ?? "");
        Select(_next, next ?? "");
        _name.Text = name;
        _name.IsEnabled = nameEditable;
        _basedOn.SelectionChanged += (_, _) => BasedOnChanged = true;
        _next.SelectionChanged += (_, _) => NextChanged = true;

        root.Children.Add(Section("Mise en forme"));
        var buttons = new StackPanel { Orientation = Orientation.Horizontal };
        var fontButton = new Button { Content = "Police…", MinWidth = 110, Margin = new Thickness(0, 0, 8, 0) };
        var paraButton = new Button { Content = "Paragraphe…", MinWidth = 110 };
        fontButton.Click += (_, _) =>
        {
            var dlg = new FontWindow(_font) { Owner = this };
            if (dlg.ShowDialog() == true && dlg.Result is { } r) { _fontOps.Add(r); Summary(); }
        };
        paraButton.Click += (_, _) =>
        {
            var dlg = new ParagraphWindow(_para) { Owner = this };
            if (dlg.ShowDialog() == true && dlg.Result is { } r) { _paraOps.Add(r); Summary(); }
        };
        buttons.Children.Add(fontButton);
        buttons.Children.Add(paraButton);
        root.Children.Add(buttons);
        _summary.Foreground = (System.Windows.Media.Brush)Application.Current.FindResource("Muted");
        root.Children.Add(_summary);
        Summary();
        root.Children.Add(Buttons(this, () =>
        {
            if (StyleName.Length == 0) { MessageBox.Show(this, "Donnez un nom au style.", Title, MessageBoxButton.OK, MessageBoxImage.Information); return false; }
            if (_fontOps.Count + _paraOps.Count > 0)
            {
                var fontOps = _fontOps.ToList();
                var paraOps = _paraOps.ToList();
                Apply = (p, r) =>
                {
                    foreach (var op in fontOps) op(p, r);
                    foreach (var op in paraOps) op(p, _para);
                };
            }
            return true;
        }));
        Content = root;
    }

    void Summary()
    {
        var cu = Cu;
        _summary.Text = $"Police : {_font.FontAscii ?? "?"} {_font.Size.ToString("0.#", cu)} pt" + (_font.Bold ? ", gras" : "") + (_font.Italic ? ", italique" : "") +
                        $"\nAlignement : {_para.Alignment}, espacement avant {_para.SpacingBefore / 20.0:0.#} pt, après {_para.SpacingAfter / 20.0:0.#} pt" +
                        (_fontOps.Count + _paraOps.Count > 0 ? $"\n{_fontOps.Count + _paraOps.Count} modification(s) en attente : appliquées par OK." : "");
    }
}

/// <summary>Boîte « Styles » : tous les styles du document, appliquer, modifier, nouveau style à partir de la sélection.</summary>
public sealed class StylesWindow : Window
{
    readonly ListBox _list = new() { Height = 320, Width = 320 };
    public string? Chosen { get; private set; }
    /// <summary>apply, modify, new, update.</summary>
    public string Action { get; private set; } = "";

    public StylesWindow(IReadOnlyList<(string Key, string Display, string Type)> styles, string? current)
    {
        Setup(this, "Styles");
        var root = new StackPanel { Margin = new Thickness(16) };
        foreach (var (key, display, type) in styles)
        {
            var item = new ListBoxItem { Content = (type == "character" ? "a  " : "¶  ") + display, Tag = key, ToolTip = type == "character" ? "Style de caractère" : "Style de paragraphe" };
            _list.Items.Add(item);
            if (key == current) _list.SelectedItem = item;
        }
        _list.MouseDoubleClick += (_, _) => Done("apply");
        var panel = new StackPanel { Orientation = Orientation.Horizontal };
        panel.Children.Add(_list);
        var actions = new StackPanel { Margin = new Thickness(12, 0, 0, 0) };
        Button B(string text, string action, bool primary = false)
        {
            var b = new Button { Content = text, MinWidth = 210, Margin = new Thickness(0, 0, 0, 6), HorizontalContentAlignment = HorizontalAlignment.Left, Padding = new Thickness(10, 4, 10, 4) };
            if (primary) b.Style = (Style)Application.Current.FindResource("PillButton");
            b.Click += (_, _) => Done(action);
            return b;
        }
        actions.Children.Add(B("Appliquer", "apply", primary: true));
        actions.Children.Add(B("Modifier…", "modify"));
        actions.Children.Add(B("Mettre à jour selon la sélection", "update"));
        actions.Children.Add(B("Nouveau style à partir de la sélection…", "new"));
        var close = new Button { Content = "Fermer", IsCancel = true, MinWidth = 210, Margin = new Thickness(0, 12, 0, 0) };
        actions.Children.Add(close);
        panel.Children.Add(actions);
        root.Children.Add(panel);
        Content = root;
    }

    void Done(string action)
    {
        if (action != "new" && _list.SelectedItem is not ListBoxItem { Tag: string key }) return;
        Chosen = (_list.SelectedItem as ListBoxItem)?.Tag as string;
        Action = action;
        DialogResult = true;
    }
}
