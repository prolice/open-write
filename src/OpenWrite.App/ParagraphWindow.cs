// Repris d'OpenPub 2.1.3 (src/OpenPub.App/ParagraphWindow.cs), adapté aux propriétés de Word (w:ind, w:spacing, enchaînements).
using System.Windows;
using System.Windows.Controls;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;
using static OpenWrite.App.DialogKit;

namespace OpenWrite.App;

/// <summary>
/// Boîte « Paragraphe » de Word : alignement, retraits, espacement avant / après, interligne, enchaînements (veuves et
/// orphelines, paragraphes solidaires, lignes solidaires, saut de page avant). Initialisée depuis le paragraphe du point
/// d'insertion ; seuls les réglages changés sont appliqués aux paragraphes de la sélection.
/// </summary>
public sealed class ParagraphWindow : Window
{
    const double TwipsPerCm = 1440 / 2.54;

    readonly ComboBox _align = Combo(150, ("Gauche", "left"), ("Centré", "center"), ("Droite", "right"), ("Justifié", "both"));
    readonly TextBox _left = Box(), _right = Box(), _special = Box(), _before = Box(), _after = Box(), _lineValue = Box();
    readonly ComboBox _specialKind = Combo(140, ("(aucun)", "none"), ("Première ligne", "first"), ("Suspendu", "hanging"));
    readonly ComboBox _lineKind = Combo(140, ("Simple", "1"), ("1,5 ligne", "1.5"), ("Double", "2"), ("Au moins", "atLeast"), ("Exactement", "exact"), ("Multiple", "multi"));
    readonly CheckBox _contextual = new() { Content = "Ne pas ajouter d'espace entre les paragraphes du même style", Margin = new Thickness(0, 4, 0, 0) };
    readonly CheckBox _widow = new() { Content = "Éviter veuves et orphelines" }, _keepNext = new() { Content = "Paragraphes solidaires" },
        _keepLines = new() { Content = "Lignes solidaires" }, _pageBreak = new() { Content = "Saut de page avant" };
    readonly Dictionary<Control, string> _start = new();

    /// <summary>Modification à appliquer à chaque paragraphe (avec sa mise en forme résolue) ; null si rien n'a changé.</summary>
    public Action<Paragraph, ParagraphFormat>? Result { get; private set; }

    public ParagraphWindow(ParagraphFormat f)
    {
        Setup(this, "Paragraphe");
        var root = new StackPanel { Margin = new Thickness(18), Width = 470 };
        root.Children.Add(Section("Général"));
        root.Children.Add(Row("Alignement", _align));
        root.Children.Add(Section("Retrait (cm)"));
        root.Children.Add(Row("À gauche", _left));
        root.Children.Add(Row("À droite", _right));
        root.Children.Add(Row("De 1re ligne", Pair(_specialKind, _special)));
        root.Children.Add(Section("Espacement"));
        root.Children.Add(Row("Avant (pt)", _before));
        root.Children.Add(Row("Après (pt)", _after));
        root.Children.Add(Row("Interligne", Pair(_lineKind, _lineValue)));
        root.Children.Add(_contextual);
        root.Children.Add(Section("Enchaînements"));
        foreach (var cb in new[] { _widow, _keepNext, _keepLines, _pageBreak }) { cb.Margin = new Thickness(0, 3, 0, 3); root.Children.Add(cb); }
        root.Children.Add(Buttons(this, Build));
        Content = root;

        Select(_align, f.Alignment switch { "start" => "left", "end" => "right", "distribute" => "both", var a => a });
        _left.Text = Cm(f.IndentLeft);
        _right.Text = Cm(f.IndentRight);
        Select(_specialKind, f.IndentFirstLine == 0 ? "none" : f.IndentFirstLine > 0 ? "first" : "hanging");
        _special.Text = Cm(Math.Abs(f.IndentFirstLine));
        _before.Text = (f.SpacingBefore / 20.0).ToString("0.#", Cu);
        _after.Text = (f.SpacingAfter / 20.0).ToString("0.#", Cu);
        string kind = f.LineRule switch
        {
            "exact" => "exact",
            "atLeast" => "atLeast",
            _ => f.Line == 240 ? "1" : f.Line == 360 ? "1.5" : f.Line == 480 ? "2" : "multi",
        };
        Select(_lineKind, kind);
        _lineValue.Text = f.LineRule == "auto" ? (f.Line / 240.0).ToString("0.##", Cu) : (Math.Abs(f.Line) / 20.0).ToString("0.#", Cu);
        _contextual.IsChecked = f.ContextualSpacing;
        _widow.IsChecked = f.WidowControl;
        _keepNext.IsChecked = f.KeepNext;
        _keepLines.IsChecked = f.KeepLines;
        _pageBreak.IsChecked = f.PageBreakBefore;
        _special.IsEnabled = Tag(_specialKind) != "none";
        _specialKind.SelectionChanged += (_, _) =>
        {
            _special.IsEnabled = Tag(_specialKind) != "none";
            if (Tag(_specialKind) != "none" && Parse(_special.Text) is null or 0) _special.Text = "1,25";
        };
        _lineValue.IsEnabled = Tag(_lineKind) is "atLeast" or "exact" or "multi";
        _lineKind.SelectionChanged += (_, _) =>
        {
            var t = Tag(_lineKind);
            _lineValue.IsEnabled = t is "atLeast" or "exact" or "multi";
            _lineValue.Text = t switch { "atLeast" or "exact" => "12", "multi" => "3", _ => double.Parse(t, System.Globalization.CultureInfo.InvariantCulture).ToString("0.##", Cu) };
        };
        foreach (var c in new Control[] { _align, _left, _right, _specialKind, _special, _before, _after, _lineKind, _lineValue, _contextual, _widow, _keepNext, _keepLines, _pageBreak })
            _start[c] = State(c);
    }

    static StackPanel Pair(UIElement a, FrameworkElement b)
    {
        var sp = new StackPanel { Orientation = Orientation.Horizontal };
        sp.Children.Add(a);
        b.Margin = new Thickness(8, 0, 0, 0);
        sp.Children.Add(b);
        return sp;
    }

    static string Cm(int twips) => (twips / TwipsPerCm).ToString("0.##", Cu);

    static string State(Control c) => c switch
    {
        TextBox t => t.Text,
        ComboBox cb => Tag(cb),
        CheckBox ch => ch.IsChecked == true ? "1" : "0",
        _ => "",
    };

    bool Changed(params Control[] cs) => cs.Any(c => State(c) != _start[c]);

    bool Build()
    {
        var ops = new List<Action<Paragraph, ParagraphFormat>>();
        string? error = null;
        int? Twips(TextBox box, double factor, double min, double max)
        {
            if (Parse(box.Text) is { } v && v >= min && v <= max) return (int)Math.Round(v * factor);
            error ??= $"Valeur incorrecte : « {box.Text} ».";
            return null;
        }
        if (Changed(_align)) { var jc = Tag(_align); ops.Add((p, _) => p.Props.SetVal("jc", jc)); }
        if (Changed(_left)) { var v = Twips(_left, TwipsPerCm, -50, 50); ops.Add((p, _) => FormatOps.SetIndent(p, left: v)); }
        if (Changed(_right)) { var v = Twips(_right, TwipsPerCm, -50, 50); ops.Add((p, _) => FormatOps.SetIndent(p, right: v)); }
        if (Changed(_specialKind, _special))
        {
            var kind = Tag(_specialKind);
            int v = kind == "none" ? 0 : Twips(_special, TwipsPerCm, 0, 50) ?? 0;
            int first = kind == "hanging" ? -v : v;
            ops.Add((p, _) => FormatOps.SetIndent(p, firstLine: first));
        }
        if (Changed(_before)) { var v = Twips(_before, 20, 0, 1584); ops.Add((p, _) => FormatOps.SetSpacing(p, before: v)); }
        if (Changed(_after)) { var v = Twips(_after, 20, 0, 1584); ops.Add((p, _) => FormatOps.SetSpacing(p, after: v)); }
        if (Changed(_lineKind, _lineValue))
        {
            var kind = Tag(_lineKind);
            (int line, string rule) = kind switch
            {
                "atLeast" => (Twips(_lineValue, 20, 0.7, 1584) ?? 240, "atLeast"),
                "exact" => (Twips(_lineValue, 20, 0.7, 1584) ?? 240, "exact"),
                "multi" => (Twips(_lineValue, 240, 0.06, 132) ?? 240, "auto"),
                _ => ((int)Math.Round(double.Parse(kind, System.Globalization.CultureInfo.InvariantCulture) * 240), "auto"),
            };
            ops.Add((p, _) => FormatOps.SetSpacing(p, line: line, lineRule: rule));
        }
        void Flag(CheckBox cb, string local, Func<ParagraphFormat, bool> has)
        {
            if (!Changed(cb)) return;
            bool on = cb.IsChecked == true;
            ops.Add((p, f) => FormatOps.SetParagraphFlag(p, local, on, has(f)));
        }
        Flag(_contextual, "contextualSpacing", f => f.ContextualSpacing);
        Flag(_widow, "widowControl", f => f.WidowControl);
        Flag(_keepNext, "keepNext", f => f.KeepNext);
        Flag(_keepLines, "keepLines", f => f.KeepLines);
        Flag(_pageBreak, "pageBreakBefore", f => f.PageBreakBefore);
        if (error is not null)
        {
            MessageBox.Show(this, error, "Paragraphe", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }
        if (ops.Count > 0) Result = (p, f) => { foreach (var op in ops) op(p, f); };
        return true;
    }
}
