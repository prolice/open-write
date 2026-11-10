using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;
using StyleOps = OpenWrite.Core.Styles.StyleOps;

namespace OpenWrite.App;

// Lot 3 (version 0.4) : volet « Mise en forme » actif — police, taille, gras / italique / souligné / barré, alignements,
// interligne, boîtes Police et Paragraphe ; il montre la mise en forme du point d'insertion. Agrandir / réduire la
// police, effacer la mise en forme (Ctrl+Espace, Ctrl+Q).
public partial class MainWindow
{
    /// <summary>Tailles proposées et paliers de Ctrl+Maj+&gt; / Ctrl+&lt; (ceux de Word).</summary>
    static readonly double[] FontSizes = { 8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 };
    static readonly double[] LineSpacings = { 1, 1.15, 1.5, 2, 2.5, 3 };
    static readonly CultureInfo Cu = CultureInfo.CurrentCulture;

    /// <summary>Mise à jour du volet en cours : ne pas traiter ses événements comme des choix de l'utilisateur.</summary>
    bool _refreshingPane;

    void InitFormatPane()
    {
        InitStylePane();
        PaneFont.ItemsSource = System.Windows.Media.Fonts.SystemFontFamilies.Select(f => f.Source).Distinct().OrderBy(n => n, StringComparer.CurrentCultureIgnoreCase).ToList();
        PaneSize.ItemsSource = FontSizes.Select(s => s.ToString(Cu)).ToList();
        PaneLineSpacing.ItemsSource = LineSpacings.Select(s => s.ToString("0.0#", Cu)).ToList();

        PaneBold.Click += (_, _) => { ToggleCharacter("b"); BackToPage(); };
        PaneItalic.Click += (_, _) => { ToggleCharacter("i"); BackToPage(); };
        PaneUnderline.Click += (_, _) => { ToggleCharacter("u"); BackToPage(); };
        PaneStrike.Click += (_, _) => { ToggleCharacter("strike"); BackToPage(); };
        PaneLeft.Click += (_, _) => { SetAlignment("left"); BackToPage(); };
        PaneCenter.Click += (_, _) => { SetAlignment("center"); BackToPage(); };
        PaneRight.Click += (_, _) => { SetAlignment("right"); BackToPage(); };
        PaneJustify.Click += (_, _) => { SetAlignment("both"); BackToPage(); };
        PaneFontDialog.Click += (_, _) => ShowFontDialog();
        PaneParagraphDialog.Click += (_, _) => ShowParagraphDialog();

        // police, taille, interligne : choix dans la liste ou saisie validée par Entrée
        Commit(PaneFont, text => SetFontFamily(text));
        Commit(PaneSize, text =>
        {
            if (double.TryParse(text.Replace('.', ','), NumberStyles.Float, CultureInfo.GetCultureInfo("fr-BE"), out var pt) && pt >= 1 && pt <= 1638) SetFontSize(pt);
            else RefreshFormatPane();
        });
        Commit(PaneLineSpacing, text =>
        {
            if (double.TryParse(text.Replace('.', ','), NumberStyles.Float, CultureInfo.GetCultureInfo("fr-BE"), out var x) && x > 0 && x <= 10) SetLineSpacing(x);
            else RefreshFormatPane();
        });
    }

    void Commit(ComboBox box, Action<string> apply)
    {
        box.SelectionChanged += (_, _) =>
        {
            if (_refreshingPane || box.SelectedItem is not string s || !box.IsDropDownOpen && !box.IsKeyboardFocusWithin) return;
            apply(s);
        };
        box.DropDownClosed += (_, _) => { if (!_refreshingPane && box.SelectedItem is string) BackToPage(); };
        box.PreviewKeyDown += (_, e) =>
        {
            if (e.Key == Key.Enter) { apply(box.Text); BackToPage(); e.Handled = true; }
            else if (e.Key == Key.Escape) { RefreshFormatPane(); BackToPage(); e.Handled = true; }
        };
    }

    void BackToPage() => _pagesView?.Focus();

    /// <summary>Dernier état affiché par le volet (rien à refaire s'il n'a pas changé).</summary>
    object? _paneState;

    /// <summary>Volet : mise en forme du point d'insertion (police, taille, attributs, alignement, interligne).</summary>
    void RefreshFormatPane()
    {
        bool enabled = _layout is not null && _caret is not null;
        if (PaneControls.IsEnabled != enabled) PaneControls.IsEnabled = enabled;
        if (_caret is not { } c || LayoutOf(c.Paragraph) is not { } pl || CaretFormat() is not { } rf) return;
        // valeurs affichées seulement (les formats résolus sont recréés à chaque remise en page du paragraphe)
        var state = (rf.FontAscii, rf.FontHAnsi, rf.Size, rf.Bold, rf.Italic, rf.Underline, rf.Strike, pl.Format.Alignment, pl.Format.Line, pl.Format.LineRule, pl.Format.StyleId);
        if (Equals(state, _paneState) && !PaneFont.IsKeyboardFocusWithin) return;
        _paneState = state;
        _refreshingPane = true;
        try
        {
            var pf = pl.Format;
            PaneFont.SelectedItem = null;
            PaneFont.Text = rf.FontAscii ?? rf.FontHAnsi ?? "";
            PaneSize.SelectedItem = null;
            PaneSize.Text = rf.Size.ToString("0.#", Cu);
            PaneBold.IsChecked = rf.Bold;
            PaneItalic.IsChecked = rf.Italic;
            PaneUnderline.IsChecked = Has(rf, "u");
            PaneStrike.IsChecked = rf.Strike;
            PaneLeft.IsChecked = pf.Alignment is "left" or "start";
            PaneCenter.IsChecked = pf.Alignment == "center";
            PaneRight.IsChecked = pf.Alignment is "right" or "end";
            PaneJustify.IsChecked = pf.Alignment is "both" or "distribute";
            PaneLineSpacing.SelectedItem = null;
            PaneLineSpacing.Text = pf.LineRule == "auto" ? (pf.Line / 240.0).ToString("0.0#", Cu) : $"{pf.Line / 20.0:0.#} pt";
            var styleName = _styles is { } st ? StyleOps.DisplayName(st.Sheet, pf.StyleId ?? st.Sheet.DefaultStyle("paragraph")?.Attribute(st.Sheet.W + "styleId")?.Value) : "Normal";
            if (!PaneStyle.IsDropDownOpen) { PaneStyle.ItemsSource = new[] { styleName }; PaneStyle.SelectedIndex = 0; }
        }
        finally { _refreshingPane = false; }
    }

    // ================================================================ commandes de caractère

    /// <summary>
    /// Applique une mise en forme de caractère au texte visé (sélection, sinon mot du point d'insertion), ou la met en
    /// attente pour la prochaine frappe entre deux mots.
    /// </summary>
    void ApplyToTarget(string label, Action<Paragraph, Run> apply)
    {
        if (CharacterTarget() is { } target) { ApplyCharacter(label, target.From, target.To, apply); return; }
        _pendingFormat.Add(apply);
        StatusText.Text = $"{label} : pour le texte que vous allez taper.";
        RefreshFormatPane();
    }

    internal void SetFontFamily(string family)
    {
        family = family.Trim();
        if (family.Length == 0) { RefreshFormatPane(); return; }
        ApplyToTarget($"Police {family}", (_, r) => FormatOps.SetFont(r, family));
    }

    internal void SetFontSize(double pt)
    {
        var half = Math.Round(pt * 2).ToString(CultureInfo.InvariantCulture);
        ApplyToTarget($"Taille {pt.ToString("0.#", Cu)} pt", (_, r) => FormatOps.SetValue(r, "sz", half));
    }

    /// <summary>Ctrl+Maj+&gt; / Ctrl+&lt; : palier de taille suivant ou précédent, pour chaque segment selon sa taille.</summary>
    void GrowFont(int direction)
    {
        if (_styles is not { } styles) return;
        ApplyToTarget(direction > 0 ? "Agrandir la police" : "Réduire la police", (p, r) =>
        {
            double size = styles.Run(p, r).Size;
            double next = direction > 0
                ? FontSizes.FirstOrDefault(s => s > size + 0.01, Math.Min(1638, size + 10))
                : FontSizes.LastOrDefault(s => s < size - 0.01, Math.Max(1, size - 1));
            FormatOps.SetValue(r, "sz", Math.Round(next * 2).ToString(CultureInfo.InvariantCulture));
        });
    }

    /// <summary>Ctrl+Espace : mise en forme directe des caractères retirée (celle du style revient).</summary>
    void ClearCharacterFormatting() => ApplyToTarget("Effacer la mise en forme des caractères", (_, r) => FormatOps.ClearCharacter(r));

    // ================================================================ commandes de paragraphe

    /// <summary>Applique une modification aux paragraphes de la sélection (ou à celui du point d'insertion), annulable.</summary>
    internal void ApplyParagraphs(string label, Action<Paragraph> apply)
    {
        if (Body is not { } body || Ordered() is not { } sel) return;
        if (ToDocPos(sel.From) is not { } from || ToDocPos(sel.To) is not { } to) { NotEditable(); return; }
        if (!SupportedRange(from, to)) return;
        var (anchor, caret) = SelectionDocPos();
        FormatOps.Apply(_journal, body, label, from, to, () => FormatOps.ApplyToParagraphs(body, from, to, apply));
        AfterFormat(label, anchor, caret);
    }

    void SetLineSpacing(double multiple) =>
        ApplyParagraphs($"Interligne {multiple.ToString("0.0#", Cu)}", p => FormatOps.SetSpacing(p, line: (int)Math.Round(multiple * 240), lineRule: "auto"));

    /// <summary>Ctrl+Q : mise en forme directe des paragraphes retirée.</summary>
    void ClearParagraphFormatting() => ApplyParagraphs("Effacer la mise en forme du paragraphe", FormatOps.ClearParagraph);

    // ================================================================ marques de mise en forme

    bool _showMarks;

    /// <summary>Ctrl+Maj+8 (ou Ctrl+*) : affiche ou masque ¶, espaces, tabulations et sauts, comme le bouton ¶ de Word.</summary>
    internal void ToggleMarks()
    {
        _showMarks = !_showMarks;
        if (_drawings is not null)
        {
            _drawings.Renderer.ShowMarks = _showMarks;
            _drawings.Clear();
            _pagesView?.InvalidateVisual();
            _thumbsView?.InvalidateVisual();
        }
        foreach (var b in _liveButtons.OfType<System.Windows.Controls.Button>().Where(b => Equals(b.Content, "Marques ¶")))
            b.Background = _showMarks ? R("Selected") : System.Windows.Media.Brushes.Transparent;
        StatusText.Text = _showMarks ? "Marques de mise en forme affichées (Ctrl+Maj+8 pour les masquer)." : "Marques de mise en forme masquées.";
    }

    // ================================================================ boîtes de dialogue

    internal void ShowFontDialog()
    {
        if (_styles is null || CaretFormat() is not { } rf) return;
        var target = CharacterTarget();
        // la boîte part du premier caractère visé
        if (target is { } t && FormatsIn(t.From, t.To).FirstOrDefault() is { } first) rf = first;
        var styles = _styles;
        var dlg = new FontWindow(rf, (p, r) => styles.Run(p, r)) { Owner = this };
        if (dlg.ShowDialog() == true && dlg.Result is { } apply) ApplyToTarget("Police", apply);
        BackToPage();
    }

    internal void ShowParagraphDialog()
    {
        if (_caret is not { } c || LayoutOf(c.Paragraph) is not { } pl) return;
        var dlg = new ParagraphWindow(pl.Format) { Owner = this };
        if (dlg.ShowDialog() == true && dlg.Result is { } apply) ApplyParagraphs("Paragraphe", p => apply(p, _styles!.Paragraph(p)));
        BackToPage();
    }
}
