// Repris d'OpenPub 2.1.4 (OpenPub.App/Output/PrintWindow.cs) : pages d'OpenWrite (OutputPages).
using System.Printing;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace OpenWrite.App.Output;

/// <summary>
/// Boite « Imprimer » avec aperçu : pages, exemplaires, disposition (une page par feuille, livret, recto verso manuel).
/// L'aperçu montre les feuilles telles qu'elles sortiront, dans l'ordre d'impression.
/// </summary>
public sealed class PrintWindow : Window
{
    readonly OutputPages _output;
    readonly PageChooser _pages;
    readonly TextBox _copies = OutputUi.Box("1", 50);
    readonly ComboBox _mode = OutputUi.Combo(
        ("Une page par feuille", PrintMode.Normal),
        ("Recto verso manuel", PrintMode.ManualDuplex),
        ("Livret (imprimante recto verso)", PrintMode.Booklet),
        ("Livret, recto verso manuel", PrintMode.BookletManualDuplex));
    readonly CheckBox _reverse = new() { Content = "Versos dans l'ordre inverse", Margin = new Thickness(0, 6, 0, 0) };
    readonly TextBlock _summary = OutputUi.Note("");
    // apercu : les feuilles rendues en vignettes, dans l'ordre d'impression, avec leur legende
    readonly WrapPanel _sheets = new() { Margin = new Thickness(4) };
    readonly ScrollViewer _preview;
    const double PreviewHeight = 300;
    readonly DispatcherTimer _refresh = new() { Interval = TimeSpan.FromMilliseconds(250) };
    readonly Size? _paper = DefaultPaper();

    public PrintOptions? Result { get; private set; }

    public PrintWindow(OutputPages pages, int activePage)
    {
        _output = pages;
        OutputUi.Style(this, "Imprimer", Help.HelpService.Print);
        Width = 1120; Height = 760;
        MinWidth = 800; MinHeight = 520;
        _pages = new PageChooser(pages.Count, activePage);
        _preview = new ScrollViewer
        {
            Content = _sheets, VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            Background = OutputUi.Res<System.Windows.Media.Brush>("WorkspaceBrush"),
        };

        var left = new StackPanel { Margin = new Thickness(18), Width = 300 };
        left.Children.Add(OutputUi.Section("Pages"));
        left.Children.Add(_pages);
        left.Children.Add(OutputUi.Section("Exemplaires"));
        left.Children.Add(OutputUi.Row("Nombre", _copies, 80));
        left.Children.Add(OutputUi.Note("Les exemplaires sont assemblés (le document entier, puis le suivant)."));
        left.Children.Add(OutputUi.Section("Disposition"));
        _mode.Width = 290;
        left.Children.Add(_mode);
        left.Children.Add(_reverse);
        left.Children.Add(_summary);
        left.Children.Add(OutputUi.Buttons(this, "Imprimer…", Build));

        var grid = new Grid();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.ColumnDefinitions.Add(new ColumnDefinition());
        var scroll = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Content = left };
        grid.Children.Add(scroll);
        var previewBox = new DockPanel { Margin = new Thickness(0, 12, 12, 12) };
        var title = OutputUi.Section("Aperçu avant impression");
        title.Margin = new Thickness(0, 6, 0, 6);
        DockPanel.SetDock(title, Dock.Top);
        previewBox.Children.Add(title);
        previewBox.Children.Add(_preview);
        Grid.SetColumn(previewBox, 1);
        grid.Children.Add(previewBox);
        Content = grid;

        _refresh.Tick += (_, _) => { _refresh.Stop(); UpdatePreview(); };
        void Later() { _refresh.Stop(); _refresh.Start(); }
        _pages.Changed += Later;
        _copies.TextChanged += (_, _) => Later();
        _mode.SelectionChanged += (_, _) => Later();
        _reverse.Checked += (_, _) => Later(); _reverse.Unchecked += (_, _) => Later();
        Loaded += (_, _) => UpdatePreview();
    }

    /// <summary>Papier de l'imprimante par defaut (1/96 de pouce), pour l'apercu du livret ; null s'il n'y en a pas.</summary>
    static Size? DefaultPaper()
    {
        try
        {
            using var server = new LocalPrintServer();
            using var q = server.DefaultPrintQueue;
            var t = q?.UserPrintTicket ?? q?.DefaultPrintTicket;
            if (t?.PageMediaSize is { Width: { } w, Height: { } h }) return new Size(w, h);
        }
        catch (Exception) { }
        return null;
    }

    /// <summary>Reglages saisis (null et message si la saisie est invalide).</summary>
    PrintOptions? Options(out string? error)
    {
        error = null;
        var pages = _pages.Pages(out error);
        if (pages is null) return null;
        if (!int.TryParse(_copies.Text.Trim(), out var copies) || copies < 1 || copies > 99) { error = "Le nombre d'exemplaires doit être compris entre 1 et 99."; return null; }
        var mode = (PrintMode)OutputUi.Tag(_mode)!;
        return new PrintOptions { Pages = pages, Copies = copies, Mode = mode, ReverseBacks = _reverse.IsChecked == true };
    }

    void UpdatePreview()
    {
        var mode = (PrintMode)OutputUi.Tag(_mode)!;
        _reverse.IsEnabled = mode is PrintMode.ManualDuplex or PrintMode.BookletManualDuplex;
        var o = Options(out var err);
        if (o is null) { _summary.Text = err; _sheets.Children.Clear(); PreviewDocument = null; return; }
        var passes = PrintService.Passes(o);
        bool booklet = mode is PrintMode.Booklet or PrintMode.BookletManualDuplex;
        string Faces(IEnumerable<PrintFace> f) => string.Join(", ", f.Take(16)) + (f.Count() > 16 ? "…" : "");
        var lines = new List<string>();
        if (booklet)
        {
            int sheets = (passes.Sum(p => p.Count) + 1) / 2;
            lines.Add($"{sheets} feuille(s) pliée(s) en deux, 2 pages par face ; pages blanches ajoutées jusqu'à un multiple de 4.");
            if (mode == PrintMode.Booklet) lines.Add("Recto verso automatique, retournement sur le bord court.");
        }
        if (passes.Count == 2)
        {
            lines.Add($"1er passage (rectos) : {Faces(passes[0])}");
            lines.Add($"2e passage (versos, après avoir retourné la pile) : {Faces(passes[1])}");
        }
        else lines.Add($"Ordre : {Faces(passes[0])}");
        _summary.Text = string.Join("\n", lines);
        // legendes : feuille et face (livret), passage (recto verso manuel), page
        var labels = new List<string>();
        for (int pi = 0; pi < passes.Count; pi++)
            for (int k = 0; k < passes[pi].Count; k++)
            {
                var f = passes[pi][k];
                string what = !f.TwoUp ? (f.Left == 0 ? "page blanche" : $"page {f.Left}")
                    : f.Left > 0 && f.Right > 0 ? $"pages {f.Left} et {f.Right}"
                    : f.Left == 0 && f.Right == 0 ? "feuille blanche"
                    : $"page {(f.Left == 0 ? "blanche" : f.Left.ToString())} et page {(f.Right == 0 ? "blanche" : f.Right.ToString())}";
                labels.Add(passes.Count == 2 ? $"{(pi == 0 ? "Recto" : "Verso")} {k + 1} — {what}"
                    : booklet ? $"Feuille {k / 2 + 1}, {(k % 2 == 0 ? "recto" : "verso")} — {what}"
                    : what[..1].ToUpperInvariant() + what[1..]);
            }
        var all = passes.SelectMany(p => p).ToList();
        var doc = PrintService.BuildDocument(_output, all, mode, _paper);
        PreviewDocument = doc;
        _sheets.Children.Clear();
        for (int i = 0; i < doc.Pages.Count; i++)
        {
            var fp = doc.Pages[i].Child;
            var size = new Size(fp.Width, fp.Height);
            fp.Measure(size);
            fp.Arrange(new Rect(size));
            fp.UpdateLayout();
            double scale = PreviewHeight / fp.Height;
            var rtb = new System.Windows.Media.Imaging.RenderTargetBitmap((int)Math.Ceiling(fp.Width * scale), (int)Math.Ceiling(fp.Height * scale),
                96 * scale, 96 * scale, System.Windows.Media.PixelFormats.Pbgra32);
            rtb.Render(fp);
            rtb.Freeze();
            var item = new StackPanel { Margin = new Thickness(10) };
            item.Children.Add(new Border
            {
                Child = new Image { Source = rtb, Width = rtb.PixelWidth, Height = rtb.PixelHeight },
                BorderBrush = OutputUi.Res<System.Windows.Media.Brush>("Line"), BorderThickness = new Thickness(1),
                Effect = new System.Windows.Media.Effects.DropShadowEffect { BlurRadius = 8, ShadowDepth = 2, Opacity = 0.18 },
            });
            item.Children.Add(new TextBlock { Text = i < labels.Count ? labels[i] : "", Margin = new Thickness(0, 6, 0, 0), FontSize = 12 });
            _sheets.Children.Add(item);
        }
    }

    bool Build()
    {
        var o = Options(out var err);
        if (o is null) { MessageBox.Show(this, err, Title, MessageBoxButton.OK, MessageBoxImage.Warning); return false; }
        Result = o;
        return true;
    }

    /// <summary>Apercu construit (auto-test).</summary>
    public System.Windows.Documents.FixedDocument? PreviewDocument { get; private set; }

    /// <summary>Vignettes affichees (auto-test).</summary>
    public int PreviewSheets => _sheets.Children.Count;

    /// <summary>Choisit une disposition et met l'apercu a jour (auto-test).</summary>
    public void SetMode(PrintMode mode)
    {
        foreach (ComboBoxItem it in _mode.Items) if ((PrintMode)it.Tag == mode) _mode.SelectedItem = it;
        UpdatePreview();
    }
}
