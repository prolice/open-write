using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using OpenWrite.App.Output;

namespace OpenWrite.App;

// Version 1.0.2 : impression (aperçu, exemplaires, recto verso manuel, livret) et export PDF direct, repris d'OpenPub
// (PDFsharp ; PLAN.md §8.13 pour le PDF complet : signets, liens internes, PDF balisé).
public partial class MainWindow
{
    /// <summary>Octets d'origine des images décodées (JPEG / PNG incorporés tels quels dans le PDF).</summary>
    readonly ConditionalWeakTable<BitmapSource, byte[]> _imageBytes = new();

    void InitOutput()
    {
        CommandBindings.Add(new CommandBinding(ApplicationCommands.Print, (_, _) => ShowPrintWindow(), (_, e) => e.CanExecute = _layout is { Pages.Count: > 0 }));
    }

    string DocumentTitle => _info?.Title ?? Path.GetFileNameWithoutExtension(_path ?? "OpenWrite");

    /// <summary>Pages du document pour l'impression et le PDF (null sans document).</summary>
    internal OutputPages? Output() =>
        _layout is { Pages.Count: > 0 } layout ? new OutputPages(layout, ImageOf, l => LinkTarget(l) is { } u && !u.StartsWith('#') ? u : null, RevisionBarBrush) : null;

    /// <summary>Page du point d'insertion (Word : « page active »), à défaut la page affichée ; à partir de 1.</summary>
    int ActivePageNumber
    {
        get
        {
            if (_caret is { } c && CaretRect(c) is { } r) return r.Page + 1;
            if (_pagesView is null) return 1;
            try
            {
                var p = Scroller.TransformToDescendant(_pagesView).Transform(new Point(0, Scroller.ViewportHeight / 3));
                return _pagesView.PageAt(p.Y) + 1;
            }
            catch (InvalidOperationException) { return 1; }
        }
    }

    /// <summary>Dernière boîte de sortie ouverte (auto-test).</summary>
    internal Window? LastOutputWindow { get; private set; }

    internal void ShowPrintWindow()
    {
        if (Output() is not { } pages) return;
        var win = new PrintWindow(pages, ActivePageNumber) { Owner = this };
        LastOutputWindow = win;
        if (win.ShowDialog() != true || win.Result is not { } o) return;
        bool done = PrintService.Print(pages, o, DocumentTitle, () =>
            MessageBox.Show(this,
                "Les rectos sont imprimés.\n\nRetournez la pile de feuilles imprimées et remettez-la dans le bac d'alimentation " +
                "(face imprimée selon votre imprimante, haut de la page vers l'intérieur), puis cliquez sur OK pour imprimer les versos.",
                "Recto verso manuel", MessageBoxButton.OKCancel, MessageBoxImage.Information) == MessageBoxResult.OK);
        if (done) StatusText.Text = "Impression envoyée à l'imprimante.";
    }

    internal void ExportPdf()
    {
        if (Output() is not { } pages) return;
        var win = new PdfExportWindow(pages.Count, ActivePageNumber) { Owner = this };
        LastOutputWindow = win;
        if (win.ShowDialog() != true || win.Result is not { } o) return;
        var dlg = new SaveFileDialog
        {
            Title = "Exporter en PDF", Filter = "Document PDF (*.pdf)|*.pdf", DefaultExt = ".pdf",
            FileName = Path.GetFileNameWithoutExtension(_path ?? "Document") + ".pdf",
            InitialDirectory = Path.GetDirectoryName(_path ?? "") is { Length: > 0 } d ? d : null,
        };
        if (dlg.ShowDialog(this) != true) return;
        try
        {
            Mouse.OverrideCursor = Cursors.Wait;
            StatusText.Text = "Création du PDF…";
            var sw = Stopwatch.StartNew();
            ExportPdfTo(dlg.FileName, o);
            StatusText.Text = $"PDF créé ({o.Pages?.Count ?? pages.Count} page(s), {new FileInfo(dlg.FileName).Length / 1048576.0:0.0} Mo, {sw.Elapsed.TotalSeconds:0.0} s) : {dlg.FileName}";
        }
        catch (Exception ex)
        {
            Mouse.OverrideCursor = null;
            MessageBox.Show(this, "Le PDF n'a pas pu être créé :\n" + ex.Message, "Exporter en PDF", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }
        finally { Mouse.OverrideCursor = null; }
        if (win.OpenAfter)
            try { Process.Start(new ProcessStartInfo(dlg.FileName) { UseShellExecute = true }); }
            catch (Exception) { /* pas de lecteur PDF : le fichier est créé quand même */ }
    }

    /// <summary>Écrit le PDF du document (toutes les pages par défaut) ; ligne de commande --pdf et auto-test.</summary>
    internal PdfStats ExportPdfTo(string path, PdfExportOptions? o = null)
    {
        var pages = Output() ?? throw new InvalidOperationException("Aucun document ouvert.");
        o ??= new PdfExportOptions();
        o = o with { Title = o.Title ?? DocumentTitle, Author = o.Author ?? _info?.Author };
        return PdfExporter.Export(pages, path, o, bmp => _imageBytes.TryGetValue(bmp, out var b) ? b : null);
    }
}
