// Repris d'OpenPub 2.1.4 (OpenPub.App/Output/PdfExportWindow.cs), sans les réglages pour l'imprimeur.
using System.Windows;
using System.Windows.Controls;

namespace OpenWrite.App.Output;

/// <summary>Boîte « Exporter en PDF » : pages à exporter, ouverture du PDF une fois créé.</summary>
public sealed class PdfExportWindow : Window
{
    readonly PageChooser _pages;
    readonly CheckBox _open = new() { Content = "Ouvrir le PDF une fois créé", IsChecked = true, Margin = new Thickness(0, 12, 0, 0) };

    public PdfExportOptions? Result { get; private set; }
    public bool OpenAfter => _open.IsChecked == true;

    public PdfExportWindow(int pageCount, int activePage)
    {
        OutputUi.Style(this, "Exporter en PDF", Help.HelpService.Pdf);
        SizeToContent = SizeToContent.WidthAndHeight;
        ResizeMode = ResizeMode.NoResize;
        _pages = new PageChooser(pageCount, activePage);

        var root = new StackPanel { Margin = new Thickness(18), Width = 400 };
        root.Children.Add(OutputUi.Section("Pages"));
        root.Children.Add(_pages);
        root.Children.Add(OutputUi.Note("Les polices sont incorporées et les liens hypertextes restent cliquables. " +
                                        "La couleur de page n'est pas exportée, comme dans Word."));
        root.Children.Add(_open);
        root.Children.Add(OutputUi.Buttons(this, "Exporter…", Build));
        Content = root;
    }

    /// <summary>Pages choisies (auto-test).</summary>
    internal PageChooser Chooser => _pages;

    bool Build()
    {
        var pages = _pages.Pages(out var err);
        if (pages is null) { MessageBox.Show(this, err, Title, MessageBoxButton.OK, MessageBoxImage.Warning); return false; }
        Result = new PdfExportOptions { Pages = pages };
        return true;
    }
}
