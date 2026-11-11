using System.IO;
using System.Windows;
using System.Windows.Automation.Peers;
using System.Windows.Automation.Provider;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;

namespace OpenWrite.App;

// Auto-test des sorties (version 1.0.2) : Fichier ▸ Imprimer… et Exporter en PDF… par les vrais contrôles, aperçu
// avant impression, PDF écrit puis relu (pages, polices incorporées, image, lien cliquable).
public static partial class SelfTest
{
    /// <summary>
    /// Ouvre une boîte modale par l'entrée <paramref name="entry"/> du menu Fichier ; <paramref name="inspect"/> est
    /// appelé pendant que la boîte est ouverte (boucle de messages de ShowDialog), puis la boîte est fermée (Annuler).
    /// </summary>
    static bool FileMenuDialog(MainWindow w, string entry, Action<Window> inspect)
    {
        var file = (Button)w.FindName("FileButton");
        ((IInvokeProvider)UIElementAutomationPeer.CreatePeerForElement(file).GetPattern(PatternInterface.Invoke)).Invoke();
        Pump();
        Idle();
        var menu = w.LastMenu;
        var item = menu?.Items.OfType<MenuItem>().FirstOrDefault(m => HeaderOf(m) == entry);
        if (menu is not { IsOpen: true } || item is null || !item.IsEnabled || !Shown(item))
        {
            if (menu is not null) menu.IsOpen = false;
            return false;
        }
        var before = w.LastOutputWindow;
        bool seen = false;
        // passe dans la boucle de messages de la boîte modale, une fois qu'elle est affichée
        void Probe()
        {
            if (w.LastOutputWindow is { IsLoaded: true } dlg && dlg != before)
            {
                seen = true;
                try { inspect(dlg); }
                finally { dlg.DialogResult = false; }
            }
            else Application.Current.Dispatcher.BeginInvoke(Probe, DispatcherPriority.ApplicationIdle);
        }
        Application.Current.Dispatcher.BeginInvoke(Probe, DispatcherPriority.ApplicationIdle);
        ((IInvokeProvider)UIElementAutomationPeer.CreatePeerForElement(item).GetPattern(PatternInterface.Invoke)).Invoke();
        Pump();
        Idle();
        if (menu.IsOpen) menu.IsOpen = false;
        return seen;
    }

    static void OutputChecks(MainWindow w, string outDir)
    {
        Console.WriteLine("Impression et PDF :");
        int pages = w.CurrentLayout!.Pages.Count;

        // Fichier ▸ Imprimer… : la boîte s'ouvre avec l'aperçu de chaque feuille ; livret = feuilles de 2 pages
        int sheets = -1, bookletSheets = -1;
        Check(FileMenuDialog(w, "Imprimer…", dlg =>
        {
            var pw = (Output.PrintWindow)dlg;
            sheets = pw.PreviewSheets;
            pw.SetMode(Output.PrintMode.Booklet);
            bookletSheets = pw.PreviewSheets;
        }) && sheets == pages && bookletSheets == Core.Output.Booklet.PaddedCount(pages) / 2,
            $"Fichier ▸ Imprimer… : aperçu de {sheets} feuille(s) pour {pages} page(s), livret : {bookletSheets} face(s)");
        Check(ApplicationCommands.Print.CanExecute(null, w), "Ctrl+P (Imprimer) disponible");

        // Fichier ▸ Exporter en PDF… : la boîte s'ouvre sur toutes les pages
        List<int>? chosen = null;
        Check(FileMenuDialog(w, "Exporter en PDF…", dlg => chosen = ((Output.PdfExportWindow)dlg).Chooser.Pages(out _)),
              $"Fichier ▸ Exporter en PDF… : boîte ouverte, {chosen?.Count ?? 0} page(s) choisies");

        // le PDF lui-même (la boîte d'enregistrement de Windows n'est pas pilotable : même code que le bouton), avec une
        // photo JPEG qui doit être incorporée telle quelle
        var bmp = new System.Windows.Media.Imaging.RenderTargetBitmap(60, 30, 96, 96, System.Windows.Media.PixelFormats.Pbgra32);
        var dv = new System.Windows.Media.DrawingVisual();
        using (var dc = dv.RenderOpen()) dc.DrawRectangle(System.Windows.Media.Brushes.SteelBlue, null, new Rect(0, 0, 60, 30));
        bmp.Render(dv);
        var enc = new System.Windows.Media.Imaging.JpegBitmapEncoder { QualityLevel = 90 };
        enc.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(bmp));
        using var jpeg = new MemoryStream();
        enc.Save(jpeg);
        w.FindNext("Deuxième", true, false);
        w.InsertPicture(jpeg.ToArray(), ".jpg", "Photo");
        Pump();
        Idle();
        var pdf = Path.Combine(outDir, "auto-test.pdf");
        if (File.Exists(pdf)) File.Delete(pdf);
        var stats = w.ExportPdfTo(pdf);
        var sum = Output.PdfExporter.Inspect(pdf);
        Check(sum.Pages == pages && sum.Sizes.All(s => s.W > 500 && s.H > 800), $"PDF : {sum.Pages} page(s) A4 ({stats})");
        Check(sum.EmbeddedFonts.Count > 0 && sum.MissingFonts.Count == 0 && stats.TextRuns > 0,
              $"PDF : texte avec polices incorporées ({string.Join(", ", sum.EmbeddedFonts.Select(f => f.Contains('+') ? f[(f.IndexOf('+') + 1)..] : f))})");
        Check(sum.Images >= 1 && stats.OriginalImages >= 1, $"PDF : {sum.Images} image(s), dont {stats.OriginalImages} JPEG incorporé(s) sans recompression");
        Check(sum.Links >= 1, $"PDF : {sum.Links} lien(s) cliquable(s)");
        Check(sum.Title is not null, $"PDF : titre « {sum.Title} »");
        // texte réellement écrit (et non dessiné en contours), avec la table de caractères qui permet de le copier
        // (le gras simulé d'une police sans variante grasse ou une police non incorporable part en contours : toléré une fois)
        Check(stats.TextRuns >= 10 && stats.TextAsOutlines <= 1 && sum.FontsWithoutUnicode == 0,
              $"PDF : texte copiable ({stats.TextRuns} morceaux, {stats.TextAsOutlines} en contours, {sum.FontsWithoutUnicode} police(s) sans table Unicode)");
        ApplicationCommands.Undo.Execute(null, w);
        Pump();
        Idle();
    }
}
