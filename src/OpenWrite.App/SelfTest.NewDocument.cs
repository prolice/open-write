using System.IO;
using System.Windows;
using System.Windows.Automation.Peers;
using System.Windows.Automation.Provider;
using System.Windows.Controls;
using System.Windows.Input;
using OpenWrite.Core.Model;
using OpenWrite.Docx;

namespace OpenWrite.App;

// Auto-test de Fichier ▸ Nouveau document (1.1.1), par le vrai menu puis par Ctrl+N : page blanche « Document1 »
// sans fichier (Enregistrer demandera un nom), frappe, copie écrite dans le dossier de l'auto-test (nouveau.docx,
// pour la faire ouvrir par Word). Dernier contrôle : le document de l'auto-test est remplacé.
public static partial class SelfTest
{
    static void NewDocumentChecks(MainWindow w, string outDir)
    {
        Console.WriteLine("Nouveau document :");
        static string TextOf(Block b) => b is Paragraph p ? string.Concat(p.Runs().SelectMany(r => r.Content).OfType<TextContent>().Select(t => t.Value)) : "";
        var file = (Button)w.FindName("FileButton");
        ((IInvokeProvider)UIElementAutomationPeer.CreatePeerForElement(file).GetPattern(PatternInterface.Invoke)).Invoke();
        Pump();
        Idle();
        var item = w.LastMenu?.Items.OfType<MenuItem>().FirstOrDefault(m => HeaderOf(m) == "Nouveau document");
        bool ok = item is { IsEnabled: true } && Shown(item);
        if (ok) { ((IInvokeProvider)UIElementAutomationPeer.CreatePeerForElement(item!).GetPattern(PatternInterface.Invoke)).Invoke(); Pump(); Idle(); }
        if (w.LastMenu is { IsOpen: true } m) m.IsOpen = false;
        var body = w.Body;
        Check(ok && body is not null && body.Paragraphs().Count() == 1 && !body.Paragraphs().First().Runs().Any() && w.WithoutFileForTest
              && w.Title.StartsWith("Document") && w.CurrentLayout?.Pages.Count == 1,
              $"Fichier ▸ Nouveau document : page blanche sans fichier (« {w.Title} »)");

        var v = w.PagesViewForTest!;
        foreach (var ch in "Bonjour")
            v.RaiseEvent(new TextCompositionEventArgs(Keyboard.PrimaryDevice, new TextComposition(InputManager.Current, v, ch.ToString()))
                { RoutedEvent = TextCompositionManager.TextInputEvent });
        Pump();
        Idle();
        Check(w.IsModified && TextOf(w.Body!.Blocks[0]) == "Bonjour", "frappe dans le document vierge");
        var copy = Path.Combine(outDir, "nouveau.docx");
        try
        {
            using var ms = new MemoryStream();
            w.SavePackageCopy(ms);
            File.WriteAllBytes(copy, ms.ToArray());
            var info = DocxInfo.Read(copy);
            Check(TextOf(DocxPackage.Open(copy).Main!.Stories().First().Blocks[0]) == "Bonjour" && info.CompatibilityMode == 15, "copie écrite et relue (nouveau.docx)");
        }
        catch (Exception ex) { Check(false, "copie du nouveau document : " + ex.Message); }

        // Ctrl+N (geste de ApplicationCommands.New, comme Ctrl+Z pour Undo) : un autre document vierge
        // (la confirmation est sautée pendant l'auto-test)
        var title = w.Title;
        bool gesture = ApplicationCommands.New.InputGestures.OfType<KeyGesture>().Any(g => g.Key == Key.N && g.Modifiers == ModifierKeys.Control);
        ApplicationCommands.New.Execute(null, w);
        Pump();
        Idle();
        Check(gesture && w.Title != title && w.Title.StartsWith("Document") && !w.IsModified && w.Body!.Paragraphs().Count() == 1, $"Ctrl+N : nouveau document vierge (« {w.Title} »)");
    }
}
