using System.IO;
using System.Windows;
using OpenWrite.Core.Update;
using OpenWrite.Docx;

namespace OpenWrite.App;

/// <summary>
/// Point d'entrée. Modes en ligne de commande (sans fenêtre), utiles pour l'automatisation et les tests :
///   OpenWrite.exe --info document.docx                 résumé du document (version 0.1)
///   OpenWrite.exe --typing-bench doc.docx [n] [--save copie.docx]   banc de saisie (milieu du document, corps répété n fois)
///   OpenWrite.exe --selftest dossier                   auto-test de bout en bout sur la vraie fenêtre (hors écran)
///   OpenWrite.exe --screenshot document.docx|- image.png [largeur hauteur] [scène]
///   OpenWrite.exe --help-topics | --gen-update-key cle.pem | --public-key cle.pem
///   OpenWrite.exe --sign-update installeur.msi version "notes" cle.pem latest.json | --verify-update latest.json
///   OpenWrite.exe --check-update [--download]
/// Sinon : ouvre la fenêtre (avec le fichier passé en argument, le cas échéant).
/// </summary>
public partial class App : Application
{
    [System.Runtime.InteropServices.DllImport("kernel32.dll")]
    static extern bool AttachConsole(int processId);

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        var args = e.Args;
        if (args.Length > 0 && args[0].StartsWith("--"))
        {
            AttachConsole(-1);                                   // sorties visibles en ligne de commande
            Console.OutputEncoding = System.Text.Encoding.UTF8;  // accents des messages et des noms de fichiers
        }
        try
        {
            if (args.Length >= 2 && args[0] == "--info")
            {
                Shutdown(PrintInfo(args[1]));
                return;
            }
            if (args.Length >= 2 && args[0] == "--gen-update-key")
            {
                // crée la clé de signature des mises à jour (à garder secrète, hors du dépôt)
                if (File.Exists(args[1])) { Console.Error.WriteLine("La clé existe déjà : " + args[1]); Shutdown(1); return; }
                var (priv, pub) = UpdateSigner.GenerateKey();
                Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(args[1]))!);
                File.WriteAllText(args[1], priv);
                Console.WriteLine(pub);
                Shutdown(0);
                return;
            }
            if (args.Length >= 2 && args[0] == "--public-key")
            {
                Console.WriteLine(UpdateSigner.PublicKeyOf(File.ReadAllText(args[1])));
                Shutdown(0);
                return;
            }
            if (args.Length >= 6 && args[0] == "--sign-update")
            {
                var m = new UpdateManifest
                {
                    Version = args[2], Url = Path.GetFileName(args[1]), Sha256 = UpdateSigner.Sha256Of(args[1]),
                    Size = new FileInfo(args[1]).Length, Notes = args[3], Published = DateTime.UtcNow,
                };
                UpdateSigner.Sign(m, File.ReadAllText(args[4]));
                File.WriteAllText(args[5], m.ToJson());
                Console.WriteLine($"latest.json signé : {m.Version}, {m.Size / 1048576.0:0.0} Mo, sha256 {m.Sha256[..12]}...");
                Shutdown(0);
                return;
            }
            if (args.Length >= 2 && args[0] == "--verify-update")
            {
                // vérifie hors ligne un latest.json (et l'installeur voisin) avec la clé publique embarquée
                var m = UpdateManifest.Parse(File.ReadAllText(args[1]));
                bool sig = UpdateSigner.Verify(m, Update.UpdateService.PublicKey ?? "");
                var msi = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(args[1]))!, m.Url);
                bool sha = File.Exists(msi) && UpdateSigner.Sha256Of(msi) == m.Sha256;
                Console.WriteLine($"signature : {(sig ? "valide" : "INVALIDE")} ; installeur : {(sha ? "empreinte conforme" : "NON CONFORME")}");
                Shutdown(sig && sha ? 0 : 1);
                return;
            }
            if (args.Length >= 1 && args[0] == "--check-update")
            {
                // sur un fil secondaire : attendre une tâche asynchrone sur le fil WPF provoquerait un interblocage
                var r = Task.Run(() => Update.UpdateService.CheckAsync(false)).GetAwaiter().GetResult();
                Console.WriteLine($"Version installée : {Update.UpdateService.CurrentVersion}; flux : {Update.UpdateService.FeedUrl ?? "(non configuré)"}");
                Console.WriteLine(r.Message + (r.Manifest is { } mm ? $" (en ligne : {mm.Version}, signature valide)" : ""));
                if (r.Manifest is { } m2 && args.Contains("--download"))
                {
                    var f = Task.Run(() => Update.UpdateService.DownloadAsync(m2, null)).GetAwaiter().GetResult();
                    Console.WriteLine("Téléchargé et vérifié (SHA-256 conforme) : " + f);
                }
                Shutdown(r.Manifest is null ? 1 : 0);
                return;
            }
            if (args.Length >= 2 && args[0] == "--layout")
            {
                Shutdown(Layout.LayoutBench.Run(args));
                return;
            }
            if (args.Length >= 1 && args[0] == "--help-topics")
            {
                foreach (var t in Help.HelpService.AllTopics) Console.WriteLine(t);
                Shutdown(0);
                return;
            }
            if (args.Length >= 2 && args[0] == "--typing-bench")
            {
                ShutdownMode = ShutdownMode.OnExplicitShutdown;
                int n = args.Length >= 3 && int.TryParse(args[2], out var rep) ? rep : 1;
                int save = Array.IndexOf(args, "--save");
                Shutdown(SelfTest.TypingBench(args[1], n, save > 0 && save + 1 < args.Length ? args[save + 1] : null));
                return;
            }
            if (args.Length >= 2 && args[0] == "--selftest")
            {
                ShutdownMode = ShutdownMode.OnExplicitShutdown;
                int code;
                try { code = SelfTest.Run(args[1]); }
                catch (Exception ex)
                {
                    Console.Error.WriteLine(ex);       // une fenêtre restée ouverte bloquerait le processus : on sort franchement
                    Environment.Exit(2);
                    return;
                }
                Shutdown(code);
                return;
            }
            if (args.Length >= 3 && args[0] == "--screenshot")
            {
                ShutdownMode = ShutdownMode.OnExplicitShutdown;
                int sw = args.Length >= 5 ? int.Parse(args[3]) : 1402, sh = args.Length >= 5 ? int.Parse(args[4]) : 948;
                var code = SelfTest.Screenshot(args[1] == "-" ? null : args[1], args[2], sw, sh, args.Length >= 6 ? args[5] : null);
                Shutdown(code);
                return;
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            Shutdown(1);
            return;
        }

        Diagnostics.CrashReporter.Install(this);   // mode fenêtre seulement : pas de rapports depuis la ligne de commande
        var win = new MainWindow { TrackRecent = true, EnableRecovery = true };
        MainWindow = win;
        win.Show();
        win.OfferRecovery();
        if (args.Length >= 1 && File.Exists(args[0])) win.OpenFile(args[0]);
    }

    /// <summary>--info : résumé d'un document, comme Fichier ▸ Informations.</summary>
    static int PrintInfo(string path)
    {
        DocxInfo info;
        try { info = DocxInfo.Read(path); }
        catch (DocxReadException ex) { Console.Error.WriteLine(ex.Message); return 1; }
        Console.WriteLine($"{Path.GetFileName(path)} : {DocxInfo.CompatibilityLabel(info.CompatibilityMode)}" +
                          (info.LegacyCompatOptions.Count > 0 ? $" ; options w:compat : {string.Join(", ", info.LegacyCompatOptions)}" : ""));
        Console.WriteLine($"  enregistré par Word : {info.SavedPages?.ToString() ?? "?"} page(s), {info.SavedWords?.ToString() ?? "?"} mots ({info.Application ?? "application inconnue"})");
        Console.WriteLine($"  compté : {info.Words} mots, {info.Paragraphs} paragraphes (dont {info.ListParagraphs} de liste), {info.Tables} tableau(x), {info.Sections} section(s)");
        Console.WriteLine($"  {info.HeadersFooters} en-tête(s)/pied(s), {info.Pictures} image(s), {info.TextBoxes} zone(s) de texte, {info.Footnotes} note(s) de bas de page, {info.Endnotes} note(s) de fin");
        Console.WriteLine($"  {info.Comments} commentaire(s), {info.Revisions} révision(s), {info.Bookmarks} signet(s), {info.Hyperlinks} lien(s), " +
                          $"{info.ContentControls} contrôle(s) de contenu, {info.LegacyFormFields} champ(s) de formulaire, {info.Equations} équation(s){(info.HasMacros ? ", macros" : "")}");
        Console.WriteLine($"  langues : {string.Join(", ", info.Languages)} ; polices : {string.Join(", ", info.Fonts)}");
        Console.WriteLine($"  {info.Parts.Count} parties : {string.Join(" ", info.Parts.Select(p => p.Uri))}");
        return 0;
    }
}
