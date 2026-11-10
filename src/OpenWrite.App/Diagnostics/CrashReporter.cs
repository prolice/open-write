// Repris d'OpenPub 2.1.3 (src/OpenPub.App/Diagnostics/CrashReporter.cs).
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Threading;
using OpenWrite.Core.Diagnostics;

namespace OpenWrite.App.Diagnostics;

/// <summary>
/// Signalement automatique des erreurs : intercepte toutes les exceptions non gerees (interface, threads,
/// taches), ecrit un rapport anonymise dans la file locale puis l'envoie au relais du site
/// (AssemblyMetadata « CrashReportUrl », fonction Azure /api/crash) qui cree ou complete l'issue GitHub.
/// Le jeton GitHub reste cote serveur : il n'est JAMAIS dans l'application (depot public, exe decompilable).
///
/// Envoi seulement depuis une version installee (pas pendant le developpement ni les tests), sauf
/// OPENWRITE_CRASH_REPORTS=1 ; desactivable par l'utilisateur (menu Fichier).
/// </summary>
public static class CrashReporter
{
    static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(15) };
    static readonly CrashStore Store = new(CrashStore.DefaultDir);
    static int _shown;   // une seule fenetre d'erreur a la fois
    // une erreur qui se repete (ex. a chaque affichage) n'est enregistree et montree qu'une fois par session
    static readonly HashSet<string> Seen = new();

    public static string? Endpoint => Assembly.GetExecutingAssembly().GetCustomAttributes<AssemblyMetadataAttribute>()
        .FirstOrDefault(a => a.Key == "CrashReportUrl")?.Value is { Length: > 0 } u ? u : null;

    public static bool Enabled
    {
        get => AppSettings.Load().CrashReports;
        set => AppSettings.Update(s => s.CrashReports = value);
    }

    static bool CanSend => Endpoint is not null && Enabled
        && (Update.UpdateService.IsInstalled || Environment.GetEnvironmentVariable("OPENWRITE_CRASH_REPORTS") == "1");

    static string Version => Update.UpdateService.CurrentVersion.ToString();

    /// <summary>Branche les gestionnaires globaux (mode fenetre uniquement) et envoie les rapports restes en attente.</summary>
    public static void Install(Application app)
    {
        app.DispatcherUnhandledException += OnDispatcher;
        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
        {
            // processus condamne : on ecrit seulement sur disque, l'envoi se fera au prochain demarrage
            if (e.ExceptionObject is Exception ex) Record(ex, "fatale", fatal: true, send: false);
        };
        TaskScheduler.UnobservedTaskException += (_, e) =>
        {
            e.SetObserved();
            Record(e.Exception, "tâche", fatal: false, send: true);
        };
        _ = Task.Run(FlushAsync);
    }

    /// <summary>Signale une erreur interceptee par le code (enregistrement, ouverture…) sans interrompre l'utilisateur.</summary>
    public static void Report(Exception ex, string context)
    {
        if (IsEnvironmental(ex)) return;   // disque plein, fichier verrouille, acces refuse : pas un defaut d'OpenWrite
        Record(ex, context, fatal: false, send: true);
    }

    static bool IsEnvironmental(Exception ex) =>
        ex is UnauthorizedAccessException or FileNotFoundException or DirectoryNotFoundException or PathTooLongException
        || (ex is IOException && ex.GetType() == typeof(IOException));

    static void OnDispatcher(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        bool first = Record(e.Exception, "interface", fatal: false, send: true);
        e.Handled = true;   // on garde l'application ouverte : l'utilisateur peut enregistrer son travail
        if (!first || Interlocked.Exchange(ref _shown, 1) == 1) return;
        try
        {
            var msg = "Une erreur inattendue s'est produite :\n\n" + e.Exception.GetBaseException().Message + "\n\n"
                + (CanSend ? "Elle a été signalée automatiquement (rapport anonyme, sans le contenu du document)."
                           : "Un rapport a été enregistré sur cet ordinateur.")
                + "\n\nEnregistrez votre document sous un autre nom par précaution. Continuer ?";
            var r = MessageBox.Show(Application.Current?.MainWindow ?? new Window(), msg, "OpenWrite — erreur",
                MessageBoxButton.YesNo, MessageBoxImage.Error);
            if (r == MessageBoxResult.No) Application.Current?.Shutdown(1);
        }
        catch { /* la fenetre d'erreur ne doit pas provoquer une seconde erreur */ }
        finally { Interlocked.Exchange(ref _shown, 0); }
    }

    /// <summary>Enregistre (et envoie) le rapport ; faux si cette erreur a deja ete signalee pendant la session.</summary>
    static bool Record(Exception ex, string context, bool fatal, bool send)
    {
        try
        {
            var report = CrashReport.From(ex, context, fatal, Version);
            lock (Seen) if (!Seen.Add(report.Signature)) return false;
            Store.Save(report);
            if (send) _ = Task.Run(FlushAsync);
        }
        catch { /* le signalement ne doit jamais aggraver la situation */ }
        return true;
    }

    static readonly SemaphoreSlim Gate = new(1, 1);

    /// <summary>Envoie les rapports en attente ; ceux qui echouent (hors ligne) restent pour la prochaine fois.</summary>
    public static async Task FlushAsync()
    {
        if (!CanSend || !await Gate.WaitAsync(0).ConfigureAwait(false)) return;
        try
        {
            foreach (var (path, report) in Store.Pending())
            {
                try
                {
                    using var content = new StringContent(report.ToJson(), Encoding.UTF8, "application/json");
                    using var resp = await Http.PostAsync(Endpoint, content).ConfigureAwait(false);
                    // 2xx : traite ; 4xx : rapport refuse par le relais (inutile de reessayer) ; 5xx/429 : plus tard
                    if (resp.IsSuccessStatusCode || ((int)resp.StatusCode is >= 400 and < 500 and not 429)) CrashStore.TryDelete(path);
                    else break;
                }
                catch (Exception e) when (e is HttpRequestException or TaskCanceledException) { break; }
            }
        }
        catch { }
        finally { Gate.Release(); }
    }
}
