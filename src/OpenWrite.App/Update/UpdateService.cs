// Repris d'OpenPub 2.1.3 (src/OpenPub.App/Update/UpdateService.cs).
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Text.Json;
using OpenWrite.Core.Update;

namespace OpenWrite.App.Update;

/// <summary>
/// Mises a jour automatiques : lit latest.json sur le site de mises a jour (Azure Static Web Apps),
/// verifie la signature avec la cle publique embarquee, telecharge l'installeur MSI, controle son
/// empreinte SHA-256 puis lance l'installation (par utilisateur, sans droits administrateur).
/// </summary>
public static class UpdateService
{
    static readonly HttpClient Http = new() { Timeout = TimeSpan.FromMinutes(10) };

    static string? Meta(string key) => Assembly.GetExecutingAssembly().GetCustomAttributes<AssemblyMetadataAttribute>()
        .FirstOrDefault(a => a.Key == key)?.Value;

    /// <summary>Adresse de latest.json (definie a la compilation).</summary>
    public static string? FeedUrl => Meta("UpdateFeed") is { Length: > 0 } u ? u : null;
    public static string? PublicKey => Meta("UpdatePublicKey") is { Length: > 0 } k ? k : null;

    public static Version CurrentVersion
    {
        get
        {
            var v = Assembly.GetExecutingAssembly().GetName().Version ?? new Version(0, 0);
            return new Version(v.Major, v.Minor, Math.Max(0, v.Build));
        }
    }

    /// <summary>L'application tourne-t-elle depuis son dossier d'installation (et pas depuis les sources) ?</summary>
    public static bool IsInstalled
    {
        get
        {
            var exe = Environment.ProcessPath ?? "";
            var installDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "OpenWrite");
            return exe.StartsWith(installDir, StringComparison.OrdinalIgnoreCase) && exe.EndsWith("OpenWrite.exe", StringComparison.OrdinalIgnoreCase);
        }
    }

    // ---------------------------------------------------------------- reglages (version ignoree, verification automatique)

    public static bool AutoCheck
    {
        get => AppSettings.Load().AutoCheck;
        set => AppSettings.Update(s => s.AutoCheck = value);
    }

    public static void SkipVersion(string version) => AppSettings.Update(s => s.SkippedVersion = version);

    // ---------------------------------------------------------------- verification

    public sealed record CheckResult(UpdateManifest? Manifest, bool IsNewer, string Message);

    /// <summary>Lit et verifie latest.json. Ne leve pas d'exception : le resultat porte un message.</summary>
    public static async Task<CheckResult> CheckAsync(bool honorSkip, CancellationToken ct = default)
    {
        if (FeedUrl is not { } feed || PublicKey is not { } key)
            return new CheckResult(null, false, "Les mises à jour ne sont pas configurées dans cette version.");
        UpdateManifest m;
        try
        {
            var json = await Http.GetStringAsync(feed + (feed.Contains('?') ? "&" : "?") + "t=" + DateTime.UtcNow.Ticks, ct).ConfigureAwait(false);
            m = UpdateManifest.Parse(json);
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or JsonException or InvalidDataException)
        {
            return new CheckResult(null, false, "Impossible de joindre le serveur de mises à jour (" + ex.Message + ").");
        }
        if (!UpdateSigner.Verify(m, key))
            return new CheckResult(null, false, "Mise à jour refusée : signature invalide.");
        bool newer = m.ParsedVersion > CurrentVersion;
        if (newer && honorSkip && AppSettings.Load().SkippedVersion == m.Version) newer = false;
        return new CheckResult(m, newer, newer ? $"OpenWrite {m.Version} est disponible." : $"OpenWrite est à jour (version {CurrentVersion}).");
    }

    /// <summary>Telecharge l'installeur dans un dossier temporaire et verifie son empreinte.</summary>
    public static async Task<string> DownloadAsync(UpdateManifest m, IProgress<double>? progress, CancellationToken ct = default)
    {
        var url = new Uri(new Uri(FeedUrl!), m.Url);
        if (url.Scheme != Uri.UriSchemeHttps) throw new InvalidOperationException("Adresse de téléchargement non sécurisée.");
        var dir = Path.Combine(Path.GetTempPath(), "OpenWrite-update");
        Directory.CreateDirectory(dir);
        var file = Path.Combine(dir, $"OpenWrite-{m.Version}.msi");
        using (var resp = await Http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, ct).ConfigureAwait(false))
        {
            resp.EnsureSuccessStatusCode();
            var total = resp.Content.Headers.ContentLength ?? m.Size;
            await using var src = await resp.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
            await using var dst = File.Create(file);
            var buf = new byte[81920];
            long done = 0;
            int n;
            while ((n = await src.ReadAsync(buf, ct).ConfigureAwait(false)) > 0)
            {
                await dst.WriteAsync(buf.AsMemory(0, n), ct).ConfigureAwait(false);
                done += n;
                if (total > 0) progress?.Report((double)done / total);
            }
        }
        var sha = UpdateSigner.Sha256Of(file);
        if (!string.Equals(sha, m.Sha256, StringComparison.OrdinalIgnoreCase) || (m.Size > 0 && new FileInfo(file).Length != m.Size))
        {
            File.Delete(file);
            throw new InvalidDataException("Le fichier téléchargé ne correspond pas à l'empreinte signée : installation annulée.");
        }
        return file;
    }

    /// <summary>
    /// Lance l'installation puis la reouverture d'OpenWrite (avec le document en cours) ; l'appelant doit
    /// ensuite fermer l'application pour liberer OpenWrite.exe.
    /// </summary>
    public static void InstallAndRestart(string msi, string? reopen)
    {
        var exe = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "OpenWrite", "OpenWrite.exe");
        var log = Path.Combine(Path.GetDirectoryName(msi)!, "install.log");
        // attend la fermeture d'OpenWrite, installe (fenetre de progression, sans question), puis relance
        var script =
            $"Start-Sleep -Seconds 2; " +
            $"$p = Start-Process msiexec.exe -ArgumentList '/i','\"{msi}\"','/passive','/norestart','/l*v','\"{log}\"' -Wait -PassThru; " +
            $"if ($p.ExitCode -eq 0 -or $p.ExitCode -eq 3010) {{ Start-Process '{exe}' {(reopen is null ? "" : $"-ArgumentList '\"{reopen.Replace("'", "''")}\"'")} }}";
        Process.Start(new ProcessStartInfo("powershell.exe", $"-NoProfile -WindowStyle Hidden -Command \"{script.Replace("\"", "\\\"")}\"")
        {
            UseShellExecute = false, CreateNoWindow = true,
        });
    }
}
