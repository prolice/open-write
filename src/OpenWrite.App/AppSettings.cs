// Repris d'OpenPub 2.1.3 (src/OpenPub.App/AppSettings.cs).
using System.IO;
using System.Text.Json;

namespace OpenWrite.App;

/// <summary>
/// Reglages de l'utilisateur, dans %APPDATA%\OpenWrite\settings.json : mises a jour (verification
/// automatique, version ignoree) et liste des documents recents.
/// </summary>
public static class AppSettings
{
    public sealed class Data
    {
        public bool AutoCheck { get; set; } = true;
        public string? SkippedVersion { get; set; }
        /// <summary>Documents ouverts ou enregistres recemment, le plus recent en tete.</summary>
        public List<string> RecentFiles { get; set; } = new();
        /// <summary>Verification orthographique pendant la frappe (soulignement rouge).</summary>
        public bool SpellCheck { get; set; } = true;
        /// <summary>Correction typographique francaise a la frappe.</summary>
        public bool AutoCorrect { get; set; } = true;
        /// <summary>Copie de secours automatique des documents modifies (toutes les 5 minutes).</summary>
        public bool AutoSave { get; set; } = true;
        /// <summary>Garder la version precedente a l'enregistrement (« nom.docx.bak »).</summary>
        public bool KeepBackup { get; set; } = true;
        /// <summary>Envoi automatique des rapports d'erreur anonymes (issue GitHub via le relais du site).</summary>
        public bool CrashReports { get; set; } = true;
    }

    static string SettingsPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OpenWrite", "settings.json");

    public static Data Load()
    {
        try { return JsonSerializer.Deserialize<Data>(File.ReadAllText(SettingsPath)) ?? new Data(); }
        catch { return new Data(); }
    }

    public static void Update(Action<Data> change)
    {
        var s = Load();
        change(s);
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
            File.WriteAllText(SettingsPath, JsonSerializer.Serialize(s));
        }
        catch { /* reglages non enregistrables : sans gravite */ }
    }
}

/// <summary>Liste des documents recents (menu Fichier).</summary>
public static class RecentFiles
{
    public const int Max = 8;

    /// <summary>Documents recents qui existent encore.</summary>
    public static IReadOnlyList<string> List() => AppSettings.Load().RecentFiles.Where(File.Exists).Take(Max).ToList();

    /// <summary>Place le document en tete de liste. Les fichiers temporaires (tests, conversions) sont ignores.</summary>
    public static void Add(string path)
    {
        string full;
        try { full = Path.GetFullPath(path); } catch { return; }
        if (full.StartsWith(Path.GetTempPath(), StringComparison.OrdinalIgnoreCase)) return;
        AppSettings.Update(s =>
        {
            s.RecentFiles.RemoveAll(p => p.Equals(full, StringComparison.OrdinalIgnoreCase));
            s.RecentFiles.Insert(0, full);
            if (s.RecentFiles.Count > Max) s.RecentFiles.RemoveRange(Max, s.RecentFiles.Count - Max);
        });
    }

    public static void Remove(string path) =>
        AppSettings.Update(s => s.RecentFiles.RemoveAll(p => p.Equals(path, StringComparison.OrdinalIgnoreCase)));

    public static void Clear() => AppSettings.Update(s => s.RecentFiles.Clear());
}
