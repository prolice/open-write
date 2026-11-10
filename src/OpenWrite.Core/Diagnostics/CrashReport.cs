// Repris d'OpenPub 2.1.3 (src/OpenPub.Core/Diagnostics/CrashReport.cs), adapté à OpenWrite.
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace OpenWrite.Core.Diagnostics;

/// <summary>
/// Rapport d'erreur anonymise, envoye au relais du site (/api/crash) qui ouvre ou complete une issue GitHub.
/// Le depot est PUBLIC : le rapport ne contient jamais le contenu du document, ni les noms de fichiers, ni le
/// nom de l'utilisateur (voir <see cref="Scrub"/>). La signature regroupe les occurrences d'une meme erreur
/// (type + premieres methodes de la pile, sans numeros de ligne) dans une seule issue.
/// </summary>
public sealed class CrashReport
{
    public string Signature { get; set; } = "";
    public string Version { get; set; } = "";
    /// <summary>Origine : « interface », « fatale », « tâche », ou une operation (« enregistrement »…).</summary>
    public string Context { get; set; } = "";
    public bool Fatal { get; set; }
    public string Type { get; set; } = "";
    public string Message { get; set; } = "";
    public string Stack { get; set; } = "";
    public string Os { get; set; } = "";
    public string Runtime { get; set; } = "";
    public string Culture { get; set; } = "";
    public DateTime TimeUtc { get; set; }

    public const int MaxMessage = 600, MaxStack = 12000;

    public static CrashReport From(Exception ex, string context, bool fatal, string version)
    {
        var inner = ex;
        while (inner is AggregateException { InnerExceptions.Count: 1 } or System.Reflection.TargetInvocationException && inner.InnerException is not null)
            inner = inner.InnerException!;
        var type = inner.GetType().FullName ?? inner.GetType().Name;
        var stack = Scrub(inner.ToString());
        return new CrashReport
        {
            Signature = SignatureOf(type, inner.StackTrace ?? ""),
            Version = version,
            Context = context,
            Fatal = fatal,
            Type = type,
            Message = Truncate(Scrub(inner.Message), MaxMessage),
            Stack = Truncate(stack, MaxStack),
            Os = Environment.OSVersion.VersionString + (Environment.Is64BitOperatingSystem ? " x64" : " x86"),
            Runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
            Culture = System.Globalization.CultureInfo.CurrentUICulture.Name,
            TimeUtc = DateTime.UtcNow,
        };
    }

    static readonly Regex FrameRx = new(@"^\s*at\s+(?<m>[^\(]+)\(", RegexOptions.Multiline);

    /// <summary>
    /// Empreinte stable d'une erreur : type + 5 premieres methodes d'OpenWrite de la pile (a defaut, les 5 premieres
    /// tout court). Ni numeros de ligne ni message : la meme erreur garde la meme signature d'une version a l'autre.
    /// </summary>
    public static string SignatureOf(string type, string stackTrace)
    {
        var frames = FrameRx.Matches(stackTrace).Select(m => m.Groups["m"].Value.Trim()).ToList();
        var own = frames.Where(f => f.StartsWith("OpenWrite.", StringComparison.Ordinal)).Take(5).ToList();
        var key = type + "|" + string.Join("|", own.Count > 0 ? own : frames.Take(5));
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(key)))[..12].ToLowerInvariant();
    }

    static readonly Regex UserDirRx = new(@"[A-Za-z]:\\(Users|Documents and Settings)\\[^\\/:*?""<>|\r\n]+", RegexOptions.IgnoreCase);
    static readonly Regex OneDriveRx = new(@"\\OneDrive[^\\]*\\", RegexOptions.IgnoreCase);
    static readonly Regex UncRx = new(@"\\\\[^\\\s]+\\[^\\\s]+");
    static readonly Regex EmailRx = new(@"[\w.+-]+@[\w-]+(\.[\w-]+)+");
    static readonly Regex FileNameRx = new(@"(?<dir>[\\/])(?<name>[^\\/:*?""<>|\r\n]+)\.(?<ext>docx?|dotx|docm|rtf|odt|json|png|jpe?g|gif|bmp|tiff?|pdf|csv|xlsx?|bak|html?)\b", RegexOptions.IgnoreCase);
    // texte entre guillemets dans un message (souvent du contenu saisi) ; les identifiants courts sont gardes
    static readonly Regex QuotedRx = new(@"(?<q>['""«“])(?<v>[^'""»”\r\n]{1,400})(['""»”])");
    static readonly Regex IdentRx = new(@"^[A-Za-z_][\w.\-]{0,47}$");
    static readonly Regex LineInfoRx = new(@" in (?<path>[^\r\n]+?):line (?<n>\d+)");

    /// <summary>
    /// Anonymise un texte d'erreur : dossier de l'utilisateur → %USERPROFILE%, partages reseau, adresses e-mail,
    /// noms de documents/images → « … », texte cite entre guillemets (sauf identifiants techniques), chemins des
    /// sources reduits au nom du fichier .cs.
    /// </summary>
    public static string Scrub(string text)
    {
        if (string.IsNullOrEmpty(text)) return "";
        var s = LineInfoRx.Replace(text, m => " in " + Path.GetFileName(m.Groups["path"].Value.Replace('/', '\\').Split('\\').Last()) + ":line " + m.Groups["n"].Value);
        s = UserDirRx.Replace(s, "%USERPROFILE%");
        s = OneDriveRx.Replace(s, @"\OneDrive\");
        s = UncRx.Replace(s, @"\\serveur\partage");
        s = EmailRx.Replace(s, "<e-mail>");
        s = FileNameRx.Replace(s, m => m.Groups["dir"].Value + "…." + m.Groups["ext"].Value);
        s = QuotedRx.Replace(s, m => IdentRx.IsMatch(m.Groups["v"].Value) ? m.Value : m.Groups["q"].Value + "…" + m.Value[^1]);
        return s;
    }

    static string Truncate(string s, int max) => s.Length <= max ? s : s[..max] + "…";

    static readonly JsonSerializerOptions Json = new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase, WriteIndented = false };

    public string ToJson() => JsonSerializer.Serialize(this, Json);

    public static CrashReport? Parse(string json)
    {
        try { return JsonSerializer.Deserialize<CrashReport>(json, Json); }
        catch (JsonException) { return null; }
    }
}

/// <summary>
/// File d'attente locale des rapports (%LOCALAPPDATA%\OpenWrite\Crashes) : chaque rapport est ecrit sur disque
/// AVANT l'envoi, si bien qu'une erreur fatale (processus tue) est envoyee au demarrage suivant.
/// </summary>
public sealed class CrashStore
{
    public string Dir { get; }
    /// <summary>Au-dela, les plus anciens sont supprimes (pas d'accumulation hors ligne).</summary>
    public const int MaxPending = 20;

    public CrashStore(string dir) => Dir = dir;

    public static string DefaultDir => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OpenWrite", "Crashes");

    public string Save(CrashReport r)
    {
        Directory.CreateDirectory(Dir);
        var path = Path.Combine(Dir, $"{r.TimeUtc:yyyyMMdd-HHmmssfff}-{r.Signature}.json");
        File.WriteAllText(path, r.ToJson());
        foreach (var old in Pending().SkipLast(MaxPending)) TryDelete(old.Path);
        return path;
    }

    /// <summary>Rapports en attente d'envoi, du plus ancien au plus recent.</summary>
    public IReadOnlyList<(string Path, CrashReport Report)> Pending()
    {
        if (!Directory.Exists(Dir)) return Array.Empty<(string, CrashReport)>();
        var list = new List<(string, CrashReport)>();
        foreach (var f in Directory.GetFiles(Dir, "*.json").Order(StringComparer.Ordinal))
        {
            CrashReport? r = null;
            try { r = CrashReport.Parse(File.ReadAllText(f)); } catch (IOException) { }
            if (r is null) TryDelete(f); else list.Add((f, r));
        }
        return list;
    }

    public static void TryDelete(string path)
    {
        try { File.Delete(path); } catch { /* sans gravite */ }
    }
}
