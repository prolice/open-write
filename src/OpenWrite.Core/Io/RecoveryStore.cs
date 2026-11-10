// Repris d'OpenPub 2.1.3 (src/OpenPub.Core/Io/RecoveryStore.cs) : la copie est écrite par l'appelant
// (le document sait s'enregistrer), le magasin ne dépend donc plus du format.
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace OpenWrite.Core.Io;

/// <summary>Copie de secours d'un document (enregistrement automatique).</summary>
public sealed record RecoveryEntry(string Id, string CopyPath, string? OriginalPath, string Title, DateTime SavedUtc);

/// <summary>
/// Copies de secours des documents ouverts, dans un dossier dédié (jamais à la place du fichier de l'utilisateur) :
/// « id.docx » + « id.json » (fichier d'origine, titre, date). Une copie restée après la fermeture d'OpenWrite
/// signale une fermeture inattendue : elle est proposée au démarrage suivant.
/// </summary>
public sealed class RecoveryStore
{
    public string Dir { get; }

    /// <summary>Extension des copies (format natif d'OpenWrite).</summary>
    public const string Extension = ".docx";

    public RecoveryStore(string dir) => Dir = dir;

    /// <summary>Dossier par défaut : %LOCALAPPDATA%\OpenWrite\Recovery.</summary>
    public static string DefaultDir => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OpenWrite", "Recovery");

    /// <summary>Identifiant stable d'un document enregistré (même fichier = même copie) ; sinon propre à la session.</summary>
    public static string IdFor(string? originalPath, string sessionId)
    {
        if (originalPath is null) return "sans-titre-" + sessionId;
        var hash = SHA1.HashData(Encoding.UTF8.GetBytes(Path.GetFullPath(originalPath).ToLowerInvariant()));
        return Convert.ToHexString(hash)[..12].ToLowerInvariant();
    }

    sealed class Info
    {
        public string? Original { get; set; }
        public string Title { get; set; } = "";
        public DateTime SavedUtc { get; set; }
    }

    /// <summary>
    /// Enregistre une copie : <paramref name="writeCopy"/> reçoit le chemin à écrire. La copie précédente n'est
    /// remplacée qu'une fois la nouvelle complète (une erreur pendant l'écriture laisse l'ancienne intacte).
    /// </summary>
    public void Save(string id, string? originalPath, string title, Action<string> writeCopy)
    {
        Directory.CreateDirectory(Dir);
        SafeFile.Write(Path.Combine(Dir, id + Extension), writeCopy, keepBackup: false);
        var info = new Info { Original = originalPath is null ? null : Path.GetFullPath(originalPath), Title = title, SavedUtc = DateTime.UtcNow };
        File.WriteAllText(Path.Combine(Dir, id + ".json"), JsonSerializer.Serialize(info));
    }

    /// <summary>Copies présentes, la plus récente en tête.</summary>
    public List<RecoveryEntry> List()
    {
        var res = new List<RecoveryEntry>();
        if (!Directory.Exists(Dir)) return res;
        foreach (var json in Directory.GetFiles(Dir, "*.json"))
        {
            var id = Path.GetFileNameWithoutExtension(json);
            var copy = Path.Combine(Dir, id + Extension);
            if (!File.Exists(copy)) { TryDelete(json); continue; }
            try
            {
                var info = JsonSerializer.Deserialize<Info>(File.ReadAllText(json)) ?? new Info();
                res.Add(new RecoveryEntry(id, copy, info.Original, info.Title, info.SavedUtc));
            }
            catch (JsonException) { TryDelete(json); }
        }
        return res.OrderByDescending(e => e.SavedUtc).ToList();
    }

    public void Delete(string id)
    {
        TryDelete(Path.Combine(Dir, id + Extension));
        TryDelete(Path.Combine(Dir, id + ".json"));
    }

    static void TryDelete(string f)
    {
        try { if (File.Exists(f)) File.Delete(f); } catch (IOException) { } catch (UnauthorizedAccessException) { }
    }
}
