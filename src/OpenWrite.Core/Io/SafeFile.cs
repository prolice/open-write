namespace OpenWrite.Core.Io;

/// <summary>
/// Enregistrement sûr (principe d'OpenPub 2.1.3, OpubPackage.Save) : le contenu est écrit dans un fichier
/// temporaire du même dossier, puis remplace la cible en une seule opération. Une erreur (disque plein, plantage
/// pendant l'écriture) laisse le fichier d'origine intact. Option : garder la version précédente (« nom.docx.bak »).
/// </summary>
public static class SafeFile
{
    public const int RetryCount = 5, RetryDelayMs = 100;

    public static void Write(string path, Action<string> write, bool keepBackup)
    {
        var full = Path.GetFullPath(path);
        var dir = Path.GetDirectoryName(full)!;
        Directory.CreateDirectory(dir);
        var tmp = Path.Combine(dir, $".{Path.GetFileName(full)}.{Guid.NewGuid():N}.tmp");
        try
        {
            write(tmp);
            if (!File.Exists(tmp)) throw new IOException("Le fichier n'a pas été écrit.");
            for (int attempt = 1; ; attempt++)
            {
                try
                {
                    // File.Replace garde les attributs et les droits du fichier d'origine
                    if (File.Exists(full)) File.Replace(tmp, full, keepBackup ? full + ".bak" : null, ignoreMetadataErrors: true);
                    else File.Move(tmp, full);
                    break;
                }
                catch (IOException) when (attempt < RetryCount)
                {
                    // fichier brièvement verrouillé par l'antivirus ou l'indexation de Windows juste après une écriture
                    // (observé sur le banc d'aller-retour : « Impossible de supprimer le fichier à remplacer ») : on réessaie
                    Thread.Sleep(RetryDelayMs * attempt);
                }
            }
        }
        finally
        {
            try { if (File.Exists(tmp)) File.Delete(tmp); } catch (IOException) { } catch (UnauthorizedAccessException) { }
        }
    }
}
