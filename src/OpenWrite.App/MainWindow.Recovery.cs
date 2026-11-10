// Repris d'OpenPub 2.1.3 (src/OpenPub.App/MainWindow.Recovery.cs), adapté au package .docx.
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Threading;
using OpenWrite.Core.Io;

namespace OpenWrite.App;

// Sûreté des documents : copie de secours automatique, récupération après une fermeture inattendue,
// détection d'un fichier modifié par un autre programme.
public partial class MainWindow
{
    /// <summary>Copies de secours actives (application normale ; jamais pendant les tests).</summary>
    public bool EnableRecovery { get; set; }

    public static TimeSpan AutoSaveInterval { get; } = TimeSpan.FromMinutes(5);

    internal RecoveryStore Recovery { get; set; } = new(RecoveryStore.DefaultDir);
    string _sessionId = Guid.NewGuid().ToString("N")[..8];
    string? _recoveryId;
    DispatcherTimer? _autoSaveTimer;
    /// <summary>Copie récupérée dont le fichier d'origine a disparu : « Enregistrer » demande un nom.</summary>
    bool _recoveredWithoutTarget;

    void EnsureAutoSaveTimer()
    {
        if (!EnableRecovery || _autoSaveTimer is not null) return;
        _autoSaveTimer = new DispatcherTimer(DispatcherPriority.Background) { Interval = AutoSaveInterval };
        _autoSaveTimer.Tick += (_, _) => AutoSaveNow();
        _autoSaveTimer.Start();
    }

    /// <summary>Enregistre une copie de secours si le document a des modifications non enregistrées.</summary>
    internal bool AutoSaveNow()
    {
        if (!EnableRecovery || !IsModified || _pkg is null || !AppSettings.Load().AutoSave) return false;
        try
        {
            var id = RecoveryStore.IdFor(_recoveredWithoutTarget ? null : _path, _sessionId);
            if (_recoveryId is not null && _recoveryId != id) Recovery.Delete(_recoveryId);
            var pkg = _pkg;
            Recovery.Save(id, _recoveredWithoutTarget ? null : _path, Path.GetFileName(_path) ?? "Sans titre", tmp =>
            {
                using var fs = File.Create(tmp);
                pkg.Save(fs);                      // copie : l'état du document (parties modifiées) ne change pas
            });
            _recoveryId = id;
            StatusText.Text = $"Copie de secours enregistrée à {DateTime.Now:HH:mm}.";
            return true;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            StatusText.Text = "Copie de secours impossible : " + ex.Message;
            return false;
        }
    }

    void DiscardRecovery()
    {
        if (_recoveryId is null) return;
        Recovery.Delete(_recoveryId);
        _recoveryId = null;
    }

    /// <summary>Au démarrage : propose les copies de secours laissées par une fermeture inattendue.</summary>
    public void OfferRecovery()
    {
        if (!EnableRecovery) return;
        foreach (var entry in Recovery.List())
        {
            var when = entry.SavedUtc.ToLocalTime().ToString("dddd d MMMM 'à' HH:mm", CultureInfo.GetCultureInfo("fr-FR"));
            var r = MessageBox.Show(this,
                $"OpenWrite n'a pas été fermé normalement. Une copie de secours de « {entry.Title} » a été enregistrée {when}.\n\n" +
                "Oui : ouvrir la copie récupérée (pensez à l'enregistrer)\nNon : supprimer cette copie\nAnnuler : décider au prochain démarrage",
                "Récupérer un document", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
            if (r == MessageBoxResult.Cancel) return;
            if (r == MessageBoxResult.No) { Recovery.Delete(entry.Id); continue; }
            OpenRecovered(entry);
            return;
        }
    }

    /// <summary>Ouvre une copie de secours : le document retrouve son fichier d'origine et reste « modifié ».</summary>
    internal void OpenRecovered(RecoveryEntry entry)
    {
        OpenFile(entry.CopyPath, remember: false);
        if (_pkg is null) return;
        bool hasTarget = entry.OriginalPath is { } o && Directory.Exists(Path.GetDirectoryName(o));
        _path = hasTarget ? entry.OriginalPath : entry.CopyPath;
        _recoveredWithoutTarget = !hasTarget;
        if (entry.Id.StartsWith("sans-titre-")) _sessionId = entry.Id["sans-titre-".Length..];
        _recoveryId = entry.Id;
        _recovered = true;
        RecordFileStamp();
        UpdateStatus();
        StatusText.Text = "Document récupéré : enregistrez-le (Ctrl+S) pour garder ces modifications.";
    }

    // ---------------------------------------------------------------- fichier modifié ailleurs

    (DateTime Time, long Size)? _fileStamp;
    bool _externalChangeWarned;

    static (DateTime, long)? Stamp(string? path)
    {
        if (path is null || !File.Exists(path)) return null;
        var fi = new FileInfo(path);
        return (fi.LastWriteTimeUtc, fi.Length);
    }

    void RecordFileStamp()
    {
        _fileStamp = Stamp(_path);
        _externalChangeWarned = false;
    }

    bool FileChangedOnDisk()
    {
        if (_path is null || _fileStamp is null) return false;
        return Stamp(_path) != _fileStamp;
    }

    void OnActivatedCheckFile(object? sender, EventArgs e)
    {
        if (_externalChangeWarned || !FileChangedOnDisk()) return;
        _externalChangeWarned = true;
        StatusText.Text = "Attention : le fichier a été modifié par un autre programme depuis son ouverture. Enregistrer remplacera cette autre version.";
    }
}
