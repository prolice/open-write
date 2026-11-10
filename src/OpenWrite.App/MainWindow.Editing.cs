using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;
using OpenWrite.Docx;

namespace OpenWrite.App;

// Version 0.2 : le document ouvert est chargé dans le modèle (DocxPackage) ; les cases à cocher des formulaires se
// cochent (annulable), le document s'enregistre sans perte (seules les parties modifiées sont régénérées, le reste est
// recopié octet pour octet), avec la version précédente en .bak et les copies de secours d'OpenPub.
public partial class MainWindow
{
    DocxPackage? _pkg;
    StyleResolver? _styles;
    readonly UndoJournal _journal = new();
    /// <summary>Document récupéré d'une copie de secours : modifié par rapport à son fichier même sans entrée d'annulation.</summary>
    bool _recovered;

    /// <summary>Pas de questions (captures, auto-test).</summary>
    internal bool SkipConfirm { get; set; }

    internal bool IsModified => _pkg is not null && (_journal.IsModified || _recovered);

    internal Story? Body => _pkg?.Main?.Stories().FirstOrDefault();

    /// <summary>Charge le modèle du document (en plus du résumé DocxInfo) ; faux si le fichier est illisible.</summary>
    bool LoadPackage(string path)
    {
        try
        {
            _pkg = DocxPackage.Open(path);
            _styles = new StyleResolver(_pkg.StyleSheet());
        }
        catch (Exception ex) when (ex is DocxReadException or IOException or UnauthorizedAccessException or System.Xml.XmlException)
        {
            _pkg = null;
            _styles = null;
            MessageBox.Show(this, $"Impossible d'ouvrir « {Path.GetFileName(path)} ».\n\n{ex.Message}", "OpenWrite", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }
        _journal.Clear();
        _recovered = false;
        return true;
    }

    void InitEditing()
    {
        CommandBindings.Add(new CommandBinding(ApplicationCommands.Save, (_, _) => Save(saveAs: false), CanSave));
        CommandBindings.Add(new CommandBinding(ApplicationCommands.SaveAs, (_, _) => Save(saveAs: true), CanSave));
        CommandBindings.Add(new CommandBinding(ApplicationCommands.Undo, (_, _) => UndoRedo(undo: true), (_, e) => e.CanExecute = _journal.CanUndo));
        CommandBindings.Add(new CommandBinding(ApplicationCommands.Redo, (_, _) => UndoRedo(undo: false), (_, e) => e.CanExecute = _journal.CanRedo));
        InputBindings.Add(new KeyBinding(ApplicationCommands.SaveAs, Key.S, ModifierKeys.Control | ModifierKeys.Shift));
        InputBindings.Add(new KeyBinding(ApplicationCommands.SaveAs, Key.F12, ModifierKeys.None));
        InputBindings.Add(new KeyBinding(ApplicationCommands.Redo, Key.Y, ModifierKeys.Control));
        // toute modification marque la partie à régénérer et met à jour l'affichage
        _journal.Changed += root =>
        {
            if (_pkg?.Stories.Values.FirstOrDefault(p => p.Stories().Contains(root)) is { } part) part.Dirty = true;
        };
        Closing += OnClosing;
        Activated += OnActivatedCheckFile;
    }

    void CanSave(object sender, CanExecuteRoutedEventArgs e) => e.CanExecute = _pkg is not null;

    // ================================================================ cases à cocher

    /// <summary>Coche ou décoche une case de formulaire (annulable). Le paragraphe est désigné par sa position.</summary>
    internal void ToggleCheckBox(Paragraph p, Field field)
    {
        if (Body is not { } body || field.CheckBox is not { } box) return;
        var pos = DocPos.Of(body, p);
        using (var tx = _journal.Begin(box.Checked ? "Décocher la case" : "Cocher la case", body, pos.ContainerPath, pos.BlockIndex, 1, pos))
        {
            box.Checked = !box.Checked;
            tx.Commit(pos);
        }
        AfterEdit($"Case « {box.Name ?? "sans nom"} » {(box.Checked ? "cochée" : "décochée")}. Ctrl+Z pour annuler.");
    }

    void UndoRedo(bool undo)
    {
        var label = undo ? _journal.UndoLabel : _journal.RedoLabel;
        if (label is null) return;
        var caret = undo ? _journal.Undo() : _journal.Redo();
        AfterEdit($"{(undo ? "Annulé" : "Rétabli")} : {label}.", caret);
    }

    /// <summary>Après une modification hors frappe (case, annuler, rétablir) : remise en page incrémentale, point d'insertion.</summary>
    void AfterEdit(string status, DocPos? caret = null)
    {
        Relayout();
        _anchor = null;
        if (caret is not null && FromDocPos(caret) is { } t) _caret = t;
        else if (_caret is { } c && _layout?.LinesOf(c.Paragraph).Count == 0) _caret = DocumentStart();
        Scroller.UpdateLayout();
        UpdateCaret(scroll: caret is not null);
        UpdateModifiedMarker();
        StatusText.Text = status;
        EnsureAutoSaveTimer();
        CommandManager.InvalidateRequerySuggested();
    }

    // ================================================================ enregistrer

    /// <summary>Enregistre (ou « Enregistrer sous ») ; renvoie faux si l'utilisateur annule ou si l'écriture échoue.</summary>
    internal bool Save(bool saveAs)
    {
        if (_pkg is null || _path is null) return false;
        var target = _path;
        if (saveAs || _recoveredWithoutTarget)
        {
            // même famille de format : un .docm garde ses macros, un modèle reste un modèle
            var ext = Path.GetExtension(_path).ToLowerInvariant();
            var (label, pattern) = ext switch
            {
                ".docm" => ("Document Word prenant en charge les macros", "*.docm"),
                ".dotx" => ("Modèle Word", "*.dotx"),
                ".dotm" => ("Modèle Word prenant en charge les macros", "*.dotm"),
                _ => ("Document Word", "*.docx"),
            };
            var dlg = new SaveFileDialog
            {
                Filter = $"{label} ({pattern})|{pattern}", FileName = Path.GetFileName(_path), DefaultExt = ext,
                InitialDirectory = Path.GetDirectoryName(_path), Title = "Enregistrer sous", AddExtension = true,
            };
            if (dlg.ShowDialog(this) != true) return false;
            target = dlg.FileName;
        }
        else if (FileChangedOnDisk() && !SkipConfirm)
        {
            var r = MessageBox.Show(this, "Le fichier a été modifié par un autre programme depuis son ouverture.\n\n" +
                                          "Oui : le remplacer quand même\nNon : enregistrer sous un autre nom",
                                    "Enregistrer", MessageBoxButton.YesNoCancel, MessageBoxImage.Warning);
            if (r == MessageBoxResult.Cancel) return false;
            if (r == MessageBoxResult.No) return Save(saveAs: true);
        }
        try
        {
            Mouse.OverrideCursor = Cursors.Wait;
            _pkg.Save(target, keepBackup: AppSettings.Load().KeepBackup && File.Exists(target));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            Mouse.OverrideCursor = null;
            MessageBox.Show(this, $"Le document n'a pas pu être enregistré :\n\n{ex.Message}\n\nLe fichier d'origine n'a pas été modifié.",
                            "Enregistrer", MessageBoxButton.OK, MessageBoxImage.Error);
            return false;
        }
        finally { Mouse.OverrideCursor = null; }
        _path = Path.GetFullPath(target);
        _journal.MarkSaved();
        _recovered = false;
        _recoveredWithoutTarget = false;
        DiscardRecovery();
        RecordFileStamp();
        if (TrackRecent) RecentFiles.Add(_path);
        UpdateStatus();
        StatusText.Text = $"Enregistré à {DateTime.Now:HH:mm}" + (AppSettings.Load().KeepBackup ? " (version précédente : .bak)." : ".");
        return true;
    }

    /// <summary>Banc de saisie : enregistre le document modifié dans un autre fichier (le document ouvert reste intact).</summary>
    internal void SaveCopy(string target) => _pkg?.Save(target, keepBackup: false);

    /// <summary>Avant de fermer ou de remplacer le document : propose d'enregistrer les modifications. Faux = ne pas continuer.</summary>
    internal bool ConfirmDiscard()
    {
        if (!IsModified || SkipConfirm) return true;
        var r = MessageBox.Show(this, $"Voulez-vous enregistrer les modifications apportées à « {Path.GetFileName(_path)} » ?",
                                "OpenWrite", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
        if (r == MessageBoxResult.Cancel) return false;
        if (r == MessageBoxResult.Yes) return Save(saveAs: false);
        DiscardRecovery();
        return true;
    }

    void OnClosing(object? sender, CancelEventArgs e)
    {
        if (!ConfirmDiscard()) { e.Cancel = true; return; }
        _autoSaveTimer?.Stop();
    }
}
