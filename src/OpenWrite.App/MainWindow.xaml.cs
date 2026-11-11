using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using OpenWrite.Docx;

namespace OpenWrite.App;

/// <summary>
/// Fenêtre principale. Ruban aux onglets de Word (MainWindow.Ribbon.cs) ; document chargé dans le modèle en flux,
/// vue des pages mises en page par OpenWrite.Layout (MainWindow.Document.cs, depuis la 0.3) ; cases à cocher, annuler,
/// enregistrer sans perte (MainWindow.Editing.cs) ; copies de secours (MainWindow.Recovery.cs) ; infrastructure d'OpenPub
/// (mises à jour, documents récents, À propos, signalement des erreurs).
/// </summary>
public partial class MainWindow : Window
{
    /// <summary>Document ouvert (résumé lu sans mise en page) et son chemin.</summary>
    DocxInfo? _info;
    string? _path;
    double _zoom = 1;

    public MainWindow()
    {
        InitializeComponent();
        InitEditing();
        InitOutput();
        InitSpelling();
        InitReview();
        InitFormatPane();
        InitRuler();
        BuildRibbon();
        ShowWelcome();
        UpdateStatus();
        Loaded += async (_, _) =>
        {
            // vérification discrète des mises à jour, quelques secondes après le démarrage
            if (!Update.UpdateService.AutoCheck || Update.UpdateService.FeedUrl is null) return;
            await Task.Delay(3000);
            await CheckForUpdates(manual: false);
        };
    }

    /// <summary>Documents récents enregistrés (application normale ; jamais pendant les tests et captures).</summary>
    public bool TrackRecent { get; set; }

    /// <summary>Liste fictive pour les captures de la documentation.</summary>
    internal IReadOnlyList<string>? RecentOverride
    {
        get => _recentOverride;
        set { _recentOverride = value; if (_info is null) ShowWelcome(); }
    }
    IReadOnlyList<string>? _recentOverride;

    IReadOnlyList<string> Recent => RecentOverride ?? (TrackRecent ? RecentFiles.List() : Array.Empty<string>());

    // ================================================================ fichiers

    public const string OpenFilter = "Documents Word (*.docx;*.docm;*.dotx;*.dotm)|*.docx;*.docm;*.dotx;*.dotm|Tous les fichiers (*.*)|*.*";

    public void OpenFile(string path) => OpenFile(path, remember: true);

    /// <param name="remember">Ajouter aux documents récents (pas pour une copie de secours).</param>
    public void OpenFile(string path, bool remember)
    {
        if (!ConfirmDiscard()) return;
        var ext = Path.GetExtension(path).ToLowerInvariant();
        if (ext is ".doc" or ".rtf" or ".odt")
        {
            MessageBox.Show(this, $"Les fichiers {ext} ne sont pas encore pris en charge : OpenWrite ouvre les documents .docx.\n\n" +
                                  "RTF et ODT arrivent avec la version 1.5, l'ancien format .doc (via LibreOffice) avec la version 1.6. " +
                                  "En attendant, enregistrez le document au format .docx depuis Word ou LibreOffice.",
                            "OpenWrite", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        DocxInfo info;
        Mouse.OverrideCursor = Cursors.Wait;
        try { info = DocxInfo.Read(path); }
        catch (Exception ex) when (ex is DocxReadException or IOException or UnauthorizedAccessException)
        {
            Mouse.OverrideCursor = null;
            MessageBox.Show(this, $"Impossible d'ouvrir « {Path.GetFileName(path)} ».\n\n{ex.Message}", "OpenWrite", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        finally { Mouse.OverrideCursor = null; }
        if (!LoadPackage(path)) return;
        DiscardRecovery();
        _info = info;
        _path = Path.GetFullPath(path);
        RecordFileStamp();
        if (TrackRecent && remember) RecentFiles.Add(_path);
        ShowDocument();
        UpdateStatus();
        StatusText.Text = "Cliquez dans la page pour écrire" + (HasCheckBoxes ? ", sur une case pour la cocher" : "") + " ; Ctrl+S pour enregistrer. "
            + $"Mis en page en {LastLayoutMs:0} ms" + (_fonts.Substitutions.Count > 0 ? $" · polices remplacées : {string.Join(", ", _fonts.Substitutions.Select(kv => $"{kv.Key} → {kv.Value}"))}" : "") + ".";
    }

    void CloseDocument()
    {
        if (!ConfirmDiscard()) return;
        StopTracking();
        _reviewPkg = null;
        _comments = null;
        ShowCommentsPane(false);
        _info = null;
        _path = null;
        _pkg = null;
        _styles = null;
        _layout = null;
        _drawings = null;
        if (_hf is not null) { _hf = null; ShowHeaderTab(false); }
        _pagesView = null;
        _journal.Clear();
        _recovered = false;
        ShowWelcome();
        UpdateStatus();
        StatusText.Text = "";
    }

    void Open_Executed(object sender, ExecutedRoutedEventArgs e) => OpenDialog();

    void OpenDialog()
    {
        var dlg = new OpenFileDialog { Filter = OpenFilter, Title = "Ouvrir un document" };
        if (_path is not null) dlg.InitialDirectory = Path.GetDirectoryName(_path);
        if (dlg.ShowDialog(this) == true) OpenFile(dlg.FileName);
    }

    void OpenRecent(string file)
    {
        if (!File.Exists(file))
        {
            MessageBox.Show(this, $"Ce document n'existe plus (déplacé ou supprimé) :\n{file}", "OpenWrite", MessageBoxButton.OK, MessageBoxImage.Warning);
            RecentFiles.Remove(file);
            if (_info is null) ShowWelcome();
            return;
        }
        if (_path is not null && string.Equals(_path, file, StringComparison.OrdinalIgnoreCase)) return;
        OpenFile(file);
    }

    void CloseDoc_Executed(object sender, ExecutedRoutedEventArgs e) => CloseDocument();
    void HasDoc_CanExecute(object sender, CanExecuteRoutedEventArgs e) => e.CanExecute = _info is not null;
    void Help_Executed(object sender, ExecutedRoutedEventArgs e) => Help.HelpService.Open(CurrentHelpTopic(), this);

    string CurrentHelpTopic() => _info is null ? Help.HelpService.OpenDocument : TabHome.IsChecked == true ? Help.HelpService.HomeTab
        : _tableDesignTab?.IsChecked == true || _tableLayoutTab?.IsChecked == true ? Help.HelpService.Tables
        : _headerTab?.IsChecked == true ? Help.HelpService.HeadersFooters
        : _objectTab?.IsChecked == true || _objSel is not null ? Help.HelpService.Objects : Help.HelpService.Roadmap;

    void Window_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
        e.Handled = true;
    }

    void Window_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(DataFormats.FileDrop) is string[] { Length: > 0 } files) OpenFile(files[0]);
    }

    // ================================================================ titre et barre d'état

    void UpdateStatus()
    {
        if (_info is { } i && _path is not null)
        {
            // comme Word : « nom.docx [Mode de compatibilité] »
            var compat = i.CompatibilityMode < 15 ? " [Mode de compatibilité]" : "";
            Title = $"{(IsModified ? "• " : "")}{Path.GetFileName(_path)}{compat} — OpenWrite";
            DocTitle.Text = Path.GetFileName(_path) + (IsModified ? " · modifié" : "");
            DocTitle.ToolTip = _path;
            PageStatus.Text = i.SavedPages is { } p ? $"{p} page{(p > 1 ? "s" : "")}" : "";
            PageStatus.ToolTip = null;
            WordStatus.Text = $"{i.Words:N0} mot{(i.Words > 1 ? "s" : "")}".Replace(' ', ' ');
            LangStatus.Text = LanguageLabel(i.Languages);
        }
        else
        {
            Title = "OpenWrite";
            DocTitle.Text = "";
            DocTitle.ToolTip = null;
            PageStatus.Text = "Aucun document";
            PageStatus.ToolTip = null;
            WordStatus.Text = LangStatus.Text = "";
        }
        PaneDocInfo.Text = DocSummary();
        RefreshPanes();
        UpdatePageStatus();
        UpdateRibbonState();
    }

    /// <summary>Langue principale affichée dans la barre d'état (la première langue de rédaction trouvée).</summary>
    static string LanguageLabel(IReadOnlyList<string> langs)
    {
        if (langs.Count == 0) return "";
        static string Name(string tag)
        {
            try { var c = System.Globalization.CultureInfo.GetCultureInfo(tag); return c.TextInfo.ToTitleCase(c.NativeName); }
            catch (System.Globalization.CultureNotFoundException) { return tag; }
        }
        var main = langs.FirstOrDefault(l => l.StartsWith("fr", StringComparison.OrdinalIgnoreCase))
                   ?? langs.FirstOrDefault(l => l.StartsWith("nl", StringComparison.OrdinalIgnoreCase)) ?? langs[0];
        return Name(main) + (langs.Count > 1 ? $" (+{langs.Count - 1})" : "");
    }

    // ================================================================ zoom

    void SetZoom(double z)
    {
        _zoom = Math.Clamp(z, 0.5, 2);
        Workspace.LayoutTransform = new ScaleTransform(_zoom, _zoom);
        ZoomStatus.Text = $"{_zoom * 100:0} %";
        if (Math.Abs(ZoomSlider.Value - _zoom * 100) > 0.5) ZoomSlider.Value = _zoom * 100;
        Dispatcher.BeginInvoke(UpdateRuler, System.Windows.Threading.DispatcherPriority.Loaded);
    }

    void ZoomSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        if (Workspace is not null) SetZoom(e.NewValue / 100);
    }

    void ZoomIn_Click(object sender, RoutedEventArgs e) => SetZoom(Math.Round(_zoom * 10 + 1) / 10);
    void ZoomOut_Click(object sender, RoutedEventArgs e) => SetZoom(Math.Round(_zoom * 10 - 1) / 10);

    void Scroller_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        if ((Keyboard.Modifiers & ModifierKeys.Control) == 0) return;
        SetZoom(_zoom * (e.Delta > 0 ? 1.1 : 1 / 1.1));
        e.Handled = true;
    }

    // ================================================================ volets

    void Pane_Checked(object sender, RoutedEventArgs e)
    {
        if (ThumbsPanel is null) return;      // pendant InitializeComponent
        ThumbsPanel.Visibility = PaneThumbs.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
        OutlinePanel.Visibility = PaneOutline.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
    }

    void CloseFormatPane_Click(object sender, RoutedEventArgs e) => ShowFormatPane(false);

    internal void ShowFormatPane(bool show)
    {
        FormatPane.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
        FormatColumn.Width = show ? new GridLength(268) : new GridLength(0);
        if (FormatPaneToggle is not null) FormatPaneToggle.IsChecked = show;
    }

    // ================================================================ mises à jour (repris d'OpenPub)

    Core.Update.UpdateManifest? _pendingUpdate;

    async Task CheckForUpdates(bool manual)
    {
        if (manual) StatusText.Text = "Recherche des mises à jour…";
        var r = await Update.UpdateService.CheckAsync(honorSkip: !manual);
        if (r.IsNewer && r.Manifest is { } m)
        {
            _pendingUpdate = m;
            UpdateText.Text = $"OpenWrite {m.Version} est disponible (vous avez la {Update.UpdateService.CurrentVersion}).";
            UpdateNotes.Text = m.Notes ?? "";
            UpdateNotes.ToolTip = m.Notes;
            UpdateBanner.Visibility = Visibility.Visible;
            UpdateNowButton.IsEnabled = UpdateLaterButton.IsEnabled = UpdateSkipButton.IsEnabled = true;
            UpdateProgress.Visibility = Visibility.Collapsed;
        }
        if (manual)
        {
            StatusText.Text = r.Message;
            if (!r.IsNewer) MessageBox.Show(this, r.Message, "Mises à jour", MessageBoxButton.OK,
                r.Manifest is null ? MessageBoxImage.Warning : MessageBoxImage.Information);
        }
    }

    async void UpdateNow_Click(object sender, RoutedEventArgs e)
    {
        if (_pendingUpdate is not { } m) return;
        UpdateNowButton.IsEnabled = UpdateLaterButton.IsEnabled = UpdateSkipButton.IsEnabled = false;
        UpdateProgress.Visibility = Visibility.Visible;
        UpdateText.Text = $"Téléchargement d'OpenWrite {m.Version}…";
        try
        {
            var file = await Update.UpdateService.DownloadAsync(m, new Progress<double>(v => UpdateProgress.Value = v));
            UpdateText.Text = "Installation… OpenWrite va redémarrer.";
            Update.UpdateService.InstallAndRestart(file, _path);
            await Task.Delay(500);
            Close();
        }
        catch (Exception ex)
        {
            UpdateText.Text = "La mise à jour a échoué.";
            UpdateNotes.Text = ex.Message;
            UpdateProgress.Visibility = Visibility.Collapsed;
            UpdateNowButton.IsEnabled = UpdateLaterButton.IsEnabled = UpdateSkipButton.IsEnabled = true;
        }
    }

    void UpdateLater_Click(object sender, RoutedEventArgs e) => UpdateBanner.Visibility = Visibility.Collapsed;

    void UpdateSkip_Click(object sender, RoutedEventArgs e)
    {
        if (_pendingUpdate is { } m) Update.UpdateService.SkipVersion(m.Version);
        UpdateBanner.Visibility = Visibility.Collapsed;
    }

    void ShowAbout() => new AboutWindow(() => _ = CheckForUpdates(manual: true)) { Owner = this }.ShowDialog();
}
