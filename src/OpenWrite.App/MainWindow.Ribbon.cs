using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;

namespace OpenWrite.App;

// Ruban aux onglets de Word (mêmes noms, même ordre, mêmes groupes) et menu Fichier.
// Version 0.1 : les groupes sont en place mais inactifs ; chacun indique la version qui l'activera
// (PLAN.md §8). Les commandes déjà disponibles (zoom, volet Mise en forme, aide) sont actives.
public partial class MainWindow
{
    /// <summary>Groupe du ruban : titre, version qui l'active, boutons (glyphe Segoe Fluent Icons, libellé).</summary>
    sealed record RibbonGroup(string Header, string Version, params (string Glyph, string Label)[] Buttons);

    static readonly (string Tab, RibbonGroup[] Groups)[] Roadmap =
    {
        ("TabHome", new RibbonGroup[]
        {
            new("Presse-papiers", "0.4", ("", "Coller"), ("", "Couper"), ("", "Copier")),
            new("Police", "0.4", ("", "Police"), ("", "Gras"), ("", "Italique"), ("", "Souligné")),
            new("Paragraphe", "0.4", ("", "Gauche"), ("", "Centrer"), ("", "Droite"), ("\uE8FD", "Puces"), ("\uF146", "Numéros"), ("\uE9D5", "Titres 1.1"), ("\uE890", "Marques ¶")),
            new("Styles", "0.5", ("", "Styles"), ("", "Modifier")),
            new("Modification", "0.4", ("", "Rechercher"), ("", "Remplacer"), ("", "Sélectionner")),
        }),
        ("TabInsert", new RibbonGroup[]
        {
            new("Pages", "0.6", ("", "Page de garde"), ("", "Saut de page")),
            new("Tableaux", "0.7", ("", "Tableau")),
            new("Illustrations", "1.0", ("", "Images"), ("", "Formes")),
            new("Liens", "0.4", ("", "Lien"), ("", "Signet")),
            new("Commentaires", "1.1", ("", "Commentaire")),
            new("En-tête et pied", "0.6", ("", "En-tête"), ("", "Pied de page"), ("", "N° de page")),
            new("Texte", "1.0", ("", "Zone de texte"), ("", "Lettrine")),
            new("Symboles", "0.4", ("", "Symbole")),
        }),
        ("TabDesign", new RibbonGroup[]
        {
            new("Mise en forme du document", "0.5", ("", "Thèmes"), ("", "Polices"), ("\uE790", "Couleurs")),
            new("Arrière-plan de la page", "0.6", ("", "Filigrane"), ("", "Couleur"), ("", "Bordures")),
        }),
        ("TabLayout", new RibbonGroup[]
        {
            new("Mise en page", "0.6", ("", "Marges"), ("", "Orientation"), ("", "Taille"), ("", "Colonnes"), ("", "Sauts")),
            new("Paragraphe", "0.4", ("", "Retraits"), ("", "Espacement")),
            new("Organiser", "1.0", ("", "Position"), ("", "Habillage")),
        }),
        ("TabReferences", new RibbonGroup[]
        {
            new("Table des matières", "1.2", ("", "Table"), ("", "Mettre à jour")),
            new("Notes de bas de page", "1.2", ("", "Note"), ("", "Note de fin")),
            new("Citations et bibliographie", "1.2", ("", "Citation"), ("", "Bibliographie")),
            new("Légendes", "1.2", ("", "Légende"), ("", "Table des illustrations")),
            new("Index", "1.2", ("", "Entrée"), ("", "Index")),
        }),
        ("TabMailings", new RibbonGroup[]
        {
            new("Démarrer la fusion", "1.4", ("", "Démarrer"), ("", "Destinataires")),
            new("Champs d'écriture et d'insertion", "1.4", ("", "Champ de fusion"), ("", "Bloc d'adresse")),
            new("Aperçu et terminer", "1.4", ("", "Aperçu"), ("", "Terminer")),
            new("Créer", "1.4", ("", "Enveloppes"), ("", "Étiquettes")),
        }),
        ("TabReview", new RibbonGroup[]
        {
            new("Vérification", "1.1", ("", "Orthographe"), ("", "Synonymes"), ("", "Statistiques")),
            new("Commentaires", "1.1", ("", "Nouveau"), ("", "Afficher")),
            new("Suivi", "1.1", ("", "Suivi des modifications")),
            new("Modifications", "1.1", ("", "Accepter"), ("", "Refuser")),
            new("Comparer", "1.1", ("", "Comparer")),
            new("Protéger", "1.1", ("", "Restreindre")),
        }),
        ("TabView", new RibbonGroup[]
        {
            new("Affichage", "1.7", ("", "Page"), ("", "Lecture"), ("", "Web"), ("", "Plan")),
            new("Fenêtre", "1.7", ("", "Nouvelle fenêtre"), ("", "Fractionner")),
        }),
    };

    readonly Dictionary<string, StackPanel> _panels = new();

    /// <summary>Bouton « Volet Mise en forme » de l'onglet Affichage (actif dès la 0.1).</summary>
    ToggleButton? FormatPaneToggle;

    /// <summary>Boutons déjà actifs (version 0.4) : actifs dès qu'un document est affiché.</summary>
    readonly List<Button> _liveButtons = new();

    /// <summary>Commandes de l'onglet Accueil déjà disponibles, par libellé de bouton.</summary>
    Dictionary<string, (Action Run, string Tip)> HomeCommands() => new()
    {
        ["Coller"] = (() => Paste(), "Coller (Ctrl+V) ; texte seul : Ctrl+Maj+V"),
        ["Police"] = (ShowFontDialog, "Police… (Ctrl+D)"),
        ["Couper"] = (Cut, "Couper (Ctrl+X)"),
        ["Copier"] = (Copy, "Copier (Ctrl+C)"),
        ["Gras"] = (() => ToggleCharacter("b"), "Gras (Ctrl+G ou Ctrl+B)"),
        ["Italique"] = (() => ToggleCharacter("i"), "Italique (Ctrl+I)"),
        ["Souligné"] = (() => ToggleCharacter("u"), "Souligné (Ctrl+U)"),
        ["Gauche"] = (() => SetAlignment("left"), "Aligner à gauche (Ctrl+Maj+G ou Ctrl+L)"),
        ["Centrer"] = (() => SetAlignment("center"), "Centrer (Ctrl+E)"),
        ["Droite"] = (() => SetAlignment("right"), "Aligner à droite (Ctrl+Maj+D ou Ctrl+R)"),
        ["Sélectionner"] = (SelectAll, "Sélectionner tout (Ctrl+A)"),
        ["Rechercher"] = (() => ShowFind(replace: false), "Rechercher (Ctrl+F)"),
        ["Remplacer"] = (() => ShowFind(replace: true), "Remplacer (Ctrl+H)"),
        ["Marques ¶"] = (ToggleMarks, "Afficher ou masquer les marques de mise en forme (Ctrl+Maj+8)"),
        ["Puces"] = (() => ToggleList(Core.Styles.ListKind.Bullet), "Liste à puces (Ctrl+Maj+L) ; Tab / Maj+Tab au début d'un élément : niveau"),
        ["Numéros"] = (() => ToggleList(Core.Styles.ListKind.Numbered), "Liste numérotée ; clic droit sur un élément : recommencer à 1, continuer"),
        ["Titres 1.1"] = (NumberHeadings, "Liste à plusieurs niveaux liée aux titres : 1. / 1.1 / 1.1.1"),
        ["Styles"] = (ShowStylesWindow, "Tous les styles : appliquer, modifier, créer"),
        ["Modifier"] = (ModifyCurrentStyle, "Modifier le style du paragraphe"),
    };

    /// <summary>Bouton du ruban qui vient d'être cliqué (pour ouvrir un menu sous lui).</summary>
    FrameworkElement? _clickedButton;

    /// <summary>Commandes de l'onglet Création déjà disponibles (menus sous le bouton).</summary>
    Dictionary<string, (Action Run, string Tip)> DesignCommands() => new()
    {
        ["Thèmes"] = (() => ShowThemesMenu(_clickedButton!), "Thèmes : polices et couleurs du document ensemble"),
        ["Polices"] = (() => ShowThemeFontsMenu(_clickedButton!), "Polices du thème : titres et corps du texte"),
        ["Couleurs"] = (() => ShowThemeColorsMenu(_clickedButton!), "Couleurs du thème : titres, accents, liens"),
    };

    /// <summary>Commandes de l'onglet Insertion déjà disponibles.</summary>
    Dictionary<string, (Action Run, string Tip)> InsertCommands() => new()
    {
        ["Lien"] = (ShowLinkDialog, "Lien hypertexte… (Ctrl+K)"),
        ["Symbole"] = (ShowSymbols, "Symbole… : caractères spéciaux, espaces et tirets insécables"),
        ["Saut de page"] = (() => InsertSpecial(Core.Editing.SpecialChar.PageBreak), "Saut de page (Ctrl+Entrée)"),
    };

    /// <summary>Commandes de l'onglet Mise en page déjà disponibles.</summary>
    Dictionary<string, (Action Run, string Tip)> LayoutCommands() => new()
    {
        ["Retraits"] = (ShowParagraphDialog, "Paragraphe… : retraits"),
        ["Marges"] = (() => ShowMarginsMenu(_clickedButton!), "Marges de la section"),
        ["Orientation"] = (() => ShowOrientationMenu(_clickedButton!), "Portrait ou paysage"),
        ["Taille"] = (() => ShowSizeMenu(_clickedButton!), "Format du papier"),
        ["Colonnes"] = (() => ShowColumnsMenu(_clickedButton!), "Colonnes de la section"),
        ["Sauts"] = (() => ShowBreaksMenu(_clickedButton!), "Sauts de page, de colonne et de section"),
        ["Espacement"] = (ShowParagraphDialog, "Paragraphe… : espacement et interligne"),
    };

    /// <summary>Active les commandes disponibles quand un document est affiché.</summary>
    void UpdateRibbonState()
    {
        foreach (var b in _liveButtons) b.IsEnabled = _layout is not null;
        UpdateStyleGallery();
    }

    StackPanel? _styleGallery;

    /// <summary>Galerie : six styles rapides avec aperçu (police, taille bornée, gras, italique, couleur du style).</summary>
    void UpdateStyleGallery()
    {
        if (_styleGallery is null) return;
        _styleGallery.Children.Clear();
        if (_layout is null || _styles is null) return;
        foreach (var (key, display, type) in StyleChoices().Where(c => c.Type == "paragraph").Take(5))
        {
            var preview = new TextBlock { Text = display, TextTrimming = TextTrimming.CharacterEllipsis, MaxWidth = 68 };
            if (_styles.Sheet.Style(key) is { } st)
            {
                var probe = new Core.Model.Paragraph { Props = new Core.Model.PropertyBag("pPr", _styles.Sheet.W) };
                probe.Props.SetVal("pStyle", key);
                var rf = _styles.ParagraphMark(probe);
                if (rf.FontAscii is { } font) preview.FontFamily = new FontFamily(font);
                preview.FontSize = Math.Clamp(rf.Size * 96 / 72 * 0.8, 10, 15);
                if (rf.Bold) preview.FontWeight = FontWeights.Bold;
                if (rf.Italic) preview.FontStyle = FontStyles.Italic;
                if (rf.Color is { Length: 6 } col) preview.Foreground = new SolidColorBrush(Color.FromRgb(Convert.ToByte(col[..2], 16), Convert.ToByte(col[2..4], 16), Convert.ToByte(col[4..], 16)));
            }
            var b = new Button
            {
                Content = preview, Width = 76, Height = 50, Margin = new Thickness(2, 0, 2, 0), Focusable = false, ToolTip = $"Style {display}",
                Background = Brushes.White, BorderBrush = (Brush)FindResource("Line"), Padding = new Thickness(4, 2, 4, 2),
            };
            b.Click += (_, _) => { ApplyStyle(key); _pagesView?.Focus(); };
            _styleGallery.Children.Add(b);
        }
    }

    void BuildRibbon()
    {
        var home = HomeCommands();
        var layoutCmds = LayoutCommands();
        var insertCmds = InsertCommands();
        var designCmds = DesignCommands();
        foreach (var (tab, groups) in Roadmap)
        {
            var panel = new StackPanel { Orientation = Orientation.Horizontal, Visibility = tab == "TabHome" ? Visibility.Visible : Visibility.Collapsed };
            foreach (var g in groups)
            {
                var buttons = new StackPanel { Orientation = Orientation.Horizontal };
                foreach (var (glyph, label) in g.Buttons)
                {
                    var button = new Button { Style = (Style)FindResource("RBig"), Tag = glyph, Content = label, IsEnabled = false };
                    if ((tab switch { "TabHome" => home, "TabLayout" => layoutCmds, "TabInsert" => insertCmds, "TabDesign" => designCmds, _ => null })?.TryGetValue(label, out var cmd) == true)
                    {
                        // le point d'insertion garde le clavier : le bouton ne prend pas le focus
                        button.Focusable = false;
                        button.ToolTip = cmd.Tip;
                        // menu sous le bouton : il garde le clavier
                        bool menu = tab == "TabDesign" || label is "Marges" or "Orientation" or "Taille" or "Colonnes" or "Sauts";
                        button.Click += (_, _) => { _clickedButton = button; cmd.Run(); if (!menu) _pagesView?.Focus(); };
                        _liveButtons.Add(button);
                    }
                    buttons.Children.Add(button);
                }
                if (tab == "TabHome" && g.Header == "Styles")
                {
                    // galerie des styles rapides, remplie à l'ouverture d'un document
                    _styleGallery = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 6, 0) };
                    buttons.Children.Insert(0, _styleGallery);
                }
                panel.Children.Add(new HeaderedContentControl
                {
                    Style = (Style)FindResource("RGroup"), Header = $"{g.Header} · {g.Version}", Content = buttons,
                    ToolTip = $"{g.Header} : disponible à partir de la version {g.Version}",
                });
                ToolTipService.SetShowOnDisabled(panel.Children[^1], true);
            }
            if (tab == "TabView") AddViewCommands(panel);
            _panels[tab] = panel;
            RibbonPanels.Children.Add(panel);
        }
    }

    /// <summary>Commandes déjà actives de l'onglet Affichage : zoom et volets.</summary>
    void AddViewCommands(StackPanel panel)
    {
        var zoom = new StackPanel { Orientation = Orientation.Horizontal };
        Button Big(string glyph, string label, string tip, Action act)
        {
            var b = new Button { Style = (Style)FindResource("RBig"), Tag = glyph, Content = label, ToolTip = tip };
            b.Click += (_, _) => act();
            return b;
        }
        zoom.Children.Add(Big("", "Zoom", "Zoom avant (Ctrl + molette)", () => SetZoom(Math.Round(_zoom * 10 + 1) / 10)));
        zoom.Children.Add(Big("", "100 %", "Revenir à 100 %", () => SetZoom(1)));
        panel.Children.Insert(1, new HeaderedContentControl { Style = (Style)FindResource("RGroup"), Header = "Zoom", Content = zoom });

        var panes = new StackPanel { Orientation = Orientation.Horizontal };
        FormatPaneToggle = new ToggleButton { Style = (Style)FindResource("RBig"), Tag = "", Content = "Mise en forme", IsChecked = true, ToolTip = "Afficher ou masquer le volet Mise en forme" };
        FormatPaneToggle.Click += (_, _) => ShowFormatPane(FormatPaneToggle.IsChecked == true);
        panes.Children.Add(FormatPaneToggle);
        panes.Children.Add(Big("", "Navigation", "Plan des titres du document", () => { PaneOutline.IsChecked = true; }));
        _rulerToggle = new ToggleButton { Style = (Style)FindResource("RBig"), Tag = "\uECC6", Content = "Règle", IsChecked = true, ToolTip = "Afficher ou masquer la règle" };
        _rulerToggle.Click += (_, _) => ToggleRuler();
        panes.Children.Add(_rulerToggle);
        panel.Children.Insert(2, new HeaderedContentControl { Style = (Style)FindResource("RGroup"), Header = "Afficher", Content = panes });
    }

    void Tab_Checked(object sender, RoutedEventArgs e)
    {
        if (_panels.Count == 0 || sender is not RadioButton tab) return;      // pendant InitializeComponent
        foreach (var (name, panel) in _panels)
            panel.Visibility = name == tab.Name ? Visibility.Visible : Visibility.Collapsed;
        RibbonScroller.ScrollToHorizontalOffset(0);
        tab.BringIntoView();
    }

    // Petits écrans : le ruban et la ligne d'onglets défilent horizontalement (repris d'OpenPub).
    void RibbonScroller_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
    {
        if (sender is not ScrollViewer sv || sv.ScrollableWidth <= 0) return;
        sv.ScrollToHorizontalOffset(sv.HorizontalOffset - e.Delta / 2.0);
        e.Handled = true;
    }

    void RibbonScroller_ScrollChanged(object sender, ScrollChangedEventArgs e)
    {
        var sv = RibbonScroller;
        RibbonLeft.Visibility = sv.HorizontalOffset > 0.5 ? Visibility.Visible : Visibility.Collapsed;
        RibbonRight.Visibility = sv.HorizontalOffset < sv.ScrollableWidth - 0.5 ? Visibility.Visible : Visibility.Collapsed;
    }

    void RibbonLeft_Click(object sender, RoutedEventArgs e) => RibbonScroller.ScrollToHorizontalOffset(RibbonScroller.HorizontalOffset - 60);
    void RibbonRight_Click(object sender, RoutedEventArgs e) => RibbonScroller.ScrollToHorizontalOffset(RibbonScroller.HorizontalOffset + 60);

    // ================================================================ menu Fichier

    /// <summary>Dernier menu ouvert (captures de la documentation).</summary>
    internal ContextMenu? LastMenu { get; private set; }

    void FileButton_Click(object sender, RoutedEventArgs e)
    {
        var menu = new ContextMenu { PlacementTarget = FileButton, Placement = PlacementMode.Bottom };
        LastMenu = menu;
        MenuItem Add(string header, string gesture, Action act, bool enabled = true, string? glyph = null)
        {
            var it = new MenuItem { Header = header, InputGestureText = gesture, IsEnabled = enabled };
            if (glyph is not null)
                it.Icon = new TextBlock { Text = glyph, FontFamily = (FontFamily)FindResource("Icons"), FontSize = 15, Foreground = (Brush)FindResource("IconBrush") };
            it.Click += (_, _) => act();
            menu.Items.Add(it);
            return it;
        }
        bool has = _info is not null;
        Add("Nouveau document", "Ctrl+N", () => { }, false, "").ToolTip = "Version 0.4";
        Add("Ouvrir…", "Ctrl+O", OpenDialog, glyph: "");
        AddRecent(menu);
        Add("Enregistrer", "Ctrl+S", () => Save(saveAs: false), has, "\uE74E");
        Add("Enregistrer sous…", "F12", () => Save(saveAs: true), has, "\uE792");
        Add("Informations sur le document", "", ShowDocInfo, has, "");
        Add("Fermer le document", "Ctrl+W", CloseDocument, has, "");
        menu.Items.Add(new Separator());
        Add("Aide en ligne", "F1", () => Help.HelpService.Open(CurrentHelpTopic(), this), glyph: "");
        Add("Feuille de route des versions", "", () => Help.HelpService.Open(Help.HelpService.Roadmap, this), glyph: "");
        menu.Items.Add(new Separator());
        Add("Rechercher des mises à jour…", "", () => _ = CheckForUpdates(manual: true), glyph: "");
        var auto = new MenuItem { Header = "Rechercher automatiquement au démarrage", IsCheckable = true, IsChecked = Update.UpdateService.AutoCheck };
        auto.Click += (_, _) => Update.UpdateService.AutoCheck = auto.IsChecked;
        menu.Items.Add(auto);
        var prefs = AppSettings.Load();
        var autosave = new MenuItem { Header = "Copie de secours automatique (toutes les 5 min)", IsCheckable = true, IsChecked = prefs.AutoSave };
        autosave.Click += (_, _) => AppSettings.Update(x => x.AutoSave = autosave.IsChecked);
        menu.Items.Add(autosave);
        var backup = new MenuItem { Header = "Garder la version précédente (.bak) à l'enregistrement", IsCheckable = true, IsChecked = prefs.KeepBackup };
        backup.Click += (_, _) => AppSettings.Update(x => x.KeepBackup = backup.IsChecked);
        menu.Items.Add(backup);
        var crash = new MenuItem { Header = "Signaler automatiquement les erreurs (rapport anonyme)", IsCheckable = true, IsChecked = AppSettings.Load().CrashReports };
        crash.Click += (_, _) => AppSettings.Update(x => x.CrashReports = crash.IsChecked);
        menu.Items.Add(crash);
        Add("À propos d'OpenWrite", "", ShowAbout, glyph: "");
        menu.Items.Add(new Separator());
        Add("Quitter", "Alt+F4", Close, glyph: "");
        menu.IsOpen = true;
    }

    /// <summary>Section « Documents récents » du menu Fichier (reprise d'OpenPub).</summary>
    void AddRecent(ContextMenu menu)
    {
        var recent = Recent;
        if (recent.Count == 0) return;
        menu.Items.Add(new Separator());
        menu.Items.Add(new MenuItem
        {
            Header = new TextBlock { Text = "DOCUMENTS RÉCENTS", FontSize = 10.5, FontWeight = FontWeights.Bold, Foreground = (Brush)FindResource("Muted") },
            IsHitTestVisible = false, Focusable = false, Padding = new Thickness(10, 4, 10, 2),
        });
        int n = 0;
        foreach (var file in recent)
        {
            n++;
            bool isCurrent = _path is not null && string.Equals(_path, file, StringComparison.OrdinalIgnoreCase);
            var head = new StackPanel();
            head.Children.Add(new TextBlock { Text = Path.GetFileName(file), FontWeight = isCurrent ? FontWeights.Bold : FontWeights.Medium, MaxWidth = 340, TextTrimming = TextTrimming.CharacterEllipsis });
            head.Children.Add(new TextBlock { Text = Path.GetDirectoryName(file), FontSize = 11, Foreground = (Brush)FindResource("Muted"), MaxWidth = 340, TextTrimming = TextTrimming.CharacterEllipsis, Margin = new Thickness(0, 1, 0, 0) });
            var it = new MenuItem
            {
                Header = head, ToolTip = file, Padding = new Thickness(10, 5, 10, 5), InputGestureText = isCurrent ? "ouvert" : "",
                Icon = new Border
                {
                    Width = 20, Height = 20, CornerRadius = new CornerRadius(6), Background = (Brush)FindResource("SoftBadge"),
                    Child = new TextBlock
                    {
                        Text = n.ToString(CultureInfo.InvariantCulture), FontSize = 11, FontWeight = FontWeights.Bold,
                        Foreground = (Brush)FindResource("SelectedText"), HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center,
                    },
                },
            };
            it.Click += (_, _) => OpenRecent(file);
            menu.Items.Add(it);
        }
        var clear = new MenuItem { Header = "Effacer la liste", Foreground = (Brush)FindResource("Muted"), FontSize = 12, Padding = new Thickness(10, 5, 10, 5) };
        clear.Click += (_, _) => { RecentFiles.Clear(); StatusText.Text = "Liste des documents récents effacée."; if (_info is null) ShowWelcome(); };
        menu.Items.Add(clear);
        menu.Items.Add(new Separator());
    }
}
