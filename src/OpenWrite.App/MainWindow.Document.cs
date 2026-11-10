using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using OpenWrite.Core.Model;
using OpenWrite.Docx;
using OpenWrite.Layout;

namespace OpenWrite.App;

// Zone de travail : accueil (ouvrir, documents récents, feuille de route) et, depuis la version 0.3, vue des pages
// du document mis en page par OpenWrite.Layout (pages dessinées à la demande, miniatures, page x sur y).
public partial class MainWindow
{
    Brush R(string key) => (Brush)FindResource(key);

    /// <summary>Étapes affichées sur l'accueil (PLAN.md §8.17).</summary>
    static readonly (string Version, string Text)[] Milestones =
    {
        ("1.0", "Images, formes et zones de texte"),
    };

    static readonly DropShadowEffect CardShadow = new() { Color = Color.FromRgb(0x27, 0x4D, 0x9D), BlurRadius = 30, ShadowDepth = 8, Direction = 270, Opacity = 0.14 };

    void ShowWelcome()
    {
        Workspace.Children.Clear();
        // l'accueil suit la largeur de la fenêtre ; seule la page (largeur fixe) peut défiler horizontalement
        Scroller.HorizontalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Disabled;
        var card = new Border
        {
            MaxWidth = 640, Background = R("Surface"), BorderBrush = R("Line"), BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(22), Padding = new Thickness(40, 34, 40, 30), Effect = CardShadow, Margin = new Thickness(0, 24, 0, 0),
        };
        var sp = new StackPanel();
        var brand = new StackPanel { Orientation = Orientation.Horizontal };
        var logo = new Image { Source = new BitmapImage(new Uri("pack://application:,,,/Assets/openwrite-128.png")), Width = 84, Height = 84 };
        RenderOptions.SetBitmapScalingMode(logo, BitmapScalingMode.HighQuality);
        brand.Children.Add(logo);
        var names = new StackPanel { VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(16, 0, 0, 0) };
        names.Children.Add(new TextBlock { Text = "OpenWrite", FontFamily = (FontFamily)FindResource("TitleFont"), FontWeight = FontWeights.ExtraBold, FontSize = 38 });
        names.Children.Add(new Border
        {
            Style = (Style)FindResource("SoonBadge"), HorizontalAlignment = HorizontalAlignment.Left, Margin = new Thickness(2, 4, 0, 0),
            Child = new TextBlock { Text = $"Version {Update.UpdateService.CurrentVersion} · aperçu", FontWeight = FontWeights.Bold, FontSize = 12.5, Foreground = R("VersionText") },
        });
        brand.Children.Add(names);
        sp.Children.Add(brand);

        sp.Children.Add(new TextBlock
        {
            Text = "Traitement de texte qui ouvre et enregistre directement les documents Word. Cette version affiche un .docx " +
                   "page par page comme dans Word ; on y écrit, met en forme, copie-colle avec Word et le web, et l'enregistrement ne perd rien du reste du document.",
            TextWrapping = TextWrapping.Wrap, FontSize = 14.5, LineHeight = 22, Foreground = R("Muted"), Margin = new Thickness(0, 20, 0, 18),
        });
        var open = new Button { Style = (Style)FindResource("PillButton"), HorizontalAlignment = HorizontalAlignment.Left, Padding = new Thickness(22, 11, 24, 11) };
        var openContent = new StackPanel { Orientation = Orientation.Horizontal };
        openContent.Children.Add(new TextBlock { Text = "", FontFamily = (FontFamily)FindResource("Icons"), FontSize = 16, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 10, 0) });
        openContent.Children.Add(new TextBlock { Text = "Ouvrir un document Word…", FontSize = 15 });
        open.Content = openContent;
        open.ToolTip = "Ctrl+O — ou faites glisser un fichier .docx dans la fenêtre";
        open.Click += (_, _) => OpenDialog();
        sp.Children.Add(open);

        if (Recent.Count > 0)
        {
            sp.Children.Add(new TextBlock { Style = (Style)FindResource("PanelTitle"), Text = "DOCUMENTS RÉCENTS", Margin = new Thickness(0, 24, 0, 6) });
            foreach (var file in Recent.Take(5))
            {
                var row = new StackPanel();
                row.Children.Add(new TextBlock { Text = Path.GetFileName(file), FontWeight = FontWeights.SemiBold, TextTrimming = TextTrimming.CharacterEllipsis });
                row.Children.Add(new TextBlock { Text = Path.GetDirectoryName(file), FontSize = 11.5, Foreground = R("Muted"), TextTrimming = TextTrimming.CharacterEllipsis });
                var b = new Button { Style = (Style)FindResource("SideButton"), Content = row, ToolTip = file, Padding = new Thickness(12, 6, 12, 6), HorizontalAlignment = HorizontalAlignment.Stretch };
                b.Click += (_, _) => OpenRecent(file);
                sp.Children.Add(b);
            }
        }

        sp.Children.Add(new TextBlock { Style = (Style)FindResource("PanelTitle"), Text = "PROCHAINES VERSIONS", Margin = new Thickness(0, 24, 0, 8) });
        var steps = new WrapPanel();
        foreach (var (v, text) in Milestones)
        {
            var chip = new StackPanel { Orientation = Orientation.Horizontal };
            chip.Children.Add(new TextBlock { Text = v, FontWeight = FontWeights.Bold, Foreground = R("SelectedText"), Margin = new Thickness(0, 0, 6, 0) });
            chip.Children.Add(new TextBlock { Text = text, Foreground = R("Ink") });
            steps.Children.Add(new Border
            {
                Background = R("Chrome"), BorderBrush = R("Line"), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(10),
                Padding = new Thickness(10, 5, 10, 6), Margin = new Thickness(0, 0, 6, 6), Child = chip,
            });
        }
        sp.Children.Add(steps);
        card.Child = sp;
        Workspace.Children.Add(card);
    }

    /// <summary>Le document a-t-il des cases à cocher de formulaire (message d'accueil de la barre d'état) ?</summary>
    bool HasCheckBoxes => Body is { } b && FieldIndex.Build(b).All.Any(f => f.CheckBox is not null);

    // ================================================================ vue des pages (version 0.3)

    /// <summary>Champs par caractère de début (clic sur une case) ; recalculé à la demande après une modification.</summary>
    Dictionary<FieldChar, Field>? _fieldsByBegin;
    readonly Layout.WpfFontProvider _fonts = new();
    DocumentLayout? _layout;
    Layout.PageDrawings? _drawings;
    Layout.PagesView? _pagesView, _thumbsView;
    readonly Dictionary<string, ImageSource?> _images = new();

    const double PtToDip = 96.0 / 72;

    /// <summary>
    /// Vue des pages : le document mis en page par le moteur du lot 2 (PLAN.md §10.7, §10.8), pages dessinées à la
    /// demande (seules les pages visibles), miniatures dans le volet gauche, page affichée dans la barre d'état.
    /// Les cases à cocher se cochent d'un clic ; le volet de droite décrit le paragraphe survolé.
    /// </summary>
    void ShowDocument()
    {
        if (_info is null || _path is null || Body is not { } body || _styles is not { } styles || _pkg is not { } pkg) return;
        Workspace.Children.Clear();
        Scroller.HorizontalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto;
        _fieldsByBegin = null;
        _fonts.AltNames = pkg.FontAltNames();
        _images.Clear();
        var sw = System.Diagnostics.Stopwatch.StartNew();
        _layout = new DocumentLayout(body, styles, _fonts, pkg.RelatedStory, id => pkg.NoteStory("footnote", id));
        _layout.LayoutAll();
        LastLayoutMs = sw.Elapsed.TotalMilliseconds;
        _drawings = new Layout.PageDrawings(_layout, new Layout.PageRenderer(ImageOf) { ShowMarks = _showMarks, HeaderMode = _hf is not null });
        RefreshHeaderLines();
        _pagesView = new Layout.PagesView(_drawings, PtToDip, 24);
        _pagesView.MouseMove += PagesView_MouseMove;
        _pagesView.MouseLeftButtonUp += PagesView_MouseLeftButtonUp;
        Workspace.Children.Add(_pagesView);
        InitTyping(_pagesView);
        BuildThumbnails();
        UpdatePageBackground();
        UpdatePageStatus();
    }

    /// <summary>Durée de la dernière mise en page complète (millisecondes), pour la barre d'état et les tests.</summary>
    internal double LastLayoutMs { get; private set; }

    internal DocumentLayout? CurrentLayout => _layout;

    /// <summary>
    /// Image d'une relation r:embed (DrawingML) ou r:id (VML) : partie image du corps du document, ou de l'en-tête /
    /// du pied <paramref name="partId"/> (r:id de la partie dans le document).
    /// </summary>
    ImageSource? ImageOf(string? partId, string rid)
    {
        if (_pkg?.Main is not { } main) return null;
        var partName = main.PartName;
        if (partId is not null)
        {
            if (!_pkg.Relationships(main.PartName).TryGetValue(partId, out var hrel) || hrel.External) return null;
            partName = hrel.Target;
        }
        var key = partName + "|" + rid;
        if (_images.TryGetValue(key, out var cached)) return cached;
        ImageSource? img = null;
        try
        {
            if (_pkg.Relationships(partName).TryGetValue(rid, out var rel) && !rel.External && _pkg.Find(rel.Target) is { } entry)
            {
                // EMF / WMF : joués par GDI (WPF ne les décode pas)
                if (Layout.Metafiles.IsMetafile(entry.Data)) { img = Layout.Metafiles.Render(entry.Data); _images[key] = img; return img; }
                var bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.StreamSource = new MemoryStream(entry.Data);
                bmp.EndInit();
                bmp.Freeze();
                img = bmp;
            }
        }
        catch (Exception) { img = null; }      // format non pris en charge par WPF (EMF, WMF…) : cadre gris
        _images[key] = img;
        return img;
    }

    /// <summary>Miniatures des pages (volet gauche) : même dessin que la vue, à petite échelle ; clic = aller à la page.</summary>
    void BuildThumbnails()
    {
        ThumbsPanel.Children.Clear();
        if (_drawings is null) return;
        _thumbsView = new Layout.PagesView(_drawings, 0.18 * PtToDip, 10, thumbnails: true) { Cursor = Cursors.Hand, HorizontalAlignment = HorizontalAlignment.Center };
        _thumbsView.MouseLeftButtonUp += (_, e) =>
        {
            var (page, _) = _thumbsView.HitPage(e.GetPosition(_thumbsView));
            if (page >= 0) GoToPage(page);
        };
        ThumbsPanel.Children.Add(_thumbsView);
    }

    /// <summary>Fait défiler la vue jusqu'au haut de la page <paramref name="i"/>.</summary>
    internal void GoToPage(int i)
    {
        if (_pagesView is null) return;
        var r = _pagesView.PageRect(i);
        if (r.IsEmpty) return;
        var top = _pagesView.TransformToAncestor(Scroller).Transform(new Point(0, r.Top));
        Scroller.ScrollToVerticalOffset(Scroller.VerticalOffset + top.Y - 12);
    }

    /// <summary>Page affichée (au tiers haut de la zone visible) : « Page x sur y » et miniature en évidence.</summary>
    void UpdatePageStatus()
    {
        if (_pagesView is null || _layout is null || _info is null) return;
        int current = 0;
        try
        {
            var p = Scroller.TransformToDescendant(_pagesView).Transform(new Point(0, Scroller.ViewportHeight / 3));
            current = _pagesView.PageAt(p.Y);
        }
        catch (InvalidOperationException) { }
        int count = _layout.Pages.Count;
        PageStatus.Text = $"Page {current + 1} sur {count}";
        PageStatus.ToolTip = _info.SavedPages is { } saved && saved != count
            ? $"Mise en page d'OpenWrite : {count} page(s) ; Word en avait enregistré {saved}."
            : "Mise en page calculée par OpenWrite.";
        if (_thumbsView is not null) _thumbsView.Current = current;
    }

    void Scroller_ScrollChanged(object sender, ScrollChangedEventArgs e)
    {
        if (e.HorizontalChange != 0 || e.ViewportWidthChange != 0 || e.ExtentWidthChange != 0) UpdateRuler();
        if (e.VerticalChange != 0 || e.ViewportHeightChange != 0) UpdatePageStatus();
    }

    /// <summary>Survol : main sur une case à cocher, curseur de texte ailleurs ; bouton enfoncé : sélection.</summary>
    void PagesView_MouseMove(object sender, MouseEventArgs e)
    {
        if (_pagesView is null || _layout is null) return;
        var (page, at) = _pagesView.HitPage(e.GetPosition(_pagesView));
        if (!_mouseSelecting && ObjectMouseMove(page, at, e)) return;
        DragSelect(e);
        if (!_mouseSelecting && HoverTableBorder(page, at)) return;
        bool onBox = page >= 0 && Layout.PagesView.HitLine(_layout.Pages[page], at) is { } hit && hit.Char >= 0
                     && hit.Line.Paragraph.Content.CheckBoxes.ContainsKey(hit.Char);
        _pagesView.Cursor = onBox ? Cursors.Hand : Cursors.IBeam;
        HoverLink(HitCharacter(e.GetPosition(_pagesView)));
    }

    /// <summary>Fin d'un clic : un clic simple sur une case à cocher de formulaire la coche (annulable par Ctrl+Z).</summary>
    void PagesView_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (_pagesView is null) return;
        var (bp, ba) = _pagesView.HitPage(e.GetPosition(_pagesView));
        if (ObjectMouseUp(bp, ba)) { e.Handled = true; return; }
        if (EndBorderDrag(bp, ba)) { e.Handled = true; return; }
        if (EndMouseSelect() && ClickAt(_pagesView.HitPage(e.GetPosition(_pagesView)))) e.Handled = true;
    }

    /// <summary>Clic sur une page (points) : coche ou décoche la case de formulaire sous le pointeur ; vrai si une case a changé.</summary>
    internal bool ClickAt((int Page, Point At) target)
    {
        if (_layout is null || target.Page < 0 || target.Page >= _layout.Pages.Count) return false;
        if (Layout.PagesView.HitLine(_layout.Pages[target.Page], target.At) is not { } hit || hit.Char < 0) return false;
        if (!hit.Line.Paragraph.Content.CheckBoxes.TryGetValue(hit.Char, out var begin) || Body is not { } body) return false;
        _fieldsByBegin ??= FieldIndex.Build(body).All.Where(f => f.Begin is not null).ToDictionary(f => f.Begin!);
        if (!_fieldsByBegin.TryGetValue(begin, out var field)) return false;
        ToggleCheckBox(hit.Line.Paragraph.Paragraph, field);
        return true;
    }

    /// <summary>Résumé de la mise en forme résolue (volet « Mise en forme »).</summary>
    string Describe(Core.Styles.ParagraphFormat pf, Core.Styles.RunFormat rf)
    {
        static string Cm(int twips) => (twips / 567.0).ToString("0.##", System.Globalization.CultureInfo.CurrentCulture) + " cm";
        static string Pt(int twips) => (twips / 20.0).ToString("0.#", System.Globalization.CultureInfo.CurrentCulture) + " pt";
        var style = pf.StyleId is { } id ? _styles!.Sheet.NameOf(id) ?? id : "Normal";
        var line = pf.LineRule switch
        {
            "exact" => "exactement " + Pt(pf.Line),
            "atLeast" => "au moins " + Pt(pf.Line),
            _ => (pf.Line / 240.0).ToString("0.##", System.Globalization.CultureInfo.CurrentCulture),
        };
        var attrs = string.Join(", ", new[] { rf.Bold ? "gras" : null, rf.Italic ? "italique" : null, rf.Underline is not null ? "souligné" : null }.Where(x => x is not null));
        return $"Style : {style}\nPolice : {rf.FontAscii ?? "?"} {rf.Size.ToString("0.#", System.Globalization.CultureInfo.CurrentCulture)} pt" +
               (attrs.Length > 0 ? $" ({attrs})" : "") + (rf.Color is { } c ? $", couleur #{c}" : "") +
               $"\nAlignement : {pf.Alignment switch { "center" => "centré", "right" => "à droite", "both" => "justifié", _ => "à gauche" }}" +
               $"\nRetraits : gauche {Cm(pf.IndentLeft)}, droite {Cm(pf.IndentRight)}, 1re ligne {Cm(pf.IndentFirstLine)}" +
               $"\nEspacement : avant {Pt(pf.SpacingBefore)}, après {Pt(pf.SpacingAfter)}, interligne {line}" +
               (pf.NumId > 0 ? $"\nListe : numéro {pf.NumId}, niveau {pf.ListLevel + 1}" : "");
    }

    /// <summary>Volet gauche : miniatures des pages (construites par ShowDocument) ; plan des titres.</summary>
    void RefreshPanes()
    {
        OutlinePanel.Children.Clear();
        if (_info is not { } info)
        {
            ThumbsPanel.Children.Clear();
            _thumbsView = null;
            ThumbsPanel.Children.Add(Hint("Ouvrez un document pour voir ses pages."));
            OutlinePanel.Children.Add(Hint("Ouvrez un document pour voir ses titres."));
            return;
        }
        // miniatures : construites avec la vue des pages (ShowDocument)
        if (info.Headings.Count == 0)
        {
            OutlinePanel.Children.Add(Hint("Aucun titre : les paragraphes mis en forme avec les styles Titre 1, Titre 2… apparaissent ici."));
            return;
        }
        foreach (var (level, title) in info.Headings.Take(300))
            OutlinePanel.Children.Add(new TextBlock
            {
                Text = title, TextTrimming = TextTrimming.CharacterEllipsis, ToolTip = title, FontSize = level == 1 ? 12.5 : 12,
                FontWeight = level == 1 ? FontWeights.SemiBold : FontWeights.Normal, Margin = new Thickness((level - 1) * 12, 3, 0, 3),
                Foreground = level == 1 ? R("Ink") : R("Muted"),
            });
    }

    TextBlock Hint(string text) => new() { Text = text, TextWrapping = TextWrapping.Wrap, Foreground = R("Muted"), FontSize = 12, LineHeight = 18 };

    /// <summary>Résumé pour le volet « Mise en forme » (rubrique « Ce document »).</summary>
    string DocSummary()
    {
        if (_info is not { } i) return "Aucun document ouvert.";
        var sb = new StringBuilder();
        sb.AppendLine(DocxInfo.CompatibilityLabel(i.CompatibilityMode));
        sb.AppendLine($"{i.Paragraphs} paragraphes · {i.Tables} tableau(x) · {i.Sections} section(s)");
        var extras = new List<string>();
        if (i.Pictures > 0) extras.Add($"{i.Pictures} image(s)");
        if (i.TextBoxes > 0) extras.Add($"{i.TextBoxes} zone(s) de texte");
        if (i.Footnotes > 0) extras.Add($"{i.Footnotes} note(s)");
        if (i.Comments > 0) extras.Add($"{i.Comments} commentaire(s)");
        if (i.Revisions > 0) extras.Add($"{i.Revisions} modification(s) suivie(s)");
        if (i.LegacyFormFields + i.ContentControls > 0) extras.Add($"{i.LegacyFormFields + i.ContentControls} champ(s) de formulaire");
        if (extras.Count > 0) sb.AppendLine(string.Join(" · ", extras));
        if (i.Author is not null) sb.AppendLine($"Auteur : {i.Author}");
        if (i.Modified is { } d) sb.Append($"Modifié le {d.ToLocalTime():d MMMM yyyy}");
        return sb.ToString().TrimEnd();
    }

    /// <summary>Fichier ▸ Informations sur le document.</summary>
    void ShowDocInfo()
    {
        if (_info is not { } i || _path is null) return;
        var sb = new StringBuilder();
        sb.AppendLine(_path).AppendLine();
        sb.AppendLine($"Format : {DocxInfo.CompatibilityLabel(i.CompatibilityMode)}");
        if (i.LegacyCompatOptions.Count > 0) sb.AppendLine($"Anciennes options de compatibilité : {string.Join(", ", i.LegacyCompatOptions)}");
        if (i.Application is not null) sb.AppendLine($"Créé avec : {i.Application}");
        if (i.Title is not null) sb.AppendLine($"Titre : {i.Title}");
        if (i.Author is not null) sb.AppendLine($"Auteur : {i.Author}");
        if (i.LastModifiedBy is not null) sb.AppendLine($"Modifié par : {i.LastModifiedBy}");
        sb.AppendLine();
        sb.AppendLine($"Selon Word : {i.SavedPages?.ToString() ?? "?"} page(s), {i.SavedWords?.ToString() ?? "?"} mots");
        sb.AppendLine($"Compté par OpenWrite : {i.Words} mots, {i.Paragraphs} paragraphes dont {i.ListParagraphs} de liste");
        sb.AppendLine($"{i.Tables} tableau(x), {i.Sections} section(s), {i.HeadersFooters} en-tête(s) et pied(s) de page");
        sb.AppendLine($"{i.Pictures} image(s), {i.TextBoxes} zone(s) de texte, {i.Equations} équation(s)");
        sb.AppendLine($"{i.Footnotes} note(s) de bas de page, {i.Endnotes} note(s) de fin, {i.Comments} commentaire(s), {i.Revisions} révision(s)");
        sb.AppendLine($"{i.Bookmarks} signet(s), {i.Hyperlinks} lien(s), {i.ContentControls} contrôle(s) de contenu, {i.LegacyFormFields} champ(s) de formulaire hérité(s)");
        if (i.HasMacros) sb.AppendLine("Contient des macros (conservées, jamais exécutées).");
        if (i.Languages.Count > 0) sb.AppendLine($"Langues : {string.Join(", ", i.Languages)}");
        if (i.Fonts.Count > 0) sb.AppendLine($"Polices : {string.Join(", ", i.Fonts)}");
        sb.AppendLine($"Parties du fichier : {i.Parts.Count}");
        MessageBox.Show(this, sb.ToString(), "Informations sur le document", MessageBoxButton.OK, MessageBoxImage.Information);
    }
}
