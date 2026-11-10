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

namespace OpenWrite.App;

// Zone de travail de la version 0.1 : accueil (ouvrir, documents récents, feuille de route) et aperçu
// d'un document en lecture seule (texte brut des premiers paragraphes, informations). La vraie vue des pages
// (DocumentView virtualisée, PLAN.md §3.3) la remplace à la version 0.3.
public partial class MainWindow
{
    Brush R(string key) => (Brush)FindResource(key);

    /// <summary>Étapes affichées sur l'accueil (PLAN.md §8.17).</summary>
    static readonly (string Version, string Text)[] Milestones =
    {
        ("0.3", "Mise en page fidèle, pages et miniatures"),
        ("0.4", "Écriture et mise en forme, copier-coller avec Word"),
        ("0.5", "Styles et listes"),
        ("0.6", "Sections, en-têtes et pieds de page"),
        ("0.7", "Tableaux"),
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
                   "mis en forme (sans pagination), coche les cases des formulaires et enregistre sans rien perdre du reste du document.",
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

    /// <summary>Nombre maximal de blocs affichés dans l'aperçu (la vue des pages virtualisée arrive avec la 0.3).</summary>
    const int PreviewBlocks = 400;

    const double Px = 96.0 / 1440;            // twips → pixels (96 ppp)

    /// <summary>
    /// Aperçu de la version 0.2 : le corps du document, mis en forme par le résolveur de styles (polices, tailles, gras,
    /// couleurs, alignement, retraits, espacements, interligne), dans une page à la largeur et aux marges de la section —
    /// mais sans pagination ni objets flottants (version 0.3 et suivantes). Les cases à cocher se cochent d'un clic.
    /// </summary>
    void ShowDocument()
    {
        if (_info is not { } info || _path is null || Body is not { } body || _styles is not { } styles) return;
        Workspace.Children.Clear();
        Scroller.HorizontalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto;
        var (pageW, marginL, marginR, marginT) = PageGeometry(body);
        var sp = new StackPanel();
        sp.Children.Add(new Border
        {
            Width = pageW, Background = R("VersionBadge"), CornerRadius = new CornerRadius(12), Padding = new Thickness(14, 8, 14, 9), Margin = new Thickness(0, 0, 0, 16),
            Child = new TextBlock
            {
                TextWrapping = TextWrapping.Wrap, Foreground = R("VersionText"), FontSize = 12.5,
                Inlines =
                {
                    new System.Windows.Documents.Run("Aperçu mis en forme, sans pagination (version 0.2). ") { FontWeight = FontWeights.Bold },
                    new System.Windows.Documents.Run("Pages, images et zones de texte arrivent avec la version 0.3. " +
                        (HasCheckBoxes ? "Les cases des formulaires se cochent d'un clic ; Ctrl+S enregistre sans rien perdre du reste." : "L'écriture arrive avec la version 0.4.")),
                },
            },
        });
        var page = new Border
        {
            Width = pageW, MinHeight = pageW * 1.414, Background = Brushes.White, BorderBrush = R("Line"), BorderThickness = new Thickness(1),
            Padding = new Thickness(marginL, marginT, marginR, marginT), Effect = CardShadow,
        };
        var content = new StackPanel();
        _fieldsByBegin = FieldIndex.Build(body).All.Where(f => f.Begin is not null).ToDictionary(f => f.Begin!);
        int shown = 0;
        foreach (var block in body.Blocks)
        {
            if (shown++ >= PreviewBlocks)
            {
                content.Children.Add(new TextBlock { Text = $"… l'aperçu se limite aux {PreviewBlocks} premiers blocs (paragraphes et tableaux).", Foreground = R("Muted"), FontStyle = FontStyles.Italic, Margin = new Thickness(0, 12, 0, 0) });
                break;
            }
            if (RenderBlock(block, null) is { } el) content.Children.Add(el);
        }
        page.Child = content;
        sp.Children.Add(page);
        Workspace.Children.Add(sp);
    }

    Dictionary<FieldChar, Field> _fieldsByBegin = new();

    /// <summary>Largeur de la page et marges (pixels) d'après la dernière section ; A4 et 2,5 cm à défaut.</summary>
    (double W, double L, double R, double T) PageGeometry(Story body)
    {
        var sect = body.FinalSection?.Element;
        var w = sect?.Name.Namespace;
        double Get(string el, string attr, double def) =>
            sect?.Element(w! + el)?.Attribute(w! + attr)?.Value is { } v && double.TryParse(v, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var d) ? d : def;
        return (Get("pgSz", "w", 11906) * Px, Get("pgMar", "left", 1417) * Px, Get("pgMar", "right", 1417) * Px, Math.Min(Get("pgMar", "top", 1417), 1700) * Px);
    }

    FrameworkElement? RenderBlock(Block block, string? tableStyle) => block switch
    {
        Paragraph p => RenderParagraph(p, tableStyle),
        Table t => RenderTable(t),
        BlockContentControl s => RenderStack(s.Content, tableStyle),
        _ => null,
    };

    FrameworkElement RenderStack(Story story, string? tableStyle)
    {
        var sp = new StackPanel();
        foreach (var b in story.Blocks)
            if (RenderBlock(b, tableStyle) is { } el) sp.Children.Add(el);
        return sp;
    }

    /// <summary>Tableau : grille de cellules de même largeur, bordures fines (la mise en page des tableaux est le lot 6).</summary>
    FrameworkElement RenderTable(Table t)
    {
        var style = t.Props.Val("tblStyle") ?? "";
        var rows = t.Items.OfType<TableRow>().ToList();
        int cols = Math.Max(1, rows.Select(r => r.Items.OfType<TableCell>().Count()).DefaultIfEmpty(1).Max());
        var grid = new Grid { Margin = new Thickness(0, 4, 0, 8) };
        for (int c = 0; c < cols; c++) grid.ColumnDefinitions.Add(new ColumnDefinition());
        for (int r = 0; r < rows.Count; r++)
        {
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            int c = 0;
            foreach (var cell in rows[r].Items.OfType<TableCell>())
            {
                var border = new Border
                {
                    BorderBrush = new SolidColorBrush(Color.FromRgb(0xBF, 0xC6, 0xD4)), BorderThickness = new Thickness(0.5), Padding = new Thickness(5, 2, 5, 2),
                    Child = RenderStack(cell.Content, style),
                };
                Grid.SetRow(border, r);
                Grid.SetColumn(border, Math.Min(c, cols - 1));
                if (c == rows[r].Items.OfType<TableCell>().Count() - 1 && c < cols - 1) Grid.SetColumnSpan(border, cols - c);
                grid.Children.Add(border);
                c++;
            }
        }
        return grid;
    }

    static readonly Dictionary<string, Color> HighlightColors = new()
    {
        ["yellow"] = Colors.Yellow, ["green"] = Colors.Lime, ["cyan"] = Colors.Cyan, ["magenta"] = Colors.Magenta, ["blue"] = Colors.Blue,
        ["red"] = Colors.Red, ["darkBlue"] = Colors.DarkBlue, ["darkCyan"] = Colors.DarkCyan, ["darkGreen"] = Colors.DarkGreen,
        ["darkMagenta"] = Colors.DarkMagenta, ["darkRed"] = Colors.DarkRed, ["darkYellow"] = Colors.Olive, ["darkGray"] = Colors.DarkGray,
        ["lightGray"] = Colors.LightGray, ["black"] = Colors.Black, ["white"] = Colors.White,
    };

    FrameworkElement RenderParagraph(Paragraph p, string? tableStyle)
    {
        var styles = _styles!;
        var pf = styles.Paragraph(p, tableStyle);
        var mark = styles.ParagraphMark(p, tableStyle);
        var tb = new TextBlock
        {
            TextWrapping = TextWrapping.Wrap,
            TextAlignment = pf.Alignment switch { "center" => TextAlignment.Center, "right" => TextAlignment.Right, "both" or "distribute" => TextAlignment.Justify, _ => TextAlignment.Left },
            Margin = new Thickness(Math.Max(0, pf.IndentLeft + Math.Min(0, pf.IndentFirstLine)) * Px, pf.SpacingBefore * Px, Math.Max(0, pf.IndentRight) * Px, pf.SpacingAfter * Px),
            FontFamily = Family(mark.FontAscii), FontSize = Math.Max(1, mark.Size) * 96 / 72,
        };
        // interligne : exact / au moins en twips ; auto en multiple d'une ligne (approximation : 1,17 × taille, lot 2 pour la métrique exacte)
        double size = mark.Size * 96 / 72;
        var firstRun = p.Runs().FirstOrDefault(r => r.Text.Length > 0);
        if (firstRun is not null) size = styles.Run(p, firstRun, tableStyle).Size * 96 / 72;
        tb.LineStackingStrategy = LineStackingStrategy.BlockLineHeight;
        tb.LineHeight = pf.LineRule switch
        {
            "exact" or "atLeast" => Math.Max(pf.Line * Px, pf.LineRule == "atLeast" ? size * 1.17 : 1),
            _ => Math.Max(1, size * 1.17 * pf.Line / 240.0),
        };
        // puce ou numéro : symbole indicatif (les compteurs de liste sont calculés à la version 0.5)
        if (pf.NumId > 0 && styles.ListMarker(p, tableStyle) is { } lm)
        {
            var lvl = styles.Sheet.ListLevel(pf.NumId, pf.ListLevel);
            var w = lvl?.Name.Namespace;
            bool bullet = lvl?.Element(w! + "numFmt")?.Attribute(w! + "val")?.Value == "bullet";
            tb.Inlines.Add(new System.Windows.Documents.Run((bullet ? "•" : "#.") + " ")
            {
                FontSize = lm.Size * 96 / 72, Foreground = lm.Color is { } lc ? Brush(lc) : Brushes.Black,
                ToolTip = bullet ? null : "Numéro de liste : calculé à partir de la version 0.5",
            });
        }
        AddInlines(tb, p, p.Inlines, tableStyle, new FieldState());
        // retrait de première ligne positif : espace de la largeur du retrait devant le premier caractère
        if (pf.IndentFirstLine > 0 && pf.NumId == 0 && tb.Inlines.FirstInline is { } fi)
            tb.Inlines.InsertBefore(fi, new System.Windows.Documents.InlineUIContainer(new Border { Width = pf.IndentFirstLine * Px, Height = 1 }));
        if (tb.Inlines.Count == 0 || tb.Inlines.All(i => i is System.Windows.Documents.Run { Text: "" })) tb.Text = "​";   // paragraphe vide : hauteur de sa ligne
        // volet de droite : mise en forme du paragraphe survolé
        tb.MouseEnter += (_, _) => PaneSelection.Text = Describe(pf, firstRun is null ? mark : styles.Run(p, firstRun, tableStyle));
        return tb;
    }

    /// <summary>État des champs pendant le parcours d'un paragraphe : le code est masqué, le résultat affiché.</summary>
    sealed class FieldState
    {
        public readonly Stack<bool> InCode = new();
        public bool Hidden => InCode.Any(c => c);
    }

    void AddInlines(TextBlock tb, Paragraph p, IEnumerable<Inline> inlines, string? tableStyle, FieldState fs)
    {
        foreach (var i in inlines)
        {
            switch (i)
            {
                case Run r:
                    var rf = _styles!.Run(p, r, tableStyle);
                    foreach (var c in r.Content)
                    {
                        switch (c)
                        {
                            case FieldChar { Kind: FieldCharKind.Begin } begin:
                                fs.InCode.Push(true);
                                if (_fieldsByBegin.TryGetValue(begin, out var f) && f.CheckBox is { } box) tb.Inlines.Add(CheckBoxInline(p, f, box, rf));
                                break;
                            case FieldChar { Kind: FieldCharKind.Separate }:
                                if (fs.InCode.Count > 0) { fs.InCode.Pop(); fs.InCode.Push(false); }
                                break;
                            case FieldChar:
                                if (fs.InCode.Count > 0) fs.InCode.Pop();
                                break;
                            case TextContent { Deleted: false } t when !fs.Hidden && !rf.Hidden:
                                tb.Inlines.Add(Styled(rf.Caps ? t.Value.ToUpperInvariant() : t.Value, rf));
                                break;
                            case TabContent when !fs.Hidden:
                                tb.Inlines.Add(Styled("  ", rf));
                                break;
                            case BreakContent when !fs.Hidden:
                                tb.Inlines.Add(new System.Windows.Documents.LineBreak());
                                break;
                            case OpaqueRunContent o when !fs.Hidden && o.Xml.Name.LocalName is "drawing" or "pict" or "object" or "AlternateContent":
                                tb.Inlines.Add(new System.Windows.Documents.Run("▣") { Foreground = R("Muted"), ToolTip = "Image ou objet : affiché à partir de la version 0.3" });
                                break;
                        }
                    }
                    break;
                case InlineContainer box:
                    AddInlines(tb, p, box.Children, tableStyle, fs);
                    break;
                case InlineContentControl s:
                    AddInlines(tb, p, s.Content, tableStyle, fs);
                    break;
            }
        }
    }

    System.Windows.Documents.Run Styled(string text, Core.Styles.RunFormat rf)
    {
        var run = new System.Windows.Documents.Run(text)
        {
            FontFamily = Family(rf.FontAscii), FontSize = Math.Max(1, rf.Size) * 96 / 72,
            FontWeight = rf.Bold ? FontWeights.Bold : FontWeights.Normal, FontStyle = rf.Italic ? FontStyles.Italic : FontStyles.Normal,
            Foreground = rf.Color is { } c ? Brush(c) : Brushes.Black,
        };
        var deco = new TextDecorationCollection();
        if (rf.Underline is not null) deco.Add(TextDecorations.Underline);
        if (rf.Strike || rf.DoubleStrike) deco.Add(TextDecorations.Strikethrough);
        if (deco.Count > 0) run.TextDecorations = deco;
        if (rf.Highlight is { } hl && HighlightColors.TryGetValue(hl, out var hc)) run.Background = new SolidColorBrush(hc);
        if (rf.VerticalAlign is "superscript" or "subscript")
        {
            run.BaselineAlignment = rf.VerticalAlign == "superscript" ? BaselineAlignment.Superscript : BaselineAlignment.Subscript;
            run.FontSize *= 0.65;
        }
        if (rf.SmallCaps) run.Typography.Capitals = FontCapitals.SmallCaps;
        return run;
    }

    /// <summary>Case à cocher d'un formulaire : ☐ / ☒ cliquable (annulable par Ctrl+Z).</summary>
    System.Windows.Documents.Inline CheckBoxInline(Paragraph p, Field field, FormCheckBox box, Core.Styles.RunFormat rf)
    {
        var glyph = new TextBlock
        {
            Text = box.Glyph, FontSize = Math.Max(10, rf.Size) * 96 / 72 * 1.1, Cursor = Cursors.Hand, Foreground = R("SelectedText"),
            ToolTip = $"Case « {box.Name ?? "sans nom"} » : cliquer pour {(box.Checked ? "décocher" : "cocher")}" + (box.Enabled ? "" : " (désactivée dans Word)"),
            Margin = new Thickness(1, 0, 1, 0),
        };
        glyph.MouseLeftButtonUp += (_, e) => { e.Handled = true; ToggleCheckBox(p, field); };
        return new System.Windows.Documents.InlineUIContainer(glyph) { BaselineAlignment = BaselineAlignment.Center };
    }

    static readonly Dictionary<string, FontFamily> Families = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Police du document (police système si elle est installée, sinon police de repli de WPF).</summary>
    static FontFamily Family(string? name)
    {
        name ??= "Calibri";
        if (!Families.TryGetValue(name, out var f)) Families[name] = f = new FontFamily(name + ", Segoe UI");
        return f;
    }

    static SolidColorBrush Brush(string rgb)
    {
        var b = new SolidColorBrush(Color.FromRgb(Convert.ToByte(rgb[..2], 16), Convert.ToByte(rgb[2..4], 16), Convert.ToByte(rgb[4..], 16)));
        b.Freeze();
        return b;
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

    /// <summary>Volet gauche : une vignette par page enregistrée par Word (sans contenu avant la 0.3) ; plan des titres.</summary>
    void RefreshPanes()
    {
        ThumbsPanel.Children.Clear();
        OutlinePanel.Children.Clear();
        if (_info is not { } info)
        {
            ThumbsPanel.Children.Add(Hint("Ouvrez un document pour voir ses pages."));
            OutlinePanel.Children.Add(Hint("Ouvrez un document pour voir ses titres."));
            return;
        }
        ThumbsPanel.Children.Add(Hint(info.SavedPages is { } p
            ? $"{p} page{(p > 1 ? "s" : "")} selon Word. Les miniatures s'afficheront avec la mise en page (version 0.3)."
            : "Les miniatures s'afficheront avec la mise en page (version 0.3)."));
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
