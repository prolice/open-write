using System.Windows;
using System.Windows.Media;
using System.Xml.Linq;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;

namespace OpenWrite.App;

// Lot 5 (version 0.6) : arrière-plan de la page — filigrane (forme WordArt placée dans les en-têtes, comme Word),
// couleur de page (w:background), bordures de page (w:pgBorders) ; numéros de ligne (w:lnNumType). Annulables.
public partial class MainWindow
{
    static readonly XNamespace V = "urn:schemas-microsoft-com:vml";
    static readonly XNamespace O = "urn:schemas-microsoft-com:office:office";
    static readonly XNamespace W10 = "urn:schemas-microsoft-com:office:word";

    internal void ShowWatermarkMenu(FrameworkElement anchor) => ShowMenu(anchor, new[] { "CONFIDENTIEL", "BROUILLON", "NE PAS COPIER", "URGENT", "ÉCHANTILLON" }
        .Select(t => (t, (UIElement?)null, (Action)(() => SetWatermark(t))))
        .Append(("Filigrane personnalisé…", null, (Action)(() =>
        {
            var dlg = new TextPromptWindow("Filigrane", "Texte du filigrane", "CONFIDENTIEL") { Owner = this };
            if (dlg.ShowDialog() == true && dlg.Value.Trim().Length > 0) SetWatermark(dlg.Value.Trim());
        })))
        .Append(("Supprimer le filigrane", null, (Action)(() => SetWatermark(null)))));

    /// <summary>
    /// Filigrane texte dans tous les en-têtes du document (un en-tête par défaut est créé dans les sections qui n'en ont
    /// pas) ; null le retire. Une seule opération annulable.
    /// </summary>
    internal void SetWatermark(string? text)
    {
        if (_pkg is null || Body is not { } body || _layout is null) return;
        var w = body.Paragraphs().First().Props.Ns;
        var caret = _caret is { } c0 && _hf is null ? ToDocPos(c0) : null;
        using (var tx = _journal.Begin(text is null ? "Supprimer le filigrane" : "Filigrane", body, Array.Empty<int>(), 0, 0, caret))
        {
            var stories = new List<Story>();
            foreach (var sect in SectionOps.All(body, w))
            {
                var ids = sect.Elements(w + "headerReference").Select(e => e.Attributes().FirstOrDefault(a => a.Name.LocalName == "id")?.Value).Where(i => i is not null).ToList();
                if (ids.Count == 0 && text is not null && ReferenceEquals(sect, SectionOps.All(body, w)[0]))
                {
                    // première section sans en-tête : en-tête par défaut créé (les suivantes en héritent)
                    tx.TrackXml(sect);
                    var (relId, story) = _pkg.CreateHeaderPart("header");
                    sect.AddFirst(new XElement(w + "headerReference", new XAttribute(w + "type", "default"), new XAttribute(LinkOps.R + "id", relId)));
                    ids.Add(relId);
                }
                foreach (var id in ids)
                    if (_pkg.RelatedStory(id!) is { } st && !stories.Contains(st)) stories.Add(st);
            }
            foreach (var story in stories)
            {
                tx.TrackStory(story);
                foreach (var p in story.Paragraphs())
                    foreach (var r in p.Runs())
                        r.Content.RemoveAll(x => x is DrawingContent o && o.Xml.Descendants().Any(e => e.Name.LocalName == "textpath"));
                foreach (var p in story.Paragraphs()) p.Inlines.RemoveAll(i => i is Run { Content.Count: 0 } r && r.Props.Element?.Element(w + "noProof") is not null);
                if (text is null) continue;
                if (story.Blocks.FirstOrDefault() is not Paragraph first) { first = new Paragraph { Props = new PropertyBag("pPr", w) }; story.Blocks.Insert(0, first); }
                var run = new Run { Props = new PropertyBag("rPr", w, new XElement(w + "rPr", new XElement(w + "noProof"))) };
                run.Content.Add(new DrawingContent(WatermarkXml(w, text)));
                first.Inlines.Insert(0, run);
                EnsureVmlNamespaces(story);
            }
            tx.Commit(caret);
        }
        RedrawAll(text is null ? "Filigrane supprimé." : $"Filigrane « {text} ». Ctrl+Z pour annuler.");
    }

    /// <summary>Forme WordArt du filigrane telle que Word l'écrit (VML, centrée dans la marge, tournée à 315°, argent semi-transparent).</summary>
    static XElement WatermarkXml(XNamespace w, string text)
    {
        double width = Math.Clamp(text.Length * 37, 150, 470), height = width * 0.6 / Math.Max(1, text.Length / 6.0);
        height = Math.Clamp(height, 60, 160);
        var formulas = new[] { "sum #0 0 10800", "prod #0 2 1", "sum 21600 0 @1", "sum 0 0 @2", "sum 21600 0 @3", "if @0 @3 0", "if @0 21600 @1",
            "if @0 0 @2", "if @0 @4 21600", "mid @5 @6", "mid @8 @5", "mid @7 @8", "mid @6 @7", "sum @6 0 @5" };
        var inv = System.Globalization.CultureInfo.InvariantCulture;
        return new XElement(w + "pict",
            new XElement(V + "shapetype", new XAttribute("id", "_x0000_t136"), new XAttribute("coordsize", "21600,21600"), new XAttribute(O + "spt", "136"),
                new XAttribute("adj", "10800"), new XAttribute("path", "m@7,l@8,m@5,21600l@6,21600e"),
                new XElement(V + "formulas", formulas.Select(f => new XElement(V + "f", new XAttribute("eqn", f)))),
                new XElement(V + "path", new XAttribute("textpathok", "t"), new XAttribute(O + "connecttype", "custom"),
                    new XAttribute(O + "connectlocs", "@9,0;@10,10800;@11,21600;@12,10800"), new XAttribute(O + "connectangles", "270,180,90,0")),
                new XElement(V + "textpath", new XAttribute("on", "t"), new XAttribute("fitshape", "t")),
                new XElement(V + "handles", new XElement(V + "h", new XAttribute("position", "#0,bottomRight"), new XAttribute("xrange", "6629,14971"))),
                new XElement(O + "lock", new XAttribute(V + "ext", "edit"), new XAttribute("text", "t"), new XAttribute("shapetype", "t"))),
            new XElement(V + "shape", new XAttribute("id", "PowerPlusWaterMarkObject" + Random.Shared.Next(100000000, 999999999)),
                new XAttribute(O + "spid", "_x0000_s2049"), new XAttribute("type", "#_x0000_t136"),
                new XAttribute("style", $"position:absolute;margin-left:0;margin-top:0;width:{width.ToString("0.#", inv)}pt;height:{height.ToString("0.#", inv)}pt;rotation:315;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin"),
                new XAttribute(O + "allowincell", "f"), new XAttribute("fillcolor", "silver"), new XAttribute("stroked", "f"),
                new XElement(V + "fill", new XAttribute("opacity", ".5")),
                new XElement(V + "textpath", new XAttribute("style", "font-family:\"Calibri\";font-size:1pt"), new XAttribute("string", text)),
                new XElement(W10 + "wrap", new XAttribute("anchorx", "margin"), new XAttribute("anchory", "margin"))));
    }

    /// <summary>Déclare v:, o: et w10: à la racine de la partie de l'en-tête (Word les y attend).</summary>
    void EnsureVmlNamespaces(Story story)
    {
        if (_pkg?.Stories.Values.FirstOrDefault(p => p.Stories().Contains(story)) is not { } part) return;
        foreach (var (prefix, ns) in new[] { ("v", V), ("o", O), ("w10", W10) })
            if (!part.RootAttributes.Any(a => a.IsNamespaceDeclaration && a.Value == ns.NamespaceName))
                part.RootAttributes.Add(new XAttribute(XNamespace.Xmlns + prefix, ns.NamespaceName));
    }

    /// <summary>Toutes les pages redessinées (arrière-plan, en-têtes) après une opération qui ne change pas la mise en page.</summary>
    void RedrawAll(string status)
    {
        _drawings?.Clear();
        UpdatePageBackground();
        _pagesView?.InvalidateVisual();
        _thumbsView?.InvalidateVisual();
        RefreshHeaderLines();
        UpdateModifiedMarker();
        EnsureAutoSaveTimer();
        StatusText.Text = status;
        System.Windows.Input.CommandManager.InvalidateRequerySuggested();
    }

    // ================================================================ couleur de page

    XElement? BackgroundElement() => _pkg?.Main?.Items.OfType<XElement>().FirstOrDefault(e => e.Name.LocalName == "background");

    static readonly (string Name, string? Rgb)[] PageColors =
    {
        ("Aucune couleur", null), ("Crème", "FFF8E7"), ("Bleu pâle", "DEEAF6"), ("Vert pâle", "E2EFD9"), ("Rose pâle", "FBE4D5"), ("Gris pâle", "EDEDED"), ("Jaune pâle", "FFF2CC"),
    };

    internal void ShowPageColorMenu(FrameworkElement anchor) => ShowMenu(anchor, PageColors.Select(c => (c.Name,
        (UIElement?)(c.Rgb is { } rgb ? new System.Windows.Controls.Border { Width = 14, Height = 14, BorderBrush = Brushes.Gray, BorderThickness = new Thickness(0.5),
            Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#" + rgb)) } : null),
        (Action)(() => SetPageColor(c.Rgb)))));

    /// <summary>Couleur de page (w:background du document, affichée par Word si w:displayBackgroundShape est posé).</summary>
    internal void SetPageColor(string? rgb)
    {
        if (_pkg?.Main is not { } main || Body is not { } body) return;
        var w = main.W;
        var bg = BackgroundElement();
        if (bg is null)
        {
            if (rgb is null) return;
            bg = new XElement(w + "background");
            int at = main.Items.FindIndex(i => i is not XElement);
            main.Items.Insert(Math.Max(0, at), bg);
        }
        var settings = _pkg.XmlPartOf("settings")?.Doc.Root;
        using (var tx = _journal.Begin(rgb is null ? "Supprimer la couleur de page" : "Couleur de page", body, Array.Empty<int>(), 0, 0, null))
        {
            tx.TrackXml(bg);
            if (settings is not null) tx.TrackXml(settings);
            bg.SetAttributeValue(w + "color", rgb);
            if (settings is not null && rgb is not null && settings.Element(w + "displayBackgroundShape") is null)
            {
                var e = new XElement(w + "displayBackgroundShape");
                if (settings.Element(w + "zoom") is { } z) z.AddAfterSelf(e); else settings.AddFirst(e);
            }
            tx.Commit();
        }
        main.Dirty = true;
        RedrawAll(rgb is null ? "Couleur de page supprimée." : "Couleur de page appliquée. Ctrl+Z pour annuler.");
    }

    /// <summary>Fond des pages de la vue : couleur de page du document (blanc par défaut).</summary>
    void UpdatePageBackground()
    {
        if (_pagesView is null) return;
        var rgb = BackgroundElement()?.Attributes().FirstOrDefault(a => a.Name.LocalName == "color")?.Value;
        _pagesView.PageBackground = rgb is { Length: 6 } ? new SolidColorBrush((Color)ColorConverter.ConvertFromString("#" + rgb)) : Brushes.White;
        if (_thumbsView is not null) _thumbsView.PageBackground = _pagesView.PageBackground;
    }

    // ================================================================ bordures de page, numéros de ligne

    internal void ShowPageBordersMenu(FrameworkElement anchor) => ShowMenu(anchor, new (string, UIElement?, Action)[]
    {
        ("Encadré simple", null, () => EditSections("Bordure de page", s => SectionOps.SetPageBorders(s, "single", 4), all: true)),
        ("Encadré double", null, () => EditSections("Bordure de page", s => SectionOps.SetPageBorders(s, "double", 4), all: true)),
        ("Pointillés", null, () => EditSections("Bordure de page", s => SectionOps.SetPageBorders(s, "dotted", 8), all: true)),
        ("Trait épais (3 pt)", null, () => EditSections("Bordure de page", s => SectionOps.SetPageBorders(s, "single", 24), all: true)),
        ("Encadré bleu", null, () => EditSections("Bordure de page", s => SectionOps.SetPageBorders(s, "single", 12, "087CFF"), all: true)),
        ("Aucune bordure", null, () => EditSections("Supprimer la bordure de page", s => SectionOps.SetPageBorders(s, null), all: true)),
    });

    internal void ShowLineNumbersMenu(FrameworkElement anchor) => ShowMenu(anchor, new (string, UIElement?, Action)[]
    {
        ("Aucun", null, () => EditSections("Sans numéros de ligne", s => SectionOps.SetLineNumbers(s, null))),
        ("Continu", null, () => EditSections("Numéros de ligne continus", s => SectionOps.SetLineNumbers(s, "continuous"))),
        ("Recommencer à chaque page", null, () => EditSections("Numéros de ligne par page", s => SectionOps.SetLineNumbers(s, "newPage"))),
        ("Recommencer à chaque section", null, () => EditSections("Numéros de ligne par section", s => SectionOps.SetLineNumbers(s, "newSection"))),
    });
}

/// <summary>Petite boîte de saisie d'un texte.</summary>
public sealed class TextPromptWindow : Window
{
    readonly System.Windows.Controls.TextBox _box = DialogKit.Box(280);
    public string Value => _box.Text;

    public TextPromptWindow(string title, string label, string value)
    {
        DialogKit.Setup(this, title);
        var root = new System.Windows.Controls.StackPanel { Margin = new Thickness(18) };
        root.Children.Add(DialogKit.Row(label, _box, 140));
        root.Children.Add(DialogKit.Buttons(this, () => true));
        Content = root;
        _box.Text = value;
        Loaded += (_, _) => { _box.Focus(); _box.SelectAll(); };
    }
}
