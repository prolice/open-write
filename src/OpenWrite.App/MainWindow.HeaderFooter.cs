using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Xml.Linq;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;
using OpenWrite.Layout;
using Page = OpenWrite.Layout.Page;

namespace OpenWrite.App;

/// <summary>Lignes que parcourt le point d'insertion : celles du corps, ou celles de l'en-tête / du pied en cours d'édition.</summary>
interface ILineSource
{
    List<LineRef> LinesOf(Paragraph p);
    PlacedLine Line(LineRef r);
    LineRef? Step(LineRef r, int delta);
    IReadOnlyList<PlacedLine> PageLines(int page);
}

/// <summary>Lignes du corps (mise en page du document).</summary>
sealed class BodyLines(DocumentLayout layout) : ILineSource
{
    public List<LineRef> LinesOf(Paragraph p) => layout.LinesOf(p);
    public PlacedLine Line(LineRef r) => layout.Line(r);
    public LineRef? Step(LineRef r, int delta) => layout.Step(r, delta);
    public IReadOnlyList<PlacedLine> PageLines(int page) => page >= 0 && page < layout.Pages.Count ? layout.Pages[page].Lines : Array.Empty<PlacedLine>();
}

/// <summary>Lignes d'un en-tête ou d'un pied sur la page où on l'édite (les autres pages n'en ont pas pour le point d'insertion).</summary>
sealed class StoryLines : ILineSource
{
    readonly int _page;
    readonly List<PlacedLine> _lines;

    public StoryLines(int page, Page placed, Story story)
    {
        _page = page;
        var mine = new HashSet<Paragraph>(story.Paragraphs(), ReferenceEqualityComparer.Instance);
        _lines = placed.Lines.Where(l => mine.Contains(l.Paragraph.Paragraph)).ToList();
    }

    public int Page => _page;
    public List<LineRef> LinesOf(Paragraph p) =>
        Enumerable.Range(0, _lines.Count).Where(k => ReferenceEquals(_lines[k].Paragraph.Paragraph, p)).Select(k => new LineRef(_page, k)).ToList();
    public PlacedLine Line(LineRef r) => _lines[r.Line];
    public LineRef? Step(LineRef r, int delta) => r.Line + delta is var k && k >= 0 && k < _lines.Count ? new LineRef(_page, k) : null;
    public IReadOnlyList<PlacedLine> PageLines(int page) => page == _page ? _lines : Array.Empty<PlacedLine>();
}

// Lot 5 (version 0.6) : en-têtes et pieds de page modifiables en place, comme Word — double-clic dans la marge du haut
// ou du bas (corps grisé), Échap ou double-clic dans le corps pour en sortir ; onglet contextuel : numéro de page,
// première page différente, pages paires et impaires différentes, lier au précédent, position, passer à l'autre.
// Un en-tête qui n'existe pas encore est créé (partie, relation, référence de la section) ; tout est annulable.
public partial class MainWindow
{
    /// <summary>En-tête ou pied en cours d'édition : genre, r:id, histoire, page où on l'édite, ses lignes.</summary>
    sealed record HeaderEdit(string Kind, string RelId, Story Story, int Page, StoryLines Lines);
    HeaderEdit? _hf;

    ILineSource Src => _box is not null ? _box.Lines : _hf is not null ? _hf.Lines : _layout is not null ? new BodyLines(_layout) : EmptyLines.Instance;

    sealed class EmptyLines : ILineSource
    {
        public static readonly EmptyLines Instance = new();
        public List<LineRef> LinesOf(Paragraph p) => new();
        public PlacedLine Line(LineRef r) => throw new InvalidOperationException();
        public LineRef? Step(LineRef r, int delta) => null;
        public IReadOnlyList<PlacedLine> PageLines(int page) => Array.Empty<PlacedLine>();
    }

    /// <summary>Suite de blocs que l'on modifie : l'en-tête ou le pied en cours d'édition, sinon le corps.</summary>
    Story? EditRoot => _box?.Story ?? _hf?.Story ?? Body;

    internal bool InHeaderMode => _hf is not null;

    /// <summary>Hauteur de l'en-tête / du pied édité à la dernière mise en page : inchangée, le corps ne bouge pas.</summary>
    double _hfHeight = -1;

    /// <summary>
    /// Après une modification de l'en-tête ou du pied édité : si sa hauteur n'a pas changé, seules les pages sont
    /// redessinées (rapide) ; sinon le document est remis en page (le corps commence plus bas ou plus haut). Vrai si traité.
    /// </summary>
    bool RelayoutHeader()
    {
        if (_hf is null || _layout is null || _drawings is null || _journal.LastXml.Count > 0
            || _journal.LastChange is not { } ch || !ReferenceEquals(ch.Root, _hf.Story)) return false;
        var story = _hf.Story;
        _layout.Invalidate(story.Blocks);
        var g = _layout.Pages[Math.Clamp(_hf.Page, 0, _layout.Pages.Count - 1)].Geometry;
        double now = _layout.StoryHeightOf(story, g.ColumnWidth);
        if (Math.Abs(now - _hfHeight) > 0.01)
        {
            _layout.LayoutAll();
            _pagesView?.InvalidateMeasure();
            _thumbsView?.InvalidateMeasure();
        }
        _hfHeight = now;
        _drawings.Clear();
        RefreshHeaderLines();
        _pagesView?.InvalidateVisual();
        _thumbsView?.InvalidateVisual();
        UpdatePageStatus();
        return true;
    }

    /// <summary>Lignes de l'en-tête / du pied refaites après une mise en page (page bornée au nombre de pages).</summary>
    void RefreshHeaderLines()
    {
        if (_hf is null || _layout is null || _layout.Pages.Count == 0) return;
        int page = Math.Clamp(_hf.Page, 0, _layout.Pages.Count - 1);
        var placed = _layout.HeaderFooter(_layout.Pages[page], _layout.Pages.Count) ?? new Page { Geometry = _layout.Pages[page].Geometry, Start = default };
        _hf = _hf with { Page = page, Lines = new StoryLines(page, placed, _hf.Story) };
        _hfHeight = _layout.StoryHeightOf(_hf.Story, _layout.Pages[page].Geometry.ColumnWidth);
    }

    /// <summary>
    /// Double-clic dans la marge du haut (en-tête) ou du bas (pied) d'une page : édition de l'en-tête ou du pied de cette
    /// page (créé s'il n'existe pas).
    /// </summary>
    internal bool EnterHeader(int page, string kind, Point? at = null)
    {
        if (_layout is null || _pkg is null || Body is not { } body || page < 0 || page >= _layout.Pages.Count) return false;
        var pg = _layout.Pages[page];
        var id = _layout.HeaderId(pg, kind);
        if (id is null)
        {
            // nouvel en-tête : partie, style « En-tête » / « Pied de page », référence dans la section de la page
            var sect = pg.Geometry.SectPr ?? SectionOps.SectPrFor(body, body.Blocks.Count, body.Paragraphs().First().Props.Ns);
            var w = sect.Name.Namespace;
            using (var tx = _journal.Begin(kind == "header" ? "Nouvel en-tête" : "Nouveau pied de page", body, Array.Empty<int>(), 0, 0, null))
            {
                tx.TrackXml(sect);
                var (relId, story) = _pkg.CreateHeaderPart(kind);
                if (StylesRoot(create: true) is { } styles && _styles is not null)
                {
                    tx.TrackXml(styles);
                    if (StyleOps.Ensure(_styles.Sheet, styles, kind) is { } sid && story.Blocks.FirstOrDefault() is Paragraph first) first.Props.SetVal("pStyle", sid);
                }
                var reference = new XElement(w + (kind + "Reference"), new XAttribute(w + "type", _layout.HeaderType(pg)), new XAttribute(LinkOps.R + "id", relId));
                var last = sect.Elements().LastOrDefault(e => e.Name.LocalName is "headerReference" or "footerReference");
                if (last is not null) last.AddAfterSelf(reference); else sect.AddFirst(reference);
                id = relId;
                tx.Commit();
            }
            _styles = new StyleResolver(_pkg.StyleSheet());
            _hf = null;
            ShowDocument();
            Scroller.UpdateLayout();
            GoToPage(page);
        }
        if (_pkg.RelatedStory(id) is not { } target) return false;
        var placed = _layout!.HeaderFooter(_layout.Pages[page], _layout.Pages.Count) ?? new Page { Geometry = _layout.Pages[page].Geometry, Start = default };
        _hf = new HeaderEdit(kind, id, target, page, new StoryLines(page, placed, target));
        _hfHeight = _layout.StoryHeightOf(target, _layout.Pages[page].Geometry.ColumnWidth);
        if (_drawings is not null) { _drawings.Renderer.HeaderMode = true; _drawings.Clear(); }
        _pagesView?.InvalidateVisual();
        _thumbsView?.InvalidateVisual();
        _anchor = null;
        _caret = at is { } p && HitText(page, p) is { } t ? t : _hf.Lines.PageLines(page) is { Count: > 0 } ls ? new TextPoint(ls[0].Paragraph.Paragraph, ls[0].Paragraph.Lines[ls[0].Line].Start) : null;
        UpdateCaret();
        ShowHeaderTab(true);
        StatusText.Text = (kind == "header" ? "En-tête" : "Pied de page") + " : tapez votre texte ; Échap ou double-clic dans la page pour revenir au document.";
        return true;
    }

    /// <summary>Retour au corps du document.</summary>
    internal void ExitHeader(Point? at = null, int page = -1)
    {
        if (_hf is null) return;
        int hfPage = _hf.Page;
        _hf = null;
        if (_drawings is not null) { _drawings.Renderer.HeaderMode = false; _drawings.Clear(); }
        _pagesView?.InvalidateVisual();
        _thumbsView?.InvalidateVisual();
        _anchor = null;
        _caret = at is { } p && page >= 0 && HitText(page, p) is { } t ? t
               : _layout?.Pages.ElementAtOrDefault(hfPage)?.Lines is { Count: > 0 } ls ? new TextPoint(ls[0].Paragraph.Paragraph, ls[0].Paragraph.Lines[ls[0].Line].Start) : DocumentStart();
        UpdateCaret();
        ShowHeaderTab(false);
        StatusText.Text = "";
    }

    /// <summary>Double-clic : dans les marges d'une page, entre dans l'en-tête / le pied ; dans le corps en mode en-tête, en sort.</summary>
    bool HeaderDoubleClick(int page, Point at)
    {
        if (_layout is null || page < 0 || page >= _layout.Pages.Count) return false;
        var pg = _layout.Pages[page];
        bool top = at.Y < pg.ContentTop, bottom = at.Y > pg.ContentBottom + pg.ReservedBottom;
        if (_hf is null)
        {
            if (top) return EnterHeader(page, "header", at);
            if (bottom) return EnterHeader(page, "footer", at);
            return false;
        }
        if (!top && !bottom) { ExitHeader(at, page); return true; }
        var kind = top ? "header" : "footer";
        if (kind != _hf.Kind || page != _hf.Page) return EnterHeader(page, kind, at);
        return false;
    }

    // ================================================================ onglet contextuel « En-tête et pied de page »

    RadioButton? _headerTab;
    StackPanel? _headerPanel;
    ToggleButton? _firstPageToggle, _evenOddToggle, _linkToggle;

    void ShowHeaderTab(bool show)
    {
        if (_headerTab is null)
        {
            _headerTab = new RadioButton { Style = (Style)FindResource("RibbonTab"), Content = "En-tête et pied de page", Name = "TabHeaderFooter", Foreground = R("SelectedText") };
            _headerTab.Checked += Tab_Checked;
            TabStrip.Children.Add(_headerTab);
            _headerPanel = BuildHeaderPanel();
            _panels["TabHeaderFooter"] = _headerPanel;
            RibbonPanels.Children.Add(_headerPanel);
        }
        _headerTab.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
        if (show) { RefreshHeaderToggles(); _headerTab.IsChecked = true; }
        else if (_headerTab.IsChecked == true) TabHome.IsChecked = true;
    }

    StackPanel BuildHeaderPanel()
    {
        var panel = new StackPanel { Orientation = Orientation.Horizontal, Visibility = Visibility.Collapsed };
        Button Big(string glyph, string label, string tip, Action<Button> act)
        {
            var b = new Button { Style = (Style)FindResource("RBig"), Tag = glyph, Content = label, ToolTip = tip, Focusable = false };
            b.Click += (_, _) => act(b);
            return b;
        }
        ToggleButton Toggle(string glyph, string label, string tip, Action act)
        {
            var t = new ToggleButton { Style = (Style)FindResource("RBig"), Tag = glyph, Content = label, ToolTip = tip, Focusable = false };
            t.Click += (_, _) => { act(); RefreshHeaderToggles(); _pagesView?.Focus(); };
            return t;
        }
        void Group(string header, params UIElement[] items)
        {
            var sp = new StackPanel { Orientation = Orientation.Horizontal };
            foreach (var i in items) sp.Children.Add(i);
            panel.Children.Add(new HeaderedContentControl { Style = (Style)FindResource("RGroup"), Header = header, Content = sp });
        }
        Group("En-tête et pied",
            Big("", "N° de page", "Insérer le numéro de page, le nombre de pages ; format des numéros", b => ShowMenu(b, new (string, UIElement?, Action)[]
            {
                ("Insérer le numéro de page", null, () => InsertPageField("PAGE")),
                ("Insérer le nombre de pages", null, () => InsertPageField("NUMPAGES")),
                ("Insérer « Page X sur Y »", null, () => { InsertText("Page "); InsertPageField("PAGE"); InsertText(" sur "); InsertPageField("NUMPAGES"); }),
                ("Format des numéros de page…", null, ShowPageNumberFormat),
            })),
            Big("", "Atteindre", "Passer de l'en-tête au pied de page et inversement", _ => { if (_hf is { } h) EnterHeader(h.Page, h.Kind == "header" ? "footer" : "header"); }));
        _firstPageToggle = Toggle("", "1re page différente", "En-tête et pied différents sur la première page de la section", ToggleFirstPage);
        _evenOddToggle = Toggle("", "Paires et impaires", "En-têtes et pieds différents sur les pages paires et impaires (tout le document)", ToggleEvenOdd);
        _linkToggle = Toggle("", "Lier au précédent", "Reprendre l'en-tête de la section précédente", ToggleLink);
        Group("Options", _firstPageToggle, _evenOddToggle, _linkToggle);
        Group("Position", Big("", "Position", "Distance de l'en-tête au haut de la page et du pied au bas", _ => ShowPageSetup()));
        Group("Fermer", Big("", "Fermer", "Revenir au document (Échap)", _ => ExitHeader()));
        return panel;
    }

    void RefreshHeaderToggles()
    {
        if (_hf is null || _layout is null || _firstPageToggle is null) return;
        var pg = _layout.Pages.ElementAtOrDefault(_hf.Page);
        _firstPageToggle.IsChecked = pg?.Geometry.TitlePage == true;
        _evenOddToggle!.IsChecked = _layout.EvenOddHeaders;
        // lié au précédent : la section n'a pas sa propre référence de ce type
        var own = pg?.Geometry.SectPr?.Elements().Any(e => e.Name.LocalName == _hf.Kind + "Reference"
            && (e.Attributes().FirstOrDefault(a => a.Name.LocalName == "type")?.Value ?? "default") == _layout.HeaderType(pg)) == true;
        bool first = _layout.Sections.Count > 0 && pg is not null && ReferenceEquals(pg.Geometry, _layout.Sections[0]);
        _linkToggle!.IsChecked = !own && !first;
        _linkToggle.IsEnabled = !first;
    }

    /// <summary>Modifie le w:sectPr de la page éditée (ou settings.xml) en une opération annulable, puis remise en page complète.</summary>
    void EditHeaderSection(string label, Action<XElement> change, XElement? extra = null)
    {
        if (_hf is null || _layout is null || Body is not { } body) return;
        var sect = _layout.Pages[_hf.Page].Geometry.SectPr ?? SectionOps.SectPrFor(body, body.Blocks.Count, body.Paragraphs().First().Props.Ns);
        int page = _hf.Page;
        var kind = _hf.Kind;
        using (var tx = _journal.Begin(label, body, Array.Empty<int>(), 0, 0, null))
        {
            tx.TrackXml(sect);
            if (extra is not null) tx.TrackXml(extra);
            change(sect);
            tx.Commit();
        }
        _styles = new StyleResolver(_pkg!.StyleSheet());
        _hf = null;
        ShowDocument();
        Scroller.UpdateLayout();
        EnterHeader(page, kind);
        StatusText.Text = $"{label}. Ctrl+Z pour annuler.";
    }

    void ToggleFirstPage() => EditHeaderSection("Première page différente", s =>
    {
        var w = s.Name.Namespace;
        if (s.Element(w + "titlePg") is { } tp) tp.Remove();
        else SectionOps.Child(s, "titlePg");
    });

    void ToggleEvenOdd()
    {
        if (_pkg?.XmlPartOf("settings")?.Doc.Root is not { } settings) return;
        var w = settings.Name.Namespace;
        EditHeaderSection("Pages paires et impaires différentes", _ =>
        {
            if (settings.Element(w + "evenAndOddHeaders") is { } eo) eo.Remove();
            else
            {
                // à sa place dans le schéma : après w:zoom… ; à défaut en tête (Word relit les deux)
                var e = new XElement(w + "evenAndOddHeaders");
                if (settings.Element(w + "mirrorMargins") is { } mm) mm.AddAfterSelf(e);
                else if (settings.Element(w + "defaultTabStop") is { } dt) dt.AddBeforeSelf(e);
                else settings.AddFirst(e);
            }
        }, extra: settings);
    }

    /// <summary>
    /// Lier au précédent : retire la référence propre de la section (l'en-tête précédent s'applique) ; délier : copie de
    /// l'en-tête hérité dans une nouvelle partie, propre à la section.
    /// </summary>
    void ToggleLink()
    {
        if (_hf is null || _layout is null || _pkg is null) return;
        var pg = _layout.Pages[_hf.Page];
        var type = _layout.HeaderType(pg);
        var kind = _hf.Kind;
        var source = _hf.Story;
        EditHeaderSection(_linkToggle?.IsChecked == true ? "Délier de la section précédente" : "Lier à la section précédente", s =>
        {
            var w = s.Name.Namespace;
            var own = s.Elements(w + (kind + "Reference")).FirstOrDefault(e => (e.Attribute(w + "type")?.Value ?? "default") == type);
            if (own is not null) { own.Remove(); return; }
            var (relId, story) = _pkg.CreateHeaderPart(kind);
            story.Blocks.Clear();
            story.Blocks.AddRange(source.Blocks.Select(b => b.Clone()));
            var reference = new XElement(w + (kind + "Reference"), new XAttribute(w + "type", type), new XAttribute(LinkOps.R + "id", relId));
            var last = s.Elements().LastOrDefault(e => e.Name.LocalName is "headerReference" or "footerReference");
            if (last is not null) last.AddAfterSelf(reference); else s.AddFirst(reference);
        });
    }

    /// <summary>Champ PAGE, NUMPAGES… au point d'insertion (en-tête, pied ou corps).</summary>
    internal void InsertPageField(string code)
    {
        if (EditRoot is not { } root || _caret is not { } c || ToDocPos(c) is not { } at) return;
        var result = code == "PAGE" ? "1" : _layout?.Pages.Count.ToString() ?? "1";
        var pos = TextOps.Edit(_journal, root, code == "PAGE" ? "Numéro de page" : "Nombre de pages", at.ContainerPath, at.BlockIndex, 1, at,
            () => FieldOps.Insert(root, at, code, result), p => p);
        AfterTyping(pos);
    }

    void InsertText(string text) => ReplaceSelection(text, "Frappe");

    void ShowPageNumberFormat()
    {
        if (_hf is null || _layout is null) return;
        var g = _layout.Pages[_hf.Page].Geometry;
        var dlg = new PageNumberWindow(g.NumberFormat, g.NumberStart) { Owner = this };
        if (dlg.ShowDialog() == true) EditHeaderSection("Format des numéros de page", s => SectionOps.SetPageNumbering(s, dlg.Start, dlg.Format));
    }
}

/// <summary>Boîte « Format des numéros de page » : format (1, i, I, a, A) et numéro de départ de la section.</summary>
public sealed class PageNumberWindow : Window
{
    readonly ComboBox _format = DialogKit.Combo(180, ("1, 2, 3…", "decimal"), ("i, ii, iii…", "lowerRoman"), ("I, II, III…", "upperRoman"), ("a, b, c…", "lowerLetter"), ("A, B, C…", "upperLetter"));
    readonly RadioButton _continue = new() { Content = "Suite de la section précédente", Margin = new Thickness(0, 4, 0, 0) };
    readonly RadioButton _restart = new() { Content = "Commencer à", Margin = new Thickness(0, 4, 8, 0) };
    readonly TextBox _start = DialogKit.Box(60);

    public string Format => DialogKit.Tag(_format);
    public int? Start => _restart.IsChecked == true && int.TryParse(_start.Text.Trim(), out var s) ? s : null;

    public PageNumberWindow(string format, int? start)
    {
        DialogKit.Setup(this, "Format des numéros de page");
        var root = new StackPanel { Margin = new Thickness(18), Width = 340 };
        root.Children.Add(DialogKit.Row("Format", _format, 90));
        root.Children.Add(DialogKit.Section("Numérotation"));
        root.Children.Add(_continue);
        var row = new StackPanel { Orientation = Orientation.Horizontal };
        row.Children.Add(_restart);
        row.Children.Add(_start);
        root.Children.Add(row);
        root.Children.Add(DialogKit.Buttons(this, () => _restart.IsChecked != true || Start is not null));
        Content = root;
        DialogKit.Select(_format, format);
        if (_format.SelectedIndex < 0) _format.SelectedIndex = 0;
        _restart.IsChecked = start is not null;
        _continue.IsChecked = start is null;
        _start.Text = (start ?? 1).ToString();
    }
}
