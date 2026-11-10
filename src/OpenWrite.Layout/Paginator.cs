using System.Globalization;
using System.Runtime.CompilerServices;
using System.Xml.Linq;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;

namespace OpenWrite.Layout;

// Étages 2 et 3 du moteur (PLAN.md §3.3) : les lignes des blocs sont placées dans les pages de chaque section
// (sauts, veuves et orphelines, paragraphes solidaires, lignes solidaires, saut de page avant), puis la mise en page
// est reprise de façon incrémentale : seule la page qui contient la modification est recalculée, jusqu'à ce qu'une
// page retrouve le même point de départ qu'avant (convergence) ; les pages suivantes sont réutilisées telles quelles.
//
// PROTOTYPE (lot 2, §8.3) : colonnes, en-têtes et pieds, notes, objets flottants et lignes de tableau fractionnables
// viennent ensuite (lots 2 à 9) ; une ligne de tableau ne se coupe pas ici.

/// <summary>
/// Géométrie d'une section (points). <see cref="Top"/> et <see cref="Bottom"/> : marges du corps, agrandies quand
/// l'en-tête ou le pied de page dépasse la marge (comme Word) ; <see cref="FirstTop"/> / <see cref="FirstBottom"/> :
/// idem pour la première page de la section quand elle a son propre en-tête (w:titlePg).
/// </summary>
public sealed record SectionGeometry(double PageWidth, double PageHeight, double Top, double Bottom, double Left, double Right, string Type)
{
    public double ColumnWidth => PageWidth - Left - Right;
    public double ContentBottom => PageHeight - Bottom;
    public double FirstTop { get; init; } = Top;
    public double FirstBottom { get; init; } = Bottom;
    /// <summary>Propriétés de la section (en-têtes, pieds…).</summary>
    public XElement? SectPr { get; init; }
    /// <summary>Distance du haut de la page à l'en-tête (w:header) et du bas au pied (w:footer), points.</summary>
    public double HeaderDistance { get; init; } = 35.4;
    public double FooterDistance { get; init; } = 35.4;

    public static SectionGeometry From(XElement? sectPr)
    {
        var w = sectPr?.Name.Namespace ?? XNamespace.None;
        double Get(string el, string attr, double def) =>
            sectPr?.Element(w + el)?.Attribute(w + attr)?.Value is { } v && double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? d / 20 : def / 20;
        // marge haute / basse négative : Word la prend en valeur absolue (le corps ne remonte pas sous l'en-tête)
        return new SectionGeometry(Get("pgSz", "w", 11906), Get("pgSz", "h", 16838), Math.Abs(Get("pgMar", "top", 1417)), Math.Abs(Get("pgMar", "bottom", 1417)),
            Get("pgMar", "left", 1417), Get("pgMar", "right", 1417), sectPr?.Element(w + "type")?.Attribute(w + "val")?.Value ?? "nextPage")
        {
            HeaderDistance = Get("pgMar", "header", 708), FooterDistance = Get("pgMar", "footer", 708), SectPr = sectPr,
        };
    }
}

/// <summary>Élément de flux : un paragraphe ou un tableau du corps (contenus des contrôles de contenu mis à plat).</summary>
public sealed class FlowItem
{
    public required Block Block { get; init; }
    /// <summary>Indice de la section (dans <see cref="DocumentLayout.Sections"/>).</summary>
    public int Section { get; init; }
    /// <summary>Le paragraphe porte la fin de sa section (w:pPr/w:sectPr).</summary>
    public bool EndsSection { get; init; }
}

/// <summary>Ligne de paragraphe placée sur une page.</summary>
public readonly record struct PlacedLine(ParagraphLayout Paragraph, int Line, double X, double Y)
{
    /// <summary>Place réservée en bas de page pour les notes appelées sur cette ligne.</summary>
    public double Notes { get; init; }
}

/// <summary>Cellule de tableau placée : rectangle (points), fond (RRGGBB) et bordures (épaisseur en points, 0 = aucune).</summary>
public readonly record struct CellBox(double X, double Y, double W, double H, string? Fill, double Top, double Bottom, double Left, double Right);

/// <summary>Point de départ d'une page : élément de flux et ligne (ou rangée de tableau) où elle commence.</summary>
/// <remarks>
/// Pour un tableau, <see cref="Line"/> est la rangée et <see cref="Offset"/> la hauteur de la rangée déjà placée sur
/// la page précédente (rangée coupée entre deux pages).
/// </remarks>
public readonly record struct FlowPosition(FlowItem? Item, int Index, int Line, double Offset = 0);

public sealed class Page
{
    public required SectionGeometry Geometry { get; init; }
    /// <summary>Première page de sa section (en-tête « première page » si w:titlePg).</summary>
    public bool FirstOfSection { get; init; }
    public double ContentTop => FirstOfSection ? Geometry.FirstTop : Geometry.Top;
    public double ContentBottom => Geometry.PageHeight - (FirstOfSection ? Geometry.FirstBottom : Geometry.Bottom) - ReservedBottom;
    /// <summary>Hauteur prise en bas du corps : objets flottants ancrés en bas de marge (tableau w:tblpYSpec="bottom"), notes.</summary>
    public double ReservedBottom { get; set; }
    /// <summary>Nombre de notes de bas de page placées sur la page.</summary>
    public int NotesCount { get; set; }
    public required FlowPosition Start { get; init; }
    public List<PlacedLine> Lines { get; } = new();
    /// <summary>Bordures des cellules de tableau (rectangles en points), pour l'affichage.</summary>
    public List<CellBox> Cells { get; } = new();
    /// <summary>Notes de bas de page de la page (w:id), dans l'ordre des appels.</summary>
    public List<string> NoteIds { get; } = new();
    public FlowPosition End { get; set; }
}

/// <summary>Mise en page d'un document : pages, caches des paragraphes, reprise incrémentale.</summary>
public sealed class DocumentLayout
{
    readonly ParagraphLayouter _lines;
    readonly Story _body;
    readonly ConditionalWeakTable<Paragraph, ParagraphLayout> _cache = new();
    List<FlowItem> _items = new();
    Dictionary<Block, int> _itemIndex = new(ReferenceEqualityComparer.Instance);
    /// <summary>Nombre de blocs du corps à la dernière construction du flux (un bloc inséré ou supprimé le change).</summary>
    int _topCount;

    /// <summary>
    /// Espace avant supprimé en haut d'une page commencée par le flux (pas après un saut de page explicite ni en tête
    /// de section) : comportement de Word, à calibrer sur le corpus avec le banc (--layout).
    /// </summary>
    public bool SuppressSpaceBeforeAtPageTop { get; set; } = true;

    public List<SectionGeometry> Sections { get; } = new();
    /// <summary>Dépassement toléré en bas de page (points), à calibrer sur le corpus (OPENWRITE_BOTTOM_TOL).</summary>
    public static double BottomTolerance { get; set; } =
        double.TryParse(Environment.GetEnvironmentVariable("OPENWRITE_BOTTOM_TOL"), NumberStyles.Float, CultureInfo.InvariantCulture, out var t) ? t : 0.01;
    /// <summary>Mise au point : chaque coupure de page est décrite sur la console (OPENWRITE_LAYOUT_TRACE=1).</summary>
    public static bool Trace { get; set; } = Environment.GetEnvironmentVariable("OPENWRITE_LAYOUT_TRACE") == "1";
    public List<Page> Pages { get; private set; } = new();
    public IReadOnlyList<FlowItem> Items => _items;
    /// <summary>Statistiques de la dernière mise en page : paragraphes recalculés, pages recalculées.</summary>
    public int LastParagraphsLaidOut { get; private set; }
    public int LastPagesLaidOut { get; private set; }

    readonly Func<string, Story?>? _related;
    readonly Func<string, Story?>? _footnote;

    /// <param name="related">En-têtes et pieds de page : histoire d'une partie liée par son r:id (null : ignorés).</param>
    /// <param name="footnote">Contenu d'une note de bas de page par son w:id ("separator" : le séparateur) ; null : notes ignorées.</param>
    public DocumentLayout(Story body, StyleResolver styles, IFontProvider fonts, Func<string, Story?>? related = null, Func<string, Story?>? footnote = null)
    {
        _body = body;
        _lines = new ParagraphLayouter(styles, fonts);
        _related = related;
        _footnote = footnote;
    }

    readonly Dictionary<string, double> _noteHeights = new();

    /// <summary>
    /// Place prise en bas de page par les notes appelées sur une ligne : leur texte, et le séparateur pour la première
    /// note de la page (notes de bas de page : prototype du lot 9, une note ne se continue pas sur la page suivante).
    /// </summary>
    double NotesHeight(LineBox line, Page page, double width)
    {
        if (line.Footnotes is null || _footnote is null) return 0;
        double h = 0;
        foreach (var id in line.Footnotes)
        {
            if (!_noteHeights.TryGetValue(id, out var nh))
            {
                _lines.NoteNumber = id;
                nh = _footnote(id) is { } story ? StoryHeight(story, width) : 0;
                _lines.NoteNumber = null;
                _noteHeights[id] = nh;
            }
            h += nh;
        }
        if (page.NotesCount == 0 && h > 0)
        {
            if (!_noteHeights.TryGetValue("separator", out var sep))
                _noteHeights["separator"] = sep = _footnote("separator") is { } s ? StoryHeight(s, width) : 0;
            h += sep;
        }
        return h;
    }

    public ParagraphLayouter Lines => _lines;

    // ------------------------------------------------------------------ flux

    void BuildItems()
    {
        // numéros de liste : recalculés à chaque reconstruction du flux (paragraphes insérés ou supprimés) ; les
        // paragraphes dont le numéro change sont remis en page
        var numbers = ListNumbering.Compute(_body, _lines.Styles);
        _renumbered.Clear();
        if (_lines.ListNumbers is { } old)
            foreach (var (p, n) in numbers)
                if (old.TryGetValue(p, out var o) && o != n) { _cache.Remove(p); _renumbered.Add(p); }
        _lines.ListNumbers = numbers;
        Sections.Clear();
        var items = new List<FlowItem>();
        var pending = new List<(Block, bool)>();
        void Walk(IEnumerable<Block> blocks)
        {
            foreach (var b in blocks)
            {
                switch (b)
                {
                    case Paragraph p:
                        var sect = p.Props.Get("sectPr");
                        pending.Add((p, sect is not null));
                        if (sect is not null) CloseSection(sect);
                        break;
                    case Table:
                        pending.Add((b, false));
                        break;
                    case BlockContentControl sdt:
                        Walk(sdt.Content.Blocks);
                        break;
                }
            }
        }
        void CloseSection(XElement? sectPr)
        {
            int index = Sections.Count;
            Sections.Add(WithHeaders(SectionGeometry.From(sectPr), sectPr));
            foreach (var (b, ends) in pending) items.Add(new FlowItem { Block = b, Section = index, EndsSection = ends });
            pending.Clear();
        }
        Walk(_body.Blocks);
        CloseSection(_body.FinalSection?.Element);
        _items = items;
        _topCount = _body.Blocks.Count;
        _itemIndex = new Dictionary<Block, int>(items.Count, ReferenceEqualityComparer.Instance);
        for (int i = 0; i < items.Count; i++) _itemIndex[items[i].Block] = i;
        _paraItem = null;
    }

    /// <summary>Paragraphes dont le numéro de liste a changé à la dernière reconstruction du flux (à remettre en page).</summary>
    readonly HashSet<Paragraph> _renumbered = new(ReferenceEqualityComparer.Instance);

    Dictionary<Paragraph, int>? _paraItem;

    /// <summary>
    /// Lignes placées d'un paragraphe du corps (cellules de tableau comprises), dans l'ordre : page et rang dans
    /// <see cref="Page.Lines"/>. Vide pour un paragraphe masqué ou hors du corps.
    /// </summary>
    public List<LineRef> LinesOf(Paragraph p)
    {
        var res = new List<LineRef>();
        if (_paraItem is null)
        {
            _paraItem = new Dictionary<Paragraph, int>(ReferenceEqualityComparer.Instance);
            for (int i = 0; i < _items.Count; i++)
                foreach (var q in new Story { Blocks = { _items[i].Block } }.Paragraphs()) _paraItem[q] = i;
        }
        if (!_paraItem.TryGetValue(p, out var idx) || Pages.Count == 0) return res;
        for (int k = PageOf(Pages, idx); k < Pages.Count; k++)
        {
            var pg = Pages[k];
            if (pg.Start.Item is not null && pg.Start.Index > idx) break;
            for (int j = 0; j < pg.Lines.Count; j++)
                if (ReferenceEquals(pg.Lines[j].Paragraph.Paragraph, p)) res.Add(new LineRef(k, j));
        }
        return res;
    }

    /// <summary>Ligne placée désignée par <paramref name="r"/>.</summary>
    public PlacedLine Line(LineRef r) => Pages[r.Page].Lines[r.Line];

    /// <summary>Ligne placée suivante (ou précédente) dans l'ordre du document, en passant d'une page à l'autre.</summary>
    public LineRef? Step(LineRef r, int delta)
    {
        int page = r.Page, line = r.Line + delta;
        while (page >= 0 && page < Pages.Count)
        {
            if (line >= 0 && line < Pages[page].Lines.Count) return new LineRef(page, line);
            page += delta;
            if (page < 0 || page >= Pages.Count) return null;
            line = delta > 0 ? 0 : Pages[page].Lines.Count - 1;
        }
        return null;
    }

    /// <summary>
    /// Le corps commence sous l'en-tête et finit au-dessus du pied quand ils dépassent les marges : Word prend
    /// max(marge du haut, distance de l'en-tête + hauteur de l'en-tête), et de même en bas. En-tête « par défaut »
    /// pour toutes les pages, « première page » pour la première si w:titlePg (pages paires : lot 5).
    /// </summary>
    SectionGeometry WithHeaders(SectionGeometry g, XElement? sectPr)
    {
        if (_related is null || sectPr is null) return g;
        var w = sectPr.Name.Namespace;
        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        double Extent(string kind, string type)
        {
            var reference = sectPr.Elements(w + kind + "Reference").FirstOrDefault(e => (e.Attribute(w + "type")?.Value ?? "default") == type);
            var id = reference?.Attributes().FirstOrDefault(a => a.Name.LocalName == "id")?.Value;
            return id is not null && _related(id) is { } story ? StoryHeight(story, g.ColumnWidth) : 0;
        }
        bool titlePg = sectPr.Element(w + "titlePg") is { } tp && tp.Attribute(w + "val")?.Value is null or "1" or "true" or "on";
        double top = Math.Max(g.Top, g.HeaderDistance + Extent("header", "default"));
        double bottom = Math.Max(g.Bottom, g.FooterDistance + Extent("footer", "default"));
        double firstTop = titlePg ? Math.Max(g.Top, g.HeaderDistance + Extent("header", "first")) : top;
        double firstBottom = titlePg ? Math.Max(g.Bottom, g.FooterDistance + Extent("footer", "first")) : bottom;
        return g with { Top = top, Bottom = bottom, FirstTop = firstTop, FirstBottom = firstBottom };
    }

    static double BorderOf(ParagraphLayout pl, string side)
    {
        var W = pl.Format.Effective.Name.Namespace;
        return BorderExtent(pl.Format.Effective.Element(W + "pBdr")?.Element(W + side), W);
    }

    static double BorderExtent(XElement? b, XNamespace W)
    {
        if (b is null || b.Attribute(W + "val")?.Value is null or "none" or "nil") return 0;
        double sz = double.TryParse(b.Attribute(W + "sz")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v / 8 : 0.5;
        double space = double.TryParse(b.Attribute(W + "space")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var sp) ? sp : 0;
        return sz + space;
    }

    // ------------------------------------------------------------------ affichage : en-têtes, pieds, notes

    /// <summary>
    /// En-tête et pied de la page, placés pour l'affichage : en-tête à sa distance du haut, pied terminé à sa distance
    /// du bas ; champs PAGE et NUMPAGES calculés. En-tête « première page » si w:titlePg, sinon « par défaut ».
    /// </summary>
    public Page? HeaderFooter(Page page, int pageNumber, int pageCount)
    {
        var g = page.Geometry;
        if (_related is null || g.SectPr is not { } sect) return null;
        var w = sect.Name.Namespace;
        bool titlePg = sect.Element(w + "titlePg") is { } tp && tp.Attribute(w + "val")?.Value is null or "1" or "true" or "on";
        string type = page.FirstOfSection && titlePg ? "first" : "default";
        Story? Find(string kind) =>
            sect.Elements(w + kind + "Reference").FirstOrDefault(e => (e.Attribute(w + "type")?.Value ?? "default") == type)
                ?.Attributes().FirstOrDefault(a => a.Name.LocalName == "id")?.Value is { } id ? _related(id) : null;
        var result = new Page { Geometry = g, Start = page.Start };
        _lines.FieldValue = code =>
        {
            var name = code.Split(' ', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()?.ToUpperInvariant();
            return name switch { "PAGE" => pageNumber.ToString(CultureInfo.InvariantCulture), "NUMPAGES" or "SECTIONPAGES" => pageCount.ToString(CultureInfo.InvariantCulture), _ => null };
        };
        try
        {
            if (Find("header") is { } header) PlaceStory(header, g.HeaderDistance, g, result, cache: false);
            if (Find("footer") is { } footer) PlaceStory(footer, g.PageHeight - g.FooterDistance - StoryHeight(footer, g.ColumnWidth), g, result, cache: false);
        }
        finally { _lines.FieldValue = null; }
        return result;
    }

    /// <summary>Notes de bas de page de la page, placées pour l'affichage (séparateur puis texte des notes).</summary>
    public Page? Notes(Page page)
    {
        if (page.NoteIds.Count == 0 || _footnote is null) return null;
        var g = page.Geometry;
        double height = page.NoteIds.Sum(id => _noteHeights.TryGetValue(id, out var h) ? h : 0) + (_noteHeights.TryGetValue("separator", out var sep) ? sep : 0);
        double y = page.Geometry.PageHeight - (page.FirstOfSection ? g.FirstBottom : g.Bottom) - height;
        var result = new Page { Geometry = g, Start = page.Start };
        if (_footnote("separator") is { } s)
        {
            result.Cells.Add(new CellBox(g.Left, y + sep / 2, g.ColumnWidth / 3, 0, null, 0.5, 0, 0, 0));
            y += sep;
        }
        foreach (var id in page.NoteIds)
        {
            if (_footnote(id) is not { } note) continue;
            _lines.NoteNumber = id;
            try { y = PlaceStory(note, y, g, result, cache: false); }
            finally { _lines.NoteNumber = null; }
        }
        return result;
    }

    /// <summary>Empile les paragraphes et tableaux d'une histoire à partir de <paramref name="top"/> ; renvoie le bas.</summary>
    double PlaceStory(Story story, double top, SectionGeometry g, Page into, bool cache)
    {
        double y = top;
        foreach (var b in story.Blocks)
        {
            switch (b)
            {
                case Paragraph p:
                    var pl = cache ? LayoutOf(p, g.ColumnWidth) : _lines.Layout(p, g.ColumnWidth);
                    if (pl.Content.Hidden) break;
                    y += pl.SpaceBefore + BorderOf(pl, "top");
                    for (int l = 0; l < pl.Lines.Count; l++)
                    {
                        into.Lines.Add(new PlacedLine(pl, l, g.Left, y));
                        y += pl.Lines[l].Height;
                    }
                    y += BorderOf(pl, "bottom") + pl.SpaceAfter;
                    break;
                case Table t:
                    foreach (var row in t.Items.OfType<TableRow>())
                    {
                        var box = MeasureRow(t, row, g.Left, g.ColumnWidth);
                        PlaceRowAt(box, y, into);
                        y += box.Height;
                    }
                    break;
                case BlockContentControl sdt:
                    y = PlaceStory(sdt.Content, y, g, into, cache);
                    break;
            }
        }
        return y;
    }

    /// <summary>Hauteur d'une histoire (en-tête, pied) : paragraphes et rangées de tableau empilés.</summary>
    double StoryHeight(Story story, double width)
    {
        double h = 0;
        ParagraphLayout? prev = null;
        foreach (var b in story.Blocks)
        {
            switch (b)
            {
                case Paragraph p:
                    var pl = LayoutOf(p, width);
                    if (pl.Content.Hidden) break;
                    double before = prev is not null && pl.Format.ContextualSpacing && prev.Format.StyleId == pl.Format.StyleId ? 0 : pl.SpaceBefore;
                    // espacement après du dernier paragraphe compris (Syllabus : corps à 53 + 13,8 + 10 + 0,75 = 77,5 pt
                    // comme dans Word) ; bordures du paragraphe (un en-tête souligné d'un filet)
                    h += before + pl.Height + pl.SpaceAfter + BorderOf(pl, "top") + BorderOf(pl, "bottom");
                    prev = pl;
                    break;
                case Table t:
                    foreach (var row in t.Items.OfType<TableRow>()) h += MeasureRow(t, row, 0, width).Height;
                    prev = null;
                    break;
                case BlockContentControl sdt:
                    h += StoryHeight(sdt.Content, width);
                    break;
            }
        }
        return h;
    }

    ParagraphLayout LayoutOf(Paragraph p, double width, string? tableStyle = null)
    {
        if (_cache.TryGetValue(p, out var l) && Math.Abs(l.Width - width) < 0.01) return l;
        l = _lines.Layout(p, width, tableStyle);
        _cache.AddOrUpdate(p, l);
        LastParagraphsLaidOut++;
        return l;
    }

    /// <summary>Oublie la mise en page des paragraphes de ces blocs (modifiés en place).</summary>
    public void Invalidate(IEnumerable<Block> blocks)
    {
        foreach (var b in blocks)
            foreach (var p in new Story { Blocks = { b } }.Paragraphs()) _cache.Remove(p);
    }

    // ------------------------------------------------------------------ pagination

    /// <summary>Mise en page complète.</summary>
    public void LayoutAll()
    {
        LastParagraphsLaidOut = 0;
        BuildItems();
        var pages = new List<Page>();
        Paginate(new FlowPosition(_items.Count > 0 ? _items[0] : null, 0, 0), pages, null, pageTopFromBreak: true);
        Pages = pages;
        LastPagesLaidOut = pages.Count;
    }

    /// <summary>
    /// Reprise après une modification des blocs <paramref name="changed"/> (modifiés en place, insérés ou remplacés) :
    /// la pagination repart de la page qui précède la première modification et s'arrête dès qu'une page retrouve le
    /// point de départ d'une ancienne page située après toutes les modifications.
    /// </summary>
    public void Update(IReadOnlyCollection<Block> changed, bool rebuildFlow = false)
    {
        LastParagraphsLaidOut = 0;
        Invalidate(changed);
        // rebuildFlow : listes touchées (numéros à recalculer) — pas de chemin rapide
        if (!rebuildFlow && Pages.Count > 0 && changed.Count > 0 && _body.Blocks.Count == _topCount && changed.All(b => b is not Paragraph p || p.Props.Get("sectPr") is null) && changed.All(_itemIndex.ContainsKey))
        {
            UpdateInPlace(changed);
            return;
        }
        var old = Pages;
        var oldIndex = new Dictionary<Block, int>(ReferenceEqualityComparer.Instance);
        for (int i = 0; i < _items.Count; i++) oldIndex[_items[i].Block] = i;
        _oldIndex = oldIndex;
        BuildItems();
        // blocs touchés, contenus des contrôles de contenu compris (les éléments de flux sont leurs blocs intérieurs)
        var changedSet = new HashSet<Block>(ReferenceEqualityComparer.Instance);
        void AddChanged(Block b)
        {
            changedSet.Add(b);
            if (b is BlockContentControl sdt) foreach (var inner in sdt.Content.Blocks) AddChanged(inner);
        }
        foreach (var b in changed) AddChanged(b);
        // premier et dernier éléments touchés (modifiés, nouveaux, ou voisins d'un élément supprimé)
        int first = -1, last = -1;
        for (int i = 0; i < _items.Count; i++)
        {
            var b = _items[i].Block;
            bool touched = changedSet.Contains(b) || !oldIndex.ContainsKey(b) || (_renumbered.Count > 0 && Renumbered(b));
            if (!touched && i > 0 && oldIndex.TryGetValue(_items[i - 1].Block, out var prevOld) && oldIndex.TryGetValue(b, out var curOld) && curOld != prevOld + 1)
                touched = true;   // un élément a disparu juste avant
            if (touched) { if (first < 0) first = i; last = i; }
        }
        if (first < 0 || old.Count == 0) { LayoutAll(); return; }
        // page de départ : celle qui contient l'élément précédant la modification, et une de plus en arrière
        // (un paragraphe solidaire du suivant peut remonter ou descendre)
        var anchor = _items[Math.Max(0, first - 1)].Block;
        int startPage = old.FindIndex(pg => pg.Start.Item is { } st && pg.End.Index >= 0 && Covers(pg, anchor));
        if (startPage < 0) { LayoutAll(); return; }
        startPage = Math.Max(0, startPage - 1);
        var start = old[startPage].Start;
        int startIndex = start.Item is null ? 0 : _items.FindIndex(it => ReferenceEquals(it.Block, start.Item.Block));
        if (startIndex < 0) { LayoutAll(); return; }
        var pages = old.Take(startPage).ToList();
        var lastBlock = _items[last].Block;
        // convergence : nouvelle page qui commence au même endroit qu'une ancienne, après la dernière modification
        var oldStarts = new Dictionary<(Block, int, double), int>();
        for (int k = startPage + 1; k < old.Count; k++)
            if (old[k].Start.Item is { } it) oldStarts.TryAdd((it.Block, old[k].Start.Line, Math.Round(old[k].Start.Offset, 2)), k);
        int lastIndex = last;
        Func<FlowPosition, int?> converge = pos =>
        {
            if (pos.Item is null || pos.Index <= lastIndex) return null;
            return oldStarts.TryGetValue((pos.Item.Block, pos.Line, Math.Round(pos.Offset, 2)), out var k) ? k : null;
        };
        int pagesBefore = pages.Count;
        var resume = Paginate(new FlowPosition(_items[startIndex], startIndex, start.Line, start.Offset), pages, converge,
            pageTopFromBreak: startPage == 0 || old[startPage - 1].End.Item is null);
        LastPagesLaidOut = pages.Count - pagesBefore;
        if (resume is { } k0)
        {
            // pages réutilisées : leurs positions de flux sont réindexées
            var index = new Dictionary<Block, int>(ReferenceEqualityComparer.Instance);
            for (int k = 0; k < _items.Count; k++) index[_items[k].Block] = k;
            for (int k = k0; k < old.Count; k++) pages.Add(Reindex(old[k], index));
        }
        Pages = pages;
    }

    /// <summary>
    /// Chemin rapide de la frappe : les blocs modifiés sont déjà des éléments du flux (pas d'insertion ni de
    /// suppression), les indices ne changent pas ; les anciennes pages sont retrouvées par recherche dichotomique.
    /// </summary>
    void UpdateInPlace(IReadOnlyCollection<Block> changed)
    {
        var old = Pages;
        int first = changed.Min(b => _itemIndex[b]), last = changed.Max(b => _itemIndex[b]);
        // page qui contient l'élément précédant la modification, et une de plus en arrière (paragraphes solidaires)
        int anchor = Math.Max(0, first - 1);
        int startPage = Math.Max(0, PageOf(old, anchor) - 1);
        var start = old[startPage].Start;
        var pages = old.GetRange(0, startPage);
        Func<FlowPosition, int?> converge = pos =>
        {
            if (pos.Item is null || pos.Index <= last) return null;
            // ancienne page qui commence exactement au même endroit (même élément, même ligne)
            int k = PageOf(old, pos.Index);
            for (; k < old.Count && old[k].Start.Index <= pos.Index; k++)
                if (old[k].Start.Index == pos.Index && old[k].Start.Line == pos.Line && Math.Abs(old[k].Start.Offset - pos.Offset) < 0.005 && k > startPage) return k;
            return null;
        };
        var resume = Paginate(start.Item is null ? start : new FlowPosition(_items[start.Index], start.Index, start.Line, start.Offset), pages, converge,
            pageTopFromBreak: startPage == 0 || old[startPage - 1].End.Item is null);
        LastPagesLaidOut = pages.Count - startPage;
        if (resume is { } k0) pages.AddRange(old.Skip(k0));
        Pages = pages;
    }

    /// <summary>Première page dont le départ est au plus l'élément <paramref name="index"/> et qui le contient.</summary>
    static int PageOf(List<Page> pages, int index)
    {
        int lo = 0, hi = pages.Count - 1;
        while (lo < hi)
        {
            int mid = (lo + hi + 1) / 2;
            if (pages[mid].Start.Index <= index && pages[mid].Start.Item is not null) lo = mid; else hi = mid - 1;
        }
        // l'élément peut commencer sur la page précédente (paragraphe coupé) : la première page qui le contient
        while (lo > 0 && pages[lo].Start.Index == index && pages[lo].Start.Line > 0) lo--;
        return lo;
    }

    bool Renumbered(Block b) => b is Paragraph p ? _renumbered.Contains(p) : new Story { Blocks = { b } }.Paragraphs().Any(_renumbered.Contains);

    /// <summary>La page contient (au moins en partie) l'élément de flux <paramref name="b"/>.</summary>
    bool Covers(Page pg, Block b) => _oldIndex.TryGetValue(b, out var i) && pg.Start.Index <= i && (pg.End.Item is null || i <= pg.End.Index);

    Dictionary<Block, int> _oldIndex = new(ReferenceEqualityComparer.Instance);

    static Page Reindex(Page p, Dictionary<Block, int> index)
    {
        FlowPosition Fix(FlowPosition f) => f.Item is not null && index.TryGetValue(f.Item.Block, out var i) ? f with { Index = i } : f;
        var copy = new Page { Geometry = p.Geometry, Start = Fix(p.Start), End = Fix(p.End) };
        copy.Lines.AddRange(p.Lines);
        copy.Cells.AddRange(p.Cells);
        copy.NoteIds.AddRange(p.NoteIds);
        copy.ReservedBottom = p.ReservedBottom;
        copy.NotesCount = p.NotesCount;
        return copy;
    }

    /// <summary>État de la pagination : page en cours, position verticale, paragraphe précédent.</summary>
    sealed class Cursor
    {
        public required Page Page;
        public required SectionGeometry Geo;
        public double Y;
        public bool Empty = true;
        /// <summary>La page a commencé par un saut explicite (ou en tête de document / de section).</summary>
        public bool TopFromBreak;
        public ParagraphLayout? Prev;
    }

    /// <summary>
    /// Pagine depuis <paramref name="from"/> en ajoutant les pages à <paramref name="pages"/>. Si <paramref name="converge"/>
    /// reconnaît le départ d'une nouvelle page, s'arrête et renvoie l'indice de l'ancienne page à réutiliser.
    /// </summary>
    int? Paginate(FlowPosition from, List<Page> pages, Func<FlowPosition, int?>? converge, bool pageTopFromBreak)
    {
        if (from.Item is null)
        {
            pages.Add(new Page { Geometry = Sections.Count > 0 ? Sections[^1] : SectionGeometry.From(null), Start = from, End = from });
            return null;
        }
        int i = from.Index;
        var geo0 = Sections[_items[i].Section];
        bool FirstOf(int item) => item == 0 || _items[item - 1].Section != _items[item].Section;
        var first0 = from.Line == 0 && FirstOf(i) && (pages.Count == 0 || pages[^1].Geometry != geo0 || pages[^1].End.Item is null);
        var page0 = new Page { Geometry = geo0, Start = from, FirstOfSection = first0 };
        var c = new Cursor { Geo = geo0, Page = page0, Y = page0.ContentTop, TopFromBreak = pageTopFromBreak };

        // page suivante ; renvoie l'indice de l'ancienne page si la mise en page converge
        int? NewPage(int item, int line, bool fromBreak, double offset = 0)
        {
            var pos = new FlowPosition(_items[item], item, line, offset);
            c.Page.End = pos;
            pages.Add(c.Page);
            if (converge?.Invoke(pos) is { } k) return k;
            // première page de la section : la section change, ou la page commence par son premier élément
            bool first = line == 0 && FirstOf(item);
            c.Geo = Sections[_items[item].Section];
            c.Page = new Page { Geometry = c.Geo, Start = pos, FirstOfSection = first };
            c.Y = c.Page.ContentTop;
            c.Empty = true;
            c.TopFromBreak = fromBreak;
            c.Prev = null;
            return null;
        }

        for (int startLine = from.Line; i < _items.Count; i++, startLine = 0)
        {
            var item = _items[i];
            // changement de section : nouvelle page, sauf section continue de même format
            if (startLine == 0 && i > 0 && _items[i - 1].Section != item.Section)
            {
                var g = Sections[item.Section];
                bool continuous = g.Type == "continuous" && g.PageWidth == c.Geo.PageWidth && g.PageHeight == c.Geo.PageHeight;
                if (!continuous && !c.Empty) { if (NewPage(i, 0, fromBreak: true) is { } k) return k; }
                else c.Geo = g;
            }
            int? conv = item.Block switch
            {
                Table t => PlaceTable(t, i, startLine, i == from.Index ? from.Offset : 0, c, NewPage),
                Paragraph p => PlaceParagraph(p, i, startLine, c, NewPage),
                _ => null,
            };
            if (conv is not null) return conv;
        }
        c.Page.End = new FlowPosition(null, _items.Count, 0);
        pages.Add(c.Page);
        return null;
    }

    delegate int? PageBreaker(int item, int line, bool fromBreak, double offset = 0);

    int? PlaceParagraph(Paragraph p, int i, int startLine, Cursor c, PageBreaker newPage)
    {
        var pl = LayoutOf(p, c.Geo.ColumnWidth);
        if (pl.Content.Hidden) return null;
        var pf = pl.Format;
        int n = pl.Lines.Count;
        double SpaceAtTop() => SuppressSpaceBeforeAtPageTop ? 0 : pl.SpaceBefore;
        if (startLine == 0 && pf.PageBreakBefore && !c.Empty)
            if (newPage(i, 0, fromBreak: true) is { } k) return k;
        // espacement avant : supprimé entre paragraphes de même style avec « contextual spacing », et en haut
        // d'une page commencée par le flux
        double before = startLine > 0 ? 0 : pl.SpaceBefore;
        if (c.Prev is not null && pf.ContextualSpacing && c.Prev.Format.StyleId == pf.StyleId) before = 0;
        if (c.Empty && !c.TopFromBreak && SuppressSpaceBeforeAtPageTop) before = 0;
        double usable = c.Page.ContentBottom - c.Page.ContentTop;
        if (startLine == 0 && !c.Empty)
        {
            // lignes solidaires : le paragraphe entier passe à la page suivante s'il tient sur une page
            bool move = pf.KeepLines && c.Y + before + pl.Height > c.Page.ContentBottom + 0.01 && pl.Height <= usable;
            // solidaire du suivant : la chaîne doit tenir avec les premières lignes du paragraphe suivant
            if (!move && pf.KeepNext)
            {
                double need = before + KeepChainHeight(i, c.Geo);
                move = c.Y + need > c.Page.ContentBottom + 0.01 && need <= usable;
            }
            if (move)
            {
                if (newPage(i, 0, fromBreak: false) is { } k) return k;
                before = SpaceAtTop();
            }
        }
        // bordure de paragraphe : en tête d'un groupe de paragraphes aux bordures identiques, la bordure du haut et
        // son espacement (w:space, points) s'ajoutent au-dessus de la première ligne ; en fin de groupe, en dessous
        var (borderTop, borderBottom) = BorderExtents(pl, i);
        c.Y += before + (startLine == 0 ? borderTop : 0);
        int line = startLine;
        while (line < n)
        {
            var lb = pl.Lines[line];
            // une ligne qui ne contient que le saut de page reste en bas de la page : elle ne crée pas de page blanche
            // (Syllabus : titre « Introduction » précédé d'un saut, gardé en page 1 par Word)
            bool breakOnly = lb.EndKind is LineEnd.PageBreak && lb.End == lb.Start;
            double notes = NotesHeight(lb, c.Page, c.Geo.ColumnWidth);
            if (c.Y + lb.FitHeight > c.Page.ContentBottom - notes + BottomTolerance && !c.Empty && !breakOnly)
            {
                if (Trace) Console.WriteLine($"    coupure : élément {i} ligne {line}/{n}, y {c.Y:0.0} + {lb.FitHeight:0.0} (ligne {lb.Height:0.0}, avant {before:0.0}) > bas {c.Page.ContentBottom:0.0}");
                int here = line - startLine;           // lignes du paragraphe déjà sur cette page
                int cut = line;
                if (pf.WidowControl && n > 1 && here > 0)
                {
                    if (n - cut == 1) cut--;                                   // veuve : une ligne de plus descend
                    if (cut - startLine == 1 && startLine == 0) cut = 0;       // orpheline : tout descend
                }
                int remove = line - cut;
                if (remove > 0)
                {
                    // les notes appelées sur les lignes qui descendent descendent avec elles
                    foreach (var gone in c.Page.Lines.GetRange(c.Page.Lines.Count - remove, remove).Where(g => g.Notes > 0))
                    {
                        c.Page.ReservedBottom -= gone.Notes;
                        c.Page.NotesCount -= gone.Paragraph.Lines[gone.Line].Footnotes?.Count ?? 0;
                        foreach (var id in gone.Paragraph.Lines[gone.Line].Footnotes ?? new()) c.Page.NoteIds.Remove(id);
                    }
                    c.Page.Lines.RemoveRange(c.Page.Lines.Count - remove, remove);
                }
                if (newPage(i, cut, fromBreak: false) is { } k) return k;
                if (cut == 0) c.Y += SpaceAtTop();
                line = cut;
                startLine = cut;
                continue;
            }
            c.Page.Lines.Add(new PlacedLine(pl, line, c.Geo.Left, c.Y) { Notes = notes });
            if (notes > 0)
            {
                c.Page.ReservedBottom += notes;
                c.Page.NotesCount += lb.Footnotes!.Count;
                c.Page.NoteIds.AddRange(lb.Footnotes!);
            }
            c.Y += lb.Height;
            c.Empty = false;
            line++;
            // saut de page dans le paragraphe : la suite (au moins la marque de paragraphe) va à la page suivante
            if (lb.EndKind is LineEnd.PageBreak or LineEnd.ColumnBreak && line < n)
            {
                if (newPage(i, line, fromBreak: true) is { } k) return k;
                startLine = line;
            }
        }
        c.Y += borderBottom + pl.SpaceAfter;
        c.Prev = pl;
        // saut de page à la fin du paragraphe : l'élément suivant commence sur une nouvelle page
        if (n > 0 && pl.Lines[^1].EndKind is LineEnd.PageBreak or LineEnd.ColumnBreak && i + 1 < _items.Count)
            if (newPage(i + 1, 0, fromBreak: true) is { } k) return k;
        return null;
    }

    /// <summary>
    /// Place prise par les bordures de paragraphe au-dessus et au-dessous du paragraphe <paramref name="i"/> : un groupe de
    /// paragraphes consécutifs aux bordures identiques (w:pBdr) forme une seule boîte, la bordure du haut ne compte
    /// qu'au premier, celle du bas qu'au dernier. Épaisseur w:sz en huitièmes de point, espacement w:space en points.
    /// </summary>
    (double Top, double Bottom) BorderExtents(ParagraphLayout pl, int i)
    {
        var W = pl.Format.Effective.Name.Namespace;
        var bdr = pl.Format.Effective.Element(W + "pBdr");
        if (bdr is null) return (0, 0);
        XElement? Neighbour(int k) => k >= 0 && k < _items.Count && _items[k].Block is Paragraph np && _items[k].Section == _items[i].Section
            ? LayoutOf(np, Sections[_items[k].Section].ColumnWidth).Format.Effective.Element(W + "pBdr") : null;
        bool SameAs(XElement? other) => other is not null && XNode.DeepEquals(other, bdr);
        double top = SameAs(Neighbour(i - 1)) ? 0 : BorderExtent(bdr.Element(W + "top"), W);
        double bottom = SameAs(Neighbour(i + 1)) ? 0 : BorderExtent(bdr.Element(W + "bottom"), W);
        return (top, bottom);
    }

    /// <summary>Hauteur de la chaîne de paragraphes solidaires qui commence en <paramref name="i"/>, avec les premières lignes du suivant.</summary>
    double KeepChainHeight(int i, SectionGeometry geo)
    {
        double h = 0;
        for (int k = i, guard = 0; k < _items.Count && guard < 8; k++, guard++)
        {
            if (_items[k].Block is not Paragraph p) break;
            var pl = LayoutOf(p, geo.ColumnWidth);
            if (k > i) h += pl.SpaceBefore;
            if (pl.Format.KeepNext && k + 1 < _items.Count && _items[k + 1].Section == _items[k].Section)
            {
                h += pl.Height + pl.SpaceAfter;
                continue;
            }
            if (k == i) return pl.Height;
            // dernier de la chaîne : sa première ligne (deux si veuves et orphelines, ou tout s'il a des lignes solidaires)
            int take = pl.Format.KeepLines ? pl.Lines.Count : Math.Min(pl.Lines.Count, pl.Format.WidowControl ? 2 : 1);
            for (int l = 0; l < take; l++) h += pl.Lines[l].Height;
            break;
        }
        return h;
    }

    /// <summary>Rangée de tableau mesurée : hauteur, et lignes des cellules (abscisse de la cellule, lignes et leur décalage vertical).</summary>
    sealed record RowBox(double Height, List<(double X, double W, List<(ParagraphLayout L, double Y)> Lines)> Cells)
    {
        /// <summary>Fond et bordures de chaque cellule (même ordre que <see cref="Cells"/>).</summary>
        public List<(string? Fill, double Top, double Bottom, double Left, double Right)> Looks { get; init; } = new();
        /// <summary>Hauteur sans les bordures horizontales.</summary>
        public double Content { get; init; }
    }

    const double CellPadding = 5.4;     // marges gauche et droite des cellules par défaut (108 twips)

    static double Twips(string? v) => double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? d / 20 : 0;

    /// <summary>Mesure une rangée : largeurs de la grille (w:gridCol, w:gridSpan), contenu des cellules, w:trHeight.</summary>
    /// <summary>
    /// Épaisseur (points) de la bordure horizontale d'une rangée : la plus épaisse de ses cellules (w:tcBorders), sinon
    /// celle du tableau (w:insideH entre deux rangées, w:top / w:bottom aux bords ; style de tableau compris).
    /// </summary>
    double RowBorder(Table t, TableRow row, string side, bool edge)
    {
        var W = t.Props.Ns;
        static double Width(XElement? b, XNamespace W) =>
            b is null || b.Attribute(W + "val")?.Value is null or "none" or "nil" ? 0
            : double.TryParse(b.Attribute(W + "sz")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v / 8 : 0.5;
        var tableBorders = t.Props.Get("tblBorders") ?? TableStyleBorders(t);
        double table = Width(tableBorders?.Element(W + (edge ? side : "insideH")), W);
        double cells = 0;
        bool any = false;
        foreach (var cell in row.Items.OfType<TableCell>())
            if (cell.Props.Get("tcBorders")?.Element(W + side) is { } cb) { cells = Math.Max(cells, Width(cb, W)); any = true; }
        return any ? Math.Max(cells, table) : table;
    }

    XElement? TableStyleBorders(Table t)
    {
        var W = t.Props.Ns;
        var sheet = _lines.Styles.Sheet;
        foreach (var st in sheet.Chain(t.Props.Val("tblStyle") ?? sheet.DefaultStyle("table")?.Attribute(W + "styleId")?.Value))
            if (st.Element(W + "tblPr")?.Element(W + "tblBorders") is { } b) return b;
        return null;
    }

    /// <summary>Fond (w:shd/@fill) et bordures d'une cellule : w:tcBorders, sinon celles du tableau (bords et intérieures confondus).</summary>
    static (string?, double, double, double, double) CellLook(Table t, TableCell cell, XElement? tableBorders, bool continued)
    {
        var W = t.Props.Ns;
        var fill = cell.Props.Get("shd")?.Attribute(W + "fill")?.Value;
        if (fill is null or "auto" || fill.Length != 6) fill = null;
        var own = cell.Props.Get("tcBorders");
        double Side(string side, string inner)
        {
            var b = own?.Element(W + side) ?? tableBorders?.Element(W + side) ?? tableBorders?.Element(W + inner);
            if (b is null || b.Attribute(W + "val")?.Value is null or "none" or "nil") return 0;
            return double.TryParse(b.Attribute(W + "sz")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? Math.Max(0.25, v / 8) : 0.5;
        }
        return (fill, continued ? 0 : Side("top", "insideH"), Side("bottom", "insideH"), Side("left", "insideV"), Side("right", "insideV"));
    }

    RowBox MeasureRow(Table t, TableRow row, double left, double columnWidth)
    {
        var W = t.Props.Ns;
        var cols = t.Grid?.Elements(W + "gridCol").Select(g => Twips(g.Attribute(W + "w")?.Value)).ToList() ?? new List<double>();
        var style = t.Props.Val("tblStyle");
        double x = left + Twips(t.Props.Get("tblInd")?.Attribute(W + "w")?.Value) - CellPadding;
        var cells = row.Items.OfType<TableCell>().ToList();
        var placed = new List<(double, double, List<(ParagraphLayout, double)>)>();
        var looks = new List<(string?, double, double, double, double)>();
        var tableBorders = t.Props.Get("tblBorders") ?? TableStyleBorders(t);
        int col = 0;
        double rowH = 0;
        foreach (var cell in cells)
        {
            int span = int.TryParse(cell.Props.Val("gridSpan"), out var s) ? s : 1;
            double w = 0;
            for (int k = 0; k < span; k++)
                w += col + k < cols.Count ? cols[col + k] : columnWidth / Math.Max(1, cells.Count);
            col += span;
            double cy = 0;
            var lines = new List<(ParagraphLayout, double)>();
            ParagraphLayout? prev = null;
            foreach (var b in cell.Content.Paragraphs())
            {
                var pl = LayoutOf(b, Math.Max(1, w - 2 * CellPadding), style);
                if (pl.Content.Hidden) continue;
                double before = pl.SpaceBefore;
                if (prev is not null && pl.Format.ContextualSpacing && prev.Format.StyleId == pl.Format.StyleId) before = 0;
                cy += before;
                lines.Add((pl, cy));
                cy += pl.Height + pl.SpaceAfter;
                prev = pl;
                if (Trace && pl.Lines.Count > 1)
                    Console.WriteLine($"    cellule {w:0.0} pt : « {pl.Content.Text[..Math.Min(40, pl.Content.Text.Length)]} » {pl.Lines.Count} lignes : "
                        + string.Join(" | ", pl.Lines.Select(l => pl.Content.Text[l.Start..l.End])));
            }
            // fusion verticale (w:vMerge sans valeur = suite) : la cellule n'impose pas de hauteur
            bool continued = cell.Props.Get("vMerge") is { } vm && vm.Attribute(W + "val")?.Value is null or "continue";
            if (continued) cy = 0;
            placed.Add((x, w, lines));
            looks.Add(CellLook(t, cell, tableBorders, continued));
            rowH = Math.Max(rowH, cy);
            x += w;
        }
        // hauteur imposée (w:trHeight : au moins, ou exacte)
        if (row.Props.Get("trHeight") is { } th)
        {
            double v = Twips(th.Attribute(W + "val")?.Value);
            rowH = th.Attribute(W + "hRule")?.Value == "exact" ? v : Math.Max(rowH, v);
        }
        // bordures horizontales : celle du bas de chaque rangée, et celle du haut pour la première (Syllabus : rangée
        // à bordure basse de 3 pt plus haute de 2,8 pt dans Word, filet de 0,5 pt : 0,7 pt)
        var rows = t.Items.OfType<TableRow>().ToList();
        int index = rows.IndexOf(row);
        // bordure partagée entre deux rangées : la plus épaisse des deux (bas de la précédente, haut de celle-ci)
        double top = index == 0 ? RowBorder(t, row, "top", edge: true)
            : Math.Max(0, RowBorder(t, row, "top", edge: false) - RowBorder(t, rows[index - 1], "bottom", edge: false));
        double borders = RowBorder(t, row, "bottom", edge: index == rows.Count - 1) + top;
        return new RowBox(rowH + borders, placed) { Content = rowH, Looks = looks };
    }

    /// <summary>Tableau : rangées entières, une rangée trop haute pour la place restante passe à la page suivante.</summary>
    int? PlaceTable(Table t, int i, int startRow, double startOffset, Cursor c, PageBreaker newPage)
    {
        var rows = t.Items.OfType<TableRow>().ToList();
        // tableau flottant (w:tblpPr) : il ne prend pas de place dans le flux ; ancré en bas de la marge, il réduit
        // la hauteur du corps de la page (Syllabus, page de garde) ; les autres positions viennent avec le lot 7
        if (t.Props.Get("tblpPr") is { } pos)
        {
            var W = t.Props.Ns;
            double height = rows.Sum(r => MeasureRow(t, r, c.Geo.Left, c.Geo.ColumnWidth).Height);
            if (pos.Attribute(W + "tblpYSpec")?.Value == "bottom" && (pos.Attribute(W + "vertAnchor")?.Value ?? "margin") == "margin")
            {
                double y = c.Page.ContentBottom - height;
                foreach (var r in rows)
                {
                    var box = MeasureRow(t, r, c.Geo.Left, c.Geo.ColumnWidth);
                    PlaceRowAt(box, y, c.Page);
                    y += box.Height;
                }
                c.Page.ReservedBottom += height + Twips(pos.Attribute(W + "bottomFromText")?.Value);
            }
            return null;
        }
        // rangées d'en-tête (w:tblHeader en tête du tableau) : répétées en haut de chaque page où le tableau continue
        int headers = 0;
        while (headers < rows.Count && rows[headers].Props.Get("tblHeader") is { } th && th.Attribute(t.Props.Ns + "val")?.Value is null or "1" or "true" or "on") headers++;
        void RepeatHeaders()
        {
            for (int h = 0; h < headers && h < rows.Count; h++)
            {
                var hb = MeasureRow(t, rows[h], c.Geo.Left, c.Geo.ColumnWidth);
                PlaceRowAt(hb, c.Y, c.Page);
                c.Y += hb.Height;
            }
        }
        if (startRow > 0 || startOffset > 0)
            if (startRow >= headers) RepeatHeaders();
        double offset = startOffset;
        for (int r = startRow; r < rows.Count; r++, offset = 0)
        {
            var box = MeasureRow(t, rows[r], c.Geo.Left, c.Geo.ColumnWidth);
            while (true)
            {
                double remaining = RowPartHeight(box, offset, null);
                if (c.Y + remaining <= c.Page.ContentBottom + 0.01 || c.Empty && !CanSplit(rows[r], t))
                {
                    PlaceRowPart(box, offset, null, c.Y, c.Page);
                    c.Y += remaining;
                    c.Empty = false;
                    break;
                }
                // rangée trop haute pour la place restante : coupée entre les lignes de ses cellules (sauf w:cantSplit),
                // sinon renvoyée entière à la page suivante
                double cut = offset + (c.Page.ContentBottom - c.Y);
                if (CanSplit(rows[r], t) && r >= headers && HasLinesBetween(box, offset, cut))
                {
                    double placed = RowPartHeight(box, offset, cut);
                    PlaceRowPart(box, offset, cut, c.Y, c.Page);
                    c.Y += placed;
                    c.Empty = false;
                    double next = LastBottomBefore(box, offset, cut);
                    if (newPage(i, r, fromBreak: false, next) is { } k) return k;
                    RepeatHeaders();
                    offset = next;
                    continue;
                }
                if (c.Empty)
                {
                    // ni coupable ni entière : placée telle quelle (elle dépasse le bas de la page)
                    PlaceRowPart(box, offset, null, c.Y, c.Page);
                    c.Y += remaining;
                    c.Empty = false;
                    break;
                }
                if (newPage(i, r, fromBreak: false, offset) is { } k2) return k2;
                RepeatHeaders();
            }
        }
        c.Prev = null;
        return null;
    }

    static bool CanSplit(TableRow row, Table t) =>
        (row.Props.Get("cantSplit") is not { } cs || cs.Attribute(t.Props.Ns + "val")?.Value is "0" or "false" or "off")
        // une rangée à hauteur minimale (w:trHeight) ne se coupe pas (Syllabus : rangée de 146 pt envoyée entière à la
        // page suivante par Word alors que 120 pt restaient ; écart ramené de -137 à -3 pt sur la page suivante)
        && row.Props.Get("trHeight") is null;

    /// <summary>Lignes des cellules d'une rangée : cellule, paragraphe, ligne, haut et bas dans la rangée.</summary>
    static IEnumerable<(int Cell, double X, ParagraphLayout L, int Line, double Top, double Bottom)> RowLines(RowBox box)
    {
        for (int ci = 0; ci < box.Cells.Count; ci++)
        {
            var (cx, _, lines) = box.Cells[ci];
            foreach (var (pl, ly) in lines)
            {
                double y = ly;
                for (int l = 0; l < pl.Lines.Count; l++)
                {
                    yield return (ci, cx, pl, l, y, y + pl.Lines[l].Height);
                    y += pl.Lines[l].Height;
                }
            }
        }
    }

    /// <summary>Haut de la première ligne restante de chaque cellule (lignes dont le bas dépasse <paramref name="from"/>).</summary>
    static double[] CellShift(RowBox box, double from)
    {
        var shift = Enumerable.Repeat(double.NaN, box.Cells.Count).ToArray();
        foreach (var x in RowLines(box))
            if (x.Bottom > from + 0.01 && double.IsNaN(shift[x.Cell])) shift[x.Cell] = x.Top;
        return shift;
    }

    static bool HasLinesBetween(RowBox box, double from, double to) =>
        RowLines(box).Any(x => x.Bottom > from + 0.01 && x.Bottom <= to + 0.01);

    /// <summary>Bas de la dernière ligne placée avant la coupure (point de reprise sur la page suivante).</summary>
    static double LastBottomBefore(RowBox box, double from, double to)
    {
        // reprise : tout ce qui est sous la coupure ; on mémorise la coupure elle-même (les lignes dont le bas la
        // dépasse continuent)
        return to;
    }

    /// <summary>Hauteur de la partie [from, to[ de la rangée (to null : jusqu'au bout, bordure et hauteur minimale comprises).</summary>
    static double RowPartHeight(RowBox box, double from, double? to)
    {
        if (from <= 0 && to is null) return box.Height;
        var shift = CellShift(box, from);
        double h = 0;
        foreach (var x in RowLines(box))
        {
            if (x.Bottom <= from + 0.01 || to is { } t && x.Bottom > t + 0.01) continue;
            h = Math.Max(h, x.Bottom - shift[x.Cell]);
        }
        if (to is null) h += box.Height - box.Content;     // bordures de la rangée, à la fin de la dernière partie
        return h;
    }

    static void PlaceRowPart(RowBox box, double from, double? to, double y, Page page)
    {
        if (from <= 0 && to is null) { PlaceRowAt(box, y, page); return; }
        var shift = CellShift(box, from);
        double h = RowPartHeight(box, from, to);
        for (int ci = 0; ci < box.Cells.Count; ci++) page.Cells.Add(Cell(box, ci, y, h));
        foreach (var x in RowLines(box))
        {
            if (x.Bottom <= from + 0.01 || to is { } t && x.Bottom > t + 0.01) continue;
            page.Lines.Add(new PlacedLine(x.L, x.Line, x.X + CellPadding, y + x.Top - shift[x.Cell]));
        }
    }

    static CellBox Cell(RowBox box, int ci, double y, double h)
    {
        var (cx, cw, _) = box.Cells[ci];
        var (fill, top, bottom, left, right) = ci < box.Looks.Count ? box.Looks[ci] : (null, 0, 0, 0, 0);
        return new CellBox(cx, y, cw, h, fill, top, bottom, left, right);
    }

    static void PlaceRowAt(RowBox box, double y, Page page)
    {
        for (int ci = 0; ci < box.Cells.Count; ci++)
        {
            var (cx, cw, lines) = box.Cells[ci];
            page.Cells.Add(Cell(box, ci, y, box.Height));
            foreach (var (pl, ly) in lines)
            {
                double yy = y + ly;
                for (int l = 0; l < pl.Lines.Count; l++)
                {
                    page.Lines.Add(new PlacedLine(pl, l, cx + CellPadding, yy));
                    yy += pl.Lines[l].Height;
                }
            }
        }
    }
}
