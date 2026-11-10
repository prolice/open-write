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
            HeaderDistance = Get("pgMar", "header", 708), FooterDistance = Get("pgMar", "footer", 708),
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
public readonly record struct PlacedLine(ParagraphLayout Paragraph, int Line, double X, double Y);

/// <summary>Point de départ d'une page : élément de flux et ligne (ou rangée de tableau) où elle commence.</summary>
public readonly record struct FlowPosition(FlowItem? Item, int Index, int Line);

public sealed class Page
{
    public required SectionGeometry Geometry { get; init; }
    /// <summary>Première page de sa section (en-tête « première page » si w:titlePg).</summary>
    public bool FirstOfSection { get; init; }
    public double ContentTop => FirstOfSection ? Geometry.FirstTop : Geometry.Top;
    public double ContentBottom => Geometry.PageHeight - (FirstOfSection ? Geometry.FirstBottom : Geometry.Bottom);
    public required FlowPosition Start { get; init; }
    public List<PlacedLine> Lines { get; } = new();
    /// <summary>Bordures des cellules de tableau (rectangles en points), pour l'affichage.</summary>
    public List<(double X, double Y, double W, double H)> Cells { get; } = new();
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
    public List<Page> Pages { get; private set; } = new();
    public IReadOnlyList<FlowItem> Items => _items;
    /// <summary>Statistiques de la dernière mise en page : paragraphes recalculés, pages recalculées.</summary>
    public int LastParagraphsLaidOut { get; private set; }
    public int LastPagesLaidOut { get; private set; }

    readonly Func<string, Story?>? _related;

    /// <param name="related">En-têtes et pieds de page : histoire d'une partie liée par son r:id (null : ignorés).</param>
    public DocumentLayout(Story body, StyleResolver styles, IFontProvider fonts, Func<string, Story?>? related = null)
    {
        _body = body;
        _lines = new ParagraphLayouter(styles, fonts);
        _related = related;
    }

    public ParagraphLayouter Lines => _lines;

    // ------------------------------------------------------------------ flux

    void BuildItems()
    {
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

    /// <summary>Hypothèse à calibrer : l'espacement après du dernier paragraphe d'un en-tête ou d'un pied ne compte pas.</summary>
    public bool TrailingSpaceExcluded { get; set; } = Environment.GetEnvironmentVariable("OPENWRITE_HF_TRAILING") != "1";

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
                    h += before + pl.Height + (ReferenceEquals(b, story.Blocks[^1]) && TrailingSpaceExcluded ? 0 : pl.SpaceAfter);
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
    public void Update(IReadOnlyCollection<Block> changed)
    {
        LastParagraphsLaidOut = 0;
        Invalidate(changed);
        if (Pages.Count > 0 && changed.Count > 0 && _body.Blocks.Count == _topCount && changed.All(b => b is not Paragraph p || p.Props.Get("sectPr") is null) && changed.All(_itemIndex.ContainsKey))
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
            bool touched = changedSet.Contains(b) || !oldIndex.ContainsKey(b);
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
        var oldStarts = new Dictionary<(Block, int), int>();
        for (int k = startPage + 1; k < old.Count; k++)
            if (old[k].Start.Item is { } it) oldStarts.TryAdd((it.Block, old[k].Start.Line), k);
        int lastIndex = last;
        Func<FlowPosition, int?> converge = pos =>
        {
            if (pos.Item is null || pos.Index <= lastIndex) return null;
            return oldStarts.TryGetValue((pos.Item.Block, pos.Line), out var k) ? k : null;
        };
        int pagesBefore = pages.Count;
        var resume = Paginate(new FlowPosition(_items[startIndex], startIndex, start.Line), pages, converge,
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
                if (old[k].Start.Index == pos.Index && old[k].Start.Line == pos.Line && k > startPage) return k;
            return null;
        };
        var resume = Paginate(start.Item is null ? start : new FlowPosition(_items[start.Index], start.Index, start.Line), pages, converge,
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

    /// <summary>La page contient (au moins en partie) l'élément de flux <paramref name="b"/>.</summary>
    bool Covers(Page pg, Block b) => _oldIndex.TryGetValue(b, out var i) && pg.Start.Index <= i && (pg.End.Item is null || i <= pg.End.Index);

    Dictionary<Block, int> _oldIndex = new(ReferenceEqualityComparer.Instance);

    static Page Reindex(Page p, Dictionary<Block, int> index)
    {
        FlowPosition Fix(FlowPosition f) => f.Item is not null && index.TryGetValue(f.Item.Block, out var i) ? f with { Index = i } : f;
        var copy = new Page { Geometry = p.Geometry, Start = Fix(p.Start), End = Fix(p.End) };
        copy.Lines.AddRange(p.Lines);
        copy.Cells.AddRange(p.Cells);
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
        int? NewPage(int item, int line, bool fromBreak)
        {
            var pos = new FlowPosition(_items[item], item, line);
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
                Table t => PlaceTable(t, i, startLine, c, NewPage),
                Paragraph p => PlaceParagraph(p, i, startLine, c, NewPage),
                _ => null,
            };
            if (conv is not null) return conv;
        }
        c.Page.End = new FlowPosition(null, _items.Count, 0);
        pages.Add(c.Page);
        return null;
    }

    delegate int? PageBreaker(int item, int line, bool fromBreak);

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
            if (c.Y + lb.Height > c.Page.ContentBottom + 0.01 && !c.Empty)
            {
                int here = line - startLine;           // lignes du paragraphe déjà sur cette page
                int cut = line;
                if (pf.WidowControl && n > 1 && here > 0)
                {
                    if (n - cut == 1) cut--;                                   // veuve : une ligne de plus descend
                    if (cut - startLine == 1 && startLine == 0) cut = 0;       // orpheline : tout descend
                }
                int remove = line - cut;
                if (remove > 0) c.Page.Lines.RemoveRange(c.Page.Lines.Count - remove, remove);
                if (newPage(i, cut, fromBreak: false) is { } k) return k;
                if (cut == 0) c.Y += SpaceAtTop();
                line = cut;
                startLine = cut;
                continue;
            }
            c.Page.Lines.Add(new PlacedLine(pl, line, c.Geo.Left, c.Y));
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
        static double Extent(XElement? b, XNamespace W)
        {
            if (b is null || b.Attribute(W + "val")?.Value is null or "none" or "nil") return 0;
            double sz = double.TryParse(b.Attribute(W + "sz")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v / 8 : 0.5;
            double space = double.TryParse(b.Attribute(W + "space")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var sp) ? sp : 0;
            return sz + space;
        }
        XElement? Neighbour(int k) => k >= 0 && k < _items.Count && _items[k].Block is Paragraph np && _items[k].Section == _items[i].Section
            ? LayoutOf(np, Sections[_items[k].Section].ColumnWidth).Format.Effective.Element(W + "pBdr") : null;
        bool SameAs(XElement? other) => other is not null && XNode.DeepEquals(other, bdr);
        double top = SameAs(Neighbour(i - 1)) ? 0 : Extent(bdr.Element(W + "top"), W);
        double bottom = SameAs(Neighbour(i + 1)) ? 0 : Extent(bdr.Element(W + "bottom"), W);
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
    sealed record RowBox(double Height, List<(double X, double W, List<(ParagraphLayout L, double Y)> Lines)> Cells);

    const double CellPadding = 5.4;     // marges gauche et droite des cellules par défaut (108 twips)

    static double Twips(string? v) => double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? d / 20 : 0;

    /// <summary>Mesure une rangée : largeurs de la grille (w:gridCol, w:gridSpan), contenu des cellules, w:trHeight.</summary>
    RowBox MeasureRow(Table t, TableRow row, double left, double columnWidth)
    {
        var W = t.Props.Ns;
        var cols = t.Grid?.Elements(W + "gridCol").Select(g => Twips(g.Attribute(W + "w")?.Value)).ToList() ?? new List<double>();
        var style = t.Props.Val("tblStyle");
        double x = left + Twips(t.Props.Get("tblInd")?.Attribute(W + "w")?.Value) - CellPadding;
        var cells = row.Items.OfType<TableCell>().ToList();
        var placed = new List<(double, double, List<(ParagraphLayout, double)>)>();
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
            }
            // fusion verticale (w:vMerge sans valeur = suite) : la cellule n'impose pas de hauteur
            if (cell.Props.Get("vMerge") is { } vm && vm.Attribute(W + "val")?.Value is null or "continue") cy = 0;
            placed.Add((x, w, lines));
            rowH = Math.Max(rowH, cy);
            x += w;
        }
        // hauteur imposée (w:trHeight : au moins, ou exacte)
        if (row.Props.Get("trHeight") is { } th)
        {
            double v = Twips(th.Attribute(W + "val")?.Value);
            rowH = th.Attribute(W + "hRule")?.Value == "exact" ? v : Math.Max(rowH, v);
        }
        return new RowBox(rowH, placed);
    }

    /// <summary>Tableau : rangées entières, une rangée trop haute pour la place restante passe à la page suivante.</summary>
    int? PlaceTable(Table t, int i, int startRow, Cursor c, PageBreaker newPage)
    {
        var rows = t.Items.OfType<TableRow>().ToList();
        for (int r = startRow; r < rows.Count; r++)
        {
            var box = MeasureRow(t, rows[r], c.Geo.Left, c.Geo.ColumnWidth);
            if (c.Y + box.Height > c.Page.ContentBottom + 0.01 && !c.Empty)
                if (newPage(i, r, fromBreak: false) is { } k) return k;
            foreach (var (cx, cw, lines) in box.Cells)
            {
                c.Page.Cells.Add((cx, c.Y, cw, box.Height));
                foreach (var (pl, ly) in lines)
                {
                    double yy = c.Y + ly;
                    for (int l = 0; l < pl.Lines.Count; l++)
                    {
                        c.Page.Lines.Add(new PlacedLine(pl, l, cx + CellPadding, yy));
                        yy += pl.Lines[l].Height;
                    }
                }
            }
            c.Y += box.Height;
            c.Empty = false;
        }
        c.Prev = null;
        return null;
    }
}
