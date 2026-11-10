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
    /// <summary>Largeur du texte entre les marges (toutes colonnes confondues).</summary>
    public double ColumnWidth => PageWidth - Left - Right;
    public double ContentBottom => PageHeight - Bottom;
    /// <summary>Colonnes (w:cols) : abscisse depuis la marge de gauche et largeur, points ; une seule par défaut.</summary>
    public IReadOnlyList<(double X, double W)> Cols { get; init; } = Array.Empty<(double, double)>();
    /// <summary>Trait vertical entre les colonnes (w:cols w:sep).</summary>
    public bool ColumnSeparator { get; init; }
    /// <summary>
    /// En-têtes et pieds de la section, hérités des sections précédentes comme dans Word : « header:default »,
    /// « footer:first », « header:even »… → r:id de la partie.
    /// </summary>
    public IReadOnlyDictionary<string, string> HeaderRefs { get; init; } = new Dictionary<string, string>();
    /// <summary>Première page différente (w:titlePg).</summary>
    public bool TitlePage { get; init; }
    /// <summary>Numérotation des pages : valeur de départ (w:pgNumType w:start, null : suite de la section précédente) et format.</summary>
    public int? NumberStart { get; init; }
    public string NumberFormat { get; init; } = "decimal";
    /// <summary>Colonne <paramref name="k"/> : abscisse depuis la marge de gauche et largeur.</summary>
    public (double X, double W) Col(int k) => Cols.Count == 0 ? (0, ColumnWidth) : Cols[Math.Clamp(k, 0, Cols.Count - 1)];
    public int ColumnCount => Math.Max(1, Cols.Count);
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
        var g = new SectionGeometry(Get("pgSz", "w", 11906), Get("pgSz", "h", 16838), Math.Abs(Get("pgMar", "top", 1417)), Math.Abs(Get("pgMar", "bottom", 1417)),
            Get("pgMar", "left", 1417), Get("pgMar", "right", 1417), sectPr?.Element(w + "type")?.Attribute(w + "val")?.Value ?? "nextPage")
        {
            HeaderDistance = Get("pgMar", "header", 708), FooterDistance = Get("pgMar", "footer", 708), SectPr = sectPr,
        };
        var pgNum = sectPr?.Element(w + "pgNumType");
        return g with
        {
            Cols = Columns(sectPr?.Element(w + "cols"), w, g.ColumnWidth), ColumnSeparator = Flag(sectPr?.Element(w + "cols")?.Attribute(w + "sep")),
            TitlePage = sectPr?.Element(w + "titlePg") is { } tp && tp.Attribute(w + "val")?.Value is null or "1" or "true" or "on",
            NumberStart = int.TryParse(pgNum?.Attribute(w + "start")?.Value, out var st) ? st : null,
            NumberFormat = pgNum?.Attribute(w + "fmt")?.Value ?? "decimal",
        };
    }

    static bool Flag(XAttribute? a) => a?.Value is "1" or "true" or "on";

    /// <summary>
    /// Colonnes d'une section : w:num colonnes égales séparées de w:space (720 twips par défaut), ou les w:col explicites
    /// si w:equalWidth vaut 0. Une colonne : liste vide.
    /// </summary>
    static IReadOnlyList<(double X, double W)> Columns(XElement? cols, XNamespace w, double textWidth)
    {
        if (cols is null) return Array.Empty<(double, double)>();
        static double Tw(XAttribute? a, double def) => a is not null && double.TryParse(a.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v / 20 : def;
        var explicitCols = cols.Elements(w + "col").ToList();
        bool equal = cols.Attribute(w + "equalWidth")?.Value is not ("0" or "false" or "off");
        var res = new List<(double, double)>();
        if (!equal && explicitCols.Count > 1)
        {
            double x = 0;
            foreach (var c in explicitCols)
            {
                double cw = Tw(c.Attribute(w + "w"), textWidth / explicitCols.Count);
                res.Add((x, cw));
                x += cw + Tw(c.Attribute(w + "space"), 0);
            }
            return res;
        }
        int n = int.TryParse(cols.Attribute(w + "num")?.Value, out var num) ? num : 1;
        if (n <= 1) return Array.Empty<(double, double)>();
        double space = Tw(cols.Attribute(w + "space"), 36);
        double width = (textWidth - (n - 1) * space) / n;
        if (width < 10) return Array.Empty<(double, double)>();
        for (int k = 0; k < n; k++) res.Add((k * (width + space), width));
        return res;
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

/// <summary>Cellule de tableau placée : rectangle (points), fond (RRGGBB) et bordures (traits ; épaisseur 0 = aucun).</summary>
public readonly record struct CellBox(double X, double Y, double W, double H, string? Fill, BorderLine Top, BorderLine Bottom, BorderLine Left, BorderLine Right)
{
    /// <summary>Cellule du modèle (la première d'une fusion verticale) ; null pour un trait qui n'est pas une cellule.</summary>
    public TableCell? Cell { get; init; }
    public Table? Table { get; init; }
    /// <summary>Rang de la rangée dans le tableau, colonne de la grille où commence la cellule.</summary>
    public int Row { get; init; }
    public int Column { get; init; }
    /// <summary>Cellule d'un tableau imbriqué dans une cellule.</summary>
    public bool Nested { get; init; }
    /// <summary>Fusion verticale : la cellule continue celle du dessus / dans celle du dessous (pas de trait entre elles).</summary>
    public bool JoinTop { get; init; }
    public bool JoinBottom { get; init; }
}

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
    /// <summary>Numéro de la page (sections qui recommencent comprises), texte affiché (format de la section), pages de la section.</summary>
    public int Number { get; set; }
    /// <summary>Numéro de la première ligne numérotée de la page (w:lnNumType), 0 si la section n'en a pas.</summary>
    public int FirstLineNumber { get; set; }
    public string NumberText { get; set; } = "";
    public int SectionPageCount { get; set; }
    /// <summary>Traits entre colonnes (abscisse, haut, bas ; points), pour l'affichage.</summary>
    public List<(double X, double Top, double Bottom)> ColumnRules { get; } = new();
    /// <summary>En-tête et pied placés pour l'affichage : lignes [First, First + Count[ de chaque partie (r:id).</summary>
    public List<(int First, int Count, string PartId)> Parts { get; } = new();
    /// <summary>Objets flottants de la page (ancrés dans le corps, puis ceux des en-têtes et pieds), avec leurs zones de texte.</summary>
    public List<FloatBox> Floats { get; } = new();
    public FlowPosition End { get; set; }
}

/// <summary>Mise en page d'un document : pages, caches des paragraphes, reprise incrémentale.</summary>
public sealed partial class DocumentLayout
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
        // en-têtes pairs et impairs (paramètre du document, w:evenAndOddHeaders)
        var settings = styles.Sheet.Settings;
        EvenOddHeaders = settings?.Element(styles.Sheet.W + "evenAndOddHeaders") is { } eo && eo.Attribute(styles.Sheet.W + "val")?.Value is null or "1" or "true" or "on";
    }

    /// <summary>Le document a des en-têtes différents pour les pages paires et impaires.</summary>
    public bool EvenOddHeaders { get; }

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

    /// <summary>Styles ou listes ajoutés (pas modifiés) : nouveau résolveur, la mise en page incrémentale reste valable.</summary>
    public void ReplaceStyles(StyleResolver styles) => _lines.ReplaceStyles(styles);

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
        // en-têtes hérités : une section sans référence d'un type reprend celle de la section précédente
        var inherited = new Dictionary<string, string>();
        void CloseSection(XElement? sectPr)
        {
            int index = Sections.Count;
            if (sectPr is not null)
                foreach (var e in sectPr.Elements().Where(e => e.Name.LocalName is "headerReference" or "footerReference"))
                {
                    var kind = e.Name.LocalName == "headerReference" ? "header" : "footer";
                    var type = e.Attributes().FirstOrDefault(a => a.Name.LocalName == "type")?.Value ?? "default";
                    if (e.Attributes().FirstOrDefault(a => a.Name.LocalName == "id")?.Value is { } id) inherited[$"{kind}:{type}"] = id;
                }
            var geo = SectionGeometry.From(sectPr) with { HeaderRefs = new Dictionary<string, string>(inherited) };
            Sections.Add(WithHeaders(geo, sectPr));
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
        if (_related is null) return g;
        double Extent(string kind, string type) =>
            g.HeaderRefs.TryGetValue($"{kind}:{type}", out var id) && _related(id) is { } story ? StoryHeight(story, g.ColumnWidth) : 0;
        bool titlePg = g.TitlePage;
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
    /// <summary>Type d'en-tête de la page : « first » (première page, w:titlePg), « even » (page paire), sinon « default ».</summary>
    public string HeaderType(Page page) =>
        page.FirstOfSection && page.Geometry.TitlePage ? "first" : EvenOddHeaders && page.Number % 2 == 0 ? "even" : "default";

    /// <summary>En-tête ou pied (r:id) d'une page, hérité au besoin ; null si la page n'en a pas.</summary>
    public string? HeaderId(Page page, string kind) => page.Geometry.HeaderRefs.TryGetValue($"{kind}:{HeaderType(page)}", out var id) ? id : null;

    public Page? HeaderFooter(Page page, int pageCount)
    {
        var g = page.Geometry;
        if (_related is null) return null;
        Story? Find(string kind) => HeaderId(page, kind) is { } id ? _related(id) : null;
        var result = new Page { Geometry = g, Start = page.Start };
        _lines.FieldValue = code =>
        {
            var name = code.Split(' ', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()?.ToUpperInvariant();
            return name switch
            {
                "PAGE" => page.NumberText,
                "NUMPAGES" => pageCount.ToString(CultureInfo.InvariantCulture),
                "SECTIONPAGES" => page.SectionPageCount.ToString(CultureInfo.InvariantCulture),
                _ => null,
            };
        };
        try
        {
            if (Find("header") is { } header)
            {
                int n0 = result.Lines.Count;
                PlaceStory(header, g.HeaderDistance, g, result, cache: false);
                result.Parts.Add((n0, result.Lines.Count - n0, HeaderId(page, "header")!));
            }
            if (Find("footer") is { } footer)
            {
                int n0 = result.Lines.Count;
                PlaceStory(footer, g.PageHeight - g.FooterDistance - StoryHeight(footer, g.ColumnWidth), g, result, cache: false);
                result.Parts.Add((n0, result.Lines.Count - n0, HeaderId(page, "footer")!));
            }
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
            result.Cells.Add(new CellBox(g.Left, y + sep / 2, g.ColumnWidth / 3, 0, null, new BorderLine(0.5), default, default, default));
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
    double PlaceStory(Story story, double top, SectionGeometry g, Page into, bool cache, double? left = null, double? width = null)
    {
        double y = top;
        double x = left ?? g.Left, colW = width ?? g.ColumnWidth;
        foreach (var b in story.Blocks)
        {
            switch (b)
            {
                case Paragraph p:
                    var pl = cache ? LayoutOf(p, colW) : _lines.Layout(p, colW);
                    if (pl.Content.Hidden) break;
                    y += pl.SpaceBefore + BorderOf(pl, "top");
                    for (int l = 0; l < pl.Lines.Count; l++)
                    {
                        into.Lines.Add(new PlacedLine(pl, l, x, y));
                        y += pl.Lines[l].Height;
                    }
                    y += BorderOf(pl, "bottom") + pl.SpaceAfter;
                    break;
                case Table t:
                    foreach (var row in MeasureTable(t, x, colW).Rows)
                    {
                        PlaceRowAt(row, y, into);
                        y += row.Height;
                    }
                    break;
                case BlockContentControl sdt:
                    y = PlaceStory(sdt.Content, y, g, into, cache, x, colW);
                    break;
            }
        }
        return y;
    }

    /// <summary>
    /// Objets flottants des en-têtes et pieds de la page : placés une fois par page, ils écartent aussi le texte du
    /// corps (logo d'en-tête posé dans la marge du texte, traits du pied d'Horaires).
    /// </summary>
    void AddHeaderFloats(Page page)
    {
        if (_related is null) return;
        var g = page.Geometry;
        foreach (var kind in new[] { "header", "footer" })
        {
            if (HeaderId(page, kind) is not { } id || _related(id) is not { } story) continue;
            double y = kind == "header" ? g.HeaderDistance : g.PageHeight - g.FooterDistance - StoryHeight(story, g.ColumnWidth);
            foreach (var b in story.Blocks)
            {
                if (b is not Paragraph p) { y += b is Table t ? MeasureTable(t, g.Left, g.ColumnWidth).Rows.Sum(r => r.Height) : 0; continue; }
                var pl = LayoutOf(p, g.ColumnWidth);
                if (pl.Content.Hidden) continue;
                y += pl.SpaceBefore;
                foreach (var (_, o) in pl.Content.Anchors)
                {
                    var (fx, fy, fw, fh) = Position(o, page, g, g.Left, g.ColumnWidth, y, y, g.Left);
                    var box = new FloatBox { Object = o, Anchor = null, PartId = id, X = fx, Y = fy, W = fw, H = fh };
                    PlaceTextBoxes(box, g);
                    page.Floats.Add(box);
                }
                y += pl.Height + pl.SpaceAfter;
            }
        }
    }

    /// <summary>Hauteur d'une histoire (en-tête, pied) : paragraphes et rangées de tableau empilés.</summary>
    /// <summary>Hauteur d'un en-tête ou d'un pied (édition en place : faut-il remettre le corps en page ?).</summary>
    public double StoryHeightOf(Story story, double width) => StoryHeight(story, width);

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
                    h += MeasureTable(t, 0, width).Rows.Sum(r => r.Height);
                    prev = null;
                    break;
                case BlockContentControl sdt:
                    h += StoryHeight(sdt.Content, width);
                    break;
            }
        }
        return h;
    }

    ParagraphLayout LayoutOf(Paragraph p, double width, CellStyle cell = default)
    {
        if (_cache.TryGetValue(p, out var l) && Math.Abs(l.Width - width) < 0.01 && l.Cell == cell) return l;
        l = _lines.Layout(p, width, cell);
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
        Renumber();
    }

    /// <summary>
    /// Numéros des pages : chaque section repart de w:pgNumType w:start si elle en a un, sinon continue ; texte au format
    /// de la section (romain, lettres…) ; nombre de pages de chaque section (champ SECTIONPAGES).
    /// </summary>
    void Renumber()
    {
        int number = 0;
        int sectionStart = 0;
        for (int k = 0; k < Pages.Count; k++)
        {
            var pg = Pages[k];
            bool newSection = k == 0 || pg.FirstOfSection || !ReferenceEquals(pg.Geometry, Pages[k - 1].Geometry) && pg.Geometry.SectPr != Pages[k - 1].Geometry.SectPr;
            if (newSection)
            {
                for (int j = sectionStart; j < k; j++) Pages[j].SectionPageCount = k - sectionStart;
                sectionStart = k;
                number = pg.Geometry.NumberStart is { } st && (k == 0 || pg.FirstOfSection) ? st : number + 1;
            }
            else number++;
            pg.Number = number;
            pg.NumberText = ListNumbering.Format(number, pg.Geometry.NumberFormat);
        }
        for (int j = sectionStart; j < Pages.Count; j++) Pages[j].SectionPageCount = Pages.Count - sectionStart;
        NumberLines();
    }

    /// <summary>
    /// Numéros de ligne (w:lnNumType) : lignes des paragraphes du corps hors tableaux, sauf w:suppressLineNumbers ; la
    /// numérotation repart à chaque page (par défaut), à chaque section ou continue.
    /// </summary>
    void NumberLines()
    {
        int count = 0;
        SectionGeometry? prev = null;
        foreach (var pg in Pages)
        {
            var g = pg.Geometry;
            var ln = g.SectPr?.Element(g.SectPr.Name.Namespace + "lnNumType");
            if (ln is null) { pg.FirstLineNumber = 0; prev = g; continue; }
            var w = ln.Name.Namespace;
            var restart = ln.Attribute(w + "restart")?.Value ?? "newPage";
            int start = int.TryParse(ln.Attribute(w + "start")?.Value, out var st) ? st + 1 : 1;
            if (restart == "newPage" || restart == "newSection" && !ReferenceEquals(prev, g) && pg.FirstOfSection || prev is null || prev.SectPr?.Element(w + "lnNumType") is null)
                count = start - 1;
            pg.FirstLineNumber = count + 1;
            count += NumberedLines(pg).Count();
            prev = g;
        }
    }

    /// <summary>Police et taille par défaut du document (numéros de ligne, comme le style « Numéro de ligne » de Word).</summary>
    public RunFormat DefaultRunFormat => _defaultRun ??= _lines.Styles.ParagraphMark(new Paragraph());
    RunFormat? _defaultRun;

    /// <summary>Lignes numérotées d'une page (paragraphes du corps hors tableaux, sans w:suppressLineNumbers).</summary>
    public IEnumerable<PlacedLine> NumberedLines(Page page) => page.Lines.Where(l =>
        _itemIndex.ContainsKey(l.Paragraph.Paragraph) && !l.Paragraph.Format.Effective.Elements().Any(e => e.Name.LocalName == "suppressLineNumbers"));

    /// <summary>
    /// Filigrane de la page : forme WordArt de l'en-tête (v:shape … v:textpath, comme Word l'écrit) — texte, police,
    /// couleur, rotation, largeur et hauteur (points) ; null s'il n'y en a pas.
    /// </summary>
    public (string Text, string Font, string Color, double Rotation, double Width, double Height, double Opacity)? Watermark(Page page)
    {
        if (_related is null || HeaderId(page, "header") is not { } id || _related(id) is not { } story) return null;
        foreach (var p in story.Paragraphs())
            foreach (var r in p.Runs())
                foreach (var c in r.Content.OfType<DrawingContent>())
                {
                    var shape = c.Xml.Descendants().FirstOrDefault(e => e.Name.LocalName == "shape" && e.Elements().Any(x => x.Name.LocalName == "textpath"));
                    if (shape is null) continue;
                    var tp = shape.Elements().First(x => x.Name.LocalName == "textpath");
                    var style = (shape.Attribute("style")?.Value ?? "").Split(';').Select(x => x.Split(':')).Where(x => x.Length == 2)
                        .ToDictionary(x => x[0].Trim(), x => x[1].Trim(), StringComparer.OrdinalIgnoreCase);
                    static double Pt(string? v) => v is not null && double.TryParse(v.Replace("pt", ""), NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? d : 0;
                    var font = System.Text.RegularExpressions.Regex.Match(tp.Attribute("style")?.Value ?? "", "font-family:\\s*\"?([^\";]+)").Groups[1].Value;
                    var fill = shape.Attribute("fillcolor")?.Value ?? "silver";
                    double opacity = shape.Elements().FirstOrDefault(x => x.Name.LocalName == "fill")?.Attribute("opacity")?.Value is { } o
                        && double.TryParse(o.TrimStart('.').Length == o.Length - 1 ? "0" + o : o, NumberStyles.Float, CultureInfo.InvariantCulture, out var op) ? op : 1;
                    return (tp.Attribute("string")?.Value ?? "", font.Length > 0 ? font : "Calibri", fill,
                            Pt(style.GetValueOrDefault("rotation")), Pt(style.GetValueOrDefault("width")), Pt(style.GetValueOrDefault("height")), opacity);
                }
        return null;
    }

    /// <summary>
    /// Reprise après une modification des blocs <paramref name="changed"/> (modifiés en place, insérés ou remplacés) :
    /// la pagination repart de la page qui précède la première modification et s'arrête dès qu'une page retrouve le
    /// point de départ d'une ancienne page située après toutes les modifications.
    /// </summary>
    /// <param name="dirty">
    /// Blocs réellement modifiés à l'intérieur des blocs <paramref name="changed"/> (paragraphes d'une cellule) : seuls
    /// leurs paragraphes sont à refaire ; null : tous ceux des blocs modifiés.
    /// </param>
    public void Update(IReadOnlyCollection<Block> changed, bool rebuildFlow = false, IEnumerable<Block>? dirty = null)
    {
        LastParagraphsLaidOut = 0;
        Invalidate(dirty ?? changed);
        // rebuildFlow : listes touchées (numéros à recalculer) — pas de chemin rapide
        if (!rebuildFlow && Pages.Count > 0 && changed.Count > 0 && _body.Blocks.Count == _topCount && changed.All(b => b is not Paragraph p || p.Props.Get("sectPr") is null) && changed.All(_itemIndex.ContainsKey))
        {
            UpdateInPlace(changed, dirty is not null && changed.Count == 1 && changed.First() is Table t ? ChangedRows(t, dirty) : null);
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
        Renumber();
    }

    /// <summary>
    /// Chemin rapide de la frappe : les blocs modifiés sont déjà des éléments du flux (pas d'insertion ni de
    /// suppression), les indices ne changent pas ; les anciennes pages sont retrouvées par recherche dichotomique.
    /// </summary>
    /// <summary>
    /// Rangées d'un tableau qui contiennent les blocs modifiés (fusions verticales comprises) ; null si la modification
    /// peut toucher tout le tableau (ajusté au contenu : les largeurs des colonnes dépendent du texte).
    /// </summary>
    static (int First, int Last)? ChangedRows(Table t, IEnumerable<Block> dirty)
    {
        var w = t.Props.Ns;
        var rows = t.Items.OfType<TableRow>().ToList();
        if (t.Props.Get("tblLayout")?.Attribute(w + "type")?.Value != "fixed"
            && rows.SelectMany(r => r.Items.OfType<TableCell>()).Any(c => c.Props.Get("tcW") is { } tcw && tcw.Attribute(w + "type")?.Value is "auto" or "nil")) return null;
        var set = new HashSet<Block>(dirty, ReferenceEqualityComparer.Instance);
        int first = int.MaxValue, last = -1;
        for (int r = 0; r < rows.Count; r++)
            foreach (var cell in rows[r].Items.OfType<TableCell>())
                if (Contains(cell.Content, set)) { first = Math.Min(first, r); last = Math.Max(last, r); }
        if (last < 0) return null;
        // fusion verticale : la cellule de départ et ses suites forment une seule hauteur
        bool Merged(int r) => rows[r].Items.OfType<TableCell>().Any(c => c.Props.Get("vMerge") is not null);
        while (first > 0 && Merged(first)) first--;
        while (last + 1 < rows.Count && Merged(last + 1)) last++;
        return (first, last);

        static bool Contains(Story s, HashSet<Block> set) => s.Blocks.Any(b => set.Contains(b)
            || b is Table nt && nt.Items.OfType<TableRow>().Any(r => r.Items.OfType<TableCell>().Any(c => Contains(c.Content, set)))
            || b is BlockContentControl sdt && Contains(sdt.Content, set));
    }

    void UpdateInPlace(IReadOnlyCollection<Block> changed, (int First, int Last)? rows = null)
    {
        // paragraphes nouveaux dans un tableau modifié (rangée, colonne, cellule fractionnée) : rattachés à son élément
        if (_paraItem is not null)
            foreach (var b in changed)
                if (b is not Paragraph)
                    foreach (var q in new Story { Blocks = { b } }.Paragraphs()) _paraItem[q] = _itemIndex[b];
        var old = Pages;
        int first = changed.Min(b => _itemIndex[b]), last = changed.Max(b => _itemIndex[b]);
        // page qui contient l'élément précédant la modification, et une de plus en arrière (paragraphes solidaires)
        int anchor = Math.Max(0, first - 1);
        int startPage = Math.Max(0, PageOf(old, anchor) - 1);
        // texte modifié dans des rangées d'un tableau : reprise à la page où commence la première de ces rangées
        // (les rangées d'avant n'ont pas changé), convergence possible dans le tableau après la dernière
        if (rows is { } rr)
        {
            int k = PageOf(old, first);
            while (k + 1 < old.Count && old[k + 1].Start.Item is not null && old[k + 1].Start.Index == first && old[k + 1].Start.Line <= rr.First) k++;
            startPage = k;
        }
        var start = old[startPage].Start;
        var pages = old.GetRange(0, startPage);
        Func<FlowPosition, int?> converge = pos =>
        {
            if (pos.Item is null || pos.Index < last || pos.Index == last && (rows is not { } r2 || pos.Line <= r2.Last)) return null;
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
        Renumber();
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
        copy.ColumnRules.AddRange(p.ColumnRules);
        copy.Floats.AddRange(p.Floats);
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
        /// <summary>Colonne en cours, haut des colonnes de la section sur cette page, début de la zone dans Page.Lines / Cells.</summary>
        public int Col;
        public double ColTop;
        public int RegionLines, RegionCells;
        public double ColX => Geo.Left + Geo.Col(Col).X;
        public double ColW => Geo.Col(Col).W;
        /// <summary>Un objet flottant ancré plus bas recouvre du texte déjà placé : la page est à refaire.</summary>
        public bool Restart;
        /// <summary>Pages refaites pour la page en cours (borné).</summary>
        public int Restarts;
        /// <summary>État au début de la page (pour la refaire) : position de départ et curseur.</summary>
        public (FlowPosition Start, SectionGeometry Geo, double Y, bool Empty, bool TopFromBreak, ParagraphLayout? Prev, int Col, double ColTop, bool First) Saved;

        public void Save(FlowPosition start) => Saved = (start, Geo, Y, Empty, TopFromBreak, Prev, Col, ColTop, Page.FirstOfSection);
    }

    /// <summary>Pages refaites au plus pour une même page (objets flottants qui se déplacent l'un l'autre).</summary>
    const int MaxRestarts = 3;

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
        var c = new Cursor { Geo = geo0, Page = page0, Y = page0.ContentTop, TopFromBreak = pageTopFromBreak, ColTop = page0.ContentTop };
        // reprise au milieu d'une section continue : sa zone de colonnes commence en haut de page (approximation)
        AddHeaderFloats(page0);
        c.Save(from);

        // colonne suivante (débordement, saut de colonne) ou page suivante ; renvoie l'indice de l'ancienne page si la
        // mise en page converge
        int? NewPage(int item, int line, bool fromBreak, double offset = 0, bool toColumn = true)
        {
            if (toColumn && c.Col + 1 < c.Geo.ColumnCount)
            {
                c.Col++;
                c.Y = c.ColTop;
                c.Empty = true;
                c.TopFromBreak = fromBreak;
                c.Prev = null;
                return null;
            }
            CloseRegion(c);
            AddCellAnchors(c.Page);
            DropOrphanFloats(c.Page);
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
            c.Col = 0;
            c.ColTop = c.Y;
            c.RegionLines = c.RegionCells = 0;
            AddHeaderFloats(c.Page);
            c.Restarts = 0;
            c.Save(pos);
            return null;
        }

        int startLine = from.Line;
        double startOffset = from.Offset;
        for (; i < _items.Count; i++, startLine = 0, startOffset = 0)
        {
            if (c.Restart)
            {
                // un objet ancré plus bas recouvre du texte déjà placé : la page est refaite, l'objet connu dès le départ
                c.Restart = false;
                c.Restarts++;
                var known = c.Page.Floats.Where(f => f.Anchor is not null || f.Table is not null).ToList();
                var s = c.Saved;
                c.Page = new Page { Geometry = s.Geo, Start = s.Start, FirstOfSection = s.First };
                AddHeaderFloats(c.Page);
                c.Page.Floats.AddRange(known);
                (c.Geo, c.Y, c.Empty, c.TopFromBreak, c.Prev, c.Col, c.ColTop) = (s.Geo, s.Y, s.Empty, s.TopFromBreak, s.Prev, s.Col, s.ColTop);
                c.RegionLines = c.RegionCells = 0;
                if (Trace) Console.WriteLine($"    page refaite ({c.Restarts}) : objet flottant ancré plus bas");
                i = s.Start.Index;
                startLine = s.Start.Line;
                startOffset = s.Start.Offset;
            }
            var item = _items[i];
            // changement de section : nouvelle page, sauf section continue de même format
            if (startLine == 0 && i > 0 && _items[i - 1].Section != item.Section)
            {
                var g = Sections[item.Section];
                bool continuous = g.Type == "continuous" && g.PageWidth == c.Geo.PageWidth && g.PageHeight == c.Geo.PageHeight;
                if (!continuous && !c.Empty) { if (NewPage(i, 0, fromBreak: true, toColumn: false) is { } k) return k; }
                else
                {
                    // section continue sur la même page : les colonnes de la section précédente sont équilibrées (Word),
                    // la nouvelle section commence en dessous
                    if (!c.Empty || c.Col > 0) Balance(c);
                    CloseRegion(c);
                    c.Geo = g;
                    c.Col = 0;
                    c.ColTop = c.Y;
                    c.RegionLines = c.Page.Lines.Count;
                    c.RegionCells = c.Page.Cells.Count;
                }
                if (!continuous) { c.Col = 0; }
            }
            int? conv = item.Block switch
            {
                Table t => PlaceTable(t, i, startLine, startOffset, c, NewPage),
                Paragraph p => PlaceParagraph(p, i, startLine, c, NewPage),
                _ => null,
            };
            if (conv is not null) return conv;
            if (c.Restart) i--;                 // la boucle repart du début de la page
        }
        CloseRegion(c);
        AddCellAnchors(c.Page);
        DropOrphanFloats(c.Page);
        c.Page.End = new FlowPosition(null, _items.Count, 0);
        pages.Add(c.Page);
        return null;
    }

    delegate int? PageBreaker(int item, int line, bool fromBreak, double offset = 0, bool toColumn = true);

    /// <summary>Fin d'une zone de colonnes (page ou section continue) : traits entre colonnes si la section en demande.</summary>
    static void CloseRegion(Cursor c)
    {
        var g = c.Geo;
        if (!g.ColumnSeparator || g.ColumnCount < 2) return;
        double bottom = c.ColTop;
        for (int k = c.RegionLines; k < c.Page.Lines.Count; k++)
        {
            var l = c.Page.Lines[k];
            bottom = Math.Max(bottom, l.Y + l.Paragraph.Lines[l.Line].Height);
        }
        if (bottom <= c.ColTop) return;
        for (int k = 1; k < g.ColumnCount; k++)
        {
            var (px, pw) = g.Col(k - 1);
            var (nx, _) = g.Col(k);
            c.Page.ColumnRules.Add((g.Left + (px + pw + nx) / 2, c.ColTop, bottom));
        }
    }

    /// <summary>
    /// Équilibre les colonnes d'une section continue sur la page (Word le fait avant la section suivante) : les lignes de
    /// la zone sont réparties sur toutes les colonnes à hauteur égale, la plus petite possible. Zone avec tableau : laissée telle quelle.
    /// </summary>
    static void Balance(Cursor c)
    {
        var g = c.Geo;
        int n = g.ColumnCount;
        if (n < 2 || c.Page.Cells.Count > c.RegionCells) return;
        var lines = c.Page.Lines;
        int start = c.RegionLines, count = lines.Count - start;
        if (count < 2) return;
        // saut de colonne explicite : la répartition voulue est gardée (Horaires, page 2)
        for (int k = start; k < lines.Count; k++)
            if (lines[k].Paragraph.Lines[lines[k].Line].EndKind == LineEnd.ColumnBreak)
            {
                // la section suivante commence sous la plus longue des colonnes
                // (une ligne qui ne porte que le saut de colonne ne compte pas : Horaires, section suivante à 365 pt)
                for (int j = start; j < lines.Count; j++)
                {
                    var lb = lines[j].Paragraph.Lines[lines[j].Line];
                    if (lb.EndKind == LineEnd.ColumnBreak && lb.End == lb.Start) continue;
                    c.Y = Math.Max(c.Y, lines[j].Y + lb.Height);
                }
                c.Col = 0;
                return;
            }
        // suite d'unités : écart au-dessus (dans la même colonne) et hauteur de chaque ligne
        var units = new List<(double Gap, double H)>();
        double prevBottom = c.ColTop, prevX = double.NaN;
        for (int k = start; k < lines.Count; k++)
        {
            var l = lines[k];
            double h = l.Paragraph.Lines[l.Line].Height;
            double gap = Math.Abs(l.X - prevX) < 0.01 || double.IsNaN(prevX) ? Math.Max(0, l.Y - prevBottom) : 0;
            units.Add((gap, h));
            prevBottom = l.Y + h;
            prevX = l.X;
        }
        double total = units.Sum(u => u.Gap + u.H);
        double step = Math.Max(1, units.Max(u => u.H) / 4);
        for (double target = total / n; target <= total + step; target += step)
        {
            // remplissage glouton : une colonne prend des lignes tant qu'elle ne dépasse pas la hauteur visée
            var col = new int[units.Count];
            var y = new double[units.Count];
            int k = 0;
            double used = 0;
            bool ok = true;
            for (int u = 0; u < units.Count; u++)
            {
                double gap = used == 0 ? 0 : units[u].Gap;
                if (used > 0 && used + gap + units[u].H > target + 0.01)
                {
                    k++;
                    used = 0;
                    gap = 0;
                    if (k >= n) { ok = false; break; }
                }
                col[u] = k;
                y[u] = c.ColTop + used + gap;
                used += gap + units[u].H;
            }
            if (!ok) continue;
            double maxBottom = c.ColTop;
            for (int u = 0; u < units.Count; u++)
            {
                var l = lines[start + u];
                double x = l.X - (c.Geo.Left + FindCol(g, l.X - c.Geo.Left)) + c.Geo.Left + g.Col(col[u]).X;
                lines[start + u] = l with { X = x, Y = y[u] };
                maxBottom = Math.Max(maxBottom, y[u] + units[u].H);
            }
            c.Y = maxBottom;
            c.Col = 0;
            return;
        }
    }

    /// <summary>Abscisse (depuis la marge) de la colonne qui contient une abscisse de ligne.</summary>
    static double FindCol(SectionGeometry g, double x)
    {
        double best = 0;
        for (int k = 0; k < g.ColumnCount; k++) if (g.Col(k).X <= x + 0.01) best = g.Col(k).X;
        return best;
    }

    int? PlaceParagraph(Paragraph p, int i, int startLine, Cursor c, PageBreaker newPage)
    {
        var pl = LayoutOf(p, c.ColW);
        if (pl.Content.Hidden) return null;
        var pf = pl.Format;
        // paragraphe vide qui ne porte qu'un saut de section continu : Word le pose au bout de la ligne précédente, sans
        // hauteur (Horaires, mesuré par COM : marque à la hauteur de la dernière ligne, x = 658 pt)
        if (startLine == 0 && pl.Content.Text.Length == 0 && pl.Content.Anchors.Count == 0 && _items[i].EndsSection && !c.Empty
            && _items[i].Section + 1 < Sections.Count && Sections[_items[i].Section + 1].Type == "continuous"
            && c.Page.Lines.Count > 0 && c.Page.Lines[^1] is var last)
        {
            c.Page.Lines.Add(new PlacedLine(pl, 0, last.X, last.Y));
            return null;
        }
        double SpaceAtTop() => SuppressSpaceBeforeAtPageTop ? 0 : pl.SpaceBefore;
        if (startLine == 0 && pf.PageBreakBefore && !c.Empty)
            if (newPage(i, 0, fromBreak: true, toColumn: false) is { } k) return k;
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
                double need = before + KeepChainHeight(i, c.ColW);
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
        // objets ancrés au paragraphe : placés avant ses lignes (le paragraphe s'habille autour d'eux) ; s'ils recouvrent
        // du texte déjà placé sur la page, la page est refaite
        if (startLine == 0 && pl.Content.Anchors.Count > 0
            && PlaceAnchors(pl, c.Page, c.Geo, c.ColX, c.ColW, c.Y, c.Page.Lines.Count) && c.Restarts < MaxRestarts)
        {
            c.Restart = true;
            return null;
        }
        pl = Rewrap(p, pl, c, startLine);
        int n = pl.Lines.Count;
        int line = startLine;
        double rowY = c.Y;
        while (line < n)
        {
            var lb = pl.Lines[line];
            if (lb.SameRow)
            {
                // suite de la rangée, de l'autre côté de l'objet
                c.Page.Lines.Add(new PlacedLine(pl, line, c.ColX, rowY));
                line++;
                continue;
            }
            // une ligne qui ne contient que le saut de page reste en bas de la page : elle ne crée pas de page blanche
            // (Syllabus : titre « Introduction » précédé d'un saut, gardé en page 1 par Word)
            // (de même un saut de colonne seul : Horaires, « août 2025 » commence par un saut de colonne en bas de page 1)
            bool breakOnly = lb.EndKind is LineEnd.PageBreak or LineEnd.ColumnBreak && lb.End == lb.Start;
            double notes = NotesHeight(lb, c.Page, c.Geo.ColumnWidth);
            if (c.Y + lb.Gap + lb.FitHeight > c.Page.ContentBottom - notes + BottomTolerance && !c.Empty && !breakOnly)
            {
                if (Trace) Console.WriteLine($"    coupure : élément {i} « {new string(pl.Content.Text.Take(20).Select(ch => ch < ' ' ? '·' : ch).ToArray())} » ligne {line}/{n} ({lb.EndKind}), y {c.Y:0.0} + {lb.FitHeight:0.0} (ligne {lb.Height:0.0}, avant {before:0.0}) > bas {c.Page.ContentBottom:0.0}");
                int here = line - startLine;           // lignes du paragraphe déjà sur cette page
                int cut = line;
                if (pf.WidowControl && n > 1 && here > 0)
                {
                    if (n - cut == 1) cut--;                                   // veuve : une ligne de plus descend
                    if (cut - startLine == 1 && startLine == 0) cut = 0;       // orpheline : tout descend
                    while (cut > startLine && pl.Lines[cut].SameRow) cut--;    // une rangée ne se coupe pas
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
                // tout le paragraphe descend : ses objets ancrés descendent avec lui
                bool moveAnchors = cut == 0 && pl.Content.Anchors.Count > 0;
                if (moveAnchors) c.Page.Floats.RemoveAll(f => ReferenceEquals(f.Anchor, p));
                if (newPage(i, cut, fromBreak: false) is { } k) return k;
                if (cut == 0) c.Y += SpaceAtTop();
                if (moveAnchors) PlaceAnchors(pl, c.Page, c.Geo, c.ColX, c.ColW, c.Y, c.Page.Lines.Count);
                pl = Rewrap(p, pl, c, cut);
                n = pl.Lines.Count;
                line = cut;
                startLine = cut;
                continue;
            }
            c.Y += lb.Gap;
            rowY = c.Y;
            c.Page.Lines.Add(new PlacedLine(pl, line, c.ColX, c.Y) { Notes = notes });
            if (notes > 0)
            {
                c.Page.ReservedBottom += notes;
                c.Page.NotesCount += lb.Footnotes!.Count;
                c.Page.NoteIds.AddRange(lb.Footnotes!);
            }
            c.Y += Math.Max(lb.Height, lb.RowHeight);
            c.Empty = false;
            line++;
            // saut de page dans le paragraphe : la suite (au moins la marque de paragraphe) va à la page suivante
            if (lb.EndKind is LineEnd.PageBreak or LineEnd.ColumnBreak && line < n)
            {
                if (newPage(i, line, fromBreak: true, toColumn: lb.EndKind == LineEnd.ColumnBreak) is { } k) return k;
                startLine = line;
                pl = Rewrap(p, pl, c, line);
                n = pl.Lines.Count;
            }
        }
        c.Y += borderBottom + pl.SpaceAfter;
        c.Prev = pl;
        // saut de page à la fin du paragraphe : l'élément suivant commence sur une nouvelle page
        if (n > 0 && pl.Lines[^1].EndKind is LineEnd.PageBreak or LineEnd.ColumnBreak && i + 1 < _items.Count)
            if (newPage(i + 1, 0, fromBreak: true, toColumn: pl.Lines[^1].EndKind == LineEnd.ColumnBreak) is { } k) return k;
        return null;
    }

    /// <summary>
    /// Lignes du paragraphe à partir de <paramref name="fromLine"/> habillées autour des objets flottants de la page
    /// (haut de la ligne au curseur) ; sans objet gênant, la mise en page ordinaire (en cache).
    /// </summary>
    ParagraphLayout Rewrap(Paragraph p, ParagraphLayout pl, Cursor c, int fromLine)
    {
        var space = c.Page.Floats.Count == 0 ? null : SpaceFor(c.Page, c.Y, c.ColX, c.ColW);
        if (space is null)
        {
            if (!pl.Wrapped) return pl;
            return fromLine == 0 ? LayoutOf(p, c.ColW) : _lines.Layout(p, c.ColW, default, null, pl, fromLine);
        }
        LastParagraphsLaidOut++;
        return _lines.Layout(p, c.ColW, default, space, fromLine > 0 ? pl : null, fromLine);
    }

    /// <summary>Objets ancrés à un paragraphe qui n'a finalement aucune ligne sur la page (descendu après coup) : retirés.</summary>
    static void DropOrphanFloats(Page page)
    {
        if (page.Floats.Count == 0) return;
        var anchors = new HashSet<Paragraph>(page.Lines.Select(l => l.Paragraph.Paragraph), ReferenceEqualityComparer.Instance);
        page.Floats.RemoveAll(f => f.Anchor is not null && !anchors.Contains(f.Anchor));
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
    double KeepChainHeight(int i, double width)
    {
        double h = 0;
        for (int k = i, guard = 0; k < _items.Count && guard < 8; k++, guard++)
        {
            if (_items[k].Block is not Paragraph p) break;
            var pl = LayoutOf(p, width);
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
}
