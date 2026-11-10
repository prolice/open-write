using System.Globalization;
using System.Xml.Linq;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;

namespace OpenWrite.Layout;

// Étage 1 du moteur (PLAN.md §3.3) : paragraphe → lignes. Le texte affiché du paragraphe est mis à plat dans une
// chaîne (codes de champs masqués, résultats affichés, texte masqué et suppressions suivies omis) avec un segment de
// mise en forme par portion ; la coupure des lignes travaille sur les avances cumulées des caractères.

/// <summary>Portion de texte de même mise en forme, résolue pour la mise en page.</summary>
public sealed class TextStyle
{
    public required RunFormat Format { get; init; }
    public required FontFace Face { get; init; }
    /// <summary>Taille de la police affichée (points ; réduite pour l'exposant et l'indice).</summary>
    public double Size { get; init; }
    /// <summary>Espacement ajouté après chaque caractère (points).</summary>
    public double Tracking { get; init; }
}

/// <summary>Caractères spéciaux du texte mis à plat.</summary>
public static class Atom
{
    public const char Tab = '\t';
    public const char LineBreak = '\n';
    public const char PageBreak = '\f';
    public const char ColumnBreak = '\v';
    /// <summary>Objet en ligne (image, forme) : sa taille est dans <see cref="ParagraphContent.Objects"/>.</summary>
    public const char Object = '￼';
}

/// <summary>Texte affiché d'un paragraphe et mise en forme de chaque caractère.</summary>
public sealed class ParagraphContent
{
    public required string Text { get; init; }
    /// <summary>Indice du style de chaque caractère dans <see cref="Styles"/>.</summary>
    public required int[] StyleOf { get; init; }
    public required List<TextStyle> Styles { get; init; }
    /// <summary>Taille des objets en ligne (points), par position dans le texte.</summary>
    public Dictionary<int, (double W, double H)> Objects { get; init; } = new();
    /// <summary>Style de la marque de paragraphe (hauteur d'une ligne vide, dernière ligne).</summary>
    public required TextStyle Mark { get; init; }
    /// <summary>Le paragraphe entier est masqué (marque et texte) : il n'occupe aucune place.</summary>
    public bool Hidden { get; init; }
}

public enum LineEnd { Wrap, Paragraph, LineBreak, PageBreak, ColumnBreak }

/// <summary>Morceau d'une ligne : caractères consécutifs de même style.</summary>
public readonly record struct LinePiece(int Start, int Length, int Style, double X, double Width);

public sealed class LineBox
{
    /// <summary>Plage du texte : [Start, End[ (espaces de fin compris).</summary>
    public int Start { get; init; }
    public int End { get; init; }
    /// <summary>Hauteur de la ligne (interligne appliqué), points.</summary>
    public double Height { get; init; }
    /// <summary>Position de la ligne de base depuis le haut de la ligne.</summary>
    public double Baseline { get; init; }
    public LineEnd EndKind { get; init; }
    /// <summary>Abscisse des morceaux depuis le bord gauche de la colonne (retraits et alignement appliqués).</summary>
    public List<LinePiece> Pieces { get; } = new();
    /// <summary>Espace ajouté à chaque espace pour la justification.</summary>
    public double ExtraPerSpace { get; init; }
}

/// <summary>Paragraphe découpé en lignes pour une largeur donnée.</summary>
public sealed class ParagraphLayout
{
    public required Paragraph Paragraph { get; init; }
    public required ParagraphFormat Format { get; init; }
    public required ParagraphContent Content { get; init; }
    public required double Width { get; init; }
    public List<LineBox> Lines { get; } = new();
    /// <summary>Puce ou numéro de liste, posé avant la première ligne.</summary>
    public (string Text, TextStyle Style, double X)? Marker { get; init; }
    public double SpaceBefore => Format.SpacingBefore / 20.0;
    public double SpaceAfter => Format.SpacingAfter / 20.0;
    public double Height => Lines.Sum(l => l.Height);
}

/// <summary>Découpe des paragraphes en lignes (comportement de Word : coupure aux espaces, espaces de fin suspendus).</summary>
public sealed class ParagraphLayouter
{
    readonly StyleResolver _styles;
    readonly IFontProvider _fonts;
    readonly XNamespace W;
    readonly Dictionary<RunFormat, TextStyle> _styleCache = new();
    /// <summary>Taquet par défaut (w:defaultTabStop), points.</summary>
    public double DefaultTabStop { get; }

    public ParagraphLayouter(StyleResolver styles, IFontProvider fonts)
    {
        _styles = styles;
        _fonts = fonts;
        W = styles.Sheet.W;
        var tab = styles.Sheet.Settings?.Element(W + "defaultTabStop")?.Attribute(W + "val")?.Value;
        DefaultTabStop = tab is not null && double.TryParse(tab, NumberStyles.Float, CultureInfo.InvariantCulture, out var t) && t > 0 ? t / 20 : 36;
    }

    public StyleResolver Styles => _styles;

    TextStyle Style(RunFormat rf)
    {
        if (_styleCache.TryGetValue(rf, out var s)) return s;
        double size = Math.Max(1, rf.Size);
        if (rf.VerticalAlign is "superscript" or "subscript") size *= 2.0 / 3;
        s = new TextStyle { Format = rf, Face = _fonts.Face(rf.FontAscii ?? rf.FontHAnsi, rf.Bold, rf.Italic), Size = size, Tracking = rf.CharacterSpacing / 20.0 };
        _styleCache[rf] = s;
        return s;
    }

    // ------------------------------------------------------------------ texte affiché

    public ParagraphContent Content(Paragraph p, string? tableStyle)
    {
        var sb = new System.Text.StringBuilder();
        var styleOf = new List<int>();
        var styles = new List<TextStyle>();
        var index = new Dictionary<TextStyle, int>();
        var objects = new Dictionary<int, (double, double)>();
        var inCode = new Stack<bool>();
        bool anyVisible = false;

        void Add(string text, TextStyle st)
        {
            if (!index.TryGetValue(st, out var i)) { i = styles.Count; styles.Add(st); index[st] = i; }
            sb.Append(text);
            for (int k = 0; k < text.Length; k++) styleOf.Add(i);
            anyVisible = true;
        }

        void Walk(IEnumerable<Inline> inlines)
        {
            foreach (var inline in inlines)
            {
                switch (inline)
                {
                    case Run r:
                        var rf = _styles.Run(p, r, tableStyle);
                        var st = Style(rf);
                        foreach (var c in r.Content)
                        {
                            bool hidden = inCode.Any(x => x) || rf.Hidden;
                            switch (c)
                            {
                                case FieldChar { Kind: FieldCharKind.Begin } fc:
                                    inCode.Push(true);
                                    if (!hidden && fc.FormData?.Element(W + "checkBox") is not null) Add("☐", st);
                                    break;
                                case FieldChar { Kind: FieldCharKind.Separate }:
                                    if (inCode.Count > 0) { inCode.Pop(); inCode.Push(false); }
                                    break;
                                case FieldChar:
                                    if (inCode.Count > 0) inCode.Pop();
                                    break;
                                case TextContent { Deleted: false } t when !hidden:
                                    Add(rf.Caps ? t.Value.ToUpper(CultureInfo.CurrentCulture) : t.Value, st);
                                    break;
                                case TabContent when !hidden:
                                    Add("\t", st);
                                    break;
                                case BreakContent b when !hidden:
                                    var type = b.Attributes.FirstOrDefault(a => a.Name.LocalName == "type")?.Value;
                                    Add(type == "page" ? "\f" : type == "column" ? "\v" : "\n", st);
                                    break;
                                case OpaqueRunContent o when !hidden && InlineObjectSize(o.Xml) is { } size:
                                    objects[sb.Length] = size;
                                    Add("￼", st);
                                    break;
                            }
                        }
                        break;
                    case InlineContainer box:
                        Walk(box.Children);
                        break;
                    case InlineContentControl sdt:
                        Walk(sdt.Content);
                        break;
                }
            }
        }

        Walk(p.Inlines);
        var mark = Style(_styles.ParagraphMark(p, tableStyle));
        return new ParagraphContent
        {
            Text = sb.ToString(), StyleOf = styleOf.ToArray(), Styles = styles, Objects = objects, Mark = mark,
            Hidden = !anyVisible && mark.Format.Hidden,
        };
    }

    /// <summary>Taille (points) d'un dessin en ligne (wp:inline/wp:extent, EMU) ; null pour un objet flottant (lot 7).</summary>
    static (double, double)? InlineObjectSize(XElement xml)
    {
        XElement? inline = xml.Name.LocalName switch
        {
            "drawing" => xml.Elements().FirstOrDefault(e => e.Name.LocalName == "inline"),
            "AlternateContent" => xml.Descendants().FirstOrDefault(e => e.Name.LocalName == "inline"),
            _ => null,
        };
        var ext = inline?.Elements().FirstOrDefault(e => e.Name.LocalName == "extent");
        if (ext is null) return null;
        static double Emu(XAttribute? a) => a is not null && long.TryParse(a.Value, out var v) ? v / 12700.0 : 0;
        return (Emu(ext.Attribute("cx")), Emu(ext.Attribute("cy")));
    }

    // ------------------------------------------------------------------ lignes

    /// <summary>Découpe le paragraphe en lignes pour une colonne de largeur <paramref name="width"/> (points).</summary>
    public ParagraphLayout Layout(Paragraph p, double width, string? tableStyle = null)
    {
        var pf = _styles.Paragraph(p, tableStyle);
        var content = Content(p, tableStyle);
        double left = pf.IndentLeft / 20.0, right = pf.IndentRight / 20.0, first = pf.IndentFirstLine / 20.0;
        double limit = width - right;

        // puce ou numéro (les compteurs de liste viennent avec le lot 4 : « 1. » indicatif)
        (string, TextStyle, double)? marker = null;
        double firstX = left + first;
        if (pf.NumId > 0 && _styles.Sheet.ListLevel(pf.NumId, pf.ListLevel) is { } lvl && _styles.ListMarker(p, tableStyle) is { } lmf)
        {
            var fmt = lvl.Element(W + "numFmt")?.Attribute(W + "val")?.Value;
            var mtext = lvl.Element(W + "lvlText")?.Attribute(W + "val")?.Value ?? "";
            if (fmt != "bullet") mtext = System.Text.RegularExpressions.Regex.Replace(mtext, "%[1-9]", "1");
            if (fmt == "none") mtext = "";
            var ms = Style(lmf);
            double mx = left + first, mw = ms.Face.Width(mtext, ms.Size);
            marker = (mtext, ms, mx);
            var suff = lvl.Element(W + "suff")?.Attribute(W + "val")?.Value ?? "tab";
            double end = mx + mw;
            firstX = suff switch
            {
                "nothing" => end,
                "space" => end + ms.Face.Advance(' ') * ms.Size,
                // tabulation : le retrait suspendu sert de taquet implicite, sinon le taquet suivant
                _ => first < 0 && end < left ? left : NextTab(pf, end),
            };
        }

        var layout = new ParagraphLayout { Paragraph = p, Format = pf, Content = content, Width = width, Marker = marker };
        if (content.Hidden) return layout;
        var text = content.Text;
        int n = text.Length;
        int pos = 0;
        bool firstLine = true;
        while (true)
        {
            double x0 = firstLine ? firstX : left;
            var (end, kind) = FillLine(content, pf, pos, x0, limit);
            layout.Lines.Add(MakeLine(content, pf, pos, end, kind, x0, limit, last: end >= n));
            firstLine = false;
            if (kind is LineEnd.LineBreak or LineEnd.PageBreak or LineEnd.ColumnBreak) pos = end + 1;
            else pos = end;
            if (pos >= n && kind is LineEnd.Paragraph or LineEnd.Wrap) break;
            if (pos >= n && kind is LineEnd.LineBreak or LineEnd.PageBreak or LineEnd.ColumnBreak)
            {
                // saut en fin de paragraphe : Word ajoute une ligne vide (la marque de paragraphe)
                layout.Lines.Add(MakeLine(content, pf, n, n, LineEnd.Paragraph, left, limit, last: true));
                break;
            }
        }
        return layout;
    }

    double Advance(ParagraphContent c, int i)
    {
        var st = c.Styles[c.StyleOf[i]];
        var ch = c.Text[i];
        if (ch == Atom.Object) return c.Objects.TryGetValue(i, out var o) ? o.W : 0;
        return st.Face.Advance(ch) * st.Size + st.Tracking;
    }

    double NextTab(ParagraphFormat pf, double x)
    {
        // taquets du paragraphe (w:tabs), puis taquets par défaut mesurés depuis le bord de la colonne
        foreach (var t in Tabs(pf))
            if (t > x + 0.01) return t;
        return (Math.Floor(x / DefaultTabStop + 1e-6) + 1) * DefaultTabStop;
    }

    readonly Dictionary<XElement, double[]> _tabCache = new();

    double[] Tabs(ParagraphFormat pf)
    {
        if (_tabCache.TryGetValue(pf.Effective, out var tabs)) return tabs;
        tabs = pf.Effective.Element(W + "tabs")?.Elements(W + "tab")
            .Where(t => t.Attribute(W + "val")?.Value is not ("clear" or "bar"))
            .Select(t => int.TryParse(t.Attribute(W + "pos")?.Value, out var v) ? v / 20.0 : -1).Where(v => v >= 0).OrderBy(v => v).ToArray() ?? Array.Empty<double>();
        _tabCache[pf.Effective] = tabs;
        return tabs;
    }

    /// <summary>Fin de la ligne qui commence en <paramref name="start"/> : coupure au dernier espace qui tient.</summary>
    (int End, LineEnd Kind) FillLine(ParagraphContent c, ParagraphFormat pf, int start, double x0, double limit)
    {
        var text = c.Text;
        int n = text.Length;
        double x = x0;
        int lastBreak = -1;            // position après le dernier espace (début du mot suivant)
        bool anyInk = false;
        for (int i = start; i < n; i++)
        {
            char ch = text[i];
            switch (ch)
            {
                case Atom.LineBreak: return (i, LineEnd.LineBreak);
                case Atom.PageBreak: return (i, LineEnd.PageBreak);
                case Atom.ColumnBreak: return (i, LineEnd.ColumnBreak);
                case Atom.Tab:
                    x = NextTab(pf, x);
                    lastBreak = i + 1;
                    // un taquet au-delà de la marge : Word passe à la ligne (sauf en tête de ligne)
                    if (x > limit + 0.01 && anyInk) return (i, LineEnd.Wrap);
                    continue;
                case ' ':
                    x += Advance(c, i);
                    lastBreak = i + 1;     // les espaces de fin de ligne débordent sans forcer la coupure
                    continue;
            }
            double w = Advance(c, i);
            if (x + w > limit + 0.01 && anyInk)
            {
                if (lastBreak > start) return (lastBreak, LineEnd.Wrap);
                return (i, LineEnd.Wrap);       // mot plus long que la ligne : coupé au caractère
            }
            x += w;
            anyInk = true;
        }
        return (n, LineEnd.Paragraph);
    }

    LineBox MakeLine(ParagraphContent c, ParagraphFormat pf, int start, int end, LineEnd kind, double x0, double limit, bool last)
    {
        // hauteur naturelle : plus grand ascent + plus grand descent des polices de la ligne (marque de paragraphe
        // comprise sur la dernière ligne, comme Word), objets en ligne posés sur la ligne de base
        double asc = 0, desc = 0, lead = 0;
        void Measure(TextStyle st)
        {
            asc = Math.Max(asc, st.Face.Ascent * st.Size);
            desc = Math.Max(desc, st.Face.Descent * st.Size);
            lead = Math.Max(lead, st.Face.ExternalLeading * st.Size);
        }
        int prevStyle = -1;
        for (int i = start; i < end; i++)
        {
            int s = c.StyleOf[i];
            if (c.Text[i] == Atom.Object && c.Objects.TryGetValue(i, out var o)) asc = Math.Max(asc, o.H);
            if (s != prevStyle) { Measure(c.Styles[s]); prevStyle = s; }
        }
        if (last || end == start) Measure(c.Mark);
        double natural = asc + desc + lead;
        double h = pf.LineRule switch
        {
            "exact" => Math.Abs(pf.Line) / 20.0,
            "atLeast" => Math.Max(natural, pf.Line / 20.0),
            _ => natural * Math.Max(pf.Line, 1) / 240.0,
        };
        // ligne de base : l'espace ajouté par l'interligne se place au-dessus du texte
        double baseline = Math.Max(0, h - desc);

        var line = new LineBox { Start = start, End = end, Height = h, Baseline = baseline, EndKind = kind };
        // morceaux de même style, abscisses (alignement à gauche ; centré / droite / justifié décalés ensuite)
        double x = x0;
        int pieceStart = start, spaces = 0;
        double pieceX = x;
        for (int i = start; i <= end; i++)
        {
            bool cut = i == end || c.StyleOf[i] != c.StyleOf[pieceStart] || c.Text[i] == Atom.Tab || (i > pieceStart && c.Text[i - 1] == Atom.Tab);
            if (cut && i > pieceStart)
            {
                line.Pieces.Add(new LinePiece(pieceStart, i - pieceStart, c.StyleOf[pieceStart], pieceX, x - pieceX));
                pieceStart = i;
                pieceX = x;
            }
            if (i == end) break;
            if (c.Text[i] == Atom.Tab) { x = NextTab(pf, x); pieceX = i + 1 <= end ? x : pieceX; pieceStart = i + 1; continue; }
            if (c.Text[i] == ' ') spaces++;
            x += Advance(c, i);
        }
        // largeur utile : sans les espaces de fin
        double ink = x;
        for (int i = end - 1; i >= start && c.Text[i] == ' '; i--) { ink -= Advance(c, i); spaces--; }
        double slack = limit - ink;
        double shift = pf.Alignment switch
        {
            "center" => slack / 2,
            "right" => slack,
            _ => 0,
        };
        double extra = pf.Alignment is "both" or "distribute" && kind == LineEnd.Wrap && spaces > 0 && slack > 0 ? slack / spaces : 0;
        if (shift != 0)
            for (int k = 0; k < line.Pieces.Count; k++)
                line.Pieces[k] = line.Pieces[k] with { X = line.Pieces[k].X + shift };
        return new LineBox { Start = start, End = end, Height = h, Baseline = baseline, EndKind = kind, ExtraPerSpace = extra }.WithPieces(line.Pieces);
    }
}

static class LineBoxExtensions
{
    public static LineBox WithPieces(this LineBox l, List<LinePiece> pieces)
    {
        l.Pieces.AddRange(pieces);
        return l;
    }
}
