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
    /// <summary>XML des objets en ligne (w:drawing…), par position, pour l'affichage.</summary>
    public Dictionary<int, XElement> ObjectXml { get; init; } = new();
    /// <summary>Cases à cocher de formulaire : position du glyphe ☐/☒ → caractère de début du champ.</summary>
    public Dictionary<int, FieldChar> CheckBoxes { get; init; } = new();
    /// <summary>Appels de note de bas de page : position dans le texte → w:id de la note.</summary>
    public Dictionary<int, string> Footnotes { get; init; } = new();
    /// <summary>Le paragraphe entier est masqué (marque et texte) : il n'occupe aucune place.</summary>
    public bool Hidden { get; init; }
    /// <summary>
    /// Position dans le modèle (décalage de <c>DocPos</c>, voir ParagraphAtoms) du début et de la fin de chaque
    /// caractère affiché. Un caractère qui représente un champ entier (case à cocher, numéro de page calculé) couvre
    /// le champ du caractère de début au caractère de fin ; un appel de note couvre sa marque.
    /// </summary>
    public int[] ModelStart { get; init; } = Array.Empty<int>();
    public int[] ModelEnd { get; init; } = Array.Empty<int>();
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
    /// <summary>
    /// Hauteur qui doit tenir sur la page : avec un interligne multiple, l'espace ajouté au-delà de la hauteur
    /// naturelle peut déborder la marge du bas (Word ; Annexe : ligne à 1,5 placée à 752,9 pt, bas du corps à 771 pt).
    /// </summary>
    public double FitHeight { get; init; }
    /// <summary>Abscisse des morceaux depuis le bord gauche de la colonne (retraits et alignement appliqués).</summary>
    public List<LinePiece> Pieces { get; } = new();
    /// <summary>Espace ajouté à chaque espace pour la justification.</summary>
    public double ExtraPerSpace { get; init; }
    /// <summary>Notes de bas de page appelées sur la ligne (w:id).</summary>
    public List<string>? Footnotes { get; set; }
    /// <summary>Fin du texte de la ligne avant alignement (retrait gauche compris, sans les espaces de fin), points.</summary>
    public double Ink { get; init; }
}

/// <summary>Paragraphe découpé en lignes pour une largeur donnée.</summary>
public sealed class ParagraphLayout
{
    public required Paragraph Paragraph { get; init; }
    public required ParagraphFormat Format { get; init; }
    public required ParagraphContent Content { get; init; }
    public required double Width { get; init; }
    /// <summary>Style du tableau et régions de la cellule qui contient le paragraphe (default hors tableau).</summary>
    public CellStyle Cell { get; init; }
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
    StyleResolver _styles;
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

    /// <summary>
    /// Nouveau résolveur après un <strong>ajout</strong> de styles ou de listes (les styles existants n'ont pas changé :
    /// les mises en forme déjà résolues restent valables).
    /// </summary>
    public void ReplaceStyles(StyleResolver styles) => _styles = styles;

    /// <summary>Numéros des listes du document (<see cref="ListNumbering"/>) ; null : « 1 » indicatif.</summary>
    public IReadOnlyDictionary<Paragraph, string>? ListNumbers { get; set; }

    /// <summary>Numéro affiché par la marque w:footnoteRef du texte d'une note (posé par l'appelant pendant sa mise en page).</summary>
    public string? NoteNumber { get; set; }

    /// <summary>
    /// Texte justifié : Word peut resserrer les espaces de la ligne pour y faire tenir un mot de plus (mesuré par COM :
    /// FrameView, espaces ramenés de 2,26 à 1,86 pt pour garder « provided »). La limite ne suit pas la largeur des
    /// espaces (22 % accepté à 10 pt, 18,8 % refusé à 12 pt) mais une quantité fixe par espace : calibrée sur le corpus,
    /// les débuts de page sont identiques entre 0,45 et 0,47 pt par espace (FrameView 6/6, Annexe 10/11), moins au-delà.
    /// OPENWRITE_JUSTIFY_SHRINK pour recalibrer.
    /// </summary>
    public static double JustifyShrink { get; set; } =
        double.TryParse(Environment.GetEnvironmentVariable("OPENWRITE_JUSTIFY_SHRINK"), NumberStyles.Float, CultureInfo.InvariantCulture, out var j) ? j : 0.46;

    /// <summary>Dépassement toléré en fin de ligne (points), à calibrer sur le corpus (OPENWRITE_FIT_TOL).</summary>
    public static double FitTolerance { get; set; } =
        double.TryParse(Environment.GetEnvironmentVariable("OPENWRITE_FIT_TOL"), NumberStyles.Float, CultureInfo.InvariantCulture, out var t) ? t : 0;

    /// <summary>Taille des minuscules en petites capitales, relative à la taille du texte (à calibrer : OPENWRITE_SMALLCAPS).</summary>
    public static double SmallCapsScale { get; set; } =
        double.TryParse(Environment.GetEnvironmentVariable("OPENWRITE_SMALLCAPS"), NumberStyles.Float, CultureInfo.InvariantCulture, out var k) ? k : 0.7;

    readonly Dictionary<RunFormat, TextStyle> _smallCache = new();

    TextStyle Style(RunFormat rf, bool small = false)
    {
        var cache = small ? _smallCache : _styleCache;
        if (cache.TryGetValue(rf, out var s)) return s;
        double size = Math.Max(1, rf.Size);
        if (rf.VerticalAlign is "superscript" or "subscript") size *= 2.0 / 3;
        if (small) size *= SmallCapsScale;
        s = new TextStyle { Format = rf, Face = _fonts.Face(rf.FontAscii ?? rf.FontHAnsi, rf.Bold, rf.Italic), Size = size, Tracking = rf.CharacterSpacing / 20.0 };
        cache[rf] = s;
        return s;
    }

    // ------------------------------------------------------------------ texte affiché

    public ParagraphContent Content(Paragraph p, CellStyle tableStyle)
    {
        var sb = new System.Text.StringBuilder();
        var styleOf = new List<int>();
        var styles = new List<TextStyle>();
        var index = new Dictionary<TextStyle, int>();
        var objects = new Dictionary<int, (double, double)>();
        var notes = new Dictionary<int, string>();
        var objectXml = new Dictionary<int, XElement>();
        var boxes = new Dictionary<int, FieldChar>();
        // champs en cours : code (w:instrText) masqué ; résultat affiché, ou remplacé par FieldValue (PAGE, NUMPAGES)
        var fields = new Stack<FieldState>();
        bool anyVisible = false;
        // correspondance avec le modèle : décalage courant, début du contenu en cours, texte (1 caractère = 1 position) ou non
        var modelStart = new List<int>();
        var modelEnd = new List<int>();
        int mo = 0, cStart = 0;
        bool perChar = false;

        void Add(string text, TextStyle st)
        {
            if (!index.TryGetValue(st, out var i)) { i = styles.Count; styles.Add(st); index[st] = i; }
            for (int k = 0; k < text.Length; k++)
            {
                modelStart.Add(perChar ? cStart + k : cStart);
                modelEnd.Add(perChar ? cStart + k + 1 : mo);
            }
            if (perChar) cStart += text.Length;
            sb.Append(text);
            for (int k = 0; k < text.Length; k++) styleOf.Add(i);
            anyVisible = true;
        }

        // petites capitales : les minuscules deviennent des capitales d'une taille réduite
        void AddSmallCaps(string text, RunFormat rf, TextStyle normal)
        {
            var small = Style(rf, small: true);
            int k = 0;
            while (k < text.Length)
            {
                bool lower = char.IsLower(text[k]);
                int e = k;
                while (e < text.Length && char.IsLower(text[e]) == lower) e++;
                Add(lower ? text[k..e].ToUpper(CultureInfo.CurrentCulture) : text[k..e], lower ? small : normal);
                k = e;
            }
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
                            bool hidden = fields.Any(f => f.InCode || f.Replaced) || rf.Hidden;
                            cStart = mo;
                            mo += ParagraphAtoms.LengthOf(c);
                            perChar = c is TextContent;
                            switch (c)
                            {
                                case FieldChar { Kind: FieldCharKind.Begin } fc:
                                    var state = new FieldState { Begin = cStart };
                                    if (!hidden && fc.FormData?.Element(W + "checkBox") is { } cb)
                                    {
                                        boxes[sb.Length] = fc;
                                        state.Covering.Add(sb.Length);
                                        Add(IsChecked(cb) ? "☒" : "☐", st);
                                    }
                                    fields.Push(state);
                                    break;
                                case FieldCode code when fields.Count > 0 && fields.Peek().InCode:
                                    fields.Peek().Code.Append(code.Value);
                                    break;
                                case FieldChar { Kind: FieldCharKind.Separate }:
                                    if (fields.Count > 0)
                                    {
                                        var f = fields.Peek();
                                        f.InCode = false;
                                        if (!fields.Skip(1).Any(x => x.InCode || x.Replaced) && FieldValue?.Invoke(f.Code.ToString().Trim()) is { } value)
                                        {
                                            for (int k = 0; k < value.Length; k++) f.Covering.Add(sb.Length + k);
                                            Add(value, st);
                                            f.Replaced = true;
                                        }
                                    }
                                    break;
                                case FieldChar:
                                    if (fields.Count > 0)
                                    {
                                        // caractères qui représentent le champ entier : du début à la fin du champ
                                        var done = fields.Pop();
                                        foreach (var k in done.Covering) { modelStart[k] = done.Begin; modelEnd[k] = mo; }
                                    }
                                    break;
                                case TextContent { Deleted: false } t when !hidden:
                                    if (rf.SmallCaps && !rf.Caps) AddSmallCaps(t.Value, rf, st);
                                    else Add(rf.Caps ? t.Value.ToUpper(CultureInfo.CurrentCulture) : t.Value, st);
                                    break;
                                case TabContent when !hidden:
                                    Add("\t", st);
                                    break;
                                case BreakContent b when !hidden:
                                    var type = b.Attributes.FirstOrDefault(a => a.Name.LocalName == "type")?.Value;
                                    Add(type == "page" ? "\f" : type == "column" ? "\v" : "\n", st);
                                    break;
                                // appel de note : son numéro (w:id, dans l'ordre du document chez Word) dans le style de l'appel ;
                                // marque dans le texte d'une note (w:footnoteRef) : idem
                                case OpaqueRunContent o when !hidden && o.Xml.Name.LocalName is "footnoteReference" or "endnoteReference" or "footnoteRef" or "endnoteRef":
                                    var id = o.Xml.Attributes().FirstOrDefault(a => a.Name.LocalName == "id")?.Value ?? NoteNumber ?? "1";
                                    if (o.Xml.Name.LocalName == "footnoteReference") notes[sb.Length] = id;
                                    Add(id, st);
                                    break;
                                case OpaqueRunContent o when !hidden && InlineObjectSize(o.Xml) is { } size:
                                    objects[sb.Length] = size;
                                    objectXml[sb.Length] = o.Xml;
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
                    default:
                        mo += ParagraphAtoms.LengthOf(inline);
                        break;
                }
            }
        }

        Walk(p.Inlines);
        var mark = Style(_styles.ParagraphMark(p, tableStyle));
        return new ParagraphContent
        {
            Text = sb.ToString(), StyleOf = styleOf.ToArray(), Styles = styles, Objects = objects, Mark = mark, Footnotes = notes,
            ObjectXml = objectXml, CheckBoxes = boxes,
            Hidden = !anyVisible && mark.Format.Hidden,
            ModelStart = modelStart.ToArray(), ModelEnd = modelEnd.ToArray(),
        };
    }

    sealed class FieldState
    {
        public readonly System.Text.StringBuilder Code = new();
        public bool InCode = true;
        public bool Replaced;
        /// <summary>Décalage du caractère de début du champ (modèle).</summary>
        public int Begin;
        /// <summary>Caractères affichés qui représentent tout le champ (case à cocher, valeur calculée).</summary>
        public readonly List<int> Covering = new();
    }

    /// <summary>
    /// Valeur d'un champ calculé par la mise en page (code → texte affiché à la place du résultat enregistré) :
    /// numéros de page des en-têtes et pieds. Null : résultats enregistrés.
    /// </summary>
    public Func<string, string?>? FieldValue { get; set; }

    /// <summary>État d'une case à cocher de formulaire : w:checked, sinon w:default.</summary>
    bool IsChecked(XElement checkBox)
    {
        static bool On(XElement? e, XNamespace W) => e is not null && (e.Attribute(W + "val")?.Value is null or "1" or "true" or "on");
        var chk = checkBox.Element(W + "checked");
        return chk is not null ? On(chk, W) : checkBox.Element(W + "default")?.Attribute(W + "val")?.Value is "1" or "true" or "on";
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
        // marges d'effet (ombre, reflet) : comptées dans la place de l'objet (Annexe : image de 105,9 pt + 0,6 pt en bas)
        var eff = inline!.Elements().FirstOrDefault(e => e.Name.LocalName == "effectExtent");
        return (Emu(ext.Attribute("cx")) + Emu(eff?.Attribute("l")) + Emu(eff?.Attribute("r")),
                Emu(ext.Attribute("cy")) + Emu(eff?.Attribute("t")) + Emu(eff?.Attribute("b")));
    }

    // ------------------------------------------------------------------ lignes

    /// <summary>Découpe le paragraphe en lignes pour une colonne de largeur <paramref name="width"/> (points).</summary>
    public ParagraphLayout Layout(Paragraph p, double width, CellStyle tableStyle = default)
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
            if (fmt != "bullet") mtext = ListNumbers is not null && ListNumbers.TryGetValue(p, out var number) ? number
                : System.Text.RegularExpressions.Regex.Replace(mtext, "%[1-9]", "1");
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

        var layout = new ParagraphLayout { Paragraph = p, Format = pf, Content = content, Width = width, Marker = marker, Cell = tableStyle };
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
            // saut de page ou de colonne en fin de paragraphe : la marque de paragraphe reste sur la ligne du saut
            // (Word 2013 et suivants ; Annexe : paragraphe « saut ¶ » en bas de page 5, page 6 commence au suivant)
            if (pos >= n && kind is LineEnd.PageBreak or LineEnd.ColumnBreak) break;
            if (pos >= n && kind is LineEnd.LineBreak)
            {
                // saut de ligne en fin de paragraphe : Word ajoute une ligne vide (la marque de paragraphe)
                layout.Lines.Add(MakeLine(content, pf, n, n, LineEnd.Paragraph, left, limit, last: true));
                break;
            }
        }
        if (TraceText is { Length: > 0 } needle && content.Text.StartsWith(needle, StringComparison.Ordinal))
        {
            Console.WriteLine($"    § « {needle} » largeur {width:0.0}, retraits {left:0.0}/{right:0.0}/{first:0.0}, alignement {pf.Alignment}");
            foreach (var l in layout.Lines)
            {
                double w = 0;
                for (int k = l.Start; k < l.End; k++) w += Advance(content, k);
                Console.WriteLine($"      {w,6:0.0} pt « {content.Text[l.Start..l.End]} »");
            }
            foreach (var st in content.Styles.Distinct())
                Console.WriteLine($"      style : {st.Face.Name} {st.Size} pt{(st.Format.Bold ? " gras" : "")}{(st.Format.Italic ? " italique" : "")}{(st.Face.Substituted ? " (remplacée)" : "")}");
        }
        return layout;
    }

    /// <summary>
    /// Largeurs de contenu d'un paragraphe (ajustement automatique des tableaux) : minimale = le plus long mot (suite de
    /// caractères sans espace ni tabulation) avec les retraits, maximale = la plus longue ligne sans coupure automatique.
    /// </summary>
    public (double Min, double Max) ContentWidths(Paragraph p, CellStyle cell = default)
    {
        var wide = Layout(p, 1e5, cell);
        if (wide.Content.Hidden) return (0, 0);
        var pf = wide.Format;
        double left = pf.IndentLeft / 20.0, right = pf.IndentRight / 20.0, first = pf.IndentFirstLine / 20.0;
        double max = wide.Lines.Count == 0 ? 0 : wide.Lines.Max(l => l.Ink) + right;
        var c = wide.Content;
        double min = 0, word = 0;
        for (int i = 0; i <= c.Text.Length; i++)
        {
            char ch = i < c.Text.Length ? c.Text[i] : ' ';
            if (ch is ' ' or Atom.Tab or Atom.LineBreak or Atom.PageBreak or Atom.ColumnBreak) { min = Math.Max(min, word); word = 0; continue; }
            word += Advance(c, i);
        }
        return (Math.Max(0, min + Math.Max(left, left + first) + right), Math.Max(0, max));
    }

    /// <summary>Mise au point : décrit la coupure des paragraphes qui commencent par ce texte (OPENWRITE_TRACE_TEXT).</summary>
    public static string? TraceText { get; set; } = Environment.GetEnvironmentVariable("OPENWRITE_TRACE_TEXT");

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
        bool justified = pf.Alignment is "both" or "distribute";
        double spaces = 0;             // largeur des espaces entre les mots de la ligne (compressibles si justifié)
        int spaceCount = 0;
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
                    double sw = Advance(c, i);
                    x += sw;
                    if (anyInk) { spaces += sw; spaceCount++; }
                    lastBreak = i + 1;     // les espaces de fin de ligne débordent sans forcer la coupure
                    continue;
            }
            double w = Advance(c, i);
            double shrink = justified ? Math.Min(JustifyShrink * spaceCount, spaces) : 0;
            if (TraceText is { Length: > 0 } tt && (tt == "*" || c.Text.StartsWith(tt, StringComparison.Ordinal)) && x + w > limit + 0.01 && x + w <= limit + shrink + 0.01)
                Console.WriteLine($"      compression acceptée à « {c.Text[Math.Max(start, i - 12)..(i + 1)]} » : {(x + w - limit) / Math.Max(1, spaceCount):0.000} pt par espace ({spaceCount} espaces)");
            if (x + w > limit + FitTolerance + shrink + 0.01 && anyInk)
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
        double asc = 0, desc = 0, lead = 0, objects = 0;
        void Measure(TextStyle st)
        {
            asc = Math.Max(asc, st.Face.Ascent * st.Size);
            desc = Math.Max(desc, st.Face.Descent * st.Size);
            lead = Math.Max(lead, st.Face.ExternalLeading * st.Size);
        }
        int prevStyle = -1;
        for (int i = start; i < end; i++)
        {
            // une tabulation ne donne pas sa hauteur à la ligne (Word ; Syllabus, en-tête de la première page : ligne
            // faite d'une tabulation en Palatino 12 pt, hauteur de la marque de paragraphe en 10 pt)
            if (c.Text[i] == Atom.Tab) continue;
            int s = c.StyleOf[i];
            // objet en ligne : la ligne prend la hauteur de l'objet si elle dépasse celle du texte, sans y ajouter le
            // descendant de la police (Annexe : ligne d'une image = hauteur de l'image, mesuré dans Word)
            if (c.Text[i] == Atom.Object && c.Objects.TryGetValue(i, out var o)) objects = Math.Max(objects, o.H);
            if (s != prevStyle) { Measure(c.Styles[s]); prevStyle = s; }
        }
        if (last || end == start) Measure(c.Mark);
        double natural = Math.Max(asc + desc + lead, objects);
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
        double fit = pf.LineRule is "exact" or "atLeast" ? h : Math.Min(h, natural);
        List<string>? lineNotes = null;
        foreach (var (pos, id) in c.Footnotes)
            if (pos >= start && pos < end) (lineNotes ??= new()).Add(id);
        return new LineBox { Start = start, End = end, Height = h, FitHeight = fit, Baseline = baseline, EndKind = kind, ExtraPerSpace = extra, Footnotes = lineNotes, Ink = ink }.WithPieces(line.Pieces);
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
