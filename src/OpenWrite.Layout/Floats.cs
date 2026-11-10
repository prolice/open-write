using OpenWrite.Core.Drawing;
using OpenWrite.Core.Model;

namespace OpenWrite.Layout;

// Objets flottants (lot 7) : placement des objets ancrés sur la page, habillage du texte autour d'eux, contenu des
// zones de texte. Un objet est ancré à un paragraphe ; il va sur la page où commence ce paragraphe. Le texte d'avant
// l'ancre sur la même page en tient compte aussi : la page est alors remise en page avec l'objet connu dès le départ.

/// <summary>Objet flottant placé sur une page (points, repère de la page).</summary>
public sealed class FloatBox
{
    public required DrawingObject Object { get; init; }
    /// <summary>Paragraphe d'ancrage (null : objet d'un en-tête ou d'un pied de page).</summary>
    public Paragraph? Anchor { get; init; }
    /// <summary>En-tête ou pied d'où vient l'objet (r:id), pour retrouver ses images ; null : corps du document.</summary>
    public string? PartId { get; init; }
    /// <summary>Cadre de l'objet (wp:extent).</summary>
    public double X { get; init; }
    public double Y { get; init; }
    public double W { get; init; }
    public double H { get; init; }
    /// <summary>Lignes des zones de texte de l'objet, placées sur la page (indice de la zone, ligne).</summary>
    public List<(int Box, PlacedLine Line)> TextLines { get; } = new();
    /// <summary>Cellules de tableaux des zones de texte.</summary>
    public List<CellBox> TextCells { get; } = new();

    /// <summary>Le texte s'écarte de l'objet (habillage carré, rapproché, au travers, haut et bas).</summary>
    public bool Wraps => Object.Wrap is WrapKind.Square or WrapKind.Tight or WrapKind.Through or WrapKind.TopAndBottom;

    /// <summary>Zone écartée du texte : cadre, marges d'effet et distances au texte.</summary>
    public (double L, double T, double R, double B) Exclusion
    {
        get
        {
            var o = Object;
            return (X - o.Effect.L - o.Distance.L, Y - o.Effect.T - o.Distance.T,
                    X + W + o.Effect.R + o.Distance.R, Y + H + o.Effect.B + o.Distance.B);
        }
    }

    /// <summary>
    /// Étendue horizontale écartée pour une bande verticale [top, bottom[ : rectangle (carré, haut et bas), ou polygone
    /// d'habillage (rapproché, au travers) ; null si la bande ne touche pas l'objet.
    /// </summary>
    public (double L, double R)? Blocked(double top, double bottom)
    {
        var (l, t, r, b) = Exclusion;
        if (bottom <= t + 0.01 || top >= b - 0.01) return null;
        if (Object.Wrap is WrapKind.Tight or WrapKind.Through && Object.WrapPolygon is { Count: >= 3 } poly)
        {
            // points du polygone (fractions du cadre) qui tombent dans la bande, et croisements des côtés avec ses bords
            double minX = double.MaxValue, maxX = double.MinValue;
            var pts = poly.Select(p => (X: X + p.X * W, Y: Y + p.Y * H)).ToList();
            for (int i = 0; i < pts.Count; i++)
            {
                var a = pts[i];
                var c = pts[(i + 1) % pts.Count];
                void Take(double x) { minX = Math.Min(minX, x); maxX = Math.Max(maxX, x); }
                if (a.Y >= top && a.Y <= bottom) Take(a.X);
                foreach (var yy in new[] { top, bottom })
                    if ((a.Y - yy) * (c.Y - yy) < 0) Take(a.X + (c.X - a.X) * (yy - a.Y) / (c.Y - a.Y));
                // côté qui traverse toute la bande
                if (Math.Min(a.Y, c.Y) <= top && Math.Max(a.Y, c.Y) >= bottom && Math.Abs(c.Y - a.Y) > 1e-9)
                {
                    Take(a.X + (c.X - a.X) * (top - a.Y) / (c.Y - a.Y));
                    Take(a.X + (c.X - a.X) * (bottom - a.Y) / (c.Y - a.Y));
                }
            }
            if (minX > maxX) return null;
            return (minX - Object.Distance.L, maxX + Object.Distance.R);
        }
        return (l, r);
    }
}

public sealed partial class DocumentLayout
{
    /// <summary>
    /// Segments libres d'une colonne [<paramref name="colL"/>, <paramref name="colR"/>] (repère de la page) pour une bande
    /// verticale, autour des objets qui habillent ; aucun segment et la hauteur à descendre si un objet « haut et bas »
    /// (ou un objet sans place à côté) occupe la bande.
    /// </summary>
    internal static (List<(double L, double R)> Segments, double Down) FreeSegments(IEnumerable<FloatBox> floats, double colL, double colR, double top, double bottom)
    {
        var segs = new List<(double L, double R)> { (colL, colR) };
        double down = 0;
        foreach (var f in floats)
        {
            if (!f.Wraps) continue;
            if (f.Blocked(top, bottom) is not { } blocked) continue;
            var (bl, br) = blocked;
            if (br <= colL || bl >= colR) continue;                     // objet hors de la colonne
            double objBottom = f.Exclusion.B;
            if (f.Object.Wrap == WrapKind.TopAndBottom)
            {
                return (new(), Math.Max(0.5, objBottom - top));
            }
            // côtés où le texte peut aller
            var side = f.Object.WrapSide;
            if (side == "largest") side = bl - colL >= colR - br ? "left" : "right";
            var cut = new List<(double L, double R)>();
            foreach (var (sl, sr) in segs)
            {
                // le texte ne va que d'un côté : l'autre côté est retiré aussi
                double xl = side == "right" ? Math.Min(sl, bl) : bl;
                double xr = side == "left" ? Math.Max(sr, br) : br;
                if (side == "right") xl = double.MinValue;
                if (side == "left") xr = double.MaxValue;
                if (sr <= xl || sl >= xr) { cut.Add((sl, sr)); continue; }
                if (sl < xl) cut.Add((sl, xl));
                if (sr > xr) cut.Add((xr, sr));
            }
            segs = cut;
            down = down == 0 ? objBottom - top : Math.Min(down, objBottom - top);
        }
        return (segs, Math.Max(down, 0));
    }

    /// <summary>
    /// Cadre d'un objet ancré sur la page : repères horizontaux (page, marge, colonne, caractère, marges gauche et
    /// droite) et verticaux (page, marge, paragraphe, ligne, marges haute et basse), alignement ou décalage.
    /// </summary>
    internal static (double X, double Y, double W, double H) Position(DrawingObject o, Page page, SectionGeometry g, double colX, double colW,
        double paraTop, double lineTop, double charX)
    {
        double w = o.Width, h = o.Height;
        double pageW = g.PageWidth, pageH = g.PageHeight;
        double top = page.ContentTop, bottom = g.PageHeight - (page.FirstOfSection ? g.FirstBottom : g.Bottom);
        // taille relative (wp14:sizeRelH/V)
        (double Base, double Size) HRef(string from) => from switch
        {
            "page" => (0, pageW),
            "margin" => (g.Left, pageW - g.Left - g.Right),
            "leftMargin" or "insideMargin" => (0, g.Left),
            "rightMargin" or "outsideMargin" => (pageW - g.Right, g.Right),
            "character" => (charX, 0),
            _ => (colX, colW),                                          // column
        };
        (double Base, double Size) VRef(string from) => from switch
        {
            "page" => (0, pageH),
            "margin" => (top, bottom - top),
            "topMargin" or "insideMargin" => (0, top),
            "bottomMargin" or "outsideMargin" => (bottom, pageH - bottom),
            "line" => (lineTop, 0),
            _ => (paraTop, 0),                                          // paragraph
        };
        if (o.RelWidth is { } rw) w = rw.Pct * HRef(rw.From == "page" ? "page" : "margin").Size;
        if (o.RelHeight is { } rh) h = rh.Pct * VRef(rh.From == "page" ? "page" : "margin").Size;

        var (hb, hs) = HRef(o.H.RelativeFrom);
        double x = o.H.Align switch
        {
            "left" or "inside" => hb,
            "center" => hb + (hs - w) / 2,
            "right" or "outside" => hb + hs - w,
            _ => o.H.Percent is { } ph ? hb + ph * hs : hb + o.H.Offset,
        };
        var (vb, vs) = VRef(o.V.RelativeFrom);
        double y = o.V.Align switch
        {
            "top" or "inside" => vb,
            "center" => vb + (vs - h) / 2,
            "bottom" or "outside" => vb + vs - h,
            _ => o.V.Percent is { } pv ? vb + pv * vs : vb + o.V.Offset,
        };
        return (x, y, w, h);
    }

    /// <summary>
    /// Place les objets ancrés au paragraphe sur la page du curseur (en haut du paragraphe : <paramref name="paraTop"/>),
    /// avec le contenu de leurs zones de texte. Renvoie vrai si un objet qui habille touche des lignes déjà placées sur
    /// la page (la page est alors à refaire).
    /// </summary>
    bool PlaceAnchors(ParagraphLayout pl, Page page, SectionGeometry g, double colX, double colW, double paraTop, int firstPlacedLine)
    {
        bool conflict = false;
        foreach (var (pos, o) in pl.Content.Anchors)
        {
            if (page.Floats.Any(f => ReferenceEquals(f.Object.Source, o.Source))) continue;
            // ligne de l'ancre (repère « ligne » et « caractère ») : première ligne du paragraphe à défaut de mieux
            double charX = colX + pl.Format.IndentLeft / 20.0;
            double lineTop = paraTop;
            double yy = paraTop;
            foreach (var l in pl.Lines)
            {
                yy += l.SameRow ? 0 : l.Gap;
                if (pos >= l.Start && (pos < l.End || l == pl.Lines[^1]))
                {
                    lineTop = yy;
                    foreach (var piece in l.Pieces)
                        if (pos >= piece.Start && pos <= piece.Start + piece.Length) charX = colX + piece.X;
                    break;
                }
                yy += l.Advance - (l.SameRow ? 0 : l.Gap);
            }
            var (x, y, w, h) = Position(o, page, g, colX, colW, paraTop, lineTop, charX);
            var box = new FloatBox { Object = o, Anchor = pl.Paragraph, X = x, Y = y, W = w, H = h };
            PlaceTextBoxes(box, g);
            page.Floats.Add(box);
            if (!box.Wraps) continue;
            // texte déjà placé sur la page (avant ce paragraphe) que l'objet recouvre
            var (el, et, er, eb) = box.Exclusion;
            for (int k = 0; k < firstPlacedLine && k < page.Lines.Count; k++)
            {
                var line = page.Lines[k];
                var lb = line.Paragraph.Lines[line.Line];
                double lt = line.Y, lbot = line.Y + lb.Height;
                if (lbot <= et + 0.01 || lt >= eb - 0.01) continue;
                double lx0 = line.X + (lb.Pieces.Count > 0 ? lb.Pieces[0].X : 0), lx1 = line.X + (lb.Pieces.Count > 0 ? lb.Pieces[^1].X + lb.Pieces[^1].Width : 0);
                if (box.Object.Wrap == WrapKind.TopAndBottom ? lx1 > el && lx0 < er || line.X + line.Paragraph.Width > el && line.X < er : lx1 > el + 0.01 && lx0 < er - 0.01)
                {
                    conflict = true;
                    break;
                }
            }
        }
        return conflict;
    }

    /// <summary>Contenu des zones de texte d'un objet (formes, groupes) placé dans leurs cadres, en points de la page.</summary>
    void PlaceTextBoxes(FloatBox box, SectionGeometry g)
    {
        void Walk(Visual v, double ox, double oy)
        {
            switch (v)
            {
                case GroupVisual grp:
                    foreach (var c in grp.Children) Walk(c, ox + grp.X, oy + grp.Y);
                    break;
                case ShapeVisual s when s.TextBox >= 0 && s.TextBox < box.Object.Source.TextBoxes.Count:
                {
                    var story = box.Object.Source.TextBoxes[s.TextBox];
                    double x = ox + s.X + s.Insets.L, width = Math.Max(1, s.W - s.Insets.L - s.Insets.R);
                    double top = oy + s.Y + s.Insets.T;
                    double h = StoryHeight(story, width);
                    double room = s.H - s.Insets.T - s.Insets.B;
                    if (s.TextAnchor == "ctr") top += Math.Max(0, (room - h) / 2);
                    else if (s.TextAnchor == "b") top += Math.Max(0, room - h);
                    var tmp = new Page { Geometry = g, Start = default };
                    PlaceStory(story, top, g, tmp, cache: true, left: x, width: width);
                    foreach (var l in tmp.Lines) box.TextLines.Add((s.TextBox, l));
                    box.TextCells.AddRange(tmp.Cells);
                    break;
                }
            }
        }
        Walk(box.Object.Root, box.X, box.Y);
    }

    /// <summary>
    /// Fonction de place des lignes d'un paragraphe habillé : haut du paragraphe sur la page, colonne, objets de la page ;
    /// les lignes qui passent sous le bas de la page ne sont pas gênées (elles iront à la page suivante).
    /// </summary>
    static LineSpace? SpaceFor(Page page, double paraTop, double colX, double colW)
    {
        var floats = page.Floats.Where(f => f.Wraps).ToList();
        if (floats.Count == 0) return null;
        double maxBottom = floats.Max(f => f.Exclusion.B);
        if (maxBottom <= paraTop) return null;
        double contentBottom = page.ContentBottom;
        return (y, h) =>
        {
            double top = paraTop + y;
            if (top >= contentBottom) return (new[] { (0.0, colW) }, 0);
            var (segs, down) = FreeSegments(floats, colX, colX + colW, top, top + h);
            return (segs.Select(s => (s.L - colX, s.R - colX)).ToList(), down);
        };
    }
}
