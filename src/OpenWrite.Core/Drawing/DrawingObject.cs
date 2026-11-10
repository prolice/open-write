using System.Xml.Linq;
using OpenWrite.Core.Model;

namespace OpenWrite.Core.Drawing;

// Vue lue d'un dessin (lot 7) : le fichier reste la référence (DrawingContent.Xml), cette vue sert à la mise en page,
// à l'affichage et à l'édition (qui réécrit le XML). Unités : points ; angles en degrés.

/// <summary>Habillage du texte autour d'un objet.</summary>
public enum WrapKind
{
    /// <summary>Objet dans la ligne de texte.</summary>
    Inline,
    /// <summary>Devant ou derrière le texte, sans habillage (wp:wrapNone, VML sans w10:wrap).</summary>
    None,
    Square,
    Tight,
    Through,
    TopAndBottom,
}

/// <summary>
/// Position d'un objet ancré sur un axe : repère (page, margin, column, character, paragraph, line, leftMargin,
/// rightMargin, topMargin, bottomMargin, insideMargin, outsideMargin), alignement (left, center, right, inside, outside,
/// top, bottom) ou décalage en points, ou pourcentage du repère (wp14:pctPosHOffset, en fraction).
/// </summary>
public sealed record AnchorAxis(string RelativeFrom, string? Align, double Offset, double? Percent = null);

/// <summary>Remplissage : none, solid, gradient (arrêts position 0..1, couleur, opacité ; angle en degrés), picture.</summary>
public sealed record FillSpec(string Kind, string? Color = null, double Alpha = 1)
{
    public List<(double Pos, string Color, double Alpha)>? Stops { get; init; }
    public double Angle { get; init; }
    /// <summary>Dégradé radial (path="circle"/"rect" DrawingML, type="gradientRadial" VML).</summary>
    public bool Radial { get; init; }
    /// <summary>Image de remplissage (r:embed).</summary>
    public string? ImageId { get; init; }
}

/// <summary>Trait : couleur RRGGBB, épaisseur (points), tirets (solid, dash, dot, dashDot, lgDash…), extrémités (flèches).</summary>
public sealed record LineSpec(string Color, double Width, string Dash = "solid", double Alpha = 1)
{
    public string? HeadEnd { get; init; }
    public string? TailEnd { get; init; }
    /// <summary>Trait double, triple… (a:ln/@cmpd : dbl, thickThin, thinThick, tri).</summary>
    public string? Compound { get; init; }
}

/// <summary>Élément graphique d'un dessin, dans le repère de son parent (points).</summary>
public abstract class Visual
{
    public double X { get; set; }
    public double Y { get; set; }
    public double W { get; set; }
    public double H { get; set; }
    /// <summary>Rotation autour du centre, en degrés (sens horaire).</summary>
    public double Rotation { get; set; }
    public bool FlipH { get; set; }
    public bool FlipV { get; set; }
}

/// <summary>Image (pic:pic, v:imagedata) : r:embed ou r:link, rognage (fractions du bord), forme du cadre.</summary>
public sealed class PictureVisual : Visual
{
    public string? ImageId { get; init; }
    public string? LinkId { get; init; }
    public (double L, double T, double R, double B) Crop { get; init; }
    /// <summary>Forme qui découpe l'image (prstGeom ; rect en général).</summary>
    public string Geometry { get; init; } = "rect";
    public LineSpec? Line { get; init; }
    /// <summary>Corrections simples : luminosité et contraste (-1..1), niveaux de gris, opacité.</summary>
    public double Brightness { get; init; }
    public double Contrast { get; init; }
    public bool Grayscale { get; init; }
    public double Alpha { get; init; } = 1;
}

/// <summary>Forme (wps:wsp, VML) : géométrie, remplissage, trait, zone de texte éventuelle.</summary>
public sealed class ShapeVisual : Visual
{
    /// <summary>Forme prédéfinie DrawingML (rect, ellipse, roundRect…), ou null si <see cref="Custom"/> / <see cref="Paths"/>.</summary>
    public string? Preset { get; init; } = "rect";
    public Dictionary<string, string>? Adjust { get; init; }
    /// <summary>Géométrie libre DrawingML (a:custGeom).</summary>
    public XElement? Custom { get; init; }
    /// <summary>Chemins déjà calculés (VML), dans le cadre 0..W × 0..H.</summary>
    public List<ShapePath>? Paths { get; init; }
    public FillSpec? Fill { get; init; }
    public LineSpec? Line { get; init; }
    /// <summary>Indice de la zone de texte dans <see cref="DrawingContent.TextBoxes"/>, -1 sans texte.</summary>
    public int TextBox { get; init; } = -1;
    /// <summary>Marges intérieures du texte (points).</summary>
    public (double L, double T, double R, double B) Insets { get; init; } = (7.2, 3.6, 7.2, 3.6);
    /// <summary>Ancrage vertical du texte : t, ctr, b.</summary>
    public string TextAnchor { get; init; } = "t";
    /// <summary>Texte vertical (vert, vert270, eaVert…), null : horizontal.</summary>
    public string? TextDirection { get; init; }
    /// <summary>La forme s'agrandit pour son texte (a:spAutoFit, mso-fit-shape-to-text).</summary>
    public bool AutoFit { get; init; }
    /// <summary>Texte WordArt (v:textpath, filigrane) : texte, police.</summary>
    public (string Text, string Font)? WordArt { get; init; }
}

/// <summary>Groupe (wpg:wgp, wpc:wpc, v:group) : enfants déjà ramenés dans le repère du groupe.</summary>
public sealed class GroupVisual : Visual
{
    public List<Visual> Children { get; } = new();
    /// <summary>Fond d'une zone de dessin (wpc:bg).</summary>
    public FillSpec? Fill { get; init; }
}

/// <summary>Objet qu'OpenWrite ne dessine pas lui-même (graphique, SmartArt, OLE) : aperçu image éventuel.</summary>
public sealed class FallbackVisual : Visual
{
    /// <summary>chart, smartart, ole, unknown.</summary>
    public required string Kind { get; init; }
    public string? ImageId { get; init; }
    /// <summary>SmartArt : r:id de la partie dessin (dsp) qui en garde la forme calculée.</summary>
    public string? DrawingId { get; init; }
}

/// <summary>Dessin lu : cadre, position et habillage s'il est ancré, contenu graphique.</summary>
public sealed class DrawingObject
{
    public required DrawingContent Source { get; init; }
    public bool Inline { get; init; }
    /// <summary>Taille du cadre (wp:extent ou style VML), points.</summary>
    public double Width { get; set; }
    public double Height { get; set; }
    /// <summary>Marges d'effet (ombre, trait épais…) autour du cadre, points.</summary>
    public (double L, double T, double R, double B) Effect { get; init; }
    public AnchorAxis H { get; init; } = new("column", null, 0);
    public AnchorAxis V { get; init; } = new("paragraph", null, 0);
    public WrapKind Wrap { get; init; } = WrapKind.Inline;
    /// <summary>Côtés habillés : bothSides, left, right, largest.</summary>
    public string WrapSide { get; init; } = "bothSides";
    /// <summary>Distance entre l'objet et le texte (haut, bas, gauche, droite ; points).</summary>
    public (double T, double B, double L, double R) Distance { get; init; }
    /// <summary>Polygone d'habillage (rapproché, au travers) en fraction du cadre (0..1).</summary>
    public List<(double X, double Y)>? WrapPolygon { get; init; }
    public bool BehindText { get; init; }
    /// <summary>Ordre d'empilement (wp:anchor/@relativeHeight, z-index VML).</summary>
    public long Z { get; init; }
    public bool AllowOverlap { get; init; } = true;
    public bool LayoutInCell { get; init; } = true;
    public bool Locked { get; init; }
    /// <summary>Largeur ou hauteur relative (wp14:sizeRelH/V : fraction du repère, repère).</summary>
    public (double Pct, string From)? RelWidth { get; init; }
    public (double Pct, string From)? RelHeight { get; init; }
    public string? Name { get; init; }
    public string? Description { get; init; }
    public bool Hidden { get; init; }
    /// <summary>Lu depuis du VML (w:pict, mc:Fallback).</summary>
    public bool Vml { get; init; }
    /// <summary>Filigrane (forme WordArt de l'en-tête) : dessiné à part.</summary>
    public bool Watermark { get; init; }
    /// <summary>Filet horizontal VML (o:hr) : part de la largeur du texte (1 = toute la largeur) ; null sinon.</summary>
    public double? HorizontalRule { get; init; }
    public required Visual Root { get; init; }
}
