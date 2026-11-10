using System.Xml.Linq;

namespace OpenWrite.Core.Model;

// Modèle de texte en flux — PROTOTYPE du lot 1 (PLAN.md §3.1, §8.2) : il valide la stratégie de préservation
// avant d'écrire le modèle définitif.
//
// Principe : chaque nœud compris (paragraphe, segment, tableau, ligne, cellule, contrôle de contenu) garde
// TOUS les attributs XML de son élément d'origine (w14:paraId, rsid…) et ses propriétés (pPr, rPr, tblPr…)
// sous forme d'un PropertyBag qui conserve l'élément XML complet ; les accès typés lisent et modifient cet
// élément à la bonne place du schéma. Tout ce qui n'est pas compris devient un îlot opaque (XElement recopié
// tel quel, déplacé avec son paragraphe). Le modèle ne dépend pas de l'Open XML SDK (System.Xml.Linq seulement).

/// <summary>Suite de blocs : corps du document, cellule, en-tête, note, commentaire, zone de texte…</summary>
public sealed class Story
{
    public List<Block> Blocks { get; } = new();
    /// <summary>Propriétés de la dernière section (w:sectPr en fin de corps) ; null hors du corps.</summary>
    public PropertyBag? FinalSection { get; set; }

    public IEnumerable<Paragraph> Paragraphs() => Blocks.SelectMany(ParagraphsOf);

    static IEnumerable<Paragraph> ParagraphsOf(Block b) => b switch
    {
        Paragraph p => new[] { p },
        Table t => t.Items.OfType<TableRow>().SelectMany(r => r.Items.OfType<TableCell>()).SelectMany(c => c.Content.Paragraphs()),
        BlockContentControl c => c.Content.Paragraphs(),
        _ => Array.Empty<Paragraph>(),
    };
}

public abstract class Block
{
    /// <summary>Attributs de l'élément d'origine, dans l'ordre (déclarations d'espaces de noms comprises).</summary>
    public List<XAttribute> Attributes { get; } = new();
}

public sealed class Paragraph : Block
{
    /// <summary>w:pPr (style, alignement, retraits, numérotation, marque de paragraphe, saut de section…).</summary>
    public PropertyBag Props { get; set; } = new("pPr");
    public List<Inline> Inlines { get; } = new();

    public string? StyleId => Props.Val("pStyle");

    /// <summary>Texte du paragraphe (segments, conteneurs ; tabulations et retours à la ligne en caractères).</summary>
    public string Text => string.Concat(Inlines.Select(i => i.Text));

    public IEnumerable<Run> Runs() => Inlines.SelectMany(i => i.Runs());
}

public sealed class Table : Block
{
    public PropertyBag Props { get; set; } = new("tblPr");
    /// <summary>w:tblGrid (largeurs des colonnes), conservé tel quel dans le prototype.</summary>
    public XElement? Grid { get; set; }
    /// <summary>Lignes, et éléments compris entre les lignes (signets, contrôles de contenu…) en îlots opaques.</summary>
    public List<object> Items { get; } = new();
}

public sealed class TableRow
{
    public List<XAttribute> Attributes { get; } = new();
    /// <summary>w:tblPrEx (exceptions de propriétés du tableau pour cette ligne).</summary>
    public PropertyBag? PropsException { get; set; }
    public PropertyBag Props { get; set; } = new("trPr");
    /// <summary>Cellules, et îlots opaques entre les cellules.</summary>
    public List<object> Items { get; } = new();
}

public sealed class TableCell
{
    public List<XAttribute> Attributes { get; } = new();
    public PropertyBag Props { get; set; } = new("tcPr");
    public Story Content { get; } = new();
}

/// <summary>Contrôle de contenu de niveau bloc (w:sdt) : propriétés conservées, contenu modélisé.</summary>
public sealed class BlockContentControl : Block
{
    public XElement? SdtPr { get; set; }
    public XElement? SdtEndPr { get; set; }
    public List<XAttribute> ContentAttributes { get; } = new();
    public Story Content { get; } = new();
}

/// <summary>Élément de niveau bloc non compris : recopié tel quel.</summary>
public sealed class OpaqueBlock : Block
{
    public OpaqueBlock(XElement xml) => Xml = xml;
    public XElement Xml { get; }
}

// ---------------------------------------------------------------- niveau ligne

public abstract class Inline
{
    public List<XAttribute> Attributes { get; } = new();
    public abstract string Text { get; }
    public virtual IEnumerable<Run> Runs() => Array.Empty<Run>();
}

public sealed class Run : Inline
{
    public PropertyBag Props { get; set; } = new("rPr");
    public List<RunContent> Content { get; } = new();
    public override string Text => string.Concat(Content.Select(c => c switch
    {
        TextContent t when !t.Deleted => t.Value,
        TabContent => "\t",
        BreakContent => "\n",
        _ => "",
    }));
    public override IEnumerable<Run> Runs() => new[] { this };
}

/// <summary>
/// Conteneur de segments : lien (w:hyperlink), révision (w:ins, w:del, w:moveFrom, w:moveTo), w:smartTag,
/// w:customXml, w:fldSimple, w:dir, w:bdo. Les éléments de tête (smartTagPr, customXmlPr) sont conservés.
/// </summary>
public sealed class InlineContainer : Inline
{
    public InlineContainer(XName name) => Name = name;
    public XName Name { get; }
    public List<XElement> Head { get; } = new();
    public List<Inline> Children { get; } = new();
    public override string Text => string.Concat(Children.Select(c => c.Text));
    public override IEnumerable<Run> Runs() => Children.SelectMany(c => c.Runs());
}

/// <summary>Contrôle de contenu dans la ligne (w:sdt) : propriétés conservées, contenu modélisé.</summary>
public sealed class InlineContentControl : Inline
{
    public XElement? SdtPr { get; set; }
    public XElement? SdtEndPr { get; set; }
    public List<XAttribute> ContentAttributes { get; } = new();
    public List<Inline> Content { get; } = new();
    public override string Text => string.Concat(Content.Select(c => c.Text));
    public override IEnumerable<Run> Runs() => Content.SelectMany(c => c.Runs());
}

/// <summary>Élément non compris dans un paragraphe (signet, commentaire, équation, mc:AlternateContent…).</summary>
public sealed class OpaqueInline : Inline
{
    public OpaqueInline(XElement xml) => Xml = xml;
    public XElement Xml { get; }
    public override string Text => "";
}

public abstract class RunContent
{
    public List<XAttribute> Attributes { get; } = new();
}

/// <summary>w:t (ou w:delText pour du texte supprimé avec suivi des modifications).</summary>
public sealed class TextContent : RunContent
{
    public string Value { get; set; } = "";
    public bool Deleted { get; init; }
    /// <summary>
    /// Valeur lue dans le fichier : tant que le texte n'a pas changé, il est réécrit exactement comme Word l'avait écrit
    /// (Word 2010 écrit parfois « &lt;w:t&gt; &lt;/w:t&gt; » sans xml:space ; on n'ajoute rien à ce qu'on n'a pas touché).
    /// </summary>
    public string? LoadedValue { get; init; }
}

public sealed class TabContent : RunContent { }

/// <summary>w:br (attributs type et clear conservés) ou w:cr.</summary>
public sealed class BreakContent : RunContent
{
    public bool CarriageReturn { get; init; }
}

/// <summary>Contenu de segment non compris (dessin, champ, symbole, note…).</summary>
public sealed class OpaqueRunContent : RunContent
{
    public OpaqueRunContent(XElement xml) => Xml = xml;
    public XElement Xml { get; }
}
