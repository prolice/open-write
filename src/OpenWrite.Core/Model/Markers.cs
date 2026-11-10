using System.Xml.Linq;

namespace OpenWrite.Core.Model;

// Marqueurs du texte en flux : caractères de champ et signets (lot 1).
//
// Décision (PLAN.md §10.3) : comme dans Word, un champ est une suite de marqueurs dans le texte — début, code,
// séparateur, résultat, fin — et non un nœud qui contiendrait son code et son résultat. Un champ peut s'étendre sur
// plusieurs paragraphes (table des matières), s'imbriquer (IF { MERGEFIELD }) et chaque marqueur porte la mise en
// forme de son segment : le représenter à plat garde l'aller-retour exact. La structure (code, résultat, imbrication)
// est calculée par FieldIndex. Même principe pour les signets (BookmarkIndex).

/// <summary>Valeur commune des marqueurs : attributs d'origine et espace de noms WordprocessingML du document.</summary>
public interface IMarker
{
    List<XAttribute> Attributes { get; }
    XNamespace Ns { get; }
}

static class MarkerAttributes
{
    public static string? Get(this IMarker m, string local) => m.Attributes.FirstOrDefault(a => a.Name == m.Ns + local)?.Value;

    public static void Set(this IMarker m, string local, string? value)
    {
        var i = m.Attributes.FindIndex(a => a.Name == m.Ns + local);
        if (value is null) { if (i >= 0) m.Attributes.RemoveAt(i); return; }
        if (i >= 0) m.Attributes[i] = new XAttribute(m.Ns + local, value);
        else m.Attributes.Add(new XAttribute(m.Ns + local, value));
    }
}

public enum FieldCharKind { Begin, Separate, End }

/// <summary>
/// w:fldChar : début, séparateur ou fin d'un champ complexe. Ses enfants (w:ffData des champs de formulaire hérités,
/// w:fldData, w:numberingChange) sont conservés tels quels ; FormCheckBox en donne une vue modifiable.
/// </summary>
public sealed class FieldChar : RunContent, IMarker
{
    public FieldChar(XNamespace ns) => Ns = ns;
    public XNamespace Ns { get; }
    public List<XElement> Children { get; } = new();

    public FieldCharKind Kind => this.Get("fldCharType") switch
    {
        "begin" => FieldCharKind.Begin,
        "separate" => FieldCharKind.Separate,
        _ => FieldCharKind.End,
    };

    /// <summary>w:dirty : Word doit mettre le champ à jour à l'ouverture.</summary>
    public bool Dirty
    {
        get => this.Get("dirty") is "1" or "true" or "on";
        set => this.Set("dirty", value ? "true" : null);
    }

    /// <summary>Données d'un champ de formulaire hérité (FORMCHECKBOX, FORMTEXT, FORMDROPDOWN) ; null sinon.</summary>
    public XElement? FormData => Children.FirstOrDefault(c => c.Name == Ns + "ffData");
}

/// <summary>w:instrText (code de champ) ou w:delInstrText (code supprimé avec suivi des modifications).</summary>
public sealed class FieldCode : RunContent
{
    public string Value { get; set; } = "";
    public string? LoadedValue { get; init; }
    public bool Deleted { get; init; }
}

/// <summary>w:bookmarkStart : w:id, w:name, et pour un signet de colonnes de tableau w:colFirst/w:colLast.</summary>
public sealed class BookmarkStart : Inline, IMarker
{
    public BookmarkStart(XNamespace ns) => Ns = ns;
    public XNamespace Ns { get; }
    public string Id { get => this.Get("id") ?? ""; set => this.Set("id", value); }
    public string Name { get => this.Get("name") ?? ""; set => this.Set("name", value); }
    public override string Text => "";
}

/// <summary>w:bookmarkEnd.</summary>
public sealed class BookmarkEnd : Inline, IMarker
{
    public BookmarkEnd(XNamespace ns) => Ns = ns;
    public XNamespace Ns { get; }
    public string Id { get => this.Get("id") ?? ""; set => this.Set("id", value); }
    public override string Text => "";
}

/// <summary>Marqueur placé entre deux blocs (signet au niveau du corps ou d'une cellule).</summary>
public sealed class InlineMarkerBlock : Block
{
    public InlineMarkerBlock(Inline marker) => Marker = marker;
    public Inline Marker { get; }
}
