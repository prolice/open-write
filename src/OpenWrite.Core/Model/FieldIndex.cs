using System.Text;
using System.Xml.Linq;

namespace OpenWrite.Core.Model;

/// <summary>
/// Champ trouvé dans le texte : complexe (marqueurs w:fldChar début / séparateur / fin, code en w:instrText) ou simple
/// (w:fldSimple). Vue calculée par <see cref="FieldIndex"/> : modifier un champ, c'est modifier ses marqueurs.
/// </summary>
public sealed class Field
{
    public FieldChar? Begin { get; internal set; }
    public FieldChar? Separate { get; internal set; }
    public FieldChar? End { get; internal set; }
    /// <summary>Champ simple : son conteneur (code dans l'attribut w:instr, résultat = contenu).</summary>
    public InlineContainer? Simple { get; internal init; }
    public Paragraph? BeginParagraph { get; internal init; }
    public Paragraph? EndParagraph { get; internal set; }
    public Field? Parent { get; internal init; }
    public List<Field> Children { get; } = new();

    internal readonly List<object> CodeParts = new();           // texte de code ou champ imbriqué
    internal readonly StringBuilder ResultText = new();
    internal bool InResult;
    /// <summary>Segments de texte du résultat qui appartiennent directement à ce champ (pas à un champ imbriqué).</summary>
    public List<(Run Run, TextContent Text)> ResultRuns { get; } = new();

    public bool IsComplete => Simple is not null || (Begin is not null && End is not null);

    /// <summary>Code du champ ; un champ imbriqué dans le code apparaît entre accolades, comme avec Alt+F9 dans Word.</summary>
    public string Code => Simple is not null ? SimpleCode : string.Concat(CodeParts.Select(p => p is Field f ? "{" + f.Code + "}" : (string)p));

    string SimpleCode => Simple!.Attributes.FirstOrDefault(a => a.Name.LocalName == "instr")?.Value ?? "";

    public FieldInstruction Instruction => FieldInstruction.Parse(Code);
    public string Type => Instruction.Type;

    /// <summary>Résultat affiché (texte, tabulations, retours ; « \r » entre deux paragraphes).</summary>
    public string Result => ResultText.ToString();

    /// <summary>Case à cocher de formulaire hérité (FORMCHECKBOX) ; null pour les autres champs.</summary>
    public FormCheckBox? CheckBox => Type == "FORMCHECKBOX" && Begin?.FormData is { } ff && ff.Element(Begin.Ns + "checkBox") is not null
        ? new FormCheckBox(ff, Begin.Ns) : null;

    public override string ToString() => $"{{{Code.Trim()}}} → « {Result} »";
}

/// <summary>
/// Case à cocher d'un formulaire hérité (w:ffData/w:checkBox) : état = w:checked s'il est présent, sinon w:default.
/// Modifie l'élément du marqueur de début : l'aller-retour garde tout le reste (nom, aide, macros d'entrée/sortie).
/// </summary>
public sealed class FormCheckBox
{
    readonly XElement _ffData;
    readonly XNamespace _w;

    internal FormCheckBox(XElement ffData, XNamespace w) { _ffData = ffData; _w = w; }

    XElement Box => _ffData.Element(_w + "checkBox")!;

    public string? Name => _ffData.Element(_w + "name")?.Attribute(_w + "val")?.Value;
    public bool Enabled => _ffData.Element(_w + "enabled") is not { } e || IsOn(e);
    public bool Default => Box.Element(_w + "default") is { } d && IsOn(d);

    static bool IsOn(XElement e) => e.Attribute(e.Name.Namespace + "val")?.Value is null or "1" or "true" or "on";

    public bool Checked
    {
        get => Box.Element(_w + "checked") is { } c ? IsOn(c) : Default;
        set
        {
            Box.Element(_w + "checked")?.Remove();
            if (value == Default) return;                        // l'état par défaut suffit : rien à écrire
            // ordre du schéma (CT_FFCheckBox) : size | sizeAuto, default, checked
            var c = new XElement(_w + "checked", value ? null : new XAttribute(_w + "val", "0"));
            var before = Box.Element(_w + "default") ?? Box.Element(_w + "size") ?? Box.Element(_w + "sizeAuto");
            if (before is null) Box.AddFirst(c); else before.AddAfterSelf(c);
        }
    }

    /// <summary>Caractère affiché pour la case (☒ cochée, ☐ vide).</summary>
    public string Glyph => Checked ? "☒" : "☐";
}

/// <summary>Index des champs d'une suite de blocs, dans l'ordre du document.</summary>
public sealed class FieldIndex
{
    public List<Field> All { get; } = new();
    public IEnumerable<Field> Roots => All.Where(f => f.Parent is null);
    /// <summary>Anomalies (marqueur sans début, champ non fermé…) : le document reste ouvert et ces marqueurs sont conservés.</summary>
    public List<string> Problems { get; } = new();

    public static FieldIndex Build(Story story)
    {
        var idx = new FieldIndex();
        var stack = new List<Field>();

        void AddResult(string s)
        {
            for (int i = stack.Count - 1; i >= 0 && stack[i].InResult; i--) stack[i].ResultText.Append(s);
        }

        Field Open(Paragraph? p, FieldChar? begin, InlineContainer? simple)
        {
            var parent = stack.Count > 0 ? stack[^1] : null;
            var f = new Field { Begin = begin, Simple = simple, BeginParagraph = p, Parent = parent, InResult = simple is not null };
            if (parent is not null)
            {
                parent.Children.Add(f);
                if (!parent.InResult) parent.CodeParts.Add(f);
            }
            idx.All.Add(f);
            stack.Add(f);
            return f;
        }

        foreach (var e in FlowWalker.Walk(story))
        {
            switch (e.Kind)
            {
                case FlowEventKind.ContainerStart when e.Item is InlineContainer { Name.LocalName: "fldSimple" } box:
                    Open(e.Paragraph, null, box);
                    break;
                case FlowEventKind.ContainerEnd when e.Item is InlineContainer { Name.LocalName: "fldSimple" } box:
                    if (stack.Count > 0 && stack[^1].Simple == box) { stack[^1].EndParagraph = e.Paragraph; stack.RemoveAt(stack.Count - 1); }
                    break;
                case FlowEventKind.ParagraphEnd:
                    if (stack.Count > 0) AddResult("\r");
                    break;
                case FlowEventKind.Content:
                    switch (e.Item)
                    {
                        case FieldChar { Kind: FieldCharKind.Begin } fc:
                            Open(e.Paragraph, fc, null);
                            break;
                        case FieldChar { Kind: FieldCharKind.Separate } fc:
                            if (stack.Count == 0 || stack[^1].Simple is not null) { idx.Problems.Add("séparateur de champ sans début"); break; }
                            stack[^1].Separate = fc;
                            stack[^1].InResult = true;
                            break;
                        case FieldChar fc:
                            if (stack.Count == 0 || stack[^1].Simple is not null) { idx.Problems.Add("fin de champ sans début"); break; }
                            stack[^1].End = fc;
                            stack[^1].EndParagraph = e.Paragraph;
                            stack.RemoveAt(stack.Count - 1);
                            break;
                        case FieldCode code when !code.Deleted:
                            if (stack.Count == 0) idx.Problems.Add($"code de champ hors d'un champ : « {code.Value.Trim()} »");
                            else if (!stack[^1].InResult) stack[^1].CodeParts.Add(code.Value);
                            break;
                        case TextContent { Deleted: false } t:
                            if (stack.Count > 0 && stack[^1].InResult && e.Run is not null) stack[^1].ResultRuns.Add((e.Run, t));
                            AddResult(t.Value);
                            break;
                        case TabContent: AddResult("\t"); break;
                        case BreakContent: AddResult("\n"); break;
                    }
                    break;
            }
        }
        foreach (var f in stack) idx.Problems.Add($"champ non fermé : {{{f.Code.Trim()}}}");
        return idx;
    }
}
