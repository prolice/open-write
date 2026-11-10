using System.Text;

namespace OpenWrite.Core.Model;

/// <summary>
/// Code de champ analysé (ECMA-376 partie 1, §17.16) : type (PAGE, REF, TOC, MERGEFIELD…), arguments et commutateurs.
/// Les commutateurs généraux \* (format), \@ (date) et \# (nombre) prennent toujours un argument ; les autres en prennent
/// un s'il est entre guillemets (TOC \o "1-3", HYPERLINK \l "ancre") — REF _Ref123 \h n'en a pas.
/// </summary>
public sealed class FieldInstruction
{
    public string Type { get; private init; } = "";
    public IReadOnlyList<string> Arguments { get; private init; } = Array.Empty<string>();
    /// <summary>Commutateurs dans l'ordre, avec leur argument éventuel (\* peut apparaître plusieurs fois).</summary>
    public IReadOnlyList<(string Switch, string? Argument)> Switches { get; private init; } = Array.Empty<(string, string?)>();

    public bool Has(string sw) => Switches.Any(s => string.Equals(s.Switch, sw, StringComparison.OrdinalIgnoreCase));
    public string? Arg(string sw) => Switches.FirstOrDefault(s => string.Equals(s.Switch, sw, StringComparison.OrdinalIgnoreCase)).Argument;

    static readonly HashSet<string> GeneralSwitches = new() { "\\*", "\\@", "\\#" };

    public static FieldInstruction Parse(string code)
    {
        var tokens = Tokenize(code);
        if (tokens.Count == 0) return new FieldInstruction();
        var args = new List<string>();
        var switches = new List<(string, string?)>();
        for (int i = 1; i < tokens.Count; i++)
        {
            var (text, quoted) = tokens[i];
            if (!quoted && text.Length >= 2 && text[0] == '\\')
            {
                string? arg = null;
                if (i + 1 < tokens.Count && !(tokens[i + 1] is { Quoted: false } n && n.Text.StartsWith('\\'))
                    && (GeneralSwitches.Contains(text) || tokens[i + 1].Quoted))
                    arg = tokens[++i].Text;
                switches.Add((text, arg));
            }
            else args.Add(text);
        }
        return new FieldInstruction { Type = tokens[0].Text.ToUpperInvariant(), Arguments = args, Switches = switches };
    }

    /// <summary>Mots du code : séparés par des blancs ; « "…" » est un seul mot (\" et \\ y sont des échappements).</summary>
    static List<(string Text, bool Quoted)> Tokenize(string code)
    {
        var res = new List<(string, bool)>();
        int i = 0;
        while (i < code.Length)
        {
            if (char.IsWhiteSpace(code[i])) { i++; continue; }
            var sb = new StringBuilder();
            if (code[i] == '"')
            {
                i++;
                while (i < code.Length && code[i] != '"')
                {
                    if (code[i] == '\\' && i + 1 < code.Length && code[i + 1] is '"' or '\\') i++;
                    sb.Append(code[i++]);
                }
                i++;                                   // guillemet fermant
                res.Add((sb.ToString(), true));
            }
            else
            {
                while (i < code.Length && !char.IsWhiteSpace(code[i]) && code[i] != '"') sb.Append(code[i++]);
                res.Add((sb.ToString(), false));
            }
        }
        return res;
    }
}
