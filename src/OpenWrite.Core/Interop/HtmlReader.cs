using System.Globalization;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using OpenWrite.Core.Editing;

namespace OpenWrite.Core.Interop;

/// <summary>
/// Lecture du HTML collé depuis un navigateur ou un courriel (format « HTML Format » du presse-papiers, lot 3) :
/// paragraphes et titres, retours à la ligne, gras, italique, souligné, barré, exposant / indice, couleur, taille et
/// police (balises et styles en ligne), alignement ; cellules de tableau séparées par des tabulations, une ligne par
/// paragraphe. Scripts, styles et commentaires ignorés ; espaces regroupés comme dans un navigateur.
/// </summary>
public static partial class HtmlReader
{
    /// <summary>Partie collée d'un « HTML Format » (entre &lt;!--StartFragment--&gt; et &lt;!--EndFragment--&gt;, sinon le corps).</summary>
    public static string Fragment(string cfHtml)
    {
        int s = cfHtml.IndexOf("<!--StartFragment", StringComparison.OrdinalIgnoreCase);
        int e = cfHtml.IndexOf("<!--EndFragment", StringComparison.OrdinalIgnoreCase);
        if (s >= 0 && e > s)
        {
            s = cfHtml.IndexOf("-->", s, StringComparison.Ordinal) + 3;
            return cfHtml[s..e];
        }
        int b = cfHtml.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
        if (b >= 0) b = cfHtml.IndexOf('>', b) + 1;
        int be = cfHtml.LastIndexOf("</body", StringComparison.OrdinalIgnoreCase);
        int h = cfHtml.IndexOf('<');
        return b > 0 && be > b ? cfHtml[b..be] : h >= 0 ? cfHtml[h..] : cfHtml;
    }

    static readonly HashSet<string> Blocks = new(StringComparer.OrdinalIgnoreCase)
    { "p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "li", "tr", "blockquote", "pre", "ul", "ol", "table", "section", "article", "header", "footer", "dt", "dd" };

    [GeneratedRegex(@"<!--.*?-->|<(/?)([a-zA-Z][a-zA-Z0-9:]*)([^>]*)>|([^<]+)", RegexOptions.Singleline)]
    private static partial Regex Token();

    [GeneratedRegex(@"([a-zA-Z-]+)\s*=\s*(""[^""]*""|'[^']*'|[^\s>]+)")]
    private static partial Regex Attribute();

    sealed record Fmt(Fragments.CharFormat Char, string? Align);

    public static List<Fragments.ParaSpec> Read(string html, out bool lastHasMark)
    {
        var paras = new List<Fragments.ParaSpec>();
        var cur = new Fragments.ParaSpec();
        var stack = new Stack<(string Tag, Fmt Fmt)>();
        var fmt = new Fmt(new Fragments.CharFormat(), null);
        int skip = 0;
        bool pre = false;
        bool needSpace = false;
        lastHasMark = false;

        void Break()
        {
            if (cur.Runs.Count > 0) paras.Add(cur);
            cur = new Fragments.ParaSpec { Align = fmt.Align };
            needSpace = false;
        }
        void Raw(string t)
        {
            if (skip > 0) return;
            cur.Align ??= fmt.Align;
            if (cur.Runs.Count > 0 && cur.Runs[^1].Format == fmt.Char) cur.Runs[^1] = (cur.Runs[^1].Text + t, fmt.Char);
            else cur.Runs.Add((t, fmt.Char));
            needSpace = false;
        }
        void Text(string t)
        {
            if (skip > 0) return;
            t = WebUtility.HtmlDecode(t);
            if (!pre)
            {
                var sb = new StringBuilder();
                foreach (var ch in t)
                {
                    if (ch is ' ' or '\t' or '\r' or '\n') { needSpace = true; continue; }
                    if (needSpace && (cur.Runs.Count > 0 || sb.Length > 0)) sb.Append(' ');
                    needSpace = false;
                    sb.Append(ch);
                }
                t = sb.ToString();
            }
            else t = t.Replace("\r\n", "\n").Replace('\n', '\v');
            if (t.Length == 0) return;
            cur.Align ??= fmt.Align;
            if (cur.Runs.Count > 0 && cur.Runs[^1].Format == fmt.Char) cur.Runs[^1] = (cur.Runs[^1].Text + t, fmt.Char);
            else cur.Runs.Add((t, fmt.Char));
        }

        foreach (Match m in Token().Matches(html))
        {
            if (m.Groups[4].Success) { Text(m.Groups[4].Value); continue; }
            if (!m.Groups[2].Success) continue;            // commentaire
            bool closing = m.Groups[1].Value == "/";
            var tag = m.Groups[2].Value.ToLowerInvariant();
            if (tag.Contains(':')) tag = tag[(tag.IndexOf(':') + 1)..];   // o:p, v:shape…
            var attrs = m.Groups[3].Value;
            bool selfClosing = attrs.TrimEnd().EndsWith('/');
            if (tag is "script" or "style" or "head" or "title" or "xml") { skip += closing ? -1 : selfClosing ? 0 : 1; skip = Math.Max(0, skip); continue; }
            if (tag == "br" && !closing) { Raw("\v"); continue; }
            if (tag is "td" or "th")
            {
                if (closing) Raw("\t");
                continue;
            }
            if (tag == "img" || selfClosing) continue;
            if (closing)
            {
                // dépile jusqu'à la balise ouverte correspondante
                if (stack.Any(x => x.Tag == tag))
                {
                    while (stack.Count > 0)
                    {
                        var (t, f) = stack.Pop();
                        fmt = f;
                        if (t == tag) break;
                    }
                }
                if (tag == "pre") pre = false;
                if (Blocks.Contains(tag))
                {
                    // une ligne de tableau : la dernière tabulation (fin de cellule) est retirée
                    if (tag == "tr" && cur.Runs.Count > 0 && cur.Runs[^1].Text.EndsWith('\t'))
                        cur.Runs[^1] = (cur.Runs[^1].Text[..^1], cur.Runs[^1].Format);
                    Break();
                }
                continue;
            }
            stack.Push((tag, fmt));
            if (Blocks.Contains(tag)) Break();
            if (tag == "pre") pre = true;
            var c = fmt.Char;
            var align = fmt.Align;
            switch (tag)
            {
                case "b" or "strong": c = c with { Bold = true }; break;
                case "i" or "em" or "cite" or "var": c = c with { Italic = true }; break;
                case "u" or "ins": c = c with { Underline = true }; break;
                case "s" or "strike" or "del": c = c with { Strike = true }; break;
                case "sup": c = c with { VertAlign = "superscript" }; break;
                case "sub": c = c with { VertAlign = "subscript" }; break;
                case "code" or "tt" or "kbd" or "samp": c = c with { Font = "Courier New" }; break;
                case "h1": c = c with { Bold = true, Size = 16 }; break;
                case "h2": c = c with { Bold = true, Size = 13 }; break;
                case "h3" or "h4" or "h5" or "h6": c = c with { Bold = true, Size = 12 }; break;
            }
            foreach (Match a in Attribute().Matches(attrs))
            {
                var name = a.Groups[1].Value.ToLowerInvariant();
                var value = WebUtility.HtmlDecode(a.Groups[2].Value.Trim('"', '\''));
                if (name == "align" && Blocks.Contains(tag)) align = Align(value) ?? align;
                if (name == "color" && tag == "font") c = c with { Color = Color(value) ?? c.Color };
                if (name == "face" && tag == "font") c = c with { Font = value.Split(',')[0].Trim().Trim('\'', '"') };
                if (name != "style") continue;
                foreach (var decl in value.Split(';'))
                {
                    int k = decl.IndexOf(':');
                    if (k < 0) continue;
                    var prop = decl[..k].Trim().ToLowerInvariant();
                    var v = decl[(k + 1)..].Trim().ToLowerInvariant().Replace("!important", "").Trim();
                    switch (prop)
                    {
                        case "font-weight": c = c with { Bold = v is "bold" or "bolder" || int.TryParse(v, out var wgt) && wgt >= 600 }; break;
                        case "font-style": c = c with { Italic = v is "italic" or "oblique" }; break;
                        case "text-decoration" or "text-decoration-line":
                            if (v.Contains("underline")) c = c with { Underline = true };
                            if (v.Contains("line-through")) c = c with { Strike = true };
                            if (v == "none") c = c with { Underline = false, Strike = false };
                            break;
                        case "color": c = c with { Color = Color(v) ?? c.Color }; break;
                        case "font-size": c = c with { Size = Size(v) ?? c.Size }; break;
                        case "font-family": c = c with { Font = decl[(k + 1)..].Split(',')[0].Trim().Trim('\'', '"') }; break;
                        case "vertical-align": c = c with { VertAlign = v == "super" ? "superscript" : v == "sub" ? "subscript" : c.VertAlign }; break;
                        case "text-align" when Blocks.Contains(tag): align = Align(v) ?? align; break;
                    }
                }
            }
            fmt = new Fmt(c, align);
            if (Blocks.Contains(tag)) cur.Align = align;
        }
        if (cur.Runs.Count > 0) paras.Add(cur);
        // sans bloc final ouvert : dernier paragraphe partiel
        lastHasMark = false;
        return paras;
    }

    static string? Align(string v) => v.Trim().ToLowerInvariant() switch
    {
        "center" => "center",
        "right" or "end" => "right",
        "justify" => "both",
        "left" or "start" => "left",
        _ => null,
    };

    internal static string? Color(string v)
    {
        v = v.Trim().ToLowerInvariant();
        if (v.StartsWith('#'))
        {
            var hex = v[1..];
            if (hex.Length == 3) hex = string.Concat(hex.Select(ch => $"{ch}{ch}"));
            return hex.Length == 6 && int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _) ? hex.ToUpperInvariant() : null;
        }
        var m = Regex.Match(v, @"rgba?\((\d+)\D+(\d+)\D+(\d+)");
        if (m.Success) return string.Concat(m.Groups.Values.Skip(1).Take(3).Select(g => Math.Clamp(int.Parse(g.Value, CultureInfo.InvariantCulture), 0, 255).ToString("X2")));
        return v switch
        {
            "black" => "000000", "red" => "FF0000", "green" => "008000", "blue" => "0000FF", "gray" or "grey" => "808080",
            "white" => "FFFFFF", "navy" => "000080", "maroon" => "800000", "purple" => "800080", "orange" => "FFA500", _ => null,
        };
    }

    static double? Size(string v)
    {
        var m = Regex.Match(v, @"([\d.]+)\s*(pt|px|em|rem)?");
        if (!m.Success || !double.TryParse(m.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var n)) return null;
        double pt = m.Groups[2].Value switch { "px" => n * 0.75, "em" or "rem" => n * 12, _ => n };
        return pt is >= 1 and <= 400 ? Math.Round(pt * 2) / 2 : null;
    }
}
