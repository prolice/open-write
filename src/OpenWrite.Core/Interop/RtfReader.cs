using System.Globalization;
using System.Text;
using OpenWrite.Core.Editing;

namespace OpenWrite.Core.Interop;

/// <summary>
/// Lecture du RTF collé depuis Word, LibreOffice, WordPad… (lot 3) : paragraphes, alignement, gras, italique,
/// souligné, barré, exposant / indice, taille, police, couleur, tabulations et retours à la ligne, caractères Unicode.
/// Les destinations non gérées (images, en-têtes, styles, codes de champs…) sont ignorées ; les résultats de champs
/// et le texte des cellules de tableau sont gardés (une cellule = un paragraphe).
/// </summary>
public static class RtfReader
{
    static RtfReader() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    sealed record State
    {
        public bool Bold, Italic, Underline, Strike;
        public int FontSize = 24;            // demi-points
        public int Font = -1, Color;
        public string? VertAlign;
        public int SkipUnicode = 1;
        /// <summary>Destination ignorée (et tout ce qu'elle contient).</summary>
        public bool Skip;
        /// <summary>Destination particulière : table des polices ou des couleurs.</summary>
        public string? Table;
    }

    static readonly HashSet<string> SkippedDestinations = new()
    {
        "stylesheet", "info", "pict", "object", "header", "headerl", "headerr", "headerf", "footer", "footerl", "footerr", "footerf",
        "footnote", "themedata", "colorschememapping", "latentstyles", "datastore", "xmlnstbl", "listtable", "listoverridetable",
        "rsidtbl", "generator", "mmathPr", "pgdsctbl", "shp", "nonshppict", "shppict", "fldinst", "filetbl", "revtbl", "annotation",
        "atnauthor", "atnid", "bkmkstart", "bkmkend", "template", "docvar", "xe", "tc", "pntext", "pntxta", "pntxtb", "listtext",
    };

    /// <summary>Paragraphes lus ; <paramref name="lastHasMark"/> : le texte se termine par une fin de paragraphe.</summary>
    public static List<Fragments.ParaSpec> Read(string rtf, out bool lastHasMark)
    {
        var paras = new List<Fragments.ParaSpec>();
        var fonts = new Dictionary<int, string>();
        var colors = new List<string?>();
        var stack = new Stack<State>();
        var st = new State();
        var cur = new Fragments.ParaSpec();
        var text = new StringBuilder();
        Fragments.CharFormat? textFormat = null;
        string? align = null;
        var tableText = new StringBuilder();
        int tableFont = -1;
        int red = 0, green = 0, blue = 0;
        bool colorSet = false;
        var encoding = Encoding.GetEncoding(1252);
        int pendingSkip = 0;
        bool hasMark = false;

        Fragments.CharFormat Format() => new()
        {
            Bold = st.Bold, Italic = st.Italic, Underline = st.Underline, Strike = st.Strike,
            Size = st.FontSize / 2.0, Font = fonts.TryGetValue(st.Font, out var f) ? f : null,
            Color = st.Color > 0 && st.Color < colors.Count ? colors[st.Color] : null, VertAlign = st.VertAlign,
        };
        void FlushText()
        {
            if (text.Length > 0 && textFormat is not null) cur.Runs.Add((text.ToString(), textFormat));
            text.Clear();
            textFormat = null;
        }
        void Emit(string s)
        {
            if (st.Skip) return;
            if (st.Table == "fonttbl") { tableText.Append(s); return; }
            if (st.Table == "colortbl") return;
            var f = Format();
            if (textFormat is not null && textFormat != f) FlushText();
            textFormat = f;
            text.Append(s);
            hasMark = false;
        }
        void EndParagraph()
        {
            if (st.Skip || st.Table is not null) return;
            FlushText();
            cur.Align = align;
            paras.Add(cur);
            cur = new Fragments.ParaSpec();
            hasMark = true;
        }

        int i = 0, n = rtf.Length;
        while (i < n)
        {
            char ch = rtf[i];
            if (ch == '{') { stack.Push(st with { }); i++; continue; }
            if (ch == '}')
            {
                if (stack.Count > 0) st = stack.Pop();
                i++;
                continue;
            }
            if (ch is '\r' or '\n') { i++; continue; }
            if (ch != '\\')
            {
                if (pendingSkip > 0) { pendingSkip--; i++; continue; }
                if (st.Table == "fonttbl" && ch == ';')
                {
                    if (tableFont >= 0) fonts[tableFont] = tableText.ToString().Trim();
                    tableText.Clear();
                    i++;
                    continue;
                }
                if (st.Table == "colortbl" && ch == ';')
                {
                    colors.Add(colorSet ? $"{red:X2}{green:X2}{blue:X2}" : null);
                    red = green = blue = 0;
                    colorSet = false;
                    i++;
                    continue;
                }
                int start = i;
                while (i < n && rtf[i] is not ('\\' or '{' or '}' or '\r' or '\n') && !(st.Table is not null && rtf[i] == ';')) i++;
                Emit(rtf[start..i]);
                continue;
            }
            // séquence de contrôle
            i++;
            if (i >= n) break;
            char c = rtf[i];
            if (!char.IsAsciiLetter(c))
            {
                i++;
                switch (c)
                {
                    case '\'':
                        if (i + 2 <= n && byte.TryParse(rtf.AsSpan(i, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var b))
                        {
                            i += 2;
                            if (pendingSkip > 0) pendingSkip--;
                            else Emit(encoding.GetString(new[] { b }));
                        }
                        break;
                    case '*': st.Skip = true; break;
                    case '~': Emit(" "); break;
                    case '-': break;
                    case '_': Emit("‑"); break;
                    case '\\' or '{' or '}': Emit(c.ToString()); break;
                    case '\r' or '\n': EndParagraph(); break;
                }
                continue;
            }
            int ws = i;
            while (i < n && char.IsAsciiLetter(rtf[i])) i++;
            var word = rtf[ws..i];
            int? param = null;
            int ps = i;
            if (i < n && (rtf[i] == '-' || char.IsAsciiDigit(rtf[i])))
            {
                i++;
                while (i < n && char.IsAsciiDigit(rtf[i])) i++;
                param = int.TryParse(rtf.AsSpan(ps, i - ps), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var pv) ? pv : null;
            }
            if (i < n && rtf[i] == ' ') i++;
            bool on = param is null or not 0;
            if (SkippedDestinations.Contains(word)) { st.Skip = true; continue; }
            switch (word)
            {
                case "fonttbl": st.Table = "fonttbl"; break;
                case "colortbl": st.Table = "colortbl"; break;
                case "ansicpg" when param is { } cp:
                    try { encoding = Encoding.GetEncoding(cp); } catch (Exception) { }
                    break;
                case "f" when st.Table == "fonttbl": tableFont = param ?? -1; tableText.Clear(); break;
                case "f": st.Font = param ?? -1; break;
                case "red": red = param ?? 0; colorSet = true; break;
                case "green": green = param ?? 0; colorSet = true; break;
                case "blue": blue = param ?? 0; colorSet = true; break;
                case "fs": st.FontSize = param ?? 24; break;
                case "cf": st.Color = param ?? 0; break;
                case "b": st.Bold = on; break;
                case "i": st.Italic = on; break;
                case "ul": st.Underline = on; break;
                case "uld" or "uldb" or "ulw" or "ulth" or "ulwave" or "uldash": st.Underline = true; break;
                case "ulnone": st.Underline = false; break;
                case "strike" or "striked": st.Strike = on; break;
                case "super": st.VertAlign = "superscript"; break;
                case "sub": st.VertAlign = "subscript"; break;
                case "nosupersub": st.VertAlign = null; break;
                case "plain": st = st with { Bold = false, Italic = false, Underline = false, Strike = false, FontSize = 24, Font = -1, Color = 0, VertAlign = null }; break;
                case "pard": align = null; break;
                case "qc": align = "center"; break;
                case "qr": align = "right"; break;
                case "qj": align = "both"; break;
                case "ql": align = null; break;
                case "par" or "sect": EndParagraph(); break;
                case "cell": EndParagraph(); break;
                case "line": Emit("\v"); break;
                case "tab": Emit("\t"); break;
                case "emdash": Emit("—"); break;
                case "endash": Emit("–"); break;
                case "bullet": Emit("•"); break;
                case "lquote": Emit("‘"); break;
                case "rquote": Emit("’"); break;
                case "ldblquote": Emit("“"); break;
                case "rdblquote": Emit("”"); break;
                case "uc": st.SkipUnicode = param ?? 1; break;
                case "u" when param is { } u:
                    Emit(((char)(u < 0 ? u + 65536 : u)).ToString());
                    pendingSkip = st.SkipUnicode;
                    break;
            }
        }
        FlushText();
        // texte après le dernier \par : paragraphe partiel ; sinon la marque du dernier paragraphe est comprise
        if (cur.Runs.Count > 0) { cur.Align = align; paras.Add(cur); hasMark = false; }
        lastHasMark = hasMark;
        return paras;
    }
}
