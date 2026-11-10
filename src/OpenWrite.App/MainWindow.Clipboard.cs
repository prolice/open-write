using System.Globalization;
using System.Net;
using System.Text;
using System.Windows;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Interop;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;
using OpenWrite.Docx;
using OpenWrite.Layout;
using Atom = OpenWrite.Layout.Atom;

namespace OpenWrite.App;

// Lot 3 (version 0.4) : presse-papiers avec mise en forme. Copier dépose le texte, le RTF et le HTML (lus par Word,
// LibreOffice, les courriels…) et, pour un collage dans le même document, le morceau du modèle lui-même (mise en forme,
// styles, tableaux, champs intacts). Coller prend, dans l'ordre : ce morceau interne, le RTF, le HTML, le texte ;
// Ctrl+Maj+V colle le texte seul.
public partial class MainWindow
{
    /// <summary>Format privé : jeton du dernier morceau copié (le presse-papiers n'a pas changé depuis).</summary>
    const string InternalFormat = "OpenWrite.Fragment";
    (string Token, DocxPackage Package, Fragment Fragment)? _internalClip;

    /// <summary>Paragraphes de la sélection avec leurs segments de même mise en forme (texte affiché).</summary>
    (List<(ParagraphFormat Format, List<(string Text, RunFormat Format)> Runs)> Paras, bool EndsWithMark) SelectedParagraphs()
    {
        var paras = new List<(ParagraphFormat, List<(string, RunFormat)>)>();
        if (!HasSelection || _layout is null || Ordered() is not { } sel || Locate(sel.From) is not { } a || Locate(sel.To) is not { } b) return (paras, false);
        ParagraphLayout? prev = null;
        List<(string, RunFormat)>? runs = null;
        for (LineRef? r = a.Ref; r is { } cur && cur.CompareTo(b.Ref) <= 0; r = Src.Step(cur, 1))
        {
            var l = Src.Line(cur);
            var box = l.Paragraph.Lines[l.Line];
            var c = l.Paragraph.Content;
            if (!ReferenceEquals(prev, l.Paragraph))
            {
                runs = new List<(string, RunFormat)>();
                paras.Add((l.Paragraph.Format, runs));
                prev = l.Paragraph;
            }
            int s = Math.Clamp(cur.Equals(a.Ref) ? sel.From.Index : box.Start, 0, c.Text.Length);
            int e = Math.Clamp(cur.Equals(b.Ref) ? sel.To.Index : box.EndKind is LineEnd.Wrap or LineEnd.Paragraph ? box.End : box.End + 1, s, c.Text.Length);
            for (int k = s; k < e; k++)
            {
                char ch = c.Text[k];
                if (ch is Atom.Object or Atom.ColumnBreak or Atom.PageBreak) continue;
                var f = c.Styles[c.StyleOf[k]].Format;
                var t = ch == Atom.LineBreak ? "\v" : ch.ToString();
                if (runs!.Count > 0 && ReferenceEquals(runs[^1].Item2, f)) runs[^1] = (runs[^1].Item1 + t, f);
                else runs.Add((t, f));
            }
        }
        // sélection terminée au début d'un paragraphe : la marque du précédent est comprise
        bool endsWithMark = sel.To.Index == 0 && paras.Count > 1 && paras[^1].Item2.Count == 0;
        if (endsWithMark) paras.RemoveAt(paras.Count - 1);
        return (paras, endsWithMark);
    }

    internal void Copy()
    {
        var text = SelectedText();
        if (text.Length == 0) return;
        var data = new DataObject();
        data.SetText(text, TextDataFormat.UnicodeText);
        if (CellSelection() is { } cells)
        {
            // cellules : tableau HTML pour les autres programmes, morceau de tableau pour le même document
            try { data.SetData(DataFormats.Html, CfHtml(CellSelectionHtml(cells))); }
            catch (Exception ex) when (ex is ArgumentException or InvalidOperationException) { }
            _internalClip = null;
            if (_pkg is { } p)
            {
                var token = Guid.NewGuid().ToString("N");
                _internalClip = (token, p, new Fragment(new List<Block> { TableOps.Extract(cells.Table, cells.R0, cells.C0, cells.R1, cells.C1) }, false));
                data.SetData(InternalFormat, token);
            }
            try { Clipboard.SetDataObject(data, copy: true); }
            catch (System.Runtime.InteropServices.ExternalException) { StatusText.Text = "Le presse-papiers est occupé par un autre programme : réessayez."; }
            return;
        }
        var (paras, endsWithMark) = SelectedParagraphs();
        try
        {
            data.SetData(DataFormats.Rtf, ToRtf(paras, endsWithMark));
            data.SetData(DataFormats.Html, CfHtml(ToHtml(paras)));
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException) { }
        _internalClip = null;
        if (EditRoot is { } body && _pkg is { } pkg && Ordered() is { } sel && ToDocPos(sel.From) is { } from && ToDocPos(sel.To) is { } to
            && SupportedRange(from, to, quiet: true))
        {
            try
            {
                var token = Guid.NewGuid().ToString("N");
                _internalClip = (token, pkg, Fragments.Extract(body, from, to));
                data.SetData(InternalFormat, token);
            }
            catch (Exception ex) when (ex is ArgumentException or NotSupportedException) { _internalClip = null; }
        }
        try { Clipboard.SetDataObject(data, copy: true); }
        catch (System.Runtime.InteropServices.ExternalException) { StatusText.Text = "Le presse-papiers est occupé par un autre programme : réessayez."; }
    }

    void Cut()
    {
        if (!HasSelection) return;
        Copy();
        DeleteSelection("Couper");
    }

    /// <summary>Colle le contenu du presse-papiers (avec sa mise en forme, sauf <paramref name="textOnly"/>).</summary>
    internal void Paste(bool textOnly = false)
    {
        // le presse-papiers peut être verrouillé un instant par le programme qui vient d'y écrire : quelques essais
        for (int attempt = 0; ; attempt++)
        {
            try { PasteOnce(textOnly); return; }
            catch (System.Runtime.InteropServices.ExternalException) when (attempt < 5) { Thread.Sleep(40); }
            catch (System.Runtime.InteropServices.ExternalException) { StatusText.Text = "Le presse-papiers est occupé par un autre programme : réessayez."; return; }
        }
    }

    /// <summary>
    /// Contenu texte d'un format du presse-papiers : chaîne, ou flux / octets selon le programme qui l'a déposé ; juste
    /// après une copie, la donnée annoncée peut manquer ou être vide un instant (quelques essais).
    /// </summary>
    static string? ReadString(string format, Encoding bytesEncoding)
    {
        for (int attempt = 0; attempt < 6; attempt++)
        {
            switch (Clipboard.GetData(format))
            {
                case string { Length: > 0 } str: return str;
                case System.IO.MemoryStream ms: return bytesEncoding.GetString(ms.ToArray()).TrimEnd('\0');
                case byte[] bytes: return bytesEncoding.GetString(bytes).TrimEnd('\0');
            }
            Thread.Sleep(30);
        }
        return null;
    }

    void PasteOnce(bool textOnly)
    {
        // sur une sélection de cellules : collage au début de la première cellule
        if (CellSelection() is { } cells && CellStart(cells.Table, cells.R0, cells.C0) is { } start && FromDocPos(start) is { } first)
        {
            _anchor = null;
            _caret = first;
        }
        Fragment? fragment = null;
        string? text = null;
        string source = "texte";
        {
            if (!textOnly && Clipboard.ContainsData(InternalFormat) && Clipboard.GetData(InternalFormat) is string token
                && _internalClip is { } clip && clip.Token == token && ReferenceEquals(clip.Package, _pkg))
            {
                fragment = clip.Fragment;
                source = "OpenWrite";
            }
            // tableau HTML (Excel, page web, autre traitement de texte) : un vrai tableau, avant le RTF qui le met à plat
            else if (!textOnly && Clipboard.ContainsData(DataFormats.Html) && ReadString(DataFormats.Html, Encoding.UTF8) is { Length: > 0 } tableHtml
                && HtmlTable(HtmlReader.Fragment(FixHtmlEncoding(tableHtml))) is { } pasted)
            {
                fragment = pasted;
                source = "tableau HTML";
            }
            else if (!textOnly && Clipboard.ContainsData(DataFormats.Rtf) && ReadString(DataFormats.Rtf, Encoding.Latin1) is { Length: > 0 } rtf)
            {
                var paras = RtfReader.Read(rtf, out var mark);
                if (paras.Count > 0) { fragment = Fragments.Build(paras, mark, Body?.Paragraphs().FirstOrDefault()?.Props.Ns); source = "RTF"; }
            }
            else if (!textOnly && Clipboard.ContainsData(DataFormats.Html) && ReadString(DataFormats.Html, Encoding.UTF8) is { Length: > 0 } html)
            {
                var paras = HtmlReader.Read(HtmlReader.Fragment(FixHtmlEncoding(html)), out var mark);
                if (paras.Count > 0) { fragment = Fragments.Build(paras, mark, Body?.Paragraphs().FirstOrDefault()?.Props.Ns); source = "HTML"; }
            }
            if (fragment is null && Clipboard.ContainsText()) text = Clipboard.GetText();
        }
        if (fragment is not null && !fragment.IsEmpty) { PasteFragment(fragment, source); return; }
        if (!string.IsNullOrEmpty(text)) ReplaceSelection(text, "Coller");
    }

    void PasteFragment(Fragment fragment, string source)
    {
        if (EditRoot is not { } body || Ordered() is not { } sel) return;
        if (ToDocPos(sel.From) is not { } from || ToDocPos(sel.To) is not { } to) { NotEditable(); return; }
        if (!SupportedRange(from, to)) return;
        DocPos pos;
        try { pos = Fragments.Paste(_journal, body, "Coller", from, to, fragment, NewParaId()); }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or InvalidOperationException)
        {
            StatusText.Text = $"Collage impossible ici : {ex.Message}";
            return;
        }
        AfterTyping(pos);
        StatusText.Text = $"Collé avec la mise en forme ({source}). Ctrl+Maj+V pour coller le texte seul.";
    }

    /// <summary>
    /// Premier tableau d'un fragment HTML (cellules &lt;td&gt; / &lt;th&gt;, colspan) en tableau du modèle : le contenu de
    /// chaque cellule garde sa mise en forme (lecteur HTML), traits simples d'un demi-point comme Word quand il colle
    /// depuis Excel. Null s'il n'y a pas de tableau.
    /// </summary>
    Fragment? HtmlTable(string html)
    {
        var m = System.Text.RegularExpressions.Regex.Match(html, @"<table\b[^>]*>(.*?)</table>", System.Text.RegularExpressions.RegexOptions.Singleline | System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (!m.Success) return null;
        var rows = System.Text.RegularExpressions.Regex.Matches(m.Groups[1].Value, @"<tr\b[^>]*>(.*?)</tr>", System.Text.RegularExpressions.RegexOptions.Singleline | System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            .Select(r => System.Text.RegularExpressions.Regex.Matches(r.Groups[1].Value, @"<t[dh]\b([^>]*)>(.*?)</t[dh]>", System.Text.RegularExpressions.RegexOptions.Singleline | System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                .Select(c => (Span: System.Text.RegularExpressions.Regex.Match(c.Groups[1].Value, @"colspan\s*=\s*""?(\d+)") is { Success: true } cs ? Math.Clamp(int.Parse(cs.Groups[1].Value), 1, 63) : 1,
                    Html: c.Groups[2].Value, Header: c.Value.StartsWith("<th", StringComparison.OrdinalIgnoreCase))).ToList())
            .Where(r => r.Count > 0).ToList();
        if (rows.Count == 0) return null;
        int cols = Math.Min(63, rows.Max(r => r.Sum(c => c.Span)));
        var w = Body?.Paragraphs().FirstOrDefault()?.Props.Ns ?? (System.Xml.Linq.XNamespace)PropertyBag.TransitionalNs;
        double width = _caret is { } cc && LayoutOf(cc.Paragraph) is { } pl ? pl.Width : 450;
        var table = TableOps.Create(w, rows.Count, cols, width, null, null, NewParaId());
        var line = new TableOps.Line("single", 4, "auto");
        var borders = new System.Xml.Linq.XElement(w + "tblBorders",
            new[] { "top", "left", "bottom", "right", "insideH", "insideV" }.Select(side => line.Element(w, side)));
        table.Props.Set(borders);
        var map = new TableOps.Map(table);
        for (int r = 0; r < rows.Count; r++)
        {
            int col = 0;
            foreach (var (span, cellHtml, header) in rows[r])
            {
                if (col >= cols || map.At(r, col) is not { } slot) break;
                var cellParas = HtmlReader.Read(header ? $"<b>{cellHtml}</b>" : cellHtml, out _);
                var built = Fragments.Build(cellParas, true, w);
                slot.Cell.Content.Blocks.Clear();
                slot.Cell.Content.Blocks.AddRange(built.Blocks.Where(b => b is Paragraph));
                if (slot.Cell.Content.Blocks.Count == 0) slot.Cell.Content.Blocks.Add(TableOps.EmptyParagraph(w, null, NewParaId()));
                col += span;
            }
        }
        // cellules fusionnées (colspan) : fusion horizontale dans la grille
        for (int r = rows.Count - 1; r >= 0; r--)
        {
            int col = 0;
            foreach (var (span, _, _) in rows[r])
            {
                if (span > 1 && col + span - 1 < cols) TableOps.Merge(table, r, col, r, col + span - 1, NewParaId());
                col += span;
            }
        }
        return new Fragment(new List<Block> { table }, false);
    }

    /// <summary>Certains programmes déposent le HTML en UTF-8 lu comme du Latin-1 (« Ã© » pour « é ») : on le redécode.</summary>
    static string FixHtmlEncoding(string html)
    {
        if (!html.Contains('Ã') && !html.Contains("â€")) return html;
        try
        {
            var bytes = Encoding.Latin1.GetBytes(html);
            var utf8 = new UTF8Encoding(false, true).GetString(bytes);
            return utf8;
        }
        catch (Exception ex) when (ex is DecoderFallbackException or EncoderFallbackException) { return html; }
    }

    // ================================================================ écriture HTML et RTF

    static string Css(RunFormat f)
    {
        var sb = new StringBuilder();
        if (f.FontAscii is { } font) sb.Append($"font-family:'{font}';");
        sb.Append(CultureInfo.InvariantCulture, $"font-size:{f.Size:0.#}pt;");
        if (f.Bold) sb.Append("font-weight:bold;");
        if (f.Italic) sb.Append("font-style:italic;");
        var deco = (f.Underline is not null and not "none" ? "underline " : "") + (f.Strike || f.DoubleStrike ? "line-through" : "");
        if (deco.Length > 0) sb.Append($"text-decoration:{deco.Trim()};");
        if (f.Color is { } col) sb.Append($"color:#{col};");
        if (f.VerticalAlign is "superscript") sb.Append("vertical-align:super;");
        if (f.VerticalAlign is "subscript") sb.Append("vertical-align:sub;");
        return sb.ToString();
    }

    static string ToHtml(List<(ParagraphFormat Format, List<(string Text, RunFormat Format)> Runs)> paras)
    {
        var sb = new StringBuilder();
        foreach (var (pf, runs) in paras)
        {
            var align = pf.Alignment switch { "center" => "center", "right" or "end" => "right", "both" or "distribute" => "justify", _ => "left" };
            sb.Append(CultureInfo.InvariantCulture, $"<p style=\"margin:0 0 {pf.SpacingAfter / 20.0:0.#}pt 0;text-align:{align}\">");
            if (runs.Count == 0) sb.Append("&nbsp;");
            foreach (var (text, f) in runs)
            {
                var enc = WebUtility.HtmlEncode(text).Replace("\t", "&emsp;").Replace("\v", "<br>");
                sb.Append($"<span style=\"{Css(f)}\">{enc}</span>");
            }
            sb.Append("</p>");
        }
        return sb.ToString();
    }

    /// <summary>Enveloppe « HTML Format » (en-tête avec les positions en octets UTF-8).</summary>
    static string CfHtml(string fragment)
    {
        const string header = "Version:0.9\r\nStartHTML:{0:D10}\r\nEndHTML:{1:D10}\r\nStartFragment:{2:D10}\r\nEndFragment:{3:D10}\r\n";
        const string pre = "<html><head><meta charset=\"utf-8\"></head><body><!--StartFragment-->";
        const string post = "<!--EndFragment--></body></html>";
        int headerLength = string.Format(CultureInfo.InvariantCulture, header, 0, 0, 0, 0).Length;
        int startHtml = headerLength;
        int startFragment = startHtml + Encoding.UTF8.GetByteCount(pre);
        int endFragment = startFragment + Encoding.UTF8.GetByteCount(fragment);
        int endHtml = endFragment + Encoding.UTF8.GetByteCount(post);
        return string.Format(CultureInfo.InvariantCulture, header, startHtml, endHtml, startFragment, endFragment) + pre + fragment + post;
    }

    static string ToRtf(List<(ParagraphFormat Format, List<(string Text, RunFormat Format)> Runs)> paras, bool endsWithMark)
    {
        var fonts = new List<string>();
        var colors = new List<string>();
        int FontIndex(string? f) { f ??= "Calibri"; int i = fonts.IndexOf(f); if (i < 0) { fonts.Add(f); i = fonts.Count - 1; } return i; }
        int ColorIndex(string? c) { if (c is null) return 0; int i = colors.IndexOf(c); if (i < 0) { colors.Add(c); i = colors.Count - 1; } return i + 1; }
        var body = new StringBuilder();
        for (int p = 0; p < paras.Count; p++)
        {
            var (pf, runs) = paras[p];
            body.Append("\\pard\\plain");
            body.Append(pf.Alignment switch { "center" => "\\qc", "right" or "end" => "\\qr", "both" or "distribute" => "\\qj", _ => "\\ql" });
            body.Append(CultureInfo.InvariantCulture, $"\\sa{pf.SpacingAfter}\\sb{pf.SpacingBefore} ");
            foreach (var (text, f) in runs)
            {
                body.Append(CultureInfo.InvariantCulture, $"{{\\f{FontIndex(f.FontAscii)}\\fs{Math.Round(f.Size * 2)}");
                if (f.Bold) body.Append("\\b");
                if (f.Italic) body.Append("\\i");
                if (f.Underline is not null and not "none") body.Append("\\ul");
                if (f.Strike || f.DoubleStrike) body.Append("\\strike");
                if (f.VerticalAlign == "superscript") body.Append("\\super");
                if (f.VerticalAlign == "subscript") body.Append("\\sub");
                if (f.Color is { } col) body.Append(CultureInfo.InvariantCulture, $"\\cf{ColorIndex(col)}");
                body.Append(' ');
                foreach (var ch in text)
                    body.Append(ch switch
                    {
                        '\\' => "\\\\", '{' => "\\{", '}' => "\\}", '\t' => "\\tab ", '\v' => "\\line ", '\u00A0' => "\\~",
                        _ when ch < 0x80 => ch.ToString(),
                        _ => $"\\u{(short)ch}?",
                    });
                body.Append('}');
            }
            if (p < paras.Count - 1 || endsWithMark) body.Append("\\par\r\n");
        }
        var sb = new StringBuilder("{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl");
        for (int i = 0; i < fonts.Count; i++) sb.Append(CultureInfo.InvariantCulture, $"{{\\f{i} {fonts[i]};}}");
        sb.Append("}{\\colortbl;");
        foreach (var c in colors.Select(c => c.Length == 6 ? c : "000000"))
            sb.Append(CultureInfo.InvariantCulture, $"\\red{Convert.ToInt32(c[..2], 16)}\\green{Convert.ToInt32(c[2..4], 16)}\\blue{Convert.ToInt32(c[4..6], 16)};");
        sb.Append("}\r\n").Append(body).Append('}');
        return sb.ToString();
    }
}
