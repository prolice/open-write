using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using OpenWrite.Core.Model;
using OpenWrite.Docx;
using OpenWrite.Layout;

namespace OpenWrite.App.Layout;

/// <summary>
/// Banc de la mise en page (lot 2), mode --layout : pages calculées comparées au nombre de pages de Word
/// (work/reference/references.csv), temps de la mise en page complète et de la reprise incrémentale après une frappe,
/// et contrôle que la reprise donne les mêmes pages qu'une mise en page complète.
///   OpenWrite --layout &lt;fichier.docx | dossier&gt; [--repeat n] [--refs references.csv] [--word-pages dossier] [--detail] [--no-suppress]
/// --word-pages : début de chaque page selon Word (tools/word-pages.ps1) ; compte les pages qui commencent au même
/// endroit (12 premiers caractères alphanumériques), --detail affiche les écarts.
/// </summary>
static class LayoutBench
{
    public static int Run(string[] args)
    {
        var target = args[1];
        int repeat = 1;
        string? refs = null;
        bool suppress = true, detail = false;
        string? wordDir = null;
        for (int i = 2; i < args.Length; i++)
        {
            if (args[i] == "--repeat" && i + 1 < args.Length) repeat = int.Parse(args[++i]);
            else if (args[i] == "--refs" && i + 1 < args.Length) refs = args[++i];
            else if (args[i] == "--no-suppress") suppress = false;
            else if (args[i] == "--word-pages" && i + 1 < args.Length) wordDir = args[++i];
            else if (args[i] == "--detail") detail = true;
        }
        var wordPages = refs is not null && File.Exists(refs) ? ReadRefs(refs) : new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var files = Directory.Exists(target) ? Directory.GetFiles(target, "*.docx").OrderBy(f => f, StringComparer.OrdinalIgnoreCase).ToArray() : new[] { target };
        var fonts = new WpfFontProvider();
        int same = 0, compared = 0, mismatch = 0, startsSame = 0, startsTotal = 0;
        Console.WriteLine($"{"document",-34} {"pages",6} {"Word",5} {"complète",9} {"frappe",8} {"pages/frappe",12} {"§/frappe",9}");
        foreach (var file in files)
        {
            string name = Path.GetFileName(file);
            try
            {
                var pkg = DocxPackage.Open(file);
                var body = pkg.Main?.Stories().FirstOrDefault();
                if (body is null) { Console.WriteLine($"{Short(name),-34} (pas de corps)"); continue; }
                if (repeat > 1)
                {
                    var copy = body.Blocks.ToList();
                    for (int r = 1; r < repeat; r++) body.Blocks.AddRange(copy.Select(b => b.Clone()));
                }
                var styles = new Core.Styles.StyleResolver(pkg.StyleSheet());
                fonts.AltNames = pkg.FontAltNames();
                // mise en page complète : une fois à froid (polices), puis mesurée
                new DocumentLayout(body, styles, fonts, pkg.RelatedStory, id => pkg.NoteStory("footnote", id)) { SuppressSpaceBeforeAtPageTop = suppress }.LayoutAll();
                var layout = new DocumentLayout(body, styles, fonts, pkg.RelatedStory, id => pkg.NoteStory("footnote", id)) { SuppressSpaceBeforeAtPageTop = suppress };
                var sw = Stopwatch.StartNew();
                layout.LayoutAll();
                double full = sw.Elapsed.TotalMilliseconds;
                int pages = layout.Pages.Count;
                var startTexts = layout.Pages.Select(StartText).ToList();     // avant les frappes simulées
                var positions = Positions(body, layout);

                // frappe : un caractère ajouté dans 20 paragraphes répartis dans le document
                var paras = body.Blocks.OfType<Paragraph>().Where(p => p.Runs().Any(r => r.Content.OfType<TextContent>().Any())).ToList();
                double typing = 0, worst = 0;
                int pagesRedone = 0, parasRedone = 0, edits = Math.Min(20, paras.Count);
                for (int k = 0; k < edits; k++)
                {
                    var p = paras[(int)((k + 0.5) * paras.Count / edits)];
                    p.Runs().SelectMany(r => r.Content.OfType<TextContent>()).First().Value += "x";
                    sw.Restart();
                    layout.Update(new Block[] { p });
                    double t = sw.Elapsed.TotalMilliseconds;
                    typing += t;
                    worst = Math.Max(worst, t);
                    pagesRedone += layout.LastPagesLaidOut;
                    parasRedone += layout.LastParagraphsLaidOut;
                }
                // contrôle : la reprise donne les mêmes pages qu'une mise en page complète
                var check = new DocumentLayout(body, styles, fonts, pkg.RelatedStory, id => pkg.NoteStory("footnote", id)) { SuppressSpaceBeforeAtPageTop = suppress };
                check.LayoutAll();
                bool consistent = check.Pages.Count == layout.Pages.Count && check.Pages.Zip(layout.Pages).All(z =>
                    ReferenceEquals(z.First.Start.Item?.Block, z.Second.Start.Item?.Block) && z.First.Start.Line == z.Second.Start.Line);
                if (!consistent) mismatch++;

                string word = "";
                if (repeat == 1 && wordPages.TryGetValue(name, out var wp))
                {
                    word = wp.ToString();
                    compared++;
                    if (wp == pages) same++;
                }
                string starts = "";
                if (repeat == 1 && wordDir is not null && File.Exists(Path.Combine(wordDir, Path.GetFileNameWithoutExtension(name) + ".json")))
                {
                    var wordStarts = System.Text.Json.JsonDocument.Parse(File.ReadAllText(Path.Combine(wordDir, Path.GetFileNameWithoutExtension(name) + ".json")))
                        .RootElement.EnumerateArray().Select(e => e.GetProperty("text").GetString() ?? "").ToList();
                    int ok = 0, n = Math.Max(wordStarts.Count, startTexts.Count);
                    for (int k = 0; k < n; k++)
                    {
                        var mine = k < startTexts.Count ? Key(startTexts[k]) : "∅";
                        var theirs = k < wordStarts.Count ? Key(wordStarts[k]) : "∅";
                        if (mine == theirs) ok++;
                        else if (detail) Console.WriteLine($"    page {k + 1} : Word « {theirs} » / OpenWrite « {mine} »");
                    }
                    startsSame += ok;
                    startsTotal += n;
                    starts = $"  débuts {ok}/{n}";
                    var parasFile = Path.Combine(wordDir, Path.GetFileNameWithoutExtension(name) + ".paragraphes.json");
                    if (File.Exists(parasFile)) starts += "  " + ComparePositions(parasFile, positions, detail);
                }
                Console.WriteLine($"{Short(name),-34} {pages,6} {word,5} {full,7:0.0} ms {(edits > 0 ? typing / edits : 0),5:0.0} ms {(edits > 0 ? pagesRedone / (double)edits : 0),12:0.0} {(edits > 0 ? parasRedone / (double)edits : 0),9:0.0}"
                    + (edits > 0 ? $"  (pire frappe {worst:0.0} ms)" : "") + (consistent ? "" : "  REPRISE DIFFÉRENTE") + starts);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{Short(name),-34} erreur : {ex.GetType().Name} {ex.Message}");
            }
        }
        if (compared > 0) Console.WriteLine($"\nNombre de pages identique à Word : {same}/{compared}");
        if (startsTotal > 0) Console.WriteLine($"Pages qui commencent au même endroit que dans Word : {startsSame}/{startsTotal}");
        if (fonts.Substitutions.Count > 0)
            Console.WriteLine("Polices remplacées : " + string.Join(", ", fonts.Substitutions.Select(kv => $"{kv.Key} → {kv.Value}")));
        return mismatch == 0 ? 0 : 1;
    }

    /// <summary>Page (1…) et position verticale de la première ligne de chaque paragraphe, dans l'ordre du document.</summary>
    static List<(string Key, int Page, double Y, double X)> Positions(Story body, DocumentLayout layout)
    {
        var first = new Dictionary<Paragraph, (int, double, double)>(ReferenceEqualityComparer.Instance);
        for (int k = 0; k < layout.Pages.Count; k++)
            foreach (var l in layout.Pages[k].Lines)
                first.TryAdd(l.Paragraph.Paragraph, (k + 1, l.Y, l.X + (l.Paragraph.Format.IndentLeft + Math.Max(0, l.Paragraph.Format.IndentFirstLine)) / 20.0));
        var list = new List<(string, int, double, double)>();
        foreach (var p in body.Paragraphs())
            if (first.TryGetValue(p, out var pos)) list.Add((Key(p.Text), pos.Item1, pos.Item2, pos.Item3));
        // OPENWRITE_BENCH_DUMP=n : positions d'OpenWrite sur la page n (mise au point)
        if (int.TryParse(Environment.GetEnvironmentVariable("OPENWRITE_BENCH_DUMP"), out var dump))
            foreach (var x in list.Where(x => x.Item2 == dump).TakeLast(Environment.GetEnvironmentVariable("OPENWRITE_BENCH_END") == "1" ? 12 : 1000).Take(80)) Console.WriteLine($"    OpenWrite p.{dump} {x.Item3,6:0.0} « {x.Item1} »");
        return list;
    }

    /// <summary>
    /// Écart vertical avec Word, paragraphe par paragraphe (appariés par leurs 12 premiers caractères, dans l'ordre) :
    /// médiane et pire écart sur les paragraphes de la même page ; --detail : premiers écarts de plus de 2 pt.
    /// </summary>
    static string ComparePositions(string file, List<(string Key, int Page, double Y, double X)> mine, bool detail)
    {
        var word = System.Text.Json.JsonDocument.Parse(File.ReadAllText(file)).RootElement.EnumerateArray()
            .Select(e => (Key: Key(e.GetProperty("text").GetString() ?? ""), Page: e.GetProperty("page").GetInt32(), Y: e.GetProperty("y").GetDouble(),
                X: e.TryGetProperty("x", out var x) ? x.GetDouble() : double.NaN))
            .Where(w => w.Key.Length >= 4).ToList();
        mine = mine.Where(m => m.Key.Length >= 4).ToList();
        var deltas = new List<double>();
        var dx = new List<double>();
        int shown = 0, otherPage = 0;
        foreach (var (wi, mi) in Align(word.Select(w => w.Key).ToList(), mine.Select(m => m.Key).ToList()))
        {
            var w = word[wi];
            var m = mine[mi];
            if (m.Page != w.Page) { otherPage++; continue; }
            double d = m.Y - w.Y;
            deltas.Add(d);
            // écart horizontal : bord gauche du paragraphe, retraits compris (Word donne la marge pour un paragraphe
            // centré : alignement non compté) — colonnes de tableau, retraits
            double ex = double.IsNaN(w.X) ? 0 : m.X - w.X;
            if (!double.IsNaN(w.X)) dx.Add(ex);
            if (detail && (Environment.GetEnvironmentVariable("OPENWRITE_BENCH_ALL") == "1" || (Math.Abs(d) > 2 || Math.Abs(ex) > 2) && shown++ < 6))
                Console.WriteLine($"    p.{w.Page} « {w.Key} » : Word {w.Y:0.0} / OpenWrite {m.Y:0.0} ({d:+0.0;-0.0})" + (double.IsNaN(w.X) ? "" : $", x Word {w.X:0.0} / OpenWrite {m.X:0.0} ({ex:+0.0;-0.0})"));
        }
        if (deltas.Count == 0) return "";
        deltas.Sort();
        return $"Δy médian {deltas[deltas.Count / 2]:+0.0;-0.0} pt, pire {deltas.MaxBy(Math.Abs):+0.0;-0.0}, {deltas.Count(d => Math.Abs(d) <= 1)}/{deltas.Count} à 1 pt, {otherPage} autre page"
            + (dx.Count > 0 ? $", Δx {dx.Count(d => Math.Abs(d) <= 1)}/{dx.Count} à 1 pt (pire {dx.MaxBy(Math.Abs):+0.0;-0.0})" : "");
    }

    /// <summary>Appariement de deux suites de clés par plus longue sous-suite commune (paires d'indices).</summary>
    static List<(int, int)> Align(List<string> a, List<string> b)
    {
        var t = new int[a.Count + 1, b.Count + 1];
        for (int i = a.Count - 1; i >= 0; i--)
            for (int j = b.Count - 1; j >= 0; j--)
                t[i, j] = a[i] == b[j] ? t[i + 1, j + 1] + 1 : Math.Max(t[i + 1, j], t[i, j + 1]);
        var res = new List<(int, int)>();
        for (int i = 0, j = 0; i < a.Count && j < b.Count;)
        {
            if (a[i] == b[j]) { res.Add((i, j)); i++; j++; }
            else if (t[i + 1, j] >= t[i, j + 1]) i++;
            else j++;
        }
        return res;
    }

    /// <summary>Texte du début de la page : première ligne placée et suivantes.</summary>
    static string StartText(Page page)
    {
        // page commencée dans un tableau : les rangées d'en-tête répétées ne comptent pas ; au milieu d'une rangée
        // coupée, Word donne le début du premier paragraphe qui continue (Range.GoTo)
        var skip = new HashSet<Paragraph>(ReferenceEqualityComparer.Instance);
        if (page.Start is { Item.Block: Table t } st)
        {
            var rows = t.Items.OfType<TableRow>().ToList();
            for (int h = 0; h < rows.Count && h < st.Line && rows[h].Props.Get("tblHeader") is not null; h++)
                foreach (var p in rows[h].Items.OfType<TableCell>().SelectMany(c => c.Content.Paragraphs())) skip.Add(p);
            if (st.Offset > 0 && page.Lines.FirstOrDefault(l => !skip.Contains(l.Paragraph.Paragraph)) is { Paragraph: { } first })
                return first.Content.Text;
        }
        var sb = new System.Text.StringBuilder();
        foreach (var l in page.Lines)
        {
            if (skip.Contains(l.Paragraph.Paragraph)) continue;
            var c = l.Paragraph.Content;
            var line = l.Paragraph.Lines[l.Line];
            // les numéros d'appel de note ne sont pas du texte pour Word (caractère spécial dans Range.Text)
            for (int k = line.Start; k < line.End; k++)
            {
                if (c.Footnotes.TryGetValue(k, out var id)) { k += id.Length - 1; continue; }
                sb.Append(c.Text[k]);
            }
            sb.Append(' ');
            if (sb.Length > 80) break;
        }
        return sb.ToString();
    }

    /// <summary>12 premiers caractères alphanumériques (sans espaces, marques de cellule, objets).</summary>
    static string Key(string s) => new string(s.Where(char.IsLetterOrDigit).Take(12).ToArray()).ToUpperInvariant();

    // noms de fichiers du corpus personnel : seulement le début dans la sortie
    static string Short(string name) => name.Length <= 34 ? name : name[..31] + "...";

    static Dictionary<string, int> ReadRefs(string path)
    {
        var d = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var line in File.ReadLines(path).Skip(1))
        {
            var m = Regex.Matches(line, "\"((?:[^\"]|\"\")*)\"");
            if (m.Count >= 2 && int.TryParse(m[1].Groups[1].Value, out var pages)) d[m[0].Groups[1].Value] = pages;
        }
        return d;
    }

    /// <summary>
    /// Mode --lines : lignes d'OpenWrite comparées à celles de Word (tools/word-lines.ps1 : page, haut, gauche, largeur et
    /// début du texte de chaque ligne du corps). Lignes appariées dans l'ordre par leur texte ; écarts de position du
    /// haut et du bord gauche (habillage autour des objets flottants).
    ///   OpenWrite --lines &lt;dossier .docx&gt; &lt;dossier JSON de Word&gt; [--detail]
    /// </summary>
    public static int Lines(string dir, string wordDir, bool detail)
    {
        var fonts = new WpfFontProvider();
        int allSame = 0, allTotal = 0;
        foreach (var file in Directory.GetFiles(dir, "*.docx").OrderBy(f => f, StringComparer.OrdinalIgnoreCase))
        {
            var json = Path.Combine(wordDir, Path.GetFileNameWithoutExtension(file) + ".lignes.json");
            if (!File.Exists(json)) continue;
            var pkg = DocxPackage.Open(file);
            var body = pkg.Main!.Stories().First();
            fonts.AltNames = pkg.FontAltNames();
            var layout = new DocumentLayout(body, new Core.Styles.StyleResolver(pkg.StyleSheet()), fonts, pkg.RelatedStory, id => pkg.NoteStory("footnote", id));
            layout.LayoutAll();
            var mine = new List<(int Page, double Top, double Left, string Text)>();
            for (int k = 0; k < layout.Pages.Count; k++)
                foreach (var l in layout.Pages[k].Lines)
                {
                    var lb = l.Paragraph.Lines[l.Line];
                    var c = l.Paragraph.Content;
                    // bord gauche de la boîte de ligne, comme Word (Line.Left) : la colonne, ou le bord libre près d'un objet
                    double left = l.X + lb.BoxLeft;
                    mine.Add((k + 1, l.Y, left, c.Text[lb.Start..Math.Min(lb.End, c.Text.Length)]));
                }
            using var doc = System.Text.Json.JsonDocument.Parse(File.ReadAllText(json));
            var word = doc.RootElement.EnumerateArray().Select(e => (Page: e.GetProperty("page").GetInt32(), Top: e.GetProperty("top").GetDouble(),
                Left: e.GetProperty("left").GetDouble(), Text: e.GetProperty("text").GetString() ?? "")).ToList();
            static string K(string s) => new string(s.Where(char.IsLetterOrDigit).Take(10).ToArray());
            int same = 0, n = Math.Min(mine.Count, word.Count);
            var lines = new List<string>();
            for (int k = 0; k < n; k++)
            {
                var (wp, wt, wl, wtext) = word[k];
                var (mp, mt, ml, mtext) = mine[k];
                bool textOk = K(wtext) == K(mtext);
                bool ok = textOk && wp == mp && Math.Abs(wt - mt) <= 1.01 && Math.Abs(wl - ml) <= 1.01;
                if (ok) same++;
                if (detail && !ok) lines.Add($"    {k + 1,3} Word p{wp} {wt,6:0.0} {wl,6:0.0} « {K(wtext)} » / OpenWrite p{mp} {mt,6:0.0} {ml,6:0.0} « {K(mtext)} »");
            }
            allSame += same;
            allTotal += Math.Max(mine.Count, word.Count);
            Console.WriteLine($"{Path.GetFileName(file),-24} lignes {mine.Count,3} (Word {word.Count,3})  identiques à 1 pt : {same}/{Math.Max(mine.Count, word.Count)}");
            foreach (var l in lines.Take(12)) Console.WriteLine(l);
        }
        Console.WriteLine($"Lignes identiques à Word (texte, page, haut et bord gauche à 1 pt) : {allSame}/{allTotal}");
        return 0;
    }
}
