using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;
using OpenWrite.Docx;

// Banc d'aller-retour .docx (prototype du lot 1, PLAN.md §8.2).
//   OpenWrite.Fidelity <dossier .docx> <dossier de sortie>          aller-retour + modifications sur tout le dossier
//   OpenWrite.Fidelity --stress <fichier.docx> <n>                  temps de lecture/écriture du corps répété n fois
//   OpenWrite.Fidelity --formats <dossier .docx> <dossier JSON>      mise en forme résolue de chaque paragraphe (StyleResolver)
//   OpenWrite.Fidelity --compare-formats <JSON OpenWrite> <JSON Word> comparaison avec tools/word-format.ps1
//   OpenWrite.Fidelity --compare-word <dossier de sortie>             compare la vue d'OpenWrite (modele/*.json) à celle de Word
//                                                                    (word-modele/*.json, tools/word-inspect.ps1)
// Sorties : <sortie>/aller-retour, modif-texte, modif-gras, modif-scission (les .docx produits) et rapport.md.

Console.OutputEncoding = Encoding.UTF8;
CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
if (args.Length >= 3 && args[0] == "--stress") return Bench.Stress(args[1], int.Parse(args[2]));
if (args.Length >= 2 && args[0] == "--compare-word") return WordCompare.Run(args[1]);
if (args.Length >= 3 && args[0] == "--formats") return FormatCompare.Export(args[1], args[2]);
if (args.Length >= 3 && args[0] == "--compare-formats") return FormatCompare.Compare(args[1], args[2]);
if (args.Length < 2) { Console.Error.WriteLine("Usage : OpenWrite.Fidelity <dossier .docx> <dossier de sortie> | --stress <fichier.docx> <n>"); return 2; }
return Bench.Run(args[0], args[1]);

static class Bench
{
    static readonly string[] Variants = { "aller-retour", "modif-texte", "modif-gras", "modif-scission", "modif-case", "modif-signet", "refaire" };

    sealed class Result
    {
        public string File = "";
        public string? Error;
        public double OpenMs, SaveMs;
        public int StoryParts, OtherParts, OtherIdentical, Dropped;
        public List<string> Diffs = new();
        public int TextChars, ModelChars;
        public Dictionary<string, int> Opaque = new();
        public Dictionary<string, string> Edits = new();     // variante → résultat
        public string Validator = "";
        public Dictionary<string, int> FieldTypes = new();
        public int Bookmarks, HiddenBookmarks;
        public List<string> MarkerProblems = new();
    }

    public static int Run(string dir, string outDir)
    {
        foreach (var v in Variants) Directory.CreateDirectory(Path.Combine(outDir, v));
        var results = new List<Result>();
        foreach (var file in Directory.GetFiles(dir, "*.docx").Order(StringComparer.OrdinalIgnoreCase))
        {
            var r = new Result { File = Path.GetFileName(file) };
            Console.WriteLine("… " + r.File);
            try { One(file, outDir, r); }
            catch (Exception ex) { r.Error = ex.GetType().Name + " : " + ex.Message; }
            results.Add(r);
            Console.WriteLine(r.Error is not null ? "   ERREUR " + r.Error
                : $"   {(r.Diffs.Count == 0 ? "identique" : r.Diffs.Count + " différence(s)")}, {r.OtherIdentical}/{r.OtherParts} parties recopiées à l'identique, " +
                  $"texte modélisé {Pct(r.ModelChars, r.TextChars)}, modifications : {string.Join(", ", r.Edits.Select(e => e.Key + " " + e.Value))}, validateur : {r.Validator}");
        }
        var report = Report(results);
        File.WriteAllText(Path.Combine(outDir, "rapport.md"), report);
        Console.WriteLine();
        Console.WriteLine(report);
        return results.All(r => r.Error is null && r.Diffs.Count == 0 && r.OtherIdentical == r.OtherParts) ? 0 : 1;
    }

    static string Pct(int a, int b) => b == 0 ? "—" : $"{100.0 * a / b:0.#} %";

    static void One(string file, string outDir, Result r)
    {
        var sw = Stopwatch.StartNew();
        var pkg = DocxPackage.Open(file);
        r.OpenMs = sw.Elapsed.TotalMilliseconds;
        r.StoryParts = pkg.Stories.Count;
        r.Dropped = pkg.Stories.Values.Sum(s => s.DroppedNodes);
        if (pkg.Unreadable.Count > 0) r.Diffs.Add("parties illisibles (recopiées) : " + string.Join(", ", pkg.Unreadable));

        // ---- 1. aller-retour sans modification, toutes les parties de texte régénérées depuis le modèle
        var rt = Path.Combine(outDir, "aller-retour", r.File);
        sw.Restart();
        pkg.Save(rt, regenerateAll: true);
        r.SaveMs = sw.Elapsed.TotalMilliseconds;
        var written = ReadZip(rt);
        if (!written.Select(e => e.Name).SequenceEqual(pkg.Entries.Select(e => e.Name))) r.Diffs.Add("liste ou ordre des entrées de l'archive modifiés");
        foreach (var e in pkg.Entries)
        {
            var after = written.FirstOrDefault(w => w.Name == e.Name).Data;
            if (after is null) continue;
            if (pkg.Stories.ContainsKey(e.Name))
            {
                foreach (var d in XmlCompare.Diff(DocxPackage.ParseXml(e.Data).Root!, DocxPackage.ParseXml(after).Root!, 10))
                    r.Diffs.Add(e.Name + " " + d);
            }
            else
            {
                r.OtherParts++;
                if (after.AsSpan().SequenceEqual(e.Data)) r.OtherIdentical++;
            }
        }

        // ---- 2. couverture : texte atteint par le modèle, îlots opaques par nom d'élément
        foreach (var e in pkg.Entries.Where(e => pkg.Stories.ContainsKey(e.Name)))
            r.TextChars += DocxPackage.ParseXml(e.Data).Descendants().Where(x => x.Name.LocalName is "t" or "delText" && x.Name.NamespaceName.EndsWith("main")).Sum(x => x.Value.Length);
        foreach (var story in pkg.Stories.Values.SelectMany(s => s.Stories())) Cover(story, r);

        // ---- 3. modifications locales (chacune sur un document rouvert)
        Edit(file, outDir, "modif-texte", r, (p, run, t, _) => { FlowEdits.InsertText(t, 0, "OW "); return (1, 1); });
        Edit(file, outDir, "modif-gras", r, (p, run, t, _) => { run.Props.SetFlag("b", !run.Props.Flag("b")); return (1, 1); });
        Edit(file, outDir, "modif-scission", r, (p, run, t, story) =>
        {
            var ids = story.Paragraphs().SelectMany(x => x.Attributes).Where(a => a.Name.LocalName is "paraId" or "textId").Select(a => a.Value);
            return FlowEdits.SplitParagraph(story, p, run, t, t.Value.Length / 2, FlowEdits.IdGenerator(ids)) is null ? (-1, -1) : (1, 2);
        });

        Edit2(file, outDir, "modif-case", r, pkg =>
        {
            // première case à cocher de formulaire du corps : inversée
            var box = FieldIndex.Build(pkg.Main!.Stories().First()).All.Select(f => f.CheckBox).FirstOrDefault(c => c is not null);
            if (box is null) return null;
            box.Checked = !box.Checked;
            return (1, 1);
        });
        Edit2(file, outDir, "modif-signet", r, pkg =>
        {
            var body = pkg.Main!.Stories().First();
            var p = body.Blocks.OfType<Paragraph>().FirstOrDefault(x => x.Text.Trim().Length > 0);
            if (p is null) return null;
            MarkerOps.AddAroundParagraph(p, "OpenWrite_essai", MarkerOps.NextId(pkg.Stories.Values.SelectMany(sp => sp.Stories())));
            return (1, 1);
        });
        ExportModel(pkg, Path.Combine(outDir, "modele", Path.ChangeExtension(r.File, ".json")), r);
        UndoRedo(file, outDir, r);

        // ---- 4. validation Open XML SDK : le résultat ne doit pas avoir plus d'erreurs que l'original
        int? before = Validate(file);
        var counts = Variants.Select(v => Validate(Path.Combine(outDir, v, r.File))).ToList();
        r.Validator = before is null ? "original illisible par le SDK" + (counts.Any(c => c is not null) ? $" ; résultats : {string.Join("/", counts.Select(c => c?.ToString() ?? "×"))}" : "")
            : counts.All(c => c is not null && c <= before) ? $"ok ({before} erreur(s) d'origine)"
            : $"ÉCHEC : {before} → {string.Join("/", counts.Select(c => c?.ToString() ?? "illisible"))}";
    }

    /// <summary>
    /// Applique une modification au premier paragraphe du corps qui a un segment de texte, enregistre (seule la partie
    /// modifiée est régénérée) et vérifie l'étendue de la différence dans le corps.
    /// </summary>
    static void Edit(string file, string outDir, string variant, Result r, Func<Paragraph, Run, TextContent, Story, (int, int)> act)
    {
        var pkg = DocxPackage.Open(file);
        var main = pkg.Main!;
        var body = main.Stories().First();
        var target = body.Blocks.OfType<Paragraph>()
            .Select(p => (p, run: p.Inlines.OfType<Run>().FirstOrDefault(x => x.Content.OfType<TextContent>().Any(t => !t.Deleted && t.Value.Trim().Length >= 2))))
            .FirstOrDefault(x => x.run is not null);
        var path = Path.Combine(outDir, variant, r.File);
        if (target.run is null) { pkg.Save(path); r.Edits[variant] = "sans objet"; return; }
        var text = target.run.Content.OfType<TextContent>().First(t => !t.Deleted && t.Value.Trim().Length >= 2);
        // partie d'origine lue avant l'enregistrement (qui met à jour le paquet en mémoire)
        var original = pkg.Find(main.PartName)!.Data;
        var expected = act(target.p, target.run, text, body);
        main.Dirty = true;
        pkg.Save(path);
        if (expected.Item1 < 0) { r.Edits[variant] = "non applicable"; return; }
        var after = ReadZip(path).First(e => e.Name == main.PartName).Data;
        var span = XmlCompare.ChangedSpan(BodyBlocks(original), BodyBlocks(after));
        r.Edits[variant] = span == expected ? "local" : $"ÉTENDUE {span.Removed}→{span.Added}";
    }

    /// <summary>
    /// Série d'opérations par le journal (frappe, scission, suppression sur deux paragraphes, frappe dans une cellule,
    /// case à cocher), puis tout annuler : le XML doit redevenir celui d'origine ; puis tout rétablir et enregistrer
    /// (variante « refaire », contrôlée par le validateur et par Word).
    /// </summary>
    static void UndoRedo(string file, string outDir, Result r)
    {
        var pkg = DocxPackage.Open(file);
        var main = pkg.Main!;
        var body = main.Stories().First();
        var original = DocxPackage.ParseXml(pkg.Find(main.PartName)!.Data).Root!;
        var j = new UndoJournal();
        var ids = FlowEdits.IdGenerator(body.Paragraphs().SelectMany(p => p.Attributes).Where(a => a.Name.LocalName is "paraId" or "textId").Select(a => a.Value));
        var texts = body.Blocks.Select((b, i) => (b, i)).Where(x => x.b is Paragraph p && p.Text.Trim().Length >= 4).Select(x => x.i).ToList();
        var path = Path.Combine(outDir, "refaire", r.File);
        if (texts.Count == 0) { pkg.Save(path); r.Edits["annuler"] = r.Edits["refaire"] = "sans objet"; return; }
        int ops = 0;
        var pos = TextOps.Type(j, body, new DocPos(new[] { texts[0] }, 2), "[OW]"); ops++;
        pos = TextOps.SplitParagraph(j, body, pos, ids); ops++;
        TextOps.Type(j, body, pos, "suite "); ops++;
        // de la 3e lettre du paragraphe scindé à la 2e lettre du paragraphe de texte suivant (blocs intermédiaires compris)
        var next = body.Blocks.Select((b, i) => (b, i)).FirstOrDefault(x => x.i > texts[0] + 1 && x.b is Paragraph p && p.Text.Length >= 2);
        if (next.b is not null) { TextOps.DeleteRange(j, body, new DocPos(new[] { texts[0] + 1 }, 3), new DocPos(new[] { next.i }, 2)); ops++; }
        // frappe dans la première cellule de tableau qui a un paragraphe
        var cellPara = body.Paragraphs().FirstOrDefault(p => !body.Blocks.Contains(p));
        if (cellPara is not null)
        {
            var cp = DocPos.Of(body, cellPara, 0);
            if (cp.Path.Count > 1) { TextOps.Type(j, body, cp, "»"); ops++; }
        }
        // case à cocher, dans une transaction
        var boxField = FieldIndex.Build(body).All.FirstOrDefault(f => f.CheckBox is not null && f.BeginParagraph is not null && body.Blocks.Contains(f.BeginParagraph));
        if (boxField is not null)
        {
            int bi = body.Blocks.IndexOf(boxField.BeginParagraph!);
            using var tx = j.Begin("Case", body, Array.Empty<int>(), bi, 1);
            FieldIndex.Build(body).All.First(f => f.CheckBox is not null && f.BeginParagraph == body.Blocks[bi]).CheckBox!.Checked ^= true;
            tx.Commit();
            ops++;
        }
        var edited = DocxPackage.ParseXml(DocxPackage.Serialize(main)).Root!;
        while (j.CanUndo) j.Undo();
        var undone = XmlCompare.Diff(original, DocxPackage.ParseXml(DocxPackage.Serialize(main)).Root!, 3);
        r.Edits["annuler"] = undone.Count == 0 ? $"identique ({ops} op.)" : "ÉCART : " + undone[0];
        while (j.CanRedo) j.Redo();
        var redone = XmlCompare.Diff(edited, DocxPackage.ParseXml(DocxPackage.Serialize(main)).Root!, 3);
        main.Dirty = true;
        pkg.Save(path);
        r.Edits["refaire"] = redone.Count == 0 ? "identique" : "ÉCART : " + redone[0];
    }

    /// <summary>Variante générale : l'action modifie le document et renvoie l'étendue attendue, ou null si sans objet.</summary>
    static void Edit2(string file, string outDir, string variant, Result r, Func<DocxPackage, (int, int)?> act)
    {
        var pkg = DocxPackage.Open(file);
        var original = pkg.Find(pkg.Main!.PartName)!.Data;
        var expected = act(pkg);
        var path = Path.Combine(outDir, variant, r.File);
        if (expected is null) { pkg.Save(path); r.Edits[variant] = "sans objet"; return; }
        pkg.Main.Dirty = true;
        pkg.Save(path);
        var after = ReadZip(path).First(e => e.Name == pkg.Main.PartName).Data;
        var span = XmlCompare.ChangedSpan(BodyBlocks(original), BodyBlocks(after));
        r.Edits[variant] = span == expected ? "local" : $"ÉTENDUE {span.Removed}→{span.Added}";
    }

    /// <summary>Vue d'OpenWrite sur les champs, cases à cocher et signets (même forme que tools/word-inspect.ps1).</summary>
    static void ExportModel(DocxPackage pkg, string path, Result r)
    {
        var fields = new List<object>();
        var boxes = new List<object>();
        var marks = new List<string>();
        foreach (var part in pkg.Stories.Values)
            foreach (var story in part.Stories())
            {
                var fi = FieldIndex.Build(story);
                var bi = BookmarkIndex.Build(story);
                r.MarkerProblems.AddRange(fi.Problems.Concat(bi.Problems).Select(x => part.PartName + " : " + x));
                foreach (var f in fi.All)
                {
                    fields.Add(new { code = WordCompare.Norm(f.Code), result = WordCompare.Norm(f.Result) });
                    r.FieldTypes[f.Type] = r.FieldTypes.GetValueOrDefault(f.Type) + 1;
                    if (f.CheckBox is { } cb) boxes.Add(new { name = cb.Name ?? "", value = cb.Checked });
                }
                // liens hypertextes : présentés par Word comme des champs HYPERLINK (calibration : \h quand w:history n'est pas activé)
                foreach (var e in FlowWalker.Walk(story).Where(e => e.Kind == FlowEventKind.ContainerStart && e.Item is InlineContainer { Name.LocalName: "hyperlink" }))
                {
                    var link = (InlineContainer)e.Item!;
                    string? A(string local) => link.Attributes.FirstOrDefault(a => a.Name.LocalName == local && a.Name.Namespace == part.W)?.Value;
                    var code = "HYPERLINK" + (pkg.HyperlinkTarget(part, link) is { } url ? $" \"{url}\"" : "")
                               + (A("anchor") is { } an ? $" \\l \"{an}\"" : "") + (A("tooltip") is { } tip ? $" \\o \"{tip}\"" : "")
                               + (A("tgtFrame") is { } fr ? $" \\t \"{fr}\"" : "") + (A("history") is "1" or "true" or "on" ? "" : " \\h");
                    fields.Add(new { code = WordCompare.Norm(code), result = WordCompare.Norm(link.Text) });
                }
                marks.AddRange(bi.All.Select(b => b.Name));
                r.Bookmarks += bi.All.Count(b => !b.Hidden);
                r.HiddenBookmarks += bi.All.Count(b => b.Hidden);
            }
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, JsonSerializer.Serialize(new { fields, checkboxes = boxes, bookmarks = marks }));
    }

    static List<XElement> BodyBlocks(byte[] xml) =>
        DocxPackage.ParseXml(xml).Root!.Elements().First(e => e.Name.LocalName == "body").Elements().ToList();

    static void Cover(Story story, Result r)
    {
        void Opaque(XElement x) { var k = XmlCompare.Short(x.Name); r.Opaque[k] = r.Opaque.GetValueOrDefault(k) + 1; }
        void Inlines(IEnumerable<Inline> list)
        {
            foreach (var i in list)
                switch (i)
                {
                    case Run run:
                        foreach (var c in run.Content)
                            if (c is TextContent t) r.ModelChars += t.Value.Length;
                            else if (c is OpaqueRunContent o) Opaque(o.Xml);
                        break;
                    case InlineContainer box: Inlines(box.Children); break;
                    case InlineContentControl s: Inlines(s.Content); break;
                    case OpaqueInline o: Opaque(o.Xml); break;
                }
        }
        foreach (var b in story.Blocks)
            switch (b)
            {
                case Paragraph p: Inlines(p.Inlines); break;
                case Table t:
                    foreach (var item in t.Items)
                        if (item is TableRow row)
                            foreach (var ci in row.Items)
                                if (ci is TableCell cell) Cover(cell.Content, r); else Opaque((XElement)ci);
                        else Opaque((XElement)item);
                    break;
                case BlockContentControl s: Cover(s.Content, r); break;
                case OpaqueBlock o: Opaque(o.Xml); break;
            }
    }

    static List<(string Name, byte[] Data)> ReadZip(string path)
    {
        using var zip = ZipFile.OpenRead(path);
        return zip.Entries.Where(e => !e.FullName.EndsWith('/')).Select(e =>
        {
            using var s = e.Open();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            return (e.FullName, ms.ToArray());
        }).ToList();
    }

    static int? Validate(string path)
    {
        try
        {
            using var doc = WordprocessingDocument.Open(path, false);
            return new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(doc).Count();
        }
        catch (Exception) { return null; }
    }

    static string Report(List<Result> results)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"# Banc d'aller-retour .docx — {DateTime.Now:yyyy-MM-dd HH:mm}");
        sb.AppendLine();
        sb.AppendLine("| Document | Parties texte | Aller-retour | Autres parties identiques | Texte modélisé | Texte | Gras | Scission | Case | Signet | Annuler tout | Rétablir tout | Champs | Signets (masqués) | Validateur SDK | Lecture / écriture |");
        sb.AppendLine("|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|");
        foreach (var r in results)
        {
            if (r.Error is not null) { sb.AppendLine($"| {r.File} | ERREUR : {r.Error} |||||||||"); continue; }
            sb.AppendLine($"| {r.File} | {r.StoryParts} | {(r.Diffs.Count == 0 ? "identique" : $"**{r.Diffs.Count} diff.**")} | {r.OtherIdentical}/{r.OtherParts} | " +
                          $"{Pct(r.ModelChars, r.TextChars)} | {r.Edits.GetValueOrDefault("modif-texte")} | {r.Edits.GetValueOrDefault("modif-gras")} | " +
                          $"{r.Edits.GetValueOrDefault("modif-scission")} | {r.Edits.GetValueOrDefault("modif-case")} | {r.Edits.GetValueOrDefault("modif-signet")} | " +
                          $"{r.Edits.GetValueOrDefault("annuler")} | {r.Edits.GetValueOrDefault("refaire")} | " +
                          $"{(r.FieldTypes.Count == 0 ? "—" : string.Join(", ", r.FieldTypes.Select(kv => $"{kv.Key} ×{kv.Value}")))} | {r.Bookmarks} ({r.HiddenBookmarks}) | " +
                          $"{r.Validator} | {r.OpenMs:0} / {r.SaveMs:0} ms |");
        }
        var diffs = results.Where(r => r.Diffs.Count > 0).ToList();
        if (diffs.Count > 0)
        {
            sb.AppendLine().AppendLine("## Différences").AppendLine();
            foreach (var r in diffs)
            {
                sb.AppendLine($"- **{r.File}**");
                foreach (var d in r.Diffs.Take(10)) sb.AppendLine($"  - `{d}`");
            }
        }
        var problems = results.Where(r => r.MarkerProblems.Count > 0).ToList();
        if (problems.Count > 0)
        {
            sb.AppendLine().AppendLine("## Anomalies de champs et de signets (conservées telles quelles)").AppendLine();
            foreach (var r in problems) sb.AppendLine($"- **{r.File}** : {string.Join(" ; ", r.MarkerProblems.Take(5))}");
        }
        var dropped = results.Where(r => r.Dropped > 0).ToList();
        if (dropped.Count > 0)
            sb.AppendLine().AppendLine("Nœuds XML ignorés (commentaires, instructions) : " + string.Join(", ", dropped.Select(r => $"{r.File} ({r.Dropped})")));
        var opaque = results.SelectMany(r => r.Opaque).GroupBy(kv => kv.Key)
            .Select(g => (Name: g.Key, Count: g.Sum(x => x.Value), Docs: g.Count())).OrderByDescending(x => x.Count).ToList();
        sb.AppendLine().AppendLine("## Îlots opaques (éléments conservés tels quels)").AppendLine();
        sb.AppendLine("| Élément | Occurrences | Documents |").AppendLine("|---|---|---|");
        foreach (var (name, count, docs) in opaque) sb.AppendLine($"| `{name}` | {count} | {docs} |");
        return sb.ToString();
    }

    // ================================================================ performances

    public static int Stress(string file, int n)
    {
        var src = DocxPackage.Open(file);
        var main = src.Find(src.Main!.PartName)!;
        var xml = DocxPackage.ParseXml(main.Data);
        var body = xml.Root!.Elements().First(e => e.Name.LocalName == "body");
        var blocks = body.Elements().Where(e => e.Name.LocalName != "sectPr").ToList();
        var last = body.Elements().LastOrDefault(e => e.Name.LocalName == "sectPr");
        for (int i = 1; i < n; i++)
            foreach (var b in blocks) (last is null ? (Action<XElement>)body.Add : last.AddBeforeSelf)(new XElement(b));
        var big = new MemoryStream();
        using (var zip = new ZipArchive(big, ZipArchiveMode.Create, true))
            foreach (var e in src.Entries)
            {
                using var s = zip.CreateEntry(e.Name).Open();
                if (e.Name == main.Name) xml.Save(s, SaveOptions.DisableFormatting); else s.Write(e.Data);
            }
        Console.WriteLine($"Corps répété {n} fois : {body.Elements().Count()} blocs, document.xml {big.Length / 1024.0 / 1024:0.0} Mo compressé");
        for (int round = 0; round < 3; round++)
        {
            big.Position = 0;
            var sw = Stopwatch.StartNew();
            var pkg = DocxPackage.Open(big);
            var open = sw.Elapsed.TotalMilliseconds;
            int paragraphs = pkg.Main!.Stories().Sum(s => s.Paragraphs().Count());
            sw.Restart();
            pkg.Main.Dirty = true;
            var outMs = new MemoryStream();
            pkg.Save(outMs);
            Console.WriteLine($"  passe {round + 1} : lecture {open:0} ms, écriture {sw.Elapsed.TotalMilliseconds:0} ms ({paragraphs} paragraphes, {outMs.Length / 1024} Ko)");
            // frappe de 1 000 caractères au milieu du document, journal compris, puis tout annuler
            var story = pkg.Main.Stories().First();
            int mid = story.Blocks.FindIndex(story.Blocks.Count / 2, b => b is Paragraph p && p.Text.Length > 0);
            var j = new UndoJournal();
            var pos = new DocPos(new[] { mid }, 0);
            sw.Restart();
            for (int k = 0; k < 1000; k++) pos = TextOps.Type(j, story, pos, k % 6 == 5 ? " " : "x");
            var typing = sw.Elapsed.TotalMilliseconds;
            int entries = j.Count;
            sw.Restart();
            while (j.CanUndo) j.Undo();
            Console.WriteLine($"           frappe : {typing / 1000:0.000} ms par caractère ({entries} entrées, regroupées par mot), tout annuler : {sw.Elapsed.TotalMilliseconds:0} ms");
        }
        return 0;
    }
}

/// <summary>Comparaison de la vue d'OpenWrite et de celle de Word (champs, cases à cocher, signets).</summary>
static class WordCompare
{
    /// <summary>Normalisation commune : blancs réduits, retours de paragraphe en espaces.</summary>
    public static string Norm(string s) => string.Join(' ', s.Replace('\r', ' ').Replace('\n', ' ').Replace('\a', ' ').Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

    sealed record View(List<FieldView> Fields, List<BoxView> Checkboxes, List<string> Bookmarks);
    sealed record FieldView(string Code, string Result);
    sealed record BoxView(string Name, bool Value);

    static readonly HashSet<string> PageFields = new() { "PAGE", "NUMPAGES", "SECTIONPAGES" };
    static string FieldInstructionType(string code) => FieldInstruction.Parse(code).Type;

    static readonly JsonSerializerOptions Json = new() { PropertyNameCaseInsensitive = true };

    static View? Load(string path)
    {
        if (!File.Exists(path)) return null;
        var v = JsonSerializer.Deserialize<View>(File.ReadAllText(path, Encoding.UTF8), Json);
        // conventions du modèle objet de Word, absentes du fichier (calibration, PLAN.md §10.3) :
        // - chaque forme flottante d'un document en mode 2010 apparaît comme un champ « SHAPE \* MERGEFORMAT » ;
        // - le signet masqué _GoBack (dernière position d'édition) est retiré à l'ouverture et recréé à l'enregistrement.
        return v is null ? null : v with
        {
            Fields = (v.Fields ?? new()).Where(f => !Norm(f.Code).StartsWith("SHAPE", StringComparison.Ordinal)).ToList(),
            Checkboxes = v.Checkboxes ?? new(),
            Bookmarks = (v.Bookmarks ?? new()).Where(b => b != "_GoBack").ToList(),
        };
    }

    public static int Run(string outDir)
    {
        int bad = 0, docs = 0;
        foreach (var mine in Directory.GetFiles(Path.Combine(outDir, "modele"), "*.json").Order())
        {
            var name = Path.GetFileName(mine);
            var a = Load(mine)!;
            var w = Load(Path.Combine(outDir, "word-modele", name));
            if (w is null) { Console.WriteLine($"{name} : pas de relevé Word"); continue; }
            docs++;
            var issues = new List<string>();
            static List<string> Sorted(IEnumerable<string> x) => x.Order(StringComparer.Ordinal).ToList();
            var codesA = Sorted(a.Fields.Select(f => Norm(f.Code)));
            var codesW = Sorted(w.Fields.Select(f => Norm(f.Code)));
            if (!codesA.SequenceEqual(codesW)) issues.Add($"codes de champs : OpenWrite [{string.Join(" | ", codesA)}] ≠ Word [{string.Join(" | ", codesW)}]");
            // PAGE, NUMPAGES, SECTIONPAGES : Word renvoie une valeur recalculée (0 dans un en-tête lu hors contexte de page),
            // pas le résultat enregistré dans le fichier — seul leur code est comparé
            static string Res(FieldView f) => Norm(f.Code) + " → " + (PageFields.Contains(FieldInstructionType(f.Code)) ? "(dynamique)" : Norm(f.Result));
            var resA = Sorted(a.Fields.Select(Res));
            var resW = Sorted(w.Fields.Select(Res));
            if (issues.Count == 0 && !resA.SequenceEqual(resW)) issues.Add($"résultats : OpenWrite [{string.Join(" | ", resA)}] ≠ Word [{string.Join(" | ", resW)}]");
            var boxA = a.Checkboxes.Select(b => $"{b.Name}={b.Value}").ToList();
            var boxW = w.Checkboxes.Select(b => $"{b.Name}={b.Value}").ToList();
            if (!boxA.SequenceEqual(boxW)) issues.Add($"cases : OpenWrite [{string.Join(", ", boxA)}] ≠ Word [{string.Join(", ", boxW)}]");
            var bmA = Sorted(a.Bookmarks);
            var bmW = Sorted(w.Bookmarks);
            if (!bmA.SequenceEqual(bmW, StringComparer.OrdinalIgnoreCase)) issues.Add($"signets : OpenWrite [{string.Join(", ", bmA)}] ≠ Word [{string.Join(", ", bmW)}]");
            Console.WriteLine($"{name} : {(issues.Count == 0 ? $"identique ({a.Fields.Count} champ(s), {a.Checkboxes.Count} case(s), {a.Bookmarks.Count} signet(s))" : "DIFFÉRENT")}");
            foreach (var i in issues) Console.WriteLine("   " + i);
            if (issues.Count > 0) bad++;
        }
        Console.WriteLine($"{docs - bad}/{docs} documents : même vue que Word");
        return bad == 0 ? 0 : 1;
    }
}

/// <summary>
/// Mise en forme résolue par OpenWrite, comparée à celle que Word donne par COM (tools/word-format.ps1) :
/// pour chaque paragraphe du corps, propriétés du paragraphe et du premier caractère (unités de Word : points).
/// </summary>
static class FormatCompare
{
    public sealed record Para(string Text, int Align, double Left, double Right, double First, double Before, double After,
        bool BeforeAuto, bool AfterAuto, double Line, int Rule, string Font, double Size, bool Bold, bool Italic, bool Underline, string Color,
        bool RowEnd = false);

    static string Norm(string s)
    {
        var t = new string(s.Where(c => c >= ' ').ToArray());
        t = string.Join(' ', t.Split(' ', StringSplitOptions.RemoveEmptyEntries));
        return t.Length > 60 ? t[..60] : t;
    }

    /// <summary>Paragraphes du corps avec le style du tableau qui les contient.</summary>
    static IEnumerable<(Paragraph P, string? Table)> Paragraphs(Story story, string? table = null)
    {
        foreach (var b in story.Blocks)
            switch (b)
            {
                case Paragraph p: yield return (p, table); break;
                case Table t:
                    var ts = t.Props.Val("tblStyle") ?? "";          // "" : tableau sans style → style de tableau par défaut
                    foreach (var row in t.Items.OfType<TableRow>())
                        foreach (var cell in row.Items.OfType<TableCell>())
                            foreach (var x in Paragraphs(cell.Content, ts)) yield return x;
                    break;
                case BlockContentControl s:
                    foreach (var x in Paragraphs(s.Content, table)) yield return x;
                    break;
            }
    }

    public static int Export(string dir, string outDir)
    {
        int themeOk = 0, themeAll = 0;
        var themeBad = new List<string>();
        Directory.CreateDirectory(outDir);
        foreach (var file in Directory.GetFiles(dir, "*.docx").Order(StringComparer.OrdinalIgnoreCase))
        {
            var pkg = DocxPackage.Open(file);
            var res = new StyleResolver(pkg.StyleSheet());
            var list = new List<Para>();
            foreach (var (p, table) in Paragraphs(pkg.Main!.Stories().First()))
            {
                var pf = res.Paragraph(p, table);
                var first = ParagraphAtoms.Of(p).FirstOrDefault(a => a.Length > 0 && a.Run is not null);
                var rf = first is null ? res.ParagraphMark(p, table) : res.Run(p, first.Run!, table);
                int rule = pf.LineRule switch
                {
                    "atLeast" => 3,
                    "exact" => 4,
                    _ => pf.Line switch { 240 => 0, 360 => 1, 480 => 2, _ => 5 },
                };
                double line = pf.LineRule is "atLeast" or "exact" ? pf.Line / 20.0 : pf.Line / 240.0 * 12;
                list.Add(new Para(Norm(p.Text), pf.Alignment switch { "center" => 1, "right" => 2, "both" => 3, "distribute" => 4, _ => 0 },
                    pf.IndentLeft / 20.0, pf.IndentRight / 20.0, pf.IndentFirstLine / 20.0, pf.SpacingBefore / 20.0, pf.SpacingAfter / 20.0,
                    pf.BeforeAutospacing, pf.AfterAutospacing, line, rule, rf.FontAscii ?? "", rf.Size, rf.Bold, rf.Italic, rf.Underline is not null, rf.Color ?? "auto"));
            }
            File.WriteAllText(Path.Combine(outDir, Path.ChangeExtension(Path.GetFileName(file), ".json")), JsonSerializer.Serialize(list));
            // couleurs de thème : Word enregistre aussi la couleur calculée dans w:val — la résolution doit la retrouver
            foreach (var part in pkg.Stories.Values)
                foreach (var e in DocxPackage.ParseXml(pkg.Find(part.PartName)!.Data).Descendants(part.W + "color"))
                {
                    if (e.Attribute(part.W + "themeColor") is null || e.Attribute(part.W + "val")?.Value is not { Length: 6 } val) continue;
                    themeAll++;
                    if (string.Equals(res.ResolveColor(e), val, StringComparison.OrdinalIgnoreCase)) themeOk++;
                    else if (themeBad.Count < 8) themeBad.Add($"{Path.GetFileName(file)} : {e} → {res.ResolveColor(e)}");
                }
        }
        Console.WriteLine($"Couleurs de thème : {themeOk}/{themeAll} identiques à la valeur enregistrée par Word (w:val)");
        foreach (var b in themeBad) Console.WriteLine("   " + b);
        return 0;
    }

    static readonly JsonSerializerOptions Json = new() { PropertyNameCaseInsensitive = true };

    static List<(int, int)> Lcs(List<string> a, List<string> b)
    {
        var dp = new int[a.Count + 1, b.Count + 1];
        for (int i = a.Count - 1; i >= 0; i--)
            for (int j = b.Count - 1; j >= 0; j--)
                dp[i, j] = a[i] == b[j] ? dp[i + 1, j + 1] + 1 : Math.Max(dp[i + 1, j], dp[i, j + 1]);
        var res = new List<(int, int)>();
        for (int i = 0, j = 0; i < a.Count && j < b.Count;)
        {
            if (a[i] == b[j]) { res.Add((i, j)); i++; j++; }
            else if (dp[i + 1, j] >= dp[i, j + 1]) i++;
            else j++;
        }
        return res;
    }

    public static int Compare(string mineDir, string wordDir)
    {
        var totals = new Dictionary<string, (int Ok, int All)>();
        var examples = new Dictionary<string, List<string>>();
        int matched = 0, ours = 0;
        foreach (var mine in Directory.GetFiles(mineDir, "*.json").Order())
        {
            var wf = Path.Combine(wordDir, Path.GetFileName(mine));
            if (!File.Exists(wf)) continue;
            var a = JsonSerializer.Deserialize<List<Para>>(File.ReadAllText(mine), Json)!;
            var w = JsonSerializer.Deserialize<List<Para>>(File.ReadAllText(wf, Encoding.UTF8), Json)!.Where(x => !x.RowEnd).ToList();
            ours += a.Count;
            // appariement par plus longue sous-suite commune des textes (Word ajoute des paragraphes de fin de ligne de
            // tableau ; un appariement glouton se tromperait de paragraphe vide)
            var pairs = Lcs(a.Select(x => x.Text).ToList(), w.Select(x => Norm(x.Text)).ToList());
            foreach (var (ia, iw) in pairs)
            {
                var p = a[ia];
                var q = w[iw];
                matched++;
                void Check(string prop, bool ok, string detail)
                {
                    var t = totals.GetValueOrDefault(prop);
                    totals[prop] = (t.Ok + (ok ? 1 : 0), t.All + 1);
                    if (!ok && (examples.TryGetValue(prop, out var ex) ? ex : examples[prop] = new()).Count < 6)
                        examples[prop].Add($"{Path.GetFileNameWithoutExtension(mine)[..Math.Min(18, Path.GetFileNameWithoutExtension(mine).Length)]} « {p.Text[..Math.Min(30, p.Text.Length)]} » : {detail}");
                }
                static bool Near(double x, double y) => Math.Abs(x - y) < 0.06;
                Check("alignement", p.Align == q.Align, $"{p.Align} / Word {q.Align}");
                Check("retrait gauche", Near(p.Left, q.Left), $"{p.Left} / Word {q.Left}");
                Check("retrait droit", Near(p.Right, q.Right), $"{p.Right} / Word {q.Right}");
                Check("première ligne", Near(p.First, q.First), $"{p.First} / Word {q.First}");
                if (!p.BeforeAuto && !q.BeforeAuto) Check("espace avant", Near(p.Before, q.Before), $"{p.Before} / Word {q.Before}");
                else Check("espace avant auto", p.BeforeAuto == q.BeforeAuto, $"{p.BeforeAuto} / Word {q.BeforeAuto}");
                if (!p.AfterAuto && !q.AfterAuto) Check("espace après", Near(p.After, q.After), $"{p.After} / Word {q.After}");
                else Check("espace après auto", p.AfterAuto == q.AfterAuto, $"{p.AfterAuto} / Word {q.AfterAuto}");
                Check("interligne", p.Rule == q.Rule && Near(p.Line, q.Line), $"{p.Rule}:{p.Line} / Word {q.Rule}:{q.Line}");
                Check("police", string.Equals(p.Font, q.Font, StringComparison.OrdinalIgnoreCase), $"{p.Font} / Word {q.Font}");
                Check("taille", Near(p.Size, q.Size), $"{p.Size} / Word {q.Size}");
                Check("gras", p.Bold == q.Bold, $"{p.Bold} / Word {q.Bold}");
                Check("italique", p.Italic == q.Italic, $"{p.Italic} / Word {q.Italic}");
                Check("souligné", p.Underline == q.Underline, $"{p.Underline} / Word {q.Underline}");
                if (q.Color != "theme") Check("couleur", string.Equals(p.Color, q.Color, StringComparison.OrdinalIgnoreCase), $"{p.Color} / Word {q.Color}");
            }
        }
        Console.WriteLine($"Paragraphes appariés avec Word : {matched}/{ours}");
        foreach (var (prop, (ok, all)) in totals)
        {
            Console.WriteLine($"  {prop,-18} {ok,5}/{all,-5} {100.0 * ok / all,6:0.0} %");
            if (examples.TryGetValue(prop, out var ex)) foreach (var e in ex) Console.WriteLine("      " + e);
        }
        return totals.Values.All(t => t.Ok == t.All) ? 0 : 1;
    }
}
