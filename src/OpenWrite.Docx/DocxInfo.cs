using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenWrite.Docx;

/// <summary>
/// Résumé d'un document .docx, lu sans le mettre en page (version 0.1 : Fichier ▸ Info, mode --info,
/// premier contact du corpus avec l'Open XML SDK). Lecture seule et tolérante : un élément illisible
/// est compté comme absent, jamais une raison de refuser le fichier.
/// </summary>
public sealed class DocxInfo
{
    public string? Title { get; init; }
    public string? Author { get; init; }
    public string? LastModifiedBy { get; init; }
    public DateTime? Modified { get; init; }
    /// <summary>Pages et mots tels que Word les a enregistrés (docProps/app.xml) ; null si absents.</summary>
    public int? SavedPages { get; init; }
    public int? SavedWords { get; init; }
    public string? Application { get; init; }

    /// <summary>
    /// Mode de compatibilité de Word : 11 (2003), 12 (2007), 14 (2010), 15 (2013 et suivants). Absent du fichier, il vaut 12,
    /// comme pour Word (relevé sur le corpus : empty.docx et utf8-bom.docx sont ouverts en mode 12).
    /// </summary>
    public int CompatibilityMode { get; init; } = 12;
    /// <summary>Anciennes options <c>w:compat</c> présentes (hors compatSetting), par nom d'élément.</summary>
    public IReadOnlyList<string> LegacyCompatOptions { get; init; } = Array.Empty<string>();

    public int Paragraphs { get; init; }
    public int Words { get; init; }
    public int Tables { get; init; }
    public int Sections { get; init; }
    public int HeadersFooters { get; init; }
    public int Pictures { get; init; }
    public int TextBoxes { get; init; }
    public int Footnotes { get; init; }
    public int Endnotes { get; init; }
    public int Comments { get; init; }
    public int Revisions { get; init; }
    public int Bookmarks { get; init; }
    public int Hyperlinks { get; init; }
    public int ListParagraphs { get; init; }
    public int ContentControls { get; init; }
    public int LegacyFormFields { get; init; }
    public int Equations { get; init; }
    public bool HasMacros { get; init; }
    public IReadOnlyList<string> Languages { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> Fonts { get; init; } = Array.Empty<string>();
    /// <summary>Parties du package (adresse, type de contenu) : tout ce qui devra être conservé à l'enregistrement.</summary>
    public IReadOnlyList<(string Uri, string ContentType)> Parts { get; init; } = Array.Empty<(string, string)>();
    /// <summary>Texte brut des premiers paragraphes du corps (aperçu de la version 0.1).</summary>
    public IReadOnlyList<string> Preview { get; init; } = Array.Empty<string>();

    /// <summary>Titres du corps (niveau de plan 1 à 9, texte), pour le volet de navigation.</summary>
    public IReadOnlyList<(int Level, string Text)> Headings { get; init; } = Array.Empty<(int, string)>();

    public const int PreviewParagraphs = 60;

    /// <summary>Lit le résumé d'un fichier. Lève <see cref="DocxReadException"/> si ce n'est pas un document Word lisible.</summary>
    public static DocxInfo Read(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return Read(fs);
    }

    public static DocxInfo Read(Stream stream)
    {
        WordprocessingDocument doc;
        try
        {
            doc = WordprocessingDocument.Open(stream, false, new OpenSettings { AutoSave = false });
        }
        catch (Exception ex) when (ex is OpenXmlPackageException or InvalidDataException or IOException or FormatException or System.Xml.XmlException or NotSupportedException or ArgumentException)
        {
            throw new DocxReadException("Ce fichier n'est pas un document Word (.docx) lisible : " + ex.Message, ex);
        }
        using (doc)
        {
            var main = doc.MainDocumentPart ?? throw new DocxReadException("Le document ne contient pas de partie principale (word/document.xml).", null);
            var body = Safe(() => main.Document?.Body);
            var bodyEls = body?.Descendants().ToList() ?? new List<OpenXmlElement>();

            // en-têtes, pieds, notes et commentaires : comptés dans les mots, langues et polices comme le fait Word
            var stories = new List<OpenXmlElement>();
            if (body is not null) stories.Add(body);
            foreach (var h in main.HeaderParts) if (Safe(() => h.Header) is { } x) stories.Add(x);
            foreach (var f in main.FooterParts) if (Safe(() => f.Footer) is { } x) stories.Add(x);
            var footnotes = Safe(() => main.FootnotesPart?.Footnotes?.Elements<Footnote>().Where(IsRealNote).ToList()) ?? new();
            var endnotes = Safe(() => main.EndnotesPart?.Endnotes?.Elements<Endnote>().Where(IsRealNote).ToList()) ?? new();
            var comments = Safe(() => main.WordprocessingCommentsPart?.Comments?.Elements<Comment>().ToList()) ?? new();
            stories.AddRange(footnotes);
            stories.AddRange(endnotes);
            var all = stories.SelectMany(s => s.Descendants()).ToList();

            var settings = Safe(() => main.DocumentSettingsPart?.Settings);
            var compat = settings?.GetFirstChild<Compatibility>();
            int mode = 12;
            var modeSetting = compat?.Elements<CompatibilitySetting>().FirstOrDefault(c => c.Name?.Value == CompatSettingNameValues.CompatibilityMode);
            if (int.TryParse(modeSetting?.Val?.Value, out var m)) mode = m;
            var legacy = compat?.ChildElements.Where(e => e is not CompatibilitySetting).Select(e => e.LocalName).Distinct().ToList() ?? new();

            var core = doc.PackageProperties;
            var app = Safe(() => doc.ExtendedFilePropertiesPart?.Properties);

            var paragraphs = bodyEls.OfType<Paragraph>().ToList();
            var outline = OutlineLevels(Safe(() => main.StyleDefinitionsPart?.Styles));
            var headings = new List<(int, string)>();
            foreach (var p in paragraphs)
            {
                if (p.Ancestors<Table>().Any() || p.Ancestors<TextBoxContent>().Any()) continue;
                var pPr = p.ParagraphProperties;
                int? lvl = pPr?.OutlineLevel?.Val?.Value is int direct ? direct
                         : pPr?.ParagraphStyleId?.Val?.Value is { } sid && outline.TryGetValue(sid, out var sl) ? sl : null;
                if (lvl is not (>= 0 and <= 8)) continue;          // 9 = corps de texte
                var text = TextOf(p).Replace('\t', ' ').Replace('\n', ' ').Trim();
                if (text.Length > 0) headings.Add((lvl.Value + 1, text));
            }
            var preview = new List<string>();
            foreach (var p in body?.Elements<Paragraph>() ?? Enumerable.Empty<Paragraph>())
            {
                if (preview.Count >= PreviewParagraphs) break;
                preview.Add(TextOf(p));
            }

            return new DocxInfo
            {
                Title = NullIfEmpty(core.Title),
                Author = NullIfEmpty(core.Creator),
                LastModifiedBy = NullIfEmpty(core.LastModifiedBy),
                Modified = core.Modified,
                SavedPages = int.TryParse(app?.Pages?.Text, out var pg) ? pg : null,
                SavedWords = int.TryParse(app?.Words?.Text, out var wd) ? wd : null,
                Application = NullIfEmpty(app?.Application?.Text),
                CompatibilityMode = mode,
                LegacyCompatOptions = legacy,
                Paragraphs = paragraphs.Count,
                Words = stories.Sum(s => CountWords(s)),
                Tables = bodyEls.OfType<Table>().Count(),
                Sections = bodyEls.OfType<SectionProperties>().Count(),
                HeadersFooters = main.HeaderParts.Count() + main.FooterParts.Count(),
                Pictures = all.Count(e => e.LocalName == "pic" && e.NamespaceUri.EndsWith("/drawingml/2006/picture", StringComparison.Ordinal))
                           + all.Count(e => e.LocalName == "imagedata"),
                // une zone de texte DrawingML est souvent doublée de sa version VML (mc:Fallback) : on ne compte que la première
                TextBoxes = all.OfType<TextBoxContent>().Count(t => !t.Ancestors().Any(a => a.LocalName == "Fallback")),
                Footnotes = footnotes.Count,
                Endnotes = endnotes.Count,
                Comments = comments.Count,
                Revisions = all.Count(e => e is InsertedRun or DeletedRun or MoveFromRun or MoveToRun or ParagraphPropertiesChange or RunPropertiesChange),
                Bookmarks = all.OfType<BookmarkStart>().Count(b => b.Name?.Value is { } n && n != "_GoBack"),
                Hyperlinks = all.OfType<Hyperlink>().Count(),
                ListParagraphs = paragraphs.Count(p => p.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value is > 0),
                ContentControls = all.Count(e => e is SdtBlock or SdtRun or SdtCell or SdtRow),
                LegacyFormFields = all.OfType<FormFieldData>().Count(),
                Equations = all.Count(e => e.LocalName == "oMath" && !e.Ancestors().Any(a => a.LocalName == "oMath")),
                HasMacros = doc.MainDocumentPart?.VbaProjectPart is not null,
                Languages = all.OfType<Languages>().Select(l => l.Val?.Value).Concat(
                                Safe(() => main.StyleDefinitionsPart?.Styles?.Descendants<Languages>().Select(l => l.Val?.Value).ToList()) ?? new())
                            .Where(v => !string.IsNullOrEmpty(v)).Select(v => v!).Distinct(StringComparer.OrdinalIgnoreCase).Order().ToList(),
                Fonts = (Safe(() => main.FontTablePart?.Fonts?.Elements<Font>().Select(f => f.Name?.Value).ToList()) ?? new())
                            .Where(v => !string.IsNullOrEmpty(v)).Select(v => v!).Distinct().Order().ToList(),
                Parts = doc.Parts.SelectMany(p => Flatten(p.OpenXmlPart)).DistinctBy(p => p.Uri).Select(p => (p.Uri.ToString(), p.ContentType)).OrderBy(p => p.Item1, StringComparer.Ordinal).ToList(),
                Preview = preview,
                Headings = headings,
            };
        }
    }

    /// <summary>Niveau de plan de chaque style de paragraphe (w:outlineLvl), hérité par la chaîne basedOn.</summary>
    static Dictionary<string, int> OutlineLevels(Styles? styles)
    {
        var res = new Dictionary<string, int>(StringComparer.Ordinal);
        if (styles is null) return res;
        var byId = styles.Elements<Style>().Where(st => st.StyleId?.Value is not null && st.Type?.Value == StyleValues.Paragraph)
                         .GroupBy(st => st.StyleId!.Value!).ToDictionary(g => g.Key, g => g.First());
        foreach (var id in byId.Keys)
        {
            var cur = byId[id];
            for (int depth = 0; cur is not null && depth < 20; depth++)      // garde-fou contre une boucle basedOn
            {
                if (cur.StyleParagraphProperties?.OutlineLevel?.Val?.Value is int lvl) { res[id] = lvl; break; }
                cur = cur.BasedOn?.Val?.Value is { } b && byId.TryGetValue(b, out var parent) ? parent : null;
            }
        }
        return res;
    }

    /// <summary>Séparateurs de notes (types separator, continuationSeparator…) : pas de vraies notes.</summary>
    static bool IsRealNote(FootnoteEndnoteType n) => n.Type is null || n.Type.Value == FootnoteEndnoteValues.Normal;

    static IEnumerable<OpenXmlPart> Flatten(OpenXmlPart part, HashSet<OpenXmlPart>? seen = null)
    {
        seen ??= new();
        if (!seen.Add(part)) yield break;
        yield return part;
        foreach (var child in part.Parts)
            foreach (var p in Flatten(child.OpenXmlPart, seen)) yield return p;
    }

    /// <summary>Texte d'un paragraphe : segments, tabulations et retours à la ligne ; le texte supprimé (révisions) est ignoré.</summary>
    public static string TextOf(OpenXmlElement paragraph)
    {
        var sb = new System.Text.StringBuilder();
        foreach (var e in paragraph.Descendants())
        {
            switch (e)
            {
                case Text t when !t.Ancestors<TextBoxContent>().Any(): sb.Append(t.Text); break;
                case TabChar when e.Parent is Run: sb.Append('\t'); break;
                case Break when e.Parent is Run: sb.Append('\n'); break;
            }
        }
        return sb.ToString();
    }

    /// <summary>Mots d'une partie : suites de caractères séparées par des espaces, tabulations ou retours (le texte supprimé ne compte pas).</summary>
    static int CountWords(OpenXmlElement story)
    {
        int n = 0;
        foreach (var p in story.Descendants<Paragraph>())
        {
            bool inWord = false;
            foreach (var e in p.Descendants())
            {
                if (e is TabChar or Break or CarriageReturn) { inWord = false; continue; }
                if (e is not Text t || t.Ancestors<Paragraph>().First() != p) continue;   // texte d'un paragraphe imbriqué : compté avec lui
                foreach (var c in t.Text)
                {
                    bool sep = char.IsWhiteSpace(c);
                    if (!sep && !inWord) n++;
                    inWord = !sep;
                }
            }
        }
        return n;
    }

    static string? NullIfEmpty(string? s) => string.IsNullOrWhiteSpace(s) ? null : s;

    /// <summary>Lecture d'une partie qui peut être endommagée : null plutôt qu'une exception.</summary>
    static T? Safe<T>(Func<T?> read) where T : class
    {
        try { return read(); }
        catch (Exception ex) when (ex is System.Xml.XmlException or InvalidOperationException or OpenXmlPackageException or InvalidDataException) { return null; }
    }

    /// <summary>Libellé du mode de compatibilité, comme dans la barre de titre de Word.</summary>
    public static string CompatibilityLabel(int mode) => mode switch
    {
        <= 11 => "Word 97-2003 (mode de compatibilité)",
        12 => "Word 2007 (mode de compatibilité)",
        14 => "Word 2010 (mode de compatibilité)",
        _ => "Word 2013 et suivants",
    };
}

public sealed class DocxReadException(string message, Exception? inner) : Exception(message, inner);
