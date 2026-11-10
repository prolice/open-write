using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using OpenWrite.Core.Io;

namespace OpenWrite.Docx;

/// <summary>
/// Package .docx (prototype du lot 1) : lu directement comme archive ZIP (tolérant : les fichiers que Word ou le SDK
/// jugent endommagés s'ouvrent quand même), parties de texte chargées dans le modèle en flux, toutes les autres
/// parties gardées octet pour octet. À l'enregistrement, seules les parties modifiées sont régénérées ; l'ordre,
/// les noms et les dates des entrées sont conservés.
/// </summary>
public sealed class DocxPackage
{
    public sealed record Entry(string Name, byte[] Data, DateTimeOffset LastWrite);

    /// <summary>Entrées de l'archive, dans l'ordre d'origine ([Content_Types].xml en premier chez Word).</summary>
    public List<Entry> Entries { get; } = new();
    /// <summary>Parties de texte, par nom d'entrée (document, en-têtes, pieds, notes, commentaires).</summary>
    public Dictionary<string, StoryPart> Stories { get; } = new(StringComparer.OrdinalIgnoreCase);
    public StoryPart? Main { get; private set; }
    /// <summary>Parties de texte qui n'ont pas pu être lues (XML invalide) : recopiées telles quelles.</summary>
    public List<string> Unreadable { get; } = new();

    public static DocxPackage Open(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return Open(fs);
    }

    public static DocxPackage Open(Stream stream)
    {
        var pkg = new DocxPackage();
        try
        {
            using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
            foreach (var e in zip.Entries)
            {
                if (e.FullName.EndsWith('/')) continue;      // dossier
                using var s = e.Open();
                using var ms = new MemoryStream();
                s.CopyTo(ms);
                pkg.Entries.Add(new Entry(e.FullName, ms.ToArray(), e.LastWriteTime));
            }
        }
        catch (InvalidDataException ex) { throw new DocxReadException("Ce fichier n'est pas un document Word (.docx) lisible : " + ex.Message, ex); }

        var mainName = pkg.Targets("", t => t.EndsWith("/officeDocument", StringComparison.Ordinal)).FirstOrDefault().Name
                       ?? throw new DocxReadException("Le document ne contient pas de partie principale (word/document.xml).", null);
        pkg.Main = pkg.Load(mainName, "document");
        foreach (var (name, type) in pkg.Targets(mainName, _ => true))
        {
            var kind = type[(type.LastIndexOf('/') + 1)..];
            if (kind is "header" or "footer" or "footnotes" or "endnotes" or "comments") pkg.Load(name, kind);
        }
        return pkg;
    }

    public Entry? Find(string name) => Entries.FirstOrDefault(e => string.Equals(e.Name, name, StringComparison.OrdinalIgnoreCase));

    static readonly XmlReaderSettings ReaderSettings = new() { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };

    public static XDocument ParseXml(byte[] data)
    {
        using var ms = new MemoryStream(data);
        using var r = XmlReader.Create(ms, ReaderSettings);
        return XDocument.Load(r, LoadOptions.PreserveWhitespace);
    }

    StoryPart? Load(string name, string kind)
    {
        if (Stories.TryGetValue(name, out var done)) return done;
        if (Find(name) is not { } entry) return null;
        try
        {
            var part = StoryMapping.Read(ParseXml(entry.Data), entry.Name, kind);
            Stories[entry.Name] = part;
            return part;
        }
        catch (XmlException)
        {
            Unreadable.Add(entry.Name);
            return null;
        }
    }

    /// <summary>Relation d'une partie : cible (nom d'entrée résolu pour une cible interne, adresse telle quelle sinon), type.</summary>
    public sealed record Relationship(string Id, string Type, string Target, bool External);

    /// <summary>Relations d'une partie (« » = relations du package), par identifiant (r:id).</summary>
    public IReadOnlyDictionary<string, Relationship> Relationships(string partName)
    {
        var res = new Dictionary<string, Relationship>(StringComparer.Ordinal);
        var dir = partName.Contains('/') ? partName[..partName.LastIndexOf('/')] : "";
        var relsName = (dir.Length > 0 ? dir + "/" : "") + "_rels/" + partName[(partName.LastIndexOf('/') + 1)..] + ".rels";
        if (Find(relsName) is not { } rels) return res;
        XDocument x;
        try { x = ParseXml(rels.Data); } catch (XmlException) { return res; }
        foreach (var rel in x.Root?.Elements().Where(e => e.Name.LocalName == "Relationship") ?? Enumerable.Empty<XElement>())
        {
            var id = rel.Attribute("Id")?.Value;
            var target = rel.Attribute("Target")?.Value;
            if (id is null || target is null) continue;
            bool external = rel.Attribute("TargetMode")?.Value == "External";
            res[id] = new Relationship(id, rel.Attribute("Type")?.Value ?? "", external ? target : Resolve(dir, Uri.UnescapeDataString(target)), external);
        }
        return res;
    }

    IEnumerable<(string Name, string Type)> Targets(string partName, Func<string, bool> typeFilter) =>
        Relationships(partName).Values.Where(r => !r.External && typeFilter(r.Type)).Select(r => (r.Target, r.Type));

    const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const string StrictRelNs = "http://purl.oclc.org/ooxml/officeDocument/relationships";

    /// <summary>Histoire d'une partie liée à la partie principale (en-tête, pied de page) par son identifiant de relation (r:id).</summary>
    public OpenWrite.Core.Model.Story? RelatedStory(string relId)
    {
        if (Main is null) return null;
        _mainRels ??= Relationships(Main.PartName);
        return _mainRels.TryGetValue(relId, out var rel) && !rel.External && Stories.TryGetValue(rel.Target, out var part) ? part.Stories().FirstOrDefault() : null;
    }

    IReadOnlyDictionary<string, Relationship>? _mainRels;

    /// <summary>Adresse d'un lien hypertexte (w:hyperlink r:id) d'une partie ; null pour un lien interne (w:anchor seul).</summary>
    public string? HyperlinkTarget(StoryPart part, OpenWrite.Core.Model.InlineContainer link)
    {
        var rid = link.Attributes.FirstOrDefault(a => a.Name.LocalName == "id" && a.Name.NamespaceName is RelNs or StrictRelNs)?.Value;
        return rid is not null && Relationships(part.PartName).TryGetValue(rid, out var r) ? r.Target : null;
    }

    static string Resolve(string dir, string target)
    {
        var parts = new List<string>();
        if (!target.StartsWith('/') && dir.Length > 0) parts.AddRange(dir.Split('/'));
        foreach (var seg in target.TrimStart('/').Split('/'))
        {
            if (seg is "" or ".") continue;
            if (seg == "..") { if (parts.Count > 0) parts.RemoveAt(parts.Count - 1); }
            else parts.Add(seg);
        }
        return string.Join('/', parts);
    }

    /// <summary>
    /// Styles, numérotation, thème et paramètres du document principal, pour le résolveur de styles (lecture seule :
    /// ces parties sont recopiées telles quelles tant que les styles ne sont pas modifiables, lot 4).
    /// </summary>
    public OpenWrite.Core.Styles.StyleSheet StyleSheet()
    {
        XElement? Part(string kind)
        {
            if (Main is null) return null;
            var name = Targets(Main.PartName, t => t.EndsWith("/" + kind, StringComparison.Ordinal)).FirstOrDefault().Name;
            if (name is null || Find(name) is not { } e) return null;
            try { return ParseXml(e.Data).Root; } catch (XmlException) { return null; }
        }
        return new OpenWrite.Core.Styles.StyleSheet(Part("styles"), Part("numbering"), Part("theme"), Part("settings"), Main?.W);
    }

    // ================================================================ écriture

    /// <summary>XML d'une partie de texte régénérée depuis le modèle (UTF-8 sans BOM, sans indentation, comme Word).</summary>
    public static byte[] Serialize(StoryPart part)
    {
        var doc = StoryMapping.Write(part);
        using var ms = new MemoryStream();
        // Entitize : un retour à la ligne dans un attribut (texte de remplacement d'une image, o:gfxdata) est écrit &#xA;,
        // sinon la normalisation des attributs le changerait en espace à la relecture (calibration corpus : convocations)
        var settings = new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false, NewLineHandling = NewLineHandling.Entitize };
        using (var w = XmlWriter.Create(ms, settings)) doc.Save(w);
        return ms.ToArray();
    }

    /// <summary>Contenu de chaque entrée tel qu'il sera écrit.</summary>
    public IEnumerable<(Entry Entry, byte[] Data, bool Regenerated)> Output(bool regenerateAll)
    {
        foreach (var e in Entries)
        {
            if (Stories.TryGetValue(e.Name, out var part) && (part.Dirty || regenerateAll)) yield return (e, Serialize(part), true);
            else yield return (e, e.Data, false);
        }
    }

    /// <summary>Écrit le package dans un flux (copie de secours, tests) ; l'état du package ne change pas.</summary>
    public void Save(Stream target, bool regenerateAll = false) => Write(target, Output(regenerateAll));

    /// <summary>
    /// Enregistrement sûr (fichier temporaire puis remplacement, .bak facultatif). Une fois le fichier écrit, le package
    /// adopte le contenu enregistré : les parties régénérées deviennent les octets de référence et ne sont plus « modifiées »
    /// (un enregistrement suivant sans modification recopie exactement ce fichier).
    /// </summary>
    public void Save(string path, bool regenerateAll = false, bool keepBackup = false)
    {
        var output = Output(regenerateAll).ToList();
        SafeFile.Write(path, tmp =>
        {
            using var fs = File.Create(tmp);
            Write(fs, output);
        }, keepBackup);
        for (int i = 0; i < Entries.Count; i++)
            if (output[i].Regenerated) Entries[i] = Entries[i] with { Data = output[i].Data };
        foreach (var part in Stories.Values) part.Dirty = false;
    }

    static void Write(Stream target, IEnumerable<(Entry Entry, byte[] Data, bool Regenerated)> output)
    {
        using var zip = new ZipArchive(target, ZipArchiveMode.Create, leaveOpen: true);
        foreach (var (e, data, _) in output)
        {
            var entry = zip.CreateEntry(e.Name, CompressionLevel.Optimal);
            entry.LastWriteTime = e.LastWrite;
            using var s = entry.Open();
            s.Write(data);
        }
    }
}
