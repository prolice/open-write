using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Security;

namespace OpenWrite.Docx;

/// <summary>
/// Document vierge (Fichier ▸ Nouveau document, Ctrl+N), comme Word en écrit un : parties relevées sur un document
/// vierge de Word (COM, Normal.dotm d'origine) — thème Office, styles par défaut (Normal, Police par défaut, Tableau
/// normal, Aucune liste ; les autres styles intégrés sont ajoutés à leur première utilisation), styles latents,
/// paramètres de compatibilité 15, marges de 2,54 cm. Ressources dans <c>Blank/</c>.
/// La langue de rédaction, le format de page (A4, ou Letter hors système métrique), les séparateurs décimal et de
/// liste suivent la langue donnée ; les identifiants du document sont nouveaux à chaque fois.
/// </summary>
public static class BlankDocument
{
    const string Ct = "application/vnd.openxmlformats-officedocument.";

    public static byte[] Create(string language, string? author = null, DateTime? nowUtc = null)
    {
        CultureInfo culture;
        try { culture = CultureInfo.GetCultureInfo(language); }
        catch (CultureNotFoundException) { culture = CultureInfo.GetCultureInfo("fr-FR"); language = "fr-FR"; }
        bool metric = true;
        try { metric = culture.IsNeutralCulture || new RegionInfo(culture.Name).IsMetric; } catch (ArgumentException) { }
        var now = (nowUtc ?? DateTime.UtcNow).ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);
        string Hex8() => Random.Shared.Next(1, int.MaxValue).ToString("X8");
        var values = new Dictionary<string, string>
        {
            ["LANG"] = language,
            ["PGW"] = metric ? "11906" : "12240",
            ["PGH"] = metric ? "16838" : "15840",
            ["PARAID"] = Hex8(),
            ["DOCID14"] = Hex8(),
            ["DOCID15"] = Guid.NewGuid().ToString("D").ToUpperInvariant(),
            ["DECIMAL"] = SecurityElement.Escape(culture.NumberFormat.NumberDecimalSeparator),
            ["LISTSEP"] = SecurityElement.Escape(culture.TextInfo.ListSeparator),
        };
        string Res(string name)
        {
            using var s = typeof(BlankDocument).Assembly.GetManifestResourceStream("OpenWrite.Docx.Blank." + name)
                          ?? throw new InvalidOperationException("Ressource manquante : " + name);
            var text = new StreamReader(s).ReadToEnd();
            foreach (var (k, v) in values) text = text.Replace("{" + k + "}", v);
            return text;
        }

        const string Rels = "http://schemas.openxmlformats.org/package/2006/relationships";
        const string OfficeRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
        const string Xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";
        var parts = new List<(string Name, string Text)>
        {
            ("[Content_Types].xml", Xml + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + $"<Override PartName=\"/word/document.xml\" ContentType=\"{Ct}wordprocessingml.document.main+xml\"/>"
                + $"<Override PartName=\"/word/styles.xml\" ContentType=\"{Ct}wordprocessingml.styles+xml\"/>"
                + $"<Override PartName=\"/word/settings.xml\" ContentType=\"{Ct}wordprocessingml.settings+xml\"/>"
                + $"<Override PartName=\"/word/webSettings.xml\" ContentType=\"{Ct}wordprocessingml.webSettings+xml\"/>"
                + $"<Override PartName=\"/word/fontTable.xml\" ContentType=\"{Ct}wordprocessingml.fontTable+xml\"/>"
                + $"<Override PartName=\"/word/theme/theme1.xml\" ContentType=\"{Ct}theme+xml\"/>"
                + "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>"
                + $"<Override PartName=\"/docProps/app.xml\" ContentType=\"{Ct}extended-properties+xml\"/></Types>"),
            ("_rels/.rels", Xml + $"<Relationships xmlns=\"{Rels}\">"
                + $"<Relationship Id=\"rId3\" Type=\"{OfficeRel}extended-properties\" Target=\"docProps/app.xml\"/>"
                + "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>"
                + $"<Relationship Id=\"rId1\" Type=\"{OfficeRel}officeDocument\" Target=\"word/document.xml\"/></Relationships>"),
            ("word/document.xml", Res("document.xml")),
            ("word/_rels/document.xml.rels", Xml + $"<Relationships xmlns=\"{Rels}\">"
                + $"<Relationship Id=\"rId3\" Type=\"{OfficeRel}webSettings\" Target=\"webSettings.xml\"/>"
                + $"<Relationship Id=\"rId2\" Type=\"{OfficeRel}settings\" Target=\"settings.xml\"/>"
                + $"<Relationship Id=\"rId1\" Type=\"{OfficeRel}styles\" Target=\"styles.xml\"/>"
                + $"<Relationship Id=\"rId5\" Type=\"{OfficeRel}theme\" Target=\"theme/theme1.xml\"/>"
                + $"<Relationship Id=\"rId4\" Type=\"{OfficeRel}fontTable\" Target=\"fontTable.xml\"/></Relationships>"),
            ("word/theme/theme1.xml", Res("theme1.xml")),
            ("word/settings.xml", Res("settings.xml")),
            ("word/styles.xml", Res("styles.xml")),
            ("word/webSettings.xml", Res("webSettings.xml")),
            ("word/fontTable.xml", Res("fontTable.xml")),
            ("docProps/core.xml", Xml + "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" "
                + "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" "
                + "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dc:title></dc:title><dc:subject></dc:subject>"
                + $"<dc:creator>{SecurityElement.Escape(author ?? "")}</dc:creator><cp:keywords></cp:keywords><dc:description></dc:description>"
                + $"<cp:lastModifiedBy>{SecurityElement.Escape(author ?? "")}</cp:lastModifiedBy><cp:revision>1</cp:revision>"
                + $"<dcterms:created xsi:type=\"dcterms:W3CDTF\">{now}</dcterms:created><dcterms:modified xsi:type=\"dcterms:W3CDTF\">{now}</dcterms:modified></cp:coreProperties>"),
            ("docProps/app.xml", Xml + "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" "
                + "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><Template>Normal.dotm</Template><TotalTime>0</TotalTime>"
                + "<Pages>1</Pages><Words>0</Words><Characters>0</Characters><Application>OpenWrite</Application><DocSecurity>0</DocSecurity>"
                + "<Lines>0</Lines><Paragraphs>0</Paragraphs><ScaleCrop>false</ScaleCrop><Company></Company><LinksUpToDate>false</LinksUpToDate>"
                + "<CharactersWithSpaces>0</CharactersWithSpaces><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged></Properties>"),
        };
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (name, text) in parts)
            {
                var e = zip.CreateEntry(name, CompressionLevel.Optimal);
                using var w = new StreamWriter(e.Open(), new System.Text.UTF8Encoding(false));
                w.Write(text);
            }
        }
        return ms.ToArray();
    }
}
