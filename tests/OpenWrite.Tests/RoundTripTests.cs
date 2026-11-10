using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using OpenWrite.Core.Model;
using OpenWrite.Docx;

namespace OpenWrite.Tests;

/// <summary>Prototype du lot 1 : aller-retour .docx sans perte et modifications locales (documents écrits ici, à la main).</summary>
public class RoundTripTests
{
    internal const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    internal const string Root = """<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:exemple:extension" mc:Ignorable="w14 x">""";

    /// <summary>Cas délicats réunis : identifiants w14, propriétés inconnues, élément d'extension, lien, révisions,
    /// contrôles de contenu, tableau avec exceptions, signets, texte non conservé par Word sans xml:space, saut de section.</summary>
    const string Body = """
        <w:body>
        <w:p w14:paraId="1A2B3C4D" w14:textId="77777777" w:rsidR="00AB12CD"><w:pPr><w:pStyle w:val="Titre1"/><w14:inconnu w14:val="1"/><w:jc w:val="center"/></w:pPr><w:bookmarkStart w:id="0" w:name="debut"/><w:r w:rsidRPr="00112233"><w:rPr><w:rFonts w:ascii="Inter"/><w:sz w:val="28"/><w:lang w:val="fr-BE"/></w:rPr><w:t xml:space="preserve">Ordre du jour </w:t></w:r><w:bookmarkEnd w:id="0"/><x:marque x:v="garder"/></w:p>
        <w:p><w:hyperlink r:id="rId9" w:history="1"><w:r><w:t>lien</w:t></w:r></w:hyperlink><w:ins w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:t>ajouté</w:t></w:r></w:ins><w:del w:id="2" w:author="A"><w:r><w:delText>retiré</w:delText></w:r></w:del><w:r><w:t> </w:t></w:r><w:r><w:tab/><w:br w:type="page"/><w:t>après</w:t><w:sym w:font="Wingdings" w:char="F0E0"/></w:r></w:p>
        <w:sdt><w:sdtPr><w:alias w:val="Bloc"/><w:id w:val="5"/></w:sdtPr><w:sdtContent><w:p><w:sdt><w:sdtPr><w:id w:val="6"/></w:sdtPr><w:sdtContent><w:r><w:t>dans le contrôle</w:t></w:r></w:sdtContent></w:sdt></w:p></w:sdtContent></w:sdt>
        <w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr w:rsidR="00AA"><w:tblPrEx><w:tblBorders/></w:tblPrEx><w:trPr><w:cantSplit/></w:trPr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>cellule</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
        <w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr><w:r><w:t>Fin de la section un</w:t></w:r></w:p>
        <w:sectPr w:rsidR="00AB12CD"><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/></w:sectPr>
        </w:body></w:document>
        """;

    internal static byte[] Docx(string documentXml, string ns = W)
    {
        var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, true))
        {
            void Add(string name, string content)
            {
                using var s = zip.CreateEntry(name).Open();
                s.Write(Encoding.UTF8.GetBytes(content));
            }
            Add("[Content_Types].xml", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>""");
            Add("_rels/.rels", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""");
            Add("word/document.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" + documentXml.Replace(W, ns));
            Add("word/custom.bin", "octets quelconques\u0001\u0002");
        }
        return ms.ToArray();
    }

    internal static string Minify(string xml) => string.Concat(xml.Split('\n').Select(l => l.Trim()));

    internal static (DocxPackage Pkg, byte[] Out) RoundTrip(byte[] docx, Action<DocxPackage>? edit = null, bool all = true)
    {
        var pkg = DocxPackage.Open(new MemoryStream(docx));
        edit?.Invoke(pkg);
        var ms = new MemoryStream();
        pkg.Save(ms, regenerateAll: all);
        return (pkg, ms.ToArray());
    }

    static byte[] Part(byte[] docx, string name)
    {
        using var zip = new ZipArchive(new MemoryStream(docx));
        using var s = zip.GetEntry(name)!.Open();
        using var ms = new MemoryStream();
        s.CopyTo(ms);
        return ms.ToArray();
    }

    internal static XElement DocXml(byte[] docx) => DocxPackage.ParseXml(Part(docx, "word/document.xml")).Root!;

    [Fact]
    public void Aller_retour_sans_aucune_difference()
    {
        var src = Docx(Minify(Root + Body));
        var (pkg, output) = RoundTrip(src);
        Assert.Empty(XmlCompare.Diff(DocXml(src), DocXml(output)));
        Assert.Equal(Part(src, "word/custom.bin"), Part(output, "word/custom.bin"));     // partie inconnue : octet pour octet

        // tout le texte est atteint par le modèle (rien n'est resté dans un îlot opaque)
        var body = pkg.Main!.Stories().Single();
        Assert.Equal(new[] { "Ordre du jour ", "lienajouté \t\naprès", "dans le contrôle", "cellule", "Fin de la section un" },
                     body.Paragraphs().Select(p => p.Text));
        Assert.Equal("Titre1", body.Paragraphs().First().StyleId);
        Assert.NotNull(body.FinalSection);
    }

    [Fact]
    public void Sans_modification_le_document_est_recopie_octet_pour_octet()
    {
        var src = Docx(Minify(Root + Body));
        var (_, output) = RoundTrip(src, all: false);
        Assert.Equal(Part(src, "word/document.xml"), Part(output, "word/document.xml"));
    }

    [Fact]
    public void Une_modification_ne_touche_que_son_paragraphe()
    {
        var src = Docx(Minify(Root + Body));
        var (_, output) = RoundTrip(src, pkg =>
        {
            var p = pkg.Main!.Stories().Single().Paragraphs().ElementAt(4);
            FlowEdits.InsertText(p.Runs().First().Content.OfType<TextContent>().First(), 0, "Vraiment : ");
            pkg.Main.Dirty = true;
        }, all: false);
        var span = XmlCompare.ChangedSpan(DocXml(src).Elements().First().Elements().ToList(), DocXml(output).Elements().First().Elements().ToList());
        Assert.Equal((1, 1), span);
    }

    [Fact]
    public void Le_texte_non_modifie_garde_sa_forme_et_le_texte_modifie_recoit_xml_space()
    {
        var src = Docx(Minify(Root + Body));
        var (_, output) = RoundTrip(src, pkg =>
        {
            var space = pkg.Main!.Stories().Single().Paragraphs().ElementAt(1).Runs().ElementAt(3).Content.OfType<TextContent>().Single();
            Assert.Equal(" ", space.Value);                          // « <w:t> </w:t> » sans xml:space, comme Word 2010
            var title = pkg.Main.Stories().Single().Paragraphs().First().Runs().First().Content.OfType<TextContent>().Single();
            title.Value = " Ordre du jour";
        });
        var ts = DocXml(output).Descendants(XName.Get("t", W)).ToList();
        Assert.Equal("preserve", ts[0].Attribute(XNamespace.Xml + "space")?.Value);
        Assert.Null(ts.Single(t => t.Value == " ").Attribute(XNamespace.Xml + "space"));
    }

    [Fact]
    public void Scinder_un_paragraphe_donne_de_nouveaux_identifiants_et_garde_le_saut_de_section_a_la_fin()
    {
        var src = Docx(Minify(Root + Body));
        Paragraph? second = null;
        var (_, output) = RoundTrip(src, pkg =>
        {
            var story = pkg.Main!.Stories().Single();
            var first = story.Paragraphs().First();
            var run = first.Runs().First();
            var t = run.Content.OfType<TextContent>().Single();
            second = FlowEdits.SplitParagraph(story, first, run, t, 6, FlowEdits.IdGenerator(new[] { "1A2B3C4D" }));
            var last = story.Blocks.OfType<Paragraph>().Last();
            Assert.NotNull(FlowEdits.SplitParagraph(story, last, last.Runs().First(), last.Runs().First().Content.OfType<TextContent>().Single(), 3, FlowEdits.IdGenerator(Array.Empty<string>())));
            pkg.Main.Dirty = true;
        });
        var ps = DocXml(output).Descendants(XName.Get("p", W)).Where(p => p.Parent!.Name.LocalName == "body").ToList();
        XNamespace w14 = FlowEdits.W14;
        Assert.Equal("Ordre ", ps[0].Value);
        Assert.Equal("du jour ", ps[1].Value);
        Assert.Equal("1A2B3C4D", ps[0].Attribute(w14 + "paraId")?.Value);
        Assert.NotEqual("1A2B3C4D", ps[1].Attribute(w14 + "paraId")?.Value);
        Assert.Equal("00AB12CD", ps[1].Attribute(XName.Get("rsidR", W))?.Value);            // les rsid suivent
        Assert.Equal("Inter", ps[1].Descendants(XName.Get("rFonts", W)).Single().Attribute(XName.Get("ascii", W))?.Value);
        // le saut de section reste sur le dernier morceau, comme dans Word
        Assert.Empty(ps[^2].Descendants(XName.Get("sectPr", W)));
        Assert.Single(ps[^1].Descendants(XName.Get("sectPr", W)));
        Assert.Equal(" de la section un", ps[^1].Value);
    }

    [Fact]
    public void Les_proprietes_sont_inserees_dans_l_ordre_du_schema()
    {
        var rPr = XElement.Parse($"""<w:rPr xmlns:w="{W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:rFonts w:ascii="Inter"/><w:sz w:val="28"/><w:lang w:val="fr-BE"/><w14:ligatures w14:val="standard"/><w:rPrChange w:id="3"/></w:rPr>""");
        var bag = new PropertyBag("rPr", W, rPr);
        bag.SetFlag("b", true);
        bag.SetVal("color", "FF0000");
        bag.Set(new XElement(XName.Get("ligatures", FlowEdits.W14.NamespaceName), new XAttribute(FlowEdits.W14 + "val", "none")));
        Assert.Equal(new[] { "rFonts", "b", "color", "sz", "lang", "ligatures", "rPrChange" }, bag.Element!.Elements().Select(e => e.Name.LocalName));
        Assert.True(bag.Flag("b"));
        bag.SetFlag("b", false);
        Assert.Null(bag.Get("b"));
    }

    [Fact]
    public void Un_retour_a_la_ligne_dans_un_attribut_est_conserve()
    {
        var xml = Root + """<w:body><w:p><w:r><w:drawing><x:image descr="ligne 1&#xA;ligne 2"/></w:drawing></w:r></w:p></w:body></w:document>""";
        var src = Docx(xml);
        var (_, output) = RoundTrip(src);
        Assert.Empty(XmlCompare.Diff(DocXml(src), DocXml(output)));
    }

    [Fact]
    public void Enregistrer_deux_fois_cumule_les_modifications_et_garde_la_version_precedente()
    {
        var dir = Path.Combine(Path.GetTempPath(), "openwrite-tests-" + Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(dir);
        try
        {
            var file = Path.Combine(dir, "doc.docx");
            File.WriteAllBytes(file, Docx(Minify(Root + Body)));
            var pkg = DocxPackage.Open(file);
            var t = pkg.Main!.Stories().Single().Paragraphs().First().Runs().First().Content.OfType<TextContent>().Single();
            FlowEdits.InsertText(t, 0, "1 ");
            pkg.Main.Dirty = true;
            pkg.Save(file, keepBackup: true);
            Assert.False(pkg.Main.Dirty);                                    // le package a adopté le contenu enregistré
            FlowEdits.InsertText(t, 0, "2 ");
            pkg.Main.Dirty = true;
            pkg.Save(file, keepBackup: true);
            var again = DocxPackage.Open(file);
            Assert.StartsWith("2 1 Ordre", again.Main!.Stories().Single().Paragraphs().First().Text);
            Assert.StartsWith("1 Ordre", DocxPackage.Open(file + ".bak").Main!.Stories().Single().Paragraphs().First().Text);
            // sans modification, un nouvel enregistrement recopie exactement le fichier
            var before = File.ReadAllBytes(file);
            again.Save(file);
            Assert.Empty(XmlCompare.Diff(DocXml(before), DocXml(File.ReadAllBytes(file))));
        }
        finally { try { Directory.Delete(dir, true); } catch { } }
    }

    [Fact]
    public void Format_strict_ISO_29500()
    {
        var src = Docx(Minify(Root + Body), "http://purl.oclc.org/ooxml/wordprocessingml/main");
        var (pkg, output) = RoundTrip(src, p => FlowEdits.InsertText(p.Main!.Stories().Single().Paragraphs().First().Runs().First().Content.OfType<TextContent>().Single(), 0, "> "));
        Assert.Equal("> Ordre du jour ", pkg.Main!.Stories().Single().Paragraphs().First().Text);
        Assert.Equal("http://purl.oclc.org/ooxml/wordprocessingml/main", DocXml(output).Name.NamespaceName);
        Assert.Single(XmlCompare.Diff(DocXml(src), DocXml(output)));      // le seul écart : le texte modifié
    }
}
