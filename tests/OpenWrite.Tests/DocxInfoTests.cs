using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenWrite.Docx;

namespace OpenWrite.Tests;

/// <summary>Résumé d'un .docx (version 0.1), sur des documents fabriqués ici avec l'Open XML SDK.</summary>
public class DocxInfoTests
{
    static MemoryStream Build(Action<MainDocumentPart> fill)
    {
        var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            doc.PackageProperties.Title = "Compte rendu";
            doc.PackageProperties.Creator = "Secrétariat";
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body());
            fill(main);
        }
        ms.Position = 0;
        return ms;
    }

    static Paragraph P(params OpenXmlElement[] runs) => new(runs);
    static Run R(string text) => new(new Text(text) { Space = SpaceProcessingModeValues.Preserve });

    [Fact]
    public void Compte_paragraphes_mots_tableaux_et_sections()
    {
        using var ms = Build(main =>
        {
            var body = main.Document.Body!;
            body.Append(P(R("Ordre "), R("du jour")));                          // un mot coupé en deux segments
            body.Append(P(new Run(new Text("Point 1")), new Run(new TabChar()), new Run(new Text("ouvert"))));
            body.Append(new Table(new TableRow(new TableCell(P(R("cellule"))))));
            body.Append(P(new ParagraphProperties(new SectionProperties())));   // saut de section
            body.Append(P(new InsertedRun(R("ajouté")) { Author = "A", Id = "1" },
                          new DeletedRun(new Run(new DeletedText("retiré"))) { Author = "A", Id = "2" }));
            body.Append(new SectionProperties());

            var settings = main.AddNewPart<DocumentSettingsPart>();
            settings.Settings = new Settings(new Compatibility(
                new DoNotExpandShiftReturn(),
                new CompatibilitySetting { Name = CompatSettingNameValues.CompatibilityMode, Uri = "http://schemas.microsoft.com/office/word", Val = "14" }));
        });

        var info = DocxInfo.Read(ms);
        Assert.Equal("Compte rendu", info.Title);
        Assert.Equal("Secrétariat", info.Author);
        Assert.Equal(5, info.Paragraphs);                     // dont celui de la cellule
        Assert.Equal(1, info.Tables);
        Assert.Equal(2, info.Sections);
        Assert.Equal(2, info.Revisions);
        Assert.Equal(14, info.CompatibilityMode);
        Assert.Equal(new[] { "doNotExpandShiftReturn" }, info.LegacyCompatOptions);
        // « Ordre du jour » (3) + « Point 1 ouvert » (3 : tabulation = séparateur) + cellule (1) + « ajouté » (1) ; « retiré » ne compte pas
        Assert.Equal("Ordre du jour", info.Preview[0]);
        Assert.Equal("Point 1\touvert", info.Preview[1]);
        Assert.Equal("", info.Preview[2]);                  // paragraphe du saut de section
        Assert.Equal("ajouté", info.Preview[3]);
        Assert.Equal(8, info.Words);
        Assert.Contains(info.Parts, p => p.Uri == "/word/document.xml");
    }

    [Fact]
    public void Ignore_les_separateurs_de_notes()
    {
        using var ms = Build(main =>
        {
            main.Document.Body!.Append(P(R("Texte"), new Run(new FootnoteReference { Id = 1 })));
            var fn = main.AddNewPart<FootnotesPart>();
            fn.Footnotes = new Footnotes(
                new Footnote(P(new Run(new SeparatorMark()))) { Type = FootnoteEndnoteValues.Separator, Id = -1 },
                new Footnote(P(new Run(new ContinuationSeparatorMark()))) { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 },
                new Footnote(P(R("La note"))) { Id = 1 });
        });
        var info = DocxInfo.Read(ms);
        Assert.Equal(1, info.Footnotes);
        Assert.Equal(12, info.CompatibilityMode);           // non précisé : mode Word 2007, comme Word
        Assert.Equal(3, info.Words);                         // « Texte » + « La note »
    }

    [Fact]
    public void Un_fichier_qui_n_est_pas_un_docx_est_refuse_clairement()
    {
        using var ms = new MemoryStream("ceci n'est pas un zip"u8.ToArray());
        var ex = Assert.Throws<DocxReadException>(() => DocxInfo.Read(ms));
        Assert.Contains("pas un document Word", ex.Message);
    }

    [Fact]
    public void Titres_par_niveau_de_plan_des_styles_herites()
    {
        using var ms = Build(main =>
        {
            var styles = main.AddNewPart<StyleDefinitionsPart>();
            styles.Styles = new Styles(
                new Style(new StyleName { Val = "heading 1" }, new StyleParagraphProperties(new OutlineLevel { Val = 0 })) { Type = StyleValues.Paragraph, StyleId = "Titre1" },
                new Style(new StyleName { Val = "Titre maison" }, new BasedOn { Val = "Titre1" }) { Type = StyleValues.Paragraph, StyleId = "MonTitre" },
                new Style(new StyleName { Val = "Normal" }) { Type = StyleValues.Paragraph, StyleId = "Normal" });
            var body = main.Document.Body!;
            body.Append(new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Titre1" }), R("Introduction")));
            body.Append(new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Normal" }), R("Texte")));
            body.Append(new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "MonTitre" }), R("Hérité")));
            body.Append(new Paragraph(new ParagraphProperties(new OutlineLevel { Val = 1 }), R("Direct")));
            body.Append(new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Titre1" })));     // titre vide : ignoré
        });
        var info = DocxInfo.Read(ms);
        Assert.Equal(new[] { (1, "Introduction"), (1, "Hérité"), (2, "Direct") }, info.Headings);
    }

    [Theory]
    [InlineData(11, "Word 97-2003 (mode de compatibilité)")]
    [InlineData(12, "Word 2007 (mode de compatibilité)")]
    [InlineData(14, "Word 2010 (mode de compatibilité)")]
    [InlineData(15, "Word 2013 et suivants")]
    public void Libelle_du_mode_de_compatibilite(int mode, string expected) =>
        Assert.Equal(expected, DocxInfo.CompatibilityLabel(mode));
}
