using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;
using OpenWrite.Docx;

namespace OpenWrite.Tests;

/// <summary>
/// Fichier ▸ Nouveau document : le document vierge est valide pour le SDK, a un paragraphe vide, les styles par défaut,
/// un thème, la langue et le format de page demandés ; on y écrit, il s'enregistre et se relit.
/// </summary>
public class BlankDocumentTests
{
    static List<string> Errors(byte[] bytes)
    {
        using var wd = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        return new OpenXmlValidator().Validate(wd).Select(e => $"{e.Part?.Uri} {e.Path?.XPath} {e.Description}").ToList();
    }

    [Fact]
    public void Document_vierge_valide_et_complet()
    {
        var bytes = BlankDocument.Create("fr-BE", "Alice");
        var errors = Errors(bytes);
        Assert.True(errors.Count == 0, string.Join("\n", errors));

        var pkg = DocxPackage.Open(new MemoryStream(bytes));
        var body = pkg.Main!.Stories().First();
        var p = Assert.Single(body.Paragraphs());
        Assert.Empty(p.Runs());
        var sheet = pkg.StyleSheet();
        Assert.NotNull(StyleOps.ByName(sheet, "Normal"));
        Assert.Equal(("Calibri Light", "Calibri"), ThemeOps.Fonts(pkg.XmlPartOf("theme")!.Doc.Root!));
        var resolved = new StyleResolver(sheet).ParagraphMark(p);
        Assert.Equal(11, resolved.Size);
        Assert.Equal("fr-BE", resolved.Language);
        var info = DocxInfo.Read(new MemoryStream(bytes));
        Assert.Equal(15, info.CompatibilityMode);
        Assert.Equal("11906", body.FinalSection!.Get("pgSz")!.Attributes().First(a => a.Name.LocalName == "w").Value);
    }

    [Fact]
    public void Format_lettre_et_separateurs_selon_la_langue()
    {
        var us = DocxPackage.Open(new MemoryStream(BlankDocument.Create("en-US")));
        var pgSz = us.Main!.Stories().First().FinalSection!.Get("pgSz")!;
        Assert.Equal("12240", pgSz.Attributes().First(a => a.Name.LocalName == "w").Value);
        var settings = us.XmlPartOf("settings")!.Doc.Root!;
        Assert.Equal(".", settings.Elements().First(e => e.Name.LocalName == "decimalSymbol").Attributes().First().Value);
        // deux documents vierges n'ont pas le même identifiant
        Assert.NotEqual(us.XmlPartOf("settings")!.Doc.ToString(),
                        DocxPackage.Open(new MemoryStream(BlankDocument.Create("en-US"))).XmlPartOf("settings")!.Doc.ToString());
    }

    [Fact]
    public void On_ecrit_dans_le_document_vierge_puis_il_se_relit()
    {
        var pkg = DocxPackage.Open(new MemoryStream(BlankDocument.Create("fr-FR")));
        var body = pkg.Main!.Stories().First();
        var journal = new UndoJournal();
        journal.Changed += root => pkg.Main!.Dirty = true;
        TextOps.Type(journal, body, new DocPos(new[] { 0 }, 0), "Bonjour le monde");
        using var ms = new MemoryStream();
        pkg.Save(ms);
        var errors = Errors(ms.ToArray());
        Assert.True(errors.Count == 0, string.Join("\n", errors));
        var again = DocxPackage.Open(new MemoryStream(ms.ToArray()));
        var text = string.Concat(again.Main!.Stories().First().Paragraphs().SelectMany(x => x.Runs()).SelectMany(r => r.Content).OfType<TextContent>().Select(t => t.Value));
        Assert.Equal("Bonjour le monde", text);
    }
}
