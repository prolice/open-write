using System.Xml.Linq;
using OpenWrite.Core.Model;
using OpenWrite.Docx;
using static OpenWrite.Tests.RoundTripTests;

namespace OpenWrite.Tests;

/// <summary>Lot 1 : champs et signets (documents écrits ici à la main, avec les cas absents du corpus).</summary>
public class MarkerTests
{
    // helpers d'écriture WordprocessingML
    static string R(string inner, string rPr = "") => $"<w:r>{(rPr.Length > 0 ? $"<w:rPr>{rPr}</w:rPr>" : "")}{inner}</w:r>";
    static string Begin(string extra = "") => R($"<w:fldChar w:fldCharType=\"begin\">{extra}</w:fldChar>");
    static string Instr(string code) => R($"<w:instrText xml:space=\"preserve\">{code}</w:instrText>");
    static string Sep => R("<w:fldChar w:fldCharType=\"separate\"/>");
    static string End => R("<w:fldChar w:fldCharType=\"end\"/>");
    static string T(string text) => R($"<w:t xml:space=\"preserve\">{text}</w:t>");
    static string P(params string[] content) => "<w:p>" + string.Concat(content) + "</w:p>";
    static string Doc(params string[] paragraphs) => Root + "<w:body>" + string.Concat(paragraphs) + "<w:sectPr/></w:body></w:document>";

    const string CheckBox = """<w:ffData><w:name w:val="CaseA"/><w:enabled/><w:calcOnExit w:val="0"/><w:checkBox><w:sizeAuto/><w:default w:val="0"/></w:checkBox></w:ffData>""";

    /// <summary>
    /// Table des matières : un champ TOC sur trois paragraphes, avec dans son résultat des liens et des champs PAGEREF
    /// imbriqués ; un IF dont le code contient un MERGEFIELD ; un champ simple ; une case à cocher ; des signets
    /// (dont un au niveau du corps et un qui s'étend sur deux paragraphes).
    /// </summary>
    static readonly string Sample = Doc(
        P(Begin(), Instr(" TOC \\o \"1-3\" \\h \\z \\u "), Sep,
          "<w:hyperlink w:anchor=\"_Toc1\" w:history=\"1\">", T("Introduction"), R("<w:tab/>"), Begin(), Instr(" PAGEREF _Toc1 \\h "), Sep, T("1"), End, "</w:hyperlink>"),
        P("<w:hyperlink w:anchor=\"_Toc2\" w:history=\"1\">", T("Suite"), R("<w:tab/>"), Begin(), Instr(" PAGEREF _Toc2 \\h "), Sep, T("2"), End, "</w:hyperlink>"),
        P(End),
        "<w:bookmarkStart w:id=\"0\" w:name=\"_Toc1\"/>",
        P(T("Introduction")),
        "<w:bookmarkEnd w:id=\"0\"/>",
        P(T("Cher "), Begin(), Instr(" IF "), Begin(), Instr(" MERGEFIELD Genre "), Sep, T("«Genre»"), End, Instr(" = \"F\" \"Madame\" \"Monsieur\" "), Sep, T("Monsieur"), End),
        P("<w:fldSimple w:instr=\" DATE \\@ &quot;d MMMM yyyy&quot; \">", T("6 novembre 2026"), "</w:fldSimple>"),
        P(Begin(CheckBox), Instr(" FORMCHECKBOX "), Sep, End, T(" J'accepte")),
        P("<w:bookmarkStart w:id=\"1\" w:name=\"Conclusion\"/>", T("Fin"), "<w:bookmarkStart w:id=\"2\" w:name=\"_Toc2\"/>", T(" un")),
        P(T("Fin deux"), "<w:bookmarkEnd w:id=\"1\"/>", "<w:bookmarkEnd w:id=\"2\"/>"));

    static Story Body(DocxPackage pkg) => pkg.Main!.Stories().Single();

    [Fact]
    public void Aller_retour_identique_avec_champs_et_signets_types()
    {
        var src = RoundTripTests.Docx(Sample);
        var (pkg, output) = RoundTrip(src);
        Assert.Empty(XmlCompare.Diff(DocXml(src), DocXml(output)));
        // plus aucun îlot opaque : champs et signets sont des nœuds du modèle
        var body = Body(pkg);
        Assert.DoesNotContain(body.Blocks, b => b is OpaqueBlock);
        Assert.Single(body.Blocks.OfType<InlineMarkerBlock>(), m => m.Marker is BookmarkStart);
        Assert.DoesNotContain(FlowWalker.Walk(body), e => e.Item is OpaqueRunContent or OpaqueInline);
    }

    [Fact]
    public void Index_des_champs_imbrication_code_et_resultat()
    {
        var (pkg, _) = RoundTrip(RoundTripTests.Docx(Sample));
        var idx = FieldIndex.Build(Body(pkg));
        Assert.Empty(idx.Problems);
        Assert.Equal(new[] { "TOC", "PAGEREF", "PAGEREF", "IF", "MERGEFIELD", "DATE", "FORMCHECKBOX" }, idx.All.Select(f => f.Type));

        var toc = idx.All[0];
        Assert.Equal(2, toc.Children.Count);
        Assert.NotSame(toc.BeginParagraph, toc.EndParagraph);                 // trois paragraphes
        Assert.Equal("Introduction\t1\rSuite\t2\r", toc.Result);
        Assert.True(toc.Instruction.Has("\\h"));
        Assert.Equal("1-3", toc.Instruction.Arg("\\o"));
        Assert.Equal("_Toc2", idx.All[2].Instruction.Arguments.Single());

        var @if = idx.All[3];
        Assert.Equal(" IF { MERGEFIELD Genre } = \"F\" \"Madame\" \"Monsieur\" ", @if.Code);
        Assert.Equal("Monsieur", @if.Result);                                  // le résultat du MERGEFIELD est dans le code, pas dans le résultat
        Assert.Same(@if, idx.All[4].Parent);
        Assert.Equal("«Genre»", idx.All[4].Result);

        var date = idx.All[5];
        Assert.NotNull(date.Simple);
        Assert.Equal("d MMMM yyyy", date.Instruction.Arg("\\@"));
        Assert.Equal("6 novembre 2026", date.Result);
    }

    [Fact]
    public void Case_a_cocher_de_formulaire_herite()
    {
        var src = RoundTripTests.Docx(Sample);
        var (pkg, output) = RoundTrip(src, p =>
        {
            var box = FieldIndex.Build(Body(p)).All.Single(f => f.Type == "FORMCHECKBOX").CheckBox!;
            Assert.Equal("CaseA", box.Name);
            Assert.False(box.Checked);
            Assert.Equal("☐", box.Glyph);
            box.Checked = true;
            p.Main!.Dirty = true;
        }, all: false);
        var ff = DocXml(output).Descendants(XName.Get("checkBox", W)).Single();
        // w:checked inséré après w:default (ordre du schéma), rien d'autre ne change
        Assert.Equal(new[] { "sizeAuto", "default", "checked" }, ff.Elements().Select(e => e.Name.LocalName));
        Assert.Single(XmlCompare.Diff(DocXml(src), DocXml(output)));
        Assert.True(FieldIndex.Build(Body(DocxPackage.Open(new MemoryStream(output)))).All.Single(f => f.Type == "FORMCHECKBOX").CheckBox!.Checked);

        // revenir à l'état par défaut retire w:checked : le XML redevient celui d'origine
        var (_, back) = RoundTrip(output, p => { FieldIndex.Build(Body(p)).All.Single(f => f.Type == "FORMCHECKBOX").CheckBox!.Checked = false; p.Main!.Dirty = true; }, all: false);
        Assert.Empty(XmlCompare.Diff(DocXml(src), DocXml(back)));
    }

    [Fact]
    public void Remplacer_le_resultat_d_un_champ()
    {
        var src = RoundTripTests.Docx(Doc(
            P(T("Page "), Begin(), Instr(" PAGE "), Sep, R("<w:t>1</w:t>", "<w:b/>"), R("<w:t>0</w:t>"), End),
            P(Begin(), Instr(" NUMPAGES "), Sep, End)));        // pas encore de résultat
        var (pkg, output) = RoundTrip(src, p =>
        {
            var fields = FieldIndex.Build(Body(p)).All;
            Assert.True(MarkerOps.SetResultText(fields[0], "12"));
            Assert.True(MarkerOps.SetResultText(fields[1], "36"));
        });
        var idx = FieldIndex.Build(Body(DocxPackage.Open(new MemoryStream(output))));
        Assert.Equal(new[] { "12", "36" }, idx.All.Select(f => f.Result));
        // le premier segment du résultat (en gras) garde sa mise en forme ; le second, vidé, a disparu
        var page = DocXml(output).Descendants(XName.Get("p", W)).First();
        Assert.Equal("Page 12", page.Value.Replace(" PAGE ", ""));
        Assert.Single(page.Descendants(XName.Get("b", W)));
        // un champ dont le résultat contient d'autres champs (table des matières) n'est pas modifié ainsi
        var (tocPkg, _) = RoundTrip(RoundTripTests.Docx(Sample));
        Assert.False(MarkerOps.SetResultText(FieldIndex.Build(Body(tocPkg)).All[0], "x"));
    }

    [Fact]
    public void Index_des_signets()
    {
        var (pkg, _) = RoundTrip(RoundTripTests.Docx(Sample));
        var idx = BookmarkIndex.Build(Body(pkg));
        Assert.Empty(idx.Problems);
        Assert.Equal(new[] { "_Toc1", "Conclusion", "_Toc2" }, idx.All.Select(b => b.Name));
        Assert.Null(idx.Find("_toc1")!.StartParagraph);                       // placé entre deux blocs
        Assert.Equal("Introduction\r", idx.Find("_TOC1")!.Text);
        Assert.Equal("Fin un\rFin deux", idx.Find("conclusion")!.Text);        // sur deux paragraphes
        Assert.True(idx.Find("_Toc2")!.Hidden);
        Assert.False(idx.Find("Conclusion")!.Hidden);
    }

    [Fact]
    public void Ajouter_renommer_et_supprimer_des_signets()
    {
        var src = RoundTripTests.Docx(Sample);
        var (pkg, output) = RoundTrip(src, p =>
        {
            var body = Body(p);
            var id = MarkerOps.NextId(new[] { body });
            Assert.Equal("3", id);
            var intro = body.Paragraphs().First(x => x.Text == "Introduction");
            MarkerOps.AddAroundParagraph(intro, "Début", id);
            var idx = BookmarkIndex.Build(body);
            MarkerOps.Rename(idx.Find("Conclusion")!, "Fin_du_texte");
            MarkerOps.Remove(body, idx.Find("_Toc1")!);
            Assert.Throws<ArgumentException>(() => MarkerOps.AddAroundParagraph(intro, "1abc", "9"));
            p.Main!.Dirty = true;
        }, all: false);
        var after = BookmarkIndex.Build(Body(DocxPackage.Open(new MemoryStream(output))));
        Assert.Empty(after.Problems);
        Assert.Equal(new[] { "Début", "Fin_du_texte", "_Toc2" }, after.All.Select(b => b.Name));
        Assert.Equal("Introduction", after.Find("Début")!.Text);
    }

    [Theory]
    [InlineData("Signet_1", true)]
    [InlineData("Étape2", true)]
    [InlineData("_Toc123", false)]          // réservé aux signets masqués créés par Word
    [InlineData("1er", false)]
    [InlineData("avec espace", false)]
    [InlineData("a2345678901234567890123456789012345678901", false)]   // 41 caractères
    public void Noms_de_signets(string name, bool ok) => Assert.Equal(ok, MarkerOps.IsValidName(name));

    [Fact]
    public void Champs_et_signets_mal_formes_sont_signales_mais_conserves()
    {
        var src = RoundTripTests.Docx(Doc(
            P(End, T("orphelin")),
            P("<w:bookmarkEnd w:id=\"7\"/>", Begin(), Instr(" PAGE "), Sep, T("3"))));   // fin manquante
        var (pkg, output) = RoundTrip(src);
        Assert.Empty(XmlCompare.Diff(DocXml(src), DocXml(output)));
        Assert.Equal(new[] { "fin de champ sans début", "champ non fermé : {PAGE}" }, FieldIndex.Build(Body(pkg)).Problems);
        Assert.Equal(new[] { "fin de signet sans début : 7" }, BookmarkIndex.Build(Body(pkg)).Problems);
    }

    [Theory]
    [InlineData(" HYPERLINK \"https://exemple.be/a b\" \\l \"ancre\" \\o \"Info\" ", "HYPERLINK", "https://exemple.be/a b", "\\l=ancre;\\o=Info")]
    [InlineData(" REF _Ref123 \\h \\* MERGEFORMAT ", "REF", "_Ref123", "\\h=;\\*=MERGEFORMAT")]
    [InlineData("date \\@ \"dddd d MMMM yyyy\"", "DATE", "", "\\@=dddd d MMMM yyyy")]
    [InlineData(" INCLUDEPICTURE \"C:\\\\Images\\\\logo.png\" \\d ", "INCLUDEPICTURE", "C:\\Images\\logo.png", "\\d=")]
    [InlineData(" SEQ Figure \\* ARABIC ", "SEQ", "Figure", "\\*=ARABIC")]
    public void Analyse_des_codes_de_champ(string code, string type, string firstArg, string switches)
    {
        var fi = FieldInstruction.Parse(code);
        Assert.Equal(type, fi.Type);
        Assert.Equal(firstArg, fi.Arguments.FirstOrDefault() ?? "");
        Assert.Equal(switches, string.Join(";", fi.Switches.Select(s => s.Switch + "=" + s.Argument)));
    }
}
