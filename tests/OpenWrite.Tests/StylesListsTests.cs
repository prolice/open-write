using System.Xml.Linq;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;
using OpenWrite.Docx;
using OpenWrite.Layout;

namespace OpenWrite.Tests;

/// <summary>Lot 4 : styles intégrés ajoutés au premier usage, nouveaux styles, listes et numéros, parties XML annulables.</summary>
public class StylesListsTests
{
    static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    const string StylesXml = """
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
          <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
        </w:styles>
        """;

    static Story Body(params string[] texts)
    {
        var xml = XDocument.Parse($"<w:document xmlns:w=\"{W.NamespaceName}\"><w:body>{string.Concat(texts.Select(t => $"<w:p><w:r><w:t>{t}</w:t></w:r></w:p>"))}</w:body></w:document>");
        return StoryMapping.Read(xml, "word/document.xml", "document").Stories().Single();
    }

    [Fact]
    public void Titre_1_ajoute_a_la_premiere_utilisation_et_resolu()
    {
        var styles = XElement.Parse(StylesXml);
        var sheet = new StyleSheet(styles);
        Assert.Null(StyleOps.ByName(sheet, "heading 1"));
        var id = StyleOps.Ensure(sheet, styles, "heading 1");
        Assert.Equal("Titre1", id);
        Assert.Equal(id, StyleOps.Ensure(sheet, styles, "heading 1"));          // une seule fois
        var fresh = new StyleSheet(styles);
        Assert.Equal("Titre 1", StyleOps.DisplayName(fresh, id));
        var p = Body("Titre").Blocks.OfType<Paragraph>().Single();
        p.Props.SetVal("pStyle", id);
        var resolver = new StyleResolver(fresh);
        var pf = resolver.Paragraph(p);
        Assert.True(pf.KeepNext);
        Assert.Equal(0, pf.OutlineLevel);
        Assert.Equal(16, resolver.Run(p, p.Runs().Single()).Size);
        // style de caractère : « Police par défaut » ajoutée comme parent
        var strong = StyleOps.Ensure(fresh, styles, "Strong");
        Assert.NotNull(StyleOps.ByName(new StyleSheet(styles), "Default Paragraph Font"));
        Assert.Equal("character", new StyleSheet(styles).TypeOf(new StyleSheet(styles).Style(strong)!));
    }

    [Fact]
    public void Nouveau_style_a_partir_de_la_selection_reprend_la_mise_en_forme_directe()
    {
        var styles = XElement.Parse(StylesXml);
        var sheet = new StyleSheet(styles);
        var p = Body("x").Blocks.OfType<Paragraph>().Single();
        p.Props.SetVal("jc", "center");
        FormatOps.SetToggle(p.Runs().Single(), "b", true);
        var (ppr, rpr) = StyleOps.Capture(p);
        var id = StyleOps.Create(sheet, styles, "paragraph", "Mon style", "Normal", null, ppr, rpr);
        p.Props.SetVal("pStyle", id);
        StyleOps.Absorb(p, ppr, rpr);
        Assert.Null(p.Props.Get("jc"));
        Assert.Null(p.Runs().Single().Props.Get("b"));
        var resolver = new StyleResolver(new StyleSheet(styles));
        Assert.Equal("center", resolver.Paragraph(p).Alignment);
        Assert.True(resolver.Run(p, p.Runs().Single()).Bold);
    }

    [Fact]
    public void Puces_et_numeros_avec_niveaux_et_recommencer()
    {
        var numbering = new XElement(W + "numbering", new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName));
        var styles = XElement.Parse(StylesXml);
        var body = Body("un", "deux", "sous", "trois", "a", "b");
        var ps = body.Blocks.OfType<Paragraph>().ToList();
        int num = ListOps.Create(numbering, W, ListKind.Numbered);
        int bul = ListOps.Create(numbering, W, ListKind.Bullet);
        Assert.NotEqual(num, bul);
        // abstractNum avant les num (ordre du schéma)
        Assert.True(numbering.Elements().TakeWhile(e => e.Name.LocalName == "abstractNum").Count() == 2);
        ListOps.Set(ps[0], num, 0); ListOps.Set(ps[1], num, 0); ListOps.Set(ps[2], num, 1); ListOps.Set(ps[3], num, 0);
        ListOps.Set(ps[4], bul, 0);
        var sheet = new StyleSheet(styles, numbering);
        Assert.True(ListOps.IsBullet(sheet, bul));
        Assert.False(ListOps.IsBullet(sheet, num));
        var numbers = ListNumbering.Compute(body, new StyleResolver(sheet));
        Assert.Equal(new[] { "1.", "2.", "a.", "3." }, ps.Take(4).Select(p => numbers[p]).ToArray());
        // recommencer à 1 à partir de « trois »
        int restarted = ListOps.Restart(numbering, W, num, 0);
        ListOps.Set(ps[3], restarted, 0);
        numbers = ListNumbering.Compute(body, new StyleResolver(new StyleSheet(styles, numbering)));
        Assert.Equal("1.", numbers[ps[3]]);
        // sortie de liste
        ListOps.Remove(ps[4], styleHasList: false);
        Assert.Null(ps[4].Props.Get("numPr"));
    }

    [Fact]
    public void Liste_de_titres_liee_aux_styles()
    {
        var numbering = new XElement(W + "numbering", new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName));
        var styles = XElement.Parse(StylesXml);
        var sheet = new StyleSheet(styles, numbering);
        var ids = new[] { "heading 1", "heading 2" }.Select(n => StyleOps.Ensure(sheet, styles, n)!).ToList();
        var fresh = new StyleSheet(styles, numbering);
        int numId = ListOps.Create(numbering, W, ListKind.Headings, ids);
        for (int l = 0; l < ids.Count; l++) ListOps.LinkStyle(fresh, fresh.Style(ids[l])!, numId, l);
        var body = Body("A", "B", "C");
        var ps = body.Blocks.OfType<Paragraph>().ToList();
        ps[0].Props.SetVal("pStyle", ids[0]); ps[1].Props.SetVal("pStyle", ids[1]); ps[2].Props.SetVal("pStyle", ids[0]);
        var numbers = ListNumbering.Compute(body, new StyleResolver(new StyleSheet(styles, numbering)));
        Assert.Equal(new[] { "1.", "1.1", "2." }, ps.Select(p => numbers[p]).ToArray());
    }

    [Fact]
    public void Partie_XML_suivie_annulee_et_retablie()
    {
        var styles = XElement.Parse(StylesXml);
        var body = Body("x");
        var journal = new UndoJournal();
        var touched = new List<XElement>();
        journal.XmlChanged += touched.Add;
        using (var tx = journal.Begin("Style", body, Array.Empty<int>(), 0, 1))
        {
            tx.TrackXml(styles);
            var id = StyleOps.Ensure(new StyleSheet(styles), styles, "Quote");
            ((Paragraph)body.Blocks[0]).Props.SetVal("pStyle", id);
            tx.Commit();
        }
        Assert.Equal(2, styles.Elements(W + "style").Count());
        Assert.Single(journal.LastXml);
        journal.Undo();
        Assert.Single(styles.Elements(W + "style"));
        Assert.Null(((Paragraph)body.Blocks[0]).Props.Val("pStyle"));
        journal.Redo();
        Assert.Equal(2, styles.Elements(W + "style").Count());
        Assert.Equal(3, touched.Count);
        // transaction qui échoue : XML remis
        try
        {
            using var tx = journal.Begin("Échec", body, Array.Empty<int>(), 0, 1);
            tx.TrackXml(styles);
            styles.RemoveNodes();
            throw new InvalidOperationException();
        }
        catch (InvalidOperationException) { }
        Assert.Equal(2, styles.Elements(W + "style").Count());
    }

    [Fact]
    public void Numerotation_creee_dans_un_document_qui_n_en_a_pas()
    {
        var docx = RoundTripTests.Docx(RoundTripTests.Root + "<w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>");
        var pkg = DocxPackage.Open(new MemoryStream(docx));
        Assert.Null(pkg.XmlPartOf("numbering"));
        var part = pkg.XmlPartOf("numbering", create: true)!;
        ListOps.Create(part.Doc.Root!, W, ListKind.Bullet);
        var ms = new MemoryStream();
        pkg.Save(ms);
        var reread = DocxPackage.Open(new MemoryStream(ms.ToArray()));
        var again = reread.XmlPartOf("numbering");
        Assert.NotNull(again);
        Assert.Single(again!.Doc.Root!.Elements(W + "abstractNum"));
        var types = DocxPackage.ParseXml(reread.Find("[Content_Types].xml")!.Data).Root!;
        Assert.Contains(types.Elements(), e => e.Attribute("PartName")?.Value == "/word/numbering.xml");
        // un document non modifié reste identique (styles chargés mais pas réécrits)
        var plain = DocxPackage.Open(new MemoryStream(docx));
        plain.XmlPartOf("styles");
        var out2 = new MemoryStream();
        plain.Save(out2);
        Assert.Equal(DocxPackage.Open(new MemoryStream(docx)).Find("word/document.xml")!.Data, DocxPackage.Open(new MemoryStream(out2.ToArray())).Find("word/document.xml")!.Data);
    }

    const string ThemeXml = """
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"><a:themeElements>
          <a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
            <a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1>
            <a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
            <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink>
            <a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
          <a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
            <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>
        </a:themeElements></a:theme>
        """;

    [Fact]
    public void Theme_change_couleur_et_police_des_titres()
    {
        var styles = XElement.Parse(StylesXml);
        var theme = XElement.Parse(ThemeXml);
        var id = StyleOps.Ensure(new StyleSheet(styles, null, theme), styles, "heading 1");
        var p = Body("T").Blocks.OfType<Paragraph>().Single();
        p.Props.SetVal("pStyle", id);
        RunFormat Rf() => new StyleResolver(new StyleSheet(styles, null, theme)).Run(p, p.Runs().Single());
        var before = Rf();
        Assert.Equal("Calibri Light", before.FontAscii);
        Assert.True(ThemeOps.SetColors(theme, "Vert", ThemeOps.Palettes.First(x => x.Name == "Vert").Colors));
        Assert.True(ThemeOps.SetFonts(theme, "Georgia", "Georgia", "Georgia"));
        var after = Rf();
        Assert.Equal("Georgia", after.FontAscii);
        Assert.NotEqual(before.Color, after.Color);
        Assert.Equal(("Georgia", "Georgia"), ThemeOps.Fonts(theme));
        Assert.Equal("549E39", ThemeOps.Accents(theme)[0].Rgb);
    }

    [Fact]
    public void Style_ajoute_remise_en_page_incrementale_identique()
    {
        var styles = XElement.Parse(StylesXml);
        var numbering = new XElement(W + "numbering", new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName));
        var body = Body(Enumerable.Range(0, 400).Select(i => string.Concat(Enumerable.Repeat("mot ", 5 + i % 40))).ToArray());
        body.FinalSection = new PropertyBag("sectPr", W, XElement.Parse($"<w:sectPr xmlns:w=\"{W.NamespaceName}\"><w:pgSz w:w=\"7200\" w:h=\"4000\"/><w:pgMar w:top=\"600\" w:bottom=\"600\" w:left=\"600\" w:right=\"600\"/></w:sectPr>"));
        var fonts = new MonoFontProvider();
        var layout = new DocumentLayout(body, new StyleResolver(new StyleSheet(styles, numbering)), fonts);
        layout.LayoutAll();
        // Titre 1 ajouté et appliqué au paragraphe 200, liste au 201 et 202
        var id = StyleOps.Ensure(new StyleSheet(styles, numbering), styles, "heading 1")!;
        var ps = body.Blocks.OfType<Paragraph>().ToList();
        ps[200].Props.SetVal("pStyle", id);
        int num = ListOps.Create(numbering, W, ListKind.Numbered);
        ListOps.Set(ps[201], num, 0);
        ListOps.Set(ps[202], num, 0);
        var resolver = new StyleResolver(new StyleSheet(styles, numbering));
        layout.ReplaceStyles(resolver);
        layout.Update(new Block[] { ps[200], ps[201], ps[202] }, rebuildFlow: true);
        var full = new DocumentLayout(body, resolver, fonts);
        full.LayoutAll();
        Assert.Equal(full.Pages.Count, layout.Pages.Count);
        for (int k = 0; k < full.Pages.Count; k++)
        {
            Assert.Same(full.Pages[k].Start.Item?.Block, layout.Pages[k].Start.Item?.Block);
            Assert.Equal(full.Pages[k].Lines.Count, layout.Pages[k].Lines.Count);
        }
        var line = layout.Pages.SelectMany(pg => pg.Lines).First(l => ReferenceEquals(l.Paragraph.Paragraph, ps[202]));
        Assert.Equal("2.", line.Paragraph.Marker?.Text);
    }
}
