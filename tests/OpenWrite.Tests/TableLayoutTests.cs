using System.Xml.Linq;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;
using OpenWrite.Docx;
using OpenWrite.Layout;

namespace OpenWrite.Tests;

/// <summary>
/// Lot 6 : styles de tableau conditionnels et mise en page des tableaux (police à chasse fixe : 5 pt par caractère à
/// 10 pt, ligne de 10 pt ; marge gauche de la page 30 pt, colonne de 300 pt).
/// </summary>
public class TableLayoutTests
{
    const string NsDecl = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";

    const string Styles = $$"""
        <w:styles {{NsDecl}}>
          <w:docDefaults>
            <w:rPrDefault><w:rPr><w:sz w:val="20"/></w:rPr></w:rPrDefault>
            <w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>
          </w:docDefaults>
          <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
          <w:style w:type="table" w:default="1" w:styleId="TableauNormal"><w:name w:val="Normal Table"/>
            <w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
          <w:style w:type="table" w:styleId="Grille"><w:name w:val="Grid Table 4 Accent 1"/><w:basedOn w:val="TableauNormal"/>
            <w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="8EAADB"/><w:insideH w:val="single" w:sz="4" w:color="8EAADB"/></w:tblBorders></w:tblPr>
            <w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="4472C4"/></w:tcPr></w:tblStylePr>
            <w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/></w:tcPr></w:tblStylePr>
            <w:tblStylePr w:type="firstCol"><w:rPr><w:b/></w:rPr></w:tblStylePr>
          </w:style>
        </w:styles>
        """;

    static string Settings(int compat) =>
        $"""<w:settings {NsDecl}><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"/></w:compat></w:settings>""";

    const string Sect = """<w:sectPr><w:pgSz w:w="7200" w:h="9000"/><w:pgMar w:top="600" w:bottom="600" w:left="600" w:right="600"/></w:sectPr>""";

    static string P(string text) => $"<w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";

    static string Tc(string content, string props = "") => $"<w:tc><w:tcPr>{props}</w:tcPr>{content}</w:tc>";

    static string Tbl(string props, int[] grid, params string[] rows) =>
        $"<w:tbl><w:tblPr>{props}</w:tblPr><w:tblGrid>{string.Concat(grid.Select(g => $"<w:gridCol w:w=\"{g}\"/>"))}</w:tblGrid>{string.Concat(rows.Select(r => $"<w:tr>{r}</w:tr>"))}</w:tbl>";

    static (Story Body, DocumentLayout Layout) Doc(int compat, params string[] blocks)
    {
        var xml = XDocument.Parse($"<w:document {NsDecl}><w:body>{string.Concat(blocks)}{Sect}</w:body></w:document>");
        var body = StoryMapping.Read(xml, "word/document.xml", "document").Stories().Single();
        var sheet = new StyleSheet(XElement.Parse(Styles), settings: XElement.Parse(Settings(compat)));
        var layout = new DocumentLayout(body, new StyleResolver(sheet), new MonoFontProvider(ascent: 0.8, descent: 0.2));
        layout.LayoutAll();
        return (body, layout);
    }

    static PlacedLine LineOf(DocumentLayout l, string text) =>
        l.Pages.SelectMany(p => p.Lines).First(x => x.Paragraph.Content.Text.StartsWith(text, StringComparison.Ordinal));

    [Fact]
    public void Regions_conditionnelles_comme_Word()
    {
        var look = TableStyles.Look(XElement.Parse($"""<w:tblLook {NsDecl} w:val="04A0"/>"""));
        Assert.Equal(TableLook.Default, look);
        Assert.Equal(TableRegion.FirstRow | TableRegion.FirstCol | TableRegion.NwCell, TableStyles.Regions(look, 0, 5, 0, 4));
        Assert.Equal(TableRegion.Band1Horz, TableStyles.Regions(look, 1, 5, 2, 4));
        Assert.Equal(TableRegion.Band2Horz, TableStyles.Regions(look, 2, 5, 2, 4));
        // bandes verticales (w:noVBand="0") sans première colonne, bandes de deux rangées
        var cols = new TableLook(false, true, false, false, false, false);
        Assert.Equal(TableRegion.Band1Horz | TableRegion.Band2Vert, TableStyles.Regions(cols, 1, 5, 1, 4, rowBand: 2));
        Assert.Equal(TableRegion.LastRow | TableRegion.Band1Vert, TableStyles.Regions(cols, 4, 5, 0, 4));
        // w:cnfStyle : cellules des coins dans l'ordre ne, nw, se, sw (relevé dans les fichiers de Word)
        Assert.Equal("101000000100", TableStyles.CnfValue(TableRegion.FirstRow | TableRegion.FirstCol | TableRegion.NwCell));
        Assert.Equal(TableRegion.FirstRow | TableRegion.NwCell, TableStyles.FromCnf("100000000100"));
    }

    [Fact]
    public void Style_conditionnel_ligne_d_en_tete_et_bandes()
    {
        var look = """<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>""";
        var (_, l) = Doc(15, Tbl($"""<w:tblStyle w:val="Grille"/>{look}""", new[] { 2000, 2000 },
            Tc(P("t1")) + Tc(P("t2")), Tc(P("a1")) + Tc(P("a2")), Tc(P("b1")) + Tc(P("b2"))));
        var header = LineOf(l, "t2").Paragraph;
        Assert.True(header.Content.Styles[0].Format.Bold);
        Assert.Equal("FFFFFF", header.Content.Styles[0].Format.Color);
        Assert.False(LineOf(l, "a2").Paragraph.Content.Styles[0].Format.Bold);
        Assert.True(LineOf(l, "a1").Paragraph.Content.Styles[0].Format.Bold);       // première colonne
        var fills = l.Pages[0].Cells.Select(c => c.Fill).ToArray();
        Assert.Equal(new[] { "4472C4", "4472C4", "D9E2F3", "D9E2F3", null, null }, fills);
        // bordures du style : trait du haut et traits intérieurs horizontaux, couleur gardée
        Assert.Equal("8EAADB", l.Pages[0].Cells[0].Top.Color);
        Assert.Equal(0.5, l.Pages[0].Cells[2].Top.Width, 3);
        Assert.False(l.Pages[0].Cells[0].Left.Visible);
    }

    [Fact]
    public void Retrait_du_tableau_selon_le_mode_de_compatibilite()
    {
        var table = Tbl("", new[] { 2000, 2000 }, Tc(P("x1")) + Tc(P("x2")));
        // Word 2013 et suivants : bord du tableau sur la marge, texte décalé de la marge de cellule
        Assert.Equal(30 + 5.4, LineOf(Doc(15, table).Layout, "x1").X, 3);
        // avant : texte de la première cellule aligné sur la marge
        Assert.Equal(30, LineOf(Doc(14, table).Layout, "x1").X, 3);
        Assert.Equal(30 + 100, LineOf(Doc(14, table).Layout, "x2").X, 3);
    }

    [Fact]
    public void Fusion_verticale_la_derniere_rangee_grandit()
    {
        // cellule fusionnée sur deux rangées d'une ligne, avec 5 lignes de texte : la seconde rangée prend le reste
        var (_, l) = Doc(15, Tbl("", new[] { 2000, 2000 },
            Tc(P("m1") + P("m2") + P("m3") + P("m4") + P("m5"), "<w:vMerge w:val=\"restart\"/>") + Tc(P("r1")),
            Tc(P(""), "<w:vMerge/>") + Tc(P("r2")),
            Tc(P("s1")) + Tc(P("s2"))), P("apres"));
        Assert.Equal(30, LineOf(l, "m1").Y, 3);
        Assert.Equal(40, LineOf(l, "r2").Y, 3);
        Assert.Equal(80, LineOf(l, "s1").Y, 3);
        Assert.Equal(90, LineOf(l, "apres").Y, 3);
        // la suite de la fusion n'affiche pas son paragraphe vide et ne trace pas de trait entre les deux rangées
        var cells = l.Pages[0].Cells;
        Assert.Equal(2, l.Pages[0].Lines.Count(x => x.Paragraph.Content.Text.Length == 0) + 2);
        var merged = cells.Where(c => c.Column == 0 && c.Row < 2).ToList();
        Assert.Equal(2, merged.Count);
        Assert.Same(merged[0].Cell, merged[1].Cell);
    }

    [Fact]
    public void Alignement_vertical_et_marges_de_cellule()
    {
        var (_, l) = Doc(15, Tbl("", new[] { 2000, 2000, 2000 },
            Tc(P("a") + P("b") + P("c")) + Tc(P("centre"), "<w:vAlign w:val=\"center\"/>") + Tc(P("bas"), "<w:vAlign w:val=\"bottom\"/>"),
            Tc(P("u")) + Tc(P("haut"), "<w:tcMar><w:top w:w=\"200\" w:type=\"dxa\"/></w:tcMar>") + Tc(P("v"))));
        Assert.Equal(40, LineOf(l, "centre").Y, 3);
        Assert.Equal(50, LineOf(l, "bas").Y, 3);
        // marge du haut d'une cellule : le texte de toute la rangée descend (Word)
        Assert.Equal(70, LineOf(l, "haut").Y, 3);
        Assert.Equal(70, LineOf(l, "u").Y, 3);
    }

    [Fact]
    public void Tableau_imbrique_et_ajustement_au_contenu()
    {
        var inner = Tbl("", new[] { 1000, 1000 }, Tc(P("i1")) + Tc(P("i2")));
        var (_, l) = Doc(15, Tbl("", new[] { 2000, 2000 }, Tc(P("o1")) + Tc(inner + P(""))), P("apres"));
        Assert.Equal(30 + 100 + 5.4 + 5.4, LineOf(l, "i1").X, 3);
        // marque de fin de cellule vide après le tableau imbriqué : pas de hauteur
        Assert.Equal(40, LineOf(l, "apres").Y, 3);

        // tableau ajusté au contenu (w:tblW et w:tcW automatiques) : chaque colonne prend sa plus longue ligne
        var auto = """<w:tcW w:w="0" w:type="auto"/>""";
        var (_, a) = Doc(15, Tbl("""<w:tblW w:w="0" w:type="auto"/>""", new[] { 3000, 3000 }, Tc(P("aa"), auto) + Tc(P("bbbb bbbb"), auto)));
        Assert.Equal(30 + 5.4 + 10 + 5.4 + 5.4, LineOf(a, "bbbb").X, 3);
    }

    [Fact]
    public void Frappe_dans_une_cellule_reprise_identique_a_une_mise_en_page_complete()
    {
        // 60 rangées sur plusieurs pages (page de 450 pt utiles), rangée d'en-tête répétée
        var rows = new List<string> { "<w:trPr><w:tblHeader/></w:trPr>" + Tc(P("titre")) + Tc(P("valeur")) };
        for (int r = 1; r < 60; r++) rows.Add(Tc(P($"r{r}")) + Tc(P(string.Join(" ", Enumerable.Repeat("mot", r % 7 + 1)))));
        var (body, l) = Doc(15, Tbl("""<w:tblLayout w:type="fixed"/>""", new[] { 2000, 4000 }, rows.ToArray()), P("fin"));
        var table = (Table)body.Blocks[0];
        Assert.True(l.Pages.Count >= 2);
        // texte allongé dans une cellule de la rangée 30 : trois lignes de plus
        var cell = table.Items.OfType<TableRow>().ElementAt(30).Items.OfType<TableCell>().Last();
        var para = (Paragraph)cell.Content.Blocks[0];
        var run = para.Runs().First();
        ((TextContent)run.Content[0]).Value += string.Concat(Enumerable.Repeat(" encore", 30));
        l.Update(new Block[] { table }, dirty: new Block[] { para });
        var fresh = new DocumentLayout(body, new StyleResolver(new StyleSheet(XElement.Parse(Styles), settings: XElement.Parse(Settings(15)))), new MonoFontProvider(ascent: 0.8, descent: 0.2));
        fresh.LayoutAll();
        Assert.Equal(fresh.Pages.Count, l.Pages.Count);
        for (int k = 0; k < fresh.Pages.Count; k++)
        {
            Assert.Equal(fresh.Pages[k].Start.Line, l.Pages[k].Start.Line);
            Assert.Equal(fresh.Pages[k].Lines.Select(x => (x.Paragraph.Paragraph, x.Line, Math.Round(x.Y, 3))),
                         l.Pages[k].Lines.Select(x => (x.Paragraph.Paragraph, x.Line, Math.Round(x.Y, 3))));
        }
    }
}
