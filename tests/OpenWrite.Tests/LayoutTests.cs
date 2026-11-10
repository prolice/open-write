using System.Xml.Linq;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;
using OpenWrite.Docx;
using OpenWrite.Layout;

namespace OpenWrite.Tests;

/// <summary>
/// Lot 2 : coupure des lignes et pagination, avec une police à chasse fixe (5 pt par caractère à 10 pt, ligne de 10 pt).
/// Page de test : colonne de 300 pt (60 caractères), 100 pt de hauteur utile (10 lignes).
/// </summary>
public class LayoutTests
{
    const string NsDecl = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";

    const string Styles = $$"""
        <w:styles {{NsDecl}}>
          <w:docDefaults>
            <w:rPrDefault><w:rPr><w:sz w:val="20"/></w:rPr></w:rPrDefault>
            <w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>
          </w:docDefaults>
          <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:widowControl/></w:pPr></w:style>
          <w:style w:type="paragraph" w:styleId="Libre"><w:name w:val="Libre"/><w:pPr><w:widowControl w:val="0"/></w:pPr></w:style>
          <w:style w:type="paragraph" w:styleId="Titre"><w:name w:val="heading 1"/><w:pPr><w:keepNext/></w:pPr></w:style>
        </w:styles>
        """;

    const string Sect = """<w:sectPr><w:pgSz w:w="7200" w:h="3200"/><w:pgMar w:top="600" w:bottom="600" w:left="600" w:right="600"/></w:sectPr>""";

    static string P(string text, string? style = null, string extra = "") =>
        $"<w:p><w:pPr>{(style is null ? "" : $"<w:pStyle w:val=\"{style}\"/>")}</w:pPr><w:r>{extra}<w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";

    /// <summary>Paragraphe de <paramref name="lines"/> lignes pleines (12 mots de 4 lettres par ligne).</summary>
    static string Lines(int lines, string? style = null) => P(string.Concat(Enumerable.Repeat("aaaa ", 12 * lines)).TrimEnd(), style);

    static (Story Body, DocumentLayout Layout) Doc(params string[] paragraphs)
    {
        var xml = XDocument.Parse($"<w:document {NsDecl}><w:body>{string.Concat(paragraphs)}{Sect}</w:body></w:document>");
        var body = StoryMapping.Read(xml, "word/document.xml", "document").Stories().Single();
        var layout = new DocumentLayout(body, new StyleResolver(new StyleSheet(XElement.Parse(Styles))), new MonoFontProvider(ascent: 0.8, descent: 0.2));
        layout.LayoutAll();
        return (body, layout);
    }

    static int[] LinesPerPage(DocumentLayout l) => l.Pages.Select(p => p.Lines.Count).ToArray();

    [Fact]
    public void Coupure_aux_espaces_et_espaces_de_fin_suspendus()
    {
        // 30 mots de 5 caractères (espace compris) : 12 par ligne (l'espace final dépasse sans forcer la coupure)
        var (_, l) = Doc(P(string.Concat(Enumerable.Repeat("aaaa ", 30)).TrimEnd()));
        var para = l.Pages[0].Lines[0].Paragraph;
        Assert.Equal(new[] { 60, 60, 29 }, para.Lines.Select(x => x.End - x.Start).ToArray());
        Assert.All(para.Lines, x => Assert.Equal(10, x.Height, 3));
    }

    [Fact]
    public void Mot_plus_long_que_la_ligne_coupe_au_caractere()
    {
        var (_, l) = Doc(P(new string('x', 130)));
        Assert.Equal(new[] { 60, 60, 10 }, l.Pages[0].Lines[0].Paragraph.Lines.Select(x => x.End - x.Start).ToArray());
    }

    [Fact]
    public void Interligne_double_exact_et_au_moins()
    {
        var (_, l) = Doc(
            "<w:p><w:pPr><w:spacing w:line=\"480\" w:lineRule=\"auto\"/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>",
            "<w:p><w:pPr><w:spacing w:line=\"300\" w:lineRule=\"exact\"/></w:pPr><w:r><w:t>b</w:t></w:r></w:p>",
            "<w:p><w:pPr><w:spacing w:line=\"100\" w:lineRule=\"atLeast\"/></w:pPr><w:r><w:t>c</w:t></w:r></w:p>");
        Assert.Equal(new[] { 20.0, 15.0, 10.0 }, l.Pages[0].Lines.Select(x => x.Paragraph.Lines[x.Line].Height).ToArray());
    }

    [Fact]
    public void Orpheline_le_paragraphe_entier_passe_a_la_page_suivante()
    {
        // 9 lignes occupées : la première ligne du paragraphe de 3 lignes tiendrait seule en bas de page
        var fill = Enumerable.Repeat(P("x"), 9);
        var (_, l) = Doc(fill.Append(Lines(3)).ToArray());
        Assert.Equal(new[] { 9, 3 }, LinesPerPage(l));
        // sans contrôle des veuves et orphelines : une ligne en bas, deux sur la page suivante
        var (_, libre) = Doc(fill.Append(Lines(3, "Libre")).ToArray());
        Assert.Equal(new[] { 10, 2 }, LinesPerPage(libre));
    }

    [Fact]
    public void Veuve_une_ligne_de_plus_descend()
    {
        // 7 lignes occupées, paragraphe de 4 lignes : 3 tiendraient, la dernière serait seule → 2 + 2
        var (_, l) = Doc(Enumerable.Repeat(P("x"), 7).Append(Lines(4)).ToArray());
        Assert.Equal(new[] { 9, 2 }, LinesPerPage(l));
        Assert.Equal(2, l.Pages[1].Start.Line);
    }

    [Fact]
    public void Saut_de_page_explicite_et_saut_avant()
    {
        var (_, l) = Doc(P("avant"), P("après", extra: "<w:br w:type=\"page\"/>"),
            "<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>nouvelle page</w:t></w:r></w:p>");
        // le saut précède « après » dans son segment : la ligne vide qu'il termine reste en page 1
        Assert.Equal(3, l.Pages.Count);
        Assert.Equal(new[] { 2, 1, 1 }, LinesPerPage(l));
    }

    [Fact]
    public void Saut_de_page_en_fin_de_paragraphe_la_marque_reste_sur_sa_ligne()
    {
        var (body, l) = Doc("<w:p><w:r><w:t>fin</w:t><w:br w:type=\"page\"/></w:r></w:p>", P("suite"));
        Assert.Equal(new[] { 1, 1 }, LinesPerPage(l));
        Assert.Same(body.Blocks[1], l.Pages[1].Start.Item?.Block);
    }

    static string Row(string cell, bool header = false, bool cantSplit = false) =>
        $"<w:tr>{(header || cantSplit ? $"<w:trPr>{(header ? "<w:tblHeader/>" : "")}{(cantSplit ? "<w:cantSplit/>" : "")}</w:trPr>" : "")}<w:tc>{cell}</w:tc></w:tr>";

    static string Table(params string[] rows) =>
        $"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"6216\"/></w:tblGrid>{string.Concat(rows)}</w:tbl>";

    [Fact]
    public void Rangee_de_tableau_coupee_entre_deux_pages_avec_en_tete_repete()
    {
        // 5 lignes occupées, puis un tableau : en-tête d'une ligne, rangée de 8 paragraphes d'une ligne
        var cell = string.Concat(Enumerable.Repeat(P("x"), 8));
        var (_, l) = Doc(Enumerable.Repeat(P("y"), 5).Append(Table(Row(P("en-tête"), header: true), Row(cell))).ToArray());
        // page 1 : 5 + en-tête + 4 lignes de la rangée ; page 2 : en-tête répété + 4 lignes restantes
        Assert.Equal(new[] { 10, 5 }, LinesPerPage(l));
        Assert.Equal(1, l.Pages[1].Start.Line);
        Assert.True(l.Pages[1].Start.Offset > 0);
        Assert.Equal(30, l.Pages[1].Lines.Min(x => x.Y), 3);
    }

    [Fact]
    public void Rangee_insecable_passe_entiere_a_la_page_suivante()
    {
        var cell = string.Concat(Enumerable.Repeat(P("x"), 8));
        var (_, l) = Doc(Enumerable.Repeat(P("y"), 5).Append(Table(Row(cell, cantSplit: true))).ToArray());
        Assert.Equal(new[] { 5, 8 }, LinesPerPage(l));
    }

    [Fact]
    public void Titre_solidaire_du_paragraphe_suivant()
    {
        // 9 lignes occupées : le titre tiendrait, mais pas les deux premières lignes du paragraphe qui suit
        var (_, l) = Doc(Enumerable.Repeat(P("x"), 9).Append(P("Titre", "Titre")).Append(Lines(3)).ToArray());
        Assert.Equal(new[] { 9, 4 }, LinesPerPage(l));
    }

    [Fact]
    public void Section_suivante_sur_une_nouvelle_page()
    {
        var sect = """<w:sectPr><w:pgSz w:w="7200" w:h="3200"/><w:pgMar w:top="600" w:bottom="600" w:left="600" w:right="600"/></w:sectPr>""";
        var (_, l) = Doc($"<w:p><w:pPr>{sect}</w:pPr><w:r><w:t>section 1</w:t></w:r></w:p>", P("section 2"));
        Assert.Equal(2, l.Pages.Count);
    }

    [Fact]
    public void Reprise_incrementale_identique_a_la_mise_en_page_complete()
    {
        var rnd = new Random(7);
        var paras = Enumerable.Range(0, 300).Select(i => Lines(1 + rnd.Next(4), i % 5 == 0 ? "Titre" : null)).ToArray();
        var (body, l) = Doc(paras);
        int pages = l.Pages.Count;
        Assert.True(pages > 50);

        // frappe au milieu : le paragraphe grossit d'une ligne
        var target = (Paragraph)body.Blocks[150];
        target.Runs().First().Content.OfType<TextContent>().First().Value += string.Concat(Enumerable.Repeat(" bbbb", 12));
        l.Update(new Block[] { target });
        Assert.True(l.LastPagesLaidOut < 10, $"pages recalculées : {l.LastPagesLaidOut}");
        AssertSame(l, body);

        // nouveau paragraphe inséré (comme une scission) puis suppression
        var xml = XDocument.Parse($"<w:document {NsDecl}><w:body>{Lines(2)}</w:body></w:document>");
        var inserted = StoryMapping.Read(xml, "word/document.xml", "document").Stories().Single().Blocks[0];
        body.Blocks.Insert(200, inserted);
        l.Update(new[] { inserted });
        AssertSame(l, body);
        body.Blocks.RemoveAt(10);
        l.Update(new[] { body.Blocks[10] });
        AssertSame(l, body);
    }

    static void AssertSame(DocumentLayout incremental, Story body)
    {
        var full = new DocumentLayout(body, new StyleResolver(new StyleSheet(XElement.Parse(Styles))), new MonoFontProvider(ascent: 0.8, descent: 0.2));
        full.LayoutAll();
        Assert.Equal(full.Pages.Count, incremental.Pages.Count);
        for (int i = 0; i < full.Pages.Count; i++)
        {
            Assert.Same(full.Pages[i].Start.Item?.Block, incremental.Pages[i].Start.Item?.Block);
            Assert.Equal(full.Pages[i].Start.Line, incremental.Pages[i].Start.Line);
            Assert.Equal(full.Pages[i].Start.Index, incremental.Pages[i].Start.Index);
            Assert.Equal(full.Pages[i].Lines.Count, incremental.Pages[i].Lines.Count);
        }
    }

    [Fact]
    public void Bordures_de_paragraphe_groupees_comptees_en_haut_et_en_bas_du_groupe()
    {
        // bordure 4/8 pt + 1 pt d'espacement : 1,5 pt au-dessus du groupe et au-dessous, rien entre deux paragraphes du groupe
        const string bdr = "<w:pBdr><w:top w:val=\"single\" w:sz=\"4\" w:space=\"1\"/><w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"1\"/></w:pBdr>";
        var (_, l) = Doc($"<w:p><w:pPr>{bdr}</w:pPr><w:r><w:t>a</w:t></w:r></w:p>", $"<w:p><w:pPr>{bdr}</w:pPr><w:r><w:t>b</w:t></w:r></w:p>", P("c"));
        Assert.Equal(new[] { 31.5, 41.5, 53.0 }, l.Pages[0].Lines.Select(x => Math.Round(x.Y, 2)).ToArray());
    }

    [Fact]
    public void Numeros_de_liste_par_niveau_avec_remise_a_zero_et_formats()
    {
        const string numbering = $$"""
            <w:numbering {{NsDecl}}>
              <w:abstractNum w:abstractNumId="0">
                <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
                <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2)"/></w:lvl>
              </w:abstractNum>
              <w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1 -"/></w:lvl></w:abstractNum>
              <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
              <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
            </w:numbering>
            """;
        static string L(int num, int lvl, string t) => $"<w:p><w:pPr><w:numPr><w:ilvl w:val=\"{lvl}\"/><w:numId w:val=\"{num}\"/></w:numPr></w:pPr><w:r><w:t>{t}</w:t></w:r></w:p>";
        var xml = XDocument.Parse($"<w:document {NsDecl}><w:body>{L(1, 0, "a")}{L(1, 1, "b")}{L(1, 1, "c")}{L(1, 0, "d")}{L(1, 1, "e")}{L(2, 0, "f")}{L(2, 0, "g")}</w:body></w:document>");
        var body = StoryMapping.Read(xml, "word/document.xml", "document").Stories().Single();
        var styles = new StyleResolver(new StyleSheet(XElement.Parse(Styles), XElement.Parse(numbering)));
        var n = ListNumbering.Compute(body, styles);
        Assert.Equal(new[] { "1.", "1.a)", "1.b)", "2.", "2.a)", "III -", "IV -" }, body.Paragraphs().Select(p => n[p]).ToArray());
    }

    [Fact]
    public void Hauteur_de_ligne_de_Word_formule_GDI()
    {
        // Times New Roman : interligne externe de 87 unités → 13,8 pt à 12 pt (valeur de Word) ; Calibri : 0 (1,2207 em)
        var tnr = new FontVerticalMetrics(2048, 1825, 443, 1825, -443, 87);
        Assert.Equal(87, tnr.ExternalLeading);
        Assert.Equal(13.8, (tnr.WinAscent + tnr.WinDescent + tnr.ExternalLeading) / 2048.0 * 12, 1);
        var calibri = new FontVerticalMetrics(2048, 1950, 550, 1536, -512, 452);
        Assert.Equal(0, calibri.ExternalLeading);
    }

    static string ColsSect(int num, string type = "nextPage", bool sep = false) =>
        $"<w:sectPr><w:type w:val=\"{type}\"/><w:pgSz w:w=\"7200\" w:h=\"3200\"/><w:pgMar w:top=\"600\" w:bottom=\"600\" w:left=\"600\" w:right=\"600\"/>" +
        $"<w:cols w:num=\"{num}\" w:space=\"400\"{(sep ? " w:sep=\"1\"" : "")}/></w:sectPr>";

    static (Story Body, DocumentLayout Layout) DocWithSect(string sect, params string[] paragraphs)
    {
        var xml = XDocument.Parse($"<w:document {NsDecl}><w:body>{string.Concat(paragraphs)}{sect}</w:body></w:document>");
        var body = StoryMapping.Read(xml, "word/document.xml", "document").Stories().Single();
        var layout = new DocumentLayout(body, new StyleResolver(new StyleSheet(XElement.Parse(Styles))), new MonoFontProvider(ascent: 0.8, descent: 0.2));
        layout.LayoutAll();
        return (body, layout);
    }

    [Fact]
    public void Deux_colonnes_le_texte_passe_a_la_colonne_suivante_puis_a_la_page()
    {
        // colonnes de (300 - 20) / 2 = 140 pt ; 10 lignes par colonne ; 25 paragraphes d'une ligne
        var (_, l) = DocWithSect(ColsSect(2, sep: true), Enumerable.Repeat(P("x"), 25).ToArray());
        Assert.Equal(2, l.Pages.Count);
        var p0 = l.Pages[0];
        Assert.Equal(20, p0.Lines.Count);
        Assert.Equal(new[] { 30.0, 190.0 }, p0.Lines.Select(x => x.X).Distinct().OrderBy(x => x).ToArray());
        Assert.Equal(p0.Lines[0].Y, p0.Lines[10].Y, 3);                      // la 2e colonne repart en haut
        Assert.Equal(140, p0.Lines[0].Paragraph.Width, 3);
        Assert.Single(p0.ColumnRules);
        Assert.Equal(180, p0.ColumnRules[0].X, 3);                          // milieu de l'espace entre les colonnes
    }

    [Fact]
    public void Saut_de_colonne_et_section_continue_equilibree()
    {
        // saut de colonne : le paragraphe suivant commence en haut de la 2e colonne
        var (_, l) = DocWithSect(ColsSect(2), P("a"), "<w:p><w:r><w:t>b</w:t><w:br w:type=\"column\"/></w:r></w:p>", P("c"));
        Assert.Single(l.Pages);
        var c = l.Pages[0].Lines.Single(x => x.Paragraph.Content.Text == "c");
        Assert.Equal(190, c.X, 3);
        Assert.Equal(l.Pages[0].Lines[0].Y, c.Y, 3);
        // section continue de 2 colonnes (6 lignes) suivie d'une section à une colonne : 3 lignes par colonne
        var twoCols = "<w:p><w:pPr>" + ColsSect(2, "continuous") + "</w:pPr><w:r><w:t>6</w:t></w:r></w:p>";
        var (_, l2) = DocWithSect("<w:sectPr><w:type w:val=\"continuous\"/><w:pgSz w:w=\"7200\" w:h=\"3200\"/><w:pgMar w:top=\"600\" w:bottom=\"600\" w:left=\"600\" w:right=\"600\"/></w:sectPr>",
            P("1"), P("2"), P("3"), P("4"), P("5"), twoCols, P("après"));
        var lines = l2.Pages[0].Lines;
        Assert.Equal(3, lines.Count(x => x.X > 100));
        var after = lines.Single(x => x.Paragraph.Content.Text == "après");
        Assert.Equal(30, after.X, 3);
        Assert.Equal(lines[0].Y + 30, after.Y, 3);                           // sous les colonnes équilibrées
    }

    [Fact]
    public void Sections_orientation_marges_colonnes_et_saut()
    {
        var w = (XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var sect = XElement.Parse($"<w:sectPr {NsDecl}><w:headerReference w:type=\"default\" r:id=\"rId8\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1417\" w:right=\"1000\" w:bottom=\"1417\" w:left=\"2000\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/></w:sectPr>");
        Core.Editing.SectionOps.SetOrientation(sect, true);
        var (pw, ph, land) = Core.Editing.SectionOps.PageSize(sect);
        Assert.True(land);
        Assert.Equal(16838 / 20.0, pw, 2);
        Assert.Equal("landscape", sect.Element(w + "pgSz")!.Attribute(w + "orient")!.Value);
        // comme Word (vérifié par COM) : haut ← gauche, bas ← droite, gauche ← bas, droite ← haut
        Assert.Equal("2000", sect.Element(w + "pgMar")!.Attribute(w + "top")!.Value);
        Assert.Equal("1000", sect.Element(w + "pgMar")!.Attribute(w + "bottom")!.Value);
        Core.Editing.SectionOps.SetColumns(sect, 3, 18, separator: true);
        Assert.Equal(3, SectionGeometry.From(sect).ColumnCount);
        Assert.True(SectionGeometry.From(sect).ColumnSeparator);
        // ordre du schéma : cols après pgMar
        Assert.True(sect.Elements().Select(e => e.Name.LocalName).ToList().IndexOf("cols") > sect.Elements().Select(e => e.Name.LocalName).ToList().IndexOf("pgMar"));

        // saut de section continu au milieu d'un paragraphe : avant = propriétés et en-têtes ; après = continu, en-têtes hérités
        var xml = XDocument.Parse($"<w:document {NsDecl}><w:body>{P("abcdef")}{sect}</w:body></w:document>");
        var body = StoryMapping.Read(xml, "word/document.xml", "document").Stories().Single();
        Core.Editing.SectionOps.InsertBreak(body, new Core.Model.DocPos(new[] { 0 }, 3), "continuous", FlowEdits.IdGenerator(Array.Empty<string>()));
        Assert.Equal(2, body.Blocks.Count);
        var first = ((Paragraph)body.Blocks[0]).Props.Get("sectPr")!;
        Assert.NotNull(first.Element(w + "headerReference"));
        var final = body.FinalSection!.Element!;
        Assert.Null(final.Element(w + "headerReference"));
        Assert.Equal("continuous", final.Element(w + "type")!.Attribute(w + "val")!.Value);
        var layout = new DocumentLayout(body, new StyleResolver(new StyleSheet(XElement.Parse(Styles))), new MonoFontProvider(), related: _ => new Story());
        layout.LayoutAll();
        Assert.Equal(2, layout.Sections.Count);
        Assert.Equal("rId8", layout.Sections[1].HeaderRefs["header:default"]);      // hérité
    }

    [Fact]
    public void Numerotation_des_pages_qui_recommence_en_romain()
    {
        // section 1 : 2 pages ; section 2 : départ à 1 en chiffres romains
        var s1 = "<w:p><w:pPr><w:sectPr><w:pgSz w:w=\"7200\" w:h=\"3200\"/><w:pgMar w:top=\"600\" w:bottom=\"600\" w:left=\"600\" w:right=\"600\"/></w:sectPr></w:pPr><w:r><w:t>fin 1</w:t></w:r></w:p>";
        var s2 = "<w:sectPr><w:pgSz w:w=\"7200\" w:h=\"3200\"/><w:pgMar w:top=\"600\" w:bottom=\"600\" w:left=\"600\" w:right=\"600\"/><w:pgNumType w:start=\"1\" w:fmt=\"upperRoman\"/></w:sectPr>";
        var paras = Enumerable.Repeat(P("x"), 14).Append(s1).Concat(Enumerable.Repeat(P("y"), 14)).ToArray();
        var (_, l) = DocWithSect(s2, paras);
        Assert.Equal(new[] { "1", "2", "I", "II" }, l.Pages.Select(p => p.NumberText).ToArray());
        Assert.Equal(new[] { 2, 2, 2, 2 }, l.Pages.Select(p => p.SectionPageCount).ToArray());
    }
}
