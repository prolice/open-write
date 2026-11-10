using System.Xml.Linq;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;
using OpenWrite.Docx;

namespace OpenWrite.Tests;

/// <summary>Lot 1 : résolution des propriétés héritées (ordre de Word, propriétés bascule, thème, numérotation).</summary>
public class StyleResolverTests
{
    const string NsDecl = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";

    const string Styles = $$"""
        <w:styles {{NsDecl}}>
          <w:docDefaults>
            <w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="22"/><w:lang w:val="fr-BE"/></w:rPr></w:rPrDefault>
            <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>
          </w:docDefaults>
          <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="left" w:pos="1000"/><w:tab w:val="left" w:pos="3000"/></w:tabs></w:pPr></w:style>
          <w:style w:type="paragraph" w:styleId="Titre1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
            <w:pPr><w:keepNext/><w:spacing w:before="240" w:after="0"/><w:ind w:left="360" w:hanging="360"/><w:outlineLvl w:val="0"/></w:pPr>
            <w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/><w:b/><w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/><w:sz w:val="32"/></w:rPr></w:style>
          <w:style w:type="paragraph" w:styleId="Titre1Rouge"><w:name w:val="Titre rouge"/><w:basedOn w:val="Titre1"/>
            <w:pPr><w:ind w:firstLine="200"/><w:tabs><w:tab w:val="clear" w:pos="1000"/></w:tabs></w:pPr><w:rPr><w:color w:val="FF0000"/></w:rPr></w:style>
          <w:style w:type="character" w:styleId="Accentuation"><w:name w:val="Strong"/><w:rPr><w:b/><w:i/></w:rPr></w:style>
          <w:style w:type="paragraph" w:styleId="ListePuces"><w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/>
            <w:pPr><w:numPr><w:numId w:val="1"/></w:numPr><w:ind w:right="100"/></w:pPr></w:style>
          <w:style w:type="table" w:default="1" w:styleId="TableauNormal"><w:name w:val="Normal Table"/></w:style>
          <w:style w:type="table" w:styleId="Grille"><w:name w:val="Table Grid"/><w:pPr><w:spacing w:after="0"/></w:pPr><w:rPr><w:sz w:val="18"/></w:rPr></w:style>
          <w:style w:type="paragraph" w:styleId="Boucle1"><w:name w:val="b1"/><w:basedOn w:val="Boucle2"/></w:style>
          <w:style w:type="paragraph" w:styleId="Boucle2"><w:name w:val="b2"/><w:basedOn w:val="Boucle1"/><w:rPr><w:i/></w:rPr></w:style>
        </w:styles>
        """;

    const string Numbering = $$"""
        <w:numbering {{NsDecl}}>
          <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/>
            <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:lvl>
            <w:lvl w:ilvl="1"><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
          <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
          <w:num w:numId="2"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:lvl w:ilvl="0"><w:pPr><w:ind w:left="2000" w:hanging="500"/></w:pPr></w:lvl></w:lvlOverride></w:num>
        </w:numbering>
        """;

    const string Theme = """
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"><a:themeElements>
          <a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
            <a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1></a:clrScheme>
          <a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
            <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>
        </a:themeElements></a:theme>
        """;

    static StyleResolver Resolver() => new(new StyleSheet(XElement.Parse(Styles), XElement.Parse(Numbering), XElement.Parse(Theme)));

    /// <summary>Paragraphes du corps d'un document écrit à la main.</summary>
    static List<Paragraph> Doc(params string[] paragraphs)
    {
        var xml = XDocument.Parse($"<w:document {NsDecl}><w:body>{string.Concat(paragraphs)}</w:body></w:document>");
        return StoryMapping.Read(xml, "word/document.xml", "document").Stories().Single().Paragraphs().ToList();
    }

    [Fact]
    public void Sans_style_le_paragraphe_prend_Normal_et_les_valeurs_par_defaut()
    {
        var p = Doc("<w:p><w:r><w:t>texte</w:t></w:r></w:p>")[0];
        var r = Resolver();
        var pf = r.Paragraph(p);
        Assert.Equal("Normal", pf.StyleId);
        Assert.Equal((0, 160, 259, "auto"), (pf.SpacingBefore, pf.SpacingAfter, pf.Line, pf.LineRule));
        var rf = r.Run(p, p.Runs().Single());
        Assert.Equal(("Calibri", 11.0, "fr-BE"), (rf.FontAscii, rf.Size, rf.Language));     // police de thème mineure
        Assert.False(rf.Bold);
        Assert.Null(rf.Color);
    }

    [Fact]
    public void Chaine_basedOn_et_fusion_par_attribut()
    {
        var p = Doc("<w:p><w:pPr><w:pStyle w:val=\"Titre1Rouge\"/><w:spacing w:after=\"40\"/></w:pPr><w:r><w:t>T</w:t></w:r></w:p>")[0];
        var r = Resolver();
        var pf = r.Paragraph(p);
        Assert.True(pf.KeepNext);                                    // hérité de Titre1
        Assert.Equal(0, pf.OutlineLevel);
        Assert.Equal(240, pf.SpacingBefore);                         // w:spacing fusionné par attribut : before de Titre1…
        Assert.Equal(40, pf.SpacingAfter);                           // …after de la mise en forme directe
        Assert.Equal(259, pf.Line);                                  // …line des valeurs par défaut
        Assert.Equal(360, pf.IndentLeft);                            // w:ind : left hérité,
        Assert.Equal(200, pf.IndentFirstLine);                       // firstLine remplace le hanging hérité
        // taquets : 1000 retiré par w:val="clear", 3000 hérité de Normal
        Assert.Equal(new[] { "3000" }, pf.Effective.Element(pf.Effective.Name.Namespace + "tabs")!.Elements().Select(t => t.Attribute(t.Name.Namespace + "pos")!.Value));
        var rf = r.Run(p, p.Runs().Single());
        Assert.Equal(("Calibri Light", 16.0, true, "FF0000"), (rf.FontAscii, rf.Size, rf.Bold, rf.Color));
    }

    [Fact]
    public void Proprietes_bascule_entre_styles_et_mise_en_forme_directe()
    {
        var ps = Doc(
            "<w:p><w:pPr><w:pStyle w:val=\"Titre1\"/></w:pPr><w:r><w:rPr><w:rStyle w:val=\"Accentuation\"/></w:rPr><w:t>a</w:t></w:r></w:p>",
            "<w:p><w:pPr><w:pStyle w:val=\"Titre1\"/></w:pPr><w:r><w:rPr><w:rStyle w:val=\"Accentuation\"/><w:b/></w:rPr><w:t>b</w:t></w:r></w:p>",
            "<w:p><w:pPr><w:pStyle w:val=\"Titre1\"/></w:pPr><w:r><w:rPr><w:b w:val=\"0\"/></w:rPr><w:t>c</w:t></w:r></w:p>");
        var r = Resolver();
        var a = r.Run(ps[0], ps[0].Runs().Single());
        Assert.False(a.Bold);                                        // gras (style de paragraphe) XOR gras (style de caractère)
        Assert.True(a.Italic);                                       // italique : un seul niveau
        Assert.True(r.Run(ps[1], ps[1].Runs().Single()).Bold);       // la mise en forme directe impose
        Assert.False(r.Run(ps[2], ps[2].Runs().Single()).Bold);
    }

    [Fact]
    public void Couleurs_et_polices_du_theme()
    {
        var p = Doc("<w:p><w:pPr><w:pStyle w:val=\"Titre1\"/></w:pPr><w:r><w:t>T</w:t></w:r>" +
                    "<w:r><w:rPr><w:rFonts w:ascii=\"Arial\"/><w:color w:val=\"000000\" w:themeColor=\"text2\"/></w:rPr><w:t>x</w:t></w:r></w:p>")[0];
        var r = Resolver();
        // accent1 4472C4 assombri à 75 % (w:themeShade="BF") en luminance : 2F5496, la valeur que Word enregistre dans w:val
        Assert.Equal("2F5496", r.Run(p, p.Runs().First()).Color);
        var x = r.Run(p, p.Runs().Last());
        Assert.Equal("44546A", x.Color);                             // text2 → dark2 (correspondance par défaut)
        Assert.Equal("Arial", x.FontAscii);                          // police explicite : remplace la police de thème héritée
        Assert.Equal("Calibri Light", x.FontHAnsi);                  // l'autre emplacement garde le thème
    }

    [Fact]
    public void Numerotation_et_style_de_liste()
    {
        var ps = Doc(
            "<w:p><w:pPr><w:pStyle w:val=\"ListePuces\"/></w:pPr><w:r><w:t>un</w:t></w:r></w:p>",
            "<w:p><w:pPr><w:pStyle w:val=\"ListePuces\"/><w:numPr><w:ilvl w:val=\"1\"/></w:numPr></w:pPr><w:r><w:t>deux</w:t></w:r></w:p>",
            "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"2\"/></w:numPr><w:ind w:left=\"3000\"/></w:pPr><w:r><w:t>trois</w:t></w:r></w:p>");
        var r = Resolver();
        var one = r.Paragraph(ps[0]);
        Assert.Equal((1, 0, 720, -360, 100), (one.NumId, one.ListLevel, one.IndentLeft, one.IndentFirstLine, one.IndentRight));
        var two = r.Paragraph(ps[1]);                                // numId du style, ilvl direct (w:numPr fusionné par enfant)
        Assert.Equal((1, 1, 1440), (two.NumId, two.ListLevel, two.IndentLeft));
        var three = r.Paragraph(ps[2]);                              // remplacement de niveau, puis retrait direct
        Assert.Equal((3000, -500), (three.IndentLeft, three.IndentFirstLine));
        Assert.Equal("Symbol", r.ListMarker(ps[0])!.FontAscii);
        Assert.Null(r.ListMarker(Doc("<w:p/>")[0]));
    }

    [Fact]
    public void Une_valeur_directe_nulle_remplace_le_retrait_du_style()
    {
        // calibration FrameView : style « Body Text » w:firstLine="720", mise en forme directe w:firstLine="0" → Word : 0
        var p = Doc("<w:p><w:pPr><w:pStyle w:val=\"Titre1Rouge\"/><w:ind w:firstLine=\"0\"/></w:pPr></w:p>")[0];
        Assert.Equal(0, Resolver().Paragraph(p).IndentFirstLine);
    }

    [Fact]
    public void Le_niveau_de_liste_passe_apres_le_style_de_paragraphe_comme_dans_Word()
    {
        var ps = Doc(
            // Annexe : style à retrait de 1re ligne positif + liste directe à retrait négatif → la liste l'emporte
            "<w:p><w:pPr><w:pStyle w:val=\"Titre1Rouge\"/><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr></w:p>",
            // FrameView, Syllabus : liste du style retirée par w:numId="0" → ni le retrait de la liste ni celui du style
            "<w:p><w:pPr><w:pStyle w:val=\"ListePuces\"/><w:numPr><w:numId w:val=\"0\"/></w:numPr></w:pPr></w:p>");
        var r = Resolver();
        var a = r.Paragraph(ps[0]);
        Assert.Equal((720, -360, 1), (a.IndentLeft, a.IndentFirstLine, a.NumId));
        var b = r.Paragraph(ps[1]);
        Assert.Equal((0, 0, 0, 0), (b.IndentLeft, b.IndentFirstLine, b.IndentRight, b.NumId));    // w:right=100 du style ignoré aussi
    }

    [Fact]
    public void Mesures_avec_unite_du_format_strict()
    {
        // strict-format.docx : w:after="8pt" w:line="12.95pt" (règle auto) → Word : 8 pt après, interligne 1,08 (259)
        var p = Doc("<w:p><w:pPr><w:spacing w:before=\"0.5in\" w:after=\"8pt\" w:line=\"12.95pt\"/><w:ind w:left=\"1cm\"/></w:pPr>" +
                    "<w:r><w:rPr><w:sz w:val=\"11pt\"/></w:rPr><w:t>x</w:t></w:r></w:p>")[0];
        var r = Resolver();
        var pf = r.Paragraph(p);
        Assert.Equal((720, 160, 259, 567), (pf.SpacingBefore, pf.SpacingAfter, pf.Line, pf.IndentLeft));
        Assert.Equal(11.0, r.Run(p, p.Runs().Single()).Size);
    }

    [Fact]
    public void Style_de_tableau_et_boucle_basedOn()
    {
        var p = Doc("<w:p><w:r><w:t>cellule</w:t></w:r></w:p>")[0];
        var r = Resolver();
        Assert.Equal(9.0, r.Run(p, p.Runs().Single(), "Grille").Size);
        Assert.Equal(0, r.Paragraph(p, "Grille").SpacingAfter);
        Assert.Equal(160, r.Paragraph(p, "").SpacingAfter);          // tableau sans style : style par défaut (vide)
        var loop = Doc("<w:p><w:pPr><w:pStyle w:val=\"Boucle1\"/></w:pPr><w:r><w:t>b</w:t></w:r></w:p>")[0];
        Assert.True(r.Run(loop, loop.Runs().Single()).Italic);       // la boucle Boucle1 ↔ Boucle2 est coupée, sans plantage
        var unknown = Doc("<w:p><w:pPr><w:pStyle w:val=\"Inconnu\"/></w:pPr></w:p>")[0];
        Assert.Equal("Normal", r.Paragraph(unknown).StyleId);        // style absent : Normal, comme Word
    }
}
