using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OpenWrite.Core.Drawing;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;
using OpenWrite.Docx;
using OpenWrite.Layout;

namespace OpenWrite.Tests;

/// <summary>Lot 7 : dessins (lecture DrawingML et VML, géométrie des formes, modifications, habillage du texte).</summary>
public class DrawingTests
{
    const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const string NsDecl = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
        + "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" "
        + "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
        + "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" "
        + "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" "
        + "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" "
        + "xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" "
        + "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" "
        + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";

    // zone de texte flottante (mc:AlternateContent : DrawingML + VML de repli), comme Word l'écrit
    const string TextBox = """
        <mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="5" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
        <wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>127000</wp:posOffset></wp:positionV>
        <wp:extent cx="1270000" cy="635000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="3" name="Zone de texte 3"/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/>
        <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="635000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>
        <wps:txbx><w:txbxContent><w:p><w:r><w:t>Dans la zone</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>
        <mc:Fallback><w:pict><v:shape id="z" style="position:absolute;margin-left:0;margin-top:10pt;width:100pt;height:50pt;z-index:5" fillcolor="white"><v:textbox><w:txbxContent><w:p><w:r><w:t>Dans la zone</w:t></w:r></w:p></w:txbxContent></v:textbox><w10:wrap type="square"/></v:shape></w:pict></mc:Fallback></mc:AlternateContent>
        """;

    static string Doc(string body) => $"<w:document {NsDecl}><w:body>{body}<w:sectPr><w:pgSz w:w=\"12000\" w:h=\"16000\"/><w:pgMar w:top=\"1000\" w:bottom=\"1000\" w:left=\"1000\" w:right=\"1000\"/></w:sectPr></w:body></w:document>";

    static Story Body(string body) =>
        StoryMapping.Read(XDocument.Parse(Doc(body)), "word/document.xml", "document").Stories().Single();

    static DrawingContent First(Story s) => DrawingOps.AllDrawings(s).First();

    [Fact]
    public void Formes_predefinies_toutes_evaluees()
    {
        int n = 0;
        foreach (var name in ShapeGeometry.PresetNames)
        {
            var paths = ShapeGeometry.PresetPaths(name, 120, 80);
            Assert.NotEmpty(paths);
            foreach (var p in paths)
                foreach (var seg in p.Segments)
                    Assert.All(seg.Points, v => Assert.True(double.IsFinite(v) && Math.Abs(v) < 1e4, $"{name} : {v}"));
            n++;
        }
        Assert.True(n >= 180, $"{n} formes");
        // rectangle : contour fermé passant par les coins ; ellipse : quatre quarts de Bézier
        var rect = ShapeGeometry.PresetPaths("rect", 100, 50).Single().Segments;
        Assert.Contains(rect, s => s.Op == 'L' && s.Points[0] == 100 && s.Points[1] == 50);
        Assert.Equal(4, ShapeGeometry.PresetPaths("ellipse", 100, 50).Single().Segments.Count(s => s.Op == 'C'));
    }

    [Fact]
    public void Zone_de_texte_lue_et_reecrite_a_l_identique()
    {
        var xml = Doc($"<w:p><w:r>{TextBox}</w:r><w:r><w:t>Texte</w:t></w:r></w:p>");
        var part = StoryMapping.Read(XDocument.Parse(xml), "word/document.xml", "document");
        var d = First(part.Stories().Single());
        // deux histoires : DrawingML puis VML ; leur texte est dans le modèle, plus dans le XML du dessin
        Assert.Equal(2, d.TextBoxes.Count);
        Assert.Equal("Dans la zone", d.TextBoxes[0].Paragraphs().Single().Text);
        Assert.DoesNotContain(d.Xml.Descendants(), e => e.Name.LocalName == "t");
        Assert.Empty(XmlCompare.Diff(XDocument.Parse(xml).Root!, StoryMapping.Write(part).Root!));

        var o = DrawingReader.Read(d)!;
        Assert.False(o.Inline);
        Assert.Equal(WrapKind.Square, o.Wrap);
        Assert.Equal(("margin", "right"), (o.H.RelativeFrom, o.H.Align));
        Assert.Equal(10, o.V.Offset, 3);
        Assert.Equal((100, 50), (o.Width, o.Height));
        var shape = Assert.IsType<ShapeVisual>(o.Root);
        Assert.Equal(0, shape.TextBox);
        Assert.Equal("FFFFFF", shape.Fill!.Color);
        Assert.Equal(0.75, shape.Line!.Width, 3);
    }

    [Fact]
    public void Vml_trait_et_groupe_en_ligne()
    {
        var body = Body("""
            <w:p><w:r><w:pict><v:group style="width:200pt;height:2pt;mso-position-horizontal-relative:char" coordorigin="0,0" coordsize="2000,20">
            <v:line style="position:absolute" from="0,10" to="2000,10" strokeweight="1pt" strokecolor="#ff0000"><v:stroke dashstyle="dot"/></v:line></v:group></w:pict></w:r></w:p>
            <w:p><w:r><w:pict><v:line style="position:absolute;mso-position-horizontal-relative:page;mso-position-vertical-relative:paragraph;z-index:-3" from="10pt,5pt" to="110pt,5pt"><w10:wrap type="none"/></v:line></w:pict></w:r></w:p>
            """);
        var all = DrawingOps.AllDrawings(body).Select(d => DrawingReader.Read(d)!).ToList();
        Assert.True(all[0].Inline && all[0].Vml);
        var g = Assert.IsType<GroupVisual>(all[0].Root);
        var line = Assert.IsType<ShapeVisual>(Assert.Single(g.Children));
        Assert.Equal(200, line.W, 3);
        Assert.Equal(1, line.Y, 3);                        // 10 / 20 de 2 pt
        Assert.Equal(("FF0000", "dot"), (line.Line!.Color, line.Line.Dash));
        // trait flottant derrière le texte (z-index négatif), repères page / paragraphe
        Assert.False(all[1].Inline);
        Assert.True(all[1].BehindText);
        Assert.Equal(("page", 10.0, "paragraph", 5.0), (all[1].H.RelativeFrom, all[1].H.Offset, all[1].V.RelativeFrom, all[1].V.Offset));
        Assert.Equal(100, all[1].Width, 3);
    }

    [Fact]
    public void Segments_libres_autour_d_un_objet()
    {
        var o = new DrawingObject
        {
            Source = new DrawingContent(new XElement("x")), Root = new ShapeVisual(), Width = 100, Height = 50,
            Wrap = WrapKind.Square, Distance = (0, 0, 9, 9),
        };
        var box = new FloatBox { Object = o, X = 200, Y = 100, W = 100, H = 50 };
        // ligne qui touche l'objet : deux segments ; au-dessus : toute la colonne
        var (segs, _) = DocumentLayout.FreeSegments(new[] { box }, 50, 550, 120, 132);
        Assert.Equal(new[] { (50.0, 191.0), (309.0, 550.0) }, segs);
        Assert.Single(DocumentLayout.FreeSegments(new[] { box }, 50, 550, 80, 92).Segments);
        // texte d'un seul côté, haut et bas
        var right = new FloatBox { Object = new DrawingObject { Source = o.Source, Root = o.Root, Wrap = WrapKind.Square, WrapSide = "right" }, X = 200, Y = 100, W = 100, H = 50 };
        Assert.Equal(new[] { (300.0, 550.0) }, DocumentLayout.FreeSegments(new[] { right }, 50, 550, 120, 132).Segments);
        var tb = new FloatBox { Object = new DrawingObject { Source = o.Source, Root = o.Root, Wrap = WrapKind.TopAndBottom }, X = 200, Y = 100, W = 100, H = 50 };
        var (none, down) = DocumentLayout.FreeSegments(new[] { tb }, 50, 550, 120, 132);
        Assert.Empty(none);
        Assert.Equal(30, down, 3);
    }

    [Fact]
    public void Texte_habille_autour_d_une_zone_de_texte()
    {
        var words = string.Join(" ", Enumerable.Repeat("mot", 300));
        var body = Body($"<w:p><w:r>{TextBox}</w:r><w:r><w:t>{words}</w:t></w:r></w:p>");
        var sheet = new StyleSheet(null);
        var layout = new DocumentLayout(body, new StyleResolver(sheet), new MonoFontProvider(ascent: 0.8, descent: 0.2));
        layout.LayoutAll();
        var page = layout.Pages[0];
        var f = Assert.Single(page.Floats);
        // zone alignée à droite de la marge (marge gauche 50 pt, texte 500 pt) ; 10 pt sous le haut du paragraphe
        Assert.Equal(450, f.X, 1);
        Assert.Equal(50 + 10, f.Y, 1);
        Assert.True(f.TextLines.Any(l => l.Line.Paragraph.Content.Text == "Dans la zone"), $"zone : {f.TextLines.Count} lignes « {string.Join("|", f.TextLines.Select(l => l.Line.Paragraph.Content.Text))} »");
        // lignes à côté de la zone : arrêtées avant son bord (moins la distance de 9 pt) ; lignes au-dessus : pleine largeur
        foreach (var l in page.Lines)
        {
            var lb = l.Paragraph.Lines[l.Line];
            double right = l.X + (lb.Pieces.Count > 0 ? lb.Pieces[^1].X + lb.Pieces[^1].Width : 0);
            if (l.Y + lb.Height > f.Y && l.Y < f.Y + f.H) Assert.True(right <= 441 + 0.01, $"ligne à {l.Y} finit à {right}");
        }
        // sous la zone : de nouveau toute la largeur
        Assert.True(page.Lines.Any(l => l.Y > f.Y + f.H && l.Paragraph.Lines[l.Line].Pieces.Sum(p => p.Width) > 420),
            string.Join(" ; ", page.Lines.Select(l => $"{l.Y:0.#}:{l.Paragraph.Lines[l.Line].Pieces.Sum(p => p.Width):0.#}")));
    }

    [Fact]
    public void Objet_ancre_dans_une_cellule_place_par_rapport_a_elle()
    {
        var body = Body($"<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"3000\"/><w:gridCol w:w=\"3000\"/></w:tblGrid>"
            + $"<w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r>{TextBox.Replace("<wp:align>right</wp:align>", "<wp:posOffset>0</wp:posOffset>").Replace("relativeFrom=\"margin\"", "relativeFrom=\"column\"")}</w:r><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>");
        var layout = new DocumentLayout(body, new StyleResolver(new StyleSheet(null)), new MonoFontProvider(ascent: 0.8, descent: 0.2));
        layout.LayoutAll();
        var f = Assert.Single(layout.Pages[0].Floats);
        var cellLine = layout.Pages[0].Lines.First(l => l.Paragraph.Content.Text == "b");
        // repère « colonne » d'un objet de cellule : le texte de la cellule (seconde colonne, marge de cellule comprise)
        Assert.Equal(cellLine.X, f.X, 1);
        Assert.True(f.X > 50 + 150 - 1, $"x = {f.X}");
        Assert.Equal(cellLine.Y + 10, f.Y, 1);
    }

    [Fact]
    public void Modifications_et_image_ajoutee_valides_pour_le_SDK()
    {
        var src = RoundTripTests.Docx($"<w:document {NsDecl}><w:body><w:p><w:r><w:t>Avant</w:t></w:r></w:p><w:p/></w:body></w:document>");
        var pkg = DocxPackage.Open(new MemoryStream(src));
        var body = pkg.Main!.Stories().Single();
        DocxPackage.EnsureDrawingNamespaces(pkg.Main);
        // image en ligne (PNG d'un pixel), puis flottante carrée, déplacée et agrandie
        var png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");
        var rid = pkg.AddImage(pkg.Main.PartName, png, "png");
        var w = (XNamespace)W;
        var d = new DrawingContent(DrawingOps.CreatePicture(w, 1, "Image 1", rid, 72, 36, "un point"));
        var run = new Run();
        run.Content.Add(d);
        ((Paragraph)body.Blocks[1]).Inlines.Add(run);
        var inline = DrawingReader.Read(d)!;
        Assert.True(inline.Inline);
        Assert.Equal(rid, Assert.IsType<PictureVisual>(inline.Root).ImageId);
        DrawingOps.ToAnchor(d, 20, 5, WrapKind.Square, false, 251659264);
        DrawingOps.SetSize(d, 144, 72);
        DrawingOps.SetPosition(d, "margin", 30, "paragraph", 12);
        DrawingOps.SetRotation(d, 30);
        DrawingOps.SetCrop(d, 0.1, 0, 0.1, 0);
        var o = DrawingReader.Read(d)!;
        Assert.False(o.Inline);
        Assert.Equal((WrapKind.Square, 144.0, 72.0), (o.Wrap, o.Width, o.Height));
        Assert.Equal(("margin", 30.0), (o.H.RelativeFrom, o.H.Offset));
        var pic = Assert.IsType<PictureVisual>(o.Root);
        Assert.Equal(30, pic.Rotation, 3);
        Assert.Equal(0.1, pic.Crop.L, 3);
        // forme et zone de texte créées comme Word
        var shape = new DrawingContent(DrawingOps.CreateShape(w, 2, "ellipse", 10, 10, 100, 60, 251659265));
        var box = new DrawingContent(DrawingOps.CreateShape(w, 3, "rect", 200, 10, 120, 50, 251659266, textBox: true));
        box.TextBoxes.Add(new Story { Blocks = { new Paragraph { Inlines = { new Run { Content = { new TextContent { Value = "Écrit" } } } } } } });
        foreach (var x in new[] { shape, box })
        {
            var r = new Run();
            r.Content.Add(x);
            ((Paragraph)body.Blocks[0]).Inlines.Add(r);
        }
        Assert.Equal("ellipse", Assert.IsType<ShapeVisual>(DrawingReader.Read(shape)!.Root).Preset);
        DrawingOps.ToInline(d);
        Assert.True(DrawingReader.Read(d)!.Inline);
        pkg.Main.Dirty = true;
        var ms = new MemoryStream();
        pkg.Save(ms);

        // le fichier est valide pour le SDK Open XML, l'image est dans le paquet, la zone de texte a son texte
        ms.Position = 0;
        using var doc = WordprocessingDocument.Open(ms, false);
        var errors = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(doc).Select(e => e.Description).ToList();
        Assert.Empty(errors);
        Assert.Single(doc.MainDocumentPart!.ImageParts);
        Assert.Contains("Écrit", doc.MainDocumentPart.Document.InnerText);
    }
}
