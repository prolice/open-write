using System.Xml.Linq;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Interop;
using OpenWrite.Core.Model;

namespace OpenWrite.Tests;

/// <summary>Collage de tableaux HTML (forme du presse-papiers d'Excel ; contenu neutre).</summary>
public class HtmlTableTests
{
    static readonly XNamespace W = PropertyBag.TransitionalNs;
    static int _id;
    static Func<string> Ids => () => (++_id).ToString("X8");

    // même structure qu'Excel : CSS dans <style>, <table> AVANT StartFragment, <col> en points, classes xlNN
    const string Excel = """
        Version:1.0
        StartHTML:0000000105
        <html><head><style>
        <!--table {mso-displayed-decimal-separator:"\,";}
        td {padding-top:1px; padding-right:1px; padding-left:1px; mso-ignore:padding; color:windowtext; font-size:10.0pt;
            font-weight:400; font-family:Arial; text-align:general; vertical-align:bottom; border:none; white-space:nowrap;}
        .xl65 {text-align:center; vertical-align:middle; border-top:.5pt solid windowtext; border-right:none;
            border-bottom:.5pt solid windowtext; border-left:1.0pt solid windowtext;}
        .xl66 {color:#00B050; font-size:12.0pt; font-weight:700; border:1.0pt solid windowtext; white-space:normal; background:#FFFF00; mso-pattern:black none;}
        .font5 {color:red; font-weight:700;}
        -->
        </style></head><body>
        <table border=0 cellpadding=0 cellspacing=0 width=200 style='border-collapse:collapse;width:150pt'>
        <!--StartFragment-->
         <col width=40 style='width:30pt'>
         <col width=80 span=2 style='width:60pt'>
         <tr height=37 style='height:28.35pt'>
          <td height=37 class=xl65 style='height:28.35pt'>Un</td>
          <td class=xl66 style='border-left:none'>Deux</td>
          <td>a<font class="font5">b</font></td>
         </tr>
         <tr>
          <td colspan=2 class=xl65>Fusion</td>
          <td>&nbsp;</td>
         </tr>
        <!--EndFragment-->
        </table></body></html>
        """;

    static Table Only(Fragment? f)
    {
        Assert.NotNull(f);
        return Assert.IsType<Table>(Assert.Single(f!.Blocks));
    }

    static List<TableCell> Cells(TableRow r) => r.Items.OfType<TableCell>().ToList();
    static string? Border(TableCell c, string side, string attr = "val") => c.Props.Get("tcBorders")?.Element(W + side)?.Attribute(W + attr)?.Value;

    [Fact]
    public void Tableau_Excel_avant_le_fragment()
    {
        var t = Only(HtmlTableReader.Read(Excel, W, 450, Ids));
        // largeurs des colonnes d'Excel en twips, tblW fixe, marges de la table à 0 (cellpadding=0)
        Assert.Equal(new[] { "600", "1200", "1200" }, t.Grid!.Elements(W + "gridCol").Select(e => e.Attribute(W + "w")!.Value));
        Assert.Equal("3000", t.Props.Get("tblW")!.Attribute(W + "w")!.Value);
        Assert.Equal("0", t.Props.Get("tblCellMar")!.Element(W + "left")!.Attribute(W + "w")!.Value);
        Assert.Null(t.Props.Get("tblBorders"));

        var rows = t.Items.OfType<TableRow>().ToList();
        Assert.Equal(2, rows.Count);
        Assert.Equal("567", rows[0].Props.Get("trHeight")!.Attribute(W + "val")!.Value);

        var a = Cells(rows[0]);
        // bordures par côté comme Word : 1 pt → 8, 0,5 pt → 4, none → nil, windowtext → auto
        Assert.Equal("8", Border(a[0], "left", "sz"));
        Assert.Equal("4", Border(a[0], "top", "sz"));
        Assert.Equal("auto", Border(a[0], "top", "color"));
        Assert.Equal("nil", Border(a[0], "right"));
        Assert.Equal("center", a[0].Props.Val("vAlign"));
        Assert.Equal("center", a[0].Content.Paragraphs().First().Props.Val("jc"));
        Assert.NotNull(a[0].Props.Get("noWrap"));
        Assert.Equal("15", a[0].Props.Get("tcMar")!.Element(W + "top")!.Attribute(W + "w")!.Value);
        Assert.Equal("0", a[0].Props.Get("tcMar")!.Element(W + "bottom")!.Attribute(W + "w")!.Value);
        // style en ligne après la classe : border-left:none l'emporte
        Assert.Equal("nil", Border(a[1], "left"));
        Assert.Equal("8", Border(a[1], "right", "sz"));
        Assert.Null(a[1].Props.Get("noWrap"));
        Assert.Equal("FFFF00", a[1].Props.Get("shd")!.Attribute(W + "fill")!.Value);
        Assert.Equal("bottom", a[2].Props.Val("vAlign"));

        // caractères : Arial 10 pt de la règle td, vert gras 12 pt de la classe, classe d'une balise interne (font5)
        var run = a[0].Content.Paragraphs().First().Inlines.OfType<Run>().Single();
        Assert.Equal("Arial", run.Props.Get("rFonts")!.Attribute(W + "ascii")!.Value);
        Assert.Equal("20", run.Props.Val("sz"));
        var green = a[1].Content.Paragraphs().First().Inlines.OfType<Run>().Single();
        Assert.Equal("00B050", green.Props.Val("color"));
        Assert.Equal("24", green.Props.Val("sz"));
        Assert.NotNull(green.Props.Get("b"));
        var rich = a[2].Content.Paragraphs().First().Inlines.OfType<Run>().ToList();
        Assert.Equal(2, rich.Count);
        Assert.Equal("FF0000", rich[1].Props.Val("color"));
        // marque de paragraphe : mise en forme de la cellule (la frappe la garde)
        Assert.Equal("20", a[0].Content.Paragraphs().First().Props.Get("rPr")!.Element(W + "sz")!.Attribute(W + "val")!.Value);

        // colspan → gridSpan, largeur cumulée
        var b = Cells(rows[1]);
        Assert.Equal(2, b.Count);
        Assert.Equal("2", b[0].Props.Val("gridSpan"));
        Assert.Equal("1800", b[0].Props.Get("tcW")!.Attribute(W + "w")!.Value);
        Assert.Equal("Fusion", TableOps.CellText(b[0]));
    }

    [Fact]
    public void Rowspan_et_texte_autour()
    {
        const string html = """
            <html><body><!--StartFragment--><p>Avant</p>
            <table border=1><tr><th>A</th><th>B</th></tr>
            <tr><td rowspan=2>x</td><td>1</td></tr><tr><td>2</td></tr></table><p>Après</p><!--EndFragment--></body></html>
            """;
        var f = HtmlTableReader.Read(html, W, 400, Ids);
        Assert.NotNull(f);
        Assert.Equal(3, f!.Blocks.Count);
        Assert.Equal("Avant", ((Paragraph)f.Blocks[0]).Text);
        Assert.Equal("Après", ((Paragraph)f.Blocks[2]).Text);
        var t = (Table)f.Blocks[1];
        // sans largeurs : largeur du texte partagée, tblW automatique, grille simple (border=1)
        Assert.Equal("auto", t.Props.Get("tblW")!.Attribute(W + "type")!.Value);
        Assert.NotNull(t.Props.Get("tblBorders"));
        var rows = t.Items.OfType<TableRow>().ToList();
        Assert.True(rows[0].Props.Flag("tblHeader"));
        Assert.NotNull(Cells(rows[0])[0].Content.Paragraphs().First().Inlines.OfType<Run>().First().Props.Get("b"));
        Assert.Equal("restart", Cells(rows[1])[0].Props.Get("vMerge")!.Attribute(W + "val")!.Value);
        var last = Cells(rows[2]);
        Assert.Equal(2, last.Count);
        Assert.NotNull(last[0].Props.Get("vMerge"));
        Assert.Null(last[0].Props.Get("vMerge")!.Attribute(W + "val"));
        Assert.Equal("2", TableOps.CellText(last[1]));
        var map = new TableOps.Map(t);
        Assert.Equal(2, map.MergeBottom(1, 0));
    }

    [Fact]
    public void Sans_tableau_null()
    {
        Assert.Null(HtmlTableReader.Read("<html><body><!--StartFragment--><p>texte</p><!--EndFragment--></body></html>", W, 400, Ids));
        // fragment dans une cellule d'un tableau de mise en page : pas de tableau collé
        Assert.Null(HtmlTableReader.Read("<table><tr><td><!--StartFragment-->mot<!--EndFragment--></td></tr></table>", W, 400, Ids));
    }
}
