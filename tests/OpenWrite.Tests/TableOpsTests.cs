using System.Xml.Linq;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;
using OpenWrite.Docx;

namespace OpenWrite.Tests;

/// <summary>Lot 6 : opérations sur les tableaux (modèle), écrites comme Word.</summary>
public class TableOpsTests
{
    static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const string NsDecl = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";

    static int _id;
    static string NewId() => (++_id).ToString("X8");

    static Story Body(params string[] blocks)
    {
        var xml = XDocument.Parse($"<w:document {NsDecl}><w:body>{string.Concat(blocks)}<w:sectPr/></w:body></w:document>");
        return StoryMapping.Read(xml, "word/document.xml", "document").Stories().Single();
    }

    static string P(string text) => $"<w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";

    /// <summary>Tableau de rows × cols cellules « r,c », colonnes de 1000 twips.</summary>
    static Table Grid(int rows, int cols)
    {
        var t = TableOps.Create(W, rows, cols, cols * 50 + 0.5, null, null, NewId);
        var map = new TableOps.Map(t);
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
            {
                var p = (Paragraph)map.Cells[r][c].Cell.Content.Blocks[0];
                var run = new Run();
                run.Content.Add(new TextContent { Value = $"{r},{c}" });
                p.Inlines.Add(run);
            }
        return t;
    }

    static string[][] Texts(Table t) => new TableOps.Map(t).Cells.Select(row => row.Select(s => TableOps.CellText(s.Cell)).ToArray()).ToArray();
    static int[] Spans(Table t, int row) => new TableOps.Map(t).Cells[row].Select(s => s.Span).ToArray();
    static int[] GridOf(Table t) => new TableOps.Map(t).GridWidths().ToArray();

    [Fact]
    public void Creation_comme_Word()
    {
        var t = TableOps.Create(W, 2, 3, 451.3, "TableGrid", null, NewId);
        Assert.Equal(new[] { 3005, 3005, 3006 }, GridOf(t));
        Assert.Equal("TableGrid", t.Props.Val("tblStyle"));
        Assert.Equal("04A0", t.Props.Get("tblLook")!.Attribute(W + "val")!.Value);
        Assert.Equal("3005", new TableOps.Map(t).Cells[0][0].Cell.Props.Get("tcW")!.Attribute(W + "w")!.Value);
        Assert.Equal(TableLook.Default, TableStyles.Look(t.Props.Get("tblLook")));
    }

    [Fact]
    public void Insertion_dans_un_paragraphe_coupe()
    {
        var body = Body(P("avantapres"));
        var pos = TableOps.InsertTable(body, new DocPos(new[] { 0 }, 5), Grid(2, 2), NewId);
        Assert.IsType<Paragraph>(body.Blocks[0]);
        Assert.IsType<Table>(body.Blocks[1]);
        Assert.Equal("avant", ((Paragraph)body.Blocks[0]).Text);
        Assert.Equal("apres", ((Paragraph)body.Blocks[2]).Text);
        Assert.Equal(new[] { 1, 0, 0, 0 }, pos.Path);
        // en tête de paragraphe : avant lui
        var body2 = Body(P("texte"));
        TableOps.InsertTable(body2, new DocPos(new[] { 0 }, 0), Grid(1, 1), NewId);
        Assert.IsType<Table>(body2.Blocks[0]);
        Assert.Equal("texte", ((Paragraph)body2.Blocks[1]).Text);
    }

    [Fact]
    public void Rangees_inserees_et_supprimees_fusions_gardees()
    {
        var t = Grid(3, 2);
        TableOps.Merge(t, 0, 0, 1, 0, NewId);                 // colonne 0 fusionnée sur les rangées 0 et 1
        TableOps.InsertRows(t, 0, below: true, 1, NewId);     // dans la fusion : elle se prolonge
        var map = new TableOps.Map(t);
        Assert.Equal(4, map.Rows.Count);
        Assert.True(map.IsContinue(map.Cells[1][0].Cell));
        Assert.Equal(2, map.MergeBottom(0, 0));
        Assert.Equal("", TableOps.CellText(map.Cells[1][1].Cell));
        // rangée de départ supprimée : la fusion recommence à la rangée suivante
        TableOps.DeleteRows(t, 0, 0);
        map = new TableOps.Map(t);
        Assert.True(map.IsRestart(map.Cells[0][0].Cell));
        Assert.Equal(1, map.MergeBottom(0, 0));
        Assert.True(TableOps.DeleteRows(t, 0, 2));
    }

    [Fact]
    public void Colonnes_inserees_et_supprimees()
    {
        var t = Grid(2, 2);
        TableOps.Merge(t, 0, 0, 0, 1, NewId);                 // première rangée : une cellule sur deux colonnes
        TableOps.InsertColumns(t, 1, 1, NewId);               // au milieu : la cellule fusionnée s'élargit
        Assert.Equal(new[] { 3 }, Spans(t, 0));
        Assert.Equal(new[] { 1, 1, 1 }, Spans(t, 1));
        Assert.Equal(new[] { "1,0", "", "1,1" }, Texts(t)[1]);
        Assert.Equal(3, GridOf(t).Length);
        TableOps.InsertColumns(t, 3, 1, NewId, fitTwips: 3000);   // à la fin, ramené à la largeur disponible
        Assert.InRange(GridOf(t).Sum(), 2996, 3004);
        Assert.False(TableOps.DeleteColumns(t, 1, 2));
        Assert.Equal(new[] { 1, 1 }, Spans(t, 0));             // la cellule fusionnée perd deux de ses trois colonnes
        Assert.Equal(new[] { "1,0", "" }, Texts(t)[1]);
    }

    [Fact]
    public void Fusion_et_fractionnement()
    {
        var t = Grid(3, 3);
        var merged = TableOps.Merge(t, 0, 0, 1, 1, NewId);
        Assert.Equal("0,0 0,1 1,0 1,1", TableOps.CellText(merged));
        var map = new TableOps.Map(t);
        Assert.Equal(new[] { 2, 1 }, Spans(t, 0));
        Assert.True(map.IsRestart(map.Cells[0][0].Cell));
        Assert.True(map.IsContinue(map.Cells[1][0].Cell));
        Assert.Equal("2000", map.Cells[0][0].Cell.Props.Get("tcW")!.Attribute(W + "w")!.Value);

        // fractionner une cellule simple en trois : la grille gagne des colonnes, les autres rangées s'élargissent
        var s = Grid(2, 2);
        TableOps.Split(s, 0, 1, 3, 1, NewId);
        Assert.Equal(new[] { 1, 1, 1, 1 }, Spans(s, 0));
        Assert.Equal(new[] { 1, 3 }, Spans(s, 1));
        Assert.Equal(4, GridOf(s).Length);
        Assert.Equal(new[] { "0,0", "0,1", "", "" }, Texts(s)[0]);
        // fractionner en deux rangées : les autres cellules de la rangée fusionnent verticalement
        var v = Grid(1, 2);
        TableOps.Split(v, 0, 0, 1, 2, NewId);
        var vm = new TableOps.Map(v);
        Assert.Equal(2, vm.Rows.Count);
        Assert.True(vm.IsRestart(vm.Cells[0][1].Cell));
        Assert.True(vm.IsContinue(vm.Cells[1][1].Cell));
        Assert.False(vm.IsContinue(vm.Cells[1][0].Cell));
    }

    [Fact]
    public void Fractionner_le_tableau_et_bordures()
    {
        var body = new Story();
        var t = Grid(3, 2);
        body.Blocks.Add(t);
        body.Blocks.Add(new Paragraph());
        var pos = TableOps.SplitTable(body, new[] { 0 }, 1, NewId);
        Assert.Equal(new[] { 1 }, pos.Path);
        Assert.Single(((Table)body.Blocks[0]).Items.OfType<TableRow>());
        Assert.Equal(2, ((Table)body.Blocks[2]).Items.OfType<TableRow>().Count());

        var b = Grid(2, 2);
        TableOps.SetBorders(b, 0, 0, 1, 1, "outside", new TableOps.Line("single", 8, "FF0000"));
        var map = new TableOps.Map(b);
        var tl = map.Cells[0][0].Cell.Props.Get("tcBorders")!;
        Assert.Equal(new[] { "top", "left" }, tl.Elements().Select(e => e.Name.LocalName).ToArray());
        Assert.Equal("FF0000", tl.Element(W + "top")!.Attribute(W + "color")!.Value);
        TableOps.SetBorders(b, 0, 0, 1, 1, "none", default);
        Assert.Equal("nil", map.Cells[1][1].Cell.Props.Get("tcBorders")!.Element(W + "bottom")!.Attribute(W + "val")!.Value);
    }

    [Fact]
    public void Texte_et_tableau_tri_formule()
    {
        var body = Body(P("nom\tprix"), P("b\t12,5"), P("a\t3"));
        var paras = body.Blocks.OfType<Paragraph>().ToList();
        var t = TableOps.FromParagraphs(paras, '\t', 300, null, NewId);
        Assert.Equal(new[] { new[] { "nom", "prix" }, new[] { "b", "12,5" }, new[] { "a", "3" } }, Texts(t));
        TableOps.Sort(t, 1, ascending: true, headerRows: 1);
        Assert.Equal(new[] { "a", "3" }, Texts(t)[1]);
        Assert.Equal(15.5, TableOps.Compute(t, 3, 1, "=SUM(ABOVE)") ?? double.NaN);   // sous le tableau : rangée 3
        Assert.Null(TableOps.Compute(t, 0, 0, "=FOO(ABOVE)"));
        Assert.Equal("15,5", TableOps.FormatNumber(15.5));
        var back = TableOps.ToText(t, "\t", NewId);
        Assert.Equal(new[] { "nom\tprix", "a\t3", "b\t12,5" }, back.Select(p => p.Text).ToArray());
    }

    [Fact]
    public void Redimensionner_une_colonne_garde_la_largeur()
    {
        var t = Grid(1, 3);
        TableOps.MoveBoundary(t, 1, 200);
        Assert.Equal(new[] { 1200, 800, 1000 }, GridOf(t));
        TableOps.MoveBoundary(t, 3, -300);                    // bord droit : seule la dernière colonne change
        Assert.Equal(new[] { 1200, 800, 700 }, GridOf(t));
        TableOps.DistributeColumns(t, 0, 2);
        Assert.Equal(new[] { 900, 900, 900 }, GridOf(t));
        TableOps.AutoFit(t, "window");
        Assert.Equal("pct", t.Props.Get("tblW")!.Attribute(W + "type")!.Value);
    }

    [Fact]
    public void Style_de_la_galerie_ecrit_comme_Word()
    {
        var root = XElement.Parse($"<w:styles {NsDecl}/>");
        var sheet = new StyleSheet(root);
        Assert.Equal(36, TableStyleOps.Gallery().Count());
        var id = TableStyleOps.Ensure(sheet, root, "Grid Table 4 Accent 1");
        Assert.Equal("GridTable4-Accent1", id);
        Assert.Equal(id, TableStyleOps.Ensure(sheet, root, "grid table 4 accent 1"));         // pas de doublon
        var normal = root.Elements(W + "style").Single(s => s.Attribute(W + "styleId")!.Value == "TableNormal");
        Assert.Equal("1", normal.Attribute(W + "default")!.Value);
        var style = root.Elements(W + "style").Single(s => s.Attribute(W + "styleId")!.Value == id);
        Assert.Equal("TableNormal", style.Element(W + "basedOn")!.Attribute(W + "val")!.Value);
        // sans thème : couleurs d'Office, nuances telles que Word les écrit (4472C4 éclairci à 60 % → 8EAADB)
        Assert.Equal("8EAADB", style.Element(W + "tblPr")!.Element(W + "tblBorders")!.Element(W + "top")!.Attribute(W + "color")!.Value);
        var header = style.Elements(W + "tblStylePr").Single(e => e.Attribute(W + "type")!.Value == "firstRow");
        Assert.Equal("4472C4", header.Element(W + "tcPr")!.Element(W + "shd")!.Attribute(W + "fill")!.Value);
        Assert.Equal("Tableau Grille 4 - Accentuation 1", TableStyleOps.French("Grid Table 4 Accent 1"));
    }
}
