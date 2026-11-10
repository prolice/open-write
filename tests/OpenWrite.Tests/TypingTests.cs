using System.Xml.Linq;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;
using OpenWrite.Docx;
using OpenWrite.Layout;

namespace OpenWrite.Tests;

/// <summary>
/// Lot 3 : point d'insertion et saisie. Correspondance texte affiché ↔ positions du modèle (champs, cases, notes),
/// géométrie du point d'insertion, caractères spéciaux, collage sur plusieurs paragraphes, frappe suivie d'une remise
/// en page incrémentale identique à une mise en page complète. Police à chasse fixe : 5 pt par caractère à 10 pt.
/// </summary>
public class TypingTests
{
    const string NsDecl = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";

    const string Styles = $$"""
        <w:styles {{NsDecl}}>
          <w:docDefaults>
            <w:rPrDefault><w:rPr><w:sz w:val="20"/></w:rPr></w:rPrDefault>
            <w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>
          </w:docDefaults>
          <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
        </w:styles>
        """;

    const string Sect = """<w:sectPr><w:pgSz w:w="7200" w:h="3200"/><w:pgMar w:top="600" w:bottom="600" w:left="600" w:right="600"/></w:sectPr>""";

    static string T(string text, string rPr = "") => $"<w:r>{(rPr.Length > 0 ? $"<w:rPr>{rPr}</w:rPr>" : "")}<w:t xml:space=\"preserve\">{text}</w:t></w:r>";
    static string P(params string[] runs) => $"<w:p>{string.Concat(runs)}</w:p>";

    static (Story Body, DocumentLayout Layout) Doc(params string[] paragraphs)
    {
        var xml = XDocument.Parse($"<w:document {NsDecl}><w:body>{string.Concat(paragraphs)}{Sect}</w:body></w:document>");
        var body = StoryMapping.Read(xml, "word/document.xml", "document").Stories().Single();
        var layout = new DocumentLayout(body, new StyleResolver(new StyleSheet(XElement.Parse(Styles))), new MonoFontProvider(ascent: 0.8, descent: 0.2));
        layout.LayoutAll();
        return (body, layout);
    }

    static ParagraphLayout Layout(DocumentLayout l, Story body, int block) => l.Line(l.LinesOf((Paragraph)body.Blocks[block])[0]).Paragraph;

    static string Text(Paragraph p) => string.Concat(p.Runs().SelectMany(r => r.Content).Select(c => c switch
    {
        TextContent t => t.Value,
        TabContent => "\t",
        BreakContent b => b.Attributes.Any(a => a.Value == "page") ? "\f" : "\n",
        _ => "",
    }));

    static readonly Func<string> Ids = FlowEdits.IdGenerator(Array.Empty<string>());

    [Fact]
    public void Correspondance_texte_affiche_et_modele_avec_champ_masque()
    {
        // « ab » + champ (code masqué, résultat « 12 ») + « cd » : le code n'est pas affiché mais occupe le modèle
        var (body, l) = Doc(P(T("ab"),
            "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> PAGE </w:instrText></w:r>",
            "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>", T("12"), "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>", T("cd")));
        var c = Layout(l, body, 0).Content;
        Assert.Equal("ab12cd", c.Text);
        // modèle : a(0) b(1) [begin](2) " PAGE "(3-8) [separate](9) 1(10) 2(11) [end](12) c(13) d(14)
        Assert.Equal(new[] { 0, 1, 10, 11, 13, 14 }, c.ModelStart);
        Assert.Equal(2, CaretGeometry.ModelOffset(c, 2));          // après « ab », avant le champ
        Assert.Equal(12, CaretGeometry.ModelOffset(c, 4));         // après « 12 », dans le résultat
        Assert.Equal(15, CaretGeometry.ModelOffset(c, 6));
        for (int i = 0; i <= c.Text.Length; i++) Assert.Equal(i, CaretGeometry.DisplayIndex(c, CaretGeometry.ModelOffset(c, i)));
        // la frappe au point d'insertion affiché 2 va après « ab », hors du champ
        var p = (Paragraph)body.Blocks[0];
        TextOps.Insert(body, new DocPos(new[] { 0 }, CaretGeometry.ModelOffset(c, 2)), "X");
        Assert.Equal("abX", p.Runs().First().Content.OfType<TextContent>().Single().Value);
    }

    [Fact]
    public void Case_a_cocher_couvre_tout_le_champ()
    {
        var box = "<w:r><w:fldChar w:fldCharType=\"begin\"><w:ffData><w:checkBox><w:default w:val=\"0\"/></w:checkBox></w:ffData></w:fldChar></w:r>" +
                  "<w:r><w:instrText xml:space=\"preserve\"> FORMCHECKBOX </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r>";
        var (body, l) = Doc(P(T("a"), box, T("b")));
        var c = Layout(l, body, 0).Content;
        Assert.Equal("a☐b", c.Text);
        // la case couvre [begin, end] : 1 à 17 (begin 1 + code 14 + end 1)
        Assert.Equal((1, 17), CaretGeometry.ModelRange(c, 1));
        Assert.Equal(17, CaretGeometry.ModelOffset(c, 2));
        // taper juste après la case n'entre pas dans le code du champ
        TextOps.Insert(body, new DocPos(new[] { 0 }, CaretGeometry.ModelOffset(c, 2)), "X");
        Assert.Equal("aXb", Text((Paragraph)body.Blocks[0]));
        Assert.Equal(" FORMCHECKBOX ", FieldIndex.Build(body).All.Single().Code);
        // Retour arrière sur la case : tout le champ disparaît
        var (b2, l2) = Doc(P(T("a"), box, T("b")));
        var c2 = Layout(l2, b2, 0).Content;
        var (s, e) = CaretGeometry.ModelRange(c2, 1);
        TextOps.Delete(b2, new DocPos(new[] { 0 }, s), new DocPos(new[] { 0 }, e));
        Assert.Empty(FieldIndex.Build(b2).All);
        Assert.Equal("ab", Text((Paragraph)b2.Blocks[0]));
    }

    [Fact]
    public void Abscisses_et_point_le_plus_proche_sur_une_ligne()
    {
        var (body, l) = Doc(P(T("abc\tde")));
        var pl = Layout(l, body, 0);
        Assert.Equal(0, CaretGeometry.X(pl, 0, 0), 3);
        Assert.Equal(15, CaretGeometry.X(pl, 0, 3), 3);            // avant la tabulation : fin de « abc »
        Assert.Equal(36, CaretGeometry.X(pl, 0, 4), 3);            // après : taquet par défaut (36 pt)
        Assert.Equal(46, CaretGeometry.X(pl, 0, 6), 3);
        Assert.Equal(1, CaretGeometry.IndexAtX(pl, 0, 6));
        Assert.Equal(6, CaretGeometry.IndexAtX(pl, 0, 100));
        Assert.Equal(4, CaretGeometry.IndexAtX(pl, 0, 35));
    }

    [Fact]
    public void Ligne_du_point_d_insertion_en_fin_de_ligne_coupee()
    {
        var (body, l) = Doc(P(T(string.Concat(Enumerable.Repeat("aaaa ", 20)).TrimEnd())));
        var pl = Layout(l, body, 0);
        int end0 = pl.Lines[0].End;
        Assert.Equal(1, CaretGeometry.LineOf(pl, end0, atLineEnd: false));
        Assert.Equal(0, CaretGeometry.LineOf(pl, end0, atLineEnd: true));
        Assert.Equal(1, CaretGeometry.LineOf(pl, pl.Content.Text.Length, atLineEnd: false));
    }

    [Fact]
    public void Mots_suivant_precedent_et_double_clic()
    {
        const string s = "Bonjour le monde, l’ami";
        Assert.Equal(8, CaretGeometry.NextWord(s, 0));
        Assert.Equal(16, CaretGeometry.NextWord(s, 11));
        Assert.Equal(11, CaretGeometry.PreviousWord(s, 16));
        Assert.Equal(0, CaretGeometry.PreviousWord(s, 3));
        Assert.Equal((8, 11), CaretGeometry.WordAt(s, 9));
        Assert.Equal((18, 23), CaretGeometry.WordAt(s, 20));       // apostrophe typographique : un seul mot
    }

    [Fact]
    public void Tabulation_et_sauts_inseres_au_milieu_du_texte()
    {
        var xml0 = XDocument.Parse($"<w:document {NsDecl}><w:body>{P(T("abcd", "<w:b/>"))}{Sect}</w:body></w:document>");
        var part = StoryMapping.Read(xml0, "word/document.xml", "document");
        var body = part.Stories().Single();
        var at = new DocPos(new[] { 0 }, 2);
        var after = TextOps.InsertSpecial(body, at, SpecialChar.Tab);
        Assert.Equal(3, after.Offset);
        after = TextOps.InsertSpecial(body, after, SpecialChar.LineBreak);
        TextOps.InsertSpecial(body, new DocPos(new[] { 0 }, 0), SpecialChar.PageBreak);
        var p = (Paragraph)body.Blocks[0];
        Assert.Equal("\fab\t\ncd", Text(p));
        Assert.Single(p.Runs());                                    // tout reste dans le segment gras
        var xml = StoryMapping.Write(part).ToString();
        Assert.Contains("w:type=\"page\"", xml);
    }

    [Fact]
    public void Collage_sur_plusieurs_paragraphes_annulable_en_une_fois()
    {
        var (body, _) = Doc(P(T("Début fin")));
        var journal = new UndoJournal();
        var pos = TextOps.Replace(journal, body, "Coller", new DocPos(new[] { 0 }, 6), new DocPos(new[] { 0 }, 6), "un\r\ndeux\tx\r\ntrois ", Ids);
        Assert.Equal(3, body.Blocks.OfType<Paragraph>().Count());
        Assert.Equal(new[] { "Début un", "deux\tx", "trois fin" }, body.Blocks.OfType<Paragraph>().Select(Text).ToArray());
        Assert.Equal(new DocPos(new[] { 2 }, 6), pos);
        journal.Undo();
        Assert.Equal("Début fin", Text((Paragraph)body.Blocks.Single()));
        // remplacement d'une sélection sur deux paragraphes
        journal.Redo();
        TextOps.Replace(journal, body, "Frappe", new DocPos(new[] { 0 }, 3), new DocPos(new[] { 2 }, 2), "Z", Ids);
        Assert.Equal("DébZois fin", Text((Paragraph)body.Blocks.Single()));
        Assert.Equal((body, 0, 1), journal.LastChange is { } ch ? (ch.Root, ch.Index, ch.Count) : default);
    }

    [Fact]
    public void Entree_en_fin_de_paragraphe_mise_en_forme_continuee_ou_style_suivant()
    {
        // même style : la marque du nouveau paragraphe reprend le gras, la frappe continue en gras
        var (body, _) = Doc(P(T("Gras", "<w:b/>")));
        var pos = TextOps.Split(body, new DocPos(new[] { 0 }, 4), Ids, _ => null);
        TextOps.Insert(body, pos, "suite");
        var run = ((Paragraph)body.Blocks[1]).Runs().Single();
        Assert.True(run.Props.Flag("b"));
        // titre suivi de « Normal » : style suivant, sans la mise en forme du titre
        var (b2, _) = Doc("<w:p><w:pPr><w:pStyle w:val=\"Titre1\"/></w:pPr>" + T("Titre", "<w:b/>") + "</w:p>");
        TextOps.Split(b2, new DocPos(new[] { 0 }, 5), Ids, id => id == "Titre1" ? "Normal" : null);
        var p2 = (Paragraph)b2.Blocks[1];
        Assert.Equal("Normal", p2.Props.Val("pStyle"));
        Assert.Null(p2.Props.Get("rPr"));
        // au milieu : rien ne change
        var (b3, _) = Doc(P(T("abcd", "<w:i/>")));
        TextOps.Split(b3, new DocPos(new[] { 0 }, 2), Ids, _ => "Autre");
        Assert.Null(((Paragraph)b3.Blocks[1]).Props.Val("pStyle"));
    }

    static Run[] Runs(Story body, int block) => ((Paragraph)body.Blocks[block]).Runs().ToArray();

    [Fact]
    public void Gras_sur_une_partie_du_segment_coupe_en_trois()
    {
        var (body, _) = Doc(P(T("abcdef", "<w:i/>")));
        var journal = new UndoJournal();
        var from = new DocPos(new[] { 0 }, 2);
        var to = new DocPos(new[] { 0 }, 4);
        FormatOps.Apply(journal, body, "Gras", from, to, () => FormatOps.ApplyToRuns(body, from, to, (_, r) => FormatOps.SetToggle(r, "b", true)));
        var runs = Runs(body, 0);
        Assert.Equal(new[] { "ab", "cd", "ef" }, runs.Select(r => r.Text).ToArray());
        Assert.Equal(new[] { false, true, false }, runs.Select(r => r.Props.Flag("b")).ToArray());
        Assert.All(runs, r => Assert.True(r.Props.Flag("i")));              // le reste de la mise en forme est gardé
        Assert.True(runs[1].Props.Flag("bCs"));
        journal.Undo();
        Assert.Equal("abcdef", Runs(body, 0).Single().Text);
    }

    [Fact]
    public void Gras_retire_alors_que_le_style_le_donne_pose_a_zero()
    {
        var (body, _) = Doc(P(T("abc", "<w:b/>"), T("def")));
        var from = new DocPos(new[] { 0 }, 0);
        var to = new DocPos(new[] { 0 }, 6);
        // « hérité » simulé : le second segment reste gras par son style
        FormatOps.ApplyToRuns(body, from, to, (_, r) => FormatOps.SetToggle(r, "b", false, run => run.Text == "def"));
        var runs = Runs(body, 0);
        Assert.Null(runs[0].Props.Get("b"));
        Assert.Equal("0", runs[1].Props.Val("b"));
    }

    [Fact]
    public void Mise_en_forme_sur_plusieurs_paragraphes_et_alignement()
    {
        var (body, _) = Doc(P(T("un deux")), P(T("trois")), P(T("quatre cinq")));
        var from = new DocPos(new[] { 0 }, 3);
        var to = new DocPos(new[] { 2 }, 6);
        int n = FormatOps.ApplyToRuns(body, from, to, (_, r) => FormatOps.SetUnderline(r, true));
        Assert.Equal(3, n);
        Assert.Equal(new[] { null, "single" }, Runs(body, 0).Select(r => r.Props.Val("u")).ToArray());
        Assert.Equal("single", Runs(body, 1).Single().Props.Val("u"));
        Assert.Equal(new[] { "single", null }, Runs(body, 2).Select(r => r.Props.Val("u")).ToArray());
        FormatOps.ApplyToParagraphs(body, from, to, p => p.Props.SetVal("jc", "center"));
        Assert.All(body.Blocks.OfType<Paragraph>(), p => Assert.Equal("center", p.Props.Val("jc")));
    }

    [Fact]
    public void Lignes_d_un_paragraphe_et_pas_a_pas_d_une_page_a_l_autre()
    {
        var paras = Enumerable.Range(0, 30).Select(i => P(T(string.Concat(Enumerable.Repeat("aaaa ", 12)).TrimEnd()))).ToArray();
        var (body, l) = Doc(paras);
        Assert.True(l.Pages.Count >= 3);
        var last = (Paragraph)body.Blocks[29];
        var refs = l.LinesOf(last);
        Assert.Single(refs);
        Assert.Equal(l.Pages.Count - 1, refs[0].Page);
        // pas à pas depuis la dernière ligne de la page 0 : première ligne de la page 1
        var end0 = new LineRef(0, l.Pages[0].Lines.Count - 1);
        Assert.Equal(new LineRef(1, 0), l.Step(end0, 1));
        Assert.Equal(end0, l.Step(new LineRef(1, 0), -1));
        Assert.Null(l.Step(new LineRef(0, 0), -1));
    }

    [Fact]
    public void Frappe_puis_remise_en_page_incrementale_identique()
    {
        var paras = Enumerable.Range(0, 200).Select(i => P(T(string.Concat(Enumerable.Repeat("aaaa ", 1 + i % 25)).TrimEnd()))).ToArray();
        var (body, l) = Doc(paras);
        var journal = new UndoJournal();
        var pos = new DocPos(new[] { 100 }, 3);
        // un mot par frappe, puis Entrée, puis Retour arrière au début du nouveau paragraphe
        foreach (var ch in "bbbb cccc dddd eeee ffff gggg hhhh iiii jjjj kkkk llll mmmm nnnn ")
        {
            pos = TextOps.Type(journal, body, pos, ch.ToString());
            Update(l, body, journal);
        }
        pos = TextOps.SplitParagraph(journal, body, pos, Ids);
        Update(l, body, journal);
        Same(l, body);
        TextOps.DeleteRange(journal, body, new DocPos(new[] { 100 }, ParagraphAtoms.Length((Paragraph)body.Blocks[100])), pos);
        Update(l, body, journal);
        Same(l, body);
        while (journal.CanUndo) { journal.Undo(); Update(l, body, journal); }
        Same(l, body);
    }

    static void Update(DocumentLayout l, Story body, UndoJournal journal)
    {
        var ch = journal.LastChange!.Value;
        l.Update(body.Blocks.GetRange(ch.Index, ch.Count));
    }

    static void Same(DocumentLayout incremental, Story body)
    {
        var full = new DocumentLayout(body, new StyleResolver(new StyleSheet(XElement.Parse(Styles))), new MonoFontProvider(ascent: 0.8, descent: 0.2));
        full.LayoutAll();
        Assert.Equal(full.Pages.Count, incremental.Pages.Count);
        for (int i = 0; i < full.Pages.Count; i++)
        {
            Assert.Same(full.Pages[i].Start.Item?.Block, incremental.Pages[i].Start.Item?.Block);
            Assert.Equal(full.Pages[i].Start.Line, incremental.Pages[i].Start.Line);
            Assert.Equal(full.Pages[i].Lines.Count, incremental.Pages[i].Lines.Count);
        }
    }
}
