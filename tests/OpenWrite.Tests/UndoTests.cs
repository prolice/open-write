using System.Xml.Linq;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Docx;
using static OpenWrite.Tests.RoundTripTests;

namespace OpenWrite.Tests;

/// <summary>Lot 1 : positions (DocPos), opérations de texte et annuler / rétablir.</summary>
public class UndoTests
{
    static string R(string inner, string rPr = "") => $"<w:r>{(rPr.Length > 0 ? $"<w:rPr>{rPr}</w:rPr>" : "")}{inner}</w:r>";
    static string T(string text, string rPr = "") => R($"<w:t xml:space=\"preserve\">{text}</w:t>", rPr);

    /// <summary>0 : « Bonjour le monde » (gras sur « le »), 1 : tableau 1 ligne 2 cellules, 2 : lien + champ + signet,
    /// 3 : paragraphe vide avec marque soulignée, 4 : texte supprimé suivi, 5 : contrôle de contenu, 6 : saut de section.</summary>
    static readonly string Sample = Root + "<w:body>" +
        "<w:p w14:paraId=\"10000001\" w14:textId=\"20000001\">" + T("Bonjour ") + T("le", "<w:b/>") + T(" monde") + "</w:p>" +
        "<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p>" + T("A1") + "</w:p></w:tc><w:tc><w:p>" + T("B1") + "</w:p><w:p>" + T("B1 suite") + "</w:p></w:tc></w:tr></w:tbl>" +
        "<w:p><w:hyperlink r:id=\"rId9\">" + T("lien ici", "<w:color w:val=\"0000FF\"/>") + "</w:hyperlink>" + R("<w:tab/>") +
            R("<w:fldChar w:fldCharType=\"begin\"/>") + R("<w:instrText> PAGE </w:instrText>") + R("<w:fldChar w:fldCharType=\"separate\"/>") + T("7") + R("<w:fldChar w:fldCharType=\"end\"/>") +
            "<w:bookmarkStart w:id=\"0\" w:name=\"fin\"/>" + T(" fin") + "<w:bookmarkEnd w:id=\"0\"/></w:p>" +
        "<w:p><w:pPr><w:rPr><w:u w:val=\"single\"/><w:ins w:id=\"5\" w:author=\"A\"/></w:rPr></w:pPr></w:p>" +
        "<w:p>" + "<w:del w:id=\"6\" w:author=\"A\">" + R("<w:delText>ancien</w:delText>") + "</w:del>" + T("nouveau") + "</w:p>" +
        "<w:sdt><w:sdtPr><w:id w:val=\"3\"/></w:sdtPr><w:sdtContent><w:p>" + T("dans le contrôle") + "</w:p></w:sdtContent></w:sdt>" +
        "<w:p><w:pPr><w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/></w:sectPr></w:pPr>" + T("Fin de section") + "</w:p>" +
        "<w:sectPr/></w:body></w:document>";

    static (DocxPackage Pkg, Story Body) Open()
    {
        var pkg = DocxPackage.Open(new MemoryStream(RoundTripTests.Docx(Sample)));
        return (pkg, pkg.Main!.Stories().Single());
    }

    static XElement Xml(DocxPackage pkg) => DocxPackage.ParseXml(DocxPackage.Serialize(pkg.Main!)).Root!;

    static DocPos P(int offset, params int[] path) => new(path, offset);

    static readonly Func<string> Ids = FlowEdits.IdGenerator(new[] { "10000001", "20000001" });

    [Fact]
    public void Positions_chemins_et_atomes()
    {
        var (_, body) = Open();
        Assert.Equal("B1 suite", P(0, 1, 0, 1, 1).Paragraph(body).Text);          // tableau, ligne 0, cellule 1, 2e paragraphe
        Assert.Equal("dans le contrôle", P(0, 5, 0).Paragraph(body).Text);         // contrôle de contenu
        var cell = P(0, 1, 0, 1, 1).Paragraph(body);
        Assert.Equal(P(3, 1, 0, 1, 1), DocPos.Of(body, cell, 3));
        Assert.True(P(5, 0) < P(0, 1, 0, 0, 0) && P(0, 1, 0, 1, 1) < P(0, 2));       // ordre de lecture
        Assert.Throws<ArgumentException>(() => P(0, 1).Paragraph(body));              // un tableau n'est pas un paragraphe
        // « lien ici » (8) + tabulation (1) + début (1) + « PAGE » (6) + séparateur (1) + « 7 » (1) + fin (1) + « fin » (4) ; signets : 0
        var p2 = P(0, 2).Paragraph(body);
        Assert.Equal(23, ParagraphAtoms.Length(p2));
        Assert.Equal(new[] { 0, 0 }, ParagraphAtoms.Of(p2).Where(a => a.Inline is BookmarkStart or BookmarkEnd).Select(a => a.Length));
    }

    [Fact]
    public void Inserer_prend_la_mise_en_forme_du_caractere_qui_precede()
    {
        var (_, body) = Open();
        var p0 = P(0, 0).Paragraph(body);
        Assert.Equal(P(13, 0), TextOps.Insert(body, P(10, 0), "!!!"));                // juste après « le » (gras) → en gras
        Assert.Equal("Bonjour le!!! monde", p0.Text);
        Assert.Equal("le!!!", p0.Runs().Single(r => r.Props.Flag("b")).Text);
        TextOps.Insert(body, P(0, 0), "« ");                                         // début : mise en forme du caractère qui suit
        Assert.Equal("« Bonjour ", p0.Runs().First().Text);
        TextOps.Insert(body, P(3, 0), "j");                                          // au milieu d'un texte
        Assert.Equal("« Bjonjour le!!! monde", p0.Text);

        var empty = P(0, 3).Paragraph(body);                                          // paragraphe vide : marque de paragraphe, sans sa révision
        TextOps.Insert(body, P(0, 3), "Nouveau");
        var run = empty.Runs().Single();
        Assert.Equal("single", run.Props.Get("u")!.Attribute(run.Props.Ns + "val")!.Value);
        Assert.Null(run.Props.Get("ins"));

        var tracked = P(0, 4).Paragraph(body);                                        // jamais dans le texte supprimé
        TextOps.Insert(body, P(6, 4), "X");
        Assert.Equal("ancien", tracked.Inlines.OfType<InlineContainer>().Single().Runs().Single().Content.OfType<TextContent>().Single().Value);
        Assert.Equal("Xnouveau", tracked.Runs().Last().Text);
        Assert.Throws<ArgumentException>(() => TextOps.Insert(body, P(0, 0), "a\tb"));
    }

    [Fact]
    public void Supprimer_dans_un_paragraphe_et_sur_plusieurs()
    {
        var (_, body) = Open();
        TextOps.Delete(body, P(5, 0), P(10, 0));                                     // « ur le » : le segment gras disparaît
        Assert.Equal("Bonjo monde", P(0, 0).Paragraph(body).Text);
        Assert.DoesNotContain(P(0, 0).Paragraph(body).Runs(), r => r.Props.Flag("b"));

        var p2 = P(0, 2).Paragraph(body);
        TextOps.Delete(body, P(19, 2), P(22, 2));                                    // « fi » de « fin » : le signet reste
        Assert.Equal(2, p2.Inlines.Count(i => i is BookmarkStart or BookmarkEnd));
        TextOps.Delete(body, P(0, 2), P(8, 2));                                      // tout le texte du lien : le lien vide disparaît
        Assert.DoesNotContain(p2.Inlines, i => i is InlineContainer);

        // du milieu du paragraphe 0 au milieu du paragraphe 6 : tableau, lien, contrôle… disparaissent, la section suit
        int before = body.Blocks.Count;
        TextOps.Delete(body, P(3, 0), P(4, 6));                                      // « Fin » + espace retirés du dernier
        Assert.Equal(before - 6, body.Blocks.Count);
        var merged = P(0, 0).Paragraph(body);
        Assert.Equal("Bonde section", merged.Text);
        Assert.NotNull(merged.Props.Get("sectPr"));
    }

    [Fact]
    public void Scinder_au_milieu_d_un_lien_duplique_le_lien()
    {
        var (_, body) = Open();
        var second = TextOps.Split(body, P(4, 2), Ids);
        Assert.Equal(P(0, 3), second);
        Assert.Equal("lien", P(0, 2).Paragraph(body).Text);
        var p3 = P(0, 3).Paragraph(body);
        Assert.StartsWith(" ici\t", p3.Text);
        Assert.Equal(" ici", p3.Inlines.OfType<InlineContainer>().Single().Text);   // second lien, mêmes attributs
        Assert.Equal("0000FF", p3.Runs().First().Props.Val("color"));
        Assert.Single(FieldIndex.Build(body).All);                                    // le champ est passé entier au second
        // fin de paragraphe : second paragraphe vide ; saut de section sur le second
        var last = TextOps.Split(body, P(14, 7), Ids);
        Assert.Equal("", last.Paragraph(body).Text);
        Assert.NotNull(last.Paragraph(body).Props.Get("sectPr"));
        Assert.Null(P(0, 7).Paragraph(body).Props.Get("sectPr"));
        Assert.NotEqual("10000001", TextOps.Split(body, P(3, 0), Ids).Paragraph(body).Attributes.Single(a => a.Name.LocalName == "paraId").Value);
    }

    [Fact]
    public void Annuler_tout_redonne_le_xml_d_origine_et_retablir_tout_redonne_le_resultat()
    {
        var (pkg, body) = Open();
        var original = Xml(pkg);
        var j = new UndoJournal();
        var pos = TextOps.Type(j, body, P(16, 0), " entier");
        pos = TextOps.SplitParagraph(j, body, pos, Ids);
        TextOps.Type(j, body, pos, "Suite");
        // le tableau est maintenant le bloc 2 : dans sa 2e cellule, de « B1 » au « B » de « B1 suite »
        TextOps.DeleteRange(j, body, P(0, 2, 0, 1, 0), P(1, 2, 0, 1, 1));
        Assert.Equal("1 suite", P(0, 2, 0, 1, 0).Paragraph(body).Text);
        TextOps.Type(j, body, P(0, 2, 0, 1, 0), ">");                               // dans une cellule
        Assert.Throws<NotSupportedException>(() => TextOps.DeleteRange(j, body, P(1, 2, 0, 0, 0), P(1, 2, 0, 1, 0)));   // d'une cellule à l'autre
        TextOps.DeleteRange(j, body, P(3, 0), P(2, 3));                              // à travers le tableau
        using (var tx = j.Begin("Case et signet", body, Array.Empty<int>(), 2, 1))
        {
            MarkerOps.AddAroundParagraph(P(0, 2).Paragraph(body), "Nouveau", "9");
            tx.Commit();
        }
        var edited = Xml(pkg);
        Assert.NotEmpty(XmlCompare.Diff(original, edited));

        while (j.CanUndo) j.Undo();
        Assert.Empty(XmlCompare.Diff(original, Xml(pkg)));
        while (j.CanRedo) j.Redo();
        Assert.Empty(XmlCompare.Diff(edited, Xml(pkg)));
    }

    [Fact]
    public void La_frappe_continue_est_regroupee_par_mot()
    {
        var (_, body) = Open();
        var j = new UndoJournal();
        var pos = P(16, 0);
        foreach (var ch in " et la lune") pos = TextOps.Type(j, body, pos, ch.ToString());
        Assert.Equal("Bonjour le monde et la lune", P(0, 0).Paragraph(body).Text);
        Assert.Equal(4, j.Count);                                                    // « ␣ », « et␣ », « la␣ », « lune » : l'espace termine le mot
        Assert.Equal(P(23, 0), j.Undo());
        Assert.Equal("Bonjour le monde et la ", P(0, 0).Paragraph(body).Text);
        // une frappe ailleurs commence une nouvelle entrée
        j.Redo();
        TextOps.Type(j, body, P(0, 0), "x");
        TextOps.Type(j, body, P(5, 0), "y");
        Assert.Equal(6, j.Count);
    }

    [Fact]
    public void Une_operation_qui_echoue_ne_laisse_rien()
    {
        var (pkg, body) = Open();
        var original = Xml(pkg);
        var j = new UndoJournal();
        void Fail()
        {
            using var tx = j.Begin("Échec", body, Array.Empty<int>(), 0, 1);
            TextOps.Insert(body, P(0, 0), "à moitié ");
            TextOps.Split(body, P(3, 0), Ids);                                        // ajoute un bloc
            throw new InvalidOperationException("panne");
        }
        Assert.Throws<InvalidOperationException>(Fail);
        Assert.Empty(XmlCompare.Diff(original, Xml(pkg)));
        Assert.False(j.CanUndo);
    }

    [Fact]
    public void Etat_enregistre()
    {
        var (_, body) = Open();
        var j = new UndoJournal();
        Assert.False(j.IsModified);
        TextOps.Type(j, body, P(0, 0), "a");
        Assert.True(j.IsModified);
        j.MarkSaved();
        Assert.False(j.IsModified);
        TextOps.Type(j, body, P(1, 0), "b");                                         // après l'enregistrement : nouvelle entrée
        Assert.Equal(2, j.Count);
        j.Undo();
        Assert.False(j.IsModified);                                                   // retour à l'état enregistré
        j.Undo();
        Assert.True(j.IsModified);
        TextOps.Type(j, body, P(0, 0), "z");                                         // l'état enregistré n'est plus atteignable
        j.Undo();
        Assert.True(j.IsModified);
    }

    [Fact]
    public void Changed_signale_la_racine_modifiee()
    {
        var (_, body) = Open();
        var j = new UndoJournal();
        var changed = new List<Story>();
        j.Changed += changed.Add;
        TextOps.Type(j, body, P(0, 0), "a");
        j.Undo();
        j.Redo();
        Assert.Equal(new[] { body, body, body }, changed);
    }
}
