using System.Xml.Linq;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Core.Review;
using OpenWrite.Docx;

namespace OpenWrite.Tests;

/// <summary>
/// Lot 8 : suivi des modifications. Frappe, suppression, Entrée, Retour arrière et mise en forme enregistrés comme Word
/// (w:ins, w:del + w:delText, marques de paragraphe, w:rPrChange, w:pPrChange) ; accepter et refuser une révision ou
/// toutes ; le XML écrit est celui que Word attend.
/// </summary>
public class ReviewTests
{
    const string NsDecl = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";
    static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    static string T(string text, string rPr = "") => $"<w:r>{(rPr.Length > 0 ? $"<w:rPr>{rPr}</w:rPr>" : "")}<w:t xml:space=\"preserve\">{text}</w:t></w:r>";
    static string P(params string[] runs) => $"<w:p>{string.Concat(runs)}</w:p>";

    static (StoryPart Part, Story Body) Doc(params string[] paragraphs)
    {
        var xml = XDocument.Parse($"<w:document {NsDecl}><w:body>{string.Concat(paragraphs)}<w:sectPr/></w:body></w:document>");
        var part = StoryMapping.Read(xml, "word/document.xml", "document");
        return (part, part.Stories().Single());
    }

    static TrackContext Ctx(string author = "Alice", int firstId = 100)
    {
        int id = firstId;
        return new TrackContext(() => id++) { Author = author, FixedDate = new DateTime(2026, 11, 11, 10, 0, 0) };
    }

    /// <summary>Texte du document tel que Word l'afficherait sans marques (insertions gardées, suppressions masquées).</summary>
    static string Final(Paragraph p) => string.Concat(p.Runs().SelectMany(r => r.Content).OfType<TextContent>().Where(t => !t.Deleted).Select(t => t.Value));

    /// <summary>Texte d'origine (suppressions visibles, insertions masquées).</summary>
    static string Original(Paragraph p)
    {
        var sb = new System.Text.StringBuilder();
        void Walk(List<Inline> list, bool inserted)
        {
            foreach (var i in list)
                switch (i)
                {
                    case Run r when !inserted: foreach (var t in r.Content.OfType<TextContent>()) sb.Append(t.Value); break;
                    case InlineContainer b: Walk(b.Children, inserted || b.Name.LocalName is "ins" or "moveTo"); break;
                }
        }
        Walk(p.Inlines, false);
        return sb.ToString();
    }

    static XElement Xml(StoryPart part) => StoryMapping.Write(part).Root!;

    static DocPos At(int block, int offset) => new(new[] { block }, offset);

    [Fact]
    public void Frappe_suivie_dans_un_seul_w_ins_avec_auteur_et_date()
    {
        var (part, body) = Doc(P(T("abc")));
        using (TrackContext.Use(Ctx()))
        {
            var pos = TextOps.Insert(body, At(0, 1), "X");
            TextOps.Insert(body, pos, "Y");
        }
        var p = (Paragraph)body.Blocks[0];
        Assert.Equal("aXYbc", Final(p));
        Assert.Equal("abc", Original(p));
        var ins = Xml(part).Descendants(W + "ins").ToList();
        Assert.Single(ins);
        Assert.Equal("Alice", (string?)ins[0].Attribute(W + "author"));
        Assert.Equal("2026-11-11T10:00:00Z", (string?)ins[0].Attribute(W + "date"));
        Assert.Equal("XY", string.Concat(ins[0].Descendants(W + "t").Select(t => t.Value)));
    }

    [Fact]
    public void Suppression_suivie_en_delText_et_propre_ajout_retire_vraiment()
    {
        var (part, body) = Doc(P(T("abcd")));
        var ctx = Ctx();
        using (TrackContext.Use(ctx))
        {
            TextOps.Delete(body, At(0, 1), At(0, 2));                 // « b » barré
            Assert.Equal(At(0, 2), ctx.LastDeletionEnd);
            TextOps.Insert(body, At(0, 4), "ZZ");                     // ajout à la fin
            TextOps.Delete(body, At(0, 5), At(0, 6));                 // un Z de l'ajout : retiré vraiment
        }
        var p = (Paragraph)body.Blocks[0];
        Assert.Equal("acdZ", Final(p));
        Assert.Equal("abcd", Original(p));
        var xml = Xml(part);
        Assert.Equal("b", xml.Descendants(W + "del").Single().Descendants(W + "delText").Single().Value);
        Assert.Empty(xml.Descendants(W + "del").Descendants(W + "t"));
        Assert.Equal("Z", xml.Descendants(W + "ins").Single().Value);
    }

    [Fact]
    public void Suppression_d_un_texte_deja_barre_ne_change_rien()
    {
        var (_, body) = Doc(P(T("abc")));
        using (TrackContext.Use(Ctx()))
        {
            TextOps.Delete(body, At(0, 0), At(0, 2));
            TextOps.Delete(body, At(0, 0), At(0, 3));
        }
        var revs = Revisions.List(body);
        Assert.Equal(new[] { RevisionKind.Deleted }, revs.Select(r => r.Kind).Distinct());
        Assert.Equal("", Final((Paragraph)body.Blocks[0]));
        Assert.Equal("abc", Original((Paragraph)body.Blocks[0]));
    }

    [Fact]
    public void Entree_suivie_marque_inseree_et_retour_arriere_sur_sa_marque_refusionne()
    {
        var (part, body) = Doc(P(T("abcd")));
        using (TrackContext.Use(Ctx()))
        {
            TextOps.Split(body, At(0, 2), FlowEdits.IdGenerator(Array.Empty<string>()));
            Assert.Equal(2, body.Blocks.Count);
            Assert.NotNull(((Paragraph)body.Blocks[0]).Props.Get("rPr")?.Element(W + "ins"));
            // Retour arrière au début du second paragraphe : sa propre marque → vraie fusion
            TextOps.Delete(body, At(0, 2), At(1, 0));
        }
        Assert.Single(body.Blocks);
        Assert.Equal("abcd", Final((Paragraph)body.Blocks[0]));
        Assert.Empty(Revisions.List(body));
        Assert.Empty(Xml(part).Descendants(W + "ins"));
    }

    [Fact]
    public void Fusion_de_paragraphes_d_un_autre_auteur_marque_supprimee_puis_acceptee()
    {
        var (_, body) = Doc(P(T("ab")), P(T("cd")));
        using (TrackContext.Use(Ctx()))
            TextOps.Delete(body, At(0, 2), At(1, 0));
        Assert.Equal(2, body.Blocks.Count);
        var r = Revisions.List(body).Single();
        Assert.Equal(RevisionKind.ParagraphDeleted, r.Kind);
        Revisions.Accept(body, r);
        Assert.Single(body.Blocks);
        Assert.Equal("abcd", Final((Paragraph)body.Blocks[0]));
    }

    [Fact]
    public void Accepter_tout_et_refuser_tout()
    {
        string Run(bool accept)
        {
            var (_, body) = Doc(P(T("Le chat dort")));
            using (TrackContext.Use(Ctx()))
            {
                TextOps.Delete(body, At(0, 3), At(0, 7));            // « chat » barré
                TextOps.Insert(body, At(0, 3), "chien");             // « chien » inséré
                TextOps.Split(body, At(0, 0), FlowEdits.IdGenerator(Array.Empty<string>()));
            }
            Assert.Equal(3, Revisions.List(body).Count);
            Assert.Equal(3, Revisions.ResolveAll(body, accept));
            Assert.Empty(Revisions.List(body));
            return string.Join("|", body.Blocks.OfType<Paragraph>().Select(Final));
        }
        Assert.Equal("|Le chien dort", Run(accept: true));
        Assert.Equal("Le chat dort", Run(accept: false));
    }

    [Fact]
    public void Refuser_d_un_seul_auteur()
    {
        var (_, body) = Doc(P(T("abc")));
        using (TrackContext.Use(Ctx("Alice", 1))) TextOps.Insert(body, At(0, 3), "A");
        using (TrackContext.Use(Ctx("Bob", 50))) TextOps.Insert(body, At(0, 0), "B");
        Assert.Equal(new[] { "Bob", "Alice" }, Revisions.List(body).Select(r => r.Author));
        Revisions.ResolveAll(body, accept: false, r => r.Author == "Bob");
        Assert.Equal("abcA", Final((Paragraph)body.Blocks[0]));
        Assert.Equal("Alice", Revisions.List(body).Single().Author);
    }

    [Fact]
    public void Mise_en_forme_suivie_garde_l_ancienne_et_se_refuse()
    {
        var (part, body) = Doc(P(T("abc", "<w:i/>")));
        using (TrackContext.Use(Ctx()))
        {
            FormatOps.ApplyToRuns(body, At(0, 0), At(0, 3), (_, r) => FormatOps.SetToggle(r, "b", true));
            FormatOps.ApplyToParagraphs(body, At(0, 0), At(0, 3), p => p.Props.SetVal("jc", "center"));
        }
        var xml = Xml(part);
        var change = xml.Descendants(W + "rPrChange").Single();
        Assert.NotNull(change.Element(W + "rPr")!.Element(W + "i"));
        Assert.Null(change.Element(W + "rPr")!.Element(W + "b"));
        Assert.Equal("Alice", (string?)change.Attribute(W + "author"));
        Assert.NotNull(xml.Descendants(W + "pPrChange").Single().Element(W + "pPr"));
        // rPrChange est le dernier enfant de w:rPr (ordre du schéma)
        Assert.Equal("rPrChange", change.Parent!.Elements().Last().Name.LocalName);

        Revisions.ResolveAll(body, accept: false);
        var run = ((Paragraph)body.Blocks[0]).Runs().Single();
        Assert.False(run.Props.Flag("b"));
        Assert.True(run.Props.Flag("i"));
        Assert.Null(((Paragraph)body.Blocks[0]).Props.Val("jc"));
    }

    [Fact]
    public void Lecture_des_revisions_de_Word_et_identifiants_suivants()
    {
        var (_, body) = Doc(
            "<w:p><w:pPr><w:rPr><w:ins w:id=\"7\" w:author=\"Word\" w:date=\"2025-01-02T03:04:00Z\"/></w:rPr></w:pPr>" +
            "<w:ins w:id=\"3\" w:author=\"Word\"><w:r><w:t>neuf</w:t></w:r></w:ins>" +
            "<w:del w:id=\"4\" w:author=\"Word\"><w:r><w:delText>vieux</w:delText></w:r></w:del></w:p>",
            P(T("suite")));
        var revs = Revisions.List(body);
        Assert.Equal(new[] { RevisionKind.Inserted, RevisionKind.Deleted, RevisionKind.ParagraphInserted }, revs.Select(r => r.Kind));
        Assert.Equal((0, 4), (revs[0].Start, revs[0].End));
        Assert.Equal((4, 9), (revs[1].Start, revs[1].End));
        Assert.Equal(new DateTime(2025, 1, 2, 3, 4, 0), revs[2].Date);
        Assert.Equal(8, TrackContext.IdsAfter(new[] { body })());
        // refuser la marque insérée : le paragraphe rejoint le suivant
        Revisions.Reject(body, revs[2]);
        Assert.Single(body.Blocks);
        Assert.Equal("neufsuite", Final((Paragraph)body.Blocks[0]));
    }

    [Fact]
    public void Collage_suivi_marque_texte_et_paragraphes()
    {
        var (_, body) = Doc(P(T("ab")));
        Fragments.ParaSpec Spec(string text) { var s = new Fragments.ParaSpec(); s.Runs.Add((text, new Fragments.CharFormat())); return s; }
        var fragment = Fragments.Build(new[] { Spec("un"), Spec("deux") }, lastHasMark: false);
        using (TrackContext.Use(Ctx()))
            Fragments.Insert(body, At(0, 1), fragment, FlowEdits.IdGenerator(Array.Empty<string>()));
        Assert.Equal("aun|deuxb", string.Join("|", body.Blocks.OfType<Paragraph>().Select(Final)));
        Assert.Equal("ab", string.Concat(body.Blocks.OfType<Paragraph>().Select(Original)));
        Assert.Contains(Revisions.List(body), r => r.Kind == RevisionKind.ParagraphInserted);
        Revisions.ResolveAll(body, accept: false);
        Assert.Equal("ab", string.Join("|", body.Blocks.OfType<Paragraph>().Select(Final)));
    }
}
