using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Docx;

namespace OpenWrite.Tests;

/// <summary>
/// Lot 8 : commentaires. Ajout sur une plage (marques et appel dans le corps, comments.xml et commentsExtended.xml
/// créés), réponse, résolution, modification, suppression d'un fil, annulation ; le document écrit est valide pour le SDK
/// et se relit à l'identique.
/// </summary>
public class CommentTests
{
    const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    static readonly XNamespace Wn = W;
    static readonly XNamespace W15 = DocxPackage.W15;

    static byte[] Source() => RoundTripTests.Docx($"""
        <w:document xmlns:w="{W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body>
        <w:p w14:paraId="10000001"><w:r><w:t xml:space="preserve">Bonjour le monde</w:t></w:r></w:p>
        <w:p w14:paraId="10000002"><w:r><w:t>Suite</w:t></w:r></w:p>
        <w:sectPr/></w:body></w:document>
        """);

    sealed class Session
    {
        public readonly DocxPackage Pkg;
        public readonly UndoJournal Journal = new();
        public readonly CommentStore Store;
        public readonly Func<string> Ids = FlowEdits.IdGenerator(new[] { "10000001", "10000002" }, 0x2000);
        public Story Body => Pkg.Main!.Stories().First();

        public Session(byte[] docx)
        {
            Pkg = DocxPackage.Open(new MemoryStream(docx));
            Store = new CommentStore(Pkg);
            // comme la fenêtre : toute modification marque sa partie à réécrire
            Journal.Changed += root => { if (Pkg.Stories.Values.FirstOrDefault(p => p.Stories().Contains(root)) is { } part) part.Dirty = true; };
            Journal.XmlChanged += root => { if (Pkg.XmlParts.FirstOrDefault(p => ReferenceEquals(p.Doc.Root, root)) is { } xp) xp.Dirty = true; };
            Journal.ListChanged += list => { if (Pkg.Stories.Values.FirstOrDefault(p => ReferenceEquals(p.Items, list)) is { } part) part.Dirty = true; };
        }

        public byte[] Save()
        {
            var ms = new MemoryStream();
            Pkg.Save(ms);
            return ms.ToArray();
        }
    }

    static DocPos At(int block, int offset) => new(new[] { block }, offset);

    static string Entry(byte[] docx, string name)
    {
        using var zip = new ZipArchive(new MemoryStream(docx));
        using var r = new StreamReader(zip.GetEntry(name)!.Open());
        return r.ReadToEnd();
    }

    [Fact]
    public void Ajouter_repondre_resoudre_puis_relire()
    {
        var s = new Session(Source());
        var id = s.Store.Add(s.Journal, s.Body, At(0, 8), At(0, 10), "Alice Martin", "AM", "Préciser\nquel monde ?", s.Ids);
        var reply = s.Store.Reply(s.Journal, s.Body, id, "Bob", "B", "D'accord", s.Ids);
        s.Store.SetDone(s.Journal, s.Body, id, true);

        var all = s.Store.All();
        Assert.Equal(2, all.Count);
        Assert.Equal("Préciser\nquel monde ?", all[0].Text);
        Assert.True(all[0].Done);
        Assert.Equal(all[0].ParaId, all[1].ParentParaId);
        var anchor = CommentStore.Anchor(s.Body, id)!;
        Assert.Equal((8, 10), (anchor.Start, anchor.End));
        Assert.Equal((8, 10), (CommentStore.Anchor(s.Body, reply)!.Start, CommentStore.Anchor(s.Body, reply)!.End));

        var output = s.Save();
        // relu : les mêmes commentaires, le corps garde ses marques et son texte
        var again = new CommentStore(DocxPackage.Open(new MemoryStream(output)));
        var re = again.All();
        Assert.Equal(new[] { "Alice Martin", "Bob" }, re.Select(c => c.Author));
        Assert.True(re[0].Done);
        Assert.Equal(re[0].ParaId, re[1].ParentParaId);
        var doc = XDocument.Parse(Entry(output, "word/document.xml"));
        Assert.Equal("Bonjour le monde", string.Concat(doc.Descendants(Wn + "p").First().Descendants(Wn + "t").Select(t => t.Value)));
        Assert.Equal(2, doc.Descendants(Wn + "commentReference").Count());
        Assert.Contains("commentsExtended", Entry(output, "word/_rels/document.xml.rels"));
        Assert.Contains("comments+xml", Entry(output, "[Content_Types].xml"));
        Assert.Equal("1", (string?)XDocument.Parse(Entry(output, "word/commentsExtended.xml")).Root!.Elements(W15 + "commentEx").First().Attribute(W15 + "done"));

        // le SDK ne trouve aucune erreur dans les parties des commentaires et le corps
        using var wd = WordprocessingDocument.Open(new MemoryStream(output), false);
        var errors = new OpenXmlValidator().Validate(wd).Select(e => $"{e.Part?.Uri} {e.Path?.XPath} {e.Description}").ToList();
        Assert.True(errors.Count == 0, string.Join("\n", errors));
    }

    [Fact]
    public void Commentaire_au_point_d_insertion_marques_dans_l_ordre()
    {
        var s = new Session(Source());
        var id = s.Store.Add(s.Journal, s.Body, At(0, 0), At(0, 0), "Alice", "A", "Ici", s.Ids);
        var names = ((Paragraph)s.Body.Blocks[0]).Inlines.Select(i => i switch
        {
            OpaqueInline o => o.Xml.Name.LocalName,
            Run r when r.Content.Any(c => c is OpaqueRunContent) => "appel",
            _ => "texte",
        }).ToList();
        Assert.Equal(new[] { "commentRangeStart", "commentRangeEnd", "appel", "texte" }, names);
        Assert.Equal((0, 0), (CommentStore.Anchor(s.Body, id)!.Start, CommentStore.Anchor(s.Body, id)!.End));
    }

    [Fact]
    public void Annuler_l_ajout_et_supprimer_un_fil()
    {
        var s = new Session(Source());
        var id = s.Store.Add(s.Journal, s.Body, At(0, 0), At(1, 5), "Alice", "A", "Tout", s.Ids);
        s.Store.Reply(s.Journal, s.Body, id, "Bob", "B", "Oui", s.Ids);
        Assert.Equal(2, s.Store.All().Count);

        s.Store.Delete(s.Journal, s.Body, id);
        Assert.Empty(s.Store.All());
        Assert.Null(CommentStore.Anchor(s.Body, id));
        Assert.Empty(s.Body.Paragraphs().SelectMany(p => p.Inlines).OfType<OpaqueInline>());

        s.Journal.Undo();                         // suppression annulée : le fil revient
        Assert.Equal(2, s.Store.All().Count);
        Assert.NotNull(CommentStore.Anchor(s.Body, id));
        s.Journal.Undo();                         // réponse
        s.Journal.Undo();                         // commentaire
        Assert.Empty(s.Store.All());
        Assert.Equal("Bonjour le monde", string.Concat(((Paragraph)s.Body.Blocks[0]).Runs().SelectMany(r => r.Content).OfType<TextContent>().Select(t => t.Value)));
        Assert.Empty(s.Body.Paragraphs().SelectMany(p => p.Inlines).OfType<OpaqueInline>());
        s.Journal.Redo();
        Assert.Single(s.Store.All());
    }

    [Fact]
    public void Modifier_le_texte_garde_le_lien_des_reponses()
    {
        var s = new Session(Source());
        var id = s.Store.Add(s.Journal, s.Body, At(0, 0), At(0, 7), "Alice", "A", "Avant", s.Ids);
        s.Store.Reply(s.Journal, s.Body, id, "Bob", "B", "Réponse", s.Ids);
        var paraId = s.Store.All()[0].ParaId;
        s.Store.Edit(s.Journal, s.Body, id, "Après\nsur deux lignes", s.Ids);
        var all = s.Store.All();
        Assert.Equal("Après\nsur deux lignes", all[0].Text);
        Assert.Equal(paraId, all[0].ParaId);
        Assert.Equal(paraId, all[1].ParentParaId);
        Assert.True(all[0].PlainText);
        s.Journal.Undo();
        Assert.Equal("Avant", s.Store.All()[0].Text);
    }
}
