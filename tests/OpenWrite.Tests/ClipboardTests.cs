using System.Xml.Linq;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Interop;
using OpenWrite.Core.Model;
using OpenWrite.Docx;

namespace OpenWrite.Tests;

/// <summary>Lot 3 : presse-papiers — lecture RTF et HTML, morceaux copiés et collés dans le modèle.</summary>
public class ClipboardTests
{
    const string NsDecl = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"";

    static string T(string text, string rPr = "") => $"<w:r>{(rPr.Length > 0 ? $"<w:rPr>{rPr}</w:rPr>" : "")}<w:t xml:space=\"preserve\">{text}</w:t></w:r>";

    static Story Body(params string[] paragraphs)
    {
        var xml = XDocument.Parse($"<w:document {NsDecl}><w:body>{string.Concat(paragraphs)}</w:body></w:document>");
        return StoryMapping.Read(xml, "word/document.xml", "document").Stories().Single();
    }

    static string[] Texts(Story s) => s.Blocks.OfType<Paragraph>().Select(p => p.Text).ToArray();
    static readonly Func<string> Ids = FlowEdits.IdGenerator(Array.Empty<string>());

    [Fact]
    public void Rtf_de_Word_paragraphes_et_mise_en_forme()
    {
        const string rtf = @"{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fswiss\fcharset0 Calibri;}{\f1 Times New Roman;}}" +
                           @"{\colortbl ;\red255\green0\blue0;}{\*\generator Riched20;}{\info{\author X}}" +
                           @"\pard\qc\f0\fs22 Bonjour {\b le monde}\par" + "\r\n" +
                           @"\pard\f1\fs24 caf\'e9 {\cf1\i rouge}\tab fin\line suite " + "\\u8364?\\par}";
        var paras = RtfReader.Read(rtf, out var mark);
        Assert.True(mark);
        Assert.Equal(2, paras.Count);
        Assert.Equal("center", paras[0].Align);
        Assert.Equal(new[] { "Bonjour ", "le monde" }, paras[0].Runs.Select(r => r.Text).ToArray());
        Assert.True(paras[0].Runs[1].Format.Bold);
        Assert.Equal("Calibri", paras[0].Runs[0].Format.Font);
        Assert.Equal(11, paras[0].Runs[0].Format.Size);
        Assert.Equal("café rouge\tfin\vsuite €", string.Concat(paras[1].Runs.Select(r => r.Text)));
        var red = paras[1].Runs.Single(r => r.Text == "rouge").Format;
        Assert.True(red.Italic);
        Assert.Equal("FF0000", red.Color);
        Assert.Null(paras[1].Align);
    }

    [Fact]
    public void Html_de_navigateur_balises_et_styles()
    {
        const string cf = "Version:0.9\r\nStartHTML:0\r\n<html><body><!--StartFragment--><h1>Titre</h1>" +
                          "<p style=\"text-align:center\">Un <b>gras</b> et <span style=\"font-style:italic; color:#c00\">italique</span>&nbsp;!</p>" +
                          "<p>ligne<br>suivante</p><table><tr><td>A</td><td>B</td></tr></table><script>x()</script><!--EndFragment--></body></html>";
        var paras = HtmlReader.Read(HtmlReader.Fragment(cf), out _);
        Assert.Equal(new[] { "Titre", "Un gras et italique !", "ligne\vsuivante", "A\tB" }, paras.Select(p => string.Concat(p.Runs.Select(r => r.Text))).ToArray());
        Assert.True(paras[0].Runs[0].Format.Bold);
        Assert.Equal("center", paras[1].Align);
        var it = paras[1].Runs.Single(r => r.Text.Trim() == "italique").Format;
        Assert.True(it.Italic);
        Assert.Equal("CC0000", it.Color);
        Assert.True(paras[1].Runs.Single(r => r.Text.Trim() == "gras").Format.Bold);
    }

    [Fact]
    public void Morceau_partiel_colle_au_milieu_d_un_paragraphe()
    {
        var body = Body($"<w:p w14:paraId=\"11111111\">{T("abc")}{T("DEF", "<w:b/>")}</w:p>", $"<w:p>{T("123456")}</w:p>");
        // copie de « cDE » (partiel) puis collage dans « 123|456 »
        var frag = Fragments.Extract(body, new DocPos(new[] { 0 }, 2), new DocPos(new[] { 0 }, 5));
        Assert.False(frag.EndsWithParagraph);
        var end = Fragments.Insert(body, new DocPos(new[] { 1 }, 3), frag, Ids);
        Assert.Equal(new[] { "abcDEF", "123cDE456" }, Texts(body));
        Assert.Equal(new DocPos(new[] { 1 }, 6), end);
        Assert.Contains(((Paragraph)body.Blocks[1]).Runs(), r => r.Text == "DE" && r.Props.Flag("b"));
    }

    [Fact]
    public void Paragraphes_entiers_colles_avec_leur_style_et_nouveaux_identifiants()
    {
        var body = Body($"<w:p w14:paraId=\"11111111\"><w:pPr><w:pStyle w:val=\"Titre1\"/></w:pPr>{T("Titre")}</w:p>",
                        $"<w:p w14:paraId=\"22222222\">{T("texte")}</w:p>", $"<w:p>{T("fin")}</w:p>");
        // les deux premiers paragraphes, marques comprises (jusqu'au début du 3e)
        var frag = Fragments.Extract(body, new DocPos(new[] { 0 }, 0), new DocPos(new[] { 2 }, 0));
        Assert.True(frag.EndsWithParagraph);
        Assert.Equal(2, frag.Blocks.Count);
        // collés au début de « fin » : deux nouveaux paragraphes avant
        var end = Fragments.Insert(body, new DocPos(new[] { 2 }, 0), frag, Ids);
        Assert.Equal(new[] { "Titre", "texte", "Titre", "texte", "fin" }, Texts(body));
        Assert.Equal("Titre1", ((Paragraph)body.Blocks[2]).StyleId);
        Assert.Equal(new DocPos(new[] { 4 }, 0), end);
        var ids = body.Blocks.OfType<Paragraph>().Select(p => p.Attributes.FirstOrDefault(a => a.Name.LocalName == "paraId")?.Value).Where(v => v is not null).ToList();
        Assert.Equal(ids.Count, ids.Distinct().Count());
    }

    [Fact]
    public void Signets_et_appels_de_note_non_dupliques()
    {
        var body = Body($"<w:p><w:bookmarkStart w:id=\"1\" w:name=\"ici\"/>{T("ab")}<w:bookmarkEnd w:id=\"1\"/><w:r><w:footnoteReference w:id=\"2\"/></w:r>{T("cd")}</w:p>");
        var frag = Fragments.Extract(body, new DocPos(new[] { 0 }, 0), new DocPos(new[] { 0 }, 5));
        var p = (Paragraph)frag.Blocks.Single();
        Assert.DoesNotContain(p.Inlines, i => i is BookmarkStart or BookmarkEnd);
        Assert.Equal("abcd", p.Text);
    }

    [Fact]
    public void Collage_RTF_annulable_en_une_fois()
    {
        var body = Body($"<w:p>{T("Avant après")}</w:p>");
        var journal = new UndoJournal();
        var paras = RtfReader.Read(@"{\rtf1\ansi {\b un}\par deux}", out var mark);
        var frag = Fragments.Build(paras, mark);
        var end = Fragments.Paste(journal, body, "Coller", new DocPos(new[] { 0 }, 6), new DocPos(new[] { 0 }, 6), frag, Ids);
        Assert.Equal(new[] { "Avant un", "deuxaprès" }, Texts(body));
        Assert.Equal(new DocPos(new[] { 1 }, 4), end);
        journal.Undo();
        Assert.Equal(new[] { "Avant après" }, Texts(body));
    }
}
