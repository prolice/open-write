using System.Diagnostics;
using System.Windows;
using System.Windows.Input;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Docx;
using OpenWrite.Layout;
using Layout = OpenWrite.App.Layout;

namespace OpenWrite.App;

// Lot 3 (version 0.4) : liens hypertextes. Ctrl+clic suit le lien (adresse web ou de courriel, signet du document) ;
// Ctrl+K insère un lien sur la sélection (ou un nouveau texte) ou modifie celui du point d'insertion ; la bulle d'un
// lien survolé montre son adresse, comme Word.
public partial class MainWindow
{
    /// <summary>Lien qui contient le caractère affiché <paramref name="index"/> du paragraphe, et sa plage dans le modèle.</summary>
    (InlineContainer Link, int Start, int End)? LinkAt(Paragraph p, int index)
    {
        if (LayoutOf(p) is not { } pl || index < 0 || index >= pl.Content.Text.Length) return null;
        return LinkOps.LinkAt(p, pl.Content.ModelStart[index]);
    }

    /// <summary>Caractère sous un point de la vue (et non le point d'insertion le plus proche) ; null hors du texte.</summary>
    TextPoint? HitCharacter(Point viewPoint)
    {
        if (_pagesView is null || _layout is null) return null;
        var (page, at) = _pagesView.HitPage(viewPoint);
        if (page < 0 || Layout.PagesView.HitLine(_layout.Pages[page], at) is not { } hit || hit.Char < 0) return null;
        return new TextPoint(hit.Line.Paragraph.Paragraph, hit.Char);
    }

    /// <summary>Adresse d'un lien : URL de sa relation, ou « #signet » pour un lien interne.</summary>
    string? LinkTarget(InlineContainer link)
    {
        var anchor = link.Attributes.FirstOrDefault(a => a.Name.LocalName == "anchor")?.Value;
        var url = _pkg?.Main is { } main ? _pkg.HyperlinkTarget(main, link) : null;
        return url is not null ? url + (anchor is not null ? "#" + anchor : "") : anchor is not null ? "#" + anchor : null;
    }

    /// <summary>Survol : bulle avec l'adresse du lien ; main si Ctrl est enfoncé.</summary>
    bool HoverLink(TextPoint? t)
    {
        if (_pagesView is null) return false;
        var link = t is { } tp ? LinkAt(tp.Paragraph, tp.Index) : null;
        if (link is null)
        {
            if (_pagesView.ToolTip is not null) _pagesView.ToolTip = null;
            return false;
        }
        var tip = $"{LinkTarget(link.Value.Link) ?? "(lien sans adresse)"}\nCtrl+clic pour suivre le lien";
        if (_pagesView.ToolTip as string != tip) _pagesView.ToolTip = tip;
        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Control)) _pagesView.Cursor = Cursors.Hand;
        return true;
    }

    /// <summary>Ctrl+clic : ouvre l'adresse (navigateur, courriel) ou va au signet visé.</summary>
    internal bool FollowLink(TextPoint t)
    {
        if (LinkAt(t.Paragraph, t.Index) is not { } found) return false;
        var anchor = found.Link.Attributes.FirstOrDefault(a => a.Name.LocalName == "anchor")?.Value;
        var url = _pkg?.Main is { } main ? _pkg.HyperlinkTarget(main, found.Link) : null;
        if (url is null && anchor is not null) return GoToBookmark(anchor);
        if (url is null) return false;
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri)) { StatusText.Text = $"Adresse illisible : {url}"; return true; }
        // adresses web et de courriel : ouvertes directement ; le reste (fichier, programme) après confirmation
        if (uri.Scheme is not ("http" or "https" or "mailto")
            && MessageBox.Show(this, $"Ce lien ouvre :\n\n{url}\n\nN'ouvrez que les liens de documents dont vous connaissez la provenance. Continuer ?",
                               "Suivre le lien", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            return true;
        try { Process.Start(new ProcessStartInfo(uri.AbsoluteUri) { UseShellExecute = true }); }
        catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or InvalidOperationException) { StatusText.Text = $"Impossible d'ouvrir {url} : {ex.Message}"; }
        return true;
    }

    internal bool GoToBookmark(string name)
    {
        if (Body is not { } body || BookmarkIndex.Build(body).Find(name) is not { StartParagraph: { } p } bm) { StatusText.Text = $"Signet « {name} » introuvable."; return true; }
        var atom = ParagraphAtoms.Of(p).FirstOrDefault(a => ReferenceEquals(a.Inline, bm.Start));
        try
        {
            if (FromDocPos(DocPos.Of(body, p, atom?.Start ?? 0)) is { } t) MoveTo(t, extend: false);
        }
        catch (ArgumentException) { }
        return true;
    }

    /// <summary>Style de caractère « Hyperlink » du document (nom intégré de Word), s'il existe.</summary>
    string? HyperlinkStyleId()
    {
        if (_styles?.Sheet is not { } sheet || sheet.Styles is null) return null;
        var w = sheet.W;
        return sheet.Styles.Elements(w + "style").FirstOrDefault(s => s.Attribute(w + "type")?.Value == "character"
            && string.Equals(s.Element(w + "name")?.Attribute(w + "val")?.Value, "Hyperlink", StringComparison.OrdinalIgnoreCase))?.Attribute(w + "styleId")?.Value;
    }

    /// <summary>Ctrl+K : insérer un lien (sur la sélection ou un nouveau texte) ou modifier / supprimer celui du point d'insertion.</summary>
    internal void ShowLinkDialog()
    {
        if (Body is not { } body || _pkg?.Main is not { } main || Ordered() is not { } sel) return;
        if (!ReferenceEquals(sel.From.Paragraph, sel.To.Paragraph)) { StatusText.Text = "Un lien se pose dans un seul paragraphe : sélectionnez moins de texte."; return; }
        var p = sel.From.Paragraph;
        if (LayoutOf(p) is not { } pl || ToDocPos(sel.From) is not { } from || ToDocPos(sel.To) is not { } to) { NotEditable(); return; }
        // lien existant sous le point d'insertion (ou au début de la sélection)
        var existing = LinkAt(p, sel.From.Index) ?? (sel.From.Index > 0 && !HasSelection ? LinkAt(p, sel.From.Index - 1) : null);
        string text = existing is { } cur0 ? TextOfRange(pl.Content, cur0.Start, cur0.End) : SelectedText();
        var bookmarks = BookmarkIndex.Build(body).All.Where(b => !b.Hidden).Select(b => b.Name).OrderBy(n => n).ToList();
        var dlg = new HyperlinkWindow(text, existing is { } e2 ? LinkTarget(e2.Link) : null, bookmarks,
            canEditText: existing is null && !HasSelection, canRemove: existing is not null) { Owner = this };
        if (dlg.ShowDialog() == true) ApplyLink(dlg.Address, dlg.Text, dlg.Remove);
        BackToPage();
    }

    /// <summary>
    /// Pose, modifie ou retire le lien de la sélection / du point d'insertion (Ctrl+K sans la boîte : auto-test).
    /// <paramref name="text"/> : texte inséré quand il n'y a pas de sélection.
    /// </summary>
    internal void ApplyLink(string? address, string text, bool remove)
    {
        if (Body is not { } body || _pkg?.Main is not { } main || Ordered() is not { } sel) return;
        if (!ReferenceEquals(sel.From.Paragraph, sel.To.Paragraph)) return;
        var p = sel.From.Paragraph;
        if (ToDocPos(sel.From) is not { } from || ToDocPos(sel.To) is not { } to) { NotEditable(); return; }
        var existing = LinkAt(p, sel.From.Index) ?? (sel.From.Index > 0 && !HasSelection ? LinkAt(p, sel.From.Index - 1) : null);
        var w = p.Props.Ns;
        string? relId = null, anchor = null;
        if (address is { Length: > 0 })
        {
            if (address.StartsWith('#')) anchor = address[1..];
            else
            {
                int hash = address.IndexOf('#');
                var url = hash > 0 && !address.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase) ? address[..hash] : address;
                if (hash > 0 && url != address) anchor = address[(hash + 1)..];
                if (!url.Contains(':')) url = (url.Contains('@') ? "mailto:" : "https://") + url;
                relId = _pkg.AddExternalRelationship(main.PartName, DocxPackage.HyperlinkType, url);
            }
        }
        var styleId = HyperlinkStyleId();
        var (anchorPos, caretPos) = SelectionDocPos();
        DocPos result = to;
        try
        {
            if (existing is { } cur)
            {
                var at = from.WithOffset(cur.Start);
                TextOps.Edit(_journal, body, remove ? "Supprimer le lien" : "Modifier le lien", at.ContainerPath, at.BlockIndex, 1, at, () =>
                {
                    // le lien est retrouvé dans le paragraphe courant (la transaction travaille sur ce bloc)
                    var para = at.Paragraph(body);
                    if (LinkOps.LinkAt(para, cur.Start) is not { } live) return 0;
                    if (remove) LinkOps.Unwrap(para, live.Link);
                    else LinkOps.Retarget(live.Link, w, relId, anchor);
                    return 0;
                }, _ => caretPos);
            }
            else
            {
                if (relId is null && anchor is null) return;
                result = TextOps.Edit(_journal, body, "Insérer un lien", from.ContainerPath, from.BlockIndex, 1, to, () =>
                {
                    var a = from;
                    var b = to;
                    if (a.Offset == b.Offset)
                    {
                        // pas de sélection : le texte à afficher est inséré
                        var label = text.Length > 0 ? text : address!;
                        b = TextOps.Insert(body, a, label);
                    }
                    var para = a.Paragraph(body);
                    var link = LinkOps.NewLink(w, relId, anchor);
                    LinkOps.Wrap(para, a.Offset, b.Offset, link);
                    foreach (var run in link.Runs())
                    {
                        if (styleId is not null) run.Props.SetVal("rStyle", styleId);
                        else { FormatOps.SetValue(run, "color", "0563C1"); FormatOps.SetUnderline(run, true); }
                    }
                    return b;
                }, x => x);
            }
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or InvalidOperationException)
        {
            StatusText.Text = $"Lien impossible ici : {ex.Message}";
            return;
        }
        if (existing is null) AfterTyping(result);
        else AfterFormat(remove ? "Lien supprimé" : "Lien modifié", anchorPos, caretPos);
    }

    static string TextOfRange(ParagraphContent c, int modelStart, int modelEnd)
    {
        var sb = new System.Text.StringBuilder();
        for (int k = 0; k < c.Text.Length; k++)
            if (c.ModelStart[k] >= modelStart && c.ModelEnd[k] <= modelEnd) sb.Append(c.Text[k]);
        return sb.ToString();
    }

    SymbolWindow? _symbols;

    /// <summary>Insertion ▸ Symbole : fenêtre non modale, chaque clic insère le caractère au point d'insertion.</summary>
    internal void ShowSymbols()
    {
        if (_layout is null) return;
        if (_symbols is null)
        {
            _symbols = new SymbolWindow(ch => ReplaceSelection(ch, "Symbole")) { Owner = this };
            _symbols.Closed += (_, _) => _symbols = null;
        }
        _symbols.Show();
        _symbols.Activate();
    }
}
