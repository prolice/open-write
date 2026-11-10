using System.Globalization;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Layout;

namespace OpenWrite.App;

// Lot 3 (version 0.4) : Rechercher / Remplacer (Ctrl+F, Ctrl+H) dans le texte affiché du corps (cellules comprises),
// dans l'ordre du document, à partir du point d'insertion, en reprenant au début. Un remplacement garde la mise en
// forme du texte remplacé ; « Remplacer tout » est une seule opération annulable.
public partial class MainWindow
{
    FindReplaceWindow? _findWindow;

    internal void ShowFind(bool replace)
    {
        if (_layout is null) return;
        if (_findWindow is null)
        {
            _findWindow = new FindReplaceWindow(this) { Owner = this };
            _findWindow.Closed += (_, _) => _findWindow = null;
        }
        var sel = SelectedText();
        _findWindow.Prepare(replace, sel.Length is > 0 and < 120 && !sel.Contains('\n') ? sel : null);
        _findWindow.Show();
        _findWindow.Activate();
    }

    /// <summary>Paragraphes du corps qui ont du texte affiché, dans l'ordre du document, avec leur contenu.</summary>
    IEnumerable<(Paragraph P, ParagraphContent C)> SearchableParagraphs()
    {
        if (Body is not { } body) yield break;
        foreach (var p in body.Paragraphs())
            if (LayoutOf(p) is { } pl && pl.Content.Text.Length > 0) yield return (p, pl.Content);
    }

    static int IndexOf(string text, string query, int start, bool matchCase, bool wholeWord)
    {
        var cmp = matchCase ? CompareOptions.None : CompareOptions.IgnoreCase | CompareOptions.IgnoreNonSpace;
        var ci = CultureInfo.CurrentCulture.CompareInfo;
        while (start <= text.Length - query.Length)
        {
            int i = ci.IndexOf(text, query, start, cmp);
            if (i < 0) return -1;
            bool ok = !wholeWord || ((i == 0 || !CaretGeometry.IsWordChar(text[i - 1])) && (i + query.Length >= text.Length || !CaretGeometry.IsWordChar(text[i + query.Length])));
            if (ok) return i;
            start = i + 1;
        }
        return -1;
    }

    /// <summary>Sélectionne l'occurrence suivante après la sélection (en reprenant au début) ; faux si aucune.</summary>
    internal bool FindNext(string query, bool matchCase, bool wholeWord)
    {
        if (query.Length == 0 || _layout is null) return false;
        var paras = SearchableParagraphs().ToList();
        if (paras.Count == 0) return false;
        int startPara = 0, startIndex = 0;
        if (Ordered() is { } sel)
        {
            int k = paras.FindIndex(x => ReferenceEquals(x.P, sel.To.Paragraph));
            if (k >= 0) { startPara = k; startIndex = sel.To.Index; }
        }
        for (int n = 0; n <= paras.Count; n++)
        {
            int k = (startPara + n) % paras.Count;
            var (p, c) = paras[k];
            int from = n == 0 ? startIndex : 0;
            int i = IndexOf(c.Text, query, from, matchCase, wholeWord);
            // dernier tour : seulement avant le point de départ
            if (i >= 0 && !(n == paras.Count && i >= startIndex))
            {
                _anchor = new TextPoint(p, i);
                _caret = new TextPoint(p, i + query.Length);
                _pendingFormat.Clear();
                UpdateCaret();
                return true;
            }
        }
        return false;
    }

    /// <summary>Remplace la sélection si c'est une occurrence, puis passe à la suivante.</summary>
    internal void ReplaceCurrent(string query, string replacement, bool matchCase, bool wholeWord)
    {
        if (HasSelection && Ordered() is { } sel && ReferenceEquals(sel.From.Paragraph, sel.To.Paragraph) && LayoutOf(sel.From.Paragraph) is { } pl
            && IndexOf(pl.Content.Text, query, sel.From.Index, matchCase, wholeWord) == sel.From.Index && sel.To.Index - sel.From.Index == query.Length
            && Body is { } body && ToRange(pl.Content, sel.From.Paragraph, sel.From.Index, query.Length) is { } range)
        {
            var pos = TextOps.Edit(_journal, body, "Remplacer", range.From.ContainerPath, range.From.BlockIndex, 1, range.To,
                () => TextOps.ReplaceText(body, range.From, range.To, replacement), x => x);
            AfterTyping(pos);
        }
        FindNext(query, matchCase, wholeWord);
    }

    /// <summary>Plage du modèle d'une occurrence (du début du 1er caractère à la fin du dernier).</summary>
    (DocPos From, DocPos To)? ToRange(ParagraphContent c, Paragraph p, int index, int length)
    {
        if (Body is not { } body) return null;
        try
        {
            var at = DocPos.Of(body, p);
            return (at.WithOffset(c.ModelStart[index]), at.WithOffset(c.ModelEnd[index + length - 1]));
        }
        catch (ArgumentException) { return null; }
    }

    /// <summary>Remplace toutes les occurrences (une seule opération annulable) ; renvoie leur nombre.</summary>
    internal int ReplaceAll(string query, string replacement, bool matchCase, bool wholeWord)
    {
        if (query.Length == 0 || Body is not { } body) return 0;
        var ranges = new List<(DocPos From, DocPos To)>();
        foreach (var (p, c) in SearchableParagraphs())
        {
            DocPos at;
            try { at = DocPos.Of(body, p); }
            catch (ArgumentException) { continue; }
            for (int i = IndexOf(c.Text, query, 0, matchCase, wholeWord); i >= 0; i = IndexOf(c.Text, query, i + query.Length, matchCase, wholeWord))
                ranges.Add((at.WithOffset(c.ModelStart[i]), at.WithOffset(c.ModelEnd[i + query.Length - 1])));
        }
        if (ranges.Count == 0) return 0;
        var caret = _caret is { } cc ? ToDocPos(cc) : null;
        TextOps.Edit(_journal, body, $"Remplacer tout ({ranges.Count})", Array.Empty<int>(), 0, body.Blocks.Count, caret, () =>
        {
            // de la fin vers le début : les positions des occurrences précédentes restent valables
            for (int k = ranges.Count - 1; k >= 0; k--) TextOps.ReplaceText(body, ranges[k].From, ranges[k].To, replacement);
            return 0;
        }, _ => caret);
        Relayout();
        _anchor = null;
        if (caret is not null && FromDocPos(caret) is { } t) _caret = t;
        else _caret = DocumentStart();
        UpdateCaret(scroll: false);
        UpdateModifiedMarker();
        EnsureAutoSaveTimer();
        return ranges.Count;
    }
}
