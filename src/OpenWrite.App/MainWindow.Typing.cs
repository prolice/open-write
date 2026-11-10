using System.Text;
using System.Windows;
using System.Windows.Input;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Layout;
using LayoutPagesView = OpenWrite.App.Layout.PagesView;
using Atom = OpenWrite.Layout.Atom;

namespace OpenWrite.App;

// Lot 3 (version 0.4) : saisie dans la vue des pages. Le point d'insertion et la sélection sont des positions dans le
// texte affiché (TextPoint : paragraphe + caractère), converties en positions du modèle (DocPos) pour chaque
// modification ; chaque modification passe par le journal d'annulation (TextOps) puis par la remise en page
// incrémentale (DocumentLayout.Update sur les blocs touchés). Navigation au clavier et à la souris de Word.
public partial class MainWindow
{
    /// <summary>Point d'insertion (extrémité mobile de la sélection).</summary>
    TextPoint? _caret;
    /// <summary>Autre extrémité de la sélection ; null : pas de sélection.</summary>
    TextPoint? _anchor;
    /// <summary>Abscisse visée par Haut / Bas (points de la page), gardée tant qu'on se déplace verticalement.</summary>
    double? _goalX;
    Func<string>? _newParaId;
    bool _mouseSelecting;
    TextPoint? _mouseDownAt;

    /// <summary>Branche la saisie sur la vue des pages (appelé par ShowDocument).</summary>
    void InitTyping(LayoutPagesView view)
    {
        view.Cursor = Cursors.IBeam;
        view.KeyDown += PagesView_KeyDown;
        view.TextInput += PagesView_TextInput;
        view.MouseLeftButtonDown += PagesView_MouseLeftButtonDown;
        view.Selection = SelectionRects;
        _newParaId = null;
        _anchor = null;
        _goalX = null;
        // point d'insertion gardé après un rechargement (annuler, rétablir) s'il désigne encore un paragraphe affiché
        if (_caret is not { } c || _layout is null || _layout.LinesOf(c.Paragraph).Count == 0) _caret = DocumentStart();
        Dispatcher.BeginInvoke(() => { view.Focus(); UpdateCaret(scroll: false); }, System.Windows.Threading.DispatcherPriority.Loaded);
    }

    internal TextPoint? CaretPoint => _caret;
    internal LayoutPagesView? PagesViewForTest => _pagesView;

    /// <summary>Banc de saisie : point d'insertion au début de la page <paramref name="page"/>, vue défilée jusque-là.</summary>
    internal void PlaceCaretOnPage(int page)
    {
        if (_layout is null || page < 0 || page >= _layout.Pages.Count || _layout.Pages[page].Lines.Count == 0) return;
        var l = _layout.Pages[page].Lines[0];
        GoToPage(page);
        MoveTo(new TextPoint(l.Paragraph.Paragraph, l.Paragraph.Lines[l.Line].Start), extend: false);
    }

    /// <summary>Capture « selection » : sélectionne quelques mots de la page (du 2e au 5e mot de la ligne la plus longue).</summary>
    internal void SelectSampleWords(int page)
    {
        if (_layout is null || page < 0 || page >= _layout.Pages.Count) return;
        var l = _layout.Pages[page].Lines.OrderByDescending(x => x.Paragraph.Lines[x.Line].End - x.Paragraph.Lines[x.Line].Start).FirstOrDefault();
        if (l.Paragraph is null) return;
        var text = l.Paragraph.Content.Text;
        int s = CaretGeometry.NextWord(text, l.Paragraph.Lines[l.Line].Start), e = s;
        for (int k = 0; k < 3; k++) e = CaretGeometry.NextWord(text, e);
        _anchor = new TextPoint(l.Paragraph.Paragraph, s);
        _caret = new TextPoint(l.Paragraph.Paragraph, Math.Max(s, e - 1));
        UpdateCaret(scroll: false);
    }

    /// <summary>Banc de saisie : corps répété n fois puis remis en page (documents de plusieurs centaines de pages).</summary>
    internal void RepeatBodyForBench(int n)
    {
        if (Body is not { } body || n <= 1) return;
        var copy = body.Blocks.Where(b => b is not Paragraph p || p.Props.Get("sectPr") is null).ToList();
        for (int r = 1; r < n; r++) body.Blocks.AddRange(copy.Select(b => b.Clone()));
        _caret = null;
        ShowDocument();
    }
    internal TextPoint? AnchorPoint => _anchor;

    // ================================================================ positions

    /// <summary>Ligne placée du point d'insertion et sa mise en page.</summary>
    (LineRef Ref, PlacedLine Line)? Locate(TextPoint t)
    {
        if (_layout is null) return null;
        var refs = _layout.LinesOf(t.Paragraph);
        if (refs.Count == 0) return null;
        var pl = _layout.Line(refs[0]).Paragraph;
        int line = CaretGeometry.LineOf(pl, Math.Min(t.Index, pl.Content.Text.Length), t.AtLineEnd);
        foreach (var r in refs)
        {
            var placed = _layout.Line(r);
            if (placed.Line == line) return (r, placed);
        }
        return (refs[0], _layout.Line(refs[0]));
    }

    ParagraphLayout? LayoutOf(Paragraph p) => _layout?.LinesOf(p) is { Count: > 0 } refs ? _layout.Line(refs[0]).Paragraph : null;

    /// <summary>Ordre du document entre deux points (ligne placée, puis caractère).</summary>
    int Compare(TextPoint a, TextPoint b)
    {
        if (ReferenceEquals(a.Paragraph, b.Paragraph)) return a.Index.CompareTo(b.Index);
        var la = Locate(a); var lb = Locate(b);
        if (la is null || lb is null) return 0;
        return la.Value.Ref.CompareTo(lb.Value.Ref);
    }

    (TextPoint From, TextPoint To)? Ordered()
    {
        if (_caret is not { } c) return null;
        if (_anchor is not { } a) return (c, c);
        return Compare(a, c) <= 0 ? (a, c) : (c, a);
    }

    bool HasSelection => _caret is { } c && _anchor is { } a && !(ReferenceEquals(a.Paragraph, c.Paragraph) && a.Index == c.Index);

    /// <summary>Position du modèle d'un point du texte affiché (paragraphes du corps seulement).</summary>
    DocPos? ToDocPos(TextPoint t)
    {
        if (Body is not { } body || LayoutOf(t.Paragraph) is not { } pl) return null;
        try { return DocPos.Of(body, t.Paragraph, CaretGeometry.ModelOffset(pl.Content, t.Index)); }
        catch (ArgumentException) { return null; }
    }

    /// <summary>Point du texte affiché d'une position du modèle (après une modification, une annulation).</summary>
    TextPoint? FromDocPos(DocPos pos)
    {
        if (Body is not { } body) return null;
        Paragraph p;
        try { p = pos.Paragraph(body); }
        catch (Exception ex) when (ex is ArgumentException or ArgumentOutOfRangeException) { return null; }
        if (LayoutOf(p) is not { } pl) return null;
        return new TextPoint(p, CaretGeometry.DisplayIndex(pl.Content, pos.Offset));
    }

    TextPoint? DocumentStart()
    {
        if (_layout is null) return null;
        for (int k = 0; k < _layout.Pages.Count; k++)
            if (_layout.Pages[k].Lines.Count > 0)
            {
                var l = _layout.Pages[k].Lines[0];
                return new TextPoint(l.Paragraph.Paragraph, l.Paragraph.Lines[l.Line].Start);
            }
        return null;
    }

    TextPoint? DocumentEnd()
    {
        if (_layout is null) return null;
        for (int k = _layout.Pages.Count - 1; k >= 0; k--)
            if (_layout.Pages[k].Lines.Count > 0)
            {
                var l = _layout.Pages[k].Lines[^1];
                return new TextPoint(l.Paragraph.Paragraph, l.Paragraph.Content.Text.Length);
            }
        return null;
    }

    /// <summary>Rectangle du point d'insertion (points de la page) : hauteur de la police du caractère qui précède.</summary>
    (int Page, Rect Rect)? CaretRect(TextPoint t)
    {
        if (Locate(t) is not { } loc) return null;
        var pl = loc.Line.Paragraph;
        var box = pl.Lines[loc.Line.Line];
        var c = pl.Content;
        int idx = Math.Clamp(t.Index, box.Start, box.End);
        double x = loc.Line.X + CaretGeometry.X(pl, loc.Line.Line, idx);
        var st = idx > 0 && idx - 1 < c.Text.Length && idx - 1 >= box.Start ? c.Styles[c.StyleOf[idx - 1]]
               : idx < c.Text.Length ? c.Styles[c.StyleOf[idx]] : c.Mark;
        double asc = st.Face.Ascent * st.Size, desc = st.Face.Descent * st.Size;
        double top = loc.Line.Y + Math.Max(0, box.Baseline - asc);
        double h = Math.Min(asc + desc, box.Height);
        return (loc.Ref.Page, new Rect(x, top, 1, Math.Max(4, h)));
    }

    /// <summary>Point du texte le plus proche d'un point d'une page (points) : ligne la plus proche, puis caractère.</summary>
    TextPoint? HitText(int page, Point at)
    {
        if (_layout is null || page < 0 || page >= _layout.Pages.Count) return null;
        PlacedLine? best = null;
        double bestScore = double.MaxValue;
        foreach (var l in _layout.Pages[page].Lines)
        {
            var box = l.Paragraph.Lines[l.Line];
            double dy = at.Y < l.Y ? l.Y - at.Y : at.Y > l.Y + box.Height ? at.Y - l.Y - box.Height : 0;
            double dx = at.X < l.X ? l.X - at.X : at.X > l.X + l.Paragraph.Width ? at.X - l.X - l.Paragraph.Width : 0;
            double score = dy * 1000 + dx;
            if (score < bestScore) { bestScore = score; best = l; }
        }
        if (best is not { } hit) return null;
        var b = hit.Paragraph.Lines[hit.Line];
        int idx = CaretGeometry.IndexAtX(hit.Paragraph, hit.Line, at.X - hit.X);
        return new TextPoint(hit.Paragraph.Paragraph, idx, idx == b.End && b.EndKind == LineEnd.Wrap && idx > b.Start);
    }

    /// <summary>Point du texte sous un point de la vue (hors des pages : page la plus proche).</summary>
    TextPoint? HitView(Point p)
    {
        if (_pagesView is null) return null;
        var (page, at) = _pagesView.HitPage(p);
        if (page < 0)
        {
            page = _pagesView.PageAt(p.Y);
            var r = _pagesView.PageRect(page);
            if (r.IsEmpty) return null;
            at = new Point((p.X - r.X) / _pagesView.Scale, (p.Y - r.Y) / _pagesView.Scale);
        }
        return HitText(page, at);
    }

    // ================================================================ affichage du point d'insertion

    /// <summary>Redessine le point d'insertion et la sélection ; fait défiler pour garder le point d'insertion visible.</summary>
    void UpdateCaret(bool scroll = true)
    {
        if (_pagesView is null) return;
        var rect = _caret is { } c ? CaretRect(c) : null;
        _pagesView.Caret = rect;
        _pagesView.DrawSelection();
        if (_caret is { } cc && LayoutOf(cc.Paragraph) is { } pl)
        {
            var ct = pl.Content;
            int i = Math.Min(cc.Index, ct.Text.Length);
            var rf = i > 0 ? ct.Styles[ct.StyleOf[i - 1]].Format : i < ct.Text.Length ? ct.Styles[ct.StyleOf[i]].Format : ct.Mark.Format;
            PaneSelection.Text = Describe(pl.Format, rf);
        }
        RefreshFormatPane();
        if (scroll && rect is { } cr) ScrollIntoView(cr.Page, cr.Rect);
    }

    void ScrollIntoView(int page, Rect rect)
    {
        if (_pagesView is null) return;
        var pr = _pagesView.PageRect(page);
        if (pr.IsEmpty) return;
        var s = _pagesView.Scale;
        var view = new Rect(pr.X + rect.X * s, pr.Y + rect.Y * s, Math.Max(1, rect.Width * s), rect.Height * s);
        Rect r;
        try { r = _pagesView.TransformToAncestor(Scroller).TransformBounds(view); }
        catch (InvalidOperationException) { return; }
        const double margin = 24;
        if (r.Top < margin) Scroller.ScrollToVerticalOffset(Scroller.VerticalOffset + r.Top - margin);
        else if (r.Bottom > Scroller.ViewportHeight - margin) Scroller.ScrollToVerticalOffset(Scroller.VerticalOffset + r.Bottom - Scroller.ViewportHeight + margin);
        if (r.Left < margin) Scroller.ScrollToHorizontalOffset(Scroller.HorizontalOffset + r.Left - margin);
        else if (r.Right > Scroller.ViewportWidth - margin) Scroller.ScrollToHorizontalOffset(Scroller.HorizontalOffset + r.Right - Scroller.ViewportWidth + margin);
    }

    /// <summary>Rectangles de la sélection sur une page (points), pour <see cref="LayoutPagesView.Selection"/>.</summary>
    IEnumerable<Rect> SelectionRects(int page)
    {
        if (!HasSelection || _layout is null || Ordered() is not { } sel) yield break;
        if (Locate(sel.From) is not { } a || Locate(sel.To) is not { } b) yield break;
        if (page < a.Ref.Page || page > b.Ref.Page) yield break;
        var lines = _layout.Pages[page].Lines;
        int first = page == a.Ref.Page ? a.Ref.Line : 0, last = page == b.Ref.Page ? b.Ref.Line : lines.Count - 1;
        for (int j = first; j <= last && j < lines.Count; j++)
        {
            var l = lines[j];
            var box = l.Paragraph.Lines[l.Line];
            bool isFirst = page == a.Ref.Page && j == a.Ref.Line, isLast = page == b.Ref.Page && j == b.Ref.Line;
            int s = isFirst ? Math.Clamp(sel.From.Index, box.Start, box.End) : box.Start;
            int e = isLast ? Math.Clamp(sel.To.Index, box.Start, box.End) : box.End;
            double x1 = CaretGeometry.X(l.Paragraph, l.Line, s), x2 = CaretGeometry.X(l.Paragraph, l.Line, e);
            // fin de paragraphe sélectionnée (la sélection continue au-delà) : marque de paragraphe
            if (!isLast && l.Line == l.Paragraph.Lines.Count - 1) x2 += 5;
            yield return new Rect(l.X + x1, l.Y, Math.Max(0, x2 - x1), box.Height);
        }
    }

    /// <summary>Texte sélectionné (presse-papiers) : paragraphes séparés par CR LF, retours à la ligne en \v comme Word.</summary>
    internal string SelectedText()
    {
        if (!HasSelection || _layout is null || Ordered() is not { } sel) return "";
        if (Locate(sel.From) is not { } a || Locate(sel.To) is not { } b) return "";
        var sb = new StringBuilder();
        ParagraphLayout? prev = null;
        for (LineRef? r = a.Ref; r is { } cur && cur.CompareTo(b.Ref) <= 0; r = _layout.Step(cur, 1))
        {
            var l = _layout.Line(cur);
            var box = l.Paragraph.Lines[l.Line];
            var text = l.Paragraph.Content.Text;
            if (prev is not null && !ReferenceEquals(prev, l.Paragraph)) sb.Append("\r\n");
            int s = cur.Equals(a.Ref) ? sel.From.Index : box.Start;
            int e = cur.Equals(b.Ref) ? sel.To.Index : box.EndKind is LineEnd.Wrap or LineEnd.Paragraph ? box.End : Math.Min(text.Length, box.End + 1);
            s = Math.Clamp(s, 0, text.Length);
            e = Math.Clamp(e, s, text.Length);
            foreach (var ch in text.AsSpan(s, e - s))
                if (ch == Atom.LineBreak) sb.Append('\v');
                else if (ch != Atom.Object && ch != Atom.ColumnBreak) sb.Append(ch);
            prev = l.Paragraph;
        }
        return sb.ToString();
    }

    // ================================================================ déplacements

    void MoveTo(TextPoint? t, bool extend, bool keepGoal = false)
    {
        if (t is null) return;
        if (extend) _anchor ??= _caret;
        else _anchor = null;
        _caret = t;
        _pendingFormat.Clear();
        if (!keepGoal) _goalX = null;
        UpdateCaret();
    }

    TextPoint? StepRight(TextPoint t, bool word)
    {
        if (LayoutOf(t.Paragraph) is not { } pl) return null;
        var text = pl.Content.Text;
        if (t.Index < text.Length) return new TextPoint(t.Paragraph, word ? CaretGeometry.NextWord(text, t.Index) : t.Index + 1);
        // fin du paragraphe : début du paragraphe suivant
        var refs = _layout!.LinesOf(t.Paragraph);
        for (LineRef? r = _layout.Step(refs[^1], 1); r is { } cur; r = _layout.Step(cur, 1))
        {
            var l = _layout.Line(cur);
            if (!ReferenceEquals(l.Paragraph.Paragraph, t.Paragraph)) return new TextPoint(l.Paragraph.Paragraph, l.Paragraph.Lines[l.Line].Start);
        }
        return t;
    }

    TextPoint? StepLeft(TextPoint t, bool word)
    {
        if (LayoutOf(t.Paragraph) is not { } pl) return null;
        var text = pl.Content.Text;
        if (t.Index > 0) return new TextPoint(t.Paragraph, word ? CaretGeometry.PreviousWord(text, t.Index) : t.Index - 1);
        var refs = _layout!.LinesOf(t.Paragraph);
        for (LineRef? r = _layout.Step(refs[0], -1); r is { } cur; r = _layout.Step(cur, -1))
        {
            var l = _layout.Line(cur);
            if (!ReferenceEquals(l.Paragraph.Paragraph, t.Paragraph))
            {
                int n = l.Paragraph.Content.Text.Length;
                return new TextPoint(l.Paragraph.Paragraph, word ? CaretGeometry.PreviousWord(l.Paragraph.Content.Text, n) : n);
            }
        }
        return t;
    }

    TextPoint? LineEdge(TextPoint t, bool end)
    {
        if (Locate(t) is not { } loc) return null;
        var box = loc.Line.Paragraph.Lines[loc.Line.Line];
        return end ? new TextPoint(t.Paragraph, box.End, box.EndKind == LineEnd.Wrap) : new TextPoint(t.Paragraph, box.Start);
    }

    /// <summary>Ligne au-dessus ou au-dessous, à l'abscisse visée.</summary>
    TextPoint? Vertical(TextPoint t, int delta)
    {
        if (_layout is null || Locate(t) is not { } loc || CaretRect(t) is not { } cr) return null;
        _goalX ??= cr.Rect.X;
        if (_layout.Step(loc.Ref, delta) is not { } target) return delta > 0 ? DocumentEnd() : DocumentStart();
        var l = _layout.Line(target);
        var box = l.Paragraph.Lines[l.Line];
        int idx = CaretGeometry.IndexAtX(l.Paragraph, l.Line, _goalX.Value - l.X);
        return new TextPoint(l.Paragraph.Paragraph, idx, idx == box.End && box.EndKind == LineEnd.Wrap && idx > box.Start);
    }

    /// <summary>Page précédente / suivante : même position dans la fenêtre, une hauteur de fenêtre plus loin.</summary>
    TextPoint? Paging(TextPoint t, int delta)
    {
        if (_pagesView is null || CaretRect(t) is not { } cr) return null;
        var pr = _pagesView.PageRect(cr.Page);
        var s = _pagesView.Scale;
        double step = Scroller.ViewportHeight / Math.Max(0.1, _zoom) * delta;
        _goalX ??= cr.Rect.X;
        var target = new Point(pr.X + _goalX.Value * s, pr.Y + (cr.Rect.Y + cr.Rect.Height / 2) * s + step);
        Scroller.ScrollToVerticalOffset(Scroller.VerticalOffset + step * _zoom);
        return HitView(target);
    }

    // ================================================================ clavier

    void PagesView_KeyDown(object sender, KeyEventArgs e)
    {
        if (_caret is not { } c || _layout is null) return;
        var mods = Keyboard.Modifiers;
        bool shift = mods.HasFlag(ModifierKeys.Shift), ctrl = mods.HasFlag(ModifierKeys.Control), alt = mods.HasFlag(ModifierKeys.Alt);
        if (alt) return;
        var key = e.Key == Key.System ? e.SystemKey : e.Key;
        if (FormattingKey(key, ctrl, shift)) { e.Handled = true; return; }
        bool handled = true;
        switch (key)
        {
            case Key.Left: MoveTo(HasSelection && !shift ? Ordered()!.Value.From : StepLeft(c, ctrl), shift); break;
            case Key.Right: MoveTo(HasSelection && !shift ? Ordered()!.Value.To : StepRight(c, ctrl), shift); break;
            case Key.Up: MoveTo(Vertical(c, -1), shift, keepGoal: true); break;
            case Key.Down: MoveTo(Vertical(c, 1), shift, keepGoal: true); break;
            case Key.Home: MoveTo(ctrl ? DocumentStart() : LineEdge(c, end: false), shift); break;
            case Key.End: MoveTo(ctrl ? DocumentEnd() : LineEdge(c, end: true), shift); break;
            case Key.PageUp: MoveTo(Paging(c, -1), shift, keepGoal: true); break;
            case Key.PageDown: MoveTo(Paging(c, 1), shift, keepGoal: true); break;
            case Key.Back: Backspace(ctrl); break;
            case Key.Delete when shift && !ctrl: Cut(); break;
            case Key.Delete: DeleteForward(ctrl); break;
            case Key.Enter:
                if (ctrl) InsertSpecial(shift ? SpecialChar.ColumnBreak : SpecialChar.PageBreak);
                else if (shift) InsertSpecial(SpecialChar.LineBreak);
                else NewParagraph();
                break;
            case Key.Tab when !ctrl: InsertSpecial(SpecialChar.Tab); break;
            case Key.Escape when HasSelection: MoveTo(c, extend: false); break;
            case Key.A when ctrl: SelectAll(); break;
            case Key.C when ctrl:
            case Key.Insert when ctrl: Copy(); break;
            case Key.X when ctrl: Cut(); break;
            case Key.V when ctrl: Paste(textOnly: shift); break;
            case Key.Insert when shift: Paste(); break;
            default: handled = false; break;
        }
        if (handled) e.Handled = true;
    }

    void PagesView_TextInput(object sender, TextCompositionEventArgs e)
    {
        var text = new string(e.Text.Where(ch => !char.IsControl(ch)).ToArray());
        if (text.Length == 0 || _caret is null) return;
        e.Handled = true;
        ReplaceSelection(text, "Frappe");
    }

    internal void SelectAll()
    {
        if (DocumentStart() is not { } s || DocumentEnd() is not { } end) return;
        _anchor = s;
        _caret = end;
        UpdateCaret(scroll: false);
    }

    // ================================================================ souris

    void PagesView_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (_pagesView is null) return;
        _pagesView.Focus();
        var p = e.GetPosition(_pagesView);
        if (HitView(p) is not { } t) return;
        // Ctrl+clic sur un lien : le suivre (comme Word)
        if (e.ClickCount == 1 && Keyboard.Modifiers.HasFlag(ModifierKeys.Control) && _pagesView.HitPage(p).Page >= 0
            && HitCharacter(p) is { } ch && FollowLink(ch)) { e.Handled = true; return; }
        if (e.ClickCount == 2 && LayoutOf(t.Paragraph) is { } pl)
        {
            // double-clic : le mot
            var (s, end) = CaretGeometry.WordAt(pl.Content.Text, t.Index);
            _anchor = new TextPoint(t.Paragraph, s);
            _caret = new TextPoint(t.Paragraph, end);
            UpdateCaret(scroll: false);
        }
        else if (e.ClickCount >= 3 && LayoutOf(t.Paragraph) is { } pl3)
        {
            // triple-clic : le paragraphe
            _anchor = new TextPoint(t.Paragraph, 0);
            _caret = new TextPoint(t.Paragraph, pl3.Content.Text.Length);
            UpdateCaret(scroll: false);
        }
        else
        {
            MoveTo(t, Keyboard.Modifiers.HasFlag(ModifierKeys.Shift));
            _mouseSelecting = true;
            _mouseDownAt = t;
            _pagesView.CaptureMouse();
        }
        e.Handled = true;
    }

    /// <summary>Glisser : étend la sélection jusqu'au pointeur.</summary>
    void DragSelect(MouseEventArgs e)
    {
        if (!_mouseSelecting || _pagesView is null || e.LeftButton != MouseButtonState.Pressed) return;
        if (HitView(e.GetPosition(_pagesView)) is not { } t) return;
        if (_caret is { } c && ReferenceEquals(c.Paragraph, t.Paragraph) && c.Index == t.Index) return;
        _anchor ??= _caret;
        _caret = t;
        UpdateCaret();
    }

    /// <summary>Fin d'un clic : vrai si le bouton a été relâché sans glisser (un clic simple sur une case la coche).</summary>
    bool EndMouseSelect()
    {
        bool click = _mouseSelecting && _caret is { } c && _mouseDownAt is { } d && ReferenceEquals(c.Paragraph, d.Paragraph) && c.Index == d.Index;
        _mouseSelecting = false;
        _mouseDownAt = null;
        _pagesView?.ReleaseMouseCapture();
        if (!HasSelection) _anchor = null;
        return click;
    }

    // ================================================================ modifications

    Func<string> NewParaId()
    {
        if (_newParaId is not null) return _newParaId;
        var ids = new List<string>();
        if (_pkg is not null)
            foreach (var part in _pkg.Stories.Values)
                foreach (var story in part.Stories())
                    foreach (var p in story.Paragraphs())
                        foreach (var a in p.Attributes)
                            if (a.Name == FlowEdits.W14 + "paraId" || a.Name == FlowEdits.W14 + "textId") ids.Add(a.Value);
        return _newParaId = FlowEdits.IdGenerator(ids, Environment.TickCount);
    }

    /// <summary>Remplace la sélection (ou insère au point d'insertion) par un texte ; une frappe simple se regroupe par mot.</summary>
    void ReplaceSelection(string text, string label)
    {
        if (Body is not { } body || Ordered() is not { } sel) return;
        if (ToDocPos(sel.From) is not { } from || ToDocPos(sel.To) is not { } to) { NotEditable(); return; }
        if (!from.ContainerPath.SequenceEqual(to.ContainerPath)) { StatusText.Text = "La sélection traverse un tableau : pas encore modifiable d'un bloc (version suivante)."; return; }
        DocPos pos;
        try
        {
            bool simple = from.Equals(to) && !text.Any(ch => ch is '\r' or '\n' or '\t' or '\v' or '\f');
            pos = !simple ? TextOps.Replace(_journal, body, label, from, to, text, NewParaId())
                : _pendingFormat.Count > 0 ? TypeWithPendingFormat(body, from, text)
                : TextOps.Type(_journal, body, from, text);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or InvalidOperationException)
        {
            StatusText.Text = $"Modification impossible ici : {ex.Message}";
            return;
        }
        AfterTyping(pos);
    }

    void NotEditable() => StatusText.Text = "Cette partie du document (en-tête, pied de page, note) n'est pas encore modifiable.";

    void DeleteSelection(string label = "Supprimer") => ReplaceSelection("", label);

    void Backspace(bool word)
    {
        if (HasSelection) { DeleteSelection(); return; }
        if (Body is not { } body || _caret is not { } c || LayoutOf(c.Paragraph) is not { } pl || ToDocPos(c) is not { } at) { NotEditable(); return; }
        var content = pl.Content;
        if (c.Index > 0)
        {
            int start = word ? CaretGeometry.ModelOffset(content, CaretGeometry.PreviousWord(content.Text, c.Index)) : CaretGeometry.ModelRange(content, c.Index - 1).Start;
            Delete(at.WithOffset(start), at);
            return;
        }
        // début de paragraphe : fusion avec le paragraphe précédent (même conteneur)
        var container = DocPos.ResolveContainer(body, at.ContainerPath);
        if (at.BlockIndex == 0 || container.Blocks[at.BlockIndex - 1] is not Paragraph prev) return;
        var prevPath = at.ContainerPath.Append(at.BlockIndex - 1).ToArray();
        Delete(new DocPos(prevPath, ParagraphAtoms.Length(prev)), at.WithOffset(0));
    }

    void DeleteForward(bool word)
    {
        if (HasSelection) { DeleteSelection(); return; }
        if (Body is not { } body || _caret is not { } c || LayoutOf(c.Paragraph) is not { } pl || ToDocPos(c) is not { } at) { NotEditable(); return; }
        var content = pl.Content;
        if (c.Index < content.Text.Length)
        {
            int end = word ? CaretGeometry.ModelOffset(content, CaretGeometry.NextWord(content.Text, c.Index)) : CaretGeometry.ModelRange(content, c.Index).End;
            Delete(at, at.WithOffset(end));
            return;
        }
        // fin de paragraphe : le paragraphe suivant (même conteneur) remonte
        var container = DocPos.ResolveContainer(body, at.ContainerPath);
        if (at.BlockIndex + 1 >= container.Blocks.Count || container.Blocks[at.BlockIndex + 1] is not Paragraph) return;
        var cur = at.Paragraph(body);
        Delete(at.WithOffset(ParagraphAtoms.Length(cur)), new DocPos(at.ContainerPath.Append(at.BlockIndex + 1).ToArray(), 0));
    }

    void Delete(DocPos from, DocPos to)
    {
        if (Body is not { } body || from.Equals(to)) return;
        try { AfterTyping(TextOps.DeleteRange(_journal, body, from, to)); }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or InvalidOperationException)
        {
            StatusText.Text = $"Suppression impossible ici : {ex.Message}";
        }
    }

    void NewParagraph()
    {
        if (HasSelection) { ReplaceSelection("\n", "Nouveau paragraphe"); return; }
        if (Body is not { } body || _caret is not { } c || ToDocPos(c) is not { } at) { NotEditable(); return; }
        AfterTyping(TextOps.SplitParagraph(_journal, body, at, NewParaId(), NextStyle));
    }

    /// <summary>Style suivant (w:next) d'un style de paragraphe : Entrée après un titre donne « Normal ».</summary>
    string? NextStyle(string? styleId)
    {
        if (_styles?.Sheet is not { } sheet) return null;
        var style = sheet.Style(styleId) ?? sheet.DefaultStyle("paragraph");
        return style?.Element(sheet.W + "next")?.Attribute(sheet.W + "val")?.Value;
    }

    void InsertSpecial(SpecialChar kind)
    {
        if (Body is not { } body || _caret is not { } c) return;
        if (kind == SpecialChar.Tab && ToDocPos(c) is { } p && p.Path.Count > 1 && body.Blocks[p.Path[0]] is Table)
        {
            StatusText.Text = "Tab dans un tableau : passage à la cellule suivante dans une prochaine version (Ctrl+Tab n'est pas encore disponible).";
            return;
        }
        if (HasSelection) { ReplaceSelection(kind switch { SpecialChar.Tab => "\t", SpecialChar.LineBreak => "\v", _ => "\f" }, "Saisie"); return; }
        if (ToDocPos(c) is not { } at) { NotEditable(); return; }
        if (kind == SpecialChar.ColumnBreak) kind = SpecialChar.PageBreak;      // colonnes : plus tard (lot 5)
        AfterTyping(TextOps.InsertSpecial(_journal, body, at, kind));
    }

    // ================================================================ après une modification

    /// <summary>Remise en page incrémentale des blocs touchés, point d'insertion à sa nouvelle place.</summary>
    void AfterTyping(DocPos caret)
    {
        Relayout();
        _anchor = null;
        _goalX = null;
        _caret = FromDocPos(caret) ?? _caret;
        Scroller.UpdateLayout();
        UpdateCaret();
        UpdateModifiedMarker();
        EnsureAutoSaveTimer();
        CommandManager.InvalidateRequerySuggested();
    }

    /// <summary>
    /// Remet en page après la dernière opération du journal : seuls les blocs touchés sont recalculés et la pagination
    /// s'arrête dès qu'elle retrouve une page d'avant (DocumentLayout.Update) ; dessins des pages refaits à la demande.
    /// </summary>
    void Relayout()
    {
        if (_layout is null || _drawings is null || _pagesView is null || Body is not { } body || _journal.LastChange is not { } ch || !ReferenceEquals(ch.Root, body))
        {
            var offset = Scroller.VerticalOffset;
            ShowDocument();
            Scroller.UpdateLayout();
            Scroller.ScrollToVerticalOffset(offset);
            return;
        }
        var sw = System.Diagnostics.Stopwatch.StartNew();
        int before = _layout.Pages.Count;
        var changed = ch.Container.Count == 0
            ? body.Blocks.GetRange(ch.Index, Math.Min(ch.Count, body.Blocks.Count - ch.Index))
            : new List<Block> { body.Blocks[ch.Container[0]] };
        _layout.Update(changed);
        LastUpdateMs = sw.Elapsed.TotalMilliseconds;
        _fieldsByBegin = null;
        // numéros de page (en-têtes, pieds) : tout redessiner si le nombre de pages change
        if (_layout.Pages.Count != before) _drawings.Clear(); else _drawings.Retain();
        _pagesView.InvalidateMeasure();
        _pagesView.InvalidateVisual();
        if (_thumbsView is not null) { _thumbsView.InvalidateMeasure(); _thumbsView.InvalidateVisual(); }
        UpdatePageStatus();
    }

    /// <summary>Durée de la dernière remise en page incrémentale (millisecondes).</summary>
    internal double LastUpdateMs { get; private set; }

    /// <summary>Titre de la fenêtre et nom du document : marque « modifié » (sans tout recalculer comme UpdateStatus).</summary>
    void UpdateModifiedMarker()
    {
        if (_info is not { } i || _path is null) return;
        var compat = i.CompatibilityMode < 15 ? " [Mode de compatibilité]" : "";
        Title = $"{(IsModified ? "• " : "")}{System.IO.Path.GetFileName(_path)}{compat} — OpenWrite";
        DocTitle.Text = System.IO.Path.GetFileName(_path) + (IsModified ? " · modifié" : "");
    }
}
