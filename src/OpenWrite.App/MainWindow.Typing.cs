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
        view.MouseRightButtonDown += (_, e) =>
        {
            // clic droit hors de la sélection : le point d'insertion vient sous le pointeur (Word)
            if (HitView(e.GetPosition(view)) is { } t && !(HasSelection && Ordered() is { } sel && Compare(sel.From, t) <= 0 && Compare(t, sel.To) <= 0))
                MoveTo(t, extend: false);
            view.Focus();
        };
        view.ContextMenuOpening += (_, _) => view.ContextMenu = BuildContextMenu();
        view.Selection = SelectionRects;
        view.ObjectSelection = ObjectOverlay;
        _objSel = null;
        _box = null;
        _newParaId = null;
        _anchor = null;
        _goalX = null;
        // point d'insertion gardé après un rechargement (annuler, rétablir) s'il désigne encore un paragraphe affiché
        if (_caret is not { } c || _layout is null || Src.LinesOf(c.Paragraph).Count == 0) _caret = DocumentStart();
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

    /// <summary>Auto-test : sélection de <paramref name="anchor"/> à <paramref name="caret"/>.</summary>
    internal void Select(TextPoint anchor, TextPoint caret)
    {
        _anchor = anchor;
        _caret = caret;
        UpdateCaret(scroll: false);
    }

    // ================================================================ positions

    /// <summary>Ligne placée du point d'insertion et sa mise en page.</summary>
    (LineRef Ref, PlacedLine Line)? Locate(TextPoint t)
    {
        if (_layout is null) return null;
        var refs = Src.LinesOf(t.Paragraph);
        if (refs.Count == 0) return null;
        var pl = Src.Line(refs[0]).Paragraph;
        int line = CaretGeometry.LineOf(pl, Math.Min(t.Index, pl.Content.Text.Length), t.AtLineEnd);
        foreach (var r in refs)
        {
            var placed = Src.Line(r);
            if (placed.Line == line) return (r, placed);
        }
        return (refs[0], Src.Line(refs[0]));
    }

    ParagraphLayout? LayoutOf(Paragraph p) => Src.LinesOf(p) is { Count: > 0 } refs ? Src.Line(refs[0]).Paragraph : null;

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
        if (EditRoot is not { } body || LayoutOf(t.Paragraph) is not { } pl) return null;
        try { return DocPos.Of(body, t.Paragraph, CaretGeometry.ModelOffset(pl.Content, t.Index)); }
        catch (ArgumentException) { return null; }
    }

    /// <summary>Point du texte affiché d'une position du modèle (après une modification, une annulation).</summary>
    TextPoint? FromDocPos(DocPos pos)
    {
        if (EditRoot is not { } body) return null;
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
            if (Src.PageLines(k).Count > 0)
            {
                var l = Src.PageLines(k)[0];
                return new TextPoint(l.Paragraph.Paragraph, l.Paragraph.Lines[l.Line].Start);
            }
        return null;
    }

    TextPoint? DocumentEnd()
    {
        if (_layout is null) return null;
        for (int k = _layout.Pages.Count - 1; k >= 0; k--)
            if (Src.PageLines(k).Count > 0)
            {
                var l = Src.PageLines(k)[^1];
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
        foreach (var l in Src.PageLines(page))
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
        ScheduleCaretUi();
        if (scroll && rect is { } cr) ScrollIntoView(cr.Page, cr.Rect);
    }

    bool _caretUiPending;

    /// <summary>
    /// Volet Mise en forme et règle : mis à jour une fois la rafale de touches traitée (priorité de fond), et seulement
    /// si ce qu'ils montrent a changé — la frappe reste à quelques millisecondes.
    /// </summary>
    void ScheduleCaretUi()
    {
        if (_caretUiPending) return;
        _caretUiPending = true;
        Dispatcher.BeginInvoke(() =>
        {
            _caretUiPending = false;
            if (_caret is { } cc && LayoutOf(cc.Paragraph) is { } pl && CaretFormat() is { } rf)
            {
                var text = Describe(pl.Format, rf);
                if (PaneSelection.Text != text) PaneSelection.Text = text;
            }
            RefreshFormatPane();
            UpdateRuler();
            UpdateTableTabs();
        }, System.Windows.Threading.DispatcherPriority.Background);
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
        // d'une cellule à une autre : rectangle de cellules (Word)
        if (CellSelection() is { } cells)
        {
            foreach (var r in CellSelectionRects(page, cells)) yield return r;
            yield break;
        }
        if (Locate(sel.From) is not { } a || Locate(sel.To) is not { } b) yield break;
        if (page < a.Ref.Page || page > b.Ref.Page) yield break;
        var lines = Src.PageLines(page);
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
        if (CellSelection() is { } cells) return CellSelectionText(cells);
        if (Locate(sel.From) is not { } a || Locate(sel.To) is not { } b) return "";
        var sb = new StringBuilder();
        ParagraphLayout? prev = null;
        for (LineRef? r = a.Ref; r is { } cur && cur.CompareTo(b.Ref) <= 0; r = Src.Step(cur, 1))
        {
            var l = Src.Line(cur);
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
        var refs = Src.LinesOf(t.Paragraph);
        for (LineRef? r = Src.Step(refs[^1], 1); r is { } cur; r = Src.Step(cur, 1))
        {
            var l = Src.Line(cur);
            if (!ReferenceEquals(l.Paragraph.Paragraph, t.Paragraph)) return new TextPoint(l.Paragraph.Paragraph, l.Paragraph.Lines[l.Line].Start);
        }
        return t;
    }

    TextPoint? StepLeft(TextPoint t, bool word)
    {
        if (LayoutOf(t.Paragraph) is not { } pl) return null;
        var text = pl.Content.Text;
        if (t.Index > 0) return new TextPoint(t.Paragraph, word ? CaretGeometry.PreviousWord(text, t.Index) : t.Index - 1);
        var refs = Src.LinesOf(t.Paragraph);
        for (LineRef? r = Src.Step(refs[0], -1); r is { } cur; r = Src.Step(cur, -1))
        {
            var l = Src.Line(cur);
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
        if (Src.Step(loc.Ref, delta) is not { } target) return delta > 0 ? DocumentEnd() : DocumentStart();
        var l = Src.Line(target);
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
        var key = e.Key == Key.System ? e.SystemKey : e.Key;
        // Ctrl+Alt+1, 2, 3 : Titre 1, 2, 3 (Word)
        if (ctrl && alt && key is Key.D1 or Key.D2 or Key.D3 or Key.NumPad1 or Key.NumPad2 or Key.NumPad3)
        {
            ApplyStyle(key is Key.D1 or Key.NumPad1 ? "heading 1" : key is Key.D2 or Key.NumPad2 ? "heading 2" : "heading 3");
            e.Handled = true;
            return;
        }
        if (alt) return;
        // objet sélectionné : Suppr, flèches, Échap ; zone de texte : Échap la sélectionne
        if (_objSel is not null)
        {
            if (ObjectKey(key, ctrl)) { e.Handled = true; return; }
            if (ctrl && key is Key.C or Key.X) { if (key == Key.X) { if (CopySelectedObject()) DeleteSelectedObject(); } else CopySelectedObject(); e.Handled = true; return; }
            if (!ctrl && key is not (Key.LeftShift or Key.RightShift or Key.LeftCtrl or Key.RightCtrl)) SelectObject(null);
        }
        if (_box is not null && key == Key.Escape) { ExitTextBox(select: true); e.Handled = true; return; }
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
            // Retour arrière au début d'un élément de liste : il sort de la liste (Word)
            case Key.Back when !ctrl && !HasSelection && c.Index == 0 && CaretInList(): EndList(); break;
            case Key.Back: Backspace(ctrl); break;
            case Key.Delete when shift && !ctrl: Cut(); break;
            case Key.Delete: DeleteForward(ctrl); break;
            // Ctrl+Maj+Entrée dans un tableau : le fractionner au-dessus de la ligne (Word)
            case Key.Enter when ctrl && shift && InTableCell(): SplitTableHere(); break;
            case Key.Enter:
                if (ctrl) InsertSpecial(shift ? SpecialChar.ColumnBreak : SpecialChar.PageBreak);
                else if (shift) InsertSpecial(SpecialChar.LineBreak);
                else NewParagraph();
                break;
            // dans un tableau : cellule suivante (Maj : précédente) ; Ctrl+Tab y insère une vraie tabulation
            // au début d'un élément de liste (ou sur plusieurs éléments) : niveau suivant / précédent
            case Key.Tab when !ctrl && CaretInList() && (c.Index == 0 || HasSelection && !ReferenceEquals(_anchor!.Value.Paragraph, c.Paragraph)):
                ChangeListLevel(shift ? -1 : 1);
                break;
            case Key.Tab when !ctrl && InTableCell(): MoveToCell(shift ? -1 : 1); break;
            case Key.N when ctrl && shift: ApplyStyle("Normal"); break;
            case Key.L when ctrl && shift: ToggleList(Core.Styles.ListKind.Bullet); break;
            case Key.Tab: InsertSpecial(SpecialChar.Tab); break;
            // Ctrl+Maj+Espace : espace insécable ; Ctrl+Maj+- : trait d'union insécable (Word)
            case Key.Space when ctrl && shift: ReplaceSelection("\u00A0", "Espace insécable"); break;
            case Key.OemMinus when ctrl && shift:
            case Key.Subtract when ctrl && shift: ReplaceSelection("\u2011", "Trait d'union insécable"); break;
            case Key.Escape when HasSelection: MoveTo(c, extend: false); break;
            case Key.Escape when _hf is not null: ExitHeader(); break;
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
        // en-têtes et pieds : double-clic dans les marges pour y entrer, dans le corps pour en sortir ; en mode en-tête,
        // un clic simple ne quitte pas l'en-tête (comme Word)
        var (hitPage, hitAt) = _pagesView.HitPage(p);
        if (e.ClickCount == 2 && hitPage >= 0 && HeaderDoubleClick(hitPage, hitAt)) { e.Handled = true; return; }
        if (_hf is not null && hitPage >= 0 && _layout is { } hl)
        {
            var hp = hl.Pages[hitPage];
            bool inMargins = hitAt.Y < hp.ContentTop || hitAt.Y > hp.ContentBottom + hp.ReservedBottom;
            if (!inMargins) { e.Handled = true; return; }
            var kind = hitAt.Y < hp.ContentTop ? "header" : "footer";
            if (hitPage != _hf.Page || kind != _hf.Kind) { EnterHeader(hitPage, kind, hitAt); e.Handled = true; return; }
        }
        // objets : poignées, objet sous le pointeur, texte d'une zone de texte
        if (hitPage >= 0 && ObjectMouseDown(hitPage, hitAt, e)) { e.Handled = true; return; }
        // bordure de cellule : redimensionner la colonne ou la ligne
        if (e.ClickCount == 1 && _hf is null && _box is null && hitPage >= 0 && StartBorderDrag(hitPage, hitAt)) { e.Handled = true; return; }
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
        if (EditRoot is not { } body || Ordered() is not { } sel) return;
        if (CellSelection() is { } cells) { ReplaceCells(cells, text.Replace("\r\n", "\n"), label); return; }
        if (ToDocPos(sel.From) is not { } from || ToDocPos(sel.To) is not { } to) { NotEditable(); return; }
        if (!SupportedRange(from, to)) return;
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

    /// <summary>
    /// Plage modifiable d'un bloc : dans un même conteneur, ou à travers des tableaux au niveau du corps (tableaux touchés
    /// pris en entier), ou d'une cellule à une autre d'un même tableau (rectangle de cellules, lot 6).
    /// </summary>
    bool SupportedRange(DocPos from, DocPos to, bool quiet = false)
    {
        if (from.ContainerPath.SequenceEqual(to.ContainerPath) || TextOps.AcrossTopLevel(from, to) is not null || CellSelection() is not null) return true;
        if (!quiet) StatusText.Text = "Cette sélection (d'un tableau imbriqué vers l'extérieur) ne peut pas être modifiée d'un bloc.";
        return false;
    }

    void NotEditable() => StatusText.Text = "Cette partie du document (en-tête, pied de page, note) n'est pas encore modifiable.";

    void DeleteSelection(string label = "Supprimer") => ReplaceSelection("", label);

    void Backspace(bool word)
    {
        if (HasSelection) { DeleteSelection(); return; }
        if (EditRoot is not { } body || _caret is not { } c || LayoutOf(c.Paragraph) is not { } pl || ToDocPos(c) is not { } at) { NotEditable(); return; }
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
        if (EditRoot is not { } body || _caret is not { } c || LayoutOf(c.Paragraph) is not { } pl || ToDocPos(c) is not { } at) { NotEditable(); return; }
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
        if (EditRoot is not { } body || from.Equals(to)) return;
        try { AfterTyping(TextOps.DeleteRange(_journal, body, from, to)); }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or InvalidOperationException)
        {
            StatusText.Text = $"Suppression impossible ici : {ex.Message}";
        }
    }

    void NewParagraph()
    {
        if (HasSelection) { ReplaceSelection("\n", "Nouveau paragraphe"); return; }
        // Entrée sur un élément de liste vide : fin de la liste (Word)
        if (CaretInList() && _caret is { } lc && LayoutOf(lc.Paragraph) is { } lpl && lpl.Content.Text.Length == 0) { EndList(); return; }
        if (EditRoot is not { } body || _caret is not { } c || ToDocPos(c) is not { } at) { NotEditable(); return; }
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
        if (EditRoot is not { } body || _caret is not { } c) return;
        if (HasSelection) { ReplaceSelection(kind switch { SpecialChar.Tab => "\t", SpecialChar.LineBreak => "\v", _ => "\f" }, "Saisie"); return; }
        if (ToDocPos(c) is not { } at) { NotEditable(); return; }
        AfterTyping(TextOps.InsertSpecial(_journal, body, at, kind));
    }

    // ================================================================ tableaux

    /// <summary>Le point d'insertion est-il dans une cellule de tableau ?</summary>
    bool InTableCell() => CellOf(_caret) is not null;

    /// <summary>Tableau, ligne et cellule qui contiennent un point (chemin : …, tableau, ligne, cellule, paragraphe).</summary>
    (Table Table, int Row, int Cell, int[] TablePath)? CellOf(TextPoint? t)
    {
        if (EditRoot is not { } body || t is null || ToDocPos(t.Value) is not { } pos || pos.Path.Count < 4) return null;
        var path = pos.Path;
        var prefix = path.Take(path.Count - 4).ToArray();
        Story parent;
        try { parent = DocPos.ResolveContainer(body, prefix); }
        catch (ArgumentException) { return null; }
        return parent.Blocks[path[^4]] is Table table ? (table, path[^3], path[^2], prefix.Append(path[^4]).ToArray()) : null;
    }

    /// <summary>
    /// Tab / Maj+Tab dans un tableau : sélectionne le contenu de la cellule suivante ou précédente (ligne suivante au bout
    /// de la ligne), comme Word ; Tab dans la dernière cellule ajoute une ligne.
    /// </summary>
    void MoveToCell(int direction)
    {
        if (EditRoot is not { } body || CellOf(_caret) is not { } at) return;
        // cellules du tableau dans l'ordre : (ligne, élément)
        var cells = new List<(int Row, int Item, TableCell Cell)>();
        for (int r = 0; r < at.Table.Items.Count; r++)
            if (at.Table.Items[r] is TableRow row)
                for (int c = 0; c < row.Items.Count; c++)
                    if (row.Items[c] is TableCell cell) cells.Add((r, c, cell));
        int k = cells.FindIndex(x => x.Row == at.Row && x.Item == at.Cell) + direction;
        if (k >= cells.Count)
        {
            // Tab dans la dernière cellule : nouvelle ligne (Word)
            _anchor = null;
            EditTable("Nouvelle ligne", (ctx, _) =>
            {
                int row = TableOps.InsertRows(ctx.Table, ctx.Map.Rows.Count - 1, below: true, 1, NewParaId());
                return CellStart(ctx.Table, row, 0);
            });
            return;
        }
        if (k < 0) { StatusText.Text = "Première cellule du tableau."; return; }
        var target = cells[k];
        var paras = target.Cell.Content.Paragraphs().ToList();
        if (paras.Count == 0) return;
        var first = FromDocPos(DocPos.Of(body, paras[0]));
        var lastP = paras[^1];
        var last = LayoutOf(lastP) is { } pl ? new TextPoint(lastP, pl.Content.Text.Length) : (TextPoint?)null;
        if (first is null || last is null) return;
        _anchor = first;
        _caret = last;
        _goalX = null;
        _pendingFormat.Clear();
        UpdateCaret();
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
        if (RelayoutTextBox()) return;
        if (RelayoutHeader()) return;
        // styles ou numérotation modifiés : résolveur refait, document entier remis en page
        bool stylesChanged = _journal.LastXml.Count > 0;
        bool additive = _xmlAdditive;
        _xmlAdditive = false;
        if (stylesChanged && _pkg is not null)
        {
            _styles = new Core.Styles.StyleResolver(_pkg.StyleSheet());
            _paneState = null;
            Dispatcher.BeginInvoke(UpdateStyleGallery, System.Windows.Threading.DispatcherPriority.Background);
            // styles ou listes seulement ajoutés : les paragraphes qui ne les utilisent pas sont inchangés
            if (additive && _layout is not null) { _layout.ReplaceStyles(_styles); stylesChanged = false; }
        }
        if (stylesChanged || _layout is null || _drawings is null || _pagesView is null || Body is not { } body || _journal.LastChange is not { } ch || !ReferenceEquals(ch.Root, body))
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
        // modification dans une cellule : seuls ses paragraphes touchés sont à recouper en lignes (le tableau est remesuré)
        List<Block>? dirty = null;
        if (ch.Container.Count > 0)
        {
            try
            {
                var inner = DocPos.ResolveContainer(body, ch.Container);
                dirty = inner.Blocks.GetRange(ch.Index, Math.Min(ch.Count, inner.Blocks.Count - ch.Index));
            }
            catch (ArgumentException) { dirty = null; }
        }
        _layout.Update(changed, rebuildFlow: _listsTouched || changed.OfType<Paragraph>().Any(p => p.Props.Get("numPr") is not null), dirty);
        _listsTouched = false;
        LastUpdateMs = sw.Elapsed.TotalMilliseconds;
        _fieldsByBegin = null;
        // numéros de page (en-têtes, pieds) : tout redessiner si le nombre de pages change
        if (_layout.Pages.Count != before) _drawings.Clear(); else _drawings.Retain();
        _pagesView.InvalidateMeasure();
        _pagesView.InvalidateVisual();
        if (_thumbsView is not null) { _thumbsView.InvalidateMeasure(); _thumbsView.InvalidateVisual(); }
        UpdatePageStatus();
    }

    /// <summary>
    /// La dernière opération n'a fait qu'ajouter des styles ou des listes (pas modifié d'existants) : remise en page
    /// incrémentale possible. Mis par l'opération, consommé par <see cref="Relayout"/> (annuler / rétablir : complète).
    /// </summary>
    bool _xmlAdditive;

    /// <summary>La dernière opération a changé des listes : numéros à recalculer (pas de chemin rapide).</summary>
    bool _listsTouched;

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
