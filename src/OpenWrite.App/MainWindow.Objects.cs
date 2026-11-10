using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using OpenWrite.Core.Drawing;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Layout;
using Page = OpenWrite.Layout.Page;
using Atom = OpenWrite.Layout.Atom;

namespace OpenWrite.App;

/// <summary>Lignes d'une zone de texte en cours d'édition (celles de l'objet flottant sur sa page).</summary>
sealed class BoxLines : ILineSource
{
    readonly int _page;
    readonly List<PlacedLine> _lines;

    public BoxLines(int page, FloatBox box, int index)
    {
        _page = page;
        _lines = box.TextLines.Where(t => t.Box == index).Select(t => t.Line).ToList();
    }

    public int Page => _page;
    public List<LineRef> LinesOf(Paragraph p) =>
        Enumerable.Range(0, _lines.Count).Where(k => ReferenceEquals(_lines[k].Paragraph.Paragraph, p)).Select(k => new LineRef(_page, k)).ToList();
    public PlacedLine Line(LineRef r) => _lines[r.Line];
    public LineRef? Step(LineRef r, int delta) => r.Line + delta is var k && k >= 0 && k < _lines.Count ? new LineRef(_page, k) : null;
    public IReadOnlyList<PlacedLine> PageLines(int page) => page == _page ? _lines : Array.Empty<PlacedLine>();
}

// Lot 7 (version 1.0) : objets graphiques — sélection au clic (poignées de taille, de rotation, ancre), déplacement et
// redimensionnement à la souris (une étape annulable chacun), clavier (Suppr, flèches, Échap), habillage, ordre,
// rotation, insertion d'images (fichier, presse-papiers), de formes et de zones de texte, édition du texte d'une zone.
public partial class MainWindow
{
    /// <summary>Objet sélectionné : dessin, paragraphe d'ancrage, page, cadre (points de la page), objet flottant ou en ligne.</summary>
    sealed record ObjSel(DrawingContent Drawing, Paragraph Anchor, int Page, Rect Frame, double Rotation, FloatBox? Float, DrawingObject Object);
    ObjSel? _objSel;

    enum DragKind { Move, Resize, Rotate }
    /// <summary>Glisser en cours sur l'objet sélectionné : genre, poignée (0..7 : coins et côtés), point de départ (page).</summary>
    (DragKind Kind, int Handle, Point Start)? _objDrag;
    Rect? _objGhost;
    double _ghostRotation;

    /// <summary>Zone de texte en cours d'édition : dessin, indice de la zone, histoire, page, lignes.</summary>
    sealed record BoxEdit(DrawingContent Drawing, int Box, Story Story, int Page, BoxLines Lines);
    BoxEdit? _box;

    internal bool HasObjectSelection => _objSel is not null;
    internal DrawingContent? SelectedDrawing => _objSel?.Drawing;
    internal bool InTextBox => _box is not null;
    internal Rect? SelectedFrame => _objSel?.Frame;

    /// <summary>Banc d'essai : déplace l'objet sélectionné comme un glisser à la souris.</summary>
    internal void MoveSelectedObjectForTest(double dx, double dy)
    {
        if (_objSel is { } s) MoveObject(s, s.Frame.X + dx, s.Frame.Y + dy);
    }

    /// <summary>Banc d'essai : taille de l'objet sélectionné comme avec une poignée.</summary>
    internal void ResizeSelectedObjectForTest(double width, double height)
    {
        if (_objSel is { } s) EditObject("Taille de l'objet", s, d => DrawingOps.SetSize(d, width, height));
    }

    /// <summary>Banc d'essai : sélectionne le premier objet de la page sous le point (points de la page).</summary>
    internal bool ClickObjectForTest(int page, Point at)
    {
        if (ObjectAt(page, at) is not { } s) return false;
        if (s.Float is { } f && TextBoxIndex(s.Object) >= 0 && InTextArea(f, at)) EnterTextBox(s, at);
        else SelectObject(s);
        return true;
    }

    // ================================================================ repérage

    /// <summary>Objets de la page sous un point : flottants devant le texte (le plus haut d'abord), en ligne, derrière le texte.</summary>
    ObjSel? ObjectAt(int page, Point at)
    {
        if (_layout is null || page < 0 || page >= _layout.Pages.Count) return null;
        var pg = _layout.Pages[page];
        // en-tête en cours d'édition : ses objets seulement ; sinon ceux du corps
        string? part = _hf?.RelId;
        IEnumerable<FloatBox> Floats(bool behind) => pg.Floats.Where(f => f.PartId == part && f.Anchor is not null == (part is null) && f.Object.BehindText == behind)
            .OrderByDescending(f => f.Object.Z);
        foreach (var f in Floats(false))
            if (Contains(f, at)) return FromFloat(f, page);
        foreach (var l in Src.PageLines(page))
        {
            var c = l.Paragraph.Content;
            var lb = l.Paragraph.Lines[l.Line];
            for (int k = lb.Start; k < lb.End && k < c.Text.Length; k++)
                if (c.Text[k] == Atom.Object && c.Drawings.TryGetValue(k, out var obj) && InlineRect(l, k) is { } r && r.Contains(at))
                    return new ObjSel(obj.Source, l.Paragraph.Paragraph, page, r, obj.Root.Rotation, null, obj);
        }
        // derrière le texte : seulement hors des lignes de texte (comme Word)
        bool onText = Src.PageLines(page).Any(l => at.Y >= l.Y && at.Y <= l.Y + l.Paragraph.Lines[l.Line].Height
            && at.X >= l.X + l.Paragraph.Lines[l.Line].Pieces.FirstOrDefault().X && at.X <= l.X + l.Paragraph.Lines[l.Line].Ink + 2);
        if (!onText)
            foreach (var f in Floats(true))
                if (Contains(f, at)) return FromFloat(f, page);
        return null;
    }

    static bool Contains(FloatBox f, Point at)
    {
        var r = new Rect(f.X, f.Y, Math.Max(f.W, 2), Math.Max(f.H, 2));
        r.Inflate(2, 2);                                   // traits : attrapés à 2 points près
        return r.Contains(at);
    }

    ObjSel FromFloat(FloatBox f, int page) =>
        new(f.Object.Source, f.Anchor!, page, new Rect(f.X, f.Y, f.W, f.H), f.Object.Root.Rotation, f, f.Object);

    /// <summary>Cadre d'un objet en ligne (points de la page) : sur la ligne de base, marges d'effet retirées.</summary>
    static Rect? InlineRect(PlacedLine l, int k)
    {
        var c = l.Paragraph.Content;
        if (!c.Objects.TryGetValue(k, out var size) || !c.Drawings.TryGetValue(k, out var obj)) return null;
        var lb = l.Paragraph.Lines[l.Line];
        double x = l.X + CaretGeometry.X(l.Paragraph, l.Line, k);
        double baseline = l.Y + lb.Baseline;
        return new Rect(x + obj.Effect.L, baseline - size.H + obj.Effect.T, Math.Max(0.5, obj.Width), Math.Max(0.5, obj.Height));
    }

    /// <summary>Retrouve la sélection après une remise en page (même dessin) ; null s'il n'est plus affiché.</summary>
    ObjSel? Reselect(DrawingContent d)
    {
        if (_layout is null) return null;
        for (int p = 0; p < _layout.Pages.Count; p++)
        {
            foreach (var f in _layout.Pages[p].Floats)
                if (ReferenceEquals(f.Object.Source, d) && f.Anchor is not null) return FromFloat(f, p);
            foreach (var l in Src.PageLines(p))
                foreach (var (k, obj) in l.Paragraph.Content.Drawings)
                    if (ReferenceEquals(obj.Source, d) && k >= l.Paragraph.Lines[l.Line].Start && k < l.Paragraph.Lines[l.Line].End && InlineRect(l, k) is { } r)
                        return new ObjSel(d, l.Paragraph.Paragraph, p, r, obj.Root.Rotation, null, obj);
        }
        return null;
    }

    /// <summary>Taille d'une poignée en points (8 pixels à l'écran, quel que soit le zoom).</summary>
    double HandleSize => 8 / Math.Max(0.1, (_pagesView?.Scale ?? 1.33) * _zoom);

    /// <summary>Poignées du cadre : 0 haut-gauche, 1 haut, 2 haut-droite, 3 droite, 4 bas-droite, 5 bas, 6 bas-gauche, 7 gauche.</summary>
    static Point[] Handles(Rect r) => new[]
    {
        r.TopLeft, new Point(r.X + r.Width / 2, r.Y), r.TopRight, new Point(r.Right, r.Y + r.Height / 2),
        r.BottomRight, new Point(r.X + r.Width / 2, r.Bottom), r.BottomLeft, new Point(r.X, r.Y + r.Height / 2),
    };

    Point RotationHandle(Rect r) => new(r.X + r.Width / 2, r.Y - 18 / Math.Max(0.1, (_pagesView?.Scale ?? 1.33) * _zoom));

    /// <summary>Poignée sous le point (-1 : aucune, 8 : rotation).</summary>
    int HandleAt(ObjSel s, Point at)
    {
        var pt = Unrotate(s, at);
        double h = HandleSize;
        var hs = Handles(s.Frame);
        for (int i = 0; i < hs.Length; i++)
            if (Math.Abs(pt.X - hs[i].X) <= h && Math.Abs(pt.Y - hs[i].Y) <= h) return i;
        var rh = RotationHandle(s.Frame);
        if ((pt - rh).Length <= h) return 8;
        return -1;
    }

    /// <summary>Point ramené dans le repère non tourné de l'objet.</summary>
    static Point Unrotate(ObjSel s, Point at)
    {
        if (s.Rotation == 0) return at;
        var c = new Point(s.Frame.X + s.Frame.Width / 2, s.Frame.Y + s.Frame.Height / 2);
        return new RotateTransform(-s.Rotation, c.X, c.Y).Transform(at);
    }

    // ================================================================ sélection et souris

    void SelectObject(ObjSel? s)
    {
        _objSel = s;
        _objDrag = null;
        _objGhost = null;
        if (s is not null)
        {
            _anchor = null;
            // le point d'insertion reste près de l'objet (hors de vue : la sélection d'objet le remplace)
            _caret = new TextPoint(s.Anchor, 0);
            StatusText.Text = s.Float is not null
                ? "Objet sélectionné : glisser pour le déplacer, poignées pour la taille ou la rotation ; Suppr pour le supprimer, Échap pour revenir au texte."
                : "Objet en ligne sélectionné : poignées pour la taille ; Suppr pour le supprimer.";
        }
        if (_pagesView is not null)
        {
            _pagesView.Caret = s is null && _caret is { } c ? CaretRect(c) : null;
            _pagesView.DrawSelection();
        }
        UpdateObjectTab();
    }

    /// <summary>Cadre de l'objet sélectionné pour le calque de sélection de la vue.</summary>
    (int Page, Rect Frame, double Rotation, bool Floating, Rect? Ghost, double GhostRotation, Point? Anchor)? ObjectOverlay()
    {
        if (_objSel is not { } s || _layout is null) return null;
        Point? anchor = null;
        if (s.Float is not null && Src.LinesOf(s.Anchor) is { Count: > 0 } refs && Src.Line(refs[0]) is var first && refs[0].Page == s.Page)
            anchor = new Point(first.X - 14, first.Y);
        return (s.Page, s.Frame, s.Rotation, s.Float is not null, _objGhost, _ghostRotation, anchor);
    }

    /// <summary>Bouton enfoncé : poignée de l'objet sélectionné, objet sous le pointeur. Vrai si pris en charge.</summary>
    bool ObjectMouseDown(int page, Point at, MouseButtonEventArgs e)
    {
        if (_layout is null || _pagesView is null || page < 0) return false;
        // zone de texte en cours d'édition : un clic dans son cadre reste dans la zone
        if (_box is not null)
        {
            if (BoxFloat() is { } bf && Contains(bf.Box, at) && bf.Page == page) return false;
            ExitTextBox(select: false);
        }
        if (_objSel is { } s && s.Page == page)
        {
            int h = HandleAt(s, at);
            if (h >= 0)
            {
                _objDrag = (h == 8 ? DragKind.Rotate : DragKind.Resize, h, at);
                _pagesView.CaptureMouse();
                return true;
            }
            if (s.Float is not null && new Rect(s.Frame.X, s.Frame.Y, s.Frame.Width, s.Frame.Height).Contains(Unrotate(s, at)))
            {
                // double-clic dans une zone de texte : écrire dedans
                if (e.ClickCount == 2 && TextBoxIndex(s.Object) >= 0) { EnterTextBox(s, at); return true; }
                _objDrag = (DragKind.Move, -1, at);
                _pagesView.CaptureMouse();
                return true;
            }
        }
        if (ObjectAt(page, at) is not { } hit) { if (_objSel is not null) SelectObject(null); return false; }
        // clic dans le texte d'une zone de texte : on écrit directement (comme Word)
        if (hit.Float is { } f && TextBoxIndex(hit.Object) >= 0 && InTextArea(f, at)) { EnterTextBox(hit, at); return true; }
        SelectObject(hit);
        if (hit.Float is not null)
        {
            _objDrag = (DragKind.Move, -1, at);
            _pagesView.CaptureMouse();
        }
        return true;
    }

    /// <summary>Le point tombe sur le texte d'une zone de texte (une de ses lignes), pas sur son bord.</summary>
    static bool InTextArea(FloatBox f, Point at)
    {
        var r = new Rect(f.X, f.Y, f.W, f.H);
        r.Inflate(-4, -4);
        return r.Contains(at) && f.TextLines.Count > 0;
    }

    static int TextBoxIndex(DrawingObject o) => o.Root switch
    {
        ShapeVisual s => s.TextBox,
        _ => -1,
    };

    /// <summary>Survol et glisser : curseurs des poignées, cadre fantôme pendant le glisser.</summary>
    bool ObjectMouseMove(int page, Point at, MouseEventArgs e)
    {
        if (_pagesView is null) return false;
        if (_objDrag is { } d && _objSel is { } s)
        {
            if (e.LeftButton != MouseButtonState.Pressed) { _objDrag = null; _objGhost = null; return true; }
            if (page != s.Page) return true;
            (_objGhost, _ghostRotation) = Ghost(s, d, at, Keyboard.Modifiers.HasFlag(ModifierKeys.Shift));
            _pagesView.DrawSelection();
            return true;
        }
        if (_objSel is { } sel && sel.Page == page)
        {
            int h = HandleAt(sel, at);
            if (h >= 0)
            {
                _pagesView.Cursor = h switch { 0 or 4 => Cursors.SizeNWSE, 2 or 6 => Cursors.SizeNESW, 1 or 5 => Cursors.SizeNS, 3 or 7 => Cursors.SizeWE, _ => Cursors.Hand };
                return true;
            }
            if (sel.Float is not null && sel.Frame.Contains(Unrotate(sel, at))) { _pagesView.Cursor = Cursors.SizeAll; return true; }
        }
        if (_box is null && ObjectAt(page, at) is { } hit)
        {
            _pagesView.Cursor = hit.Float is { } f && TextBoxIndex(hit.Object) >= 0 && InTextArea(f, at) ? Cursors.IBeam : hit.Float is not null ? Cursors.SizeAll : Cursors.Arrow;
            return true;
        }
        return false;
    }

    /// <summary>Cadre résultant d'un glisser (et rotation) ; Maj : proportions gardées (toujours aux coins d'une image).</summary>
    (Rect, double) Ghost(ObjSel s, (DragKind Kind, int Handle, Point Start) d, Point at, bool shift)
    {
        var r = s.Frame;
        switch (d.Kind)
        {
            case DragKind.Move:
                return (new Rect(r.X + at.X - d.Start.X, r.Y + at.Y - d.Start.Y, r.Width, r.Height), s.Rotation);
            case DragKind.Rotate:
            {
                var c = new Point(r.X + r.Width / 2, r.Y + r.Height / 2);
                double angle = Math.Atan2(at.X - c.X, -(at.Y - c.Y)) * 180 / Math.PI;
                if (shift) angle = Math.Round(angle / 15) * 15;
                return (r, (angle + 360) % 360);
            }
            default:
            {
                var p = Unrotate(s, at);
                double l = r.Left, t = r.Top, rr = r.Right, b = r.Bottom;
                int h = d.Handle;
                if (h is 0 or 6 or 7) l = Math.Min(p.X, rr - 1);
                if (h is 2 or 3 or 4) rr = Math.Max(p.X, l + 1);
                if (h is 0 or 1 or 2) t = Math.Min(p.Y, b - 1);
                if (h is 4 or 5 or 6) b = Math.Max(p.Y, t + 1);
                bool keep = shift ^ (s.Object.Root is PictureVisual && h is 0 or 2 or 4 or 6);
                if (keep && r.Width > 0 && r.Height > 0)
                {
                    double ratio = r.Width / r.Height;
                    double w = rr - l, hh = b - t;
                    if (w / hh > ratio) w = hh * ratio; else hh = w / ratio;
                    if (h is 0 or 6 or 7) l = rr - w; else rr = l + w;
                    if (h is 0 or 1 or 2) t = b - hh; else b = t + hh;
                }
                return (new Rect(l, t, rr - l, b - t), s.Rotation);
            }
        }
    }

    /// <summary>Bouton relâché : déplacement, taille ou rotation appliqués (une étape annulable).</summary>
    bool ObjectMouseUp(int page, Point at)
    {
        if (_objDrag is not { } d || _objSel is not { } s || _pagesView is null) return false;
        _pagesView.ReleaseMouseCapture();
        var (ghost, rotation) = Ghost(s, d, at, Keyboard.Modifiers.HasFlag(ModifierKeys.Shift));
        _objDrag = null;
        _objGhost = null;
        bool moved = d.Kind switch
        {
            DragKind.Move => (at - d.Start).Length > 0.75,
            DragKind.Rotate => Math.Abs(rotation - s.Rotation) > 0.5,
            _ => Math.Abs(ghost.Width - s.Frame.Width) > 0.5 || Math.Abs(ghost.Height - s.Frame.Height) > 0.5 || (ghost.TopLeft - s.Frame.TopLeft).Length > 0.5,
        };
        if (!moved || page != s.Page) { _pagesView.DrawSelection(); return true; }
        switch (d.Kind)
        {
            case DragKind.Move: MoveObject(s, ghost.X, ghost.Y); break;
            case DragKind.Rotate: EditObject("Rotation", s, dc => DrawingOps.SetRotation(dc, Math.Round(rotation, 1))); break;
            default:
                EditObject("Taille de l'objet", s, dc =>
                {
                    DrawingOps.SetSize(dc, ghost.Width, ghost.Height);
                    // poignée de gauche ou du haut d'un objet flottant : son coin bouge aussi
                    if (s.Float is { } f && (Math.Abs(ghost.X - s.Frame.X) > 0.01 || Math.Abs(ghost.Y - s.Frame.Y) > 0.01))
                        SetFloatPosition(dc, f, ghost.X, ghost.Y);
                });
                break;
        }
        return true;
    }

    // ================================================================ modifications

    /// <summary>
    /// Modification d'un dessin, annulable : transaction sur le bloc de premier niveau qui contient son paragraphe
    /// d'ancrage (corps, ou en-tête en cours d'édition), remise en page, objet resélectionné.
    /// </summary>
    void EditObject(string label, ObjSel s, Action<DrawingContent> op, bool reselect = true)
    {
        if (EditRoot is not { } root) return;
        DocPos pos;
        try { pos = DocPos.Of(root, s.Anchor, 0); }
        catch (ArgumentException) { NotEditable(); return; }
        int top = pos.Path[0];
        try
        {
            using var tx = _journal.Begin(label, root, Array.Empty<int>(), top, 1, pos);
            op(s.Drawing);
            tx.Commit(pos);
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or FormatException)
        {
            StatusText.Text = $"{label} impossible : {ex.Message}";
            return;
        }
        AfterObjectEdit(s.Drawing, reselect, label);
    }

    void AfterObjectEdit(DrawingContent d, bool reselect, string label)
    {
        Relayout();
        _drawings?.Clear();
        _pagesView?.InvalidateVisual();
        UpdateModifiedMarker();
        EnsureAutoSaveTimer();
        if (reselect && Reselect(d) is { } again) SelectObject(again);
        else SelectObject(null);
        StatusText.Text = $"{label}. Ctrl+Z pour annuler.";
    }

    /// <summary>
    /// Position d'un objet flottant (coin haut gauche, points de la page) écrite dans ses repères : horizontal page,
    /// marge ou colonne (sinon la colonne), vertical page ou marge (sinon le paragraphe d'ancrage).
    /// </summary>
    void SetFloatPosition(DrawingContent dc, FloatBox f, double x, double y, double? paraTop = null)
    {
        var o = f.Object;
        var page = _layout!.Pages[Math.Clamp(_objSel?.Page ?? 0, 0, _layout.Pages.Count - 1)];
        var g = page.Geometry;
        string hFrom = o.H.RelativeFrom is "page" or "margin" ? o.H.RelativeFrom : "column";
        double hBase = hFrom switch { "page" => 0, "margin" => g.Left, _ => ColumnLeft(page, f) };
        string vFrom = o.V.RelativeFrom is "page" or "margin" ? o.V.RelativeFrom : "paragraph";
        double vBase = vFrom switch { "page" => 0, "margin" => page.ContentTop, _ => paraTop ?? ParagraphTop(f.Anchor!, page) };
        DrawingOps.SetPosition(dc, hFrom, x - hBase, vFrom, y - vBase);
    }

    /// <summary>Bord gauche de la colonne de l'ancre (la première ligne du paragraphe d'ancrage).</summary>
    double ColumnLeft(Page page, FloatBox f) =>
        f.Anchor is { } p && Src.LinesOf(p) is { Count: > 0 } refs ? Src.Line(refs[0]).X : page.Geometry.Left;

    /// <summary>Haut du paragraphe sur la page (sa première ligne, saut éventuel compris).</summary>
    double ParagraphTop(Paragraph p, Page page) =>
        Src.LinesOf(p) is { Count: > 0 } refs ? Src.Line(refs[0]).Y - Src.Line(refs[0]).Paragraph.Lines[Src.Line(refs[0]).Line].Gap : page.ContentTop;

    /// <summary>
    /// Déplacement d'un objet flottant. Ancré au paragraphe (repère vertical « paragraphe ») et posé ailleurs sur la page :
    /// l'ancre passe au paragraphe le plus proche du nouvel emplacement (comme Word), en une seule étape annulable.
    /// </summary>
    void MoveObject(ObjSel s, double x, double y)
    {
        if (s.Float is not { } f || EditRoot is not { } root || _layout is null) return;
        var page = _layout.Pages[s.Page];
        // paragraphe de premier niveau de la page le plus proche du haut de l'objet
        Paragraph? target = null;
        double targetTop = 0;
        if (f.Object.V.RelativeFrom is "paragraph" or "line" && _hf is null)
            foreach (var l in page.Lines)
            {
                if (!ReferenceEquals(l.Paragraph.Paragraph, l.Paragraph.Paragraph) || l.Paragraph.Lines[l.Line].SameRow) continue;
                if (l.Y - l.Paragraph.Lines[l.Line].Gap > y + 0.01 && target is not null) break;
                if (root.Blocks.Contains(l.Paragraph.Paragraph) && l.Line == 0) { target = l.Paragraph.Paragraph; targetTop = l.Y - l.Paragraph.Lines[l.Line].Gap; }
            }
        if (target is null || ReferenceEquals(target, s.Anchor) || !root.Blocks.Contains(s.Anchor))
        {
            EditObject("Déplacer l'objet", s, dc => SetFloatPosition(dc, f, x, y));
            return;
        }
        int a = root.Blocks.IndexOf(s.Anchor), b = root.Blocks.IndexOf(target);
        int first = Math.Min(a, b), count = Math.Abs(a - b) + 1;
        try
        {
            using var tx = _journal.Begin("Déplacer l'objet", root, Array.Empty<int>(), first, count, null);
            // le segment qui porte le dessin passe au début du paragraphe d'arrivée
            var run = s.Anchor.Runs().First(r => r.Content.Contains(s.Drawing));
            run.Content.Remove(s.Drawing);
            var moved = new Run { Props = run.Props.Clone() };
            moved.Content.Add(s.Drawing);
            if (run.Content.Count == 0) RemoveInline(s.Anchor.Inlines, run);
            target.Inlines.Insert(0, moved);
            SetFloatPosition(s.Drawing, f, x, y, targetTop);
            tx.Commit();
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException)
        {
            StatusText.Text = $"Déplacement impossible : {ex.Message}";
            return;
        }
        AfterObjectEdit(s.Drawing, true, "Objet déplacé (ancré au paragraphe le plus proche)");
    }

    static bool RemoveInline(List<Inline> inlines, Inline target)
    {
        if (inlines.Remove(target)) return true;
        foreach (var i in inlines)
            if (i is InlineContainer c && RemoveInline(c.Children, target) || i is InlineContentControl sdt && RemoveInline(sdt.Content, target)) return true;
        return false;
    }

    /// <summary>Supprime l'objet sélectionné (son segment s'il ne porte rien d'autre).</summary>
    internal void DeleteSelectedObject()
    {
        if (_objSel is not { } s) return;
        var anchor = s.Anchor;
        EditObject("Supprimer l'objet", s, d =>
        {
            var run = anchor.Runs().First(r => r.Content.Contains(d));
            run.Content.Remove(d);
            if (run.Content.Count == 0) RemoveInline(anchor.Inlines, run);
        }, reselect: false);
        _caret = LayoutOf(anchor) is not null ? new TextPoint(anchor, 0) : DocumentStart();
        UpdateCaret();
    }

    /// <summary>Touches quand un objet est sélectionné : Suppr / Retour arrière, flèches (déplacer), Échap. Vrai si traitée.</summary>
    bool ObjectKey(Key key, bool ctrl)
    {
        if (_objSel is not { } s) return false;
        switch (key)
        {
            case Key.Delete or Key.Back:
                DeleteSelectedObject();
                return true;
            case Key.Escape:
                SelectObject(null);
                UpdateCaret();
                return true;
            case Key.Left or Key.Right or Key.Up or Key.Down when s.Float is { } f:
                double step = ctrl ? 1 : 7.2;
                double dx = key == Key.Left ? -step : key == Key.Right ? step : 0, dy = key == Key.Up ? -step : key == Key.Down ? step : 0;
                EditObject("Déplacer l'objet", s, dc => SetFloatPosition(dc, f, s.Frame.X + dx, s.Frame.Y + dy));
                return true;
        }
        return false;
    }

    // ================================================================ habillage, ordre, rotation

    /// <summary>Choix d'habillage de Word : en ligne, carré, rapproché, au travers, haut et bas, derrière, devant.</summary>
    internal void SetObjectWrap(string mode)
    {
        if (_objSel is not { } s) return;
        if (mode == "inline")
        {
            if (s.Float is null) return;
            EditObject("Aligné sur le texte", s, DrawingOps.ToInline);
            return;
        }
        var (wrap, behind) = mode switch
        {
            "square" => (WrapKind.Square, false), "tight" => (WrapKind.Tight, false), "through" => (WrapKind.Through, false),
            "topBottom" => (WrapKind.TopAndBottom, false), "behind" => (WrapKind.None, true), _ => (WrapKind.None, false),
        };
        if (s.Float is null)
        {
            // en ligne → flottant à la même place (depuis la colonne et le haut du paragraphe)
            var page = _layout!.Pages[s.Page];
            double colX = Src.LinesOf(s.Anchor) is { Count: > 0 } refs ? Src.Line(refs[0]).X : page.Geometry.Left;
            double top = ParagraphTop(s.Anchor, page);
            long z = DrawingOps.MaxZ(AllStories()) + 1024;
            EditObject("Habillage", s, d => DrawingOps.ToAnchor(d, s.Frame.X - colX, s.Frame.Y - top, wrap, behind, z));
            return;
        }
        EditObject("Habillage", s, d => DrawingOps.SetWrap(d, wrap, behind));
    }

    /// <summary>Histoires du document (corps, en-têtes et pieds) pour les identifiants et l'ordre d'empilement.</summary>
    IEnumerable<Story> AllStories() => _pkg is null ? Array.Empty<Story>() : _pkg.Stories.Values.SelectMany(p => p.Stories());

    internal void ArrangeObject(string how)
    {
        if (_objSel is not { Float: { } f } s || _layout is null) return;
        var zs = _layout.Pages[s.Page].Floats.Where(x => x.PartId == f.PartId).Select(x => x.Object.Z).OrderBy(z => z).ToList();
        long z = how switch
        {
            "front" => (zs.LastOrDefault() + 1024),
            "back" => Math.Max(0, zs.FirstOrDefault() - 1024),
            "forward" => zs.FirstOrDefault(v => v > f.Object.Z, f.Object.Z) + 1,
            _ => Math.Max(0, zs.LastOrDefault(v => v < f.Object.Z, f.Object.Z) - 1),
        };
        EditObject(how is "front" or "forward" ? "Avancer" : "Reculer", s, d => DrawingOps.SetZ(d, z));
    }

    internal void RotateObject(double delta)
    {
        if (_objSel is not { } s) return;
        EditObject("Rotation", s, d => DrawingOps.SetRotation(d, s.Rotation + delta));
    }

    internal void FlipObject(bool horizontal)
    {
        if (_objSel is not { } s) return;
        EditObject(horizontal ? "Retourner horizontalement" : "Retourner verticalement", s, d => DrawingOps.Flip(d, horizontal));
    }

    /// <summary>Position rapide sur la page (comme le menu Position de Word) : neuf emplacements, habillage carré.</summary>
    internal void PlaceObject(string h, string v)
    {
        if (_objSel is not { } s) return;
        EditObject("Position", s, d =>
        {
            if (s.Float is null) DrawingOps.ToAnchor(d, 0, 0, WrapKind.Square, false, DrawingOps.MaxZ(AllStories()) + 1024);
            DrawingOps.SetAlignment(d, "margin", h, "margin", v);
        });
    }

    // ================================================================ insertion

    /// <summary>Insertion ▸ Images ▸ Cet appareil… : image en ligne au point d'insertion, réduite à la largeur du texte.</summary>
    internal void InsertPictureFromFile()
    {
        var dlg = new OpenFileDialog
        {
            Title = "Insérer une image",
            Filter = "Images|*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.tif;*.tiff;*.ico|Tous les fichiers|*.*",
        };
        if (dlg.ShowDialog(this) != true) return;
        byte[] data;
        try { data = File.ReadAllBytes(dlg.FileName); }
        catch (IOException ex) { StatusText.Text = $"Image illisible : {ex.Message}"; return; }
        InsertPicture(data, Path.GetExtension(dlg.FileName), Path.GetFileNameWithoutExtension(dlg.FileName));
    }

    /// <summary>Image du presse-papiers (copie d'écran, image copiée d'un autre programme) : insérée en PNG.</summary>
    internal bool PasteImage()
    {
        BitmapSource? img = null;
        try { if (Clipboard.ContainsImage()) img = Clipboard.GetImage(); }
        catch (System.Runtime.InteropServices.ExternalException) { }
        if (img is null) return false;
        var enc = new PngBitmapEncoder();
        enc.Frames.Add(BitmapFrame.Create(img));
        using var ms = new MemoryStream();
        enc.Save(ms);
        InsertPicture(ms.ToArray(), ".png", "Image");
        return true;
    }

    /// <summary>Image insérée en ligne au point d'insertion (taille d'origine d'après sa résolution, réduite à la largeur du texte).</summary>
    internal void InsertPicture(byte[] data, string extension, string name)
    {
        if (_pkg?.Main is not { } main || EditRoot is not { } root || _caret is not { } c) return;
        double w, h;
        try
        {
            var frame = BitmapDecoder.Create(new MemoryStream(data), BitmapCreateOptions.None, BitmapCacheOption.OnLoad).Frames[0];
            double dpiX = frame.DpiX > 1 ? frame.DpiX : 96, dpiY = frame.DpiY > 1 ? frame.DpiY : 96;
            w = frame.PixelWidth * 72.0 / dpiX;
            h = frame.PixelHeight * 72.0 / dpiY;
        }
        catch (Exception ex) when (ex is NotSupportedException or FileFormatException or IOException or ArgumentException)
        {
            StatusText.Text = "Format d'image non pris en charge.";
            return;
        }
        // réduite à la largeur du texte (et à la hauteur de la page), comme Word
        double maxW = LayoutOf(c.Paragraph)?.Width ?? 450, maxH = _layout is { Pages.Count: > 0 } lay ? lay.Pages[0].ContentBottom - lay.Pages[0].ContentTop : 650;
        double k = Math.Min(1, Math.Min(maxW / Math.Max(1, w), maxH / Math.Max(1, h)));
        w *= k;
        h *= k;
        var partName = _hf is not null && _pkg.Relationships(main.PartName).TryGetValue(_hf.RelId, out var hr) ? hr.Target : main.PartName;
        var rid = _pkg.AddImage(partName, data, extension);
        InsertDrawing(DrawingOps.CreatePicture(main.W, DrawingOps.MaxDocPrId(AllStories()) + 1, name, rid, w, h), "Insérer une image");
    }

    /// <summary>Forme flottante devant le texte, près du point d'insertion (taille par défaut de Word : 1 pouce).</summary>
    internal void InsertShape(string preset)
    {
        if (_pkg?.Main is not { } main || _caret is not { } c) return;
        bool line = preset is "line" or "straightConnector1";
        double x = Locate(c) is { } loc ? CaretGeometry.X(loc.Line.Paragraph, loc.Line.Line, c.Index) : 0;
        var xml = DrawingOps.CreateShape(main.W, DrawingOps.MaxDocPrId(AllStories()) + 1, preset, Math.Max(0, x), 4, 72, line ? 0 : 72, DrawingOps.MaxZ(AllStories()) + 1024);
        InsertDrawing(xml, "Insérer une forme");
    }

    /// <summary>Zone de texte flottante (habillage carré) près du point d'insertion ; on écrit aussitôt dedans.</summary>
    internal void InsertTextBox()
    {
        if (_pkg?.Main is not { } main || _caret is not { } c) return;
        double x = Locate(c) is { } loc ? CaretGeometry.X(loc.Line.Paragraph, loc.Line.Line, c.Index) : 0;
        var width = Math.Min(180, (LayoutOf(c.Paragraph)?.Width ?? 450) - x);
        if (width < 72) { x = 0; width = 180; }
        var xml = DrawingOps.CreateShape(main.W, DrawingOps.MaxDocPrId(AllStories()) + 1, "rect", x, 4, width, 72, DrawingOps.MaxZ(AllStories()) + 1024, textBox: true);
        var d = InsertDrawing(xml, "Insérer une zone de texte");
        if (d is not null && Reselect(d) is { } s) EnterTextBox(s, null);
    }

    /// <summary>Dessin inséré au point d'insertion dans son propre segment ; sélectionné ensuite.</summary>
    DrawingContent? InsertDrawing(System.Xml.Linq.XElement xml, string label)
    {
        if (_pkg?.Main is not { } main || EditRoot is not { } root || Ordered() is not { } sel) return null;
        if (ToDocPos(sel.From) is not { } from || ToDocPos(sel.To) is not { } to) { NotEditable(); return null; }
        var part = _hf is not null ? _pkg.Stories.Values.FirstOrDefault(p => p.Stories().Contains(_hf.Story)) ?? main : main;
        Docx.DocxPackage.EnsureDrawingNamespaces(part);
        part.Dirty = true;
        var w = main.W;
        // la zone de texte reçoit son histoire (le contenu w:txbxContent est dans le modèle, pas dans le XML)
        var d = new DrawingContent(xml);
        foreach (var box in DrawingReader.TextBoxContents(xml))
        {
            var story = new Story();
            story.Blocks.Add(new Paragraph { Props = new PropertyBag("pPr", w) });
            d.TextBoxes.Add(story);
            box.RemoveNodes();
        }
        var run = new Run { Props = new PropertyBag("rPr", w, new System.Xml.Linq.XElement(w + "rPr", new System.Xml.Linq.XElement(w + "noProof"))) };
        run.Content.Add(d);
        var fragment = new Fragment(new List<Block> { new Paragraph { Props = new PropertyBag("pPr", w), Inlines = { run } } }, false);
        DocPos pos;
        try { pos = Fragments.Paste(_journal, root, label, from, to, fragment, NewParaId()); }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or InvalidOperationException)
        {
            StatusText.Text = $"{label} impossible ici : {ex.Message}";
            return null;
        }
        // Paste clone le morceau : le dessin inséré est celui du document, juste avant la nouvelle position
        AfterTyping(pos);
        var inserted = DrawingNear(root, pos) ?? d;
        if (Reselect(inserted) is { } s) SelectObject(s);
        StatusText.Text = $"{label}. Ctrl+Z pour annuler.";
        return inserted;
    }

    /// <summary>Dessin du paragraphe d'une position (le dernier avant le décalage).</summary>
    static DrawingContent? DrawingNear(Story root, DocPos pos)
    {
        try
        {
            var p = pos.Paragraph(root);
            return p.Runs().SelectMany(r => r.Content).OfType<DrawingContent>().LastOrDefault();
        }
        catch (Exception ex) when (ex is ArgumentException or ArgumentOutOfRangeException) { return null; }
    }

    // ================================================================ zones de texte

    /// <summary>Objet flottant de la zone en cours d'édition et sa page.</summary>
    (FloatBox Box, int Page)? BoxFloat()
    {
        if (_box is null || _layout is null) return null;
        for (int p = 0; p < _layout.Pages.Count; p++)
            foreach (var f in _layout.Pages[p].Floats)
                if (ReferenceEquals(f.Object.Source, _box.Drawing)) return (f, p);
        return null;
    }

    /// <summary>Entre dans le texte d'une zone de texte (clic dans la zone, double-clic sur la forme, nouvelle zone).</summary>
    void EnterTextBox(ObjSel s, Point? at)
    {
        int index = TextBoxIndex(s.Object);
        if (s.Float is not { } f || index < 0 || index >= s.Drawing.TextBoxes.Count) return;
        _objSel = null;
        _objDrag = null;
        _box = new BoxEdit(s.Drawing, index, s.Drawing.TextBoxes[index], s.Page, new BoxLines(s.Page, f, index));
        _anchor = null;
        _caret = at is { } p && HitText(s.Page, p) is { } t ? t
            : _box.Lines.PageLines(s.Page).Count > 0 ? new TextPoint(_box.Lines.PageLines(s.Page)[^1].Paragraph.Paragraph, _box.Lines.PageLines(s.Page)[^1].Paragraph.Content.Text.Length)
            : null;
        UpdateCaret(scroll: false);
        UpdateObjectTab();
        StatusText.Text = "Zone de texte : écrivez ; Échap pour sélectionner la zone, clic hors de la zone pour revenir au document.";
    }

    /// <summary>Quitte le texte de la zone (Échap : la zone reste sélectionnée).</summary>
    void ExitTextBox(bool select)
    {
        if (_box is null) return;
        var d = _box.Drawing;
        SyncFallback(d);
        _box = null;
        _anchor = null;
        if (select && Reselect(d) is { } s) SelectObject(s);
        else { _caret = DocumentStart(); SelectObject(null); }
        UpdateCaret(scroll: false);
    }

    /// <summary>
    /// Zone de texte d'un mc:AlternateContent : la variante VML de repli reçoit une copie du texte modifié (les deux
    /// variantes restent identiques, comme quand Word enregistre).
    /// </summary>
    static void SyncFallback(DrawingContent d)
    {
        var boxes = DrawingReader.TextBoxContents(d.Xml);
        int choice = boxes.Count(b => !b.Ancestors().Any(a => a.Name.LocalName == "Fallback"));
        for (int k = choice; k < boxes.Count && k - choice < choice && k < d.TextBoxes.Count; k++)
        {
            var src = d.TextBoxes[k - choice];
            var dst = d.TextBoxes[k];
            dst.Blocks.Clear();
            dst.Blocks.AddRange(src.Blocks.Select(b => b.Clone()));
        }
    }

    /// <summary>
    /// Après une modification du texte d'une zone : le paragraphe d'ancrage est remis en page (la zone et son habillage
    /// suivent), ses lignes reprises pour le point d'insertion. Vrai si traité.
    /// </summary>
    bool RelayoutTextBox()
    {
        if (_box is null || _layout is null || Body is not { } body || _journal.LastChange is not { } ch || !ReferenceEquals(ch.Root, _box.Story)) return false;
        SyncFallback(_box.Drawing);
        _layout.Invalidate(_box.Story.Blocks);
        // bloc de premier niveau qui porte le dessin
        var top = body.Blocks.FirstOrDefault(b => new Story { Blocks = { b } }.Paragraphs().Any(p => p.Runs().Any(r => r.Content.Contains(_box.Drawing))));
        if (top is null) { _layout.LayoutAll(); }
        else _layout.Update(new[] { top });
        _drawings?.Clear();
        if (BoxFloat() is { } bf) _box = _box with { Page = bf.Page, Lines = new BoxLines(bf.Page, bf.Box, _box.Box) };
        _pagesView?.InvalidateMeasure();
        _pagesView?.InvalidateVisual();
        _thumbsView?.InvalidateVisual();
        UpdatePageStatus();
        return true;
    }

    // ================================================================ onglet contextuel et menus

    RadioButton? _objectTab;
    StackPanel? _objectPanel;

    /// <summary>Onglet « Format de l'image » / « Format de la forme » : visible quand un objet est sélectionné.</summary>
    void UpdateObjectTab()
    {
        bool show = _objSel is not null || _box is not null;
        if (_objectTab is null)
        {
            if (!show) return;
            _objectTab = new RadioButton { Style = (Style)FindResource("RibbonTab"), Content = "Format de l'image", Name = "TabObjectFormat", Foreground = R("SelectedText") };
            _objectTab.Checked += Tab_Checked;
            TabStrip.Children.Add(_objectTab);
            _objectPanel = BuildObjectPanel();
            _panels["TabObjectFormat"] = _objectPanel;
            RibbonPanels.Children.Add(_objectPanel);
        }
        var vis = show ? Visibility.Visible : Visibility.Collapsed;
        bool picture = (_objSel?.Object.Root ?? (_box is not null ? BoxFloat()?.Box.Object.Root : null)) is PictureVisual;
        _objectTab.Content = picture ? "Format de l'image" : "Format de la forme";
        if (_objectTab.Visibility != vis)
        {
            _objectTab.Visibility = vis;
            if (!show && _objectTab.IsChecked == true) TabHome.IsChecked = true;
        }
    }

    StackPanel BuildObjectPanel()
    {
        var panel = new StackPanel { Orientation = Orientation.Horizontal, Visibility = Visibility.Collapsed };
        TGroup(panel, "Ajuster",
            TBig("", "Corrections", "Luminosité et contraste de l'image", b => ShowMenu(b, new (string, UIElement?, Action)[]
            {
                ("Normal", null, () => Corrections(0, 0, null)), ("Plus lumineux (+20 %)", null, () => Corrections(0.2, 0, null)),
                ("Plus sombre (-20 %)", null, () => Corrections(-0.2, 0, null)), ("Plus contrasté (+20 %)", null, () => Corrections(0, 0.2, null)),
                ("Moins contrasté (-20 %)", null, () => Corrections(0, -0.2, null)),
            })),
            TBig("", "Couleur", "Couleurs de l'image : d'origine ou niveaux de gris", b => ShowMenu(b, new (string, UIElement?, Action)[]
            {
                ("Couleurs d'origine", null, () => Corrections(null, null, false)), ("Niveaux de gris", null, () => Corrections(null, null, true)),
            })));
        TGroup(panel, "Styles",
            TBig("", "Remplissage", "Couleur de remplissage de la forme", b => ShowColorMenu(b, "Remplissage", c => EditSelected("Remplissage", d => DrawingOps.SetFill(d, c)))),
            TBig("", "Contour", "Couleur et épaisseur du contour (de la bordure d'une image)", ShowOutlineMenu),
            TBig("", "Forme", "Changer la forme (ou découper l'image selon une forme)", ShowGeometryMenu));
        TGroup(panel, "Accessibilité",
            TBig("", "Texte de remplacement", "Description de l'objet pour les lecteurs d'écran", _ => ShowAltText()));
        TGroup(panel, "Organiser",
            TBig("", "Position", "Placer l'objet sur la page (habillage carré)", ShowObjectPositionMenu),
            TBig("", "Habillage", "Habillage du texte autour de l'objet", ShowWrapMenu),
            TBig("", "Avancer", "Avancer, mettre au premier plan", b => ShowMenu(b, new (string, UIElement?, Action)[]
            {
                ("Avancer", null, () => ArrangeObject("forward")), ("Mettre au premier plan", null, () => ArrangeObject("front")),
            })),
            TBig("", "Reculer", "Reculer, mettre à l'arrière-plan", b => ShowMenu(b, new (string, UIElement?, Action)[]
            {
                ("Reculer", null, () => ArrangeObject("backward")), ("Mettre à l'arrière-plan", null, () => ArrangeObject("back")),
            })),
            TBig("", "Rotation", "Faire pivoter ou retourner l'objet", b => ShowMenu(b, new (string, UIElement?, Action)[]
            {
                ("Faire pivoter à droite de 90°", null, () => RotateObject(90)), ("Faire pivoter à gauche de 90°", null, () => RotateObject(-90)),
                ("Retourner verticalement", null, () => FlipObject(false)), ("Retourner horizontalement", null, () => FlipObject(true)),
            })));
        TGroup(panel, "Taille",
            TBig("", "Rogner", "Rogner l'image (pourcentage de chaque bord)", _ => ShowCropDialog()),
            TBig("", "Taille", "Hauteur et largeur exactes", _ => ShowSizeDialog()));
        return panel;
    }

    void EditSelected(string label, Action<DrawingContent> op)
    {
        if (_objSel is { } s) EditObject(label, s, op);
        else if (_box is not null && BoxFloat() is { } bf)
        {
            var d = _box.Drawing;
            ExitTextBox(select: true);
            if (_objSel is { } s2 && ReferenceEquals(s2.Drawing, d)) EditObject(label, s2, op);
        }
        else StatusText.Text = "Sélectionnez d'abord une image ou une forme (clic dessus).";
    }

    void Corrections(double? brightness, double? contrast, bool? gray)
    {
        if (_objSel?.Object.Root is not PictureVisual p) { StatusText.Text = "Corrections : sélectionnez une image."; return; }
        EditSelected("Corrections de l'image", d => DrawingOps.SetCorrections(d, brightness ?? (gray is null ? 0 : p.Brightness), contrast ?? (gray is null ? 0 : p.Contrast), gray ?? p.Grayscale));
    }

    void ShowObjectPositionMenu(FrameworkElement b) => ShowMenu(b,
        from v in new[] { ("top", "En haut"), ("center", "Au milieu"), ("bottom", "En bas") }
        from h in new[] { ("left", "à gauche"), ("center", "au centre"), ("right", "à droite") }
        select ($"{v.Item2} {h.Item2}", (UIElement?)null, (Action)(() => PlaceObject(h.Item1, v.Item1))));

    void ShowWrapMenu(FrameworkElement b) => ShowMenu(b, new (string, UIElement?, Action)[]
    {
        ("Aligné sur le texte", null, () => SetObjectWrap("inline")), ("Carré", null, () => SetObjectWrap("square")),
        ("Rapproché", null, () => SetObjectWrap("tight")), ("Au travers", null, () => SetObjectWrap("through")),
        ("Haut et bas", null, () => SetObjectWrap("topBottom")), ("Derrière le texte", null, () => SetObjectWrap("behind")),
        ("Devant le texte", null, () => SetObjectWrap("front")),
    });

    static readonly (string Name, string Rgb)[] ShapeColors =
    {
        ("Blanc", "FFFFFF"), ("Noir", "000000"), ("Gris", "A6A6A6"), ("Bleu", "4472C4"), ("Orange", "ED7D31"), ("Vert", "70AD47"),
        ("Jaune", "FFC000"), ("Rouge", "C00000"), ("Bleu clair", "9DC3E6"), ("Bleu OpenWrite", "087CFF"),
    };

    void ShowColorMenu(FrameworkElement b, string what, Action<string?> apply) => ShowMenu(b,
        new[] { ($"Sans {what.ToLowerInvariant()}", (UIElement?)null, (Action)(() => apply(null))) }
        .Concat(ShapeColors.Select(c => (c.Name, (UIElement?)new Border { Width = 14, Height = 14, Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#" + c.Rgb)), BorderBrush = Brushes.Gray, BorderThickness = new Thickness(0.5) },
            (Action)(() => apply(c.Rgb))))));

    void ShowOutlineMenu(FrameworkElement b)
    {
        var current = _objSel?.Object.Root switch { ShapeVisual s => s.Line, PictureVisual p => p.Line, _ => null };
        double width = current?.Width ?? 0.75;
        ShowMenu(b, new[] { ("Sans contour", (UIElement?)null, (Action)(() => EditSelected("Contour", d => DrawingOps.SetLine(d, null)))) }
            .Concat(ShapeColors.Select(c => (c.Name, (UIElement?)new Border { Width = 14, Height = 14, BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#" + c.Rgb)), BorderThickness = new Thickness(3) },
                (Action)(() => EditSelected("Contour", d => DrawingOps.SetLine(d, c.Rgb, width))))))
            .Concat(new[] { 0.5, 1, 1.5, 2.25, 3, 4.5, 6 }.Select(wd => ($"Épaisseur {wd:0.##} pt", (UIElement?)null,
                (Action)(() => EditSelected("Contour", d => DrawingOps.SetLine(d, current?.Color ?? "000000", wd)))))));
    }

    static readonly (string Preset, string Name)[] GalleryShapes =
    {
        ("rect", "Rectangle"), ("roundRect", "Rectangle à coins arrondis"), ("ellipse", "Ellipse"), ("triangle", "Triangle isocèle"),
        ("rtTriangle", "Triangle rectangle"), ("diamond", "Losange"), ("pentagon", "Pentagone"), ("hexagon", "Hexagone"),
        ("octagon", "Octogone"), ("parallelogram", "Parallélogramme"), ("trapezoid", "Trapèze"), ("plus", "Croix"),
        ("star5", "Étoile à 5 branches"), ("star8", "Étoile à 8 branches"), ("heart", "Cœur"), ("smileyFace", "Émoticône"),
        ("cloud", "Nuage"), ("sun", "Soleil"), ("moon", "Lune"), ("lightningBolt", "Éclair"),
        ("rightArrow", "Flèche : droite"), ("leftArrow", "Flèche : gauche"), ("upArrow", "Flèche : haut"), ("downArrow", "Flèche : bas"),
        ("leftRightArrow", "Flèche : gauche-droite"), ("chevron", "Chevron"), ("homePlate", "Pentagone (flèche)"),
        ("wedgeRectCallout", "Bulle : rectangle"), ("wedgeRoundRectCallout", "Bulle : rectangle arrondi"), ("wedgeEllipseCallout", "Bulle : ovale"),
        ("cloudCallout", "Bulle : nuage"), ("flowChartProcess", "Organigramme : processus"), ("flowChartDecision", "Organigramme : décision"),
        ("flowChartTerminator", "Organigramme : terminaison"), ("line", "Ligne"),
    };

    void ShowGeometryMenu(FrameworkElement b) => ShowMenu(b, GalleryShapes.Where(g => g.Preset != "line")
        .Select(g => (g.Name, (UIElement?)ShapeIcon(g.Preset), (Action)(() => EditSelected("Modifier la forme", d => DrawingOps.SetGeometry(d, g.Preset))))));

    /// <summary>Petit aperçu d'une forme prédéfinie (galerie des formes, menus).</summary>
    static UIElement ShapeIcon(string preset)
    {
        var group = new GeometryGroup();
        foreach (var p in ShapeGeometry.PresetPaths(preset, 16, 14))
        {
            var fig = new PathFigure { IsClosed = false };
            var geo = new PathGeometry();
            foreach (var s in p.Segments)
            {
                var a = s.Points;
                switch (s.Op)
                {
                    case 'M': fig = new PathFigure { StartPoint = new Point(a[0], a[1]) }; geo.Figures.Add(fig); break;
                    case 'L': fig.Segments.Add(new LineSegment(new Point(a[0], a[1]), true)); break;
                    case 'C': fig.Segments.Add(new BezierSegment(new Point(a[0], a[1]), new Point(a[2], a[3]), new Point(a[4], a[5]), true)); break;
                    case 'Q': fig.Segments.Add(new QuadraticBezierSegment(new Point(a[0], a[1]), new Point(a[2], a[3]), true)); break;
                    case 'Z': fig.IsClosed = true; break;
                }
            }
            group.Children.Add(geo);
        }
        return new System.Windows.Shapes.Path
        {
            Data = group, Width = 18, Height = 16, Stretch = Stretch.None, Margin = new Thickness(1),
            Fill = new SolidColorBrush(Color.FromRgb(0xC5, 0xD9, 0xF1)), Stroke = new SolidColorBrush(Color.FromRgb(0x2F, 0x52, 0x8F)), StrokeThickness = 1,
        };
    }

    /// <summary>Insertion ▸ Formes : galerie des formes prédéfinies (insérée près du point d'insertion, devant le texte).</summary>
    internal void ShowShapesGallery(FrameworkElement anchor)
    {
        var wrap = new WrapPanel { Width = 300, Margin = new Thickness(6) };
        var popup = new System.Windows.Controls.Primitives.Popup { PlacementTarget = anchor, Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom, StaysOpen = false, AllowsTransparency = true };
        foreach (var (preset, name) in GalleryShapes)
        {
            var b = new Button { Content = ShapeIcon(preset), ToolTip = name, Width = 28, Height = 26, Margin = new Thickness(1), Padding = new Thickness(0), Focusable = false };
            b.Click += (_, _) => { popup.IsOpen = false; InsertShape(preset); _pagesView?.Focus(); };
            wrap.Children.Add(b);
        }
        popup.Child = new Border
        {
            Background = R("Surface"), BorderBrush = R("Line"), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(6),
            Child = new StackPanel { Children = { new TextBlock { Text = "Formes", FontWeight = FontWeights.SemiBold, Margin = new Thickness(8, 6, 8, 0) }, wrap } },
        };
        popup.IsOpen = true;
    }

    /// <summary>Insertion ▸ Images : depuis un fichier, ou l'image du presse-papiers.</summary>
    internal void ShowPicturesMenu(FrameworkElement anchor) => ShowMenu(anchor, new (string, UIElement?, Action)[]
    {
        ("Cet appareil…", null, InsertPictureFromFile),
        ("Coller l'image du presse-papiers", null, () => { if (!PasteImage()) StatusText.Text = "Le presse-papiers ne contient pas d'image."; }),
    });

    void ShowAltText()
    {
        if (_objSel is not { } s) { StatusText.Text = "Sélectionnez d'abord une image ou une forme."; return; }
        var w = new AltTextWindow(s.Object.Description ?? "") { Owner = this };
        if (w.ShowDialog() == true) EditObject("Texte de remplacement", s, d => DrawingOps.SetAltText(d, w.Description));
    }

    void ShowSizeDialog()
    {
        if (_objSel is not { } s) { StatusText.Text = "Sélectionnez d'abord une image ou une forme."; return; }
        var w = new ObjectSizeWindow(s.Frame.Width, s.Frame.Height, s.Rotation, s.Object.Root is PictureVisual) { Owner = this };
        if (w.ShowDialog() != true) return;
        EditObject("Taille de l'objet", s, d =>
        {
            DrawingOps.SetSize(d, w.ResultWidth, w.ResultHeight);
            if (Math.Abs(w.ResultRotation - s.Rotation) > 0.01) DrawingOps.SetRotation(d, w.ResultRotation);
        });
    }

    void ShowCropDialog()
    {
        if (_objSel is not { Object.Root: PictureVisual p } s) { StatusText.Text = "Rogner : sélectionnez une image."; return; }
        var w = new CropWindow(p.Crop) { Owner = this };
        if (w.ShowDialog() != true) return;
        // la taille de l'image affichée suit la partie gardée (comme Word)
        var (l, t, r, b) = p.Crop;
        double fullW = s.Frame.Width / Math.Max(0.01, 1 - l - r), fullH = s.Frame.Height / Math.Max(0.01, 1 - t - b);
        var (nl, nt, nr, nb) = w.Result;
        EditObject("Rogner", s, d =>
        {
            DrawingOps.SetCrop(d, nl, nt, nr, nb);
            DrawingOps.SetSize(d, fullW * (1 - nl - nr), fullH * (1 - nt - nb));
        });
    }

    /// <summary>Entrées du menu contextuel sur un objet sélectionné.</summary>
    bool AddObjectContextItems(ContextMenu menu)
    {
        if (_objSel is null) return false;
        MenuItem Item(ItemsControl parent, string header, Action act)
        {
            var mi = new MenuItem { Header = header };
            mi.Click += (_, _) => { act(); _pagesView?.Focus(); };
            parent.Items.Add(mi);
            return mi;
        }
        Item(menu, "Couper", Cut);
        Item(menu, "Copier", Copy);
        Item(menu, "Supprimer", DeleteSelectedObject);
        menu.Items.Add(new Separator());
        var wrap = new MenuItem { Header = "Habillage" };
        foreach (var (label, mode) in new[] { ("Aligné sur le texte", "inline"), ("Carré", "square"), ("Rapproché", "tight"), ("Au travers", "through"),
                     ("Haut et bas", "topBottom"), ("Derrière le texte", "behind"), ("Devant le texte", "front") })
            Item(wrap, label, () => SetObjectWrap(mode));
        menu.Items.Add(wrap);
        var order = new MenuItem { Header = "Ordre" };
        Item(order, "Mettre au premier plan", () => ArrangeObject("front"));
        Item(order, "Avancer", () => ArrangeObject("forward"));
        Item(order, "Reculer", () => ArrangeObject("backward"));
        Item(order, "Mettre à l'arrière-plan", () => ArrangeObject("back"));
        menu.Items.Add(order);
        var rot = new MenuItem { Header = "Rotation" };
        Item(rot, "Faire pivoter à droite de 90°", () => RotateObject(90));
        Item(rot, "Faire pivoter à gauche de 90°", () => RotateObject(-90));
        Item(rot, "Retourner verticalement", () => FlipObject(false));
        Item(rot, "Retourner horizontalement", () => FlipObject(true));
        menu.Items.Add(rot);
        menu.Items.Add(new Separator());
        Item(menu, "Taille et rotation…", ShowSizeDialog);
        if (_objSel.Object.Root is PictureVisual) Item(menu, "Rogner…", ShowCropDialog);
        Item(menu, "Texte de remplacement…", ShowAltText);
        return true;
    }

    /// <summary>Copie de l'objet sélectionné (presse-papiers interne : un segment avec le dessin) ; vrai si traité.</summary>
    bool CopySelectedObject()
    {
        if (_objSel is not { } s) return false;
        var run = s.Anchor.Runs().FirstOrDefault(r => r.Content.Contains(s.Drawing));
        if (run is null) return false;
        var copy = new Run { Props = run.Props.Clone() };
        copy.Content.Add(s.Drawing.Clone());
        var w = s.Anchor.Props.Ns;
        var fragment = new Fragment(new List<Block> { new Paragraph { Props = new PropertyBag("pPr", w), Inlines = { copy } } }, false);
        SetInternalClipboard(fragment, s.Object.Description ?? s.Object.Name ?? "Objet");
        StatusText.Text = "Objet copié.";
        return true;
    }
}
