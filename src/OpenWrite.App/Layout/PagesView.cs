using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using OpenWrite.Layout;
using Page = OpenWrite.Layout.Page;

namespace OpenWrite.App.Layout;

/// <summary>Dessins des pages d'un document, calculés à la demande et gardés (partagés par la vue et les miniatures).</summary>
public sealed class PageDrawings
{
    readonly DocumentLayout _layout;
    readonly PageRenderer _renderer;
    readonly Dictionary<Page, Drawing> _cache = new(ReferenceEqualityComparer.Instance);

    public PageDrawings(DocumentLayout layout, PageRenderer renderer)
    {
        _layout = layout;
        _renderer = renderer;
    }

    public DocumentLayout Layout => _layout;
    public IReadOnlyList<Page> Pages => _layout.Pages;

    /// <summary>Dessin de la page <paramref name="i"/>, en points (sans le fond).</summary>
    public Drawing Get(int i)
    {
        var page = _layout.Pages[i];
        if (_cache.TryGetValue(page, out var d)) return d;
        var visual = new DrawingVisual();
        using (var dc = visual.RenderOpen()) _renderer.Draw(dc, page, _layout, i + 1, _layout.Pages.Count);
        d = visual.Drawing ?? new DrawingGroup();
        d.Freeze();
        _cache[page] = d;
        return d;
    }

    /// <summary>Oublie les dessins (après une nouvelle mise en page : numéros et contenus changent).</summary>
    public void Clear() => _cache.Clear();
}

/// <summary>
/// Vue des pages, les unes sous les autres (PLAN.md §3.3, « DocumentView ») : seules les pages visibles dans le
/// ScrollViewer qui la contient sont dessinées, à chaque défilement. Sert aussi aux miniatures (petite échelle,
/// numéro sous chaque page). Coordonnées des pages en points ; <see cref="Scale"/> convertit en unités WPF.
/// </summary>
public sealed class PagesView : FrameworkElement
{
    readonly PageDrawings _drawings;
    readonly List<Rect> _rects = new();
    ScrollViewer? _scroller;

    public PagesView(PageDrawings drawings, double scale, double gap, bool thumbnails = false)
    {
        _drawings = drawings;
        Scale = scale;
        Gap = gap;
        Thumbnails = thumbnails;
        Loaded += (_, _) =>
        {
            _scroller = FindScroller(this);
            if (_scroller is not null) _scroller.ScrollChanged += OnScroll;
        };
        Unloaded += (_, _) =>
        {
            if (_scroller is not null) _scroller.ScrollChanged -= OnScroll;
            _scroller = null;
        };
    }

    void OnScroll(object? sender, ScrollChangedEventArgs e)
    {
        if (e.VerticalChange != 0 || e.HorizontalChange != 0 || e.ViewportHeightChange != 0 || e.ExtentHeightChange != 0) InvalidateVisual();
    }

    static ScrollViewer? FindScroller(DependencyObject d)
    {
        for (var p = VisualTreeHelper.GetParent(d); p is not null; p = VisualTreeHelper.GetParent(p))
            if (p is ScrollViewer sv) return sv;
        return null;
    }

    /// <summary>Unités WPF par point (96/72 à 100 %).</summary>
    public double Scale { get; }
    public double Gap { get; }
    public bool Thumbnails { get; }
    /// <summary>Page mise en évidence (miniatures : page affichée dans la vue principale).</summary>
    public int Current { get; set { if (field != value) { field = value; InvalidateVisual(); } } } = -1;

    public PageDrawings Drawings => _drawings;
    public int PageCount => _drawings.Pages.Count;

    /// <summary>Rectangle de la page <paramref name="i"/> dans la vue.</summary>
    public Rect PageRect(int i) => i >= 0 && i < _rects.Count ? _rects[i] : Rect.Empty;

    const double LabelHeight = 16;

    protected override Size MeasureOverride(Size availableSize)
    {
        _rects.Clear();
        double maxW = _drawings.Pages.Count == 0 ? 0 : _drawings.Pages.Max(p => p.Geometry.PageWidth) * Scale;
        double y = 0;
        foreach (var p in _drawings.Pages)
        {
            double w = p.Geometry.PageWidth * Scale, h = p.Geometry.PageHeight * Scale;
            _rects.Add(new Rect((maxW - w) / 2, y, w, h));
            y += h + Gap + (Thumbnails ? LabelHeight : 0);
        }
        return new Size(maxW, Math.Max(0, y - Gap));
    }

    /// <summary>Partie visible de la vue (coordonnées de la vue) ; toute la vue si elle n'est pas dans un ScrollViewer.</summary>
    Rect Visible()
    {
        if (_scroller is null || !IsVisible) return new Rect(RenderSize);
        try
        {
            var t = _scroller.TransformToDescendant(this);
            return t.TransformBounds(new Rect(0, 0, _scroller.ViewportWidth, _scroller.ViewportHeight));
        }
        catch (InvalidOperationException) { return new Rect(RenderSize); }
    }

    static readonly Brush Shadow = Freeze(new SolidColorBrush(Color.FromArgb(0x30, 0x27, 0x4D, 0x9D)));
    static readonly Pen PageBorder = Freeze(new Pen(Freeze(new SolidColorBrush(Color.FromRgb(0xD5, 0xDC, 0xEA))), 1));
    static readonly Pen CurrentBorder = Freeze(new Pen(Freeze(new SolidColorBrush(Color.FromRgb(0x08, 0x7C, 0xFF))), 2.5));

    protected override void OnRender(DrawingContext dc)
    {
        var visible = Visible();
        visible.Inflate(0, 200);
        for (int i = 0; i < _rects.Count; i++)
        {
            var r = _rects[i];
            if (r.Bottom < visible.Top || r.Top > visible.Bottom) continue;
            dc.DrawRectangle(Shadow, null, new Rect(r.X + 2, r.Y + 3, r.Width, r.Height));
            dc.DrawRectangle(Brushes.White, i == Current ? CurrentBorder : PageBorder, r);
            dc.PushClip(new RectangleGeometry(r));
            dc.PushTransform(new MatrixTransform(Scale, 0, 0, Scale, r.X, r.Y));
            dc.DrawDrawing(_drawings.Get(i));
            dc.Pop();
            dc.Pop();
            if (Thumbnails)
            {
                var label = new FormattedText((i + 1).ToString(CultureInfo.CurrentCulture), CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
                    new Typeface("Segoe UI"), 11, i == Current ? CurrentBorder.Brush : Brushes.DimGray, 1.0);
                dc.DrawText(label, new Point(r.X + (r.Width - label.Width) / 2, r.Bottom + 2));
            }
        }
    }

    /// <summary>Page sous un point de la vue et position correspondante sur la page (points) ; -1 hors des pages.</summary>
    public (int Page, Point At) HitPage(Point p)
    {
        for (int i = 0; i < _rects.Count; i++)
            if (_rects[i].Contains(p)) return (i, new Point((p.X - _rects[i].X) / Scale, (p.Y - _rects[i].Y) / Scale));
        return (-1, default);
    }

    /// <summary>Page dont le haut est le plus proche au-dessus d'une ordonnée de la vue (page affichée).</summary>
    public int PageAt(double y)
    {
        int best = 0;
        for (int i = 0; i < _rects.Count; i++)
            if (_rects[i].Top - Gap / 2 <= y) best = i; else break;
        return best;
    }

    /// <summary>Ligne de texte sous un point de la page (points), avec la position du caractère.</summary>
    public static (PlacedLine Line, int Char)? HitLine(Page page, Point at)
    {
        foreach (var l in page.Lines)
        {
            var box = l.Paragraph.Lines[l.Line];
            if (at.Y < l.Y || at.Y >= l.Y + box.Height) continue;
            double shift = 0;
            var c = l.Paragraph.Content;
            foreach (var piece in box.Pieces)
            {
                double x = l.X + piece.X + shift;
                int spaces = 0;
                double cx = x;
                for (int k = piece.Start; k < piece.Start + piece.Length; k++)
                {
                    var st = c.Styles[c.StyleOf[k]];
                    double adv = c.Text[k] == Atom.Object && c.Objects.TryGetValue(k, out var o) ? o.W
                        : st.Face.Advance(c.Text[k]) * st.Size + st.Tracking + (c.Text[k] == ' ' ? box.ExtraPerSpace : 0);
                    if (c.Text[k] == ' ') spaces++;
                    if (at.X >= cx && at.X < cx + adv) return (l, k);
                    cx += adv;
                }
                shift += spaces * box.ExtraPerSpace;
            }
            // dans la hauteur de la ligne mais hors du texte : la ligne, sans caractère
            if (at.X >= l.X - 4 && at.X <= l.X + l.Paragraph.Width + 4) return (l, -1);
        }
        return null;
    }

    static T Freeze<T>(T f) where T : Freezable { f.Freeze(); return f; }
}
