using System.Globalization;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace OpenWrite.App.Layout;

/// <summary>Repère de la règle que l'on fait glisser.</summary>
public enum RulerHandle { None, FirstLine, Hanging, Left, Right, Tab }

/// <summary>
/// Règle horizontale de Word (lot 3) : marges de la colonne en gris, graduations en centimètres depuis la marge de gauche,
/// retraits du paragraphe du point d'insertion (première ligne, suspendu, gauche, droit) et taquets de tabulation.
/// Les repères se déplacent à la souris (pas de 0,125 cm ; Alt : sans pas) ; un clic sur la règle pose un taquet
/// gauche, un taquet tiré hors de la règle disparaît, un double-clic ouvre la boîte Paragraphe.
/// Toutes les valeurs sont en points de la page ; <see cref="Update"/> donne l'origine et l'échelle d'affichage.
/// </summary>
public sealed class RulerBar : FrameworkElement
{
    const double CmPerPt = 2.54 / 72;
    double _origin, _scale = 1, _pageWidth;
    double _textLeft, _textRight;         // bords de la colonne (cellule) du paragraphe, points de la page
    double _left, _first, _right;         // retraits : gauche, première ligne (relative), droit (points)
    List<double> _tabs = new();
    bool _active;

    RulerHandle _drag;
    int _dragTab = -1;
    double _dragValue;
    bool _dragOut;

    /// <summary>Fin d'un déplacement : repère, nouvelle valeur (points depuis le bord gauche de la colonne ; retrait droit : depuis le bord droit).</summary>
    public event Action<RulerHandle, int, double?>? HandleMoved;
    /// <summary>Clic sur une zone libre : position d'un nouveau taquet (points depuis le bord gauche de la colonne).</summary>
    public event Action<double>? TabAdded;
    public event Action? DoubleClicked;

    public RulerBar()
    {
        Height = 24;
        SnapsToDevicePixels = true;
        ToolTip = "Règle : faites glisser les retraits et les taquets ; un clic pose un taquet ; double-clic : Paragraphe…";
    }

    /// <summary>Recale la règle : origine de la page (unités de la règle), échelle (unités par point), colonne et retraits.</summary>
    public void Update(bool active, double origin, double scale, double pageWidth, double textLeft, double textRight,
        double left, double first, double right, IEnumerable<double> tabs)
    {
        var tabList = tabs.ToList();
        if (active == _active && origin == _origin && scale == _scale && pageWidth == _pageWidth && textLeft == _textLeft && textRight == _textRight
            && left == _left && first == _first && right == _right && tabList.SequenceEqual(_tabs)) return;
        _active = active;
        _origin = origin;
        _scale = scale;
        _pageWidth = pageWidth;
        _textLeft = textLeft;
        _textRight = textRight;
        _left = left;
        _first = first;
        _right = right;
        _tabs = tabList;
        InvalidateVisual();
    }

    double X(double pagePt) => _origin + pagePt * _scale;

    static readonly Brush Back = Freeze(new SolidColorBrush(Color.FromRgb(0xF3, 0xF6, 0xFB)));
    static readonly Brush MarginBrush = Freeze(new SolidColorBrush(Color.FromRgb(0xD9, 0xE0, 0xEC)));
    static readonly Brush Ink = Freeze(new SolidColorBrush(Color.FromRgb(0x44, 0x4F, 0x6A)));
    static readonly Brush Handle = Freeze(new SolidColorBrush(Color.FromRgb(0xFF, 0xFF, 0xFF)));
    static readonly Pen HandlePen = Freeze(new Pen(Freeze(new SolidColorBrush(Color.FromRgb(0x5B, 0x74, 0xA8))), 1));
    static readonly Pen TickPen = Freeze(new Pen(Freeze(new SolidColorBrush(Color.FromRgb(0x7A, 0x86, 0xA0))), 1));
    static readonly Typeface Face = new("Segoe UI");

    protected override void OnRender(DrawingContext dc)
    {
        double h = ActualHeight, w = ActualWidth;
        dc.DrawRectangle(Back, null, new Rect(0, 0, w, h));
        if (!_active || _scale <= 0) return;
        double top = 3, bottom = h - 3;
        // page : marges en gris, colonne en blanc
        dc.DrawRectangle(MarginBrush, null, new Rect(X(0), top, _pageWidth * _scale, bottom - top));
        dc.DrawRectangle(Brushes.White, null, new Rect(X(_textLeft), top, Math.Max(0, (_textRight - _textLeft) * _scale), bottom - top));
        // graduations : 0 à la marge de gauche, un nombre par centimètre, traits par quart de centimètre
        double ptPerQuarter = 0.25 / CmPerPt;
        int qMin = (int)Math.Floor(-_textLeft / ptPerQuarter), qMax = (int)Math.Ceiling((_pageWidth - _textLeft) / ptPerQuarter);
        for (int q = qMin; q <= qMax; q++)
        {
            double x = X(_textLeft + q * ptPerQuarter);
            if (q % 4 == 0)
            {
                if (q == 0) continue;
                var label = new FormattedText(Math.Abs(q / 4).ToString(CultureInfo.CurrentCulture), CultureInfo.CurrentCulture, FlowDirection.LeftToRight, Face, 9, Ink, 1.0);
                dc.DrawText(label, new Point(x - label.Width / 2, (h - label.Height) / 2));
            }
            else
            {
                double len = q % 2 == 0 ? 5 : 2.5;
                dc.DrawLine(TickPen, new Point(x + 0.5, h / 2 - len / 2), new Point(x + 0.5, h / 2 + len / 2));
            }
        }
        // taquets (L)
        for (int i = 0; i < _tabs.Count; i++)
        {
            double tx = X(_textLeft + (i == _dragTab && _drag == RulerHandle.Tab ? _dragValue : _tabs[i]));
            if (i == _dragTab && _dragOut) continue;
            var pen = new Pen(Ink, 1.6);
            dc.DrawLine(pen, new Point(tx, bottom - 8), new Point(tx, bottom - 1));
            dc.DrawLine(pen, new Point(tx, bottom - 1), new Point(tx + 6, bottom - 1));
        }
        // retraits
        double left = _drag is RulerHandle.Left or RulerHandle.Hanging ? _dragValue : _left;
        double first = _drag == RulerHandle.FirstLine ? _dragValue : _drag == RulerHandle.Left ? _dragValue + (_first) : _left + _first;
        if (_drag == RulerHandle.Hanging) first = _left + _first;           // la première ligne reste en place
        double right = _drag == RulerHandle.Right ? _dragValue : _right;
        DrawTriangle(dc, X(_textLeft + first), top, down: true);
        DrawTriangle(dc, X(_textLeft + left), bottom - 7, down: false);
        dc.DrawRectangle(Handle, HandlePen, new Rect(X(_textLeft + left) - 4.5, bottom - 4, 9, 4));
        DrawTriangle(dc, X(_textRight - right), bottom - 7, down: false);
    }

    static void DrawTriangle(DrawingContext dc, double x, double y, bool down)
    {
        var g = new StreamGeometry();
        using (var c = g.Open())
        {
            if (down) { c.BeginFigure(new Point(x - 4.5, y), true, true); c.LineTo(new Point(x + 4.5, y), true, false); c.LineTo(new Point(x, y + 6), true, false); }
            else { c.BeginFigure(new Point(x - 4.5, y + 6), true, true); c.LineTo(new Point(x + 4.5, y + 6), true, false); c.LineTo(new Point(x, y), true, false); }
        }
        g.Freeze();
        dc.DrawGeometry(Handle, HandlePen, g);
    }

    // ------------------------------------------------------------------ souris

    (RulerHandle Handle, int Tab) HitHandle(Point p)
    {
        double h = ActualHeight;
        bool near(double pagePt) => Math.Abs(p.X - X(pagePt)) <= 5;
        if (p.Y < h / 2 && near(_textLeft + _left + _first)) return (RulerHandle.FirstLine, -1);
        if (p.Y >= h - 7 && near(_textLeft + _left)) return (RulerHandle.Left, -1);
        if (p.Y >= h / 2 && near(_textLeft + _left)) return (RulerHandle.Hanging, -1);
        if (p.Y >= h / 2 && near(_textRight - _right)) return (RulerHandle.Right, -1);
        for (int i = 0; i < _tabs.Count; i++)
            if (p.Y >= h / 2 && p.X >= X(_textLeft + _tabs[i]) - 2 && p.X <= X(_textLeft + _tabs[i]) + 7) return (RulerHandle.Tab, i);
        return (RulerHandle.None, -1);
    }

    /// <summary>Valeur sous la souris : points depuis le bord gauche de la colonne (retrait droit : depuis le bord droit), au pas de 0,125 cm.</summary>
    double ValueAt(Point p, RulerHandle handle)
    {
        double pagePt = (p.X - _origin) / _scale;
        double v = handle == RulerHandle.Right ? _textRight - pagePt : pagePt - _textLeft;
        if (!Keyboard.Modifiers.HasFlag(ModifierKeys.Alt))
        {
            double step = 0.125 / CmPerPt;
            v = Math.Round(v / step) * step;
        }
        return v;
    }

    protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e)
    {
        if (!_active) return;
        if (e.ClickCount == 2) { DoubleClicked?.Invoke(); e.Handled = true; return; }
        var p = e.GetPosition(this);
        var (handle, tab) = HitHandle(p);
        if (handle == RulerHandle.None)
        {
            double pagePt = (p.X - _origin) / _scale;
            if (pagePt > _textLeft && pagePt < _textRight && p.Y >= ActualHeight / 3) TabAdded?.Invoke(ValueAt(p, RulerHandle.Tab));
            e.Handled = true;
            return;
        }
        _drag = handle;
        _dragTab = tab;
        _dragValue = handle switch
        {
            RulerHandle.FirstLine => _left + _first,
            RulerHandle.Left or RulerHandle.Hanging => _left,
            RulerHandle.Right => _right,
            _ => _tabs[tab],
        };
        CaptureMouse();
        e.Handled = true;
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        var p = e.GetPosition(this);
        if (_drag == RulerHandle.None)
        {
            Cursor = _active && HitHandle(p).Handle != RulerHandle.None ? Cursors.SizeWE : Cursors.Arrow;
            return;
        }
        _dragValue = ValueAt(p, _drag);
        if (_drag == RulerHandle.Left) _dragValue = Math.Max(_dragValue, Math.Max(-_textLeft, -_textLeft - _first));
        _dragOut = _drag == RulerHandle.Tab && (p.Y < -12 || p.Y > ActualHeight + 12);
        InvalidateVisual();
    }

    protected override void OnMouseLeftButtonUp(MouseButtonEventArgs e)
    {
        if (_drag == RulerHandle.None) return;
        ReleaseMouseCapture();
        var (handle, tab, value, removed) = (_drag, _dragTab, _dragValue, _dragOut);
        _drag = RulerHandle.None;
        _dragTab = -1;
        _dragOut = false;
        InvalidateVisual();
        HandleMoved?.Invoke(handle, tab, removed ? null : value);
        e.Handled = true;
    }

    static T Freeze<T>(T f) where T : Freezable { f.Freeze(); return f; }
}
