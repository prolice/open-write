using System.Windows;
using System.Windows.Controls.Primitives;
using System.Xml.Linq;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Layout;
using OpenWrite.App.Layout;

namespace OpenWrite.App;

// Lot 3 (version 0.4) : règle horizontale. Elle suit la colonne (ou la cellule) du paragraphe du point d'insertion ;
// retraits et taquets s'y déplacent à la souris et s'appliquent aux paragraphes de la sélection (annulable).
public partial class MainWindow
{
    bool _showRuler = true;
    ToggleButton? _rulerToggle;

    void InitRuler()
    {
        Ruler.HandleMoved += RulerHandleMoved;
        Ruler.TabAdded += pos => SetTabs(tabs => tabs.Add((pos, "left", null)), "Taquet ajouté");
        Ruler.DoubleClicked += ShowParagraphDialog;
        Scroller.SizeChanged += (_, _) => UpdateRuler();
    }

    internal void ToggleRuler()
    {
        _showRuler = !_showRuler;
        if (_rulerToggle is not null) _rulerToggle.IsChecked = _showRuler;
        UpdateRuler();
    }

    /// <summary>Recale la règle sur la page et le paragraphe du point d'insertion.</summary>
    void UpdateRuler()
    {
        RulerBorder.Visibility = _showRuler && _layout is not null ? Visibility.Visible : Visibility.Collapsed;
        if (RulerBorder.Visibility != Visibility.Visible || _pagesView is null || _caret is not { } c || Locate(c) is not { } loc)
        {
            Ruler.Update(false, 0, 1, 0, 0, 0, 0, 0, 0, Array.Empty<double>());
            return;
        }
        var page = _layout!.Pages[loc.Ref.Page];
        var rect = _pagesView.PageRect(loc.Ref.Page);
        double origin;
        try { origin = _pagesView.TranslatePoint(new Point(rect.X, 0), Ruler).X; }
        catch (InvalidOperationException) { return; }
        var pl = loc.Line.Paragraph;
        var f = pl.Format;
        Ruler.Update(true, origin, _pagesView.Scale * _zoom, page.Geometry.PageWidth, loc.Line.X, loc.Line.X + pl.Width,
            f.IndentLeft / 20.0, f.IndentFirstLine / 20.0, f.IndentRight / 20.0, Tabs(f.Effective).Select(t => t.Pos));
    }

    /// <summary>Taquets effectifs d'un w:pPr fusionné : position (points), alignement, points de suite.</summary>
    static List<(double Pos, string Align, string? Leader)> Tabs(XElement? ppr)
    {
        var list = new List<(double, string, string?)>();
        if (ppr is null) return list;
        var w = ppr.Name.Namespace;
        foreach (var t in ppr.Element(w + "tabs")?.Elements(w + "tab") ?? Enumerable.Empty<XElement>())
        {
            var val = t.Attribute(w + "val")?.Value ?? "left";
            if (val is "clear" or "bar" || !int.TryParse(t.Attribute(w + "pos")?.Value, out var pos)) continue;
            list.Add((pos / 20.0, val, t.Attribute(w + "leader")?.Value));
        }
        return list.OrderBy(x => x.Item1).ToList();
    }

    /// <summary>Auto-test : comme un déplacement de repère, ou un clic qui pose un taquet (<paramref name="handle"/> = None).</summary>
    internal void RulerAction(RulerHandle handle, int tab, double? value)
    {
        if (handle == RulerHandle.None && value is { } v) SetTabs(tabs => tabs.Add((v, "left", null)), "Taquet ajouté");
        else RulerHandleMoved(handle, tab, value);
    }

    void RulerHandleMoved(RulerHandle handle, int tab, double? value)
    {
        if (_caret is not { } c || LayoutOf(c.Paragraph) is not { } pl) return;
        var f = pl.Format;
        int Tw(double pt) => (int)Math.Round(pt * 20);
        switch (handle)
        {
            case RulerHandle.FirstLine when value is { } v:
                ApplyParagraphs("Retrait de première ligne", p => FormatOps.SetIndent(p, firstLine: Tw(v) - f.IndentLeft));
                break;
            case RulerHandle.Left when value is { } v:
                ApplyParagraphs("Retrait gauche", p => FormatOps.SetIndent(p, left: Tw(v)));
                break;
            case RulerHandle.Hanging when value is { } v:
                // la première ligne reste où elle est : son retrait relatif compense
                int firstAbs = f.IndentLeft + f.IndentFirstLine;
                ApplyParagraphs("Retrait suspendu", p => FormatOps.SetIndent(p, left: Tw(v), firstLine: firstAbs - Tw(v)));
                break;
            case RulerHandle.Right when value is { } v:
                ApplyParagraphs("Retrait droit", p => FormatOps.SetIndent(p, right: Tw(v)));
                break;
            case RulerHandle.Tab:
                SetTabs(tabs =>
                {
                    if (tab < 0 || tab >= tabs.Count) return;
                    var old = tabs[tab];
                    tabs.RemoveAt(tab);
                    if (value is { } v && v > 0) tabs.Add((v, old.Align, old.Leader));
                }, value is null ? "Taquet supprimé" : "Taquet déplacé");
                break;
        }
    }

    /// <summary>
    /// Modifie les taquets des paragraphes de la sélection : liste effective changée, puis écrite en mise en forme directe
    /// (taquets ajoutés, et « clear » pour ceux du style qui disparaissent).
    /// </summary>
    void SetTabs(Action<List<(double Pos, string Align, string? Leader)>> change, string label)
    {
        if (_styles is not { } styles) return;
        ApplyParagraphs(label, p =>
        {
            var tabs = Tabs(styles.Paragraph(p).Effective);
            change(tabs);
            var w = p.Props.Ns;
            // taquets du style seul (sans ceux du paragraphe)
            var bare = new Paragraph { Props = p.Props.Clone() };
            bare.Props.Remove("tabs");
            var inherited = Tabs(styles.Paragraph(bare).Effective);
            var el = new XElement(w + "tabs");
            foreach (var (pos, align, leader) in tabs.OrderBy(t => t.Pos))
            {
                if (inherited.Any(i => Math.Abs(i.Pos - pos) < 0.05 && i.Align == align && i.Leader == leader)) continue;
                var t = new XElement(w + "tab", new XAttribute(w + "val", align));
                if (leader is not null) t.Add(new XAttribute(w + "leader", leader));
                t.Add(new XAttribute(w + "pos", (int)Math.Round(pos * 20)));
                el.Add(t);
            }
            foreach (var i in inherited.Where(i => !tabs.Any(t => Math.Abs(t.Pos - i.Pos) < 0.05)))
                el.Add(new XElement(w + "tab", new XAttribute(w + "val", "clear"), new XAttribute(w + "pos", (int)Math.Round(i.Pos * 20))));
            if (el.HasElements) p.Props.Set(new XElement(w + "tabs", el.Elements().OrderBy(e => int.Parse(e.Attribute(w + "pos")!.Value))));
            else p.Props.Remove("tabs");
        });
    }
}
