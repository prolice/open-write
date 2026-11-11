using System.Windows;
using System.Windows.Media;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Layout;
using Page = OpenWrite.Layout.Page;

namespace OpenWrite.App.Output;

/// <summary>
/// Pages d'un document mis en page, telles qu'elles sortent sur papier ou en PDF : même dessin que l'écran
/// (<see cref="OpenWrite.App.Layout.PageRenderer"/>), sans les repères d'édition (marques de mise en forme, quadrillage des
/// tableaux, limites de l'en-tête). Comme Word avec ses réglages par défaut, la couleur de page n'est pas imprimée.
/// </summary>
public sealed class OutputPages
{
    readonly DocumentLayout _layout;
    readonly OpenWrite.App.Layout.PageRenderer _renderer;
    readonly Func<InlineContainer, string?>? _linkTarget;

    /// <param name="image">Images des relations (comme pour la vue des pages).</param>
    /// <param name="linkTarget">Adresse d'un lien hypertexte (null : lien sans adresse externe, non reporté dans le PDF).</param>
    public OutputPages(DocumentLayout layout, Func<string?, string, ImageSource?> image, Func<InlineContainer, string?>? linkTarget = null, Brush? revisionBar = null)
    {
        _layout = layout;
        // marques de révision comme à l'écran (Word imprime le document avec ses marques)
        _renderer = new OpenWrite.App.Layout.PageRenderer(image) { ShowMarks = false, HeaderMode = false, ForOutput = true, RevisionBar = revisionBar };
        _linkTarget = linkTarget;
    }

    public int Count => _layout.Pages.Count;

    /// <summary>Format de la page <paramref name="number"/> (à partir de 1), en points : chaque section a le sien.</summary>
    public Size SizeOf(int number)
    {
        var g = _layout.Pages[number - 1].Geometry;
        return new Size(g.PageWidth, g.PageHeight);
    }

    /// <summary>Dessin de la page <paramref name="number"/> (à partir de 1), en points, découpé au bord de la page.</summary>
    public DrawingGroup Render(int number)
    {
        var page = _layout.Pages[number - 1];
        var size = SizeOf(number);
        var group = new DrawingGroup();
        // le quadrillage des tableaux sans bordure est un repère d'écran (Word ne l'imprime pas)
        bool grid = OpenWrite.App.Layout.PageRenderer.ShowGridlines;
        OpenWrite.App.Layout.PageRenderer.ShowGridlines = false;
        try
        {
            using var dc = group.Open();
            dc.PushClip(new RectangleGeometry(new Rect(size)));
            _renderer.Draw(dc, page, _layout, number, _layout.Pages.Count);
            dc.Pop();
        }
        finally { OpenWrite.App.Layout.PageRenderer.ShowGridlines = grid; }
        group.Freeze();
        return group;
    }

    /// <summary>
    /// Zones des liens hypertextes du corps de la page (points, depuis le coin haut gauche) et leur adresse :
    /// une zone par morceau de ligne couvert par un même lien.
    /// </summary>
    public List<(Rect Area, string Url)> Links(int number)
    {
        var res = new List<(Rect, string)>();
        if (_linkTarget is null) return res;
        var page = _layout.Pages[number - 1];
        foreach (var l in page.Lines)
        {
            var pl = l.Paragraph;
            var box = pl.Lines[l.Line];
            var c = pl.Content;
            InlineContainer? cur = null;
            int start = 0;
            void Close(int end)
            {
                if (cur is null || end <= start) return;
                if (_linkTarget(cur) is { Length: > 0 } url)
                {
                    double x1 = CaretGeometry.X(pl, l.Line, start), x2 = CaretGeometry.X(pl, l.Line, end);
                    if (x2 > x1) res.Add((new Rect(l.X + x1, l.Y, x2 - x1, box.Height), url));
                }
                cur = null;
            }
            for (int i = box.Start; i < box.End && i < c.Text.Length; i++)
            {
                var link = c.ModelStart.Length > i ? LinkOps.LinkAt(pl.Paragraph, c.ModelStart[i])?.Link : null;
                if (ReferenceEquals(link, cur)) continue;
                Close(i);
                if (link is not null) { cur = link; start = i; }
            }
            Close(Math.Min(box.End, c.Text.Length));
        }
        return res;
    }
}
