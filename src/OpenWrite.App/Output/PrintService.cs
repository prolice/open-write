// Repris d'OpenPub 2.1.4 (OpenPub.App/Output/PrintService.cs) : pages d'OpenWrite (OutputPages), format de chaque
// page respecté (sections en paysage), fond blanc (la couleur de page n'est pas imprimée, comme Word par défaut).
using System.Printing;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using OpenWrite.Core.Output;

namespace OpenWrite.App.Output;

/// <summary>Disposition de l'impression.</summary>
public enum PrintMode
{
    /// <summary>Une page par feuille.</summary>
    Normal,
    /// <summary>Recto verso manuel : rectos, puis la pile retournée, versos.</summary>
    ManualDuplex,
    /// <summary>Livret : 2 pages par face d'une feuille paysage, dans l'ordre de pliage.</summary>
    Booklet,
    /// <summary>Livret sur une imprimante sans recto verso automatique : faces avant, puis faces arrière.</summary>
    BookletManualDuplex,
}

public sealed record PrintOptions
{
    /// <summary>Pages choisies (numéros à partir de 1).</summary>
    public required IReadOnlyList<int> Pages { get; init; }
    public int Copies { get; init; } = 1;
    public PrintMode Mode { get; init; } = PrintMode.Normal;
    /// <summary>Recto verso manuel : imprimer les versos dans l'ordre inverse (imprimantes qui empilent face dessus).</summary>
    public bool ReverseBacks { get; init; }
}

/// <summary>Une face imprimée : une page (Right = -1) ou deux pages côte à côte (0 = page blanche).</summary>
public readonly record struct PrintFace(int Left, int Right)
{
    public static readonly PrintFace Blank = new(0, -1);
    public bool TwoUp => Right >= 0;
    public override string ToString() => TwoUp ? $"{(Left == 0 ? "blanc" : Left)}-{(Right == 0 ? "blanc" : Right)}" : Left == 0 ? "blanc" : Left.ToString();
}

/// <summary>Une page dessinée pour l'impression : le dessin de la page (points) à l'échelle voulue, sur fond blanc.</summary>
sealed class PrintedPage : FrameworkElement
{
    readonly Drawing _drawing;
    readonly Size _pt;
    readonly double _scale;

    public PrintedPage(Drawing drawing, Size pointSize, double scale)
    {
        _drawing = drawing;
        _pt = pointSize;
        _scale = scale;
        Width = pointSize.Width * PrintService.DipPerPoint * scale;
        Height = pointSize.Height * PrintService.DipPerPoint * scale;
    }

    protected override void OnRender(DrawingContext dc)
    {
        dc.DrawRectangle(Brushes.White, null, new Rect(0, 0, Width, Height));
        dc.PushTransform(new ScaleTransform(PrintService.DipPerPoint * _scale, PrintService.DipPerPoint * _scale));
        dc.PushClip(new RectangleGeometry(new Rect(_pt)));
        dc.DrawDrawing(_drawing);
        dc.Pop();
        dc.Pop();
    }
}

/// <summary>
/// Impression : choix des pages, exemplaires (assemblés), livret (imposition 2 pages par face), recto verso manuel guidé.
/// Les faces sont des FixedPage WPF (même dessin que l'écran, sans les repères), envoyées à l'imprimante choisie.
/// </summary>
public static class PrintService
{
    public const double DipPerPoint = 96.0 / 72;

    /// <summary>
    /// Passages d'impression : un seul (normal, livret sur imprimante recto verso) ou deux (rectos puis versos).
    /// Les exemplaires sont assemblés : le document entier est répété.
    /// </summary>
    public static List<List<PrintFace>> Passes(PrintOptions o)
    {
        var pages = o.Pages;
        int copies = Math.Max(1, o.Copies);
        List<PrintFace> OneCopy()
        {
            if (o.Mode is PrintMode.Booklet or PrintMode.BookletManualDuplex)
            {
                // livret des pages choisies, numérotées dans leur ordre : la page k du livret est pages[k-1]
                return Booklet.Sides(pages.Count).Select(s => new PrintFace(s.Left == 0 ? 0 : pages[s.Left - 1], s.Right == 0 ? 0 : pages[s.Right - 1])).ToList();
            }
            return pages.Select(p => new PrintFace(p, -1)).ToList();
        }
        var faces = new List<PrintFace>();
        for (int c = 0; c < copies; c++)
        {
            var one = OneCopy();
            // recto verso : chaque exemplaire commence sur une nouvelle feuille
            if (o.Mode == PrintMode.ManualDuplex && one.Count % 2 == 1 && c < copies - 1) one.Add(PrintFace.Blank);
            faces.AddRange(one);
        }
        if (o.Mode is PrintMode.ManualDuplex or PrintMode.BookletManualDuplex)
        {
            var blank = o.Mode == PrintMode.BookletManualDuplex ? new PrintFace(0, 0) : PrintFace.Blank;
            var (fronts, backs) = ManualDuplex.Split(faces, blank, o.ReverseBacks);
            return new List<List<PrintFace>> { fronts, backs };
        }
        return new List<List<PrintFace>> { faces };
    }

    /// <summary>Feuille d'un livret (1/96 de pouce) : papier de l'imprimante (A4 à défaut), tourné en paysage.</summary>
    static Size BookletSheet(Size? paper)
    {
        var p = paper ?? new Size(595.276 * DipPerPoint, 841.89 * DipPerPoint);
        return new Size(Math.Max(p.Width, p.Height), Math.Min(p.Width, p.Height));
    }

    /// <summary>Document WPF d'un passage : une FixedPage par face (unités : 1/96 de pouce).</summary>
    public static FixedDocument BuildDocument(OutputPages pages, IEnumerable<PrintFace> faces, PrintMode mode, Size? paper = null)
    {
        var doc = new FixedDocument();
        bool booklet = mode is PrintMode.Booklet or PrintMode.BookletManualDuplex;
        var first = pages.SizeOf(1);
        doc.DocumentPaginator.PageSize = booklet ? BookletSheet(paper) : new Size(first.Width * DipPerPoint, first.Height * DipPerPoint);
        foreach (var face in faces)
        {
            FixedPage fp;
            if (face.TwoUp)
            {
                var sheet = BookletSheet(paper);
                fp = new FixedPage { Width = sheet.Width, Height = sheet.Height, Background = Brushes.White };
                // deux pages côte à côte, mises à l'échelle de la demi-feuille, centrées
                double pw = first.Width * DipPerPoint, ph = first.Height * DipPerPoint;
                double scale = Math.Min(sheet.Width / 2 / pw, sheet.Height / ph);
                double ox = (sheet.Width - 2 * pw * scale) / 2, oy = (sheet.Height - ph * scale) / 2;
                int slot = 0;
                foreach (var n in new[] { face.Left, face.Right })
                {
                    if (n > 0)
                    {
                        var view = new PrintedPage(pages.Render(n), pages.SizeOf(n), scale) { Tag = n };
                        FixedPage.SetLeft(view, ox + slot * pw * scale);
                        FixedPage.SetTop(view, oy);
                        fp.Children.Add(view);
                    }
                    slot++;
                }
            }
            else
            {
                // chaque page garde son format (une section en paysage sort en paysage)
                var size = pages.SizeOf(face.Left > 0 ? face.Left : 1);
                fp = new FixedPage { Width = size.Width * DipPerPoint, Height = size.Height * DipPerPoint, Background = Brushes.White };
                if (face.Left > 0) fp.Children.Add(new PrintedPage(pages.Render(face.Left), size, 1) { Tag = face.Left });
            }
            fp.Tag = face;
            var pc = new PageContent();
            ((System.Windows.Markup.IAddChild)pc).AddChild(fp);
            doc.Pages.Add(pc);
        }
        return doc;
    }

    /// <summary>Faces d'un document construit (dans l'ordre), pour l'auto-test et l'aperçu.</summary>
    public static List<PrintFace> FacesOf(FixedDocument doc) => doc.Pages.Select(p => (PrintFace)p.Child.Tag).ToList();

    /// <summary>
    /// Imprime : boîte d'impression de Windows (choix de l'imprimante), puis un ou deux passages.
    /// <paramref name="confirmSecondPass"/> affiche la consigne « retournez la pile » et renvoie faux pour abandonner.
    /// </summary>
    public static bool Print(OutputPages pages, PrintOptions o, string title, Func<bool> confirmSecondPass)
    {
        var dlg = new PrintDialog { UserPageRangeEnabled = false };
        bool booklet = o.Mode is PrintMode.Booklet or PrintMode.BookletManualDuplex;
        var first = pages.SizeOf(o.Pages.Count > 0 ? o.Pages[0] : 1);
        try
        {
            if (booklet) dlg.PrintTicket.PageOrientation = PageOrientation.Landscape;
            else
            {
                dlg.PrintTicket.PageMediaSize = new PageMediaSize(first.Width * DipPerPoint, first.Height * DipPerPoint);
                dlg.PrintTicket.PageOrientation = first.Width > first.Height ? PageOrientation.Landscape : PageOrientation.Portrait;
            }
            if (o.Mode == PrintMode.Booklet) dlg.PrintTicket.Duplexing = Duplexing.TwoSidedShortEdge;
            dlg.PrintTicket.CopyCount = 1;   // les exemplaires sont dans le document (assemblés)
        }
        catch (Exception) { /* imprimante sans ces réglages : valeurs par défaut */ }
        if (dlg.ShowDialog() != true) return false;
        Size? paper = null;
        if (dlg.PrintTicket.PageMediaSize is { Width: { } mw, Height: { } mh }) paper = new Size(mw, mh);
        var passes = Passes(o);
        for (int i = 0; i < passes.Count; i++)
        {
            if (i > 0 && !confirmSecondPass()) return false;
            if (passes[i].Count == 0) continue;
            var doc = BuildDocument(pages, passes[i], o.Mode, paper);
            dlg.PrintDocument(doc.DocumentPaginator, passes.Count > 1 ? $"{title} ({(i == 0 ? "rectos" : "versos")})" : title);
        }
        return true;
    }
}
