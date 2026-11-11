// Repris d'OpenPub 2.1.4 (OpenPub.App/Output/PdfExporter.cs) : export direct par PDFsharp, sans les réglages pour
// l'imprimeur (fond perdu, traits de coupe, PDF/X-4) qui n'ont pas de sens pour un traitement de texte.
using System.IO;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace OpenWrite.App.Output;

/// <summary>Réglages de l'export PDF.</summary>
public sealed record PdfExportOptions
{
    /// <summary>Pages à exporter (numéros à partir de 1) ; null = toutes.</summary>
    public IReadOnlyList<int>? Pages { get; init; }
    public string? Title { get; init; }
    public string? Author { get; init; }
    /// <summary>Résolution des parties rastérisées (dégradés, effets).</summary>
    public double RasterDpi { get; init; } = 300;
}

/// <summary>
/// Export PDF direct (PDFsharp, licence MIT), sans imprimante virtuelle : chaque page est dessinée par
/// <see cref="OutputPages.Render"/> puis traduite par <see cref="DrawingToPdf"/> (texte avec les vraies polices
/// incorporées en sous-ensemble, traits vectoriels, images d'origine) ; les liens hypertextes restent cliquables.
/// Signets, liens internes et PDF balisé : lot 12.
/// </summary>
public static class PdfExporter
{
    public static PdfStats Export(OutputPages pages, string path, PdfExportOptions o, Func<System.Windows.Media.Imaging.BitmapSource, byte[]?>? original = null)
    {
        PdfFontResolver.Install();
        var stats = new PdfStats();
        var doc = new PdfDocument();
        doc.Options.CompressContentStreams = true;
        doc.Options.FlateEncodeMode = PdfFlateEncodeMode.BestCompression;
        var now = DateTime.Now;
        doc.Info.Title = o.Title ?? Path.GetFileNameWithoutExtension(path);
        if (!string.IsNullOrEmpty(o.Author)) doc.Info.Author = o.Author;
        doc.Info.Creator = "OpenWrite " + Update.UpdateService.CurrentVersion;
        doc.Info.CreationDate = now;
        doc.Info.ModificationDate = now;
        var fonts = new PdfFontCache();
        var images = new PdfImageCache { Original = original };
        var numbers = o.Pages ?? Enumerable.Range(1, pages.Count).ToList();
        foreach (var n in numbers)
        {
            var size = pages.SizeOf(n);
            var pdfPage = doc.AddPage();
            pdfPage.Width = XUnit.FromPoint(size.Width);
            pdfPage.Height = XUnit.FromPoint(size.Height);
            var drawing = pages.Render(n);
            using (var g = XGraphics.FromPdfPage(pdfPage, XGraphicsUnit.Point))
            {
                var writer = new DrawingToPdf(g, doc, fonts, images, stats) { RasterDpi = o.RasterDpi };
                try { writer.Draw(drawing); }
                finally { writer.Detach(doc); }
            }
            foreach (var (r, url) in pages.Links(n))
            {
                pdfPage.AddWebLink(new PdfRectangle(new XRect(r.X, size.Height - r.Y - r.Height, r.Width, r.Height)), url);
                stats.Links++;
            }
        }
        stats.OriginalImages = images.Originals;
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        using var ms = new MemoryStream();
        doc.Save(ms, false);
        // écrit d'un bloc : un PDF ouvert dans un lecteur fait échouer l'écriture sans laisser de fichier tronqué
        File.WriteAllBytes(path, ms.ToArray());
        return stats;
    }

    /// <summary>Résumé d'un PDF relu par PDFsharp (diagnostic, auto-test).</summary>
    public sealed record PdfSummary(int Pages, List<(double W, double H)> Sizes, HashSet<string> EmbeddedFonts, HashSet<string> MissingFonts,
        int Images, int Links, string? Title, int FontsWithoutUnicode);

    public static PdfSummary Inspect(string path)
    {
        using var doc = PdfReader.Open(path, PdfDocumentOpenMode.Import);
        var sizes = new List<(double, double)>();
        var embedded = new HashSet<string>(StringComparer.Ordinal);
        var missing = new HashSet<string>(StringComparer.Ordinal);
        int images = 0, links = 0, noUnicode = 0;
        var seen = new HashSet<PdfDictionary>(ReferenceEqualityComparer.Instance);
        foreach (var p in doc.Pages)
        {
            sizes.Add((Math.Round(p.Width.Point, 2), Math.Round(p.Height.Point, 2)));
            var res = p.Elements.GetDictionary("/Resources");
            if (res?.Elements.GetDictionary("/Font") is { } fontDict)
                foreach (var key in fontDict.Elements.Keys)
                {
                    if (fontDict.Elements.GetDictionary(key) is not { } f || !seen.Add(f)) continue;
                    var name = f.Elements.GetName("/BaseFont");
                    // police composite (Type0) : descripteur dans la police descendante
                    var desc = f.Elements.GetDictionary("/FontDescriptor");
                    if (desc is null && f.Elements.GetArray("/DescendantFonts") is { Elements.Count: > 0 } arr && arr.Elements.GetDictionary(0) is { } df)
                        desc = df.Elements.GetDictionary("/FontDescriptor");
                    bool emb = desc is not null && (desc.Elements.ContainsKey("/FontFile2") || desc.Elements.ContainsKey("/FontFile") || desc.Elements.ContainsKey("/FontFile3"));
                    (emb ? embedded : missing).Add(name.TrimStart('/'));
                    if (!f.Elements.ContainsKey("/ToUnicode")) noUnicode++;
                }
            if (res?.Elements.GetDictionary("/XObject") is { } xo)
                foreach (var key in xo.Elements.Keys)
                    if (xo.Elements.GetDictionary(key) is { } x && x.Elements.GetName("/Subtype") == "/Image" && seen.Add(x)) images++;
            if (p.Elements.GetArray("/Annots") is { } annots)
                for (int i = 0; i < annots.Elements.Count; i++)
                    if (annots.Elements.GetDictionary(i) is { } a && a.Elements.GetName("/Subtype") == "/Link") links++;
        }
        var title = doc.Info.Title;
        return new PdfSummary(doc.PageCount, sizes, embedded, missing, images, links, string.IsNullOrEmpty(title) ? null : title, noUnicode);
    }
}
