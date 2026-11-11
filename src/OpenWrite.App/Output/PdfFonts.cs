// Repris d'OpenPub 2.1.4 (OpenPub.App/Output/PdfFonts.cs).
using System.IO;
using System.Windows.Media;
using OpenWrite.Core.Output;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;

namespace OpenWrite.App.Output;

/// <summary>
/// Polices du PDF : chaque police WPF rencontree dans le rendu (GlyphTypeface d'un GlyphRun : police systeme ou
/// police du dossier fonts/ du paquet) est enregistree sous un nom unique ; PDFsharp la relit ici depuis le vrai
/// fichier et l'incorpore en sous-ensemble.
/// </summary>
public sealed class PdfFontResolver : IFontResolver
{
    public static readonly PdfFontResolver Instance = new();
    static readonly object Gate = new();
    static bool _installed;

    readonly Dictionary<string, byte[]> _faces = new();
    readonly Dictionary<string, (string Face, bool Italic)> _families = new();
    readonly Dictionary<string, string?> _byKey = new();

    /// <summary>Installe le resolveur une fois pour toutes (PDFsharp n'en accepte qu'un par processus).</summary>
    public static void Install()
    {
        lock (Gate)
        {
            if (_installed) return;
            GlobalFontSettings.FontResolver = Instance;
            _installed = true;
        }
    }

    public FontResolverInfo? ResolveTypeface(string familyName, bool bold, bool italic)
    {
        lock (_faces)
        {
            if (_families.TryGetValue(familyName, out var f))
                return new FontResolverInfo(f.Face, false, f.Italic);
            // police demandee par nom (rare : PDFsharp lui-meme) : premiere police enregistree
            var any = _families.Values.FirstOrDefault();
            return any.Face is null ? null : new FontResolverInfo(any.Face);
        }
    }

    public byte[]? GetFont(string faceName)
    {
        lock (_faces) return _faces.TryGetValue(faceName, out var b) ? b : null;
    }

    /// <summary>
    /// Nom de famille PDFsharp pour une police WPF, ou null si elle ne peut pas etre incorporee comme texte
    /// (contours CFF, licence interdisant l'incorporation, fichier illisible) : le texte est alors dessine en contours vectoriels.
    /// </summary>
    public string? Register(GlyphTypeface gt)
    {
        var uri = gt.FontUri;
        var en = System.Globalization.CultureInfo.GetCultureInfo("en-US");
        string? fam = gt.Win32FamilyNames.TryGetValue(en, out var fn) ? fn : gt.Win32FamilyNames.Values.FirstOrDefault();
        string? sty = gt.Win32FaceNames.TryGetValue(en, out var sn) ? sn : gt.Win32FaceNames.Values.FirstOrDefault();
        // une collection .ttc contient plusieurs polices dans un meme fichier : la cle inclut famille et style
        var key = uri.AbsoluteUri + "|" + fam + "|" + sty;
        lock (_faces)
        {
            if (!_byKey.TryGetValue(key, out var face))
            {
                face = null;
                try
                {
                    byte[] data;
                    using (var s = gt.GetFontStream())
                    using (var ms = new MemoryStream())
                    {
                        s.CopyTo(ms);
                        data = ms.ToArray();
                    }
                    if (SfntFont.IsCollection(data))
                    {
                        int index = SfntFont.FindFace(data, fam, sty);
                        data = SfntFont.ExtractFace(data, index);
                    }
                    if (!SfntFont.HasCffOutlines(data) && !SfntFont.EmbeddingRestricted(data) && SfntFont.Tables(data).Any(t => t.Tag == "glyf"))
                    {
                        face = "openwrite-" + _faces.Count;
                        _faces[face] = data;
                    }
                }
                catch (Exception) { face = null; }
                _byKey[key] = face;
            }
            if (face is null) return null;
            // l'italique simule par WPF (police sans variante italique) est reproduit par PDFsharp
            bool italic = (gt.StyleSimulations & StyleSimulations.ItalicSimulation) != 0;
            var family = italic ? face + "-i" : face;
            _families[family] = (face, italic);
            return family;
        }
    }
}

/// <summary>Polices XFont par taille, pour un document PDF.</summary>
public sealed class PdfFontCache
{
    readonly Dictionary<(string, double), XFont> _fonts = new();
    static readonly XPdfFontOptions Options = new(PdfFontEncoding.Unicode, PdfFontEmbedding.TryComputeSubset);

    public XFont Get(string family, double size)
    {
        var key = (family, Math.Round(size, 4));
        if (!_fonts.TryGetValue(key, out var f)) _fonts[key] = f = new XFont(family, size, XFontStyleEx.Regular, Options);
        return f;
    }
}
