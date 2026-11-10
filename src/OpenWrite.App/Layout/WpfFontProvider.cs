using System.IO;
using System.Windows;
using System.Windows.Media;
using OpenWrite.Layout;

namespace OpenWrite.App.Layout;

/// <summary>
/// Polices pour la mise en page, lues par WPF (GlyphTypeface) : avances des glyphes (sans crénage, comme Word par
/// défaut) et métriques verticales du fichier de police (formule GDI de la hauteur de ligne, voir <see cref="FontVerticalMetrics"/>).
/// Une police absente est remplacée par une police métriquement compatible si elle est installée (Carlito pour
/// Calibri, Caladea pour Cambria…), sinon par Arial ; les remplacements sont relevés dans <see cref="Substitutions"/>.
/// </summary>
public sealed class WpfFontProvider : IFontProvider
{
    static readonly Dictionary<string, string[]> Compatible = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Calibri"] = new[] { "Carlito" },
        ["Calibri Light"] = new[] { "Carlito" },
        ["Cambria"] = new[] { "Caladea" },
        ["Arial"] = new[] { "Liberation Sans" },
        ["Times New Roman"] = new[] { "Liberation Serif" },
        ["Courier New"] = new[] { "Liberation Mono" },
        ["Aptos"] = new[] { "Calibri", "Carlito" },
    };

    readonly Dictionary<(string, bool, bool), FontFace> _faces = new();
    Dictionary<string, List<GlyphTypeface>>? _cloud;

    /// <summary>
    /// Polices « cloud » d'Office (Aptos…), téléchargées par Word dans %LOCALAPPDATA%\Microsoft\FontCache\4\CloudFonts\&lt;famille&gt;
    /// mais pas installées dans Windows : WPF ne les voit pas, on les charge depuis leurs fichiers.
    /// </summary>
    Dictionary<string, List<GlyphTypeface>> Cloud()
    {
        if (_cloud is not null) return _cloud;
        _cloud = new Dictionary<string, List<GlyphTypeface>>(StringComparer.OrdinalIgnoreCase);
        var root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "FontCache", "4", "CloudFonts");
        try
        {
            if (!Directory.Exists(root)) return _cloud;
            foreach (var dir in Directory.GetDirectories(root))
            {
                var list = new List<GlyphTypeface>();
                foreach (var file in Directory.GetFiles(dir, "*.ttf"))
                {
                    try { list.Add(new GlyphTypeface(new Uri(file))); } catch (Exception) { }
                }
                if (list.Count > 0) _cloud[Path.GetFileName(dir)] = list;
            }
        }
        catch (Exception) { }
        return _cloud;
    }

    /// <summary>Variante la plus proche (style puis graisse) d'une police cloud.</summary>
    GlyphTypeface? CloudFace(string family, int weight, bool italic)
    {
        if (!Cloud().TryGetValue(family, out var list)) return null;
        return list.OrderBy(g => (g.Style == FontStyles.Normal) == italic ? 1 : 0)
            .ThenBy(g => Math.Abs(g.Weight.ToOpenTypeWeight() - weight))
            .ThenBy(g => g.StyleSimulations != StyleSimulations.None ? 1 : 0)
            .FirstOrDefault();
    }
    readonly HashSet<string> _installed = Fonts.SystemFontFamilies.Select(f => f.Source).ToHashSet(StringComparer.OrdinalIgnoreCase);

    IReadOnlyDictionary<string, string> _altNames = new Dictionary<string, string>();

    /// <summary>
    /// Remplaçants proposés par le document ouvert (fontTable.xml, w:altName), essayés en premier pour une police absente,
    /// comme Word (Syllabus : ZapfChancery → Monotype Corsiva).
    /// </summary>
    public IReadOnlyDictionary<string, string> AltNames
    {
        get => _altNames;
        set { _altNames = value; _faces.Clear(); }
    }

    /// <summary>Police demandée → police utilisée, pour les polices absentes.</summary>
    public Dictionary<string, string> Substitutions { get; } = new(StringComparer.OrdinalIgnoreCase);

    static readonly (string Suffix, int Weight)[] WeightSuffixes =
    {
        (" Thin", 100), (" ExtraLight", 200), (" Light", 300), (" Medium", 500), (" SemiBold", 600), (" Semibold", 600),
        (" ExtraBold", 800), (" Black", 900), ("-Bold", 700), ("-Regular", 400),
    };

    public FontFace Face(string? family, bool bold, bool italic)
    {
        family ??= "Times New Roman";
        var key = (family, bold, italic);
        if (_faces.TryGetValue(key, out var face)) return face;
        // « Calibri Light » : pour WPF, une graisse de la famille « Calibri » ; « Calibri-Bold », « X-Italic » :
        // noms de polices PostScript (documents convertis depuis un PDF)
        string name = family;
        int weight = bold ? 700 : 400;
        bool ital = italic;
        if (!_installed.Contains(name))
        {
            if (name.EndsWith("-Italic", StringComparison.OrdinalIgnoreCase)) { name = name[..^7]; ital = true; }
            else if (name.EndsWith("-BoldItalic", StringComparison.OrdinalIgnoreCase)) { name = name[..^11]; ital = true; weight = 700; }
            foreach (var (suffix, w) in WeightSuffixes)
                if (name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) && _installed.Contains(name[..^suffix.Length]))
                {
                    name = name[..^suffix.Length];
                    weight = bold && w < 700 ? 700 : w;
                    break;
                }
        }
        string used = name;
        if (!_installed.Contains(name) && CloudFace(name, weight, ital) is { } cloud)
        {
            face = new GlyphFace(cloud, family, substituted: false);
            _faces[key] = face;
            return face;
        }
        if (!_installed.Contains(name))
        {
            used = _altNames.TryGetValue(name, out var docAlt) && _installed.Contains(docAlt) ? docAlt
                : Compatible.TryGetValue(name, out var alts) ? alts.FirstOrDefault(_installed.Contains) ?? "Arial" : "Arial";
            Substitutions[family] = used;
        }
        var style = ital ? FontStyles.Italic : FontStyles.Normal;
        if (!new Typeface(new FontFamily(used), style, FontWeight.FromOpenTypeWeight(weight), FontStretches.Normal).TryGetGlyphTypeface(out var gt))
        {
            used = "Arial";
            new Typeface(new FontFamily(used), style, FontWeight.FromOpenTypeWeight(weight), FontStretches.Normal).TryGetGlyphTypeface(out gt);
        }
        face = new GlyphFace(gt!, used == name ? family : used, used != name);
        _faces[key] = face;
        return face;
    }

    /// <summary>Police WPF : avances pour la mise en page, glyphes pour l'affichage (PageRenderer).</summary>
    public sealed class GlyphFace : FontFace
    {
        public GlyphTypeface Typeface { get; }

        /// <summary>Glyphe du caractère (polices de symboles : plage F000–F0FF) ; null s'il manque à la police.</summary>
        public ushort? Glyph(char c) =>
            _map.TryGetValue(c, out var g) || c < 0x100 && _map.TryGetValue(0xF000 + c, out g) ? g : null;

        readonly IDictionary<int, ushort> _map;
        readonly IDictionary<ushort, double> _advances;
        readonly double[] _fast = new double[0x500];
        readonly Dictionary<char, double> _slow = new();

        [System.Diagnostics.CodeAnalysis.SetsRequiredMembers]
        public GlyphFace(GlyphTypeface gt, string name, bool substituted)
        {
            Name = name;
            Substituted = substituted;
            Typeface = gt;
            _map = gt.CharacterToGlyphMap;
            _advances = gt.AdvanceWidths;
            Array.Fill(_fast, double.NaN);
            FontVerticalMetrics m = default;
            bool ok = false;
            try
            {
                using var s = gt.GetFontStream();
                using var ms = new MemoryStream();
                s.CopyTo(ms);
                ok = FontVerticalMetrics.TryRead(ms.ToArray(), name, out m);
            }
            catch (Exception) { ok = false; }
            if (ok)
            {
                var (a, dsc, lead) = m.WordMetrics;
                Ascent = a / (double)m.UnitsPerEm;
                Descent = dsc / (double)m.UnitsPerEm;
                ExternalLeading = lead / (double)m.UnitsPerEm;
            }
            else
            {
                Ascent = gt.Baseline;
                Descent = gt.Height - gt.Baseline;
            }
        }

        public override double Advance(char c)
        {
            if (c < _fast.Length)
            {
                var v = _fast[c];
                if (double.IsNaN(v)) _fast[c] = v = Lookup(c);
                return v;
            }
            if (!_slow.TryGetValue(c, out var w)) _slow[c] = w = Lookup(c);
            return w;
        }

        double Lookup(char c)
        {
            if (c == '­') return 0;
            // police de symboles (Symbol, Wingdings) : caractères rangés en F000–F0FF
            if (_map.TryGetValue(c, out var g) || c < 0x100 && _map.TryGetValue(0xF000 + c, out g)) return _advances[g];
            return 0.6;      // glyphe absent (remplacé à l'affichage par une police de secours)
        }
    }
}
