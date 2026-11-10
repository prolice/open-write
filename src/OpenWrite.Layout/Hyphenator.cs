// Repris d'OpenPub 2.1.3 (src/OpenPub.Layout/Hyphenator.cs).
using System.Globalization;
using System.Reflection;
using System.Text;

namespace OpenWrite.Layout;

/// <summary>
/// Cesure par motifs de Liang (format des dictionnaires LibreOffice hyph_xx.dic).
/// Le dictionnaire francais embarque provient de LibreOffice (via pyphen).
/// </summary>
public sealed class Hyphenator
{
    readonly Dictionary<string, byte[]> _patterns = new();
    readonly int _maxLen;
    public int LeftMin { get; } = 2;
    public int RightMin { get; } = 3;      // comme le prototype : pas de coupure laissant 2 lettres seules
    public int MinWordLength { get; set; } = 6;

    Hyphenator(IEnumerable<string> lines)
    {
        foreach (var raw in lines)
        {
            var line = raw.Trim();
            if (line.Length == 0 || line.StartsWith('%') || line == "NEXTLEVEL" || line == "UTF-8" ||
                line.StartsWith("LEFTHYPHENMIN") || line.StartsWith("RIGHTHYPHENMIN") ||
                line.StartsWith("COMPOUND") || line.StartsWith("ISO8859"))
                continue;
            var slash = line.IndexOf('/');
            if (slash >= 0) line = line[..slash];       // regles de remplacement non gerees
            var letters = new StringBuilder();
            var values = new List<byte> { 0 };
            foreach (var c in line)
            {
                if (char.IsDigit(c)) values[^1] = (byte)(c - '0');
                else { letters.Append(c); values.Add(0); }
            }
            var key = letters.ToString();
            if (key.Length == 0) continue;
            if (!_patterns.TryGetValue(key, out var existing))
                _patterns[key] = values.ToArray();
            else
                for (int i = 0; i < existing.Length && i < values.Count; i++)
                    existing[i] = Math.Max(existing[i], values[i]);
            _maxLen = Math.Max(_maxLen, key.Length);
        }
    }

    public static Hyphenator FromLines(IEnumerable<string> lines) => new(lines);

    static Hyphenator? _french;
    public static Hyphenator French => _french ??= LoadEmbedded("hyph_fr.dic");

    static Hyphenator LoadEmbedded(string name)
    {
        var asm = Assembly.GetExecutingAssembly();
        var res = asm.GetManifestResourceNames().First(n => n.EndsWith(name, StringComparison.OrdinalIgnoreCase));
        using var s = asm.GetManifestResourceStream(res)!;
        using var r = new StreamReader(s, Encoding.UTF8);
        var lines = new List<string>();
        while (r.ReadLine() is { } l) lines.Add(l);
        return new Hyphenator(lines);
    }

    readonly Dictionary<string, int[]> _cache = new();

    /// <summary>Positions de coupure possibles (indices dans le mot : coupure avant word[i]).</summary>
    public int[] Points(string word)
    {
        if (word.Length < MinWordLength) return Array.Empty<int>();
        if (_cache.TryGetValue(word, out var cached)) return cached;
        var lower = word.ToLower(CultureInfo.GetCultureInfo("fr-FR")).Replace('’', '\'');
        var w = "." + lower + ".";
        var values = new byte[w.Length + 1];
        for (int i = 0; i < w.Length; i++)
            for (int len = 1; len <= _maxLen && i + len <= w.Length; len++)
                if (_patterns.TryGetValue(w.Substring(i, len), out var pv))
                    for (int k = 0; k < pv.Length; k++)
                        values[i + k] = Math.Max(values[i + k], pv[k]);
        var points = new List<int>();
        for (int i = LeftMin; i <= word.Length - RightMin; i++)
            if (values[i + 1] % 2 == 1 && char.IsLetter(word[i - 1]) && char.IsLetter(word[i]))
                points.Add(i);
        var res = points.ToArray();
        _cache[word] = res;
        return res;
    }

    public string Insert(string word, string hyphen = "-")
    {
        var pts = Points(word);
        var sb = new StringBuilder();
        int last = 0;
        foreach (var p in pts) { sb.Append(word, last, p - last).Append(hyphen); last = p; }
        return sb.Append(word, last, word.Length - last).ToString();
    }
}
