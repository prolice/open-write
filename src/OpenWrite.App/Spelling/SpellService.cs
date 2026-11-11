// Repris d'OpenPub 2.1.5 (OpenPub.App/Editing/SpellService.cs). OpenWrite : langue par défaut de Windows, langue non
// installée = aucune vérification (et non le français), liste des langues prises en charge, langue voisine (fr-BE → fr-FR).
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace OpenWrite.App.Spelling;

/// <summary>
/// Correcteur orthographique de Windows (API Spell Checking, spellcheck.h) : dictionnaires fournis par Windows (aucun
/// fichier à embarquer), dictionnaire personnel de l'utilisateur partagé avec les autres applications. Un vérificateur
/// par langue ; résultats gardés par texte.
/// </summary>
public sealed class SpellService
{
    readonly ISpellCheckerFactory _factory;
    readonly Dictionary<string, ISpellChecker?> _checkers = new(StringComparer.OrdinalIgnoreCase);
    readonly Dictionary<(string, string), List<(int Start, int Length)>> _cache = new();
    readonly HashSet<string> _ignored = new(StringComparer.Ordinal);
    List<string>? _supported;

    SpellService(ISpellCheckerFactory f) => _factory = f;

    /// <summary>Service partagé par les fenêtres (null si l'API de Windows est absente).</summary>
    public static readonly Lazy<SpellService?> Shared = new(TryCreate);

    /// <summary>Service disponible (Windows 8 ou plus) ; null si l'API est absente.</summary>
    public static SpellService? TryCreate()
    {
        try
        {
            var type = Type.GetTypeFromCLSID(new Guid("7AB36653-1796-484B-BDFA-E74F1DB7C1DC"));
            if (type is null || Activator.CreateInstance(type) is not ISpellCheckerFactory f) return null;
            return new SpellService(f);
        }
        catch { return null; }
    }

    /// <summary>Langue par défaut quand le document n'en indique aucune : celle de Windows.</summary>
    public static string DefaultTag => CultureInfo.CurrentUICulture.Name is { Length: > 0 } n ? n : "fr-FR";

    /// <summary>Langues dont Windows a le dictionnaire (« fr-FR », « en-US »…).</summary>
    public IReadOnlyList<string> SupportedLanguages
    {
        get
        {
            if (_supported is not null) return _supported;
            _supported = new List<string>();
            try
            {
                _factory.get_SupportedLanguages(out var e);
                var buf = new string[1];
                while (e.Next(1, buf, IntPtr.Zero) == 0) _supported.Add(buf[0]);
            }
            catch (COMException) { }
            return _supported;
        }
    }

    /// <summary>Langue vérifiable pour une étiquette du document : elle-même, ou une variante de la même langue.</summary>
    public string? Resolve(string? lang)
    {
        var tag = string.IsNullOrEmpty(lang) ? DefaultTag : lang;
        if (Checker(tag) is not null) return tag;
        return null;
    }

    ISpellChecker? Checker(string tag)
    {
        if (_checkers.TryGetValue(tag, out var c)) return c;
        c = Create(tag);
        if (c is null)
        {
            // fr-BE, nl-BE, en-BE… : le dictionnaire de la même langue (fr-FR de préférence pour « fr »)
            var primary = tag.Split('-')[0];
            var near = SupportedLanguages.Where(s => s.Split('-')[0].Equals(primary, StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(s => s.Equals(primary + "-" + primary, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            if (near is not null && !near.Equals(tag, StringComparison.OrdinalIgnoreCase)) c = Checker(near);
        }
        // langue sans dictionnaire installé : rien n'est souligné (comme Word sans les outils de vérification)
        _checkers[tag] = c;
        return c;
    }

    ISpellChecker? Create(string tag)
    {
        try
        {
            _factory.IsSupported(tag, out int ok);
            if (ok != 0) { _factory.CreateSpellChecker(tag, out var sc); return sc; }
        }
        catch (Exception) { }
        return null;
    }

    /// <summary>Mots mal orthographiés (début, longueur) dans un texte, pour une langue du document.</summary>
    public List<(int Start, int Length)> Check(string text, string? lang)
    {
        var tag = string.IsNullOrEmpty(lang) ? DefaultTag : lang;
        if (_cache.TryGetValue((tag, text), out var res)) return res;
        res = new List<(int, int)>();
        if (text.Trim().Length > 0 && Checker(tag) is { } c)
        {
            try
            {
                // les caractères de contrôle (saut de ligne, tabulation) deviennent des espaces : positions inchangées
                var clean = new string(text.Select(ch => char.IsControl(ch) || ch is '­' or '￼' ? ' ' : ch).ToArray());
                c.Check(clean, out var errors);
                while (errors.Next(out var err) == 0 && err is not null)
                {
                    err.get_StartIndex(out uint st);
                    err.get_Length(out uint len);
                    if (st + len > clean.Length || len == 0) continue;
                    var word = clean.Substring((int)st, (int)len);
                    if (!_ignored.Contains(word)) res.Add(((int)st, (int)len));
                }
            }
            catch (COMException) { /* vérificateur indisponible : pas de soulignement */ }
        }
        if (_cache.Count > 4000) _cache.Clear();
        _cache[(tag, text)] = res;
        return res;
    }

    public List<string> Suggest(string word, string? lang)
    {
        var list = new List<string>();
        if (Checker(string.IsNullOrEmpty(lang) ? DefaultTag : lang) is not { } c) return list;
        try
        {
            c.Suggest(word, out var e);
            var buf = new string[1];
            while (list.Count < 8 && e.Next(1, buf, IntPtr.Zero) == 0) list.Add(buf[0]);
        }
        catch (COMException) { }
        return list;
    }

    /// <summary>Ignorer ce mot jusqu'à la fermeture d'OpenWrite (Word : « Ignorer tout »).</summary>
    public void Ignore(string word)
    {
        _ignored.Add(word);
        _cache.Clear();
    }

    /// <summary>Ajouter au dictionnaire personnel de Windows (partagé avec les autres applications).</summary>
    public void Add(string word, string? lang)
    {
        try { Checker(string.IsNullOrEmpty(lang) ? DefaultTag : lang)?.Add(word); } catch (COMException) { }
        _cache.Clear();
    }

    // ------------------------------------------------------------------ interop (spellcheck.h)

    [ComImport, Guid("8E018A9D-2415-4677-BF08-794EA61F94BB"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface ISpellCheckerFactory
    {
        void get_SupportedLanguages(out IEnumString value);
        void IsSupported([MarshalAs(UnmanagedType.LPWStr)] string languageTag, out int value);
        void CreateSpellChecker([MarshalAs(UnmanagedType.LPWStr)] string languageTag, out ISpellChecker value);
    }

    [ComImport, Guid("B6FD0B71-E2BC-4653-8D05-F197E412770B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface ISpellChecker
    {
        void get_LanguageTag([MarshalAs(UnmanagedType.LPWStr)] out string value);
        void Check([MarshalAs(UnmanagedType.LPWStr)] string text, out IEnumSpellingError value);
        void Suggest([MarshalAs(UnmanagedType.LPWStr)] string word, out IEnumString value);
        void Add([MarshalAs(UnmanagedType.LPWStr)] string word);
        void Ignore([MarshalAs(UnmanagedType.LPWStr)] string word);
        void AutoCorrect([MarshalAs(UnmanagedType.LPWStr)] string from, [MarshalAs(UnmanagedType.LPWStr)] string to);
        void GetOptionValue([MarshalAs(UnmanagedType.LPWStr)] string optionId, out byte value);
        void get_OptionIds(out IEnumString value);
        void get_Id([MarshalAs(UnmanagedType.LPWStr)] out string value);
        void get_LocalizedName([MarshalAs(UnmanagedType.LPWStr)] out string value);
    }

    [ComImport, Guid("803E3BD4-2828-4410-8290-418D1D73C762"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IEnumSpellingError
    {
        [PreserveSig] int Next(out ISpellingError? value);
    }

    [ComImport, Guid("B7C82D61-FBE8-4B47-9B27-6C0D2E0DE0A3"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface ISpellingError
    {
        void get_StartIndex(out uint value);
        void get_Length(out uint value);
        void get_CorrectiveAction(out int value);
        void get_Replacement([MarshalAs(UnmanagedType.LPWStr)] out string value);
    }
}
