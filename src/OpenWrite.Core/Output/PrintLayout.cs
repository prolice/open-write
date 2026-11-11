// Repris d'OpenPub 2.1.4 (OpenPub.Core/Output/PrintLayout.cs) : choix des pages, livret, recto verso manuel (sans le fond perdu).
using System.Globalization;

namespace OpenWrite.Core.Output;

/// <summary>
/// Logique des sorties sans dependance a WPF : choix des pages, imposition en livret, recto verso manuel.
/// Unites : points (1/72 de pouce), origine en haut a gauche.
/// </summary>
public static class PageRange
{
    /// <summary>
    /// Lit une liste de pages « 1-3; 5, 8- » (numeros a partir de 1, dans l'ordre saisi, sans doublon).
    /// Vide ou « toutes » = toutes les pages. Renvoie null et un message en francais si la saisie est invalide.
    /// </summary>
    public static List<int>? Parse(string? text, int pageCount, out string? error)
    {
        error = null;
        var all = Enumerable.Range(1, Math.Max(0, pageCount)).ToList();
        if (string.IsNullOrWhiteSpace(text) || text.Trim().Equals("toutes", StringComparison.OrdinalIgnoreCase)) return all;
        var result = new List<int>();
        foreach (var raw in text.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries))
        {
            var part = raw.Trim();
            int a, b;
            var dash = part.IndexOfAny(new[] { '-', '–', '—' });
            if (dash < 0)
            {
                if (!int.TryParse(part, NumberStyles.None, CultureInfo.InvariantCulture, out a)) { error = $"« {part} » n'est pas un numéro de page."; return null; }
                b = a;
            }
            else
            {
                var left = part[..dash].Trim();
                var right = part[(dash + 1)..].Trim();
                a = left.Length == 0 ? 1 : int.TryParse(left, NumberStyles.None, CultureInfo.InvariantCulture, out var l) ? l : -1;
                b = right.Length == 0 ? pageCount : int.TryParse(right, NumberStyles.None, CultureInfo.InvariantCulture, out var r) ? r : -1;
                if (a < 0 || b < 0) { error = $"« {part} » n'est pas une plage de pages valide (exemple : 2-4)."; return null; }
                if (b < a) (a, b) = (b, a);
            }
            if (a < 1 || b > pageCount) { error = $"La page {(a < 1 ? a : b)} n'existe pas : le document a {pageCount} page(s)."; return null; }
            for (int p = a; p <= b; p++) if (!result.Contains(p)) result.Add(p);
        }
        if (result.Count == 0) { error = "Aucune page choisie."; return null; }
        return result;
    }
}

/// <summary>Une face de feuille : deux pages cote a cote (0 = page blanche).</summary>
public readonly record struct SheetSide(int Sheet, bool Front, int Left, int Right);

/// <summary>Imposition en livret : feuilles pliees en deux, 2 pages par face, dans l'ordre de pliage.</summary>
public static class Booklet
{
    /// <summary>Nombre de pages du livret : multiple de 4 superieur ou egal (pages blanches ajoutees a la fin).</summary>
    public static int PaddedCount(int pages) => pages <= 0 ? 0 : (pages + 3) / 4 * 4;

    /// <summary>
    /// Faces a imprimer, dans l'ordre (recto puis verso de chaque feuille). Pour 8 pages : 8-1, 2-7, 6-3, 4-5.
    /// Les numeros superieurs a <paramref name="pages"/> valent 0 (page blanche).
    /// </summary>
    public static List<SheetSide> Sides(int pages)
    {
        int n = PaddedCount(pages);
        int Real(int p) => p <= pages ? p : 0;
        var sides = new List<SheetSide>();
        for (int s = 0; s < n / 4; s++)
        {
            sides.Add(new SheetSide(s + 1, true, Real(n - 2 * s), Real(1 + 2 * s)));
            sides.Add(new SheetSide(s + 1, false, Real(2 + 2 * s), Real(n - 1 - 2 * s)));
        }
        return sides;
    }
}

/// <summary>Recto verso manuel : d'abord toutes les faces avant, puis (pile retournee) toutes les faces arriere.</summary>
public static class ManualDuplex
{
    /// <summary>
    /// Separe une suite de faces (recto, verso, recto, verso...) en deux passages. Les versos sont completes par
    /// une face blanche (<paramref name="blank"/>) pour que chaque feuille de la pile en recoive une ;
    /// <paramref name="reverseBacks"/> inverse l'ordre du second passage (imprimantes qui empilent face dessus).
    /// </summary>
    public static (List<T> Fronts, List<T> Backs) Split<T>(IReadOnlyList<T> faces, T blank, bool reverseBacks = false)
    {
        var fronts = new List<T>();
        var backs = new List<T>();
        for (int i = 0; i < faces.Count; i++) (i % 2 == 0 ? fronts : backs).Add(faces[i]);
        while (backs.Count < fronts.Count) backs.Add(blank);
        if (reverseBacks) backs.Reverse();
        return (fronts, backs);
    }
}
