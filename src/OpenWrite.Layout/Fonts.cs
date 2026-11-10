namespace OpenWrite.Layout;

/// <summary>
/// Police résolue pour la mise en page : avances des caractères et métriques verticales, en fraction de l'em
/// (multiplier par la taille en points). Fournie par la couche d'affichage (WPF) ou par les tests (police à chasse fixe).
/// </summary>
public abstract class FontFace
{
    /// <summary>Nom de la police réellement utilisée (différent du nom demandé en cas de substitution).</summary>
    public required string Name { get; init; }
    /// <summary>La police demandée est absente : <see cref="Name"/> est une police de remplacement.</summary>
    public bool Substituted { get; init; }
    /// <summary>Hauteur au-dessus de la ligne de base (usWinAscent / unitsPerEm).</summary>
    public required double Ascent { get; init; }
    /// <summary>Hauteur sous la ligne de base (usWinDescent / unitsPerEm).</summary>
    public required double Descent { get; init; }
    /// <summary>Interligne externe à la manière de GDI (tmExternalLeading / unitsPerEm).</summary>
    public double ExternalLeading { get; init; }

    /// <summary>Hauteur d'une ligne simple de Word, en em : ascent + descent + interligne externe (formule GDI).</summary>
    public double LineHeight => Ascent + Descent + ExternalLeading;

    /// <summary>Avance du caractère, en em.</summary>
    public abstract double Advance(char c);

    public double Width(ReadOnlySpan<char> text, double sizePt)
    {
        double w = 0;
        foreach (var c in text) w += Advance(c);
        return w * sizePt;
    }
}

/// <summary>Accès aux polices (avec substitution quand une police manque).</summary>
public interface IFontProvider
{
    FontFace Face(string? family, bool bold, bool italic);
}

/// <summary>Police à chasse fixe pour les tests : chaque caractère avance de <see cref="Advance"/> em.</summary>
public sealed class MonoFace : FontFace
{
    public double CharAdvance { get; init; } = 0.5;
    public override double Advance(char c) => c == '­' ? 0 : CharAdvance;
}

public sealed class MonoFontProvider : IFontProvider
{
    readonly MonoFace _face;
    public MonoFontProvider(double advance = 0.5, double ascent = 0.8, double descent = 0.2, double leading = 0)
        => _face = new MonoFace { Name = "Mono", Ascent = ascent, Descent = descent, ExternalLeading = leading, CharAdvance = advance };
    public FontFace Face(string? family, bool bold, bool italic) => _face;
}

/// <summary>Métriques verticales lues dans les tables 'head', 'hhea' et 'OS/2' d'un fichier TrueType/OpenType (ou d'une collection .ttc).</summary>
public readonly record struct FontVerticalMetrics(int UnitsPerEm, int WinAscent, int WinDescent, int HheaAscender, int HheaDescender, int HheaLineGap)
{
    /// <summary>OS/2 fsSelection bit 7 (USE_TYPO_METRICS) : la police demande ses métriques typo.</summary>
    public bool UseTypoMetrics { get; init; }
    public int TypoAscender { get; init; }
    public int TypoDescender { get; init; }
    public int TypoLineGap { get; init; }

    /// <summary>
    /// Ascent, descent et interligne externe retenus par Word, en unités : métriques typo si la police les demande
    /// (Aptos : 1,2207 em, vérifié sur le corpus : 17,0 pt par ligne à 12 pt × 1,158), sinon formule GDI.
    /// </summary>
    public (int Ascent, int Descent, int Leading) WordMetrics =>
        UseTypoMetrics && TypoAscender - TypoDescender > 0 ? (TypoAscender, -TypoDescender, Math.Max(0, TypoLineGap)) : (WinAscent, WinDescent, ExternalLeading);

    /// <summary>
    /// tmExternalLeading de GDI : max(0, lineGap − ((winAscent + winDescent) − (ascender − descender))).
    /// Calibri : 0 (ligne 1,2207 em) ; Times New Roman : 87 unités (ligne 1,15 em : 13,8 pt à 12 pt, comme Word).
    /// </summary>
    public int ExternalLeading => Math.Max(0, HheaLineGap - ((WinAscent + WinDescent) - (HheaAscender - HheaDescender)));

    public static bool TryRead(byte[] d, string? family, out FontVerticalMetrics m)
    {
        m = default;
        try
        {
            if (d.Length < 12) return false;
            int start = 0;
            if (d[0] == 't' && d[1] == 't' && d[2] == 'c' && d[3] == 'f')
            {
                // collection : la police dont le nom de famille correspond, sinon la première
                int n = (int)U32(d, 8);
                start = (int)U32(d, 12);
                for (int i = 0; i < n; i++)
                {
                    int off = (int)U32(d, 12 + 4 * i);
                    if (family is not null && string.Equals(FamilyName(d, off), family, StringComparison.OrdinalIgnoreCase)) { start = off; break; }
                }
            }
            int head = Table(d, start, "head"), hhea = Table(d, start, "hhea"), os2 = Table(d, start, "OS/2");
            if (head < 0 || hhea < 0) return false;
            int upm = U16(d, head + 18);
            short asc = (short)U16(d, hhea + 4), desc = (short)U16(d, hhea + 6), gap = (short)U16(d, hhea + 8);
            int winAsc = os2 >= 0 ? U16(d, os2 + 74) : asc, winDesc = os2 >= 0 ? U16(d, os2 + 76) : -desc;
            if (upm <= 0) return false;
            m = new FontVerticalMetrics(upm, winAsc, winDesc, asc, desc, gap);
            if (os2 >= 0 && U16(d, os2) >= 4)
                m = m with
                {
                    UseTypoMetrics = (U16(d, os2 + 62) & 0x80) != 0,
                    TypoAscender = (short)U16(d, os2 + 68), TypoDescender = (short)U16(d, os2 + 70), TypoLineGap = (short)U16(d, os2 + 72),
                };
            return true;
        }
        catch (IndexOutOfRangeException) { return false; }
    }

    static int U16(byte[] d, int o) => (d[o] << 8) | d[o + 1];
    static uint U32(byte[] d, int o) => (uint)((d[o] << 24) | (d[o + 1] << 16) | (d[o + 2] << 8) | d[o + 3]);

    static int Table(byte[] d, int start, string tag)
    {
        int count = U16(d, start + 4);
        for (int i = 0; i < count; i++)
        {
            int rec = start + 12 + 16 * i;
            if (d[rec] == tag[0] && d[rec + 1] == tag[1] && d[rec + 2] == tag[2] && d[rec + 3] == tag[3]) return (int)U32(d, rec + 8);
        }
        return -1;
    }

    static string? FamilyName(byte[] d, int start)
    {
        int name = Table(d, start, "name");
        if (name < 0) return null;
        int count = U16(d, name + 2), stor = U16(d, name + 4);
        string? any = null;
        for (int j = 0; j < count; j++)
        {
            int r = name + 6 + 12 * j;
            if (U16(d, r) != 3 || U16(d, r + 6) != 1) continue;
            var s = System.Text.Encoding.BigEndianUnicode.GetString(d, name + stor + U16(d, r + 10), U16(d, r + 8));
            if (U16(d, r + 4) == 0x409) return s;
            any ??= s;
        }
        return any;
    }
}
