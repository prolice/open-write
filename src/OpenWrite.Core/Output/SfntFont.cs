// Repris d'OpenPub 2.1.4 (OpenPub.Core/Output/SfntFont.cs), espace de noms seul changé.
using System.Text;

namespace OpenWrite.Core.Output;

/// <summary>
/// Outils minimaux sur les fichiers de police TrueType/OpenType (sfnt), pour l'incorporation dans les PDF :
/// collections .ttc (une police extraite en fichier .ttf autonome), contours CFF, noms de famille et de style.
/// </summary>
public static class SfntFont
{
    static int U16(byte[] d, int o) => (d[o] << 8) | d[o + 1];
    static uint U32(byte[] d, int o) => (uint)((d[o] << 24) | (d[o + 1] << 16) | (d[o + 2] << 8) | d[o + 3]);

    public static bool IsCollection(byte[] d) => d.Length >= 12 && d[0] == 't' && d[1] == 't' && d[2] == 'c' && d[3] == 'f';

    /// <summary>Nombre de polices du fichier (1 pour un .ttf/.otf).</summary>
    public static int FaceCount(byte[] d) => IsCollection(d) ? (int)U32(d, 8) : 1;

    static int FaceOffset(byte[] d, int index) => IsCollection(d) ? (int)U32(d, 12 + 4 * index) : 0;

    /// <summary>Tables (etiquette, decalage, longueur) de la police <paramref name="index"/> du fichier.</summary>
    public static List<(string Tag, int Offset, int Length)> Tables(byte[] d, int index = 0)
    {
        int face = FaceOffset(d, index);
        int n = U16(d, face + 4);
        var list = new List<(string, int, int)>(n);
        for (int i = 0; i < n; i++)
        {
            int rec = face + 12 + 16 * i;
            list.Add((Encoding.ASCII.GetString(d, rec, 4), (int)U32(d, rec + 8), (int)U32(d, rec + 12)));
        }
        return list;
    }

    /// <summary>Contours PostScript (table CFF) : non pris en charge par la bibliotheque PDF.</summary>
    public static bool HasCffOutlines(byte[] d, int index = 0) => Tables(d, index).Any(t => t.Tag is "CFF " or "CFF2");

    /// <summary>
    /// Licence de la police (OS/2 fsType) interdisant toute incorporation (« restricted license ») : le texte est
    /// alors dessine en contours dans le PDF.
    /// </summary>
    public static bool EmbeddingRestricted(byte[] d, int index = 0)
    {
        var os2 = Tables(d, index).FirstOrDefault(t => t.Tag == "OS/2");
        if (os2.Tag is null || os2.Length < 10) return false;
        return (U16(d, os2.Offset + 8) & 0x000F) == 0x0002;
    }

    /// <summary>Famille et style (anglais de preference, plate-forme Windows) de la police <paramref name="index"/>.</summary>
    public static (string? Family, string? Style) Names(byte[] d, int index = 0)
    {
        var name = Tables(d, index).FirstOrDefault(t => t.Tag == "name");
        if (name.Tag is null) return (null, null);
        int b = name.Offset, count = U16(d, b + 2), stor = U16(d, b + 4);
        string? fam = null, sty = null, famAny = null, styAny = null;
        for (int j = 0; j < count; j++)
        {
            int r = b + 6 + 12 * j;
            int pid = U16(d, r), lid = U16(d, r + 4), nid = U16(d, r + 6), len = U16(d, r + 8), o = U16(d, r + 10);
            if (pid != 3 || (nid != 1 && nid != 2)) continue;
            var s = Encoding.BigEndianUnicode.GetString(d, b + stor + o, len);
            if (nid == 1) { famAny ??= s; if (lid == 0x409) fam = s; }
            else { styAny ??= s; if (lid == 0x409) sty = s; }
        }
        return (fam ?? famAny, sty ?? styAny);
    }

    /// <summary>
    /// Police <paramref name="index"/> d'une collection, recopiee en fichier sfnt autonome (tables alignees sur
    /// 4 octets) ; un fichier simple est renvoye tel quel.
    /// </summary>
    public static byte[] ExtractFace(byte[] d, int index)
    {
        if (!IsCollection(d)) return d;
        int face = FaceOffset(d, index);
        var tables = Tables(d, index);
        int n = tables.Count;
        int headerLen = 12 + 16 * n;
        int total = headerLen + tables.Sum(t => (t.Length + 3) & ~3);
        var o = new byte[total];
        Array.Copy(d, face, o, 0, 12);                // version sfnt, nombre de tables, champs de recherche
        int pos = headerLen;
        for (int i = 0; i < n; i++)
        {
            int src = face + 12 + 16 * i, dst = 12 + 16 * i;
            Array.Copy(d, src, o, dst, 8);              // etiquette et somme de controle
            var (_, off, len) = tables[i];
            o[dst + 8] = (byte)(pos >> 24); o[dst + 9] = (byte)(pos >> 16); o[dst + 10] = (byte)(pos >> 8); o[dst + 11] = (byte)pos;
            o[dst + 12] = (byte)(len >> 24); o[dst + 13] = (byte)(len >> 16); o[dst + 14] = (byte)(len >> 8); o[dst + 15] = (byte)len;
            Array.Copy(d, off, o, pos, len);
            pos += (len + 3) & ~3;
        }
        return o;
    }

    /// <summary>Indice, dans une collection, de la police dont la famille et le style correspondent (0 a defaut).</summary>
    public static int FindFace(byte[] d, string? family, string? style)
    {
        int n = FaceCount(d);
        for (int i = 0; i < n; i++)
        {
            var (f, s) = Names(d, i);
            if (string.Equals(f, family, StringComparison.OrdinalIgnoreCase) && (style is null || string.Equals(s, style, StringComparison.OrdinalIgnoreCase)))
                return i;
        }
        for (int i = 0; i < n; i++)
            if (string.Equals(Names(d, i).Family, family, StringComparison.OrdinalIgnoreCase)) return i;
        return 0;
    }
}
