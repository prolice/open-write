using System.Globalization;
using System.Xml.Linq;

namespace OpenWrite.Core.Styles;

/// <summary>
/// Régions d'un tableau auxquelles s'appliquent les mises en forme conditionnelles d'un style de tableau
/// (w:tblStylePr) : ligne d'en-tête, ligne des totaux, première et dernière colonnes, bandes, cellules des coins.
/// </summary>
[Flags]
public enum TableRegion
{
    None = 0,
    FirstRow = 1, LastRow = 2, FirstCol = 4, LastCol = 8,
    Band1Vert = 16, Band2Vert = 32, Band1Horz = 64, Band2Horz = 128,
    NwCell = 256, NeCell = 512, SwCell = 1024, SeCell = 2048,
}

/// <summary>Style du tableau qui contient un paragraphe et régions de sa cellule (mise en forme conditionnelle).</summary>
public readonly record struct CellStyle(string? TableStyleId, TableRegion Regions = TableRegion.None)
{
    public static implicit operator CellStyle(string? tableStyleId) => new(tableStyleId);
}

/// <summary>Options de style du tableau (w:tblLook) : quelles régions conditionnelles sont actives.</summary>
public readonly record struct TableLook(bool FirstRow, bool LastRow, bool FirstColumn, bool LastColumn, bool NoHBand, bool NoVBand)
{
    /// <summary>Options de Word pour un nouveau tableau (w:val="04A0") : en-tête, première colonne, lignes à bandes.</summary>
    public static TableLook Default => new(true, false, true, false, false, true);
}

public static class TableStyles
{
    /// <summary>
    /// Ordre d'application des mises en forme conditionnelles (la suivante l'emporte) : bandes verticales, bandes
    /// horizontales, première / dernière colonne, première / dernière ligne, cellules des coins (ECMA-376 §17.7.6).
    /// </summary>
    public static readonly (TableRegion Region, string Type)[] Order =
    {
        (TableRegion.Band1Vert, "band1Vert"), (TableRegion.Band2Vert, "band2Vert"),
        (TableRegion.Band1Horz, "band1Horz"), (TableRegion.Band2Horz, "band2Horz"),
        (TableRegion.FirstCol, "firstCol"), (TableRegion.LastCol, "lastCol"),
        (TableRegion.FirstRow, "firstRow"), (TableRegion.LastRow, "lastRow"),
        (TableRegion.NeCell, "neCell"), (TableRegion.NwCell, "nwCell"),
        (TableRegion.SeCell, "seCell"), (TableRegion.SwCell, "swCell"),
    };

    /// <summary>
    /// w:tblLook : attributs explicites (Word 2010+) sinon bits de w:val (0x20 première ligne, 0x40 dernière ligne,
    /// 0x80 première colonne, 0x100 dernière colonne, 0x200 sans bandes horizontales, 0x400 sans bandes verticales).
    /// Absent : aucune région (bandes comprises : w:noHBand et w:noVBand valent 0).
    /// </summary>
    public static TableLook Look(XElement? tblLook)
    {
        if (tblLook is null) return new(false, false, false, false, false, false);
        var W = tblLook.Name.Namespace;
        int val = int.TryParse(tblLook.Attribute(W + "val")?.Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var v) ? v : 0;
        bool Get(string attr, int bit) => tblLook.Attribute(W + attr)?.Value is { } a ? a is "1" or "true" or "on" : (val & bit) != 0;
        return new(Get("firstRow", 0x20), Get("lastRow", 0x40), Get("firstColumn", 0x80), Get("lastColumn", 0x100), Get("noHBand", 0x200), Get("noVBand", 0x400));
    }

    /// <summary>
    /// Régions d'une cellule (rangée et colonne comptées depuis 0). Les bandes ne comptent pas la ligne d'en-tête ni la
    /// première colonne quand elles sont actives (Word : cnfStyle « oddHBand » sur la rangée qui suit l'en-tête), et
    /// s'arrêtent avant la ligne des totaux et la dernière colonne actives ; <paramref name="rowBand"/> /
    /// <paramref name="colBand"/> = w:tblStyleRowBandSize / w:tblStyleColBandSize.
    /// </summary>
    public static TableRegion Regions(TableLook look, int row, int rowCount, int col, int colCount, int rowBand = 1, int colBand = 1)
    {
        var r = TableRegion.None;
        bool firstRow = look.FirstRow && row == 0;
        bool lastRow = look.LastRow && row == rowCount - 1 && !firstRow;
        bool firstCol = look.FirstColumn && col == 0;
        bool lastCol = look.LastColumn && col == colCount - 1 && !firstCol;
        if (firstRow) r |= TableRegion.FirstRow;
        if (lastRow) r |= TableRegion.LastRow;
        if (firstCol) r |= TableRegion.FirstCol;
        if (lastCol) r |= TableRegion.LastCol;
        if (!look.NoHBand && !firstRow && !lastRow)
        {
            int k = row - (look.FirstRow ? 1 : 0);
            r |= k / Math.Max(1, rowBand) % 2 == 0 ? TableRegion.Band1Horz : TableRegion.Band2Horz;
        }
        if (!look.NoVBand && !firstCol && !lastCol)
        {
            int k = col - (look.FirstColumn ? 1 : 0);
            r |= k / Math.Max(1, colBand) % 2 == 0 ? TableRegion.Band1Vert : TableRegion.Band2Vert;
        }
        if (firstRow && firstCol) r |= TableRegion.NwCell;
        if (firstRow && lastCol) r |= TableRegion.NeCell;
        if (lastRow && firstCol) r |= TableRegion.SwCell;
        if (lastRow && lastCol) r |= TableRegion.SeCell;
        return r;
    }

    /// <summary>
    /// Valeur de w:cnfStyle/@w:val (12 chiffres) pour des régions : cache écrit par Word. Ordre relevé dans les fichiers
    /// de Word (attributs nommés à côté) : les cellules des coins viennent dans l'ordre ne, nw, se, sw.
    /// </summary>
    public static string CnfValue(TableRegion r)
    {
        var order = new[]
        {
            TableRegion.FirstRow, TableRegion.LastRow, TableRegion.FirstCol, TableRegion.LastCol,
            TableRegion.Band1Vert, TableRegion.Band2Vert, TableRegion.Band1Horz, TableRegion.Band2Horz,
            TableRegion.NeCell, TableRegion.NwCell, TableRegion.SeCell, TableRegion.SwCell,
        };
        return string.Concat(order.Select(f => (r & f) != 0 ? '1' : '0'));
    }

    /// <summary>Régions décrites par une valeur w:cnfStyle/@w:val (inverse de <see cref="CnfValue"/>).</summary>
    public static TableRegion FromCnf(string? val)
    {
        if (val is null) return TableRegion.None;
        var order = new[]
        {
            TableRegion.FirstRow, TableRegion.LastRow, TableRegion.FirstCol, TableRegion.LastCol,
            TableRegion.Band1Vert, TableRegion.Band2Vert, TableRegion.Band1Horz, TableRegion.Band2Horz,
            TableRegion.NeCell, TableRegion.NwCell, TableRegion.SeCell, TableRegion.SwCell,
        };
        var r = TableRegion.None;
        for (int k = 0; k < Math.Min(val.Length, order.Length); k++) if (val[k] == '1') r |= order[k];
        return r;
    }
}
