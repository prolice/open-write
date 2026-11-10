using System.Xml.Linq;
using OpenWrite.Core.Model;

namespace OpenWrite.Core.Editing;

/// <summary>
/// Insertion de champs (lot 5) : PAGE, NUMPAGES, SECTIONPAGES… écrits comme Word (champ complexe : début, code,
/// séparateur, résultat, fin), dans la mise en forme du caractère qui précède.
/// </summary>
public static class FieldOps
{
    /// <summary>Insère le champ <paramref name="code"/> (« PAGE ») avec un résultat provisoire ; renvoie la position après le champ.</summary>
    public static DocPos Insert(Story root, DocPos pos, string code, string result)
    {
        var p = pos.Paragraph(root);
        var w = p.Props.Ns;
        // mise en forme : celle du segment qui porte le caractère précédent (sinon la marque de paragraphe)
        FormatOps.SplitRunAt(p, pos.Offset);
        var atoms = ParagraphAtoms.Of(p);
        var before = atoms.LastOrDefault(a => a.Run is not null && a.End <= pos.Offset && a.Length > 0 && !a.InDeletion);
        var props = before?.Run?.Props.Clone() ?? new PropertyBag("rPr", w, p.Props.Get("rPr") is { } m ? new XElement(m) : null);
        Run R(RunContent c) { var r = new Run { Props = props.Clone() }; r.Content.Add(c); return r; }
        FieldChar Char(string type) { var fc = new FieldChar(w); fc.Set("fldCharType", type); return fc; }
        var runs = new List<Inline>
        {
            R(Char("begin")),
            R(new FieldCode { Value = $" {code} " }),
            R(Char("separate")),
            R(new TextContent { Value = result }),
            R(Char("end")),
        };
        // place : après le dernier élément qui finit au point d'insertion (dans sa liste), sinon en tête
        var anchor = atoms.LastOrDefault(a => a.End <= pos.Offset && a.Length > 0 && !a.InDeletion);
        if (anchor is null) p.Inlines.InsertRange(0, runs);
        else anchor.Owner.InsertRange(anchor.Owner.IndexOf(anchor.Inline) + 1, runs);
        return pos.WithOffset(pos.Offset + 3 + code.Length + 2 + result.Length);
    }
}
