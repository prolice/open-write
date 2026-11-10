using System.Windows;
using System.Xml.Linq;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;

namespace OpenWrite.App;

// Lot 5 (version 0.6) : onglet Mise en page — marges, orientation, taille, colonnes, sauts de page, de colonne et de
// section ; boîte « Mise en page ». Les w:sectPr modifiés sont suivis par le journal (annulable) ; la mise en page
// change partout : le document entier est remis en page.
public partial class MainWindow
{
    const double PtPerCm = 72 / 2.54;

    /// <summary>w:sectPr de la section du point d'insertion (paragraphe de premier niveau qui contient le point d'insertion).</summary>
    XElement? CaretSection()
    {
        if (Body is not { } body || _caret is not { } c || ToDocPos(c) is not { } pos) return null;
        return SectionOps.SectPrFor(body, pos.Path[0], body.Paragraphs().FirstOrDefault()?.Props.Ns ?? Core.Model.PropertyBag.TransitionalNs);
    }

    /// <summary>Modifie la section du point d'insertion (ou toutes), en une opération annulable.</summary>
    internal void EditSections(string label, Action<XElement> change, bool all = false)
    {
        if (Body is not { } body || CaretSection() is not { } current) return;
        var w = current.Name.Namespace;
        var targets = all ? SectionOps.All(body, w) : new List<XElement> { current };
        var caret = _caret is { } c ? ToDocPos(c) : null;
        var (anchor, caretPos) = SelectionDocPos();
        using (var tx = _journal.Begin(label, body, Array.Empty<int>(), 0, 0, caret))
        {
            foreach (var t in targets) tx.TrackXml(t);
            foreach (var t in targets) change(t);
            tx.Commit(caret);
        }
        AfterFormat(label, anchor, caretPos);
    }

    static readonly (string Name, double Top, double Bottom, double Left, double Right)[] MarginPresets =
    {
        ("Normales", 2.5, 2.5, 2.5, 2.5), ("Étroites", 1.27, 1.27, 1.27, 1.27), ("Modérées", 2.54, 2.54, 1.91, 1.91), ("Larges", 2.54, 2.54, 5.08, 5.08),
    };

    internal void ShowMarginsMenu(FrameworkElement anchor) => ShowMenu(anchor, MarginPresets.Select(m =>
        ($"{m.Name} — haut {m.Top.ToString(Cu)}, bas {m.Bottom.ToString(Cu)}, gauche {m.Left.ToString(Cu)}, droite {m.Right.ToString(Cu)} cm", (UIElement?)null,
         (Action)(() => EditSections($"Marges {m.Name.ToLowerInvariant()}", s => SectionOps.SetMargins(s, m.Top * PtPerCm, m.Bottom * PtPerCm, m.Left * PtPerCm, m.Right * PtPerCm)))))
        .Append(("Marges personnalisées…", null, (Action)(() => ShowPageSetup()))));

    internal void ShowOrientationMenu(FrameworkElement anchor) => ShowMenu(anchor, new (string, UIElement?, Action)[]
    {
        ("Portrait", null, () => EditSections("Orientation portrait", s => SectionOps.SetOrientation(s, false))),
        ("Paysage", null, () => EditSections("Orientation paysage", s => SectionOps.SetOrientation(s, true))),
    });

    internal void ShowSizeMenu(FrameworkElement anchor) => ShowMenu(anchor, PageSetupWindow.Papers.Select(p =>
        ($"{p.Name} — {p.Wcm.ToString(Cu)} × {p.Hcm.ToString(Cu)} cm", (UIElement?)null,
         (Action)(() => EditSections($"Format {p.Name}", s => SectionOps.SetPaper(s, p.Wcm * PtPerCm, p.Hcm * PtPerCm)))))
        .Append(("Autres formats…", null, (Action)(() => ShowPageSetup()))));

    internal void ShowColumnsMenu(FrameworkElement anchor) => ShowMenu(anchor, new (string, UIElement?, Action)[]
    {
        ("Une", null, () => EditSections("Une colonne", s => SectionOps.SetColumns(s, 1))),
        ("Deux", null, () => EditSections("Deux colonnes", s => SectionOps.SetColumns(s, 2))),
        ("Trois", null, () => EditSections("Trois colonnes", s => SectionOps.SetColumns(s, 3))),
        ("Gauche (étroite à gauche)", null, () => EditSections("Colonnes : gauche", s => SectionOps.SetColumns(s, 2, uneven: "left"))),
        ("Droite (étroite à droite)", null, () => EditSections("Colonnes : droite", s => SectionOps.SetColumns(s, 2, uneven: "right"))),
        ("Autres colonnes…", null, () => ShowPageSetup(columnsOnly: true)),
    });

    internal void ShowBreaksMenu(FrameworkElement anchor) => ShowMenu(anchor, new (string, UIElement?, Action)[]
    {
        ("Saut de page (Ctrl+Entrée)", null, () => InsertSpecial(SpecialChar.PageBreak)),
        ("Saut de colonne (Ctrl+Maj+Entrée)", null, () => InsertSpecial(SpecialChar.ColumnBreak)),
        ("Saut de section : page suivante", null, () => InsertSectionBreak("nextPage")),
        ("Saut de section : continu", null, () => InsertSectionBreak("continuous")),
        ("Saut de section : page paire", null, () => InsertSectionBreak("evenPage")),
        ("Saut de section : page impaire", null, () => InsertSectionBreak("oddPage")),
    });

    /// <summary>Saut de section au point d'insertion (corps du document).</summary>
    internal void InsertSectionBreak(string type)
    {
        if (Body is not { } body || _caret is not { } c || ToDocPos(c) is not { } pos) return;
        if (pos.Path.Count != 1) { StatusText.Text = "Un saut de section ne se place pas dans un tableau."; return; }
        var w = body.Paragraphs().First().Props.Ns;
        var after = SectionOps.SectPrFor(body, pos.BlockIndex, w);
        DocPos result;
        using (var tx = _journal.Begin("Saut de section", body, Array.Empty<int>(), pos.BlockIndex, 1, pos))
        {
            // le w:sectPr de la section (paragraphe suivant ou corps) change de type : suivi par le journal
            tx.TrackXml(after);
            result = SectionOps.InsertBreak(body, pos, type, NewParaId());
            tx.Commit(result);
        }
        AfterTyping(result);
        StatusText.Text = type switch
        {
            "continuous" => "Saut de section continu inséré. Ctrl+Z pour annuler.",
            "evenPage" => "Saut de section (page paire) inséré.",
            "oddPage" => "Saut de section (page impaire) inséré.",
            _ => "Saut de section (page suivante) inséré.",
        };
    }

    /// <summary>Boîte « Mise en page » (ou « Colonnes »), appliquée à la section ou au document.</summary>
    internal void ShowPageSetup(bool columnsOnly = false)
    {
        if (CaretSection() is not { } sect) return;
        var w = sect.Name.Namespace;
        var (pw, ph, landscape) = SectionOps.PageSize(sect);
        var cols = sect.Element(w + "cols");
        int n = int.TryParse(cols?.Attribute(w + "num")?.Value, out var cn) ? cn : 1;
        double space = double.TryParse(cols?.Attribute(w + "space")?.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var sp) ? sp / 20 : 36;
        bool sep = cols?.Attribute(w + "sep")?.Value is "1" or "true" or "on";
        var dlg = new PageSetupWindow(SectionOps.Margins(sect), pw, ph, landscape, n, space, sep, columnsOnly) { Owner = this };
        if (dlg.ShowDialog() != true) { BackToPage(); return; }
        EditSections(columnsOnly ? "Colonnes" : "Mise en page", s =>
        {
            if (!columnsOnly)
            {
                var m = dlg.Margins;
                SectionOps.SetMargins(s, m.Top, m.Bottom, m.Left, m.Right, m.Header, m.Footer);
                SectionOps.SetPaper(s, dlg.Paper.W, dlg.Paper.H);
                SectionOps.SetOrientation(s, dlg.Landscape);
            }
            SectionOps.SetColumns(s, dlg.Columns, dlg.ColumnSpace, separator: dlg.Separator);
        }, all: dlg.AllSections);
        BackToPage();
    }
}
