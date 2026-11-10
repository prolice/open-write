using System.Windows.Input;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;
using OpenWrite.Layout;

namespace OpenWrite.App;

// Lot 3 (version 0.4) : mise en forme directe — gras, italique, souligné, alignement — au clavier (raccourcis de Word
// en français et en anglais) et par le ruban. Comme Word : sur la sélection ; sans sélection, sur le mot qui contient
// le point d'insertion ; entre deux mots, la mise en forme s'applique à la prochaine frappe.
public partial class MainWindow
{
    /// <summary>Mise en forme en attente au point d'insertion (Ctrl+G entre deux mots) : appliquée au texte tapé ensuite.</summary>
    readonly List<Action<Paragraph, Run>> _pendingFormat = new();

    /// <summary>Raccourcis de mise en forme ; vrai si la touche a été traitée.</summary>
    bool FormattingKey(Key key, bool ctrl, bool shift)
    {
        if (!ctrl) return false;
        switch (key)
        {
            case Key.D when !shift: ShowFontDialog(); return true;
            case Key.F when !shift: ShowFind(replace: false); return true;
            case Key.K when !shift: ShowLinkDialog(); return true;
            case Key.H when !shift: ShowFind(replace: true); return true;
            case Key.D8 when shift:
            case Key.Multiply: ToggleMarks(); return true;
            case Key.Space when !shift: ClearCharacterFormatting(); return true;
            case Key.Q when !shift: ClearParagraphFormatting(); return true;
            // Ctrl+Maj+> et Ctrl+< (touche &lt;&gt; des claviers européens), Ctrl+Maj+. et Ctrl+Maj+, (claviers américains), Ctrl+] et Ctrl+[
            case Key.OemBackslash: GrowFont(shift ? 1 : -1); return true;
            case Key.OemPeriod when shift:
            case Key.Oem6 when !shift: GrowFont(1); return true;
            case Key.OemComma when shift:
            case Key.Oem4 when !shift: GrowFont(-1); return true;
            case Key.G when !shift:
            case Key.B when !shift: ToggleCharacter("b"); return true;
            case Key.I when !shift: ToggleCharacter("i"); return true;
            case Key.U when !shift: ToggleCharacter("u"); return true;
            case Key.E when !shift: SetAlignment("center"); return true;
            case Key.J when !shift: SetAlignment("both"); return true;
            case Key.L when !shift:
            case Key.G when shift: SetAlignment("left"); return true;
            case Key.R when !shift:
            case Key.D when shift: SetAlignment("right"); return true;
        }
        return false;
    }

    static bool Has(RunFormat f, string what) => what switch
    {
        "b" => f.Bold,
        "i" => f.Italic,
        "u" => f.Underline is not null and not "none",
        "strike" => f.Strike,
        _ => false,
    };

    static string LabelOf(string what) => what switch { "b" => "Gras", "i" => "Italique", "strike" => "Barré", _ => "Souligné" };

    /// <summary>Texte visé par une mise en forme de caractère : la sélection, sinon le mot du point d'insertion.</summary>
    (TextPoint From, TextPoint To)? CharacterTarget()
    {
        if (HasSelection) return Ordered();
        if (_caret is not { } c || LayoutOf(c.Paragraph) is not { } pl) return null;
        var t = pl.Content.Text;
        int i = c.Index;
        if (i <= 0 || i >= t.Length || !CaretGeometry.IsWordChar(t[i - 1]) || !CaretGeometry.IsWordChar(t[i])) return null;
        var (s, e) = CaretGeometry.WordAt(t, i);
        while (e > s && t[e - 1] == ' ') e--;
        return (new TextPoint(c.Paragraph, s), new TextPoint(c.Paragraph, e));
    }

    /// <summary>Mise en forme résolue des caractères d'une plage (au plus 20 000), dans l'ordre.</summary>
    IEnumerable<RunFormat> FormatsIn(TextPoint from, TextPoint to)
    {
        if (_layout is null || Locate(from) is not { } a || Locate(to) is not { } b) yield break;
        int n = 0;
        for (LineRef? r = a.Ref; r is { } cur && cur.CompareTo(b.Ref) <= 0 && n < 20000; r = Src.Step(cur, 1))
        {
            var l = Src.Line(cur);
            var box = l.Paragraph.Lines[l.Line];
            var c = l.Paragraph.Content;
            int s = cur.Equals(a.Ref) ? Math.Max(box.Start, from.Index) : box.Start;
            int e = cur.Equals(b.Ref) ? Math.Min(box.End, to.Index) : box.End;
            for (int k = s; k < e && k < c.Text.Length; k++, n++)
                if (!char.IsWhiteSpace(c.Text[k])) yield return c.Styles[c.StyleOf[k]].Format;
        }
    }

    /// <summary>Format du caractère qui précède le point d'insertion (celui que prendra la frappe).</summary>
    RunFormat? CaretFormat()
    {
        if (_caret is not { } c || LayoutOf(c.Paragraph) is not { } pl) return null;
        var ct = pl.Content;
        int i = Math.Min(c.Index, ct.Text.Length);
        return i > 0 ? ct.Styles[ct.StyleOf[i - 1]].Format : i < ct.Text.Length ? ct.Styles[ct.StyleOf[i]].Format : ct.Mark.Format;
    }

    /// <summary>Gras, italique ou souligné : retiré si tout le texte visé l'a déjà, appliqué sinon.</summary>
    internal void ToggleCharacter(string what)
    {
        if (_styles is not { } styles) return;
        if (CharacterTarget() is not { } target)
        {
            // entre deux mots : pour la frappe qui suit
            bool current = CaretFormat() is { } cf && Has(cf, what);
            bool want = !current;
            _pendingFormat.Add((p, r) => Set(r, p, want));
            StatusText.Text = $"{LabelOf(what)} {(want ? "activé" : "désactivé")} pour le texte que vous allez taper.";
            RefreshFormatPane();
            return;
        }
        var formats = FormatsIn(target.From, target.To).ToList();
        bool on = formats.Count == 0 || !formats.All(f => Has(f, what));
        ApplyCharacter(LabelOf(what), target.From, target.To, (p, r) => Set(r, p, on));

        void Set(Run r, Paragraph p, bool value)
        {
            Func<Run, bool> inherited = run => Has(styles.Run(p, run), what);
            if (what == "u") FormatOps.SetUnderline(r, value, inherited);
            else FormatOps.SetToggle(r, what, value, inherited);
        }
    }

    /// <summary>Applique une mise en forme de caractère à une plage (annulable), sélection gardée.</summary>
    void ApplyCharacter(string label, TextPoint a, TextPoint b, Action<Paragraph, Run> apply)
    {
        if (EditRoot is not { } body || ToDocPos(a) is not { } from || ToDocPos(b) is not { } to) { NotEditable(); return; }
        if (!SupportedRange(from, to)) return;
        var (anchor, caret) = SelectionDocPos();
        FormatOps.Apply(_journal, body, label, from, to, () => FormatOps.ApplyToRuns(body, from, to, apply));
        AfterFormat(label, anchor, caret);
    }

    /// <summary>Alignement des paragraphes de la sélection (ou du paragraphe du point d'insertion) ; déjà aligné ainsi : à gauche.</summary>
    internal void SetAlignment(string jc)
    {
        if (EditRoot is not { } body || Ordered() is not { } sel) return;
        if (ToDocPos(sel.From) is not { } from || ToDocPos(sel.To) is not { } to) { NotEditable(); return; }
        if (!SupportedRange(from, to)) return;
        var paragraphs = new List<Paragraph>();
        FormatOps.ApplyToParagraphs(body, from, to, paragraphs.Add);
        bool already = paragraphs.All(p => _styles?.Paragraph(p).Alignment == jc);
        var value = already && jc != "left" ? "left" : jc;
        var label = value switch { "center" => "Centrer", "right" => "Aligner à droite", "both" => "Justifier", _ => "Aligner à gauche" };
        var (anchor, caret) = SelectionDocPos();
        FormatOps.Apply(_journal, body, label, from, to, () => FormatOps.ApplyToParagraphs(body, from, to, p => p.Props.SetVal("jc", value)));
        AfterFormat(label, anchor, caret);
    }

    (DocPos? Anchor, DocPos? Caret) SelectionDocPos() =>
        (_anchor is { } a ? ToDocPos(a) : null, _caret is { } c ? ToDocPos(c) : null);

    /// <summary>Après une mise en forme : remise en page, même sélection (les positions ne changent pas).</summary>
    void AfterFormat(string label, DocPos? anchor, DocPos? caret)
    {
        Relayout();
        _anchor = anchor is not null ? FromDocPos(anchor) : null;
        if (caret is not null && FromDocPos(caret) is { } c) _caret = c;
        Scroller.UpdateLayout();
        UpdateCaret();
        UpdateModifiedMarker();
        StatusText.Text = $"{label}. Ctrl+Z pour annuler.";
        EnsureAutoSaveTimer();
        CommandManager.InvalidateRequerySuggested();
    }

    /// <summary>Frappe avec une mise en forme en attente : insertion et mise en forme en une seule opération annulable.</summary>
    DocPos TypeWithPendingFormat(Story body, DocPos at, string text)
    {
        var apply = _pendingFormat.ToList();
        _pendingFormat.Clear();
        return TextOps.Edit(_journal, body, "Frappe", at.ContainerPath, at.BlockIndex, 1, at, () =>
        {
            var end = TextOps.Insert(body, at, text);
            FormatOps.ApplyToRuns(body, at, end, (p, r) => { foreach (var f in apply) f(p, r); });
            return end;
        }, p => p);
    }
}
