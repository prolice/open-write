using System.Globalization;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Xml.Linq;
using OpenWrite.App.Spelling;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Layout;

namespace OpenWrite.App;

// Lot 8 (version 1.1), orthographe : correcteur de Windows (SpellService, repris d'OpenPub), fautes soulignées d'un trait
// ondulé rouge à l'écran seulement (ni impression ni PDF), dans la langue de chaque passage (w:lang, w:noProof respecté) ;
// clic droit : suggestions, Ignorer tout, Ajouter au dictionnaire ; F7 : vérification de tout le document ;
// Révision ▸ Langue : langue de vérification de la sélection, vérification pendant la frappe. Le réglage du document
// « masquer les fautes d'orthographe » (w:hideSpellingErrors) est respecté comme dans Word.
public partial class MainWindow
{
    /// <summary>Une faute : caractères affichés du paragraphe, et langue dans laquelle le mot a été vérifié.</summary>
    internal readonly record struct SpellError(int Start, int Length, string Lang);

    sealed class SpellResult
    {
        public required List<SpellError> Errors { get; init; }
        public required List<(int, int)> Ranges { get; init; }
    }

    /// <summary>Fautes par contenu de paragraphe (un contenu est refait à chaque remise en page du paragraphe).</summary>
    ConditionalWeakTable<ParagraphContent, SpellResult> _spellCache = new();
    static readonly SpellResult NoErrors = new() { Errors = new(), Ranges = new() };

    /// <summary>Le document demande de masquer les fautes (Word : Options ▸ Vérification, propre au document).</summary>
    bool _hideSpellingErrors;

    internal static SpellService? Spell => SpellService.Shared.Value;

    /// <summary>Soulignement des fautes actif : réglage de l'utilisateur, document, correcteur de Windows présent.</summary>
    internal bool SpellingShown => AppSettings.Load().SpellCheck && !_hideSpellingErrors && Spell is not null;

    void InitSpelling()
    {
        PreviewKeyDown += (_, e) =>
        {
            if (e.Key == Key.F7 && Keyboard.Modifiers == ModifierKeys.None && _layout is not null)
            {
                ShowSpelling();
                e.Handled = true;
            }
        };
    }

    /// <summary>À l'ouverture d'un document : réglage w:hideSpellingErrors, fautes oubliées.</summary>
    void ReadSpellingSettings()
    {
        _spellCache = new();
        var settings = _pkg?.XmlPartOf("settings")?.Doc.Root;
        var hide = settings?.Elements().FirstOrDefault(e => e.Name.LocalName == "hideSpellingErrors");
        _hideSpellingErrors = hide is not null && (hide.Attributes().FirstOrDefault(a => a.Name.LocalName == "val")?.Value ?? "true") is not ("0" or "false" or "off");
    }

    /// <summary>Fournisseur des fautes pour le dessin des pages (l'impression et le PDF n'en ont pas).</summary>
    IReadOnlyList<(int Start, int Length)> MisspellingsForDrawing(ParagraphContent c) => SpellingShown ? Misspelled(c).Ranges : NoErrors.Ranges;

    SpellResult Misspelled(ParagraphContent c)
    {
        if (_spellCache.TryGetValue(c, out var r)) return r;
        r = Compute(c);
        _spellCache.AddOrUpdate(c, r);
        return r;
    }

    internal List<SpellError> SpellingErrors(ParagraphContent c) => Spell is null ? new() : Misspelled(c).Errors;

    static SpellResult Compute(ParagraphContent c)
    {
        var text = c.Text;
        if (Spell is not { } spell || c.Hidden || text.Length == 0 || c.StyleOf.Length < text.Length) return NoErrors;
        // langue de chaque caractère (null : ne pas vérifier) ; un passage par langue, les autres caractères en espaces
        var langs = new string?[text.Length];
        var distinct = new List<string>();
        for (int i = 0; i < text.Length; i++)
        {
            var f = c.Styles[c.StyleOf[i]].Format;
            // texte supprimé (suivi des modifications) ou marqué « ne pas vérifier » : pas de vérification
            var lang = f.NoProof || c.Revisions?[i] == 2 ? null : string.IsNullOrEmpty(f.Language) ? SpellService.DefaultTag : f.Language;
            langs[i] = lang;
            if (lang is not null && !distinct.Contains(lang, StringComparer.OrdinalIgnoreCase)) distinct.Add(lang);
        }
        var errors = new List<SpellError>();
        foreach (var lang in distinct)
        {
            var masked = distinct.Count == 1 && langs.All(l => l is not null)
                ? text
                : new string(text.Select((ch, i) => string.Equals(langs[i], lang, StringComparison.OrdinalIgnoreCase) ? ch : ' ').ToArray());
            foreach (var (st, len) in spell.Check(masked, lang))
                if (!Ignorable(text, st, len)) errors.Add(new SpellError(st, len, lang));
        }
        if (errors.Count == 0) return NoErrors;
        errors.Sort((a, b) => a.Start.CompareTo(b.Start));
        return new SpellResult { Errors = errors, Ranges = errors.Select(e => (e.Start, e.Length)).ToList() };
    }

    /// <summary>
    /// Mots que Word ne signale pas avec ses options par défaut : en majuscules, avec des chiffres, adresses Internet,
    /// de courriel et de fichiers.
    /// </summary>
    static bool Ignorable(string text, int start, int length)
    {
        var word = text.Substring(start, length);
        if (word.Any(char.IsDigit)) return true;
        if (word.Length > 1 && word.Any(char.IsLetter) && word.Where(char.IsLetter).All(char.IsUpper)) return true;
        int a = start, b = start + length;
        while (a > 0 && !char.IsWhiteSpace(text[a - 1])) a--;
        while (b < text.Length && !char.IsWhiteSpace(text[b])) b++;
        var token = text[a..b];
        return token.Contains('@') || token.Contains("://") || token.StartsWith("www.", StringComparison.OrdinalIgnoreCase)
               || token.Contains(":\\") || token.StartsWith("\\\\");
    }

    /// <summary>Fautes oubliées et pages redessinées (mot ignoré ou ajouté, langue changée, réglage modifié).</summary>
    internal void RefreshSpelling()
    {
        _spellCache = new();
        if (_drawings is null) return;
        _drawings.Clear();
        _pagesView?.InvalidateVisual();
        _thumbsView?.InvalidateVisual();
    }

    /// <summary>Faute sous le point d'insertion (ou exactement sélectionnée) : paragraphe, contenu et faute.</summary>
    (Paragraph P, ParagraphContent C, SpellError E)? ErrorAtCaret()
    {
        if (!SpellingShown || _caret is not { } c || LayoutOf(c.Paragraph) is not { } pl) return null;
        int from = c.Index, to = c.Index;
        if (HasSelection)
        {
            if (Ordered() is not { } sel || !ReferenceEquals(sel.From.Paragraph, sel.To.Paragraph)) return null;
            (from, to) = (sel.From.Index, sel.To.Index);
        }
        foreach (var e in SpellingErrors(pl.Content))
            if (e.Start <= from && to <= e.Start + e.Length) return (c.Paragraph, pl.Content, e);
        return null;
    }

    /// <summary>Clic droit sur une faute : suggestions (un clic remplace), Ignorer tout, Ajouter au dictionnaire.</summary>
    void AddSpellingItems(ContextMenu menu)
    {
        if (ErrorAtCaret() is not { } at || Spell is not { } spell) return;
        var word = at.C.Text.Substring(at.E.Start, at.E.Length);
        var suggestions = spell.Suggest(word, at.E.Lang);
        MenuItem Item(string header, Action act, bool bold = false)
        {
            var mi = new MenuItem { Header = header, FontWeight = bold ? FontWeights.SemiBold : FontWeights.Normal };
            mi.Click += (_, _) => { act(); _pagesView?.Focus(); };
            menu.Items.Add(mi);
            return mi;
        }
        foreach (var s in suggestions.Take(6)) Item(s, () => ReplaceWord(at.P, at.C, at.E.Start, at.E.Length, s), bold: true);
        if (suggestions.Count == 0) menu.Items.Add(new MenuItem { Header = "(aucune suggestion)", IsEnabled = false });
        Item("Ignorer tout", () => { spell.Ignore(word); RefreshSpelling(); StatusText.Text = $"« {word} » ne sera plus signalé jusqu'à la fermeture d'OpenWrite."; });
        Item("Ajouter au dictionnaire", () => { spell.Add(word, at.E.Lang); RefreshSpelling(); StatusText.Text = $"« {word} » ajouté au dictionnaire personnel de Windows."; });
        Item("Orthographe…", ShowSpelling).InputGestureText = "F7";
        menu.Items.Add(new Separator());
    }

    /// <summary>Remplace un mot (correction d'orthographe) : le nouveau texte prend la mise en forme du mot, une étape annulable.</summary>
    internal void ReplaceWord(Paragraph p, ParagraphContent c, int start, int length, string replacement)
    {
        if (EditRoot is not { } root || length <= 0 || start + length > c.ModelEnd.Length) return;
        DocPos at;
        try { at = DocPos.Of(root, p); }
        catch (ArgumentException) { NotEditable(); return; }
        var from = at.WithOffset(c.ModelStart[start]);
        var to = at.WithOffset(c.ModelEnd[start + length - 1]);
        var pos = TextOps.Edit(_journal, root, "Corriger l'orthographe", from.ContainerPath, from.BlockIndex, 1, to,
            () => TextOps.ReplaceText(root, from, to, replacement), x => x);
        AfterTyping(pos);
    }

    /// <summary>Remplace la faute proposée par F7 si le texte n'a pas changé depuis (sinon, rien).</summary>
    internal void ReplaceSpellingError(Paragraph p, SpellError e, string word, string replacement)
    {
        if (LayoutOf(p) is not { } pl || e.Start + e.Length > pl.Content.Text.Length || pl.Content.Text.Substring(e.Start, e.Length) != word) return;
        ReplaceWord(p, pl.Content, e.Start, e.Length, replacement);
        // la suite de la vérification part après le mot remplacé
        _anchor = null;
    }

    /// <summary>Dessin à l'écran de la page <paramref name="i"/> (auto-test : traits ondulés présents).</summary>
    internal System.Windows.Media.Drawing? ScreenDrawingForTest(int i) => _drawings?.Get(i);

    // ================================================================ F7 : tout le document

    SpellingWindow? _spellingWindow;
    /// <summary>Point de départ de la vérification (F7) et passage au début du document déjà fait.</summary>
    TextPoint? _spellStart;
    bool _spellWrapped;

    internal SpellingWindow? SpellingWindowForTest => _spellingWindow;

    internal void ShowSpelling()
    {
        if (_layout is null) return;
        if (Spell is null)
        {
            MessageBox.Show(this, "Le correcteur orthographique de Windows n'est pas disponible sur ce PC.\n\nParamètres ▸ Heure et langue ▸ Langue : ajoutez la langue voulue avec ses outils de vérification.",
                "Orthographe", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        if (_spellingWindow is null || !_spellingWindow.IsLoaded)
        {
            _spellingWindow = new SpellingWindow(this) { Owner = this };
            _spellingWindow.Closed += (_, _) => { _spellingWindow = null; _pagesView?.Focus(); };
            _spellingWindow.Show();
        }
        _spellStart = HasSelection && Ordered() is { } sel ? sel.From : _caret;
        _spellWrapped = false;
        _spellingWindow.Activate();
        _spellingWindow.Next();
    }

    /// <summary>
    /// Faute suivante après la sélection, jusqu'à la fin du document puis du début jusqu'au point de départ ; elle est
    /// sélectionnée (la vue défile jusqu'à elle). Null : vérification terminée.
    /// </summary>
    internal (string Word, string Lang, Paragraph P, ParagraphContent C, SpellError E)? NextSpellingError()
    {
        if (Spell is null) return null;
        var paras = SearchableParagraphs().ToList();
        if (paras.Count == 0) return null;
        var here = HasSelection && Ordered() is { } sel ? sel.To : _caret;
        int K(TextPoint? t) => t is { } x ? paras.FindIndex(q => ReferenceEquals(q.P, x.Paragraph)) : -1;
        var cur = K(here) is >= 0 and var hk ? (K: hk, I: here!.Value.Index) : (K: 0, I: 0);
        var start = K(_spellStart) is >= 0 and var sk ? (K: sk, I: _spellStart!.Value.Index) : (K: 0, I: 0);
        static int Cmp((int K, int I) a, (int K, int I) b) => a.K != b.K ? a.K.CompareTo(b.K) : a.I.CompareTo(b.I);
        while (true)
        {
            for (int k = cur.K; k < paras.Count; k++)
            {
                var (p, c) = paras[k];
                foreach (var e in SpellingErrors(c))
                {
                    if (k == cur.K && e.Start < cur.I) continue;
                    // après être reparti du début : on s'arrête au point de départ
                    if (_spellWrapped && Cmp((k, e.Start), start) >= 0) return null;
                    _anchor = new TextPoint(p, e.Start);
                    _caret = new TextPoint(p, e.Start + e.Length);
                    _pendingFormat.Clear();
                    UpdateCaret();
                    return (c.Text.Substring(e.Start, e.Length), e.Lang, p, c, e);
                }
                if (_spellWrapped && k >= start.K) return null;
            }
            if (_spellWrapped || Cmp(start, (0, 0)) == 0) return null;
            // fin du document : on reprend au début, jusqu'au point de départ
            _spellWrapped = true;
            cur = (0, 0);
        }
    }

    // ================================================================ Révision ▸ Langue

    /// <summary>Clic sur la langue dans la barre d'état : boîte Langue (comme Word).</summary>
    void LangStatus_Click(object sender, MouseButtonEventArgs e) => ShowLanguageDialog();

    internal void ShowLanguageMenu(FrameworkElement anchor)
    {
        var menu = new ContextMenu { PlacementTarget = anchor, Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom };
        LastMenu = menu;
        var set = new MenuItem { Header = "Définir la langue de vérification…" };
        set.Click += (_, _) => ShowLanguageDialog();
        menu.Items.Add(set);
        var live = new MenuItem { Header = "Vérifier l'orthographe pendant la frappe", IsCheckable = true, IsChecked = AppSettings.Load().SpellCheck };
        live.Click += (_, _) => SetSpellingShown(live.IsChecked);
        menu.Items.Add(live);
        menu.IsOpen = true;
    }

    internal void SetSpellingShown(bool on)
    {
        AppSettings.Update(x => x.SpellCheck = on);
        RefreshSpelling();
        StatusText.Text = !on ? "Les fautes d'orthographe ne sont plus soulignées (F7 vérifie toujours le document)."
            : _hideSpellingErrors ? "Ce document demande de masquer les fautes d'orthographe (réglage enregistré par Word) : F7 pour les vérifier."
            : Spell is null ? "Le correcteur orthographique de Windows n'est pas disponible sur ce PC." : "Les fautes d'orthographe sont soulignées en rouge.";
    }

    /// <summary>Langues proposées : celles des dictionnaires de Windows, puis les langues courantes de Belgique et d'Europe.</summary>
    internal static List<(string Tag, string Name, bool Checked)> LanguageChoices()
    {
        var supported = Spell?.SupportedLanguages ?? Array.Empty<string>();
        var tags = supported.Concat(new[] { "fr-BE", "fr-FR", "nl-BE", "nl-NL", "en-GB", "en-US", "de-DE", "de-BE", "es-ES", "it-IT", "pt-PT" })
            // variantes régionales à code numérique (fr-015, en-029) : sans intérêt dans la liste
            .Where(t => !t.Split('-').Skip(1).Any(part => part.All(char.IsDigit)))
            .Distinct(StringComparer.OrdinalIgnoreCase);
        var res = new List<(string, string, bool)>();
        foreach (var tag in tags)
        {
            string name;
            try { var n = CultureInfo.GetCultureInfo(tag).NativeName; name = n.Length > 0 ? char.ToUpper(n[0]) + n[1..] : tag; }
            catch (CultureNotFoundException) { name = tag; }
            res.Add((tag, name, Spell?.Resolve(tag) is not null));
        }
        return res.OrderBy(x => x.Item2, StringComparer.CurrentCultureIgnoreCase).ToList();
    }

    internal void ShowLanguageDialog()
    {
        if (_layout is null) return;
        var current = CaretFormat();
        var dlg = new LanguageWindow(LanguageChoices(), current?.Language ?? SpellService.DefaultTag, current?.NoProof == true) { Owner = this };
        if (dlg.ShowDialog() != true || dlg.ChosenTag is not string tag) { _pagesView?.Focus(); return; }
        SetProofingLanguage(tag, dlg.NoProof);
        _pagesView?.Focus();
    }

    /// <summary>Langue de vérification (w:lang w:val, les autres écritures gardées) et w:noProof de la sélection ou du mot.</summary>
    internal void SetProofingLanguage(string tag, bool noProof)
    {
        void Apply(Paragraph p, Run r)
        {
            var ns = r.Props.Ns;
            if (r.Props.Get("lang") is { } lang) lang.SetAttributeValue(ns + "val", tag);
            else r.Props.Set(new XElement(ns + "lang", new XAttribute(ns + "val", tag)));
            if (noProof) r.Props.Set(new XElement(ns + "noProof"));
            else r.Props.Remove("noProof");
        }
        var name = LanguageChoices().FirstOrDefault(x => x.Tag.Equals(tag, StringComparison.OrdinalIgnoreCase)).Name ?? tag;
        if (CharacterTarget() is not { } target)
        {
            _pendingFormat.Add(Apply);
            StatusText.Text = $"Langue « {name} » pour le texte que vous allez taper.";
            return;
        }
        ApplyCharacter("Langue", target.From, target.To, Apply);
        RefreshSpelling();
        StatusText.Text = noProof ? "Orthographe non vérifiée pour ce texte." : $"Langue de vérification : {name}.";
    }
}
