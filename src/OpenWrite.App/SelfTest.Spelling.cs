using System.Windows;
using System.Windows.Automation.Peers;
using System.Windows.Automation.Provider;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using OpenWrite.Core.Model;

namespace OpenWrite.App;

// Auto-test de l'orthographe (lot 8) : un mot inventé est souligné à l'écran (pas dans le PDF), proposé par F7
// (onglet Révision ▸ Orthographe) et remplacé, n'est plus signalé en « Ne pas vérifier l'orthographe », puis
// « Ignorer tout » par le clic droit. Jamais « Ajouter au dictionnaire » ni le réglage de frappe : ils modifieraient
// le dictionnaire de Windows et les réglages de l'utilisateur.
public static partial class SelfTest
{
    /// <summary>Traits ondulés rouges (fautes) dans un dessin de page.</summary>
    static int Squiggles(Drawing? d) => d switch
    {
        DrawingGroup g => g.Children.Sum(Squiggles),
        GeometryDrawing { Pen.Brush: SolidColorBrush { Color: { R: 0xE0, G: 0x1B, B: 0x24 } } } => 1,
        _ => 0,
    };

    static void SpellingChecks(MainWindow w)
    {
        Console.WriteLine("Orthographe :");
        if (MainWindow.Spell is not { } spell || spell.SupportedLanguages.Count == 0)
        {
            Console.WriteLine("  (correcteur de Windows absent : étape non vérifiée)");
            return;
        }
        const string Fake = "xqzvbnt";
        var tag = new[] { "fr-BE", "fr-FR", Spelling.SpellService.DefaultTag, "en-US" }.FirstOrDefault(t => spell.SupportedLanguages.Contains(t, StringComparer.OrdinalIgnoreCase))
                  ?? spell.SupportedLanguages[0];
        void Undo() { ApplicationCommands.Undo.Execute(null, w); Pump(); Idle(); }
        static string TextOf(Paragraph p) => string.Concat(p.Runs().SelectMany(r => r.Content).OfType<TextContent>().Select(t => t.Value));
        Paragraph? Para() => w.Body!.Paragraphs().FirstOrDefault(p => TextOf(p).Contains(Fake));
        global::OpenWrite.Layout.ParagraphContent? Content() => Para() is { } p && w.CurrentLayout is { } l && l.LinesOf(p) is { Count: > 0 } refs ? l.Line(refs[0]).Paragraph.Content : null;
        bool Flagged() => Content() is { } c && w.SpellingErrors(c).Any(e => c.Text.Substring(e.Start, e.Length) == Fake);

        // mot inventé ajouté après « Deuxième », dans une langue dont Windows a le dictionnaire
        w.FindNext("Deuxième", true, false);
        w.ReplaceCurrent("Deuxième", "Deuxième " + Fake, true, false);
        Pump();
        w.FindNext(Fake, true, false);
        w.SetProofingLanguage(tag, noProof: false);
        Pump();
        Idle();
        Check(Flagged(), $"mot inconnu signalé ({tag})");
        int page = w.CurrentLayout!.LinesOf(Para()!)[0].Page;
        int onScreen = Squiggles(w.ScreenDrawingForTest(page)), inPdf = Squiggles(w.Output()!.Render(page + 1));
        Check(onScreen > 0 && inPdf == 0, $"trait ondulé à l'écran ({onScreen}), pas à l'impression ni dans le PDF ({inPdf})");

        // Ne pas vérifier l'orthographe (w:noProof), puis Ctrl+Z
        w.FindNext(Fake, true, false);
        w.SetProofingLanguage(tag, noProof: true);
        Pump();
        Check(!Flagged(), "« Ne pas vérifier l'orthographe » : plus signalé");
        Undo();
        Check(Flagged(), "Ctrl+Z : de nouveau signalé");

        // Révision ▸ Langue : menu affiché
        Check(SelectTab(w, "Révision") && ClickButton(w, "Langue") && w.LastMenu is { IsOpen: true } lm
              && lm.Items.OfType<MenuItem>().Any(m => HeaderOf(m) == "Définir la langue de vérification…" && Shown(m)),
              "Révision ▸ Langue : menu affiché");
        if (w.LastMenu is { IsOpen: true } open) open.IsOpen = false;
        Pump();

        // F7 par le bouton Révision ▸ Orthographe : la boîte propose le mot (après d'éventuelles autres fautes), Remplacer
        w.FindNext("Formulaire", true, false);
        Check(ClickButton(w, "Orthographe") && w.SpellingWindowForTest is not null, "Révision ▸ Orthographe : boîte ouverte");
        var sw = w.SpellingWindowForTest;
        for (int n = 0; n < 60 && sw?.CurrentWord is { } word && word != Fake; n++) sw.Next();
        Check(sw?.CurrentWord == Fake && w.SelectedText() == Fake, $"F7 : « {sw?.CurrentWord} » proposé et sélectionné dans le texte");
        if (sw is not null)
        {
            sw.Replacement = "essai";
            var replace = Descendants<Button>(sw).First(b => b.Content as string == "Remplacer");
            ((IInvokeProvider)UIElementAutomationPeer.CreatePeerForElement(replace).GetPattern(PatternInterface.Invoke)).Invoke();
            Pump();
            Idle();
            Check(w.Body!.Paragraphs().Any(p => TextOf(p).Contains("Deuxième essai")) && Para() is null, "F7 ▸ Remplacer : mot remplacé");
            sw.Close();
            Pump();
            Undo();
            Check(Para() is not null, "Ctrl+Z : mot d'origine revenu");
        }
        SelectTab(w, "Accueil");

        // clic droit sur le mot : Ignorer tout, Ajouter au dictionnaire affichés ; Ignorer tout
        w.FindNext(Fake, true, false);
        var menu = w.OpenContextMenuForTest();
        Pump();
        Idle();
        Check(menu is { IsOpen: true } && menu.Items.OfType<MenuItem>().Any(m => HeaderOf(m) == "Ajouter au dictionnaire" && Shown(m)),
              "clic droit sur une faute : suggestions, Ignorer tout, Ajouter au dictionnaire");
        Check(InvokeMenu(menu, "Ignorer tout") && !Flagged(), "clic droit ▸ Ignorer tout : plus signalé");
        Undo();   // langue posée
        Undo();   // mot ajouté
        Check(Para() is null, "Ctrl+Z : texte d'origine");
    }
}
