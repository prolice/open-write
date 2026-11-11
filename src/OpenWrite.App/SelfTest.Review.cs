using System.IO;
using System.Windows;
using System.Windows.Automation.Peers;
using System.Windows.Automation.Provider;
using System.Windows.Controls;
using System.Windows.Input;
using OpenWrite.Core.Model;
using OpenWrite.Core.Review;
using OpenWrite.Docx;

namespace OpenWrite.App;

// Auto-test de la révision (lot 8), par les vrais contrôles : Révision ▸ Suivi des modifications, frappe et suppression
// suivies, son propre ajout effacé vraiment, Révision ▸ Refuser ▸ toutes ; Révision ▸ Nouveau, volet Commentaires
// (publier, répondre, résoudre, supprimer), Ctrl+Z. Une copie du document avec une modification suivie et un fil de
// commentaires est écrite dans le dossier de l'auto-test (revision.docx) pour la faire ouvrir par Word.
// L'affichage des révisions n'est jamais changé par le menu (il est gardé dans les réglages de l'utilisateur).
public static partial class SelfTest
{
    static void Invoke(Button b)
    {
        ((IInvokeProvider)UIElementAutomationPeer.CreatePeerForElement(b).GetPattern(PatternInterface.Invoke)).Invoke();
        Pump();
        Idle();
    }

    static void ReviewChecks(MainWindow w, string outDir)
    {
        Console.WriteLine("Révision :");
        void Key(Key k)
        {
            var v = w.PagesViewForTest!;
            v.RaiseEvent(new KeyEventArgs(Keyboard.PrimaryDevice, PresentationSource.FromVisual(v)!, 0, k) { RoutedEvent = Keyboard.KeyDownEvent });
            Pump();
        }
        void Type(string text)
        {
            var v = w.PagesViewForTest!;
            foreach (var ch in text)
                v.RaiseEvent(new TextCompositionEventArgs(Keyboard.PrimaryDevice, new TextComposition(InputManager.Current, v, ch.ToString()))
                    { RoutedEvent = TextCompositionManager.TextInputEvent });
            Pump();
        }
        void Undo() { ApplicationCommands.Undo.Execute(null, w); Pump(); Idle(); }
        static string Final(Paragraph p) => string.Concat(p.Runs().SelectMany(r => r.Content).OfType<TextContent>().Where(t => !t.Deleted).Select(t => t.Value));
        Paragraph Para() => w.Body!.Paragraphs().First(p => string.Concat(p.Runs().SelectMany(r => r.Content).OfType<TextContent>().Select(t => t.Value)).Contains("Deuxième"));
        List<Revision> Revs() => Revisions.List(w.Body!);
        var original = Final(Para());

        // suivi des modifications par le bouton du ruban
        Check(SelectTab(w, "Révision") && ClickButton(w, "Suivi des modifications") && w.Tracking, "Révision ▸ Suivi des modifications : activé");
        var settings = w.SettingsForTest;
        Check(settings?.Elements().Select(e => e.Name.LocalName).ToList() is { } names && names.IndexOf("trackRevisions") is var k and >= 0 && k < names.IndexOf("defaultTabStop"),
              "settings.xml : w:trackRevisions à sa place (avant w:defaultTabStop)");
        w.FindNext("Deuxième", true, false);
        Key(System.Windows.Input.Key.Right);         // fin du mot
        Type(" ajouté");
        Check(Revs().Any(r => r.Kind == RevisionKind.Inserted && r.Author == MainWindow.UserName) && Final(Para()).Contains("Deuxième ajouté"),
              $"frappe suivie : insertion au nom de « {MainWindow.UserName} »");
        Key(System.Windows.Input.Key.Back);
        Key(System.Windows.Input.Key.Back);
        Check(Final(Para()).Contains("Deuxième ajou") && !Final(Para()).Contains("ajouté") && !Revs().Any(r => r.Kind == RevisionKind.Deleted),
              "Retour arrière dans son propre ajout : effacé vraiment (pas barré)");
        w.FindNext("Deuxième", true, false);
        Key(System.Windows.Input.Key.Delete);
        Check(Revs().Any(r => r.Kind == RevisionKind.Deleted) && !Final(Para()).Contains("Deuxième"), "Suppr sur une sélection : texte barré (suppression suivie)");
        // affichage : le menu propose les quatre vues ; toutes les marques montre le texte barré
        bool clicked = ClickButton(w, "Marques");
        Check(clicked && w.LastMenu is { IsOpen: true } vm && new[] { "Marques simples", "Toutes les marques", "Aucune marque", "Original" }
              .All(h => vm.Items.OfType<MenuItem>().Any(m => HeaderOf(m) == h && Shown(m))), "Révision ▸ Marques : quatre vues");
        if (w.LastMenu is { IsOpen: true } open) open.IsOpen = false;
        var before = w.View;
        w.SetRevisionView(global::OpenWrite.Layout.RevisionView.All, persist: false);
        var content = w.CurrentLayout!.Line(w.CurrentLayout.LinesOf(Para())[0]).Paragraph.Content;
        Check(content.Text.Contains("Deuxième") && content.Revisions is { } rv && rv.Contains((byte)2) && rv.Contains((byte)1), "toutes les marques : texte supprimé barré et inséré visibles");
        w.SetRevisionView(before, persist: false);

        // copie pour Word : modification suivie + fil de commentaires (écrite plus bas)
        // Révision ▸ Refuser ▸ toutes
        Check(ClickButton(w, "Refuser") && InvokeMenu(w.LastMenu, "Refuser toutes les modifications") && Revs().Count == 0 && Final(Para()) == original,
              $"Révision ▸ Refuser ▸ toutes : texte d'origine (« {Final(Para())} »)");
        Undo();
        Check(Revs().Count > 0, "Ctrl+Z : modifications revenues");

        // commentaires : nouveau, publier, répondre, résoudre (volet) — sur « ligne » (« Deuxième » est barré)
        w.FindNext("ligne", true, false);
        Check(ClickButton(w, "Nouveau") && w.CommentsPaneShown && w.CommentsPaneForTest?.Editor is not null, "Révision ▸ Nouveau : volet Commentaires, carte à remplir");
        if (w.CommentsPaneForTest?.Editor is { } ed)
        {
            ed.Box.Text = "Relire ce mot";
            Invoke(ed.Ok);
        }
        var comments = w.Comments?.All() ?? new();
        var anchor = comments.Count == 1 ? CommentStore.Anchor(w.Body!, comments[0].Id) : null;
        Check(comments.Count == 1 && comments[0].Text == "Relire ce mot" && comments[0].Author == MainWindow.UserName && anchor is { } an
              && ReferenceEquals(an.StartParagraph, an.EndParagraph) && an.End - an.Start == "ligne".Length,
              "commentaire publié, ancré sur le texte sélectionné");
        var pane = w.CommentsPaneForTest!;
        Check(pane.Cards == 1, "volet : une carte");
        var reply = Descendants<Button>(pane).FirstOrDefault(b => b.Content as string == "Répondre" && Shown(b));
        if (reply is not null) Invoke(reply);
        if (pane.Editor is { } re) { re.Box.Text = "D'accord"; Invoke(re.Ok); }
        comments = w.Comments!.All();
        Check(comments.Count == 2 && comments[1].ParentParaId == comments[0].ParaId, "volet ▸ Répondre : réponse reliée au fil");
        var resolve = Descendants<Button>(pane).FirstOrDefault(b => b.Content as string == "Résoudre" && Shown(b));
        if (resolve is not null) Invoke(resolve);
        Check(w.Comments!.All()[0].Done, "volet ▸ Résoudre : fil résolu");

        // copie pour Word (modification suivie + fil de commentaires)
        var copy = Path.Combine(outDir, "revision.docx");
        try
        {
            using var ms = new MemoryStream();
            w.SavePackageCopy(ms);
            File.WriteAllBytes(copy, ms.ToArray());
            var again = DocxPackage.Open(copy);
            Check(new CommentStore(again).All().Count == 2 && again.Main!.Stories().SelectMany(s => Revisions.List(s)).Any(), "copie écrite et relue : révisions et commentaires");
        }
        catch (Exception ex) { Check(false, "copie pour Word : " + ex.Message); }

        // supprimer le fil, puis tout annuler
        var del = Descendants<Button>(pane).FirstOrDefault(b => b.Content as string == "Supprimer" && Shown(b));
        if (del is not null) Invoke(del);
        Check(w.Comments!.All().Count == 0 && anchor is not null && CommentStore.Anchor(w.Body!, w.Comments.All().FirstOrDefault()?.Id ?? "x") is null, "volet ▸ Supprimer : fil supprimé");
        Undo();
        Check(w.Comments!.All().Count == 2, "Ctrl+Z : fil revenu");
        Undo(); Undo(); Undo();                       // résolu, réponse, commentaire
        Check(w.Comments!.All().Count == 0 && !w.Body!.Paragraphs().SelectMany(p => p.Inlines).OfType<OpaqueInline>().Any(o => o.Xml.Name.LocalName.StartsWith("comment")),
              "Ctrl+Z ×3 : plus de commentaire ni de marque");
        w.ShowCommentsPane(false);

        // fin du suivi par le bouton ; toutes les modifications refusées
        Check(SelectTab(w, "Révision") && ClickButton(w, "Suivi des modifications") && !w.Tracking && w.SettingsForTest?.Elements().Any(e => e.Name.LocalName == "trackRevisions") == false,
              "Révision ▸ Suivi des modifications : arrêté (retiré de settings.xml)");
        w.ResolveAll(accept: false);
        Check(Revs().Count == 0 && Final(Para()) == original, "document revenu à son texte");
        SelectTab(w, "Accueil");
    }
}
