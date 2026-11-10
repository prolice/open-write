using System.Windows;
using System.Windows.Automation;
using System.Windows.Automation.Peers;
using System.Windows.Automation.Provider;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;

namespace OpenWrite.App;

// Auto-test de l'interface : les commandes passent par les vrais contrôles (menus contextuels, sous-menus, onglets et
// boutons du ruban) au moyen de l'automatisation de l'interface (UI Automation), comme un lecteur d'écran ou un clic.
// Les autres étapes de l'auto-test appellent les méthodes de la fenêtre directement : elles ne voient pas un menu qui
// ne s'ouvre pas ou un bouton mal branché.
public static partial class SelfTest
{
    static AutomationPeer Peer(UIElement e) => UIElementAutomationPeer.CreatePeerForElement(e);

    /// <summary>Éléments visuels de type <typeparamref name="T"/> sous <paramref name="root"/>.</summary>
    static IEnumerable<T> Descendants<T>(DependencyObject root) where T : DependencyObject
    {
        var stack = new Stack<DependencyObject>();
        stack.Push(root);
        while (stack.Count > 0)
        {
            var d = stack.Pop();
            if (d is T t) yield return t;
            if (d is not Visual and not System.Windows.Media.Media3D.Visual3D) continue;
            for (int i = VisualTreeHelper.GetChildrenCount(d) - 1; i >= 0; i--) stack.Push(VisualTreeHelper.GetChild(d, i));
        }
    }

    /// <summary>L'entrée est réellement affichée (dans une fenêtre, visible).</summary>
    static bool Shown(UIElement e) => e.IsVisible && PresentationSource.FromVisual(e) is not null && e.RenderSize.Width > 0;

    static void Expand(MenuItem mi)
    {
        ((IExpandCollapseProvider)Peer(mi).GetPattern(PatternInterface.ExpandCollapse)).Expand();
        Pump();
        Idle();
    }

    static void Collapse(MenuItem mi)
    {
        if (!mi.IsSubmenuOpen) return;
        ((IExpandCollapseProvider)Peer(mi).GetPattern(PatternInterface.ExpandCollapse)).Collapse();
        Pump();
    }

    static string HeaderOf(MenuItem mi) => mi.Header as string ?? mi.Header?.ToString() ?? "";

    /// <summary>
    /// Ouvre chaque sous-menu d'un menu (récursivement) et vérifie que ses entrées s'affichent ; renvoie les problèmes.
    /// </summary>
    static List<string> AuditMenu(ItemsControl menu, string path)
    {
        var problems = new List<string>();
        foreach (var mi in menu.Items.OfType<MenuItem>().ToList())
        {
            if (!mi.HasItems || !mi.IsEnabled) continue;
            string name = $"{path} ▸ {HeaderOf(mi)}";
            Expand(mi);
            var children = mi.Items.OfType<MenuItem>().ToList();
            int hidden = children.Count(c => !Shown(c));
            if (!mi.IsSubmenuOpen) problems.Add($"{name} : le sous-menu ne s'ouvre pas");
            else if (hidden > 0) problems.Add($"{name} : {hidden} entrée(s) sur {children.Count} non affichée(s)");
            else problems.AddRange(AuditMenu(mi, name));
            Collapse(mi);
        }
        return problems;
    }

    /// <summary>Choisit une entrée de menu par son chemin (« Supprimer », « Lignes ») en ouvrant les sous-menus ; vrai si cliquée.</summary>
    static bool InvokeMenu(ContextMenu? menu, params string[] path)
    {
        if (menu is null || !menu.IsOpen) return false;
        ItemsControl current = menu;
        for (int i = 0; i < path.Length; i++)
        {
            var mi = current.Items.OfType<MenuItem>().FirstOrDefault(m => HeaderOf(m) == path[i]);
            if (mi is null || !mi.IsEnabled || !Shown(mi)) { menu.IsOpen = false; return false; }
            if (i < path.Length - 1)
            {
                Expand(mi);
                current = mi;
                continue;
            }
            ((IInvokeProvider)Peer(mi).GetPattern(PatternInterface.Invoke)).Invoke();
            Pump();
            Idle();
        }
        if (menu.IsOpen) menu.IsOpen = false;
        Pump();
        return true;
    }

    /// <summary>Onglet du ruban (« Accueil », « Disposition »…) choisi comme d'un clic ; vrai s'il est affiché et choisi.</summary>
    static bool SelectTab(Window w, string label)
    {
        var tab = Descendants<RadioButton>(w).FirstOrDefault(r => r.Content as string == label && Shown(r));
        if (tab is null) return false;
        ((ISelectionItemProvider)Peer(tab).GetPattern(PatternInterface.SelectionItem)).Select();
        Pump();
        Idle();
        return tab.IsChecked == true;
    }

    /// <summary>Bouton affiché du ruban, par son libellé, cliqué ; vrai s'il existe, est actif et a été cliqué.</summary>
    static bool ClickButton(Window w, string label)
    {
        var b = Descendants<ButtonBase>(w).FirstOrDefault(x => x.Content as string == label && Shown(x));
        if (b is null || !b.IsEnabled) return false;
        if (Peer(b).GetPattern(PatternInterface.Invoke) is IInvokeProvider inv) inv.Invoke();
        else if (Peer(b).GetPattern(PatternInterface.Toggle) is IToggleProvider tog) tog.Toggle();
        else return false;
        Pump();
        Idle();
        return true;
    }

    /// <summary>Étape « interface » de l'auto-test : menus contextuels (texte, tableau, objet) et menus du ruban des tableaux.</summary>
    static void InterfaceChecks(MainWindow w)
    {
        Console.WriteLine("Interface (menus et ruban, par l'automatisation de l'interface) :");
        Table? TableBlock() => w.Body!.Blocks.OfType<Table>().FirstOrDefault();
        (int Rows, int Cols) Size() => TableBlock() is { } t ? (new TableOps.Map(t).Rows.Count, new TableOps.Map(t).Columns) : (0, 0);
        void Undo() { ApplicationCommands.Undo.Execute(null, w); Pump(); Idle(); }

        // menu contextuel dans le texte : il s'ouvre, ses entrées s'affichent
        w.FindNext("Deuxième", true, false);
        var menu = w.OpenContextMenuForTest();
        Pump();
        Idle();
        Check(menu is { IsOpen: true } && menu.Items.OfType<MenuItem>().Any(m => HeaderOf(m) == "Coller" && Shown(m)), "menu contextuel du texte affiché");
        if (menu is not null) menu.IsOpen = false;

        // menu contextuel dans un tableau : chaque sous-menu s'ouvre et montre ses entrées
        var start = Size();
        w.FindNext("A2", true, false);
        menu = w.OpenContextMenuForTest();
        Pump();
        Idle();
        var problems = menu is null ? new List<string> { "menu absent" } : AuditMenu(menu, "tableau");
        Check(menu is not null && menu.Items.OfType<MenuItem>().Any(m => HeaderOf(m) == "Supprimer") && problems.Count == 0,
              "menu contextuel du tableau : sous-menus Insérer et Supprimer affichés" + (problems.Count > 0 ? " (" + string.Join(" ; ", problems) + ")" : ""));
        Check(InvokeMenu(menu, "Insérer", "Lignes en dessous") && Size().Rows == start.Rows + 1,
              $"clic droit ▸ Insérer ▸ Lignes en dessous : {Size().Rows} lignes (avant : {start.Rows})");
        w.FindNext("A2", true, false);
        Check(InvokeMenu(w.OpenContextMenuForTest(), "Supprimer", "Lignes") && Size().Rows == start.Rows,
              $"clic droit ▸ Supprimer ▸ Lignes : {Size().Rows} lignes");
        w.FindNext("A2", true, false);
        Check(InvokeMenu(w.OpenContextMenuForTest(), "Supprimer", "Colonnes") && Size().Cols == start.Cols - 1,
              $"clic droit ▸ Supprimer ▸ Colonnes : {Size().Cols} colonnes (avant : {start.Cols})");
        w.FindNext("A2", true, false);
        Check(InvokeMenu(w.OpenContextMenuForTest(), "Supprimer", "Tableau") && TableBlock() is null, "clic droit ▸ Supprimer ▸ Tableau : tableau supprimé");
        Undo();
        Undo();
        Undo();
        Undo();
        Check(Size() == start, $"Ctrl+Z ×4 : tableau d'origine ({Size().Rows} × {Size().Cols})");

        // ruban : onglet contextuel « Disposition », bouton Supprimer, menu du bouton
        w.FindNext("A2", true, false);
        Pump();
        Check(SelectTab(w, "Disposition"), "onglet Disposition affiché dans un tableau et choisi");
        Check(ClickButton(w, "Supprimer") && w.LastMenu is { IsOpen: true }, "Disposition ▸ Supprimer : menu ouvert");
        Check(InvokeMenu(w.LastMenu, "Supprimer le tableau") && TableBlock() is null, "Disposition ▸ Supprimer ▸ Supprimer le tableau");
        Undo();
        Check(Size() == start, "Ctrl+Z : tableau revenu");
        w.FindNext("A2", true, false);
        SelectTab(w, "Disposition");
        Check(ClickButton(w, "Supprimer") && InvokeMenu(w.LastMenu, "Supprimer les colonnes") && Size().Cols == start.Cols - 1, "Disposition ▸ Supprimer ▸ Supprimer les colonnes");
        Undo();
        SelectTab(w, "Accueil");

        // objet : menu contextuel (Habillage ▸, Ordre ▸, Rotation ▸)
        int Drawings() => w.Body!.Paragraphs().SelectMany(p => p.Runs()).SelectMany(r => r.Content).OfType<DrawingContent>().Count();
        int floatsBefore = w.CurrentLayout!.Pages[0].Floats.Count, drawingsBefore = Drawings();
        w.FindNext("Deuxième", true, false);
        w.InsertShape("rect");
        Pump();
        menu = w.OpenContextMenuForTest();
        Pump();
        Idle();
        problems = menu is null ? new List<string> { "menu absent" } : AuditMenu(menu, "objet");
        Check(menu is not null && menu.Items.OfType<MenuItem>().Any(m => HeaderOf(m) == "Habillage") && problems.Count == 0,
              "menu contextuel d'un objet : sous-menus affichés" + (problems.Count > 0 ? " (" + string.Join(" ; ", problems) + ")" : ""));
        Check(InvokeMenu(menu, "Habillage", "Carré") && w.CurrentLayout!.Pages[0].Floats.Any(f => f.Object.Wrap == Core.Drawing.WrapKind.Square),
              "clic droit ▸ Habillage ▸ Carré");
        Undo();
        Undo();
        Check(w.CurrentLayout!.Pages[0].Floats.Count == floatsBefore && Drawings() == drawingsBefore, "Ctrl+Z ×2 : forme retirée");
    }
}
