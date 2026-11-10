// Repris d'OpenPub 2.1.3 (src/OpenPub.App/Help/HelpService.cs) : mêmes principes, rubriques d'OpenWrite.
using System.Diagnostics;
using System.Reflection;
using System.Windows;

namespace OpenWrite.App.Help;

/// <summary>
/// Liens vers la documentation en ligne (portail, dossier /aide/). Les cibles devront correspondre aux ancres
/// générées par tools/build-docs.py (lot 0, publication) : ne pas les renommer sans mettre la documentation à jour.
/// </summary>
public static class HelpService
{
    // pages et sections de la documentation (version 0.1)
    public const string Home = "index.html";
    public const string Interface = "interface.html#ruban";
    public const string FileMenu = "interface.html#fichier";
    public const string HomeTab = "interface.html#accueil";
    public const string Panels = "interface.html#volets";
    public const string OpenDocument = "documents.html#ouvrir";
    public const string Info = "documents.html#informations";
    public const string Roadmap = "feuille-de-route.html#versions";
    public const string Updates = "mises-a-jour.html#automatiques";
    public const string Shortcuts = "raccourcis.html#general";
    public const string Writing = "documents.html#ecrire";
    public const string FindReplace = "documents.html#rechercher";
    public const string Formatting = "raccourcis.html#mise-en-forme";
    public const string Clipboard = "documents.html#presse-papiers";
    public const string Links = "documents.html#liens";
    public const string Styles = "documents.html#styles";
    public const string Lists = "documents.html#listes";
    public const string Themes = "documents.html#themes";
    public const string PageSetup = "documents.html#mise-en-page";
    public const string HeadersFooters = "documents.html#en-tetes";
    public const string Tables = "documents.html#tableaux";

    public static IReadOnlyList<string> AllTopics => typeof(HelpService).GetFields(BindingFlags.Public | BindingFlags.Static)
        .Where(f => f.IsLiteral && f.FieldType == typeof(string)).Select(f => (string)f.GetRawConstantValue()!).ToList();

    /// <summary>Adresse de base de l'aide (définie à la compilation, voir OpenWrite.App.csproj) ; null tant que le portail n'existe pas.</summary>
    public static string? BaseUrl => Assembly.GetExecutingAssembly().GetCustomAttributes<AssemblyMetadataAttribute>()
        .FirstOrDefault(a => a.Key == "HelpBase")?.Value is { Length: > 0 } u ? u : null;

    public static string? UrlOf(string topic) => BaseUrl is { } b ? b.TrimEnd('/') + "/" + topic : null;

    /// <summary>Ouvre une page d'aide dans le navigateur par défaut (ou explique qu'elle n'est pas encore publiée).</summary>
    public static void Open(string topic, Window? owner = null)
    {
        if (UrlOf(topic) is not { } url)
        {
            var msg = "L'aide en ligne d'OpenWrite sera publiée avec le portail de téléchargement.";
            if (owner is not null) MessageBox.Show(owner, msg, "Aide", MessageBoxButton.OK, MessageBoxImage.Information);
            else MessageBox.Show(msg, "Aide", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
        catch (Exception) { /* pas de navigateur : rien à faire */ }
    }
}
