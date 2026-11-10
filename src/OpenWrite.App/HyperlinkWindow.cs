using System.Windows;
using System.Windows.Controls;
using static OpenWrite.App.DialogKit;

namespace OpenWrite.App;

/// <summary>Boîte « Lien hypertexte » (Ctrl+K) : texte à afficher, adresse (web, courriel) ou signet du document.</summary>
public sealed class HyperlinkWindow : Window
{
    readonly TextBox _text = Box(320), _address = Box(320);
    readonly ComboBox _bookmarks = new() { Width = 320 };

    public string Text => _text.Text;
    public string? Address => _address.Text.Trim() is { Length: > 0 } a ? a : null;
    /// <summary>« Supprimer le lien » choisi.</summary>
    public bool Remove { get; private set; }

    public HyperlinkWindow(string text, string? address, IReadOnlyList<string> bookmarks, bool canEditText, bool canRemove)
    {
        Setup(this, "Lien hypertexte");
        var root = new StackPanel { Margin = new Thickness(18) };
        root.Children.Add(Row("Texte à afficher", _text, 150));
        root.Children.Add(Row("Adresse", _address, 150));
        root.Children.Add(new TextBlock
        {
            Text = "Adresse web (https://…), de courriel (nom@exemple.be) ou #signet.", FontSize = 11.5, Margin = new Thickness(150, 0, 0, 6),
            Foreground = (System.Windows.Media.Brush)Application.Current.FindResource("Muted"),
        });
        if (bookmarks.Count > 0)
        {
            _bookmarks.Items.Add(new ComboBoxItem { Content = "(choisir un signet)", Tag = "" });
            foreach (var b in bookmarks) _bookmarks.Items.Add(new ComboBoxItem { Content = b, Tag = b });
            _bookmarks.SelectedIndex = 0;
            _bookmarks.SelectionChanged += (_, _) => { if (Tag(_bookmarks) is { Length: > 0 } b) _address.Text = "#" + b; };
            root.Children.Add(Row("Signet du document", _bookmarks, 150));
        }
        var buttons = Buttons(this, () =>
        {
            if (Address is null && !Remove)
            {
                MessageBox.Show(this, "Indiquez l'adresse du lien.", "Lien hypertexte", MessageBoxButton.OK, MessageBoxImage.Information);
                return false;
            }
            return true;
        });
        if (canRemove)
        {
            var remove = new Button { Content = "Supprimer le lien", MinWidth = 120, Margin = new Thickness(0, 0, 8, 0) };
            remove.Click += (_, _) => { Remove = true; DialogResult = true; };
            buttons.Children.Insert(0, remove);
        }
        root.Children.Add(buttons);
        Content = root;
        _text.Text = text;
        _text.IsEnabled = canEditText;
        _address.Text = address ?? "";
        Loaded += (_, _) => { (canEditText && text.Length == 0 ? _text : _address).Focus(); _address.SelectAll(); };
    }
}
