using System.Windows;
using System.Windows.Controls;
using static OpenWrite.App.DialogKit;

namespace OpenWrite.App;

// Boîtes des objets graphiques (lot 7) : texte de remplacement, taille et rotation, rognage. Valeurs en centimètres et
// en degrés, comme Word en français.

/// <summary>Boîte « Texte de remplacement » : description lue par les lecteurs d'écran.</summary>
public sealed class AltTextWindow : Window
{
    readonly TextBox _text = new() { Width = 360, Height = 90, TextWrapping = TextWrapping.Wrap, AcceptsReturn = true, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
    public string Description => _text.Text.Trim();

    public AltTextWindow(string current)
    {
        Setup(this, "Texte de remplacement");
        var root = new StackPanel { Margin = new Thickness(18), Width = 380 };
        root.Children.Add(new TextBlock
        {
            Text = "Décrivez l'objet et son rôle pour une personne qui ne le voit pas (une ou deux phrases).",
            TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 0, 0, 8),
        });
        root.Children.Add(_text);
        _text.Text = current;
        root.Children.Add(Buttons(this, () => true));
        Content = root;
        Loaded += (_, _) => { _text.Focus(); _text.SelectAll(); };
    }
}

/// <summary>Boîte « Taille » : hauteur, largeur (proportions gardées si demandé), rotation.</summary>
public sealed class ObjectSizeWindow : Window
{
    const double PtPerCm = 72 / 2.54;
    readonly TextBox _w = Box(), _h = Box(), _rot = Box();
    readonly CheckBox _ratio = new() { Content = "Conserver les proportions", Margin = new Thickness(0, 6, 0, 0) };
    public double ResultWidth { get; private set; }
    public double ResultHeight { get; private set; }
    public double ResultRotation { get; private set; }

    public ObjectSizeWindow(double width, double height, double rotation, bool picture)
    {
        Setup(this, "Taille");
        var root = new StackPanel { Margin = new Thickness(18), Width = 320 };
        string Cm(double pt) => (pt / PtPerCm).ToString("0.##", Cu);
        root.Children.Add(Section("Taille"));
        root.Children.Add(Row("Hauteur (cm)", _h));
        root.Children.Add(Row("Largeur (cm)", _w));
        root.Children.Add(_ratio);
        root.Children.Add(Section("Rotation"));
        root.Children.Add(Row("Rotation (°)", _rot));
        _w.Text = Cm(width);
        _h.Text = Cm(height);
        _rot.Text = rotation.ToString("0.#", Cu);
        _ratio.IsChecked = picture;
        bool sync = false;
        // proportions gardées : changer une dimension recalcule l'autre
        void Linked(TextBox from, TextBox to, double ratio)
        {
            from.TextChanged += (_, _) =>
            {
                if (sync || _ratio.IsChecked != true || Parse(from.Text) is not { } v || v <= 0) return;
                sync = true;
                to.Text = (v * ratio).ToString("0.##", Cu);
                sync = false;
            };
        }
        if (width > 0 && height > 0)
        {
            Linked(_w, _h, height / width);
            Linked(_h, _w, width / height);
        }
        root.Children.Add(Buttons(this, () =>
        {
            if (Parse(_w.Text) is not { } w || Parse(_h.Text) is not { } h || w <= 0 || h <= 0 || w > 60 || h > 60)
            {
                MessageBox.Show(this, "Largeur et hauteur : de 0,01 à 60 cm.", Title, MessageBoxButton.OK, MessageBoxImage.Information);
                return false;
            }
            ResultWidth = w * PtPerCm;
            ResultHeight = h * PtPerCm;
            ResultRotation = Parse(_rot.Text) ?? rotation;
            return true;
        }));
        Content = root;
        Loaded += (_, _) => { _h.Focus(); _h.SelectAll(); };
    }
}

/// <summary>Boîte « Rogner » : part retirée de chaque bord de l'image, en pour cent.</summary>
public sealed class CropWindow : Window
{
    readonly TextBox _l = Box(), _t = Box(), _r = Box(), _b = Box();
    public (double L, double T, double R, double B) Result { get; private set; }

    public CropWindow((double L, double T, double R, double B) crop)
    {
        Setup(this, "Rogner");
        var root = new StackPanel { Margin = new Thickness(18), Width = 300 };
        root.Children.Add(new TextBlock { Text = "Part de l'image retirée de chaque bord (%). 0 partout : image entière.", TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 0, 0, 8) });
        string P(double v) => (v * 100).ToString("0.#", Cu);
        root.Children.Add(Row("Haut (%)", _t));
        root.Children.Add(Row("Bas (%)", _b));
        root.Children.Add(Row("Gauche (%)", _l));
        root.Children.Add(Row("Droite (%)", _r));
        (_l.Text, _t.Text, _r.Text, _b.Text) = (P(crop.L), P(crop.T), P(crop.R), P(crop.B));
        root.Children.Add(Buttons(this, () =>
        {
            double? V(TextBox x) => Parse(x.Text) is { } v && v >= 0 && v < 100 ? v / 100 : null;
            if (V(_l) is not { } l || V(_t) is not { } t || V(_r) is not { } r || V(_b) is not { } b || l + r >= 0.98 || t + b >= 0.98)
            {
                MessageBox.Show(this, "De 0 à 99 % par bord, en gardant une partie de l'image.", Title, MessageBoxButton.OK, MessageBoxImage.Information);
                return false;
            }
            Result = (l, t, r, b);
            return true;
        }));
        Content = root;
    }
}
