using System.Windows;
using System.Windows.Controls;
using static OpenWrite.App.DialogKit;

namespace OpenWrite.App;

/// <summary>Boîte « Mise en page » (lot 5) : marges, distances de l'en-tête et du pied, orientation, papier, colonnes ; section ou document.</summary>
public sealed class PageSetupWindow : Window
{
    const double PtPerCm = 72 / 2.54;
    public static readonly (string Name, double Wcm, double Hcm)[] Papers =
    {
        ("A4", 21, 29.7), ("A5", 14.8, 21), ("A3", 29.7, 42), ("B5", 17.6, 25), ("Lettre US", 21.59, 27.94), ("Légal US", 21.59, 35.56), ("Enveloppe DL", 11, 22),
    };

    readonly TextBox _top = Box(), _bottom = Box(), _left = Box(), _right = Box(), _header = Box(), _footer = Box(), _pw = Box(), _ph = Box(), _cols = Box(50), _space = Box();
    readonly ComboBox _orient = Combo(140, ("Portrait", "portrait"), ("Paysage", "landscape"));
    readonly ComboBox _paper = new() { Width = 160 };
    readonly ComboBox _scope = Combo(200, ("Cette section", "section"), ("Tout le document", "all"));
    readonly CheckBox _sep = new() { Content = "Trait entre les colonnes", Margin = new Thickness(0, 4, 0, 0) };

    public (double Top, double Bottom, double Left, double Right, double Header, double Footer) Margins { get; private set; }
    public bool Landscape => Tag(_orient) == "landscape";
    /// <summary>Papier en points, sens portrait.</summary>
    public (double W, double H) Paper { get; private set; }
    public int Columns { get; private set; }
    public double ColumnSpace { get; private set; }
    public bool Separator => _sep.IsChecked == true;
    public bool AllSections => Tag(_scope) == "all";

    public PageSetupWindow((double Top, double Bottom, double Left, double Right, double Header, double Footer) m, double pageW, double pageH, bool landscape,
        int columns, double columnSpace, bool separator, bool columnsTab = false)
    {
        Setup(this, columnsTab ? "Colonnes" : "Mise en page");
        var root = new StackPanel { Margin = new Thickness(18), Width = 420 };
        if (!columnsTab)
        {
            root.Children.Add(Section("Marges (cm)"));
            root.Children.Add(Row("Haut", _top)); root.Children.Add(Row("Bas", _bottom));
            root.Children.Add(Row("Gauche", _left)); root.Children.Add(Row("Droite", _right));
            root.Children.Add(Row("En-tête (du bord)", _header)); root.Children.Add(Row("Pied (du bord)", _footer));
            root.Children.Add(Section("Papier"));
            root.Children.Add(Row("Orientation", _orient));
            foreach (var (name, w, h) in Papers) _paper.Items.Add(new ComboBoxItem { Content = $"{name} ({w.ToString(Cu)} × {h.ToString(Cu)} cm)", Tag = name });
            _paper.Items.Add(new ComboBoxItem { Content = "Personnalisé", Tag = "" });
            root.Children.Add(Row("Format", _paper));
            root.Children.Add(Row("Largeur (cm)", _pw)); root.Children.Add(Row("Hauteur (cm)", _ph));
        }
        root.Children.Add(Section("Colonnes"));
        root.Children.Add(Row("Nombre", _cols));
        root.Children.Add(Row("Espacement (cm)", _space));
        root.Children.Add(_sep);
        root.Children.Add(Section("Appliquer"));
        root.Children.Add(Row("À", _scope));
        root.Children.Add(Buttons(this, Build));
        Content = root;

        string Cm(double pt) => (pt / PtPerCm).ToString("0.##", Cu);
        _top.Text = Cm(m.Top); _bottom.Text = Cm(m.Bottom); _left.Text = Cm(m.Left); _right.Text = Cm(m.Right); _header.Text = Cm(m.Header); _footer.Text = Cm(m.Footer);
        Select(_orient, landscape ? "landscape" : "portrait");
        double small = Math.Min(pageW, pageH), big = Math.Max(pageW, pageH);
        _pw.Text = Cm(small); _ph.Text = Cm(big);
        var match = Papers.FirstOrDefault(p => Math.Abs(p.Wcm * PtPerCm - small) < 2 && Math.Abs(p.Hcm * PtPerCm - big) < 2);
        Select(_paper, match.Name ?? "");
        _paper.SelectionChanged += (_, _) =>
        {
            if (Papers.FirstOrDefault(p => p.Name == Tag(_paper)) is { Name: not null } p) { _pw.Text = p.Wcm.ToString(Cu); _ph.Text = p.Hcm.ToString(Cu); }
        };
        _cols.Text = columns.ToString(Cu);
        _space.Text = Cm(columnSpace);
        _sep.IsChecked = separator;
    }

    bool Build()
    {
        double? Cm(TextBox b, double min, double max) => Parse(b.Text) is { } v && v >= min && v <= max ? v * PtPerCm : null;
        var top = Cm(_top, 0, 50); var bottom = Cm(_bottom, 0, 50); var left = Cm(_left, 0, 50); var right = Cm(_right, 0, 50);
        var header = Cm(_header, 0, 50); var footer = Cm(_footer, 0, 50); var pw = Cm(_pw, 2, 300); var ph = Cm(_ph, 2, 300);
        var space = Cm(_space, 0, 20);
        int cols = int.TryParse(_cols.Text.Trim(), out var n) ? n : 0;
        bool full = _top.IsLoaded || _top.Parent is not null;
        if ((full && (top is null || bottom is null || left is null || right is null || header is null || footer is null || pw is null || ph is null)) || space is null || cols is < 1 or > 12)
        {
            MessageBox.Show(this, "Une valeur est incorrecte (marges et papier en centimètres, 1 à 12 colonnes).", Title, MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }
        Margins = (top ?? 0, bottom ?? 0, left ?? 0, right ?? 0, header ?? 0, footer ?? 0);
        Paper = (pw ?? 0, ph ?? 0);
        Columns = cols;
        ColumnSpace = space.Value;
        return true;
    }
}
