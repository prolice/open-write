using System.Windows;
using System.Windows.Controls;
using static OpenWrite.App.DialogKit;

namespace OpenWrite.App;

// Boîtes des tableaux (lot 6) : insérer un tableau, fractionner les cellules, propriétés du tableau, trier, formule,
// séparateur (convertir texte ↔ tableau). Valeurs en centimètres comme Word en français.

/// <summary>Boîte « Insérer un tableau » : colonnes, lignes, comportement de l'ajustement automatique.</summary>
public sealed class InsertTableWindow : Window
{
    readonly TextBox _cols = Box(60), _rows = Box(60);
    readonly ComboBox _fit = Combo(260, ("Largeur de colonne fixe", "fixed"), ("Ajuster au contenu", "content"), ("Ajuster à la fenêtre", "window"));
    public int Columns { get; private set; }
    public int Rows { get; private set; }
    public string Fit => Tag(_fit);

    public InsertTableWindow(int cols = 5, int rows = 2)
    {
        Setup(this, "Insérer un tableau");
        var root = new StackPanel { Margin = new Thickness(18), Width = 380 };
        root.Children.Add(Section("Taille du tableau"));
        root.Children.Add(Row("Nombre de colonnes", _cols, 160));
        root.Children.Add(Row("Nombre de lignes", _rows, 160));
        root.Children.Add(Section("Ajustement automatique"));
        root.Children.Add(Row("Comportement", _fit, 110));
        _cols.Text = cols.ToString(Cu);
        _rows.Text = rows.ToString(Cu);
        root.Children.Add(Buttons(this, () =>
        {
            if (!int.TryParse(_cols.Text, out var c) || c < 1 || c > 63 || !int.TryParse(_rows.Text, out var r) || r < 1 || r > 500)
            {
                MessageBox.Show(this, "De 1 à 63 colonnes et de 1 à 500 lignes.", Title, MessageBoxButton.OK, MessageBoxImage.Information);
                return false;
            }
            (Columns, Rows) = (c, r);
            return true;
        }));
        Content = root;
        Loaded += (_, _) => { _cols.Focus(); _cols.SelectAll(); };
    }
}

/// <summary>Boîte « Fractionner les cellules » : nombre de colonnes et de lignes.</summary>
public sealed class SplitCellsWindow : Window
{
    readonly TextBox _cols = Box(60), _rows = Box(60);
    public int Columns { get; private set; }
    public int Rows { get; private set; }

    public SplitCellsWindow()
    {
        Setup(this, "Fractionner les cellules");
        var root = new StackPanel { Margin = new Thickness(18), Width = 300 };
        root.Children.Add(Row("Nombre de colonnes", _cols, 160));
        root.Children.Add(Row("Nombre de lignes", _rows, 160));
        _cols.Text = "2";
        _rows.Text = "1";
        root.Children.Add(Buttons(this, () =>
        {
            if (!int.TryParse(_cols.Text, out var c) || c < 1 || c > 63 || !int.TryParse(_rows.Text, out var r) || r < 1 || r > 100) return false;
            (Columns, Rows) = (c, r);
            return true;
        }));
        Content = root;
        Loaded += (_, _) => { _cols.Focus(); _cols.SelectAll(); };
    }
}

/// <summary>
/// Boîte « Propriétés du tableau » : tableau (alignement, retrait), ligne (hauteur, règle, fractionnement, en-tête
/// répété), colonne (largeur), cellule (alignement vertical), marges de cellule par défaut. Seules les valeurs
/// modifiées sont appliquées (null sinon).
/// </summary>
public sealed class TablePropertiesWindow : Window
{
    const double PtPerCm = 72 / 2.54;
    readonly ComboBox _align = Combo(150, ("À gauche", "left"), ("Centré", "center"), ("À droite", "right"));
    readonly TextBox _indent = Box(), _height = Box(), _width = Box(), _mTop = Box(), _mBottom = Box(), _mLeft = Box(), _mRight = Box();
    readonly ComboBox _rule = Combo(150, ("Au moins", "atLeast"), ("Exactement", "exact"));
    readonly CheckBox _allowBreak = new() { Content = "Autoriser le fractionnement des lignes sur plusieurs pages", Margin = new Thickness(0, 4, 0, 0) };
    readonly CheckBox _header = new() { Content = "Répéter en haut de chaque page en tant que ligne d'en-tête", Margin = new Thickness(0, 4, 0, 0) };
    readonly ComboBox _valign = Combo(150, ("Haut", "top"), ("Centre", "center"), ("Bas", "bottom"));

    public string? Alignment { get; private set; }
    public double? Indent { get; private set; }
    /// <summary>Hauteur de ligne (points ; 0 = automatique), règle.</summary>
    public double? RowHeight { get; private set; }
    public string Rule => Tag(_rule);
    public bool? AllowBreak { get; private set; }
    public bool? HeaderRow { get; private set; }
    public double? ColumnWidth { get; private set; }
    public string? VerticalAlignment { get; private set; }
    public (double? Top, double? Bottom, double? Left, double? Right) Margins { get; private set; }

    public TablePropertiesWindow(string align, double indent, double? rowHeight, string rule, bool allowBreak, bool header, double columnWidth, string valign,
        (double Top, double Bottom, double Left, double Right) margins)
    {
        Setup(this, "Propriétés du tableau");
        var root = new StackPanel { Margin = new Thickness(18), Width = 470 };
        root.Children.Add(Section("Tableau"));
        root.Children.Add(Row("Alignement", _align));
        root.Children.Add(Row("Retrait à gauche (cm)", _indent));
        root.Children.Add(Section("Ligne"));
        root.Children.Add(Row("Hauteur (cm, vide = auto)", _height, 170));
        root.Children.Add(Row("Hauteur de ligne", _rule, 170));
        root.Children.Add(_allowBreak);
        root.Children.Add(_header);
        root.Children.Add(Section("Colonne"));
        root.Children.Add(Row("Largeur (cm)", _width));
        root.Children.Add(Section("Cellule"));
        root.Children.Add(Row("Alignement vertical", _valign));
        root.Children.Add(Section("Marges de cellule par défaut (cm)"));
        var mar = new WrapPanel();
        mar.Children.Add(Row("Haut", _mTop, 50)); mar.Children.Add(Row("  Bas", _mBottom, 50));
        mar.Children.Add(Row("Gauche", _mLeft, 50)); mar.Children.Add(Row("  Droite", _mRight, 50));
        root.Children.Add(mar);

        string Cm(double pt) => (pt / PtPerCm).ToString("0.##", Cu);
        Select(_align, align);
        _indent.Text = Cm(indent);
        _height.Text = rowHeight is { } h && h > 0 ? Cm(h) : "";
        Select(_rule, rule);
        _allowBreak.IsChecked = allowBreak;
        _header.IsChecked = header;
        _width.Text = Cm(columnWidth);
        Select(_valign, valign);
        _mTop.Text = Cm(margins.Top); _mBottom.Text = Cm(margins.Bottom); _mLeft.Text = Cm(margins.Left); _mRight.Text = Cm(margins.Right);

        root.Children.Add(Buttons(this, () =>
        {
            double? Read(TextBox b, string before) => b.Text == before ? null : Parse(b.Text) is { } v ? v * PtPerCm : null;
            if (Tag(_align) != align) Alignment = Tag(_align);
            Indent = Read(_indent, Cm(indent));
            var hText = rowHeight is { } h0 && h0 > 0 ? Cm(h0) : "";
            if (_height.Text.Trim() != hText || Tag(_rule) != rule) RowHeight = _height.Text.Trim().Length == 0 ? 0 : Parse(_height.Text) is { } hv ? hv * PtPerCm : null;
            if (_allowBreak.IsChecked != allowBreak) AllowBreak = _allowBreak.IsChecked == true;
            if (_header.IsChecked != header) HeaderRow = _header.IsChecked == true;
            ColumnWidth = Read(_width, Cm(columnWidth));
            if (Tag(_valign) != valign) VerticalAlignment = Tag(_valign);
            Margins = (Read(_mTop, Cm(margins.Top)), Read(_mBottom, Cm(margins.Bottom)), Read(_mLeft, Cm(margins.Left)), Read(_mRight, Cm(margins.Right)));
            return true;
        }));
        Content = root;
    }
}

/// <summary>Boîte « Trier » : colonne, ordre, ligne d'en-tête.</summary>
public sealed class SortWindow : Window
{
    readonly ComboBox _col = new() { Width = 220 };
    readonly ComboBox _order = Combo(160, ("Croissant", "asc"), ("Décroissant", "desc"));
    readonly CheckBox _header = new() { Content = "Ma liste a une ligne d'en-tête", Margin = new Thickness(0, 8, 0, 0) };
    public int Column => _col.SelectedIndex;
    public bool Ascending => Tag(_order) == "asc";
    public bool Header => _header.IsChecked == true;

    public SortWindow(IReadOnlyList<string> columns, int current, bool header)
    {
        Setup(this, "Trier");
        var root = new StackPanel { Margin = new Thickness(18), Width = 380 };
        foreach (var c in columns) _col.Items.Add(c);
        _col.SelectedIndex = Math.Clamp(current, 0, Math.Max(0, columns.Count - 1));
        root.Children.Add(Row("Trier par", _col, 100));
        root.Children.Add(Row("Ordre", _order, 100));
        _header.IsChecked = header;
        root.Children.Add(_header);
        root.Children.Add(Buttons(this, () => _col.SelectedIndex >= 0));
        Content = root;
    }
}

/// <summary>Boîte « Formule » : formule de tableau (=SUM(ABOVE)…), résultat calculé à l'insertion.</summary>
public sealed class FormulaWindow : Window
{
    readonly TextBox _formula = Box(260);
    public string Formula => _formula.Text.Trim();

    public FormulaWindow(string suggestion)
    {
        Setup(this, "Formule");
        var root = new StackPanel { Margin = new Thickness(18), Width = 380 };
        root.Children.Add(Row("Formule", _formula, 80));
        root.Children.Add(new TextBlock
        {
            Text = "SUM, AVERAGE, MIN, MAX, COUNT ou PRODUCT de ABOVE (au-dessus), BELOW, LEFT (à gauche) ou RIGHT.",
            TextWrapping = TextWrapping.Wrap, Foreground = (System.Windows.Media.Brush)Application.Current.FindResource("Muted"), Margin = new Thickness(0, 6, 0, 0),
        });
        _formula.Text = suggestion;
        root.Children.Add(Buttons(this, () => Formula.Length > 1));
        Content = root;
        Loaded += (_, _) => { _formula.Focus(); _formula.SelectAll(); };
    }
}

/// <summary>Séparateur de texte (convertir le texte en tableau, le tableau en texte).</summary>
public sealed class SeparatorWindow : Window
{
    readonly ComboBox _sep = Combo(200, ("Tabulations", "\t"), ("Points-virgules", ";"), ("Virgules", ","), ("Autre", "other"));
    readonly TextBox _other = Box(40);
    public string Separator => Tag(_sep) == "other" ? (_other.Text.Length > 0 ? _other.Text[..1] : "-") : Tag(_sep);

    public SeparatorWindow(string title, string? info = null)
    {
        Setup(this, title);
        var root = new StackPanel { Margin = new Thickness(18), Width = 360 };
        if (info is not null) root.Children.Add(new TextBlock { Text = info, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 0, 0, 8) });
        root.Children.Add(Row("Séparer le texte par", _sep, 140));
        root.Children.Add(Row("Autre caractère", _other, 140));
        _other.Text = "-";
        root.Children.Add(Buttons(this, () => true));
        Content = root;
    }
}
