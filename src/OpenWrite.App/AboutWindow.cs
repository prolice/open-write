// Repris d'OpenPub 2.1.3 (src/OpenPub.App/AboutWindow.cs).
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace OpenWrite.App;

/// <summary>
/// Fenetre « À propos d'OpenWrite » aux couleurs de la charte : bandeau en degrade avec le logo,
/// informations sur la version et l'installation, credits, actions (aide, mises a jour).
/// </summary>
public sealed class AboutWindow : Window
{
    static Brush R(string key) => (Brush)Application.Current.FindResource(key);

    /// <param name="checkUpdates">Action « Rechercher des mises à jour » (la fenetre se ferme d'abord).</param>
    public AboutWindow(Action? checkUpdates = null)
    {
        Title = "À propos d'OpenWrite";
        Width = 560;
        SizeToContent = SizeToContent.Height;
        ResizeMode = ResizeMode.NoResize;
        WindowStyle = WindowStyle.None;
        AllowsTransparency = true;
        Background = Brushes.Transparent;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        ShowInTaskbar = false;
        FontFamily = (FontFamily)Application.Current.FindResource("UiFont");
        Foreground = R("Ink");
        FontSize = 13;

        var version = Update.UpdateService.CurrentVersion.ToString();
        var root = new Grid();
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        // ------------------------------------------------ bandeau
        var header = new Grid { Height = 156, ClipToBounds = true, Background = R("Accent") };
        header.SizeChanged += (_, e) => header.Clip = TopRounded(e.NewSize, 16);
        // bulles decoratives, comme sur le portail
        header.Children.Add(Bubble(190, -40, -70, 0.10));
        header.Children.Add(Bubble(120, 300, 70, 0.08));
        header.Children.Add(Bubble(60, 250, -18, 0.10));

        var brand = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(28, 0, 28, 0) };
        var tile = new Border
        {
            Width = 76, Height = 76, CornerRadius = new CornerRadius(20), Background = Brushes.White,
            Effect = new DropShadowEffect { Color = Color.FromRgb(0x1B, 0x1F, 0x6B), BlurRadius = 22, ShadowDepth = 6, Direction = 270, Opacity = 0.30 },
            Child = new Image
            {
                Source = new BitmapImage(new Uri("pack://application:,,,/Assets/openwrite-128.png")),
                Width = 56, Height = 56, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center,
            },
        };
        RenderOptions.SetBitmapScalingMode(tile.Child, BitmapScalingMode.HighQuality);
        brand.Children.Add(tile);
        var names = new StackPanel { VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(18, 0, 0, 0) };
        names.Children.Add(new TextBlock
        {
            Text = "OpenWrite", Foreground = Brushes.White, FontSize = 30, FontWeight = FontWeights.ExtraBold,
            FontFamily = (FontFamily)Application.Current.FindResource("TitleFont"),
        });
        names.Children.Add(new TextBlock { Text = "Traitement de texte pour Windows", Foreground = new SolidColorBrush(Color.FromArgb(0xE0, 255, 255, 255)), FontSize = 13.5 });
        names.Children.Add(new Border
        {
            HorizontalAlignment = HorizontalAlignment.Left, Margin = new Thickness(0, 10, 0, 0), Padding = new Thickness(10, 3, 10, 4),
            CornerRadius = new CornerRadius(9), Background = R("VersionBadge"),
            Child = new TextBlock { Text = $"Version {version}", Foreground = R("VersionText"), FontSize = 12, FontWeight = FontWeights.Bold },
        });
        brand.Children.Add(names);
        header.Children.Add(brand);

        var close = new Button
        {
            Content = new TextBlock { Text = "\uE711", FontFamily = (FontFamily)Application.Current.FindResource("Icons"), FontSize = 12, Foreground = Brushes.White },
            Width = 32, Height = 32, Padding = new Thickness(0), HorizontalAlignment = HorizontalAlignment.Right, VerticalAlignment = VerticalAlignment.Top,
            HorizontalContentAlignment = HorizontalAlignment.Center, Margin = new Thickness(0, 12, 12, 0),
            Background = new SolidColorBrush(Color.FromArgb(0x26, 255, 255, 255)), BorderThickness = new Thickness(0), ToolTip = "Fermer (Échap)",
        };
        close.Click += (_, _) => Close();
        header.Children.Add(close);
        header.MouseLeftButtonDown += (_, e) => { if (e.ButtonState == MouseButtonState.Pressed) DragMove(); };
        Grid.SetRow(header, 0);
        root.Children.Add(header);

        // ------------------------------------------------ contenu
        var body = new StackPanel { Margin = new Thickness(28, 22, 28, 20) };
        body.Children.Add(new TextBlock
        {
            Text = "OpenWrite remplace Microsoft Word pour un usage hors ligne : il ouvre et enregistre directement " +
                   "les documents .docx, sans rien perdre de ce qu'il ne sait pas encore modifier.",
            TextWrapping = TextWrapping.Wrap, LineHeight = 20, Foreground = R("Muted"),
        });

        var info = new Grid { Margin = new Thickness(0, 18, 0, 0) };
        info.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(124) });
        info.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        void Row(string glyph, string label, UIElement value)
        {
            int r = info.RowDefinitions.Count;
            info.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var l = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 7, 0, 7) };
            l.Children.Add(new TextBlock { Text = glyph, FontFamily = (FontFamily)Application.Current.FindResource("Icons"), FontSize = 14, Foreground = R("IconBrush"), VerticalAlignment = VerticalAlignment.Center, Width = 24 });
            l.Children.Add(new TextBlock { Text = label, Foreground = R("Muted"), VerticalAlignment = VerticalAlignment.Center });
            Grid.SetRow(l, r);
            info.Children.Add(l);
            if (value is FrameworkElement fe) { fe.VerticalAlignment = VerticalAlignment.Center; fe.Margin = new Thickness(0, 7, 0, 7); }
            Grid.SetRow(value, r);
            Grid.SetColumn(value, 1);
            info.Children.Add(value);
            if (r > 0)
            {
                var sep = new Border { Height = 1, Background = R("Line"), VerticalAlignment = VerticalAlignment.Top };
                Grid.SetRow(sep, r);
                Grid.SetColumnSpan(sep, 2);
                info.Children.Add(sep);
            }
        }
        TextBlock Value(string t) => new() { Text = t, FontWeight = FontWeights.Medium, TextWrapping = TextWrapping.Wrap };

        Row("\uE946", "Version", Value($"OpenWrite {version}"));
        Row("\uE7B8", "Installation", Value(Update.UpdateService.IsInstalled ? "Installée pour cet utilisateur" : "Version portable / développement"));
        Row("\uE895", "Mises à jour", Value(Update.UpdateService.FeedUrl is null ? "Non configurées"
            : Update.UpdateService.AutoCheck ? "Vérifiées automatiquement au démarrage" : "Vérification manuelle (menu Fichier)"));
        Row("\uE7F8", "Environnement", Value($".NET {Environment.Version.ToString(3)} · Windows {Environment.OSVersion.Version.Major}.{Environment.OSVersion.Version.Minor} (build {Environment.OSVersion.Version.Build})"));
        if (SiteUrl() is { } site)
        {
            var link = new TextBlock { Text = new Uri(site).Host, Foreground = R("Blue"), FontWeight = FontWeights.Medium, Cursor = Cursors.Hand, TextDecorations = TextDecorations.Underline, ToolTip = site, TextTrimming = TextTrimming.CharacterEllipsis };
            link.MouseLeftButtonUp += (_, _) => Launch(site);
            Row("\uE774", "Site", link);
        }
        body.Children.Add(new Border
        {
            Background = R("Chrome"), BorderBrush = R("Line"), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(12),
            Padding = new Thickness(14, 4, 14, 4), Margin = new Thickness(0, 18, 0, 0), Child = info,
        });
        info.Margin = new Thickness(0);

        body.Children.Add(new TextBlock
        {
            Text = "Polices de l'interface : DM Sans et Plus Jakarta Sans (SIL Open Font License). " +
                   "Lecture des .docx : Open XML SDK (licence MIT). Césure : motifs français hyph_fr (LibreOffice). Icônes : Segoe Fluent Icons.",
            TextWrapping = TextWrapping.Wrap, FontSize = 11.5, Foreground = R("Muted"), Margin = new Thickness(2, 14, 2, 0), LineHeight = 17,
        });
        Grid.SetRow(body, 1);
        root.Children.Add(body);

        // ------------------------------------------------ boutons
        var footer = new Border
        {
            Background = R("Chrome"), BorderBrush = R("Line"), BorderThickness = new Thickness(0, 1, 0, 0),
            CornerRadius = new CornerRadius(0, 0, 16, 16), Padding = new Thickness(20, 14, 20, 14),
        };
        var buttons = new DockPanel { LastChildFill = false };
        var help = IconButton("\uE897", "Aide en ligne");
        help.Click += (_, _) => Help.HelpService.Open(Help.HelpService.FileMenu, this);
        buttons.Children.Add(help);
        if (checkUpdates is not null && Update.UpdateService.FeedUrl is not null)
        {
            var upd = IconButton("\uE895", "Mises à jour…");
            upd.Margin = new Thickness(8, 0, 0, 0);
            upd.ToolTip = "Rechercher des mises à jour";
            upd.Click += (_, _) => { Close(); checkUpdates(); };
            buttons.Children.Add(upd);
        }
        var ok = new Button { Content = "Fermer", Style = (Style)Application.Current.FindResource("PillButton"), IsDefault = true, IsCancel = true, MinWidth = 96, Padding = new Thickness(18, 8, 18, 8) };
        ok.Click += (_, _) => Close();
        DockPanel.SetDock(ok, Dock.Right);
        buttons.Children.Add(ok);
        footer.Child = buttons;
        Grid.SetRow(footer, 2);
        root.Children.Add(footer);

        // carte arrondie avec ombre douce (marge transparente pour l'ombre)
        Content = new Border
        {
            Margin = new Thickness(20), CornerRadius = new CornerRadius(16), Background = R("Surface"),
            BorderBrush = R("Line"), BorderThickness = new Thickness(1),
            Effect = new DropShadowEffect { Color = Color.FromRgb(0x2E, 0x52, 0x9A), BlurRadius = 30, ShadowDepth = 8, Direction = 270, Opacity = 0.22 },
            Child = root,
        };
    }

    static Ellipse Bubble(double size, double x, double y, double opacity) => new()
    {
        Width = size, Height = size, Fill = new SolidColorBrush(Color.FromArgb((byte)(opacity * 255), 255, 255, 255)),
        HorizontalAlignment = HorizontalAlignment.Left, VerticalAlignment = VerticalAlignment.Top, Margin = new Thickness(x, y, 0, 0),
        IsHitTestVisible = false,
    };

    /// <summary>Rectangle aux coins superieurs arrondis (le bandeau suit l'arrondi de la carte).</summary>
    static Geometry TopRounded(Size s, double r)
    {
        var g = new StreamGeometry();
        using (var c = g.Open())
        {
            c.BeginFigure(new Point(0, s.Height), true, true);
            c.LineTo(new Point(0, r), false, false);
            c.ArcTo(new Point(r, 0), new Size(r, r), 0, false, SweepDirection.Clockwise, false, false);
            c.LineTo(new Point(s.Width - r, 0), false, false);
            c.ArcTo(new Point(s.Width, r), new Size(r, r), 0, false, SweepDirection.Clockwise, false, false);
            c.LineTo(new Point(s.Width, s.Height), false, false);
        }
        g.Freeze();
        return g;
    }

    static Button IconButton(string glyph, string text)
    {
        var sp = new StackPanel { Orientation = Orientation.Horizontal };
        sp.Children.Add(new TextBlock { Text = glyph, FontFamily = (FontFamily)Application.Current.FindResource("Icons"), FontSize = 14, Foreground = R("IconBrush"), VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 8, 0) });
        sp.Children.Add(new TextBlock { Text = text, VerticalAlignment = VerticalAlignment.Center });
        return new Button { Content = sp, Padding = new Thickness(12, 7, 14, 7) };
    }

    /// <summary>Page de telechargement : racine du site qui heberge latest.json.</summary>
    static string? SiteUrl()
    {
        if (Update.UpdateService.FeedUrl is not { } feed || !Uri.TryCreate(feed, UriKind.Absolute, out var u)) return null;
        return new Uri(u, "./").ToString();
    }

    static void Launch(string url)
    {
        try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); } catch { /* pas de navigateur */ }
    }
}
