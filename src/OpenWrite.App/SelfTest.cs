// Captures repries d'OpenPub 2.1.3 (src/OpenPub.App/SelfTest.cs : Pump, ForceContentSize, Capture, Screenshot).
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using System.IO.Compression;
using System.Windows.Input;
using OpenWrite.Core.Model;
using OpenWrite.Docx;

namespace OpenWrite.App;

/// <summary>
/// Captures d'écran de la vraie fenêtre, hors écran (documentation, portail, vérification visuelle) :
/// OpenWrite.exe --screenshot document.docx|- image.png [largeur hauteur] [scène].
/// Scènes : nom d'un onglet (TabInsert, TabReview…), « file-menu », « about », « page-N » (vue défilée à la page N),
/// « selection » (quelques mots sélectionnés sur la première page).
/// Auto-test de bout en bout (OpenWrite.exe --selftest dossier) : la vraie fenêtre, hors écran, sur un document fabriqué.
/// </summary>
public static class SelfTest
{
    static void Pump() => Application.Current.Dispatcher.Invoke(() => { }, DispatcherPriority.Background);

    static void Idle() => Application.Current.Dispatcher.Invoke(() => { }, DispatcherPriority.ApplicationIdle);

    /// <summary>
    /// Autorise une fenêtre plus grande que l'écran (Windows la limite à la taille de l'écran : 1024 x 768 sur les
    /// machines de GitHub) puis dimensionne la fenêtre pour que son contenu fasse exactement <paramref name="size"/>.
    /// </summary>
    static void ForceContentSize(Window w, Size size)
    {
        var hwnd = new System.Windows.Interop.WindowInteropHelper(w).EnsureHandle();
        if (Hooked.Add(hwnd))
            System.Windows.Interop.HwndSource.FromHwnd(hwnd).AddHook(AllowLargeWindow);
        var content = (FrameworkElement)w.Content;
        for (int i = 0; i < 3; i++)
        {
            Pump();
            double chromeW = w.ActualWidth - content.ActualWidth, chromeH = w.ActualHeight - content.ActualHeight;
            w.Width = size.Width + Math.Max(0, chromeW);
            w.Height = size.Height + Math.Max(0, chromeH);
            w.UpdateLayout();
        }
        Pump();
        if (content.ActualWidth < size.Width - 1 || content.ActualHeight < size.Height - 1)
            Console.Error.WriteLine($"Avertissement : capture limitée à {content.ActualWidth:0}x{content.ActualHeight:0} par la taille de l'écran " +
                                    $"(demandé : {size.Width:0}x{size.Height:0}).");
    }

    static readonly HashSet<IntPtr> Hooked = new();

    [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
    struct MinMaxInfo { public int rx, ry, maxSizeX, maxSizeY, maxPosX, maxPosY, minTrackX, minTrackY, maxTrackX, maxTrackY; }

    static IntPtr AllowLargeWindow(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        const int WM_GETMINMAXINFO = 0x0024;
        if (msg == WM_GETMINMAXINFO)
        {
            var mmi = System.Runtime.InteropServices.Marshal.PtrToStructure<MinMaxInfo>(lParam);
            mmi.maxTrackX = mmi.maxSizeX = 6000;
            mmi.maxTrackY = mmi.maxSizeY = 4000;
            System.Runtime.InteropServices.Marshal.StructureToPtr(mmi, lParam, false);
            handled = true;
        }
        return IntPtr.Zero;
    }

    static void Save(Visual v, int w, int h, string path)
    {
        var rtb = new RenderTargetBitmap(w, h, 96, 96, PixelFormats.Pbgra32);
        rtb.Render(v);
        var enc = new PngBitmapEncoder();
        enc.Frames.Add(BitmapFrame.Create(rtb));
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        using var fs = File.Create(path);
        enc.Save(fs);
    }

    // ================================================================ auto-test (versions 0.2 à 0.4)

    static int _ok, _ko;

    static void Check(bool cond, string what)
    {
        if (cond) _ok++; else _ko++;
        Console.WriteLine($"  {(cond ? "OK " : "ÉCHEC")} {what}");
    }

    /// <summary>Document de test : un titre, un paragraphe en gras, une case à cocher de formulaire, un tableau.</summary>
    static void WriteTestDocument(string path)
    {
        const string W = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";
        var doc = $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:document {W}><w:body>" +
            "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val=\"36\"/></w:rPr><w:t>Formulaire d'essai</w:t></w:r></w:p>" +
            "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"><w:ffData><w:name w:val=\"Accord\"/><w:enabled/><w:checkBox><w:sizeAuto/><w:default w:val=\"0\"/></w:checkBox></w:ffData></w:fldChar></w:r>" +
            "<w:r><w:instrText xml:space=\"preserve\"> FORMCHECKBOX </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r>" +
            "<w:r><w:t xml:space=\"preserve\"> J'accepte les conditions.</w:t></w:r></w:p>" +
            "<w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
            "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1417\" w:right=\"1417\" w:bottom=\"1417\" w:left=\"1417\"/></w:sectPr></w:body></w:document>";
        using var fs = File.Create(path);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Create);
        void Add(string name, string content)
        {
            using var sw = new StreamWriter(zip.CreateEntry(name).Open(), new System.Text.UTF8Encoding(false));
            sw.Write(content);
        }
        Add("[Content_Types].xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>");
        Add("_rels/.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
        Add("word/document.xml", doc);
    }

    public static int Run(string outDir)
    {
        Directory.CreateDirectory(outDir);
        var file = Path.Combine(outDir, "auto-test.docx");
        foreach (var f in new[] { file, file + ".bak" }) if (File.Exists(f)) File.Delete(f);
        WriteTestDocument(file);
        var original = File.ReadAllBytes(file);
        Console.WriteLine("Auto-test OpenWrite " + Update.UpdateService.CurrentVersion);
        var w = new MainWindow { Left = -6000, Top = 0, Width = 1402, Height = 948, ShowActivated = false, WindowStartupLocation = WindowStartupLocation.Manual, SkipConfirm = true };
        w.Show();
        w.OpenFile(file);
        Pump();
        var body = w.Body;
        Check(body is not null, "document ouvert dans le modèle");
        if (body is null) { w.Close(); return 1; }
        Field Box() => FieldIndex.Build(w.Body!).All.Single(f => f.CheckBox is not null);
        Paragraph BoxParagraph() => Box().BeginParagraph!;
        Check(!w.IsModified && !w.Title.StartsWith("•"), "document non modifié à l'ouverture");
        Check(Box().CheckBox!.Checked == false, "case décochée à l'ouverture");

        // vue des pages (0.3) : le document est paginé, la case est cliquée là où elle est dessinée
        var layout = w.CurrentLayout;
        Check(layout is { Pages.Count: 1 }, $"vue des pages : {layout?.Pages.Count ?? 0} page(s), mise en page en {w.LastLayoutMs:0} ms");
        (int, Point)? boxAt = null;
        if (layout is not null)
            foreach (var l in layout.Pages[0].Lines)
                foreach (var (pos, _) in l.Paragraph.Content.CheckBoxes)
                {
                    var box = l.Paragraph.Lines[l.Line];
                    if (pos < box.Start || pos >= box.End) continue;
                    var piece = box.Pieces.First(p => pos >= p.Start && pos < p.Start + p.Length);
                    var st = l.Paragraph.Content.Styles[piece.Style];
                    double x = l.X + piece.X;
                    for (int k = piece.Start; k < pos; k++) x += st.Face.Advance(l.Paragraph.Content.Text[k]) * st.Size;
                    boxAt = (0, new Point(x + 2, l.Y + box.Height / 2));
                }
        Check(boxAt is not null, "case dessinée sur la page");
        if (boxAt is { } at) w.ClickAt(at); else w.ToggleCheckBox(BoxParagraph(), Box());
        Pump();
        Check(Box().CheckBox!.Checked, "clic sur la page : case cochée");
        Check(w.IsModified && w.Title.StartsWith("•"), "titre : document modifié");
        Check(ApplicationCommands.Undo.CanExecute(null, w), "Annuler disponible");

        ApplicationCommands.Undo.Execute(null, w);
        Pump();
        Check(!Box().CheckBox!.Checked && !w.IsModified, "Ctrl+Z : case décochée, document revenu à l'état enregistré");
        ApplicationCommands.Redo.Execute(null, w);
        Pump();
        Check(Box().CheckBox!.Checked && w.IsModified, "Ctrl+Y : case recochée");

        Check(w.Save(saveAs: false), "Ctrl+S : enregistré");
        Check(!w.IsModified && !w.Title.StartsWith("•"), "titre : plus de modification en attente");
        var saved = DocxPackage.Open(file);
        var savedBox = FieldIndex.Build(saved.Main!.Stories().First()).All.Single(f => f.CheckBox is not null).CheckBox!;
        Check(savedBox.Checked && savedBox.Name == "Accord", "fichier relu : case cochée, nom conservé");
        var diffs = XmlCompare.Diff(DocxPackage.ParseXml(DocxPackage.Open(new MemoryStream(original)).Find("word/document.xml")!.Data).Root!,
                                    DocxPackage.ParseXml(saved.Find("word/document.xml")!.Data).Root!);
        Check(diffs.Count == 1 && diffs[0].Contains("checked"), "fichier relu : seule la case a changé (" + string.Join(" ; ", diffs) + ")");
        Check(!AppSettings.Load().KeepBackup || File.Exists(file + ".bak"), "version précédente gardée (.bak)");

        // saisie (0.4) : touches et caractères envoyés à la vue des pages comme par le clavier
        Pump();
        Idle();
        var view = w.PagesViewForTest!;
        void Key(Key k)
        {
            view.RaiseEvent(new KeyEventArgs(Keyboard.PrimaryDevice, PresentationSource.FromVisual(view)!, 0, k) { RoutedEvent = Keyboard.KeyDownEvent });
            Pump();
        }
        void Type(string text)
        {
            foreach (var ch in text)
                view.RaiseEvent(new TextCompositionEventArgs(Keyboard.PrimaryDevice, new TextComposition(InputManager.Current, view, ch.ToString()))
                    { RoutedEvent = TextCompositionManager.TextInputEvent });
            Pump();
        }
        static string TextOf(Block b) => b is Paragraph p ? string.Concat(p.Runs().SelectMany(r => r.Content).OfType<TextContent>().Select(t => t.Value)) : "";
        Check(w.CaretPoint is { Index: 0 } c0 && ReferenceEquals(c0.Paragraph, w.Body!.Blocks[1]), "Ctrl+Y : point d'insertion sur le paragraphe de la case");
        Key(System.Windows.Input.Key.Up);
        Key(System.Windows.Input.Key.End);
        Type(" complet");
        Check(TextOf(w.Body!.Blocks[0]) == "Formulaire d'essai complet", $"frappe en fin de ligne : « {TextOf(w.Body!.Blocks[0])} »");
        Key(System.Windows.Input.Key.Enter);
        Type("Deuxième ligne");
        Check(w.Body!.Blocks.Count == 4 && TextOf(w.Body!.Blocks[1]) == "Deuxième ligne", "Entrée : nouveau paragraphe, frappe dedans");
        Check(w.LastUpdateMs < 100, $"remise en page incrémentale : {w.LastUpdateMs:0.0} ms");
        Key(System.Windows.Input.Key.Home);
        Key(System.Windows.Input.Key.Back);
        Check(w.Body!.Blocks.Count == 3 && TextOf(w.Body!.Blocks[0]) == "Formulaire d'essai completDeuxième ligne", "Retour arrière en début de paragraphe : fusion");
        ApplicationCommands.Undo.Execute(null, w);
        Pump();
        Check(w.Body!.Blocks.Count == 4 && w.CaretPoint is { Index: 0 } cp && ReferenceEquals(cp.Paragraph, w.Body!.Blocks[1]), "Ctrl+Z : paragraphes séparés, point d'insertion revenu");
        // mise en forme : « Deuxième ligne » continue le titre en gras (Entrée en fin de texte gras, comme Word) ;
        // point d'insertion dans « Deuxième » → le mot perd le gras ; paragraphe aligné à droite
        for (int k = 0; k < 3; k++) Key(System.Windows.Input.Key.Right);
        w.ToggleCharacter("b");
        Pump();
        var runs2 = ((Paragraph)w.Body!.Blocks[1]).Runs().ToList();
        Check(runs2.Count == 2 && runs2[0].Text == "Deuxième" && !runs2[0].Props.Flag("b") && runs2[1].Props.Flag("b"), "Ctrl+G sur le mot du point d'insertion : gras retiré au mot seul");
        w.SetAlignment("right");
        Pump();
        Check(((Paragraph)w.Body!.Blocks[1]).Props.Val("jc") == "right" && w.CaretPoint is { Index: 3 }, "Aligner à droite, point d'insertion gardé");
        ApplicationCommands.Undo.Execute(null, w);
        Pump();
        Check(((Paragraph)w.Body!.Blocks[1]).Props.Val("jc") == "center", "Ctrl+Z : alignement d'avant");
        Check(w.Save(saveAs: false), "Ctrl+S après la frappe : enregistré");
        var reread = DocxPackage.Open(file).Main!.Stories().First();
        Check(TextOf(reread.Blocks[0]) == "Formulaire d'essai complet" && TextOf(reread.Blocks[1]) == "Deuxième ligne"
              && !((Paragraph)reread.Blocks[1]).Runs().First().Props.Flag("b")
              && FieldIndex.Build(reread).All.Single(f => f.CheckBox is not null).CheckBox!.Checked, "fichier relu : texte saisi, case toujours cochée");

        Idle();
        var content = (FrameworkElement)w.Content;
        Save(content, (int)content.ActualWidth, (int)content.ActualHeight, Path.Combine(outDir, "auto-test.png"));
        w.Close();
        Console.WriteLine($"{_ok} vérification(s) réussie(s), {_ko} échec(s)");
        return _ko == 0 ? 0 : 1;
    }

    /// <summary>
    /// Banc de saisie dans la vraie fenêtre (hors écran) : OpenWrite.exe --typing-bench document.docx [répétitions].
    /// Frappe au milieu du document, Entrée, Retour arrière, gras, centrage : durée de chaque touche, affichage compris.
/// Ne modifie pas le fichier ; --save copie.docx enregistre le résultat ailleurs (pour le faire ouvrir par Word).
    /// </summary>
    public static int TypingBench(string file, int repeat, string? saveCopy = null)
    {
        var w = new MainWindow { Left = -6000, Top = 0, Width = 1402, Height = 948, ShowActivated = false, WindowStartupLocation = WindowStartupLocation.Manual, SkipConfirm = true };
        w.Show();
        w.OpenFile(file);
        Pump();
        w.RepeatBodyForBench(repeat);
        Pump();
        Idle();
        var layout = w.CurrentLayout!;
        Console.WriteLine($"{Path.GetFileName(file)} ×{repeat} : {layout.Pages.Count} pages, mise en page complète {w.LastLayoutMs:0} ms");
        int page = layout.Pages.Count / 2;
        w.PlaceCaretOnPage(page);
        Pump();
        Idle();
        var view = w.PagesViewForTest!;
        var times = new List<double>();
        var updates = new List<double>();
        void Measure(Action a)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            a();
            Idle();                      // jusqu'à l'affichage compris
            times.Add(sw.Elapsed.TotalMilliseconds);
            updates.Add(w.LastUpdateMs);
        }
        void Key(Key k) => view.RaiseEvent(new KeyEventArgs(Keyboard.PrimaryDevice, PresentationSource.FromVisual(view)!, 0, k) { RoutedEvent = Keyboard.KeyDownEvent });
        void Char(char ch) => view.RaiseEvent(new TextCompositionEventArgs(Keyboard.PrimaryDevice, new TextComposition(InputManager.Current, view, ch.ToString()))
            { RoutedEvent = TextCompositionManager.TextInputEvent });
        foreach (var ch in "Essai de frappe au milieu du document, avec une phrase assez longue pour changer de ligne. ") Measure(() => Char(ch));
        Measure(() => Key(System.Windows.Input.Key.Enter));
        foreach (var ch in "Nouveau paragraphe.") Measure(() => Char(ch));
        for (int i = 0; i < 20; i++) Measure(() => Key(System.Windows.Input.Key.Back));
        for (int i = 0; i < 10; i++) Measure(() => Key(System.Windows.Input.Key.Down));
        // mise en forme du mot du point d'insertion, puis copie enregistrée (à faire ouvrir par Word)
        Measure(() => Key(System.Windows.Input.Key.Left));
        Measure(() => Key(System.Windows.Input.Key.Left));
        Measure(() => w.ToggleCharacter("b"));
        Measure(() => w.SetAlignment("center"));
        if (saveCopy is not null) w.SaveCopy(saveCopy);
        times.Sort();
        updates.Sort();
        Console.WriteLine($"{times.Count} touches : médiane {times[times.Count / 2]:0.0} ms, 95e centile {times[(int)(times.Count * 0.95)]:0.0} ms, max {times[^1]:0.0} ms " +
                          $"(remise en page seule : médiane {updates[updates.Count / 2]:0.00} ms, max {updates[^1]:0.00} ms) ; {layout.Pages.Count} → {w.CurrentLayout!.Pages.Count} pages");
        w.Close();
        return 0;
    }

    public static int Screenshot(string? file, string png, int width = 1402, int height = 948, string? scene = null)
    {
        var w = new MainWindow { Left = -6000, Top = 0, Width = width, Height = height, ShowActivated = false, WindowStartupLocation = WindowStartupLocation.Manual };
        // documents récents fictifs : jamais les vrais fichiers de l'utilisateur dans une capture
        w.RecentOverride = new[] { @"C:\Users\Moi\Documents\Club\compte-rendu-octobre.docx", @"C:\Users\Moi\Documents\Club\convocation.docx" };
        w.Show();
        if (file is not null) w.OpenFile(file);
        var target = new Size(width, height);
        ForceContentSize(w, target);
        switch (scene)
        {
            case "about":
            {
                var about = new AboutWindow(() => { }) { Owner = w, Left = -6000, Top = 0, ShowActivated = false, WindowStartupLocation = WindowStartupLocation.Manual };
                about.Show();
                Pump();
                Idle();
                ((FrameworkElement)about.Content).Margin = new Thickness(0);   // sans la marge réservée à l'ombre
                Pump();
                var c = (FrameworkElement)about.Content;
                Save(c, (int)c.ActualWidth, (int)c.ActualHeight, png);
                about.Close();
                w.Close();
                return 0;
            }
            case "file-menu":
            {
                typeof(MainWindow).GetMethod("FileButton_Click", BindingFlags.NonPublic | BindingFlags.Instance)!
                    .Invoke(w, new object[] { w, new RoutedEventArgs() });
                Pump();
                Idle();
                var menu = w.LastMenu!;
                Save(menu, (int)Math.Ceiling(menu.ActualWidth), (int)Math.Ceiling(menu.ActualHeight), png);
                menu.IsOpen = false;
                w.Close();
                return 0;
            }
            default:
                if (scene is not null && w.FindName(scene) is RadioButton rb) rb.IsChecked = true;
                // « selection » : quelques mots sélectionnés sur la première page
                if (scene == "selection")
                {
                    Pump();
                    w.SelectSampleWords(0);
                }
                // « page-N » : vue des pages défilée jusqu'à la page N
                if (scene is not null && scene.StartsWith("page-") && int.TryParse(scene[5..], out var n))
                {
                    Pump();
                    w.GoToPage(n - 1);
                }
                break;
        }
        Pump();
        Idle();
        var content = (FrameworkElement)w.Content;
        Save(content, (int)Math.Min(target.Width, content.ActualWidth), (int)Math.Min(target.Height, content.ActualHeight), png);
        w.Close();
        return 0;
    }
}
