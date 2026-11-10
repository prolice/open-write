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
/// « selection » (quelques mots sélectionnés sur la première page), « marques » (idem, marques de mise en forme affichées),
/// boîtes « symbole », « police », « paragraphe », « en-tete » (édition de l'en-tête de la première page).
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
        // la vue est refaite quand les styles changent (remise en page complète) : toujours la vue courante
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
        // rechercher / remplacer (0.4)
        Check(w.FindNext("ligne", false, false) && w.SelectedText() == "ligne", "Rechercher : occurrence sélectionnée");
        Check(w.ReplaceAll("ligne", "LIGNE", true, true) == 1 && TextOf(w.Body!.Blocks[1]) == "Deuxième LIGNE", "Remplacer tout");
        ApplicationCommands.Undo.Execute(null, w);
        Pump();
        Check(TextOf(w.Body!.Blocks[1]) == "Deuxième ligne", "Ctrl+Z : remplacement annulé d'un coup");
        // presse-papiers (0.4) : copier dépose texte, RTF et HTML ; coller dans le document garde la mise en forme
        w.FindNext("Deuxième ligne", true, false);
        w.Copy();
        Check(Clipboard.ContainsText() && Clipboard.ContainsData(DataFormats.Rtf) && Clipboard.ContainsData(DataFormats.Html)
              && (Clipboard.GetData(DataFormats.Rtf) as string ?? "").Contains("\\b"), "Copier : texte, RTF (gras) et HTML");
        File.WriteAllText(Path.Combine(outDir, "copie.rtf"), Clipboard.GetData(DataFormats.Rtf) as string ?? "", System.Text.Encoding.ASCII);
        Key(System.Windows.Input.Key.End);
        Key(System.Windows.Input.Key.Enter);
        w.Paste();
        Pump();
        var pasted = (Paragraph)w.Body!.Blocks[2];
        Check(TextOf(pasted) == "Deuxième ligne" && pasted.Runs().Any(r => r.Text.Contains("ligne") && r.Props.Flag("b")), "Coller (interne) : texte et gras");
        // comme un autre programme : le presse-papiers ne contient plus que du RTF (on attend qu'il soit visible)
        for (int k = 0; k < 20; k++)
        {
            try
            {
                Clipboard.SetDataObject(new DataObject(DataFormats.Rtf, "{\\rtf1\\ansi {\\i Italique} venu d'ailleurs}"), true);
                if (Clipboard.ContainsData(DataFormats.Rtf) && !Clipboard.ContainsData("OpenWrite.Fragment")) break;
            }
            catch (System.Runtime.InteropServices.ExternalException) { }
            System.Threading.Thread.Sleep(50);
            Pump();
        }
        w.Paste();
        Pump();
        Check(TextOf(w.Body!.Blocks[2]).EndsWith("Italique venu d'ailleurs") && ((Paragraph)w.Body!.Blocks[2]).Runs().Any(r => r.Text == "Italique" && r.Props.Flag("i")),
              $"Coller du RTF d'un autre programme : italique gardé ({w.StatusText.Text} ; « {TextOf(w.Body!.Blocks[2])} »)");
        // comme Excel : HTML dont la balise <table> précède <!--StartFragment-->, mise en forme dans des classes CSS, plus
        // un RTF qui met le tableau à plat (ne doit pas être choisi)
        const string excelHtml = "Version:1.0\r\nStartHTML:0000000000\r\n<html><head><style><!--td {font-size:10.0pt;font-family:Arial;border:none;}"
            + " .xl70 {color:#00B050;font-weight:700;text-align:center;vertical-align:middle;border:1.0pt solid windowtext;}--></style></head><body>"
            + "<table border=0 cellpadding=0 cellspacing=0><!--StartFragment--><col width=80 style='width:60pt'><col width=120 style='width:90pt'>"
            + "<tr height=37 style='height:28.35pt'><td class=xl70>Un</td><td class=xl70 style='border-left:none'>Deux</td></tr>"
            + "<tr><td>3</td><td>4</td></tr><!--EndFragment--></table></body></html>";
        for (int k = 0; k < 20; k++)
        {
            try
            {
                var data = new DataObject();
                data.SetData(DataFormats.Html, excelHtml);
                data.SetData(DataFormats.Rtf, "{\\rtf1\\ansi Un\\par Deux\\par}");
                data.SetData(DataFormats.UnicodeText, "Un\tDeux\r\n3\t4\r\n");
                Clipboard.SetDataObject(data, true);
                if (Clipboard.ContainsData(DataFormats.Html) && !Clipboard.ContainsData("OpenWrite.Fragment")) break;
            }
            catch (System.Runtime.InteropServices.ExternalException) { }
            System.Threading.Thread.Sleep(50);
            Pump();
        }
        w.Paste();
        Pump();
        var excel = w.Body!.Blocks.OfType<Table>().FirstOrDefault(t => t.Items.OfType<TableRow>().FirstOrDefault()?.Items.OfType<TableCell>().FirstOrDefault() is { } c0
            && Core.Editing.TableOps.CellText(c0) == "Un");
        var green = excel?.Items.OfType<TableRow>().First().Items.OfType<TableCell>().First().Content.Paragraphs().First().Runs().FirstOrDefault();
        Check(excel is not null && excel.Grid?.Elements().Count() == 2 && excel.Grid.Elements().First().Attribute(excel.Props.Ns + "w")?.Value == "1200"
              && green?.Props.Val("color") == "00B050" && green.Props.Flag("b"),
              $"Coller un tableau d'Excel : vrai tableau, largeurs et mise en forme gardées ({w.StatusText.Text})");
        ApplicationCommands.Undo.Execute(null, w);
        Pump();
        Check(!w.Body!.Blocks.OfType<Table>().Any(t => ReferenceEquals(t, excel)) && w.Body!.Blocks.OfType<Table>().Count() == 1, "Ctrl+Z : tableau collé retiré");
        ApplicationCommands.Undo.Execute(null, w);
        ApplicationCommands.Undo.Execute(null, w);
        ApplicationCommands.Undo.Execute(null, w);
        Pump();
        Check(w.Body!.Blocks.Count == 4, "Ctrl+Z ×3 : collages et nouveau paragraphe annulés");
        // lien hypertexte (Ctrl+K) sur « conditions »
        w.FindNext("conditions", true, true);
        w.ApplyLink("exemple.be", "", remove: false);
        Pump();
        var linkPara = (Paragraph)w.Body!.Blocks[2];
        var link = linkPara.Inlines.OfType<InlineContainer>().FirstOrDefault(c => c.Name.LocalName == "hyperlink");
        Check(link is not null && string.Concat(link.Runs().Select(r => r.Text)) == "conditions" && TextOf(linkPara) == " J'accepte les conditions.",
              "Ctrl+K : lien posé sur le mot, texte inchangé");
        // règle : retrait gauche à 36 pt et taquet à 100 pt sur « Deuxième ligne »
        w.FindNext("Deuxième", true, false);
        w.RulerAction(Layout.RulerHandle.Left, -1, 36);
        w.RulerAction(Layout.RulerHandle.None, -1, 100);
        var ruled = (Paragraph)w.Body!.Blocks[1];
        Check(ruled.Props.Get("ind")?.Attributes().FirstOrDefault(a => a.Name.LocalName == "left")?.Value == "720"
              && ruled.Props.Get("tabs")?.Elements().Any(t => t.Attributes().Any(a => a.Name.LocalName == "pos" && a.Value == "2000")) == true,
              "Règle : retrait gauche et taquet posés");
        ApplicationCommands.Undo.Execute(null, w);
        ApplicationCommands.Undo.Execute(null, w);
        Pump();
        // sélection qui sort d'un tableau : de « conditions » jusqu'à « A1 », remplacée par la frappe ; le tableau part entier
        w.FindNext("conditions", true, true);
        var selStart = w.AnchorPoint!.Value;
        w.FindNext("A1", true, true);
        w.Select(selStart, w.CaretPoint!.Value);
        Type("X");
        Check(w.Body!.Blocks.OfType<Table>().Count() == 0 && TextOf(w.Body!.Blocks[2]) == " J'accepte les X",
              $"Frappe sur une sélection qui traverse le tableau : « {TextOf(w.Body!.Blocks[2])} »");
        ApplicationCommands.Undo.Execute(null, w);
        Pump();
        Check(w.Body!.Blocks.OfType<Table>().Count() == 1 && TextOf(w.Body!.Blocks[2]) == " J'accepte les conditions.", "Ctrl+Z : tableau revenu");
        // tableau : Tab passe à la cellule suivante (contenu sélectionné), Maj+Tab revient… sans modifier le document
        w.FindNext("A1", true, true);
        Key(System.Windows.Input.Key.Right);
        Key(System.Windows.Input.Key.Tab);
        Check(w.SelectedText() == "B1", $"Tab dans un tableau : cellule suivante sélectionnée (« {w.SelectedText()} »)");
        // tableaux (0.7) : Tab dans la dernière cellule, sélection de cellules, gras, fusion, style, colonne
        var tbl = w.Body!.Blocks.OfType<Table>().First();
        string Cell(int r, int c) => Core.Editing.TableOps.CellText(new Core.Editing.TableOps.Map(tbl).Cells[r][c].Cell);
        Key(System.Windows.Input.Key.Tab);
        Check(tbl.Items.OfType<TableRow>().Count() == 2, "Tab dans la dernière cellule : nouvelle ligne");
        Type("A2");
        Key(System.Windows.Input.Key.Tab);
        Type("B2");
        Check(Cell(1, 0) == "A2" && Cell(1, 1) == "B2", $"frappe dans la nouvelle ligne : « {Cell(1, 0)} » « {Cell(1, 1)} »");
        w.FindNext("A1", true, true);
        var cellFrom = w.AnchorPoint!.Value;
        w.FindNext("B2", true, true);
        w.Select(cellFrom, w.CaretPoint!.Value);
        Check(w.CellRangeForTest == (0, 0, 1, 1) && w.SelectedText() == "A1\tB1\r\nA2\tB2\r\n", "sélection d'une cellule à l'autre : rectangle de 2 × 2 cellules");
        w.ToggleCharacter("b");
        Check(tbl.Items.OfType<TableRow>().SelectMany(r => r.Items.OfType<TableCell>()).All(c => c.Content.Paragraphs().First().Runs().All(r => r.Props.Flag("b"))),
              "gras sur les quatre cellules sélectionnées");
        w.MergeCells();
        var merged = new Core.Editing.TableOps.Map(tbl);
        Check(merged.Cells[0].Count == 1 && merged.Cells[0][0].Span == 2 && merged.IsRestart(merged.Cells[0][0].Cell) && Cell(0, 0) == "A1 B1 A2 B2",
              $"fusion : une cellule de 2 × 2 (« {Cell(0, 0)} »)");
        ApplicationCommands.Undo.Execute(null, w);
        Pump();
        tbl = w.Body!.Blocks.OfType<Table>().First();          // annuler remet une copie du bloc
        Check(new Core.Editing.TableOps.Map(tbl).Cells[0].Count == 2 && Cell(1, 1) == "B2", "Ctrl+Z : cellules séparées");
        w.FindNext("A2", true, true);
        w.ApplyTableStyle("Grid Table 4 Accent 1");
        Pump();
        var tableStyle = tbl.Props.Val("tblStyle");
        Check(tableStyle is not null && w.StylesXmlForTest?.Elements().Any(e => e.Attributes().Any(a => a.Value == tableStyle)
              && e.Elements().Any(n => n.Name.LocalName == "name" && n.Attributes().First().Value == "Grid Table 4 Accent 1")) == true,
              "style de tableau Grille 4 - Accentuation 1 écrit dans styles.xml");
        var headerLine = w.CurrentLayout!.Pages[0].Lines.First(l => l.Paragraph.Content.Text == "A1");
        Check(headerLine.Paragraph.Content.Styles[0].Format.Color == "FFFFFF", "ligne d'en-tête du style : texte blanc");
        w.InsertColumns(right: true);
        tbl = w.Body!.Blocks.OfType<Table>().First();
        Check(new Core.Editing.TableOps.Map(tbl).Columns == 3, "colonne insérée à droite");
        // styles et listes (0.5) : Titre 1 sur le titre, puces sur « Deuxième ligne » et la case, niveau, fin de liste
        w.FindNext("Formulaire", true, false);
        w.ApplyStyle("heading 1");
        Pump();
        var sheet5 = new Core.Styles.StyleSheet(w.StylesXmlForTest, null);
        var h1 = Core.Styles.StyleOps.ByName(sheet5, "heading 1")?.Attribute(sheet5.W + "styleId")?.Value;
        Check(h1 is not null && ((Paragraph)w.Body!.Blocks[0]).StyleId == h1, "Titre 1 appliqué (style ajouté à styles.xml)");
        w.FindNext("Deuxième", true, false);
        var listStart = w.AnchorPoint!.Value;
        w.FindNext("conditions", true, true);
        w.Select(listStart, w.CaretPoint!.Value);
        w.ToggleList(Core.Styles.ListKind.Bullet);
        Pump();
        bool InList(int b) => ((Paragraph)w.Body!.Blocks[b]).Props.Get("numPr") is not null;
        Check(InList(1) && InList(2), "Puces sur deux paragraphes");
        w.FindNext("J'accepte", true, false);
        Key(System.Windows.Input.Key.Home);
        Key(System.Windows.Input.Key.Tab);
        Check(((Paragraph)w.Body!.Blocks[2]).Props.Get("numPr")?.Elements().FirstOrDefault(e => e.Name.LocalName == "ilvl")?.Attributes().First().Value == "1",
              "Tab au début d'un élément de liste : niveau 2");
        Key(System.Windows.Input.Key.End);
        Key(System.Windows.Input.Key.Enter);
        Key(System.Windows.Input.Key.Enter);
        Check(w.Body!.Blocks.Count == 5 && !InList(3), "Entrée sur un élément vide : fin de la liste");
        ApplicationCommands.Undo.Execute(null, w);
        ApplicationCommands.Undo.Execute(null, w);
        Pump();
        Check(w.Body!.Blocks.Count == 4 && InList(2), "Ctrl+Z : élément vide retiré");
        // en-tête (0.6) : créé par double-clic dans la marge du haut, frappe, numéro de page, sortie par Échap
        Check(w.EnterHeader(0, "header"), "Double-clic dans la marge : en-tête créé et édité");
        Type("Club ");
        w.InsertPageField("PAGE");
        Key(System.Windows.Input.Key.Escape);
        Pump();
        var hdrPkg = w.PackageForTest!;
        var hdrStory = hdrPkg.Stories.Values.FirstOrDefault(p => p.Kind == "header")?.Stories().First();
        Check(!w.InHeaderMode && hdrStory is not null && TextOf(hdrStory.Blocks[0]).StartsWith("Club ")
              && FieldIndex.Build(hdrStory).All.Any(f => f.Code.Trim() == "PAGE"), "En-tête : texte et champ PAGE, retour au document");
        // filigrane (dans les en-têtes, un seul Ctrl+Z), couleur de page
        bool HasWatermark() => w.PackageForTest!.Stories.Values.Where(p => p.Kind == "header").SelectMany(p => p.Stories()).SelectMany(st => st.Paragraphs())
            .SelectMany(p => p.Runs()).SelectMany(r => r.Content).OfType<OpaqueRunContent>().Any(o => o.Xml.Descendants().Any(e => e.Name.LocalName == "textpath"));
        w.SetWatermark("BROUILLON");
        Check(HasWatermark(), "Filigrane posé dans l'en-tête");
        ApplicationCommands.Undo.Execute(null, w);
        Pump();
        Check(!HasWatermark(), "Ctrl+Z : filigrane retiré d'un coup");
        ApplicationCommands.Redo.Execute(null, w);
        Pump();
        w.SetPageColor("FFF8E7");
        Check(HasWatermark() && w.PackageForTest!.Main!.Items.OfType<System.Xml.Linq.XElement>().Any(e => e.Name.LocalName == "background"), "Ctrl+Y : filigrane revenu ; couleur de page");
        // mise en page : paysage et deux colonnes, puis Ctrl+Z
        w.EditSections("Paysage", s => { Core.Editing.SectionOps.SetOrientation(s, true); Core.Editing.SectionOps.SetColumns(s, 2); });
        Pump();
        Check(w.CurrentLayout!.Pages[0].Geometry.PageWidth > w.CurrentLayout!.Pages[0].Geometry.PageHeight && w.CurrentLayout!.Pages[0].Geometry.ColumnCount == 2,
              "Mise en page : paysage, deux colonnes");
        ApplicationCommands.Undo.Execute(null, w);
        Pump();
        Check(w.CurrentLayout!.Pages[0].Geometry.PageWidth < w.CurrentLayout!.Pages[0].Geometry.PageHeight, "Ctrl+Z : portrait, une colonne");
        Check(w.Save(saveAs: false), "Ctrl+S après la frappe : enregistré");
        var reread = DocxPackage.Open(file).Main!.Stories().First();
        Check(TextOf(reread.Blocks[0]) == "Formulaire d'essai complet" && TextOf(reread.Blocks[1]) == "Deuxième ligne"
              && !((Paragraph)reread.Blocks[1]).Runs().First().Props.Flag("b")
              && FieldIndex.Build(reread).All.Single(f => f.CheckBox is not null).CheckBox!.Checked, "fichier relu : texte saisi, case toujours cochée");
        var rereadPkg = DocxPackage.Open(file);
        var rereadLink = ((Paragraph)rereadPkg.Main!.Stories().First().Blocks[2]).Inlines.OfType<InlineContainer>().FirstOrDefault(c => c.Name.LocalName == "hyperlink");
        var reHeader = rereadPkg.Stories.Values.FirstOrDefault(p => p.Kind == "header")?.Stories().First();
        Check(reHeader is not null && FieldIndex.Build(reHeader).All.Any(f => f.Code.Trim() == "PAGE")
              && rereadPkg.Main!.Stories().First().FinalSection?.Element?.Elements().Any(e => e.Name.LocalName == "headerReference") == true,
              "fichier relu : en-tête avec numéro de page, référencé par la section");
        Check(rereadPkg.Main!.Items.OfType<System.Xml.Linq.XElement>().Any(e => e.Name.LocalName == "background" && e.Attributes().Any(x => x.Value == "FFF8E7")),
              "fichier relu : couleur de page");
        var reStyles = rereadPkg.XmlPartOf("styles")?.Doc.Root;
        Check(reStyles is not null && rereadPkg.XmlPartOf("numbering") is not null
              && Core.Styles.StyleOps.ByName(new Core.Styles.StyleSheet(reStyles, null), "heading 1") is not null, "fichier relu : styles.xml et numbering.xml créés");
        Check(rereadLink is not null && rereadPkg.HyperlinkTarget(rereadPkg.Main!, rereadLink) == "https://exemple.be", "fichier relu : lien vers https://exemple.be");
        var reTable = rereadPkg.Main!.Stories().First().Blocks.OfType<Table>().First();
        Check(new Core.Editing.TableOps.Map(reTable).Columns == 3 && reTable.Items.OfType<TableRow>().Count() == 2 && reTable.Props.Val("tblStyle") is not null,
              "fichier relu : tableau de 2 lignes × 3 colonnes avec son style");

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
        var labelled = new List<(double Ms, string What)>();
        string what = "";
        void Measure(Action a)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            a();
            Idle();                      // jusqu'à l'affichage compris
            times.Add(sw.Elapsed.TotalMilliseconds);
            updates.Add(w.LastUpdateMs);
            labelled.Add((sw.Elapsed.TotalMilliseconds, what));
        }
        void Key(Key k)
        {
            var v = w.PagesViewForTest!;
            v.RaiseEvent(new KeyEventArgs(Keyboard.PrimaryDevice, PresentationSource.FromVisual(v)!, 0, k) { RoutedEvent = Keyboard.KeyDownEvent });
        }
        void Char(char ch)
        {
            var v = w.PagesViewForTest!;
            v.RaiseEvent(new TextCompositionEventArgs(Keyboard.PrimaryDevice, new TextComposition(InputManager.Current, v, ch.ToString()))
                { RoutedEvent = TextCompositionManager.TextInputEvent });
        }
        what = "frappe";
        foreach (var ch in "Essai de frappe au milieu du document, avec une phrase assez longue pour changer de ligne. ") Measure(() => Char(ch));
        what = "Entrée";
        Measure(() => Key(System.Windows.Input.Key.Enter));
        what = "frappe 2";
        foreach (var ch in "Nouveau paragraphe.") Measure(() => Char(ch));
        what = "Retour arrière";
        for (int i = 0; i < 20; i++) Measure(() => Key(System.Windows.Input.Key.Back));
        what = "Bas";
        for (int i = 0; i < 10; i++) Measure(() => Key(System.Windows.Input.Key.Down));
        what = "mise en forme";
        // mise en forme du mot du point d'insertion, puis copie enregistrée (à faire ouvrir par Word)
        Measure(() => Key(System.Windows.Input.Key.Left));
        Measure(() => Key(System.Windows.Input.Key.Left));
        Measure(() => w.ToggleCharacter("b"));
        Measure(() => w.SetAlignment("center"));
        // lot 4 : style intégré, liste numérotée, couleurs du thème (remise en page complète)
        what = "style Titre 2";
        Measure(() => w.ApplyStyle("heading 2"));
        what = "numérotation";
        Measure(() => w.ToggleList(Core.Styles.ListKind.Numbered));
        what = "couleurs du thème";
        Measure(() => w.SetThemeColors(3));
        // lot 5 : en-tête de la première page (créé s'il manque), frappe dedans, sortie ; colonnes de la section
        what = "entrer dans l'en-tête";
        Measure(() => w.EnterHeader(0, "header"));
        what = "frappe dans l'en-tête";
        foreach (var ch in "En-tête ") Measure(() => Char(ch));
        what = "numéro de page";
        Measure(() => w.InsertPageField("PAGE"));
        what = "sortir de l'en-tête";
        Measure(() => w.ExitHeader());
        what = "deux colonnes";
        Measure(() => w.EditSections("Deux colonnes", s => Core.Editing.SectionOps.SetColumns(s, 2)));
        // lot 6 : tableau inséré, frappe et Tab, style, fusion d'une ligne, ligne et colonne, fractionnement
        what = "insérer un tableau";
        Measure(() => w.InsertTable(3, 4));
        what = "frappe dans le tableau";
        foreach (var word in new[] { "Un", "Deux", "Trois", "Quatre", "Cinq" })
        {
            foreach (var ch in word) Measure(() => Char(ch));
            Measure(() => Key(System.Windows.Input.Key.Tab));
        }
        what = "style de tableau";
        Measure(() => w.ApplyTableStyle("Grid Table 4 Accent 1"));
        what = "fusion";
        Measure(() => w.SelectInTable("row"));
        Measure(() => w.MergeCells());
        what = "ligne et colonne";
        Measure(() => w.InsertRows(below: true));
        Measure(() => w.InsertColumns(right: false));
        what = "fractionner";
        Measure(() => w.SplitCells(2, 2));
        if (saveCopy is not null) w.SaveCopy(saveCopy);
        Console.WriteLine("plus lentes : " + string.Join(", ", labelled.OrderByDescending(x => x.Ms).Take(4).Select(x => $"{x.What} {x.Ms:0} ms")));
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
            case "symbole" or "police" or "paragraphe":
            {
                Window dlg = scene switch
                {
                    "symbole" => new SymbolWindow(_ => { }),
                    "police" => new FontWindow(new Core.Styles.RunFormat { FontAscii = "Calibri", Size = 11, Bold = true, Effective = new System.Xml.Linq.XElement("rPr") }),
                    _ => new ParagraphWindow(new Core.Styles.ParagraphFormat { Alignment = "both", IndentFirstLine = 709, SpacingAfter = 160, Line = 276, KeepNext = true, Effective = new System.Xml.Linq.XElement("pPr") }),
                };
                dlg.Owner = w;
                dlg.Left = -6000;
                dlg.ShowActivated = false;
                dlg.WindowStartupLocation = WindowStartupLocation.Manual;
                dlg.Show();
                Pump();
                Idle();
                var c = (FrameworkElement)dlg.Content;
                Save(c, (int)c.ActualWidth, (int)c.ActualHeight, png);
                dlg.Close();
                w.Close();
                return 0;
            }
            case "tableau-galerie" or "tableau-grille":
            {
                // fenêtres surgissantes des tableaux : galerie des styles, grille d'insertion (leur contenu seul)
                Pump();
                w.ShowTableTabForCapture(layout: false, selectCells: false);
                Pump();
                if (scene == "tableau-galerie") w.ShowTableStylesGallery(w);
                else w.ShowTableMenu(w);
                Pump();
                Idle();
                var child = (FrameworkElement)w.LastPopup!.Child;
                Save(child, (int)Math.Ceiling(child.ActualWidth), (int)Math.Ceiling(child.ActualHeight), png);
                w.LastPopup.IsOpen = false;
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
                if (scene == "arriere-plan")
                {
                    Pump();
                    w.SetWatermark("BROUILLON");
                    w.SetPageColor("FFF8E7");
                    w.EditSections("Bordure", s => Core.Editing.SectionOps.SetPageBorders(s, "double", 4), all: true);
                    w.EditSections("Numéros de ligne", s => Core.Editing.SectionOps.SetLineNumbers(s, "newPage"));
                    Pump();
                    w.GoToPage(1);
                }
                if (scene is "tableau-creation" or "tableau-disposition")
                {
                    Pump();
                    w.ShowTableTabForCapture(layout: scene == "tableau-disposition", selectCells: scene == "tableau-disposition");
                }
                if (scene == "en-tete")
                {
                    Pump();
                    w.EnterHeader(0, "header");
                }
                if (scene is "selection" or "marques")
                {
                    Pump();
                    w.SelectSampleWords(0);
                    if (scene == "marques") w.ToggleMarks();
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
