using System.Xml.Linq;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;
using OpenWrite.Layout;

namespace OpenWrite.App;

// Lot 4 (version 0.5) : styles et listes. Appliquer un style (galerie, volet, Ctrl+Alt+1…), styles intégrés écrits dans
// styles.xml au premier usage ; puces, numéros, liste de titres ; niveaux (Tab / Maj+Tab), fin de liste (Entrée sur un
// élément vide, Retour arrière au début), recommencer / continuer la numérotation. Chaque opération est annulable, styles.xml
// et numbering.xml compris (instantanés XML du journal).
public partial class MainWindow
{
    XElement? StylesRoot(bool create = false) => _pkg?.XmlPartOf("styles", create)?.Doc.Root;
    XElement? NumberingRoot(bool create = false) => _pkg?.XmlPartOf("numbering", create)?.Doc.Root;
    internal XElement? StylesXmlForTest => StylesRoot();
    internal OpenWrite.Docx.DocxPackage? PackageForTest => _pkg;

    /// <summary>
    /// Transaction sur les blocs de la sélection (tableaux touchés compris) où l'opération peut aussi modifier styles.xml
    /// et numbering.xml (suivis par le journal) ; sélection gardée ensuite.
    /// </summary>
    bool EditSelection(string label, Action<EditTransaction, Story, DocPos, DocPos> op, bool lists = false, bool additive = false)
    {
        if (EditRoot is not { } body || Ordered() is not { } sel) return false;
        if (ToDocPos(sel.From) is not { } from || ToDocPos(sel.To) is not { } to) { NotEditable(); return false; }
        if (!SupportedRange(from, to)) return false;
        var (f, t) = FormatOps.Normalize(from, to);
        var (anchor, caret) = SelectionDocPos();
        try
        {
            var ranges = TargetRanges(from, to);
            using var tx = _journal.Begin(label, body, f.ContainerPath, f.BlockIndex, t.BlockIndex - f.BlockIndex + 1, caret);
            foreach (var (a, b) in ranges) op(tx, body, a, b);
            tx.Commit(caret);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or InvalidOperationException or FormatException)
        {
            StatusText.Text = $"{label} impossible ici : {ex.Message}";
            return false;
        }
        _listsTouched = lists;
        _xmlAdditive = additive;
        AfterFormat(label, anchor, caret);
        return true;
    }

    /// <summary>Paragraphes de la plage (tableaux traversés compris).</summary>
    static List<Paragraph> ParagraphsIn(Story body, DocPos from, DocPos to)
    {
        var list = new List<Paragraph>();
        FormatOps.ApplyToParagraphs(body, from, to, list.Add);
        return list;
    }

    // ================================================================ appliquer un style

    /// <summary>
    /// Applique un style : nom intégré de Word (« heading 1 ») ou identifiant du document. Style de paragraphe : aux
    /// paragraphes de la sélection ; style de caractère : au texte visé (sélection ou mot du point d'insertion).
    /// </summary>
    internal void ApplyStyle(string nameOrId)
    {
        if (_styles?.Sheet is not { } sheet) return;
        var existing = sheet.Style(nameOrId) ?? StyleOps.ByName(sheet, nameOrId);
        string? type = existing is not null ? sheet.TypeOf(existing) : StyleOps.BuiltInType(nameOrId);
        if (type is null) { StatusText.Text = $"Style « {nameOrId} » inconnu."; return; }
        var display = existing?.Attribute(sheet.W + "styleId")?.Value is { } eid ? StyleOps.DisplayName(sheet, eid)
                    : StyleOps.FrenchNames.TryGetValue(nameOrId, out var fr) ? fr : nameOrId;
        var label = $"Style {display}";
        if (type == "character")
        {
            if (CharacterTarget() is not { } target) { StatusText.Text = "Sélectionnez le texte auquel appliquer ce style de caractère."; return; }
            var (savedAnchor, savedCaret) = (_anchor, _caret);
            _anchor = target.From;
            _caret = target.To;
            EditSelection(label, (tx, body, from, to) =>
            {
                var id = ResolveStyleId(tx, nameOrId, existing);
                bool isDefault = StyleOps.ByName(_styles!.Sheet, "Default Paragraph Font")?.Attribute(sheet.W + "styleId")?.Value == id;
                FormatOps.ApplyToRuns(body, from, to, (_, r) => r.Props.SetVal("rStyle", isDefault ? null : id));
            }, additive: true);
            if (savedAnchor is null) { _anchor = null; _caret = savedCaret; UpdateCaret(scroll: false); }
            return;
        }
        EditSelection(label, (tx, body, from, to) =>
        {
            var id = ResolveStyleId(tx, nameOrId, existing);
            var normal = sheet.DefaultStyle("paragraph")?.Attribute(sheet.W + "styleId")?.Value;
            foreach (var p in ParagraphsIn(body, from, to)) p.Props.SetVal("pStyle", id == normal ? null : id);
        }, lists: true, additive: true);
    }

    /// <summary>Identifiant du style (style intégré ajouté à styles.xml s'il n'y est pas encore, dans la transaction).</summary>
    string ResolveStyleId(EditTransaction tx, string nameOrId, XElement? existing)
    {
        var sheet = _styles!.Sheet;
        if (existing?.Attribute(sheet.W + "styleId")?.Value is { } id) return id;
        var root = StylesRoot(create: true) ?? throw new InvalidOperationException("le document n'a pas de styles");
        tx.TrackXml(root);
        return StyleOps.Ensure(sheet, root, nameOrId) ?? throw new InvalidOperationException($"style « {nameOrId} » inconnu");
    }

    /// <summary>Styles du volet et de la galerie : rapides de Word puis styles qFormat du document, (identifiant ou nom intégré, nom affiché, type).</summary>
    internal List<(string Key, string Display, string Type)> StyleChoices(bool all = false)
    {
        var res = new List<(string, string, string)>();
        if (_styles?.Sheet is not { } sheet) return res;
        var w = sheet.W;
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var name in StyleOps.QuickStyles)
        {
            if (StyleOps.ByName(sheet, name) is { } st && st.Attribute(w + "styleId")?.Value is { } id)
            {
                if (seen.Add(id)) res.Add((id, StyleOps.DisplayName(sheet, id), sheet.TypeOf(st) ?? "paragraph"));
            }
            else if (StyleOps.BuiltInType(name) is { } type && seen.Add(name))
                res.Add((name, StyleOps.FrenchNames.TryGetValue(name, out var fr) ? fr : name, type));
        }
        foreach (var st in sheet.Styles?.Elements(w + "style") ?? Enumerable.Empty<XElement>())
        {
            var type = sheet.TypeOf(st);
            if (type is not ("paragraph" or "character") || st.Attribute(w + "styleId")?.Value is not { } id || seen.Contains(id)) continue;
            bool quick = st.Element(w + "qFormat") is not null && st.Element(w + "hidden") is null && st.Element(w + "semiHidden") is null;
            if (!quick && !all) continue;
            seen.Add(id);
            res.Add((id, StyleOps.DisplayName(sheet, id), type));
        }
        return res;
    }

    // ================================================================ listes

    /// <summary>Puces / numéros : retire la liste si toute la sélection en est déjà une de ce genre, sinon l'applique.</summary>
    internal void ToggleList(ListKind kind)
    {
        if (_styles is not { } styles || EditRoot is not { } body || Ordered() is not { } sel || ToDocPos(sel.From) is not { } f0 || ToDocPos(sel.To) is not { } t0) return;
        if (!SupportedRange(f0, t0)) return;
        var (nf, nt) = FormatOps.Normalize(f0, t0);
        var paras = ParagraphsIn(body, nf, nt);
        bool Is(Paragraph p) => styles.Paragraph(p) is { NumId: > 0 } pf && ListOps.IsBullet(styles.Sheet, pf.NumId) == (kind == ListKind.Bullet);
        bool remove = paras.Count > 0 && paras.All(Is);
        var label = remove ? "Retirer la liste" : kind == ListKind.Bullet ? "Puces" : "Numérotation";
        // liste à continuer : celle du paragraphe qui précède, si elle est du même genre
        int? previous = null;
        var all = body.Paragraphs().ToList();
        int first = paras.Count > 0 ? all.IndexOf(paras[0]) : -1;
        if (first > 0 && styles.Paragraph(all[first - 1]) is { NumId: > 0 } prev && ListOps.IsBullet(styles.Sheet, prev.NumId) == (kind == ListKind.Bullet)) previous = prev.NumId;
        EditSelection(label, (tx, b, from, to) =>
        {
            var sheet = _styles!.Sheet;
            var w = sheet.W;
            if (remove)
            {
                foreach (var p in ParagraphsIn(b, from, to))
                {
                    var bare = new Paragraph { Props = p.Props.Clone() };
                    bare.Props.Remove("numPr");
                    ListOps.Remove(p, styleHasList: _styles.Paragraph(bare).NumId > 0);
                    // « Paragraphe de liste » posé avec la liste : retiré avec elle
                    if (StyleOps.ByName(sheet, "List Paragraph")?.Attribute(w + "styleId")?.Value is { } lp && p.Props.Val("pStyle") == lp) p.Props.Remove("pStyle");
                }
                return;
            }
            var numbering = NumberingRoot(create: true) ?? throw new InvalidOperationException("numérotation impossible");
            tx.TrackXml(numbering);
            int numId = previous ?? ListOps.Create(numbering, w, kind);
            var normal = sheet.DefaultStyle("paragraph")?.Attribute(w + "styleId")?.Value;
            string? listParagraph = null;
            foreach (var p in ParagraphsIn(b, from, to))
            {
                int ilvl = _styles.Paragraph(p) is { NumId: > 0 } cur ? cur.ListLevel : 0;
                ListOps.Set(p, numId, ilvl);
                // comme Word : un paragraphe Normal mis en liste prend le style « Paragraphe de liste »
                if (p.Props.Val("pStyle") is null || p.Props.Val("pStyle") == normal)
                {
                    if (listParagraph is null)
                    {
                        if (StylesRoot(create: true) is { } sr) { tx.TrackXml(sr); listParagraph = StyleOps.Ensure(sheet, sr, "List Paragraph"); }
                    }
                    if (listParagraph is not null) p.Props.SetVal("pStyle", listParagraph);
                }
            }
        }, lists: true, additive: true);
    }

    /// <summary>Paragraphe du point d'insertion dans une liste ?</summary>
    bool CaretInList() => _caret is { } c && LayoutOf(c.Paragraph) is { } pl && pl.Format.NumId > 0;

    /// <summary>Tab / Maj+Tab au début d'un élément de liste : niveau suivant / précédent (0 à 8).</summary>
    internal void ChangeListLevel(int delta)
    {
        if (_styles is not { } styles) return;
        EditSelection(delta > 0 ? "Abaisser d'un niveau" : "Remonter d'un niveau", (_, body, from, to) =>
        {
            foreach (var p in ParagraphsIn(body, from, to))
                if (styles.Paragraph(p) is { NumId: > 0 } pf)
                    ListOps.Set(p, pf.NumId, Math.Clamp(pf.ListLevel + delta, 0, 8));
        }, lists: true);
    }

    /// <summary>Entrée sur un élément vide ou Retour arrière au début : le paragraphe sort de la liste (comme Word).</summary>
    internal void EndList()
    {
        if (_styles is not { } styles) return;
        EditSelection("Fin de la liste", (_, body, from, to) =>
        {
            foreach (var p in ParagraphsIn(body, from, to))
            {
                var bare = new Paragraph { Props = p.Props.Clone() };
                bare.Props.Remove("numPr");
                ListOps.Remove(p, styleHasList: styles.Paragraph(bare).NumId > 0);
            }
        }, lists: true);
    }

    /// <summary>Recommencer à 1 : nouvelle instance de la liste pour ce paragraphe et les suivants de la même liste.</summary>
    internal void RestartNumbering()
    {
        if (_styles is not { } styles || _caret is not { } c || styles.Paragraph(c.Paragraph) is not { NumId: > 0 } pf) return;
        RenumberFrom(c.Paragraph, pf.NumId, "Recommencer la numérotation", (root, w) => ListOps.Restart(root, w, pf.NumId, pf.ListLevel));
    }

    /// <summary>Continuer la numérotation : reprend la liste précédente de même définition (ou de même genre).</summary>
    internal void ContinueNumbering()
    {
        if (_styles is not { } styles || Body is not { } body || _caret is not { } c || styles.Paragraph(c.Paragraph) is not { NumId: > 0 } pf) return;
        var all = body.Paragraphs().ToList();
        int k = all.IndexOf(c.Paragraph);
        var w = styles.Sheet.W;
        var root = NumberingRoot();
        if (root is null) return;
        var abs = ListOps.AbstractOf(root, w, pf.NumId);
        bool bullet = ListOps.IsBullet(styles.Sheet, pf.NumId);
        int? target = null;
        for (int i = k - 1; i >= 0 && target is null; i--)
            if (styles.Paragraph(all[i]) is { NumId: > 0 } q && q.NumId != pf.NumId
                && (ListOps.AbstractOf(root, w, q.NumId) == abs || ListOps.IsBullet(styles.Sheet, q.NumId) == bullet)) target = q.NumId;
        if (target is null) { StatusText.Text = "Aucune liste précédente à continuer."; return; }
        RenumberFrom(c.Paragraph, pf.NumId, "Continuer la numérotation", (_, _) => target.Value);
    }

    void RenumberFrom(Paragraph start, int oldNumId, string label, Func<XElement, XNamespace, int> newNum)
    {
        if (_styles is not { } styles || Body is not { } body || NumberingRoot() is not { } root) return;
        // tout le corps : les éléments suivants de la liste peuvent être loin
        var caret = _caret is { } c ? ToDocPos(c) : null;
        using (var tx = _journal.Begin(label, body, Array.Empty<int>(), 0, body.Blocks.Count, caret))
        {
            tx.TrackXml(root);
            int numId = newNum(root, styles.Sheet.W);
            bool started = false;
            foreach (var p in body.Paragraphs())
            {
                if (ReferenceEquals(p, start)) started = true;
                if (started && styles.Paragraph(p) is { NumId: var n, ListLevel: var l } && n == oldNumId) ListOps.Set(p, numId, l);
            }
            tx.Commit(caret);
        }
        _listsTouched = true;
        AfterFormat(label, null, caret);
    }

    /// <summary>Liste à plusieurs niveaux liée aux titres : Titre 1 → « 1. », Titre 2 → « 1.1 », Titre 3 → « 1.1.1 ».</summary>
    internal void NumberHeadings()
    {
        if (_styles?.Sheet is not { } sheet || Body is not { } body) return;
        var caret = _caret is { } c ? ToDocPos(c) : null;
        try
        {
            using var tx = _journal.Begin("Numéroter les titres", body, Array.Empty<int>(), 0, 0, caret);
            var styles = StylesRoot(create: true)!;
            var numbering = NumberingRoot(create: true)!;
            tx.TrackXml(styles);
            tx.TrackXml(numbering);
            var ids = new[] { "heading 1", "heading 2", "heading 3" }.Select(n => StyleOps.Ensure(sheet, styles, n)!).ToList();
            var fresh = new StyleSheet(styles, numbering, sheet.Theme, sheet.Settings, sheet.W);
            int numId = ListOps.Create(numbering, sheet.W, ListKind.Headings, ids);
            for (int l = 0; l < ids.Count; l++) ListOps.LinkStyle(fresh, fresh.Style(ids[l])!, numId, l);
            tx.Commit(caret);
        }
        catch (Exception ex) when (ex is InvalidOperationException or NullReferenceException)
        {
            StatusText.Text = $"Numérotation des titres impossible : {ex.Message}";
            return;
        }
        AfterFormat("Titres numérotés (Titre 1, 2, 3)", null, caret);
    }

    // ================================================================ styles : boîtes, modifier, créer

    /// <summary>Style de paragraphe du point d'insertion (identifiant ; Normal si aucun).</summary>
    string? CurrentParagraphStyle() =>
        _caret is { } c && LayoutOf(c.Paragraph) is { } pl ? pl.Format.StyleId ?? _styles?.Sheet.DefaultStyle("paragraph")?.Attribute(_styles.Sheet.W + "styleId")?.Value : null;

    internal void ShowStylesWindow()
    {
        if (_layout is null) return;
        var dlg = new StylesWindow(StyleChoices(all: true), CurrentParagraphStyle()) { Owner = this };
        if (dlg.ShowDialog() == true)
        {
            switch (dlg.Action)
            {
                case "apply" when dlg.Chosen is { } k: ApplyStyle(k); break;
                case "modify" when dlg.Chosen is { } k: ModifyStyle(k); break;
                case "update" when dlg.Chosen is { } k: UpdateStyleFromSelection(k); break;
                case "new": NewStyleFromSelection(); break;
            }
        }
        BackToPage();
    }

    internal void ModifyCurrentStyle()
    {
        if (CurrentParagraphStyle() is { } id) ModifyStyle(id);
    }

    /// <summary>Mise en forme résolue d'un style (paragraphe : la sienne et celle de sa marque ; caractère : celle d'un segment).</summary>
    (ParagraphFormat Para, RunFormat Font) StyleFormats(string id, string type)
    {
        var w = _styles!.Sheet.W;
        var probe = new Paragraph { Props = new PropertyBag("pPr", w) };
        if (type == "paragraph") probe.Props.SetVal("pStyle", id);
        var run = new Run { Props = new PropertyBag("rPr", w) };
        if (type == "character") run.Props.SetVal("rStyle", id);
        return (_styles.Paragraph(probe), type == "paragraph" ? _styles.ParagraphMark(probe) : _styles.Run(probe, run));
    }

    /// <summary>Modifier un style (nom intégré ou identifiant) : la boîte, puis styles.xml en une opération annulable.</summary>
    internal void ModifyStyle(string key)
    {
        if (_styles?.Sheet is not { } sheet || Body is not { } body) return;
        var existing = sheet.Style(key) ?? StyleOps.ByName(sheet, key);
        var type = existing is not null ? sheet.TypeOf(existing) ?? "paragraph" : StyleOps.BuiltInType(key) ?? "paragraph";
        var w = sheet.W;
        string? id = existing?.Attribute(w + "styleId")?.Value;
        var display = id is not null ? StyleOps.DisplayName(sheet, id) : StyleOps.FrenchNames.TryGetValue(key, out var fr) ? fr : key;
        var builtIn = existing?.Element(w + "name")?.Attribute(w + "val")?.Value is { } nm && StyleOps.FrenchNames.ContainsKey(nm) || existing is null;
        var (pf, rf) = id is not null ? StyleFormats(id, type) : (_styles.Paragraph(new Paragraph()), _styles.ParagraphMark(new Paragraph()));
        var paraChoices = StyleChoices(all: true).Where(c => c.Type == type && c.Key != id && sheet.Style(c.Key) is not null).Select(c => (c.Key, c.Display)).ToList();
        var dlg = new StyleWindow($"Modifier le style « {display} »", display, !builtIn, paraChoices,
            existing?.Element(w + "basedOn")?.Attribute(w + "val")?.Value, existing?.Element(w + "next")?.Attribute(w + "val")?.Value, rf, pf) { Owner = this };
        if (dlg.ShowDialog() != true) { BackToPage(); return; }
        var caret = _caret is { } c ? ToDocPos(c) : null;
        var (anchorPos, caretPos) = SelectionDocPos();
        using (var tx = _journal.Begin($"Modifier le style {display}", body, Array.Empty<int>(), 0, 0, caret))
        {
            var root = StylesRoot(create: true)!;
            tx.TrackXml(root);
            var styleId = id ?? StyleOps.Ensure(sheet, root, key)!;
            var style = root.Elements(w + "style").First(s => s.Attribute(w + "styleId")?.Value == styleId);
            if (!builtIn && dlg.StyleName != display) StyleOps.SetStyleValue(style, w, "name", dlg.StyleName);
            if (dlg.BasedOnChanged) StyleOps.SetStyleValue(style, w, "basedOn", dlg.BasedOn);
            if (dlg.NextChanged) StyleOps.SetStyleValue(style, w, "next", dlg.Next);
            if (dlg.Apply is { } apply)
            {
                var (para, run) = StyleOps.Editable(sheet, style);
                apply(para, run);
                StyleOps.Tidy(sheet, style);
            }
            tx.Commit(caret);
        }
        AfterFormat($"Style {display} modifié", anchorPos, caretPos);
        UpdateStyleGallery();
        BackToPage();
    }

    /// <summary>« Mettre à jour pour correspondre à la sélection » : la mise en forme directe du paragraphe passe dans son style.</summary>
    internal void UpdateStyleFromSelection(string key)
    {
        if (_styles?.Sheet is not { } sheet || _caret is not { } c) return;
        var w = sheet.W;
        EditSelection("Mettre à jour le style", (tx, body, from, to) =>
        {
            var root = StylesRoot(create: true)!;
            tx.TrackXml(root);
            var id = (sheet.Style(key) ?? StyleOps.ByName(sheet, key))?.Attribute(w + "styleId")?.Value ?? StyleOps.Ensure(sheet, root, key)!;
            var style = root.Elements(w + "style").First(s => s.Attribute(w + "styleId")?.Value == id);
            var p = c.Paragraph;
            var (ppr, rpr) = StyleOps.Capture(p);
            StyleOps.Merge(sheet, style, ppr, rpr);
            p.Props.SetVal("pStyle", id);
            StyleOps.Absorb(p, ppr, rpr);
        });
    }

    /// <summary>Nouveau style de paragraphe à partir de la sélection (nom demandé), appliqué au paragraphe.</summary>
    internal void NewStyleFromSelection(string? name = null)
    {
        if (_styles?.Sheet is not { } sheet || _caret is not { } c || LayoutOf(c.Paragraph) is not { } pl) return;
        var w = sheet.W;
        var current = CurrentParagraphStyle();
        if (name is null)
        {
            var dlg = new StyleWindow("Nouveau style", "Style1", true, StyleChoices(all: true).Where(x => x.Type == "paragraph" && sheet.Style(x.Key) is not null)
                .Select(x => (x.Key, x.Display)).ToList(), current, null, CaretFormat() ?? pl.Content.Mark.Format, pl.Format) { Owner = this };
            if (dlg.ShowDialog() != true) { BackToPage(); return; }
            name = dlg.StyleName;
        }
        EditSelection($"Nouveau style {name}", (tx, body, from, to) =>
        {
            var root = StylesRoot(create: true)!;
            tx.TrackXml(root);
            var p = c.Paragraph;
            var (ppr, rpr) = StyleOps.Capture(p);
            var id = StyleOps.Create(sheet, root, "paragraph", name, current, null, ppr, rpr);
            p.Props.SetVal("pStyle", id);
            StyleOps.Absorb(p, ppr, rpr);
        }, additive: true);
        UpdateStyleGallery();
    }

    /// <summary>Volet : liste des styles (choisir = appliquer).</summary>
    void InitStylePane()
    {
        PaneStyle.IsEnabled = true;
        PaneStyle.IsEditable = false;
        PaneStyle.DropDownOpened += (_, _) =>
        {
            _refreshingPane = true;
            var current = CurrentParagraphStyle();
            PaneStyle.ItemsSource = StyleChoices(all: true).Select(x => new ComboBoxItemData(x.Key, (x.Type == "character" ? "a  " : "¶  ") + x.Display)).ToList();
            PaneStyle.SelectedItem = ((List<ComboBoxItemData>)PaneStyle.ItemsSource).FirstOrDefault(i => i.Key == current);
            _refreshingPane = false;
        };
        PaneStyle.SelectionChanged += (_, _) =>
        {
            if (_refreshingPane || !PaneStyle.IsDropDownOpen || PaneStyle.SelectedItem is not ComboBoxItemData d) return;
            ApplyStyle(d.Key);
            BackToPage();
        };
    }

    sealed record ComboBoxItemData(string Key, string Display)
    {
        public override string ToString() => Display;
    }

    // ================================================================ menu contextuel de la page

    /// <summary>Ouvre le menu contextuel de la page au point d'insertion, comme un clic droit (auto-test de l'interface).</summary>
    internal System.Windows.Controls.ContextMenu? OpenContextMenuForTest()
    {
        if (_pagesView is null) return null;
        var menu = BuildContextMenu();
        _pagesView.ContextMenu = menu;
        menu.PlacementTarget = _pagesView;
        menu.Placement = System.Windows.Controls.Primitives.PlacementMode.MousePoint;
        menu.IsOpen = true;
        return menu;
    }

    System.Windows.Controls.ContextMenu BuildContextMenu()
    {
        var menu = new System.Windows.Controls.ContextMenu();
        if (AddObjectContextItems(menu)) return menu;
        AddSpellingItems(menu);
        AddReviewItems(menu);
        void Item(string header, string? gesture, Action act, bool enabled = true)
        {
            var mi = new System.Windows.Controls.MenuItem { Header = header, InputGestureText = gesture ?? "", IsEnabled = enabled };
            mi.Click += (_, _) => { act(); _pagesView?.Focus(); };
            menu.Items.Add(mi);
        }
        Item("Couper", "Ctrl+X", Cut, HasSelection);
        Item("Copier", "Ctrl+C", Copy, HasSelection);
        Item("Coller", "Ctrl+V", () => Paste());
        Item("Coller le texte seul", "Ctrl+Maj+V", () => Paste(textOnly: true));
        menu.Items.Add(new System.Windows.Controls.Separator());
        Item("Police…", "Ctrl+D", ShowFontDialog);
        Item("Paragraphe…", null, ShowParagraphDialog);
        Item("Lien…", "Ctrl+K", ShowLinkDialog);
        if (CaretInList())
        {
            menu.Items.Add(new System.Windows.Controls.Separator());
            Item("Recommencer à 1", null, RestartNumbering);
            Item("Continuer la numérotation", null, ContinueNumbering);
            Item("Abaisser d'un niveau", "Tab", () => ChangeListLevel(1));
            Item("Remonter d'un niveau", "Maj+Tab", () => ChangeListLevel(-1));
        }
        AddTableContextItems(menu);
        menu.Items.Add(new System.Windows.Controls.Separator());
        Item("Styles…", null, ShowStylesWindow);
        return menu;
    }
}
