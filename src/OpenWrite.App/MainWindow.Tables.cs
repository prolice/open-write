using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Xml.Linq;
using OpenWrite.Core.Editing;
using OpenWrite.Core.Model;
using OpenWrite.Core.Styles;
using OpenWrite.Layout;

namespace OpenWrite.App;

// Lot 6 (version 0.7) : tableaux comme dans Word — Insertion ▸ Tableau (grille, boîte, texte → tableau), sélection
// de cellules (rectangle, comme Word), onglets contextuels « Création de tableau » (styles de la galerie et options de
// style, trame, bordures) et « Disposition » (sélectionner, insérer, supprimer, fusionner, fractionner, ajustement,
// distribuer, alignement, trier, formule, convertir en texte, propriétés), Tab dans la dernière cellule, bordures de
// colonnes et de rangées tirées à la souris, quadrillage. Chaque opération est une seule étape annulable.
public partial class MainWindow
{
    /// <summary>Tableau du point d'insertion : tableau, chemin du bloc, grille, rangée et colonne de la cellule.</summary>
    sealed record TableCtx(Table Table, int[] Path, TableOps.Map Map, int Row, int Col, TableCell Cell);

    /// <summary>Sélection de cellules (rectangle de la grille, fusions comprises).</summary>
    sealed record CellSel(Table Table, int[] Path, int R0, int C0, int R1, int C1);

    TableCtx? TableAt(TextPoint? t)
    {
        if (CellOf(t) is not { } at) return null;
        if (at.Table.Items[at.Row] is not TableRow row || row.Items[at.Cell] is not TableCell cell) return null;
        var map = new TableOps.Map(at.Table);
        int r = map.Rows.IndexOf(row);
        return map.Find(cell) is { } f ? new TableCtx(at.Table, at.TablePath, map, r, f.Col, cell) : null;
    }

    /// <summary>
    /// Sélection qui va d'une cellule à une autre du même tableau (au niveau le plus profond où leurs cellules
    /// diffèrent) : rectangle de cellules, comme Word. Null pour une sélection de texte ordinaire.
    /// </summary>
    CellSel? CellSelection()
    {
        if (!HasSelection || _anchor is not { } a || _caret is not { } c || EditRoot is not { } root || ReferenceEquals(a.Paragraph, c.Paragraph)) return null;
        if (ToDocPos(a) is not { } pa || ToDocPos(c) is not { } pc) return null;
        var story = root;
        int i = 0;
        while (i < pa.Path.Count - 1 && i < pc.Path.Count - 1 && pa.Path[i] == pc.Path[i] && pa.Path[i] < story.Blocks.Count)
        {
            var block = story.Blocks[pa.Path[i]];
            if (block is Table t && i + 3 < pa.Path.Count && i + 3 < pc.Path.Count)
            {
                if (pa.Path[i + 1] != pc.Path[i + 1] || pa.Path[i + 2] != pc.Path[i + 2])
                {
                    var map = new TableOps.Map(t);
                    if (t.Items[pa.Path[i + 1]] is not TableRow ra || ra.Items[pa.Path[i + 2]] is not TableCell ca
                        || t.Items[pc.Path[i + 1]] is not TableRow rc || rc.Items[pc.Path[i + 2]] is not TableCell cc
                        || map.Find(ca) is not { } fa || map.Find(cc) is not { } fc) return null;
                    var (r0, c0, r1, c1) = map.Normalize(fa.Row, fa.Col, fc.Row, fc.Col);
                    var sa = map.At(fa.Row, fa.Col)!.Value;
                    var sc = map.At(fc.Row, fc.Col)!.Value;
                    (r0, c0, r1, c1) = map.Normalize(r0, Math.Min(c0, Math.Min(sa.Col, sc.Col)), r1, Math.Max(c1, Math.Max(sa.Col + sa.Span - 1, sc.Col + sc.Span - 1)));
                    return new CellSel(t, pa.Path.Take(i + 1).ToArray(), r0, c0, r1, c1);
                }
                if (t.Items[pa.Path[i + 1]] is not TableRow row || row.Items[pa.Path[i + 2]] is not TableCell cell) return null;
                story = cell.Content;
                i += 3;
                continue;
            }
            if (block is BlockContentControl sdt) { story = sdt.Content; i++; continue; }
            break;
        }
        return null;
    }

    /// <summary>Cellules (premières de leur fusion) d'une sélection de cellules.</summary>
    static List<TableCell> CellsOf(CellSel cs)
    {
        var map = new TableOps.Map(cs.Table);
        return map.In(cs.R0, cs.C0, cs.R1, cs.C1).Select(x => x.Slot.Cell).Where(c => !map.IsContinue(c)).Distinct().ToList();
    }

    /// <summary>
    /// Plages visées par une mise en forme : un paragraphe à la fois dans une sélection de cellules (rectangle), sinon la
    /// sélection elle-même.
    /// </summary>
    List<(DocPos From, DocPos To)> TargetRanges(DocPos from, DocPos to)
    {
        if (CellSelection() is not { } cs || EditRoot is not { } root) return new() { (from, to) };
        var res = new List<(DocPos, DocPos)>();
        foreach (var cell in CellsOf(cs))
            foreach (var p in cell.Content.Paragraphs())
            {
                var start = DocPos.Of(root, p);
                res.Add((start, start.WithOffset(ParagraphAtoms.Length(p))));
            }
        return res;
    }

    /// <summary>Rectangles des cellules sélectionnées sur une page (affichage de la sélection).</summary>
    IEnumerable<Rect> CellSelectionRects(int page, CellSel cs)
    {
        if (_layout is null || _hf is not null || page < 0 || page >= _layout.Pages.Count) yield break;
        var set = new HashSet<TableCell>(CellsOf(cs), ReferenceEqualityComparer.Instance);
        foreach (var b in _layout.Pages[page].Cells)
            if (ReferenceEquals(b.Table, cs.Table) && b.Cell is { } cell && set.Contains(cell)) yield return new Rect(b.X, b.Y, b.W, b.H);
    }

    /// <summary>Texte d'une sélection de cellules : cellules séparées par des tabulations, rangées par des retours (Word).</summary>
    static string CellSelectionText(CellSel cs)
    {
        var map = new TableOps.Map(cs.Table);
        var sb = new System.Text.StringBuilder();
        for (int r = cs.R0; r <= cs.R1; r++)
        {
            var cells = map.Cells[r].Where(s => s.Col + s.Span - 1 >= cs.C0 && s.Col <= cs.C1).Select(s => map.IsContinue(s.Cell) ? "" : TableOps.CellText(s.Cell));
            sb.Append(string.Join("\t", cells)).Append("\r\n");
        }
        return sb.ToString();
    }

    /// <summary>HTML d'une sélection de cellules (tableau simple : texte des cellules), pour les autres programmes.</summary>
    static string CellSelectionHtml(CellSel cs)
    {
        var map = new TableOps.Map(cs.Table);
        var sb = new System.Text.StringBuilder("<table border=\"1\" cellspacing=\"0\" cellpadding=\"4\" style=\"border-collapse:collapse\">");
        for (int r = cs.R0; r <= cs.R1; r++)
        {
            sb.Append("<tr>");
            foreach (var s in map.Cells[r].Where(s => s.Col + s.Span - 1 >= cs.C0 && s.Col <= cs.C1))
            {
                if (map.IsContinue(s.Cell)) continue;
                int rows = map.IsRestart(s.Cell) ? map.MergeBottom(r, s.Col) - r + 1 : 1;
                sb.Append("<td").Append(s.Span > 1 ? $" colspan=\"{s.Span}\"" : "").Append(rows > 1 ? $" rowspan=\"{rows}\"" : "").Append('>');
                sb.Append(string.Join("<br>", s.Cell.Content.Paragraphs().Select(p => System.Net.WebUtility.HtmlEncode(p.Text))));
                sb.Append("</td>");
            }
            sb.Append("</tr>");
        }
        return sb.Append("</table>").ToString();
    }

    // ================================================================ opérations

    /// <summary>
    /// Modifie le tableau du point d'insertion (ou de la sélection de cellules) en une étape annulable : l'opération reçoit
    /// le contexte et la sélection, et renvoie la nouvelle position du point d'insertion (null : sélection gardée).
    /// </summary>
    void EditTable(string label, Func<TableCtx, CellSel?, DocPos?> op, bool trackStyles = false)
    {
        if (EditRoot is not { } body) return;
        var cs = CellSelection();
        var ctx = cs is not null ? Ctx(cs) : TableAt(_caret);
        if (ctx is null) { StatusText.Text = "Placez le point d'insertion dans un tableau."; return; }
        var container = ctx.Path.Take(ctx.Path.Length - 1).ToArray();
        int index = ctx.Path[^1];
        var (savedAnchor, savedCaret) = (_anchor, _caret);
        DocPos? caret;
        try
        {
            using var tx = _journal.Begin(label, body, container, index, 1, _caret is { } c ? ToDocPos(c) : null);
            if (trackStyles && StylesRoot(create: true) is { } styles) tx.TrackXml(styles);
            caret = op(ctx, cs);
            tx.Commit(caret);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or InvalidOperationException or IndexOutOfRangeException)
        {
            StatusText.Text = $"{label} impossible : {ex.Message}";
            return;
        }
        if (trackStyles) _xmlAdditive = true;
        Relayout();
        if (caret is not null) { _anchor = null; _caret = FromDocPos(caret) ?? _caret; }
        else
        {
            // les paragraphes gardés sont les mêmes objets : sélection inchangée s'ils sont encore affichés
            _anchor = savedAnchor is { } a && LayoutOf(a.Paragraph) is not null ? a : null;
            _caret = savedCaret is { } c2 && LayoutOf(c2.Paragraph) is not null ? c2 : DocumentStart();
        }
        _goalX = null;
        Scroller.UpdateLayout();
        UpdateCaret();
        UpdateModifiedMarker();
        StatusText.Text = $"{label}. Ctrl+Z pour annuler.";
        EnsureAutoSaveTimer();
        CommandManager.InvalidateRequerySuggested();
        UpdateTableTabs();
    }

    TableCtx? Ctx(CellSel cs)
    {
        var map = new TableOps.Map(cs.Table);
        return map.At(cs.R0, cs.C0) is { } s ? new TableCtx(cs.Table, cs.Path, map, cs.R0, s.Col, s.Cell) : null;
    }

    /// <summary>Rectangle visé : la sélection de cellules, sinon la cellule du point d'insertion.</summary>
    static (int R0, int C0, int R1, int C1) Rect(TableCtx ctx, CellSel? cs)
    {
        if (cs is not null) return (cs.R0, cs.C0, cs.R1, cs.C1);
        var span = ctx.Map.At(ctx.Row, ctx.Col)?.Span ?? 1;
        return (ctx.Row, ctx.Col, ctx.Row, ctx.Col + span - 1);
    }

    /// <summary>Début de la cellule (rangée, colonne) d'un tableau.</summary>
    DocPos? CellStart(Table t, int row, int col)
    {
        if (EditRoot is not { } root) return null;
        var map = new TableOps.Map(t);
        row = Math.Clamp(row, 0, map.Rows.Count - 1);
        var slot = map.At(row, col) ?? (map.Cells[row].Count > 0 ? map.Cells[row][0] : null);
        if (slot is not { } s) return null;
        var cell = s.Cell;
        var top = map.MergeTop(row, s.Col);
        if (top != row && map.At(top, s.Col) is { } first) cell = first.Cell;
        return cell.Content.Paragraphs().FirstOrDefault() is { } p ? DocPos.Of(root, p) : null;
    }

    /// <summary>Largeur disponible pour le tableau (twips) : colonne de la section, pour un tableau du corps.</summary>
    int? AvailableTwips(TableCtx ctx)
    {
        if (ctx.Path.Length != 1 || _layout is null || _caret is not { } c || Locate(c) is not { } loc) return null;
        return (int)Math.Round(_layout.Pages[loc.Ref.Page].Geometry.ColumnWidth * 20);
    }

    // ---------------------------------------------------------------- insérer un tableau

    /// <summary>Insère un tableau « Grille du tableau » au point d'insertion (comme Insertion ▸ Tableau de Word).</summary>
    internal void InsertTable(int rows, int cols, string fit = "fixed")
    {
        if (EditRoot is not { } body || Ordered() is not { } sel || _styles is null || _pkg is null) return;
        if (ToDocPos(sel.From) is not { } at) { NotEditable(); return; }
        double width = LayoutOf(sel.From.Paragraph)?.Width ?? 450;
        var w = at.Paragraph(body).Props.Ns;
        DocPos pos;
        try
        {
            using var tx = _journal.Begin("Insérer un tableau", body, at.ContainerPath, at.BlockIndex, 1, at);
            var stylesRoot = StylesRoot(create: true);
            string? styleId = null;
            if (stylesRoot is not null)
            {
                tx.TrackXml(stylesRoot);
                styleId = TableStyleOps.Ensure(_styles.Sheet, stylesRoot, "Table Grid");
            }
            var table = TableOps.Create(w, rows, cols, width, styleId, null, NewParaId(), autoFit: fit == "content");
            if (fit == "window") TableOps.AutoFit(table, "window");
            pos = TableOps.InsertTable(body, at, table, NewParaId());
            tx.Commit(pos);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or InvalidOperationException)
        {
            StatusText.Text = $"Insertion du tableau impossible ici : {ex.Message}";
            return;
        }
        _xmlAdditive = true;
        AfterTyping(pos);
        StatusText.Text = $"Tableau de {cols} × {rows} inséré. Tab passe d'une cellule à l'autre ; Ctrl+Z pour annuler.";
        UpdateTableTabs();
    }

    internal void ShowInsertTableDialog()
    {
        var dlg = new InsertTableWindow { Owner = this };
        if (dlg.ShowDialog() == true) InsertTable(dlg.Rows, dlg.Columns, dlg.Fit);
        _pagesView?.Focus();
    }

    /// <summary>Menu Insertion ▸ Tableau : grille de 10 × 8 cases (survol : taille), boîte, conversion du texte.</summary>
    internal void ShowTableMenu(FrameworkElement anchor)
    {
        var popup = new Popup { PlacementTarget = anchor, Placement = PlacementMode.Bottom, StaysOpen = false, AllowsTransparency = true };
        var root = new StackPanel { Margin = new Thickness(8) };
        var title = new TextBlock { Text = "Insérer un tableau", FontWeight = FontWeights.SemiBold, Margin = new Thickness(2, 0, 0, 6) };
        root.Children.Add(title);
        var grid = new UniformGrid { Rows = 8, Columns = 10 };
        var cells = new Border[8, 10];
        for (int r = 0; r < 8; r++)
            for (int c = 0; c < 10; c++)
            {
                int rr = r, cc = c;
                var b = new Border { Width = 18, Height = 18, Margin = new Thickness(1), BorderBrush = R("Line"), BorderThickness = new Thickness(1), Background = Brushes.White, Cursor = Cursors.Hand };
                b.MouseEnter += (_, _) =>
                {
                    title.Text = $"Tableau de {cc + 1} × {rr + 1}";
                    for (int i = 0; i < 8; i++)
                        for (int j = 0; j < 10; j++)
                        {
                            bool on = i <= rr && j <= cc;
                            cells[i, j].Background = on ? R("SoftBadge") : Brushes.White;
                            cells[i, j].BorderBrush = on ? R("Blue") : R("Line");
                        }
                };
                b.MouseLeftButtonUp += (_, _) => { popup.IsOpen = false; InsertTable(rr + 1, cc + 1); _pagesView?.Focus(); };
                cells[r, c] = b;
                grid.Children.Add(b);
            }
        root.Children.Add(grid);
        Button Link(string text, Action act)
        {
            var b = new Button { Content = text, Margin = new Thickness(0, 6, 0, 0), HorizontalContentAlignment = HorizontalAlignment.Left, Padding = new Thickness(8, 4, 8, 4) };
            b.Click += (_, _) => { popup.IsOpen = false; act(); };
            return b;
        }
        root.Children.Add(Link("Insérer un tableau…", ShowInsertTableDialog));
        root.Children.Add(Link("Convertir le texte en tableau…", ConvertTextToTable));
        popup.Child = new Border { Background = R("Chrome"), BorderBrush = R("Line"), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(6), Child = root };
        popup.IsOpen = true;
        LastPopup = popup;
    }


    /// <summary>Dernière fenêtre surgissante ouverte (captures de la documentation).</summary>
    internal Popup? LastPopup { get; private set; }

    /// <summary>Paragraphes sélectionnés (même conteneur) → tableau, cellules coupées au séparateur choisi.</summary>
    internal void ConvertTextToTable()
    {
        if (EditRoot is not { } body || Ordered() is not { } sel || _styles is null) return;
        if (ToDocPos(sel.From) is not { } from || ToDocPos(sel.To) is not { } to) { NotEditable(); return; }
        if (!from.ContainerPath.SequenceEqual(to.ContainerPath)) { StatusText.Text = "Sélectionnez des paragraphes qui ne sont pas dans un tableau."; return; }
        var container = DocPos.ResolveContainer(body, from.ContainerPath);
        int first = from.BlockIndex, last = to.BlockIndex;
        if (to.Offset == 0 && last > first) last--;
        var paras = container.Blocks.GetRange(first, last - first + 1).OfType<Paragraph>().ToList();
        if (paras.Count != last - first + 1) { StatusText.Text = "La sélection contient autre chose que des paragraphes."; return; }
        bool tabs = paras.Any(p => p.Text.Contains('\t'));
        var dlg = new SeparatorWindow("Convertir le texte en tableau", $"{paras.Count} paragraphe(s) : une ligne de tableau par paragraphe.") { Owner = this };
        if (!tabs) dlg.Loaded += (_, _) => { };
        if (dlg.ShowDialog() != true) return;
        double width = LayoutOf(sel.From.Paragraph)?.Width ?? 450;
        DocPos pos;
        try
        {
            using var tx = _journal.Begin("Convertir le texte en tableau", body, from.ContainerPath, first, last - first + 1, from);
            var stylesRoot = StylesRoot(create: true);
            string? styleId = null;
            if (stylesRoot is not null) { tx.TrackXml(stylesRoot); styleId = TableStyleOps.Ensure(_styles.Sheet, stylesRoot, "Table Grid"); }
            // un saut de section porté par le dernier paragraphe reste sur un paragraphe vide après le tableau
            var sect = paras[^1].Props.Get("sectPr");
            var table = TableOps.FromParagraphs(paras, dlg.Separator[0], width, styleId, NewParaId());
            container.Blocks.RemoveRange(first, paras.Count);
            container.Blocks.Insert(first, table);
            if (sect is not null || first + 1 >= container.Blocks.Count || container.Blocks[first + 1] is not Paragraph)
            {
                var p = TableOps.EmptyParagraph(table.Props.Ns, null, NewParaId());
                if (sect is not null) p.Props.Set(new XElement(sect));
                container.Blocks.Insert(first + 1, p);
            }
            pos = new DocPos(from.ContainerPath.Concat(new[] { first, 0, 0, 0 }).ToArray(), 0);
            tx.Commit(pos);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or InvalidOperationException)
        {
            StatusText.Text = $"Conversion impossible : {ex.Message}";
            return;
        }
        _xmlAdditive = true;
        AfterTyping(pos);
        UpdateTableTabs();
    }

    // ---------------------------------------------------------------- rangées, colonnes, cellules

    internal void InsertRows(bool below) => EditTable(below ? "Insérer en dessous" : "Insérer au-dessus", (ctx, cs) =>
    {
        var (r0, _, r1, _) = Rect(ctx, cs);
        TableOps.InsertRows(ctx.Table, below ? r1 : r0, below, r1 - r0 + 1, NewParaId());
        return null;
    });

    internal void InsertColumns(bool right) => EditTable(right ? "Insérer à droite" : "Insérer à gauche", (ctx, cs) =>
    {
        var (_, c0, _, c1) = Rect(ctx, cs);
        TableOps.InsertColumns(ctx.Table, right ? c1 + 1 : c0, c1 - c0 + 1, NewParaId(), AvailableTwips(ctx));
        return null;
    });

    internal void DeleteRows() => EditTable("Supprimer les lignes", (ctx, cs) =>
    {
        var (r0, _, r1, _) = Rect(ctx, cs);
        if (TableOps.DeleteRows(ctx.Table, r0, r1)) return RemoveTable(ctx);
        return CellStart(ctx.Table, Math.Min(r0, new TableOps.Map(ctx.Table).Rows.Count - 1), 0);
    });

    internal void DeleteColumns() => EditTable("Supprimer les colonnes", (ctx, cs) =>
    {
        var (_, c0, _, c1) = Rect(ctx, cs);
        if (TableOps.DeleteColumns(ctx.Table, c0, c1)) return RemoveTable(ctx);
        return CellStart(ctx.Table, ctx.Row, Math.Max(0, c0 - 1));
    });

    internal void DeleteTable() => EditTable("Supprimer le tableau", (ctx, _) => RemoveTable(ctx));

    /// <summary>Retire le bloc du tableau ; point d'insertion au paragraphe qui le suivait.</summary>
    DocPos? RemoveTable(TableCtx ctx)
    {
        if (EditRoot is not { } root) return null;
        var container = DocPos.ResolveContainer(root, ctx.Path.Take(ctx.Path.Length - 1).ToArray());
        int index = ctx.Path[^1];
        container.Blocks.RemoveAt(index);
        if (container.Blocks.Count == 0 || index >= container.Blocks.Count || container.Blocks[index] is not Paragraph)
            container.Blocks.Insert(index, TableOps.EmptyParagraph(ctx.Table.Props.Ns, null, NewParaId()));
        return new DocPos(ctx.Path.Take(ctx.Path.Length - 1).Append(index).ToArray(), 0);
    }

    internal void MergeCells()
    {
        if (CellSelection() is null) { StatusText.Text = "Sélectionnez les cellules à fusionner (glisser d'une cellule à l'autre)."; return; }
        EditTable("Fusionner les cellules", (ctx, cs) =>
        {
            var (r0, c0, r1, c1) = Rect(ctx, cs);
            TableOps.Merge(ctx.Table, r0, c0, r1, c1, NewParaId());
            return CellStart(ctx.Table, r0, c0);
        });
    }

    /// <summary>Fractionner la cellule du point d'insertion (boîte, ou nombres donnés : auto-test, banc).</summary>
    internal void SplitCells(int cols = 0, int rows = 0)
    {
        if (TableAt(_caret) is null) { StatusText.Text = "Placez le point d'insertion dans la cellule à fractionner."; return; }
        if (cols <= 0)
        {
            var dlg = new SplitCellsWindow { Owner = this };
            if (dlg.ShowDialog() != true) return;
            (cols, rows) = (dlg.Columns, dlg.Rows);
        }
        _anchor = null;
        EditTable("Fractionner les cellules", (ctx, _) =>
        {
            TableOps.Split(ctx.Table, ctx.Row, ctx.Col, cols, Math.Max(1, rows), NewParaId());
            return CellStart(ctx.Table, ctx.Row, ctx.Col);
        });
    }

    /// <summary>Captures : point d'insertion dans le premier tableau du document, onglet de tableau affiché.</summary>
    internal void ShowTableTabForCapture(bool layout, bool selectCells)
    {
        if (Body?.Blocks.OfType<Table>().FirstOrDefault() is not { } t || t.Items.OfType<TableRow>().ElementAtOrDefault(1) is not { } row
            || row.Items.OfType<TableCell>().FirstOrDefault()?.Content.Paragraphs().FirstOrDefault() is not { } p) return;
        _anchor = null;
        _caret = new TextPoint(p, 0);
        if (selectCells && CellStart(t, 2, 1) is { } end && FromDocPos(end) is { } e)
        {
            _anchor = _caret;
            _caret = e;
        }
        UpdateCaret();
        UpdateTableTabs();
        if (layout) { if (_tableLayoutTab is not null) _tableLayoutTab.IsChecked = true; }
        else if (_tableDesignTab is not null) _tableDesignTab.IsChecked = true;
    }

    /// <summary>Sélection de cellules en cours (rangées et colonnes de la grille), pour l'auto-test.</summary>
    internal (int R0, int C0, int R1, int C1)? CellRangeForTest => CellSelection() is { } cs ? (cs.R0, cs.C0, cs.R1, cs.C1) : null;

    internal void SplitTableHere() => EditTable("Fractionner le tableau", (ctx, _) =>
    {
        if (EditRoot is not { } root) return null;
        if (ctx.Row == 0) throw new InvalidOperationException("placez le point d'insertion sous la première ligne");
        return TableOps.SplitTable(root, ctx.Path, ctx.Row, NewParaId());
    });

    // ---------------------------------------------------------------- sélectionner

    internal void SelectInTable(string what)
    {
        var cs = CellSelection();
        var ctx = cs is not null ? Ctx(cs) : TableAt(_caret);
        if (ctx is null || EditRoot is not { } root) return;
        var map = ctx.Map;
        var (r0, c0, r1, c1) = Rect(ctx, cs);
        (r0, c0, r1, c1) = what switch
        {
            "row" => (r0, 0, r1, map.Columns - 1),
            "column" => (0, c0, map.Rows.Count - 1, c1),
            "table" => (0, 0, map.Rows.Count - 1, map.Columns - 1),
            _ => (r0, c0, r1, c1),
        };
        (r0, c0, r1, c1) = map.Normalize(r0, c0, r1, c1);
        var first = map.In(r0, c0, r1, c1).First().Slot.Cell;
        var last = map.In(r0, c0, r1, c1).Last().Slot.Cell;
        var p0 = first.Content.Paragraphs().First();
        var p1 = last.Content.Paragraphs().Last();
        _anchor = new TextPoint(p0, 0);
        _caret = LayoutOf(p1) is { } pl ? new TextPoint(p1, pl.Content.Text.Length) : new TextPoint(p1, 0);
        _goalX = null;
        UpdateCaret();
    }

    // ---------------------------------------------------------------- taille

    internal void AutoFitTable(string mode) => EditTable(mode switch { "content" => "Ajuster au contenu", "window" => "Ajuster à la fenêtre", _ => "Largeur de colonne fixe" }, (ctx, _) =>
    {
        TableOps.AutoFit(ctx.Table, mode, CurrentGrid(ctx.Table));
        return null;
    });

    /// <summary>Largeurs des colonnes affichées (twips) : celles de la mise en page (ajustement automatique compris).</summary>
    List<int>? CurrentGrid(Table t)
    {
        if (_layout is null) return null;
        var boxes = _layout.Pages.SelectMany(p => p.Cells).Where(b => ReferenceEquals(b.Table, t) && b.Cell is not null).ToList();
        if (boxes.Count == 0) return null;
        var map = new TableOps.Map(t);
        var grid = map.GridWidths();
        foreach (var b in boxes)
            if (map.At(b.Row, b.Column) is { Span: 1 } s && s.Col == b.Column && b.Column < grid.Count) grid[b.Column] = (int)Math.Round(b.W * 20);
        return grid;
    }

    internal void DistributeColumns() => EditTable("Distribuer les colonnes", (ctx, cs) =>
    {
        var (_, c0, _, c1) = Rect(ctx, cs);
        if (cs is null) (c0, c1) = (0, ctx.Map.Columns - 1);
        TableOps.DistributeColumns(ctx.Table, c0, c1);
        return null;
    });

    internal void DistributeRows() => EditTable("Distribuer les lignes", (ctx, cs) =>
    {
        var (r0, _, r1, _) = Rect(ctx, cs);
        if (cs is null) (r0, r1) = (0, ctx.Map.Rows.Count - 1);
        var heights = new List<double>();
        for (int r = r0; r <= r1; r++) heights.Add(RowHeight(ctx.Table, r));
        double h = heights.Count > 0 ? heights.Average() : 0;
        for (int r = r0; r <= r1; r++) TableOps.SetRowHeight(ctx.Map.Rows[r], h);
        return null;
    });

    /// <summary>Hauteur affichée d'une rangée (points), parties sur plusieurs pages additionnées.</summary>
    double RowHeight(Table t, int row)
    {
        if (_layout is null) return 0;
        double h = 0;
        foreach (var page in _layout.Pages)
        {
            var mine = page.Cells.Where(b => ReferenceEquals(b.Table, t) && b.Row == row && b.Cell is not null).ToList();
            if (mine.Count > 0) h += mine.Max(b => b.H);
        }
        return h;
    }

    internal void ShowTableProperties()
    {
        var ctx = CellSelection() is { } cs0 ? Ctx(cs0) : TableAt(_caret);
        if (ctx is null || _styles is null) { StatusText.Text = "Placez le point d'insertion dans un tableau."; return; }
        var t = ctx.Table;
        var w = t.Props.Ns;
        var stylePr = _styles.TableStyleProps(t.Props.Val("tblStyle") ?? "");
        var row = ctx.Map.Rows[ctx.Row];
        double Tw(XElement? e, string attr = "w") => double.TryParse(e?.Attribute(w + attr)?.Value, out var v) ? v / 20 : 0;
        var mar = t.Props.Get("tblCellMar") ?? stylePr?.Element(w + "tblCellMar");
        XElement? Side(string a, string b) => mar?.Element(w + a) ?? mar?.Element(w + b);
        var th = row.Props.Get("trHeight");
        var grid = CurrentGrid(t) ?? ctx.Map.GridWidths();
        var dlg = new TablePropertiesWindow(t.Props.Val("jc") ?? "left", Tw(t.Props.Get("tblInd")), th is null ? null : Tw(th, "val"),
            th?.Attribute(w + "hRule")?.Value == "exact" ? "exact" : "atLeast", !row.Props.Flag("cantSplit"), row.Props.Flag("tblHeader"),
            ctx.Col < grid.Count ? grid[ctx.Col] / 20.0 : 50, ctx.Cell.Props.Val("vAlign") ?? "top",
            (Tw(Side("top", "top")), Tw(Side("bottom", "bottom")), Side("left", "start") is { } l ? Tw(l) : 5.4, Side("right", "end") is { } r ? Tw(r) : 5.4))
        { Owner = this };
        if (dlg.ShowDialog() != true) return;
        EditTable("Propriétés du tableau", (c, cs) =>
        {
            var (r0, c0, r1, c1) = Rect(c, cs);
            if (dlg.Alignment is not null || dlg.Indent is not null) TableOps.SetTableAlignment(c.Table, dlg.Alignment, dlg.Indent);
            for (int r = r0; r <= r1; r++)
            {
                var rowObj = c.Map.Rows[r];
                if (dlg.RowHeight is { } h) TableOps.SetRowHeight(rowObj, h > 0 ? h : null, dlg.Rule);
                if (dlg.AllowBreak is { } ab) TableOps.SetAllowBreak(rowObj, ab);
                if (dlg.HeaderRow is { } hr) TableOps.SetHeaderRow(rowObj, hr);
            }
            if (dlg.ColumnWidth is { } cw) TableOps.SetColumnWidth(c.Table, c0, c1, (int)Math.Round(cw * 20));
            if (dlg.VerticalAlignment is { } va)
                foreach (var cell in new TableOps.Map(c.Table).In(r0, c0, r1, c1).Select(x => x.Slot.Cell)) TableOps.SetAlignment(cell, va, null);
            var m = dlg.Margins;
            if (m.Top is not null || m.Bottom is not null || m.Left is not null || m.Right is not null) TableOps.SetCellMargins(c.Table, m.Top, m.Bottom, m.Left, m.Right);
            return null;
        });
    }

    // ---------------------------------------------------------------- alignement, données

    /// <summary>Alignement des cellules : vertical (top, center, bottom) et horizontal (left, center, right).</summary>
    internal void AlignCells(string vertical, string horizontal) => EditTable("Alignement des cellules", (ctx, cs) =>
    {
        var (r0, c0, r1, c1) = Rect(ctx, cs);
        foreach (var cell in ctx.Map.In(r0, c0, r1, c1).Select(x => x.Slot.Cell).Distinct()) TableOps.SetAlignment(cell, vertical, horizontal);
        return null;
    });

    internal void ToggleHeaderRows()
    {
        if (TableAt(_caret) is not { } ctx0) return;
        bool on = !ctx0.Map.Rows[ctx0.Row].Props.Flag("tblHeader");
        EditTable(on ? "Répéter les lignes d'en-tête" : "Ne plus répéter les lignes d'en-tête", (ctx, cs) =>
        {
            // Word : les rangées d'en-tête vont de la première jusqu'à la sélection
            var (_, _, r1, _) = Rect(ctx, cs);
            for (int r = 0; r <= r1; r++) TableOps.SetHeaderRow(ctx.Map.Rows[r], on);
            return null;
        });
    }

    internal void SortTable()
    {
        if (TableAt(_caret) is not { } ctx0) { StatusText.Text = "Placez le point d'insertion dans le tableau à trier."; return; }
        var look = TableStyles.Look(ctx0.Table.Props.Get("tblLook"));
        var header = ctx0.Map.Cells.FirstOrDefault();
        var names = Enumerable.Range(0, ctx0.Map.Columns).Select(c =>
            look.FirstRow && ctx0.Map.At(0, c) is { } s && TableOps.CellText(s.Cell) is { Length: > 0 } h ? $"{h} (colonne {c + 1})" : $"Colonne {c + 1}").ToList();
        var dlg = new SortWindow(names, ctx0.Col, look.FirstRow) { Owner = this };
        if (dlg.ShowDialog() != true) return;
        _anchor = null;
        EditTable("Trier", (ctx, _) =>
        {
            TableOps.Sort(ctx.Table, dlg.Column, dlg.Ascending, dlg.Header ? 1 : 0);
            return null;
        });
    }

    internal void ConvertTableToText()
    {
        if (TableAt(_caret) is null && CellSelection() is null) return;
        var dlg = new SeparatorWindow("Convertir le tableau en texte") { Owner = this };
        if (dlg.ShowDialog() != true) return;
        EditTable("Convertir en texte", (ctx, _) =>
        {
            if (EditRoot is not { } root) return null;
            var container = DocPos.ResolveContainer(root, ctx.Path.Take(ctx.Path.Length - 1).ToArray());
            int index = ctx.Path[^1];
            var paras = TableOps.ToText(ctx.Table, dlg.Separator, NewParaId());
            container.Blocks.RemoveAt(index);
            container.Blocks.InsertRange(index, paras);
            return new DocPos(ctx.Path.Take(ctx.Path.Length - 1).Append(index).ToArray(), 0);
        });
    }

    /// <summary>Formule (=SUM(ABOVE)…) : champ inséré au point d'insertion avec son résultat calculé.</summary>
    internal void InsertFormula()
    {
        if (TableAt(_caret) is not { } ctx0 || EditRoot is not { } body || _caret is not { } c || ToDocPos(c) is not { } at) { StatusText.Text = "Placez le point d'insertion dans une cellule."; return; }
        // Word propose ABOVE s'il y a des nombres au-dessus, sinon LEFT
        var suggestion = TableOps.Compute(ctx0.Table, ctx0.Row, ctx0.Col, "=SUM(ABOVE)") is > 0 || ctx0.Row > 0 ? "=SUM(ABOVE)" : "=SUM(LEFT)";
        var dlg = new FormulaWindow(suggestion) { Owner = this };
        if (dlg.ShowDialog() != true) return;
        var formula = dlg.Formula.StartsWith('=') ? dlg.Formula : "=" + dlg.Formula;
        if (TableOps.Compute(ctx0.Table, ctx0.Row, ctx0.Col, formula) is not { } value)
        {
            StatusText.Text = "Formule non comprise : SUM, AVERAGE, MIN, MAX, COUNT ou PRODUCT de ABOVE, BELOW, LEFT ou RIGHT.";
            return;
        }
        var pos = TextOps.Edit(_journal, body, "Formule", at.ContainerPath, at.BlockIndex, 1, at,
            () => FieldOps.Insert(body, at, formula, TableOps.FormatNumber(value)), p => p);
        AfterTyping(pos);
    }

    // ---------------------------------------------------------------- création : styles, trame, bordures

    internal void ApplyTableStyle(string name)
    {
        if (_styles is null) return;
        _anchor = _anchor;
        EditTable($"Style de tableau {TableStyleOps.French(name) ?? name}", (ctx, _) =>
        {
            if (StylesRoot(create: true) is not { } root) return null;
            var id = name == "" ? null : TableStyleOps.Ensure(_styles.Sheet, root, name);
            TableOps.SetStyle(ctx.Table, id);
            // tableau sans options de style : celles de Word par défaut (en-tête, première colonne, lignes à bandes)
            if (ctx.Table.Props.Get("tblLook") is null) TableOps.SetLook(ctx.Table, TableLook.Default);
            return null;
        }, trackStyles: true);
    }

    internal void ToggleLook(string which)
    {
        if (TableAt(_caret) is not { } ctx0 && CellSelection() is null) return;
        EditTable("Options de style du tableau", (ctx, _) =>
        {
            var l = TableStyles.Look(ctx.Table.Props.Get("tblLook"));
            l = which switch
            {
                "firstRow" => l with { FirstRow = !l.FirstRow },
                "lastRow" => l with { LastRow = !l.LastRow },
                "firstColumn" => l with { FirstColumn = !l.FirstColumn },
                "lastColumn" => l with { LastColumn = !l.LastColumn },
                "bandRows" => l with { NoHBand = !l.NoHBand },
                _ => l with { NoVBand = !l.NoVBand },
            };
            TableOps.SetLook(ctx.Table, l);
            return null;
        });
    }

    internal void ShadeCells(string? fill) => EditTable(fill is null ? "Sans trame" : "Trame", (ctx, cs) =>
    {
        var (r0, c0, r1, c1) = Rect(ctx, cs);
        foreach (var cell in ctx.Map.In(r0, c0, r1, c1).Select(x => x.Slot.Cell).Distinct()) TableOps.SetShading(cell, fill);
        return null;
    });

    /// <summary>Trait utilisé par le menu Bordures (style, épaisseur, couleur), comme le « stylet » de Word.</summary>
    TableOps.Line _borderPen = new("single", 4, "auto");

    internal void SetCellBorders(string which) => EditTable("Bordures", (ctx, cs) =>
    {
        var (r0, c0, r1, c1) = Rect(ctx, cs);
        TableOps.SetBorders(ctx.Table, r0, c0, r1, c1, which, _borderPen);
        return null;
    });

    static readonly (string Name, string? Rgb)[] ShadeColors =
    {
        ("Aucune couleur", null), ("Gris clair", "D9D9D9"), ("Gris", "A6A6A6"), ("Bleu clair", "DEEAF6"), ("Bleu", "9CC2E5"), ("Vert clair", "E2EFD9"),
        ("Vert", "A8D08D"), ("Jaune", "FFF2CC"), ("Orange clair", "FBE4D5"), ("Rouge clair", "F4CCCC"), ("Violet clair", "E4DFEC"),
    };

    internal void ShowShadingMenu(FrameworkElement anchor) => ShowMenu(anchor, ShadeColors.Select(c => (c.Name,
        (UIElement?)(c.Rgb is { } rgb ? new Border { Width = 14, Height = 14, BorderBrush = Brushes.Gray, BorderThickness = new Thickness(0.5),
            Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#" + rgb)) } : null),
        (Action)(() => ShadeCells(c.Rgb)))));

    internal void ShowBordersMenu(FrameworkElement anchor)
    {
        var items = new List<(string, UIElement?, Action)>
        {
            ("Bordure inférieure", null, () => SetCellBorders("bottom")),
            ("Bordure supérieure", null, () => SetCellBorders("top")),
            ("Bordure gauche", null, () => SetCellBorders("left")),
            ("Bordure droite", null, () => SetCellBorders("right")),
            ("Aucune bordure", null, () => SetCellBorders("none")),
            ("Toutes les bordures", null, () => SetCellBorders("all")),
            ("Bordures extérieures", null, () => SetCellBorders("outside")),
            ("Bordures intérieures", null, () => SetCellBorders("inside")),
            ("Bordure horizontale intérieure", null, () => SetCellBorders("insideH")),
            ("Bordure verticale intérieure", null, () => SetCellBorders("insideV")),
        };
        foreach (var (label, pen) in new (string, TableOps.Line)[]
        {
            ("Trait : simple ½ pt", new("single", 4, "auto")), ("Trait : simple 1 pt", new("single", 8, "auto")), ("Trait : simple 2¼ pt", new("single", 18, "auto")),
            ("Trait : double", new("double", 4, "auto")), ("Trait : pointillés", new("dotted", 4, "auto")), ("Trait : tirets", new("dashed", 4, "auto")),
            ("Trait : bleu ½ pt", new("single", 4, "2F5496")), ("Trait : gris ½ pt", new("single", 4, "A6A6A6")),
        })
            items.Add((label + (_borderPen == pen ? "  ✓" : ""), null, () => { _borderPen = pen; StatusText.Text = $"{label} : choisissez maintenant les bordures à tracer."; }));
        ShowMenu(anchor, items);
    }

    /// <summary>Galerie des styles de tableau : aperçu de chaque style (en-tête, bandes, couleurs du thème du document).</summary>
    internal void ShowTableStylesGallery(FrameworkElement anchor)
    {
        if (_styles is null || _pkg is null) return;
        var popup = new Popup { PlacementTarget = anchor, Placement = PlacementMode.Bottom, StaysOpen = false, AllowsTransparency = true };
        var panel = new WrapPanel { Width = 7 * 76, Margin = new Thickness(6) };
        // feuille de styles de travail : chaque style y est ajouté pour l'aperçu, sans toucher au document
        var sheet = _styles.Sheet;
        var scratch = sheet.Styles is { } st ? new XElement(st) : new XElement(sheet.W + "styles");
        var look = TableAt(_caret) is { } ctx ? TableStyles.Look(ctx.Table.Props.Get("tblLook")) : TableLook.Default;
        foreach (var (name, french) in TableStyleOps.Gallery())
        {
            var id = TableStyleOps.Ensure(sheet, scratch, name);
            var preview = StylePreview(new StyleResolver(new StyleSheet(scratch, sheet.Numbering, sheet.Theme, sheet.Settings, sheet.W)), id, look);
            var b = new Button { Content = preview, Width = 72, Height = 54, Margin = new Thickness(2), ToolTip = french, Focusable = false, Background = Brushes.White, Padding = new Thickness(4) };
            b.Click += (_, _) => { popup.IsOpen = false; ApplyTableStyle(name); _pagesView?.Focus(); };
            panel.Children.Add(b);
        }
        var clear = new Button { Content = "Effacer le style (Tableau Normal)", Margin = new Thickness(8, 0, 8, 8), HorizontalAlignment = HorizontalAlignment.Left, Padding = new Thickness(8, 3, 8, 3) };
        clear.Click += (_, _) => { popup.IsOpen = false; ApplyTableStyle(""); _pagesView?.Focus(); };
        var root = new StackPanel();
        root.Children.Add(new ScrollViewer { Content = panel, MaxHeight = 420, VerticalScrollBarVisibility = ScrollBarVisibility.Auto });
        root.Children.Add(clear);
        popup.Child = new Border { Background = R("Chrome"), BorderBrush = R("Line"), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(6), Child = root };
        popup.IsOpen = true;
        LastPopup = popup;
    }

    /// <summary>Aperçu d'un style de tableau : 4 × 5 cellules avec trame, traits et couleur du texte résolus.</summary>
    static UIElement StylePreview(StyleResolver resolver, string? styleId, TableLook look)
    {
        const int rows = 4, cols = 5;
        const double cw = 12, ch = 9;
        var canvas = new Canvas { Width = cols * cw, Height = rows * ch };
        var w = resolver.Sheet.W;
        var stylePr = resolver.TableStyleProps(styleId);
        var tblBorders = stylePr?.Element(w + "tblBorders");
        Brush? ColorOf(string? hex) => hex is { Length: 6 } h ? new SolidColorBrush((Color)ColorConverter.ConvertFromString("#" + h)) : null;
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
            {
                var cs = new CellStyle(styleId, TableStyles.Regions(look, r, rows, c, cols));
                var level = resolver.TableLevel(cs);
                var shd = level.TcPr?.Element(w + "shd");
                var fill = shd?.Attribute(w + "fill")?.Value;
                var rect = new System.Windows.Shapes.Rectangle { Width = cw, Height = ch, Fill = ColorOf(fill) ?? Brushes.White };
                Canvas.SetLeft(rect, c * cw);
                Canvas.SetTop(rect, r * ch);
                canvas.Children.Add(rect);
                // « texte » : petit trait de la couleur du texte de la cellule
                var rf = level.RPr;
                var color = rf?.Element(w + "color")?.Attribute(w + "val")?.Value;
                bool bold = rf?.Element(w + "b") is { } b && b.Attribute(w + "val")?.Value is null or "1" or "true";
                var text = new System.Windows.Shapes.Rectangle { Width = cw - 5, Height = bold ? 2 : 1.2, Fill = ColorOf(color) ?? Brushes.DimGray };
                Canvas.SetLeft(text, c * cw + 2.5);
                Canvas.SetTop(text, r * ch + ch / 2 - 1);
                canvas.Children.Add(text);
                // traits : bordures de la région (sinon du tableau)
                var tc = level.TcPr?.Element(w + "tcBorders");
                void Edge(string side, string table, double x1, double y1, double x2, double y2)
                {
                    var e = tc?.Element(w + side) ?? tblBorders?.Element(w + table);
                    if (e is null || e.Attribute(w + "val")?.Value is null or "nil" or "none") return;
                    var col = e.Attribute(w + "color")?.Value;
                    canvas.Children.Add(new System.Windows.Shapes.Line { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, StrokeThickness = 0.8, Stroke = ColorOf(col == "auto" ? "000000" : col) ?? Brushes.Black });
                }
                Edge("top", r == 0 ? "top" : "insideH", c * cw, r * ch, (c + 1) * cw, r * ch);
                Edge("left", c == 0 ? "left" : "insideV", c * cw, r * ch, c * cw, (r + 1) * ch);
                if (r == rows - 1) Edge("bottom", "bottom", c * cw, (r + 1) * ch, (c + 1) * cw, (r + 1) * ch);
                if (c == cols - 1) Edge("right", "right", (c + 1) * cw, r * ch, (c + 1) * cw, (r + 1) * ch);
            }
        return canvas;
    }

    // ================================================================ onglets contextuels

    RadioButton? _tableDesignTab, _tableLayoutTab;
    StackPanel? _tableDesignPanel, _tableLayoutPanel;
    readonly Dictionary<string, ToggleButton> _lookToggles = new();
    ToggleButton? _headerRowsToggle, _gridlinesToggle;
    Table? _tabsTable;

    /// <summary>Onglets « Création de tableau » et « Disposition » : visibles quand le point d'insertion est dans un tableau.</summary>
    void UpdateTableTabs()
    {
        var ctx = _layout is null ? null : CellSelection() is { } cs ? Ctx(cs) : TableAt(_caret);
        bool show = ctx is not null;
        if (_tableDesignTab is null)
        {
            if (!show) return;
            _tableDesignTab = new RadioButton { Style = (Style)FindResource("RibbonTab"), Content = "Création de tableau", Name = "TabTableDesign", Foreground = R("SelectedText") };
            _tableLayoutTab = new RadioButton { Style = (Style)FindResource("RibbonTab"), Content = "Disposition", Name = "TabTableLayout", Foreground = R("SelectedText") };
            _tableDesignTab.Checked += Tab_Checked;
            _tableLayoutTab.Checked += Tab_Checked;
            TabStrip.Children.Add(_tableDesignTab);
            TabStrip.Children.Add(_tableLayoutTab);
            _tableDesignPanel = BuildTableDesignPanel();
            _tableLayoutPanel = BuildTableLayoutPanel();
            _panels["TabTableDesign"] = _tableDesignPanel;
            _panels["TabTableLayout"] = _tableLayoutPanel;
            RibbonPanels.Children.Add(_tableDesignPanel);
            RibbonPanels.Children.Add(_tableLayoutPanel);
        }
        var vis = show ? Visibility.Visible : Visibility.Collapsed;
        if (_tableDesignTab.Visibility != vis)
        {
            _tableDesignTab.Visibility = vis;
            _tableLayoutTab!.Visibility = vis;
            if (!show && (_tableDesignTab.IsChecked == true || _tableLayoutTab.IsChecked == true)) TabHome.IsChecked = true;
        }
        if (ctx is null) { _tabsTable = null; return; }
        _tabsTable = ctx.Table;
        var look = TableStyles.Look(ctx.Table.Props.Get("tblLook"));
        void Set(string k, bool v) { if (_lookToggles.TryGetValue(k, out var t) && t.IsChecked != v) t.IsChecked = v; }
        Set("firstRow", look.FirstRow); Set("lastRow", look.LastRow); Set("firstColumn", look.FirstColumn); Set("lastColumn", look.LastColumn);
        Set("bandRows", !look.NoHBand); Set("bandColumns", !look.NoVBand);
        if (_headerRowsToggle is not null) _headerRowsToggle.IsChecked = ctx.Map.Rows[ctx.Row].Props.Flag("tblHeader");
    }

    Button TBig(string glyph, string label, string tip, Action<Button> act)
    {
        var b = new Button { Style = (Style)FindResource("RBig"), Tag = glyph, Content = label, ToolTip = tip, Focusable = false };
        b.Click += (_, _) => { act(b); };
        return b;
    }

    ToggleButton TToggle(string glyph, string label, string tip, Action act)
    {
        var t = new ToggleButton { Style = (Style)FindResource("RBig"), Tag = glyph, Content = label, ToolTip = tip, Focusable = false };
        t.Click += (_, _) => { act(); _pagesView?.Focus(); };
        return t;
    }

    static void TGroup(StackPanel panel, string header, params UIElement[] items)
    {
        var sp = new StackPanel { Orientation = Orientation.Horizontal };
        foreach (var i in items) sp.Children.Add(i);
        panel.Children.Add(new HeaderedContentControl { Style = (Style)Application.Current.FindResource("RGroup"), Header = header, Content = sp });
    }

    StackPanel BuildTableDesignPanel()
    {
        var panel = new StackPanel { Orientation = Orientation.Horizontal, Visibility = Visibility.Collapsed };
        var options = new UniformGrid { Columns = 2, Rows = 3, Margin = new Thickness(4, 2, 4, 2) };
        foreach (var (key, label, tip) in new[]
        {
            ("firstRow", "Ligne d'en-tête", "Mise en forme de la première ligne"), ("firstColumn", "Première colonne", "Mise en forme de la première colonne"),
            ("lastRow", "Ligne des totaux", "Mise en forme de la dernière ligne"), ("lastColumn", "Dernière colonne", "Mise en forme de la dernière colonne"),
            ("bandRows", "Lignes à bandes", "Une ligne sur deux mise en forme"), ("bandColumns", "Colonnes à bandes", "Une colonne sur deux mise en forme"),
        })
        {
            var t = new ToggleButton
            {
                Content = label, ToolTip = tip, Focusable = false, Margin = new Thickness(2, 1, 2, 1), Padding = new Thickness(8, 0, 8, 0), Height = 21, FontSize = 11.5,
                HorizontalContentAlignment = HorizontalAlignment.Left, VerticalContentAlignment = VerticalAlignment.Center,
            };
            t.Click += (_, _) => { ToggleLook(key); _pagesView?.Focus(); };
            _lookToggles[key] = t;
            options.Children.Add(t);
        }
        TGroup(panel, "Options de style du tableau", options);
        TGroup(panel, "Styles de tableau", TBig("\uE771", "Styles", "Galerie des styles de tableau (couleurs du thème du document)", ShowTableStylesGallery));
        TGroup(panel, "Trame et bordures",
            TBig("\uE7E6", "Trame", "Couleur de fond des cellules sélectionnées", ShowShadingMenu),
            TBig("\uE80A", "Bordures", "Bordures des cellules sélectionnées ; choisir d'abord le trait", ShowBordersMenu));
        return panel;
    }

    StackPanel BuildTableLayoutPanel()
    {
        var panel = new StackPanel { Orientation = Orientation.Horizontal, Visibility = Visibility.Collapsed };
        _gridlinesToggle = TToggle("\uF0E2", "Quadrillage", "Afficher le quadrillage des tableaux sans bordures (non imprimé)", () =>
        {
            Layout.PageRenderer.ShowGridlines = !Layout.PageRenderer.ShowGridlines;
            _drawings?.Clear();
            _pagesView?.InvalidateVisual();
        });
        _gridlinesToggle.IsChecked = Layout.PageRenderer.ShowGridlines;
        TGroup(panel, "Tableau",
            TBig("\uE8B3", "Sélectionner", "Sélectionner la cellule, la ligne, la colonne ou le tableau", b => ShowMenu(b, new (string, UIElement?, Action)[]
            {
                ("Sélectionner la cellule", null, () => SelectInTable("cell")), ("Sélectionner la colonne", null, () => SelectInTable("column")),
                ("Sélectionner la ligne", null, () => SelectInTable("row")), ("Sélectionner le tableau", null, () => SelectInTable("table")),
            })),
            _gridlinesToggle,
            TBig("\uE713", "Propriétés", "Propriétés du tableau, des lignes, des colonnes et des cellules", _ => ShowTableProperties()));
        TGroup(panel, "Lignes et colonnes",
            TBig("\uE74D", "Supprimer", "Supprimer des lignes, des colonnes ou le tableau", b => ShowMenu(b, new (string, UIElement?, Action)[]
            {
                ("Supprimer les colonnes", null, DeleteColumns), ("Supprimer les lignes", null, DeleteRows), ("Supprimer le tableau", null, DeleteTable),
            })),
            TBig("\uE74A", "Au-dessus", "Insérer une ligne au-dessus", _ => { InsertRows(below: false); _pagesView?.Focus(); }),
            TBig("\uE74B", "En dessous", "Insérer une ligne en dessous (Tab dans la dernière cellule)", _ => { InsertRows(below: true); _pagesView?.Focus(); }),
            TBig("\uE72B", "À gauche", "Insérer une colonne à gauche", _ => { InsertColumns(right: false); _pagesView?.Focus(); }),
            TBig("\uE72A", "À droite", "Insérer une colonne à droite", _ => { InsertColumns(right: true); _pagesView?.Focus(); }));
        TGroup(panel, "Fusionner",
            TBig("\uE8A0", "Fusionner", "Fusionner les cellules sélectionnées", _ => { MergeCells(); _pagesView?.Focus(); }),
            TBig("\uEA5B", "Fractionner", "Fractionner la cellule en plusieurs colonnes ou lignes", _ => SplitCells()),
            TBig("\uE81E", "Fractionner le tableau", "Couper le tableau au-dessus de la ligne du point d'insertion (Ctrl+Maj+Entrée)", _ => { SplitTableHere(); _pagesView?.Focus(); }));
        TGroup(panel, "Taille de cellule",
            TBig("\uECA5", "Ajustement", "Ajuster au contenu, à la fenêtre, ou largeur de colonne fixe", b => ShowMenu(b, new (string, UIElement?, Action)[]
            {
                ("Ajuster au contenu", null, () => AutoFitTable("content")), ("Ajuster à la fenêtre", null, () => AutoFitTable("window")),
                ("Largeur de colonne fixe", null, () => AutoFitTable("fixed")),
            })),
            TBig("\uE8E4", "Distribuer les lignes", "Même hauteur pour les lignes sélectionnées (tout le tableau sans sélection)", _ => { DistributeRows(); _pagesView?.Focus(); }),
            TBig("\uE8A9", "Distribuer les colonnes", "Même largeur pour les colonnes sélectionnées (tout le tableau sans sélection)", _ => { DistributeColumns(); _pagesView?.Focus(); }));
        TGroup(panel, "Alignement",
            TBig("\uE8E3", "Aligner", "Alignement du texte dans les cellules (9 positions)", b => ShowMenu(b,
                from v in new[] { ("top", "en haut"), ("center", "au centre"), ("bottom", "en bas") }
                from h in new[] { ("left", "à gauche"), ("center", "centré"), ("right", "à droite") }
                select ($"{char.ToUpper(v.Item2[0])}{v.Item2[1..]}, {h.Item2}", (UIElement?)null, (Action)(() => AlignCells(v.Item1, h.Item1))))),
            TBig("\uE8A1", "Marges", "Marges de cellule par défaut du tableau", _ => ShowTableProperties()));
        _headerRowsToggle = TToggle("\uE8BC", "Répéter l'en-tête", "Répéter les lignes d'en-tête en haut de chaque page", ToggleHeaderRows);
        TGroup(panel, "Données",
            TBig("\uE8CB", "Trier", "Trier les lignes selon une colonne", _ => SortTable()),
            _headerRowsToggle,
            TBig("\uE8D2", "Convertir en texte", "Remplacer le tableau par du texte (cellules séparées par des tabulations)", _ => ConvertTableToText()),
            TBig("\uE8EF", "Formule", "Formule : =SUM(ABOVE), =AVERAGE(LEFT)…", _ => InsertFormula()));
        return panel;
    }

    /// <summary>Entrées du menu contextuel dans un tableau.</summary>
    void AddTableContextItems(ContextMenu menu)
    {
        if (TableAt(_caret) is null && CellSelection() is null) return;
        MenuItem Item(ItemsControl parent, string header, Action act, bool enabled = true)
        {
            var mi = new MenuItem { Header = header, IsEnabled = enabled };
            mi.Click += (_, _) => { act(); _pagesView?.Focus(); };
            parent.Items.Add(mi);
            return mi;
        }
        menu.Items.Add(new Separator());
        var insert = new MenuItem { Header = "Insérer" };
        Item(insert, "Lignes au-dessus", () => InsertRows(false));
        Item(insert, "Lignes en dessous", () => InsertRows(true));
        Item(insert, "Colonnes à gauche", () => InsertColumns(false));
        Item(insert, "Colonnes à droite", () => InsertColumns(true));
        menu.Items.Add(insert);
        var delete = new MenuItem { Header = "Supprimer" };
        Item(delete, "Lignes", DeleteRows);
        Item(delete, "Colonnes", DeleteColumns);
        Item(delete, "Tableau", DeleteTable);
        menu.Items.Add(delete);
        Item(menu, "Fusionner les cellules", MergeCells, CellSelection() is not null);
        Item(menu, "Fractionner les cellules…", () => SplitCells());
        Item(menu, "Propriétés du tableau…", ShowTableProperties);
    }

    /// <summary>
    /// Frappe ou suppression sur une sélection de cellules : le contenu des cellules est effacé (leur mise en forme de
    /// paragraphe gardée), le texte tapé va dans la première (Word).
    /// </summary>
    void ReplaceCells(CellSel cs, string text, string label) => EditTable(label, (ctx, _) =>
    {
        if (EditRoot is not { } root) return null;
        var cells = CellsOf(cs);
        foreach (var cell in cells)
        {
            var like = cell.Content.Paragraphs().LastOrDefault();
            var p = new Paragraph { Props = like?.Props.Clone() ?? new PropertyBag("pPr", ctx.Table.Props.Ns) };
            p.Props.Remove("sectPr");
            if (like is not null)
                foreach (var a in like.Attributes)
                    p.Attributes.Add(a.Name == FlowEdits.W14 + "paraId" || a.Name == FlowEdits.W14 + "textId" ? new XAttribute(a.Name, NewParaId()()) : new XAttribute(a));
            cell.Content.Blocks.Clear();
            cell.Content.Blocks.Add(p);
        }
        var pos = DocPos.Of(root, (Paragraph)cells[0].Content.Blocks[0]);
        return text.Length > 0 ? TextOps.InsertMultiline(root, pos, text, NewParaId()) : pos;
    });

    // ================================================================ souris : bordures de colonnes et de rangées

    /// <summary>Limite tirée à la souris : tableau, colonne de grille (bord gauche de la limite) ou rangée, position de départ.</summary>
    sealed record BorderDrag(Table Table, int[] Path, bool Vertical, int Boundary, int Row, double Start, int Page);
    BorderDrag? _borderDrag;

    /// <summary>Bordure de cellule sous le pointeur (à 2,5 points près) : limite de colonne ou bas de rangée.</summary>
    BorderDrag? BorderAt(int page, Point at)
    {
        if (_layout is null || _hf is not null || page < 0 || page >= _layout.Pages.Count || EditRoot is not { } root) return null;
        const double tol = 2.5;
        foreach (var b in _layout.Pages[page].Cells)
        {
            if (b.Cell is null || b.Table is null || b.Nested) continue;
            if (at.Y < b.Y - tol || at.Y > b.Y + b.H + tol || at.X < b.X - tol || at.X > b.X + b.W + tol) continue;
            bool right = Math.Abs(at.X - (b.X + b.W)) <= tol, left = Math.Abs(at.X - b.X) <= tol && b.Column > 0;
            bool bottom = Math.Abs(at.Y - (b.Y + b.H)) <= tol;
            if (!right && !left && !bottom) continue;
            int[] path;
            try { path = TablePathOf(root, b.Table); }
            catch (ArgumentException) { continue; }
            var map = new TableOps.Map(b.Table);
            if (right || left)
            {
                var slot = map.At(b.Row, b.Column);
                int boundary = right ? b.Column + (slot?.Span ?? 1) : b.Column;
                return new BorderDrag(b.Table, path, true, boundary, b.Row, at.X, page);
            }
            return new BorderDrag(b.Table, path, false, 0, b.Row, at.Y, page);
        }
        return null;
    }

    static int[] TablePathOf(Story root, Table t)
    {
        for (int i = 0; i < root.Blocks.Count; i++)
            if (ReferenceEquals(root.Blocks[i], t)) return new[] { i };
        throw new ArgumentException("tableau imbriqué");
    }

    /// <summary>Survol : curseur de redimensionnement sur une bordure de cellule. Vrai si le curseur a été posé.</summary>
    bool HoverTableBorder(int page, Point at)
    {
        if (_pagesView is null) return false;
        if (_borderDrag is not null) return true;
        if (BorderAt(page, at) is not { } b) return false;
        _pagesView.Cursor = b.Vertical ? Cursors.SizeWE : Cursors.SizeNS;
        return true;
    }

    /// <summary>Bouton enfoncé sur une bordure : début du redimensionnement. Vrai si pris en charge.</summary>
    bool StartBorderDrag(int page, Point at)
    {
        if (BorderAt(page, at) is not { } b || _pagesView is null) return false;
        _borderDrag = b;
        _pagesView.CaptureMouse();
        StatusText.Text = b.Vertical ? "Relâchez pour fixer la largeur de la colonne (le tableau garde sa largeur)." : "Relâchez pour fixer la hauteur de la ligne.";
        return true;
    }

    /// <summary>Bouton relâché : largeur de colonne ou hauteur de rangée appliquée (une étape annulable).</summary>
    bool EndBorderDrag(int page, Point at)
    {
        if (_borderDrag is not { } d || _pagesView is null) return false;
        _borderDrag = null;
        _pagesView.ReleaseMouseCapture();
        double delta = (d.Vertical ? at.X - d.Start : at.Y - d.Start);
        if (page != d.Page || Math.Abs(delta) < 0.5) { StatusText.Text = ""; return true; }
        if (EditRoot is not { } body) return true;
        var (savedAnchor, savedCaret) = (_anchor, _caret);
        try
        {
            using var tx = _journal.Begin(d.Vertical ? "Largeur de colonne" : "Hauteur de ligne", body, Array.Empty<int>(), d.Path[0], 1, null);
            if (d.Vertical)
            {
                var grid = CurrentGrid(d.Table);
                if (grid is not null) TableOps.SetGrid(d.Table, grid);
                TableOps.MoveBoundary(d.Table, d.Boundary, (int)Math.Round(delta * 20));
            }
            else
            {
                var map = new TableOps.Map(d.Table);
                double h = Math.Max(4, RowHeight(d.Table, d.Row) + delta);
                TableOps.SetRowHeight(map.Rows[d.Row], h);
            }
            tx.Commit();
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException) { StatusText.Text = $"Redimensionnement impossible : {ex.Message}"; return true; }
        Relayout();
        _anchor = savedAnchor;
        _caret = savedCaret;
        Scroller.UpdateLayout();
        UpdateCaret(scroll: false);
        UpdateModifiedMarker();
        StatusText.Text = (d.Vertical ? "Largeur de colonne modifiée" : "Hauteur de ligne modifiée") + ". Ctrl+Z pour annuler.";
        return true;
    }
}
