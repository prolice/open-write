# Fabrique, avec Word (COM), des documents de calibration des tableaux (lot 6) au contenu neutre :
#   styles.docx    : tableaux aux styles intégrés de Word (en-tête, bandes, première colonne, ligne des totaux) ;
#   fusions.docx   : cellules fusionnées horizontalement et verticalement, alignement vertical, marges de cellule,
#                    tableau imbriqué, tableau centré et retrait ;
#   ajustement.docx: tableaux « ajuster au contenu » (largeurs calculées par Word) ;
#   long.docx      : tableau de 90 rangées sur plusieurs pages, rangée d'en-tête répétée, rangées hautes.
# Noms de styles : ceux d'un Word en français.
# Usage : powershell.exe -ExecutionPolicy Bypass -File tools\word-make-tables.ps1 <dossier de sortie>
param([Parameter(Mandatory)][string]$Dest)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $Dest | Out-Null

$words = 'alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau upsilon phi chi psi omega'.Split(' ')
function Lorem([int]$n, [int]$seed) { ($(for ($k = 0; $k -lt $n; $k++) { $words[($seed * 7 + $k * 3) % $words.Length] }) -join ' ') }

function EndRange($doc) { $r = $doc.Content; $r.Collapse(0); return $r }
function Para($doc, [string]$text) {
    $r = EndRange $doc
    $r.InsertAfter($text)
    $r.InsertParagraphAfter()
}
function NewTable($doc, [int]$rows, [int]$cols) {
    $r = EndRange $doc
    $t = $doc.Tables.Add($r, $rows, $cols, 1, 0)   # wdWord9TableBehavior, wdAutoFitFixed
    return $t
}
function Fill($t, [string]$prefix) {
    for ($r = 1; $r -le $t.Rows.Count; $r++) {
        for ($c = 1; $c -le $t.Columns.Count; $c++) { $t.Cell($r, $c).Range.Text = "$prefix L$r C$c" }
    }
}
# style par son nom local, espaces insécables comprises (« Tableau Grille 4 - Accentuation 1 » s'écrit avec U+00A0)
function StyleOf($doc, [string]$name) {
    # -ceq : « Tableau Liste 3 » (List Table 3) et « Tableau liste 3 » (ancien Table List 3) ne diffèrent que par la casse
    foreach ($st in $doc.Styles) { if ($st.NameLocal.Replace([char]0xA0, ' ') -ceq $name) { return $st } }
    throw "style introuvable : $name"
}
function Save($doc, [string]$name) {
    $path = [IO.Path]::Combine($Dest, $name)
    $doc.SaveAs2($path, 16)                          # wdFormatDocumentDefault
    $doc.Close([ref]0)
    Write-Host "  $name"
}

$grid = 'Grille du tableau'                            # Table Grid
$word = New-Object -ComObject Word.Application
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0

    # ---------------------------------------------------------------- styles
    $doc = $word.Documents.Add()
    # noms locaux (Word en français) : Grid Table 4 - Accent 1, Grid Table 5 Dark - Accent 2, List Table 3 - Accent 5,
    # Plain Table 1, Grid Table 1 Light, List Table 4 - Accent 6, Grid Table 6 Colorful
    $styles = @('Tableau Grille 4 - Accentuation 1', 'Tableau Grille 5 Foncé - Accentuation 2', 'Tableau Liste 3 - Accentuation 5', 'Tableau simple 1', 'Tableau Grille 1 Clair', 'Tableau Liste 4 - Accentuation 6', 'Tableau Grille 6 Couleur')
    $i = 0
    foreach ($s in $styles) {
        $i++
        Para $doc "Tableau $i"
        $t = NewTable $doc 5 4
        Fill $t "S$i"
        $t.Style = (StyleOf $doc $s)
        $t.ApplyStyleHeadingRows = $true
        $t.ApplyStyleFirstColumn = ($i % 2 -eq 1)
        $t.ApplyStyleLastRow = ($i -ge 4)
        $t.ApplyStyleLastColumn = ($i -eq 5)
        $t.ApplyStyleRowBands = ($i -ne 6)
        $t.ApplyStyleColumnBands = ($i -eq 6 -or $i -eq 7)
    }
    Para $doc "Fin des styles"
    Save $doc 'styles.docx'

    # ---------------------------------------------------------------- fusions
    $doc = $word.Documents.Add()
    Para $doc "Fusion horizontale et verticale"
    $t = NewTable $doc 5 4
    Fill $t "F"
    $t.Style = (StyleOf $doc $grid)
    $t.Cell(1, 1).Merge($t.Cell(1, 2))
    $t.Cell(2, 1).Merge($t.Cell(4, 1))
    $t.Cell(2, 1).Range.Text = (Lorem 30 1)
    $t.Cell(3, 3).Merge($t.Cell(4, 4))
    Para $doc "Alignement vertical"
    $t = NewTable $doc 2 3
    Fill $t "V"
    $t.Style = (StyleOf $doc $grid)
    $t.Cell(1, 1).Range.Text = (Lorem 40 2)
    $t.Cell(1, 2).VerticalAlignment = 1    # centre
    $t.Cell(1, 3).VerticalAlignment = 3    # bas
    $t.Rows.Item(2).Height = 60
    $t.Rows.Item(2).HeightRule = 2         # exactement
    $t.Cell(2, 2).VerticalAlignment = 1
    Para $doc "Marges de cellule"
    $t = NewTable $doc 3 3
    Fill $t "M"
    $t.Style = (StyleOf $doc $grid)
    $t.TopPadding = 6; $t.BottomPadding = 4; $t.LeftPadding = 12; $t.RightPadding = 2
    $t.Cell(2, 2).TopPadding = 14
    $t.Cell(2, 2).LeftPadding = 20
    Para $doc "Tableau imbrique"
    $t = NewTable $doc 2 2
    Fill $t "E"
    $t.Style = (StyleOf $doc $grid)
    $inner = $t.Cell(2, 2).Tables.Add($t.Cell(2, 2).Range, 2, 2, 1, 0)
    Fill $inner "I"
    $t.Cell(1, 2).Range.Text = (Lorem 25 3)
    Para $doc "Tableau centre"
    $t = NewTable $doc 2 2
    Fill $t "C"
    $t.Style = (StyleOf $doc $grid)
    $t.PreferredWidthType = 3; $t.PreferredWidth = 250     # wdPreferredWidthPoints
    $t.Rows.Alignment = 1
    Para $doc "Tableau en retrait"
    $t = NewTable $doc 2 2
    Fill $t "R"
    $t.Style = (StyleOf $doc $grid)
    $t.PreferredWidthType = 2; $t.PreferredWidth = 60      # pourcentage
    $t.Rows.LeftIndent = 40
    Para $doc "Fin des fusions"
    Save $doc 'fusions.docx'

    # ---------------------------------------------------------------- ajustement
    $doc = $word.Documents.Add()
    $texts = @(
        @('Nom', 'Description longue qui doit passer sur plusieurs lignes si la place manque vraiment', 'Prix'),
        @('A', 'B', 'C'),
        @('Un', 'deux mots', 'trois mots ici')
    )
    $i = 0
    foreach ($row in $texts) {
        $i++
        Para $doc "Ajustement $i"
        $t = NewTable $doc 3 3
        for ($r = 1; $r -le 3; $r++) { for ($c = 1; $c -le 3; $c++) { $t.Cell($r, $c).Range.Text = $row[$c - 1] + $(if ($r -gt 1) { " $r" } else { '' }) } }
        $t.Style = (StyleOf $doc $grid)
        $t.AutoFitBehavior(1)                                 # wdAutoFitContent
    }
    Para $doc "Ajustement fenetre"
    $t = NewTable $doc 3 3
    Fill $t "W"
    $t.Style = (StyleOf $doc $grid)
    $t.AutoFitBehavior(2)                                     # wdAutoFitWindow
    Para $doc "Fin des ajustements"
    Save $doc 'ajustement.docx'

    # ---------------------------------------------------------------- long
    $doc = $word.Documents.Add()
    Para $doc "Tableau long"
    $t = NewTable $doc 90 3
    $t.Style = (StyleOf $doc 'Tableau Grille 4 - Accentuation 1')
    for ($r = 1; $r -le 90; $r++) {
        $t.Cell($r, 1).Range.Text = "Rangee $r"
        $t.Cell($r, 2).Range.Text = (Lorem (3 + ($r * 5) % 40) $r)
        $t.Cell($r, 3).Range.Text = "$($r * 3)"
    }
    $t.Rows.Item(1).HeadingFormat = -1
    $t.Rows.Item(20).AllowBreakAcrossPages = 0
    $t.Cell(20, 2).Range.Text = (Lorem 160 20)
    Para $doc "Fin du tableau long"
    Save $doc 'long.docx'

    # ---------------------------------------------------------------- galerie (définitions des styles intégrés)
    $doc = $word.Documents.Add()
    $families = @('Tableau Grille 1 Clair', 'Tableau Grille 4', 'Tableau Grille 5 Foncé', 'Tableau Liste 3', 'Tableau Liste 4')
    foreach ($f in $families) {
        foreach ($n in @('', ' - Accentuation 1', ' - Accentuation 2', ' - Accentuation 3', ' - Accentuation 4', ' - Accentuation 5', ' - Accentuation 6')) {
            Para $doc "$f$n"
            $t = NewTable $doc 2 2
            $t.Style = (StyleOf $doc "$f$n")
        }
    }
    Para $doc "Fin de la galerie"
    Save $doc 'galerie.docx'
} catch { Write-Host ("ligne " + $_.InvocationInfo.ScriptLineNumber + " : " + $_.InvocationInfo.Line); throw } finally {
    $word.Quit([ref]0)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
