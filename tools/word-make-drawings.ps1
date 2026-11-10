# Fabrique, avec Word (COM), des documents de calibration des objets flottants (lot 7) au contenu neutre : une forme
# rectangulaire (ou une ellipse) ancrée dans un texte de plusieurs paragraphes, avec un habillage et une position donnés.
#   carre-gauche, carre-droite, carre-milieu      habillage carré, objet contre la marge gauche / droite / centré
#   cote-gauche, cote-droit, cote-large           habillage carré au milieu, texte d'un seul côté (left, right, largest)
#   haut-bas                                      haut et bas
#   rapproche                                     ellipse, habillage rapproché
#   retrait                                       paragraphes en retrait de 36 pt, objet contre la marge gauche
#   distances                                     distances au texte de 0 / 18 pt
#   ancre-bas                                     objet ancré au 5e paragraphe, placé en haut de la marge (page refaite)
#   page                                          objet placé par rapport à la page (coin haut droit)
#   devant, derriere                              sans habillage
#   zone-texte                                    zone de texte avec deux paragraphes, habillage carré
# Usage : powershell.exe -ExecutionPolicy Bypass -File tools\word-make-drawings.ps1 <dossier de sortie>
param([Parameter(Mandatory)][string]$Dest)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $Dest | Out-Null

$words = 'alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau upsilon phi chi psi omega'.Split(' ')
function Lorem([int]$n, [int]$seed) { ($(for ($k = 0; $k -lt $n; $k++) { $words[($seed * 7 + $k * 3) % $words.Length] }) -join ' ') }

function Body($doc, [int]$paras, [double]$indent = 0) {
    $r = $doc.Content
    for ($i = 0; $i -lt $paras; $i++) {
        $r.InsertAfter((Lorem 60 $i))
        if ($i -lt $paras - 1) { $r.InsertParagraphAfter() }
    }
    if ($indent -ne 0) { $doc.Content.ParagraphFormat.LeftIndent = $indent }
}

# type : 1 rectangle, 9 ellipse ; wrap : 0 carré, 1 rapproché, 2 au travers, 4 haut et bas, 5 derrière, 3 devant
function AddShape($doc, [int]$anchorPara, [int]$type, [double]$left, [double]$top, [double]$w, [double]$h,
                  [int]$relH, [int]$relV, [int]$wrap, [int]$side = 0, [double]$dist = -1) {
    $anchor = $doc.Paragraphs.Item($anchorPara).Range
    $s = $doc.Shapes.AddShape($type, $left, $top, $w, $h, $anchor)
    $s.RelativeHorizontalPosition = $relH        # 0 marge, 1 page, 2 colonne, 3 caractère
    $s.RelativeVerticalPosition = $relV          # 0 marge, 1 page, 2 paragraphe, 3 ligne
    $s.Left = $left
    $s.Top = $top
    $s.WrapFormat.Type = $wrap
    $s.WrapFormat.Side = $side                   # 0 deux côtés, 1 gauche, 2 droite, 3 le plus large
    if ($dist -ge 0) {
        $s.WrapFormat.DistanceLeft = $dist; $s.WrapFormat.DistanceRight = $dist
        $s.WrapFormat.DistanceTop = $dist; $s.WrapFormat.DistanceBottom = $dist
    }
    $s.Fill.ForeColor.RGB = 0xC0D0E0
    return $s
}

function Save($doc, [string]$name) {
    $path = [IO.Path]::Combine($Dest, $name)
    $doc.SaveAs2($path, 16)
    $doc.Close([ref]0)
    Write-Host "  $name"
}

$word = New-Object -ComObject Word.Application
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $cases = @(
        @{ n = 'carre-gauche.docx';  a = 2; t = 1; l = 0;   tp = 40; w = 120; h = 100; rh = 0; rv = 2; wr = 0; s = 0 },
        @{ n = 'carre-droite.docx';  a = 2; t = 1; l = 330; tp = 40; w = 120; h = 100; rh = 0; rv = 2; wr = 0; s = 0 },
        @{ n = 'carre-milieu.docx';  a = 2; t = 1; l = 170; tp = 40; w = 120; h = 100; rh = 0; rv = 2; wr = 0; s = 0 },
        @{ n = 'cote-gauche.docx';   a = 2; t = 1; l = 170; tp = 40; w = 120; h = 100; rh = 0; rv = 2; wr = 0; s = 1 },
        @{ n = 'cote-droit.docx';    a = 2; t = 1; l = 170; tp = 40; w = 120; h = 100; rh = 0; rv = 2; wr = 0; s = 2 },
        @{ n = 'cote-large.docx';    a = 2; t = 1; l = 120; tp = 40; w = 120; h = 100; rh = 0; rv = 2; wr = 0; s = 3 },
        @{ n = 'haut-bas.docx';      a = 2; t = 1; l = 100; tp = 30; w = 150; h = 80;  rh = 0; rv = 2; wr = 4; s = 0 },
        @{ n = 'rapproche.docx';     a = 2; t = 9; l = 0;   tp = 20; w = 160; h = 160; rh = 0; rv = 2; wr = 1; s = 0 },
        @{ n = 'ancre-bas.docx';     a = 5; t = 1; l = 0;   tp = 0;  w = 140; h = 90;  rh = 0; rv = 0; wr = 0; s = 0 },
        @{ n = 'page.docx';          a = 1; t = 1; l = 420; tp = 30; w = 150; h = 120; rh = 1; rv = 1; wr = 0; s = 0 },
        @{ n = 'devant.docx';        a = 2; t = 1; l = 100; tp = 20; w = 150; h = 80;  rh = 0; rv = 2; wr = 3; s = 0 },
        @{ n = 'derriere.docx';      a = 2; t = 1; l = 100; tp = 20; w = 150; h = 80;  rh = 0; rv = 2; wr = 5; s = 0 }
    )
    foreach ($c in $cases) {
        $doc = $word.Documents.Add()
        Body $doc 7
        $null = AddShape $doc $c.a $c.t $c.l $c.tp $c.w $c.h $c.rh $c.rv $c.wr $c.s
        Save $doc $c.n
    }
    # retrait : paragraphes en retrait gauche de 36 pt, objet contre la marge
    $doc = $word.Documents.Add()
    Body $doc 7 36
    $null = AddShape $doc 2 1 0 30 110 90 0 2 0 0
    Save $doc 'retrait.docx'
    # distances : 18 pt tout autour
    $doc = $word.Documents.Add()
    Body $doc 7
    $null = AddShape $doc 2 1 0 30 110 90 0 2 0 0 18
    Save $doc 'distances.docx'
    # zone de texte : deux paragraphes dans une zone de 160 × 90, habillage carré à droite
    $doc = $word.Documents.Add()
    Body $doc 7
    $anchor = $doc.Paragraphs.Item(2).Range
    $tb = $doc.Shapes.AddTextbox(1, 290, 30, 160, 90, $anchor)
    $tb.RelativeHorizontalPosition = 0; $tb.RelativeVerticalPosition = 2
    $tb.Left = 290; $tb.Top = 30
    $tb.WrapFormat.Type = 0
    $tb.TextFrame.TextRange.Text = (Lorem 12 3) + "`r" + (Lorem 8 5)
    Save $doc 'zone-texte.docx'
}
finally {
    $word.Quit()
}
