# Relève la mise en forme que Word applique à chaque paragraphe du corps (comparaison avec le résolveur de styles d'OpenWrite).
# Usage : powershell.exe -ExecutionPolicy Bypass -File tools\word-format.ps1 <dossier .docx> <dossier JSON> [délai par fichier en s]
#
# - Les .docx sont ouverts en LECTURE SEULE et refermés sans enregistrer ; une instance de Word par document.
# - Sortie : <nom>.json = liste { text, align, left, right, first, before, after, beforeAuto, afterAuto, line, rule,
#   font, size, bold, italic, underline, color } (points ; premier caractère pour la police) ;
#   comparaison : OpenWrite.Fidelity --compare-formats <JSON OpenWrite> <JSON Word>.
param(
    [Parameter(Mandatory)][string]$Source,
    [Parameter(Mandatory)][string]$Dest,
    [int]$Timeout = 300,
    [string]$One
)
$ErrorActionPreference = 'Stop'

if ($One) {
    $word = New-Object -ComObject Word.Application
    try {
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $doc = $word.Documents.Open($One, $false, $true, $false, 'x-pas-de-mot-de-passe')
        $list = New-Object System.Collections.Generic.List[object]
        foreach ($p in $doc.Paragraphs) {
            $f = $p.Format
            $r = $p.Range
            $c = $r.Characters.Item(1).Font
            $col = [int]$c.Color
            if ($col -eq -16777216) { $color = 'auto' }
            elseif ($col -ge 0 -and $col -le 0xFFFFFF) { $color = '{0:X2}{1:X2}{2:X2}' -f ($col -band 0xFF), (($col -shr 8) -band 0xFF), (($col -shr 16) -band 0xFF) }
            else { $color = 'theme' }      # couleur de thème : la valeur COM n'est pas fiable (comparée à w:val par OpenWrite.Fidelity)
            $list.Add([pscustomobject]@{
                text = [string]$r.Text; align = [int]$f.Alignment
                rowEnd = [bool]$r.Characters.Item(1).Information(31)          # wdAtEndOfRowMarker : marque de fin de ligne de tableau
                left = [double]$f.LeftIndent; right = [double]$f.RightIndent; first = [double]$f.FirstLineIndent
                before = [double]$f.SpaceBefore; after = [double]$f.SpaceAfter
                beforeAuto = [bool]($f.SpaceBeforeAuto -ne 0); afterAuto = [bool]($f.SpaceAfterAuto -ne 0)
                line = [double]$f.LineSpacing; rule = [int]$f.LineSpacingRule
                font = [string]$c.Name; size = [double]$c.Size
                bold = [bool]($c.Bold -eq -1); italic = [bool]($c.Italic -eq -1); underline = [bool]($c.Underline -ne 0)
                color = $color
            })
        }
        $doc.Close([ref]0)
        # PowerShell 5.1 : convertir un tableau explicite, puis forcer une chaîne (sinon « types des arguments »)
        $json = [string](ConvertTo-Json -InputObject $list.ToArray() -Depth 3)
        [IO.File]::WriteAllText([IO.Path]::Combine($Dest, [IO.Path]::GetFileNameWithoutExtension($One) + '.json'), $json, [Text.UTF8Encoding]::new($false))
    } finally {
        $word.Quit([ref]0)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    exit 0
}

New-Item -ItemType Directory -Force -Path $Dest | Out-Null
foreach ($f in Get-ChildItem -Path $Source -Filter *.docx | Sort-Object Name) {
    Write-Host ("… " + $f.Name)
    $avant = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
    $p = Start-Process powershell.exe -PassThru -WindowStyle Hidden `
        -ArgumentList ('-NoProfile -ExecutionPolicy Bypass -File "{0}" -Source "{1}" -Dest "{2}" -One "{3}"' -f $PSCommandPath, $Source, $Dest, $f.FullName)
    if (-not $p.WaitForExit($Timeout * 1000)) {
        # délai dépassé : seule l'instance de Word lancée pour ce fichier est arrêtée
        Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $avant -notcontains $_.Id } | Stop-Process -Force -ErrorAction SilentlyContinue
        Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
        Write-Host "   délai dépassé"
    }
}
