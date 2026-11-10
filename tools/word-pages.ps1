# Relève la pagination de Word (comparaison avec celle d'OpenWrite, banc --layout --word-pages).
# Usage : powershell.exe -ExecutionPolicy Bypass -File tools\word-pages.ps1 <dossier .docx> <dossier JSON> [délai par fichier en s]
#
# - Les .docx sont ouverts en LECTURE SEULE et refermés sans enregistrer ; une instance de Word par document.
# - Sortie : <nom>.json = liste { page, text } : les 60 premiers caractères du corps de chaque page (sans en-têtes,
#   pieds ni zones de texte), d'après Range.GoTo(wdGoToPage) ;
#   <nom>.paragraphes.json = liste { text, page, y, x } : page, position verticale (points depuis le haut de la page,
#   haut de la première ligne) et horizontale (premier caractère) de chaque paragraphe du corps, cellules de tableau comprises.
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
        $doc.Repaginate()
        $pages = [int]$doc.ComputeStatistics(2)                          # wdStatisticPages
        $end = $doc.Content.End
        for ($i = 1; $i -le $pages; $i++) {
            $r = $doc.GoTo(1, 1, $i)                                     # wdGoToPage, wdGoToAbsolute
            $start = $r.Start
            $r.SetRange($start, [Math]::Min($start + 60, $end))
            $list.Add([pscustomobject]@{ page = $i; text = [string]$r.Text })
        }
        $paras = New-Object System.Collections.Generic.List[object]
        foreach ($p in $doc.Paragraphs) {
            $r = $p.Range
            $c = $r.Characters.Item(1)
            $paras.Add([pscustomobject]@{
                text = [string]$r.Text.Substring(0, [Math]::Min(40, $r.Text.Length))
                page = [int]$c.Information(3)                              # wdActiveEndPageNumber
                y = [double]$c.Information(6)                              # wdVerticalPositionRelativeToPage
                x = [double]$c.Information(5)                              # wdHorizontalPositionRelativeToPage
            })
        }
        $pjson = [string](ConvertTo-Json -InputObject $paras.ToArray() -Depth 3)
        [IO.File]::WriteAllText([IO.Path]::Combine($Dest, [IO.Path]::GetFileNameWithoutExtension($One) + '.paragraphes.json'), $pjson, [Text.UTF8Encoding]::new($false))
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
