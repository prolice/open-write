# Relève ce que Word voit des champs, cases à cocher et signets de chaque .docx (comparaison avec OpenWrite).
# Usage : powershell.exe -ExecutionPolicy Bypass -File tools\word-inspect.ps1 <dossier .docx> <dossier JSON> [délai par fichier en s]
#
# - Les .docx sont ouverts en LECTURE SEULE et refermés sans enregistrer ; les champs ne sont jamais mis à jour.
# - Une instance de Word par document (processus enfant), arrêtée après le délai si elle bloque.
# - Sortie : <nom>.json = { fields: [{code, result}], checkboxes: [{name, value}], bookmarks: [noms, masqués compris] }
#   même forme que tests/OpenWrite.Fidelity (dossier modele/) ; comparaison : OpenWrite.Fidelity --compare-word.
param(
    [Parameter(Mandatory)][string]$Source,
    [Parameter(Mandatory)][string]$Dest,
    [int]$Timeout = 90,
    [string]$One
)
$ErrorActionPreference = 'Stop'

if ($One) {
    $word = New-Object -ComObject Word.Application
    try {
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $doc = $word.Documents.Open($One, $false, $true, $false, 'x-pas-de-mot-de-passe')
        $fields = @()
        # toutes les parties de texte : corps, notes, en-têtes et pieds (chaque section), commentaires…
        $seen = @{}
        foreach ($story in $doc.StoryRanges) {
            $r = $story
            while ($r -ne $null) {
                $key = "$($r.StoryType):$($r.Start):$($r.End):$($r.Text.GetHashCode())"
                if (-not $seen.ContainsKey($key)) {
                    $seen[$key] = 1
                    foreach ($f in $r.Fields) { $fields += [pscustomobject]@{ code = [string]$f.Code.Text; result = [string]$f.Result.Text } }
                }
                $r = $r.NextStoryRange
            }
        }
        $boxes = @()
        foreach ($ff in $doc.FormFields) {
            if ($ff.Type -eq 71) { $boxes += [pscustomobject]@{ name = [string]$ff.Name; value = [bool]$ff.CheckBox.Value } }   # wdFieldFormCheckBox
        }
        $doc.Bookmarks.ShowHidden = $true
        $marks = @()
        foreach ($b in $doc.Bookmarks) { $marks += [string]$b.Name }
        $doc.Close([ref]0)
        $out = [pscustomobject]@{ fields = @($fields); checkboxes = @($boxes); bookmarks = @($marks) }
        [IO.File]::WriteAllText([IO.Path]::Combine($Dest, [IO.Path]::GetFileNameWithoutExtension($One) + '.json'),
            (ConvertTo-Json $out -Depth 4), [Text.UTF8Encoding]::new($false))
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
