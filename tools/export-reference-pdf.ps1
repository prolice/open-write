# Produit les PDF de référence du corpus avec Microsoft Word (poste sous licence).
# Usage : powershell.exe -ExecutionPolicy Bypass -File tools\export-reference-pdf.ps1 <dossier .docx> <dossier PDF> [délai par fichier en s]
#
# - Les .docx sont ouverts en LECTURE SEULE et refermés sans enregistrer.
# - Chaque document est traité dans SA PROPRE instance de Word (processus enfant) : si un document bloque
#   (boîte de dialogue invisible…), le script ne ferme que l'instance Word qu'il a lancée, après le délai.
# - Les champs ne sont jamais mis à jour (USERADDRESS interroge le carnet d'adresses Outlook et ouvre
#   une boîte de dialogue) : le PDF montre les résultats de champ enregistrés par Word.
param(
    [Parameter(Mandatory)][string]$Source,
    [Parameter(Mandatory)][string]$Dest,
    [int]$Timeout = 120,
    [string]$One          # usage interne : traite un seul fichier (processus enfant)
)
$ErrorActionPreference = 'Stop'

if ($One) {
    $pidFile = Join-Path $Dest ('.' + [IO.Path]::GetFileName($One) + '.pid')
    $logFile = $pidFile -replace '\.pid$', '.log'
    function Etape($m) { Add-Content -Path $logFile -Value ((Get-Date -Format 'HH:mm:ss.fff') + ' ' + $m) }
    Etape 'lancement de Word'
    $avant = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
    $word = New-Object -ComObject Word.Application
    $wp = Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $avant -notcontains $_.Id } | Select-Object -First 1
    if ($wp) { Set-Content -Path $pidFile -Value $wp.Id }
    try {
        Etape 'Word lancé'; $word.Visible = $true
        $word.DisplayAlerts = 0                    # wdAlertsNone
        # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument)
        $doc = $word.Documents.Open($One, $false, $true, $false, 'x-pas-de-mot-de-passe')
        Etape 'ouvert'; $doc.Fields.Locked = $true                 # le document n'est pas enregistré : verrou sans effet sur le fichier
        Etape 'champs verrouillés'; $pages = $doc.ComputeStatistics(2)         # wdStatisticPages
        $words = $doc.ComputeStatistics(0)         # wdStatisticWords
        Etape 'statistiques'; $compat = $doc.CompatibilityMode
        [string]$pdf = [IO.Path]::Combine($Dest, [IO.Path]::GetFileNameWithoutExtension($One) + '.pdf')   # chaîne .NET pure : un PSObject (Join-Path) fige ExportAsFixedFormat
        # ExportAsFixedFormat(Fichier, 17=PDF, Ouvrir, 0=impression, 0=tout, De, À, 0=contenu, Propriétés, IRM, 1=signets des titres)
        Etape 'export'; $doc.ExportAsFixedFormat($pdf, 17, $false, 0, 0, 1, 1, 0, $true, $true, 1)
        Etape 'exporté'; $doc.Close([ref]0)                         # wdDoNotSaveChanges ([ref] obligatoire en PowerShell, sinon exception et Word reste ouvert)
        Set-Content -Path ($pidFile -replace '\.pid$', '.res') -Value "ok;$pages;$words;$compat"
    } finally {
        Etape 'fermeture de Word'; $word.Quit([ref]0); Etape 'fini'
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
        Remove-Item $pidFile -ErrorAction SilentlyContinue
    }
    exit 0
}

New-Item -ItemType Directory -Force -Path $Dest | Out-Null
$journal = @()
foreach ($f in Get-ChildItem -Path $Source -Filter *.docx | Sort-Object Name) {
    Write-Host ("… " + $f.Name)
    $out = Join-Path $Dest ('.' + $f.Name + '.out')
    $res = Join-Path $Dest ('.' + $f.Name + '.res')
    # Arguments entre guillemets : les noms de fichiers contiennent des espaces
    $p = Start-Process powershell.exe -PassThru -WindowStyle Minimized -RedirectStandardError ($out + '.err') `
        -ArgumentList ('-NoProfile -ExecutionPolicy Bypass -File "{0}" -Source "{1}" -Dest "{2}" -One "{3}"' -f $PSCommandPath, $Source, $Dest, $f.FullName)
    $t = [Diagnostics.Stopwatch]::StartNew()
    if ($p.WaitForExit($Timeout * 1000)) {
        $r = (Get-Content $res -ErrorAction SilentlyContinue | Where-Object { $_ -like 'ok;*' } | Select-Object -Last 1)
        if ($r) { $c = $r.Split(';'); $etat = 'ok'; $pages = $c[1]; $mots = $c[2]; $compat = $c[3] }
        else { $etat = 'échec : ' + ((Get-Content ($out + '.err') -ErrorAction SilentlyContinue) -join ' ').Trim(); $pages = ''; $mots = ''; $compat = '' }
    } else {
        # Délai dépassé : on ne ferme QUE l'instance de Word lancée pour ce fichier, puis son processus parent.
        $pidFile = Join-Path $Dest ('.' + $f.Name + '.pid')
        if (Test-Path $pidFile) { Stop-Process -Id ([int](Get-Content $pidFile)) -Force -ErrorAction SilentlyContinue; Remove-Item $pidFile }
        Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
        $etat = "délai dépassé ($Timeout s)"; $pages = ''; $mots = ''; $compat = ''
    }
    Remove-Item $res, ($out + '.err') -ErrorAction SilentlyContinue
    $l = [pscustomobject]@{ Fichier = $f.Name; Pages = $pages; Mots = $mots; Compat = $compat; Etat = $etat; Secondes = [math]::Round($t.Elapsed.TotalSeconds, 1) }
    $journal += $l
    Write-Host ("   " + $l.Etat + "  " + $l.Pages + " p.  " + $l.Secondes + " s")
    $journal | Export-Csv -Path (Join-Path $Dest 'references.csv') -NoTypeInformation -Encoding UTF8
}
$journal | Format-Table -AutoSize | Out-String -Width 200
