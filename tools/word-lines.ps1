# Relève les lignes de texte telles que Word les a mises en page (objets flottants, habillage : lot 7) : pour chaque
# page, chaque ligne du corps avec sa position (haut, gauche ; points depuis le coin de la page), sa largeur et le
# début de son texte. Lecture seule (Pages / Rectangles / Lines du volet en mode Page).
# Usage : powershell.exe -ExecutionPolicy Bypass -File tools\word-lines.ps1 <dossier .docx> <dossier JSON>
#   sortie : <nom>.lignes.json = liste { page, top, left, width, text }
param([Parameter(Mandatory)][string]$Source, [Parameter(Mandatory)][string]$Dest)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $Dest | Out-Null
$word = New-Object -ComObject Word.Application
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    foreach ($f in Get-ChildItem -Path $Source -Filter *.docx) {
        $doc = $word.Documents.Open($f.FullName, $false, $true, $false)
        try {
            $win = $doc.ActiveWindow
            $win.View.Type = 3                                         # wdPrintView
            $doc.Repaginate()
            $list = New-Object System.Collections.Generic.List[object]
            $pane = $win.ActivePane
            for ($p = 1; $p -le $pane.Pages.Count; $p++) {
                $page = $pane.Pages.Item($p)
                foreach ($rect in $page.Rectangles) {
                    if ($rect.RectangleType -ne 0) { continue }       # wdTextRectangle (corps)
                    foreach ($line in $rect.Lines) {
                        $text = ''
                        try { $text = [string]$line.Range.Text } catch { }
                        try { if ($line.Range.StoryType -ne 1) { continue } } catch { }   # corps seulement (pas les zones de texte)
                        $list.Add([pscustomobject]@{
                            page = $p; top = [double]$line.Top; left = [double]$line.Left; width = [double]$line.Width
                            text = $text.Substring(0, [Math]::Min(30, $text.Length))
                        })
                    }
                }
            }
            $out = [IO.Path]::Combine($Dest, [IO.Path]::GetFileNameWithoutExtension($f.Name) + '.lignes.json')
            [IO.File]::WriteAllText($out, (ConvertTo-Json -InputObject $list.ToArray() -Depth 3), [Text.Encoding]::UTF8)
            Write-Host "  $($f.Name) : $($list.Count) lignes"
        }
        finally { $doc.Close([ref]0) }
    }
}
finally { $word.Quit() }
