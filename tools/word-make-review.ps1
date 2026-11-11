# Fabrique, avec Word (COM), un document de calibration du lot 8 au contenu neutre : revisions.docx
#   modifications suivies de deux relecteurs (texte inséré, mot supprimé, mise en forme, paragraphe inséré,
#   paragraphes fusionnés), commentaires avec réponse, fil résolu.
# Le nom d'utilisateur de Word est remplacé le temps du script (« Relecteur A », « Relecteur B ») puis remis.
# Écrit aussi revisions.txt : ce que Word voit (révisions : type, auteur, texte ; commentaires : auteur, texte, résolu).
# Usage : powershell.exe -ExecutionPolicy Bypass -File tools\word-make-review.ps1 <dossier de sortie>
param([Parameter(Mandatory)][string]$Dest)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $Dest | Out-Null
$path = [IO.Path]::Combine($Dest, 'revisions.docx')

$word = New-Object -ComObject Word.Application
$oldName = $word.UserName
$oldInitials = $word.UserInitials
# « Toujours utiliser ces valeurs, même connecté à Office » : sinon Word signe avec le compte Microsoft
$oldLocal = $word.Options.UseLocalUserInfo
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.Options.UseLocalUserInfo = $true
    $doc = $word.Documents.Add()
    $texts = @(
        'Premier paragraphe : alpha beta gamma delta epsilon.',
        'Second paragraphe : zeta eta theta iota kappa.',
        'Troisième paragraphe : lambda mu nu xi omicron.',
        'Quatrième paragraphe : pi rho sigma tau.',
        'Cinquième paragraphe : upsilon phi chi psi omega.'
    )
    foreach ($t in $texts) { $r = $doc.Content; $r.Collapse(0); $r.InsertAfter($t); $r.InsertParagraphAfter() }

    $doc.TrackRevisions = $true
    $word.UserName = 'Relecteur A'; $word.UserInitials = 'RA'
    # insertion au milieu du premier paragraphe
    $p1 = $doc.Paragraphs(1).Range
    $at = $doc.Range($p1.Start + 20, $p1.Start + 20); $at.InsertAfter('NOUVEAU ')
    # suppression d'un mot du deuxième
    $p2 = $doc.Paragraphs(2).Range
    $f = $p2.Duplicate; $null = $f.Find.Execute('theta '); if ($f.Find.Found) { $f.Delete() | Out-Null }
    # mise en forme dans le troisième
    $p3 = $doc.Paragraphs(3).Range
    $g = $p3.Duplicate; $null = $g.Find.Execute('omicron'); if ($g.Find.Found) { $g.Bold = $true }

    $word.UserName = 'Relecteur B'; $word.UserInitials = 'RB'
    # paragraphe inséré après le quatrième
    $p4 = $doc.Paragraphs(4).Range
    $e = $doc.Range($p4.End - 1, $p4.End - 1); $e.InsertParagraphAfter(); $e2 = $doc.Paragraphs(5).Range; $e2.InsertBefore('Paragraphe ajouté par B.')
    # fusion : marque du dernier paragraphe d'origine supprimée
    $last = $doc.Paragraphs(6).Range
    $m = $doc.Range($last.End - 1, $last.End); $m.Delete() | Out-Null

    # commentaires (pas suivis) : un fil avec réponse, un fil résolu
    $doc.TrackRevisions = $false
    $word.UserName = 'Relecteur A'; $word.UserInitials = 'RA'
    $c1 = $doc.Comments.Add($doc.Paragraphs(2).Range.Words(1), 'Faut-il garder ce mot ?')
    $word.UserName = 'Relecteur B'; $word.UserInitials = 'RB'
    $null = $c1.Replies.Add($c1.Scope, 'Oui, gardons-le.')
    $c2 = $doc.Comments.Add($doc.Paragraphs(4).Range.Words(1), 'Vérifier la numérotation.')
    $c2.Done = $true

    # ce que Word voit
    $lines = @()
    foreach ($rev in $doc.Revisions) { $lines += "révision $($rev.Type) | $($rev.Author) | $($rev.Range.Text)" }
    foreach ($c in $doc.Comments) { $lines += "commentaire | $($c.Author) | $($c.Range.Text) | résolu=$($c.Done) | parent=$(if ($c.Ancestor) { $c.Ancestor.Range.Text } else { '' })" }
    [IO.File]::WriteAllLines([IO.Path]::Combine($Dest, 'revisions.txt'), $lines, (New-Object Text.UTF8Encoding $false))

    $doc.SaveAs2($path, 16)
    $doc.Close([ref]0)
    Write-Host "  revisions.docx ($($lines.Count) éléments)"
}
finally {
    $word.UserName = $oldName
    $word.UserInitials = $oldInitials
    $word.Options.UseLocalUserInfo = $oldLocal
    $word.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
