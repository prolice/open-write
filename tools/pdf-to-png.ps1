# Convertit les pages d'un PDF (référence exportée par Word) en images PNG, avec le lecteur PDF de Windows (WinRT).
# Usage : powershell.exe -ExecutionPolicy Bypass -File tools\pdf-to-png.ps1 <fichier.pdf> <dossier de sortie> [pages, ex. 1,3] [largeur px]
param(
    [Parameter(Mandatory)][string]$Pdf,
    [Parameter(Mandatory)][string]$Dest,
    [string]$Pages = '',
    [int]$Width = 1000
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
[void][Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
[void][Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
[void][Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]

$asTask = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' } | Select-Object -First 1
$asTaskAction = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' } | Select-Object -First 1
function Await($op, [Type]$type) { $t = $asTask.MakeGenericMethod($type).Invoke($null, @($op)); $t.Wait(); $t.Result }
function AwaitAction($op) { $t = $asTaskAction.Invoke($null, @($op)); $t.Wait() }

New-Item -ItemType Directory -Force -Path $Dest | Out-Null
$file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync((Resolve-Path $Pdf).Path)) ([Windows.Storage.StorageFile])
$doc = Await ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
$wanted = if ($Pages) { $Pages.Split(',') | ForEach-Object { [int]$_ } } else { 1..$doc.PageCount }
$name = [IO.Path]::GetFileNameWithoutExtension($Pdf)
foreach ($n in $wanted) {
    if ($n -lt 1 -or $n -gt $doc.PageCount) { continue }
    $page = $doc.GetPage($n - 1)
    $opts = New-Object Windows.Data.Pdf.PdfPageRenderOptions
    $opts.DestinationWidth = $Width
    $opts.DestinationHeight = [uint32]($Width * $page.Size.Height / $page.Size.Width)
    $stream = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
    AwaitAction ($page.RenderToStreamAsync($stream, $opts))
    $net = [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($stream.GetInputStreamAt(0))
    $out = [IO.Path]::Combine($Dest, "$name-$n.png")
    $fs = [IO.File]::Create($out)
    $net.CopyTo($fs)
    $fs.Close()
    $page.Dispose()
    Write-Host "  $out"
}
