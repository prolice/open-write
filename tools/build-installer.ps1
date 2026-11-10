# Construit l'executable autonome puis l'installeur MSI d'OpenWrite.
# Usage (PowerShell, depuis la racine du depot) : powershell -ExecutionPolicy Bypass -File tools\build-installer.ps1
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$dotnet = Join-Path $env:LOCALAPPDATA 'Microsoft\dotnet\dotnet.exe'
if (-not (Test-Path $dotnet)) { $dotnet = 'dotnet' }
Push-Location $root
try {
  & $dotnet test tests\OpenWrite.Tests -nologo -v q
  if ($LASTEXITCODE -ne 0) { throw "Les tests ont echoue" }
  & $dotnet publish src\OpenWrite.App\OpenWrite.App.csproj -c Release -r win-x64 --self-contained true `
      -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true `
      -p:DebugType=none -o dist\OpenWrite -nologo -v q
  if ($LASTEXITCODE -ne 0) { throw "Echec de la publication" }
  & $dotnet build installer\OpenWrite.Installer.wixproj -c Release -nologo -v q
  if ($LASTEXITCODE -ne 0) { throw "Echec de la construction de l'installeur" }
  Copy-Item installer\bin\x64\Release\fr-FR\*.msi dist\ -ErrorAction SilentlyContinue
  Copy-Item installer\bin\x64\Release\*.msi dist\ -ErrorAction SilentlyContinue
  Get-ChildItem dist\*.msi | ForEach-Object { "Installeur : $($_.FullName) ($([math]::Round($_.Length / 1MB, 1)) Mo)" }
} finally { Pop-Location }
