<#
.SYNOPSIS
  Builds dist\FormulaExplorerSetup.exe: a single self-contained installer (runs without .NET installed) that embeds
  the packed 64-bit add-in. Run it on any machine with 64-bit Excel; FormulaExplorerSetup.exe /uninstall removes it
  (it is also listed in Settings > Apps > Installed apps).
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$xll = & (Join-Path $PSScriptRoot 'build.ps1') -Configuration Release | Select-Object -Last 1
if (-not (Test-Path -LiteralPath $xll)) { throw "Packed add-in not found: $xll" }
$dist = Join-Path $root 'dist'
dotnet publish (Join-Path $root 'tools\FormulaExplorer.Setup\FormulaExplorer.Setup.csproj') -c Release -o $dist --nologo
if ($LASTEXITCODE -ne 0) { throw 'dotnet publish failed.' }
Get-ChildItem $dist -Exclude 'FormulaExplorerSetup.exe' | Remove-Item -Recurse -Force
$exe = Join-Path $dist 'FormulaExplorerSetup.exe'
"{0} ({1:N1} MB)" -f $exe, ((Get-Item $exe).Length / 1MB)
