<#
.SYNOPSIS
  Builds the whole solution (Release by default) and prints the path of the packed 64-bit add-in.
.DESCRIPTION
  The packed .xll (FormulaExplorer64-packed.xll) is a single file that contains the add-in, Core and the Excel interop
  assembly; it is the only file that has to be copied to another machine.
#>
[CmdletBinding()]
param(
    [ValidateSet('Release', 'Debug')]
    [string]$Configuration = 'Release'
)
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$solution = Join-Path $root 'FormulaExplorer.slnx'

# Out-Host keeps dotnet's output off the pipeline, so the only output object is the path below.
dotnet build $solution -c $Configuration -nologo | Out-Host
if ($LASTEXITCODE -ne 0) { throw "dotnet build failed with exit code $LASTEXITCODE." }

$xll = Join-Path $root "src\FormulaExplorer.AddIn\bin\$Configuration\net10.0-windows\publish\FormulaExplorer64-packed.xll"
if (-not (Test-Path -LiteralPath $xll)) { throw "Build succeeded but the packed add-in was not found: $xll" }

Write-Host "Packed add-in: $xll"
$xll
