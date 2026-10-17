# Runs scripts\install.ps1 (or uninstall.ps1) OUTSIDE any MSIX package the calling shell may run in, so the registry value and
# the copied .xll land where Excel can see them (see Invoke-Unpackaged in FeTest.ps1), then verifies the result.
# Usage: Install-Unpackaged.ps1 [-Uninstall] [-Rebuild]
param([switch]$Uninstall, [switch]$Rebuild)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\FeTest.ps1"

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$out = Join-Path $PSScriptRoot 'out'
New-Item -ItemType Directory -Force $out | Out-Null
$log = Join-Path $out 'install-log.txt'
$script = if ($Uninstall) { 'uninstall.ps1' } else { 'install.ps1' }
$extra = if ($Rebuild -and -not $Uninstall) { ' -Rebuild' } else { '' }

Invoke-Unpackaged ("cmd.exe /c powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$root\scripts\$script`"$extra > `"$log`" 2>&1") -TimeoutMs 600000
Get-Content $log

# Verification from outside the package: the OPEN values and the installed file.
$check = Join-Path $out 'install-check.txt'
Invoke-Unpackaged ("cmd.exe /c (reg query HKCU\Software\Microsoft\Office\16.0\Excel\Options /f OPEN /v & dir `"%LOCALAPPDATA%\FormulaExplorer`") > `"$check`" 2>&1")
Get-Content $check
