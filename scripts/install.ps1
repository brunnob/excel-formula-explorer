<#
.SYNOPSIS
  Installs Formula Explorer for the current user: copies the packed 64-bit add-in to
  %LOCALAPPDATA%\FormulaExplorer\FormulaExplorer64.xll and makes Excel load it at startup.
.DESCRIPTION
  Builds the solution first when the packed add-in is missing or older than the sources (use -Rebuild to force it,
  -NoBuild to never build). Startup loading is registered as an OPEN value under
  HKCU\Software\Microsoft\Office\16.0\Excel\Options (the next free OPEN, OPEN1, OPEN2...), idempotently.
  Excel must not be running.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$Rebuild,
    [switch]$NoBuild
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'ExcelOptions.ps1')

$root = Split-Path -Parent $PSScriptRoot
Assert-ExcelNotRunning

# 1. The packed add-in: build if needed.
$packed = Join-Path $root 'src\FormulaExplorer.AddIn\bin\Release\net10.0-windows\publish\FormulaExplorer64-packed.xll'
$needsBuild = $Rebuild.IsPresent
if (-not $needsBuild -and -not $NoBuild) {
    if (-not (Test-Path -LiteralPath $packed)) {
        $needsBuild = $true
    } else {
        $built = (Get-Item -LiteralPath $packed).LastWriteTimeUtc
        $newer = Get-ChildItem -Path (Join-Path $root 'src') -Recurse -File -Include *.cs, *.csproj, *.png, *.props |
            Where-Object { $_.FullName -notmatch '\\(bin|obj)\\' -and $_.LastWriteTimeUtc -gt $built } |
            Select-Object -First 1
        $needsBuild = $null -ne $newer
    }
}
if ($needsBuild) {
    if ($NoBuild) { throw '-NoBuild and -Rebuild cannot be combined.' }
    # The build rewrites bin/obj, so -WhatIf must not run it.
    if ($PSCmdlet.ShouldProcess($packed, 'Build the solution (Release)')) {
        $packed = & (Join-Path $PSScriptRoot 'build.ps1') -Configuration Release | Select-Object -Last 1
    }
}
if (-not (Test-Path -LiteralPath $packed)) {
    # Under -WhatIf the build may have been skipped, so a missing output is expected there.
    if ($WhatIfPreference) { Write-Warning "Packed add-in not found: $packed (it would be built first)." }
    else { throw "Packed add-in not found: $packed (run scripts\build.ps1)." }
}

# 2. Copy ONLY the packed 64-bit add-in (it already contains Core and the interop assembly).
$target = Get-InstalledXllPath
if ($PSCmdlet.ShouldProcess($target, 'Copy the packed add-in')) {
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $target) | Out-Null
    Copy-Item -LiteralPath $packed -Destination $target -Force
    Write-Host "Copied add-in to $target"
}

# 3. Register it to open when Excel starts.
if ($PSCmdlet.ShouldProcess("$script:DefaultOptionsKey", "Register $target to load at Excel startup")) {
    $name = Add-OpenEntry $script:DefaultOptionsKey $target
    Write-Host "Registered at Excel startup ($name under $script:DefaultOptionsKey)."
}

# 4. Friendly warnings about the requirements (they do not stop the installation).
try {
    $runtimes = & dotnet --list-runtimes 2>$null
    if (-not ($runtimes -match 'Microsoft\.WindowsDesktop\.App 10\.')) {
        Write-Warning '.NET 10 Desktop Runtime (x64) was not found. Install it from https://dotnet.microsoft.com/download/dotnet/10.0 before starting Excel.'
    }
} catch {
    Write-Warning '.NET 10 Desktop Runtime (x64) could not be verified. Install it from https://dotnet.microsoft.com/download/dotnet/10.0 if Excel reports a load error.'
}
$platform = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue).Platform
if ($platform -and $platform -ne 'x64') {
    Write-Warning "Excel reports platform '$platform'; Formula Explorer needs 64-bit Excel."
}

Write-Host 'Done. Start Excel: the "Studio" tab appears after "View".'
