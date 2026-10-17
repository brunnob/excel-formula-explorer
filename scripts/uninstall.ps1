<#
.SYNOPSIS
  Removes Formula Explorer for the current user: the Excel startup entry and the installed add-in file.
.DESCRIPTION
  Use -RemoveUserData to also delete the settings and log (%LOCALAPPDATA%\FormulaExplorer\settings.json, log.txt).
  Excel must not be running.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$RemoveUserData
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'ExcelOptions.ps1')

Assert-ExcelNotRunning

$target = Get-InstalledXllPath
$directory = Get-InstallDirectory

if ($PSCmdlet.ShouldProcess($script:DefaultOptionsKey, "Remove the startup entry of $target")) {
    if (Remove-OpenEntry $script:DefaultOptionsKey $target) {
        Write-Host 'Removed the Excel startup entry.'
    } else {
        Write-Host 'No Excel startup entry found.'
    }
}

if (Test-Path -LiteralPath $target) {
    if ($PSCmdlet.ShouldProcess($target, 'Delete the add-in file')) {
        Remove-Item -LiteralPath $target -Force
        Write-Host "Deleted $target"
    }
}

if ($RemoveUserData) {
    if ((Test-Path -LiteralPath $directory) -and $PSCmdlet.ShouldProcess($directory, 'Delete the settings and log')) {
        Remove-Item -LiteralPath $directory -Recurse -Force
        Write-Host "Deleted $directory"
    }
} elseif ((Test-Path -LiteralPath $directory) -and -not (Get-ChildItem -LiteralPath $directory -Force)) {
    if ($PSCmdlet.ShouldProcess($directory, 'Delete the empty folder')) { Remove-Item -LiteralPath $directory -Force }
}

Write-Host 'Done.'
