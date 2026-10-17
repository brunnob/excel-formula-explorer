# Helpers (dot-sourced by install.ps1 and uninstall.ps1) for Excel's "open at startup" registry entries.
#
# Excel opens every file named by the values OPEN, OPEN1, OPEN2, ... under
#   HKCU\Software\Microsoft\Office\16.0\Excel\Options
# when it starts. The value data of an add-in is  /R "C:\path\add-in.xll"  (/R = open read-only).
# Excel reads the sequence in order, so the numbering is kept contiguous.

$script:DefaultOptionsKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'

function Get-OpenValueNames([string]$KeyPath) {
    # OPEN, OPEN1, OPEN2, ... sorted by their numeric suffix.
    if (-not (Test-Path -LiteralPath $KeyPath)) { return @() }
    $key = Get-Item -LiteralPath $KeyPath
    @($key.GetValueNames() | Where-Object { $_ -match '^OPEN\d*$' } |
        Sort-Object { if ($_ -eq 'OPEN') { -1 } else { [int]$_.Substring(4) } })
}

function ConvertTo-OpenValueData([string]$XllPath) {
    '/R "' + $XllPath + '"'
}

function Test-OpenValueMatchesXll([string]$Data, [string]$XllPath) {
    # Tolerates quotes/spacing differences, compares the path case-insensitively.
    $path = ($Data -replace '^\s*/R\s+', '').Trim().Trim('"')
    [string]::Equals($path, $XllPath, [System.StringComparison]::OrdinalIgnoreCase)
}

function Find-OpenEntry([string]$KeyPath, [string]$XllPath) {
    foreach ($name in Get-OpenValueNames $KeyPath) {
        $data = [string](Get-ItemProperty -LiteralPath $KeyPath -Name $name).$name
        if (Test-OpenValueMatchesXll $data $XllPath) { return $name }
    }
    $null
}

function Add-OpenEntry([string]$KeyPath, [string]$XllPath) {
    # Idempotent. Returns the name of the value that now points at the add-in.
    $existing = Find-OpenEntry $KeyPath $XllPath
    if ($existing) { return $existing }

    if (-not (Test-Path -LiteralPath $KeyPath)) { New-Item -Path $KeyPath -Force | Out-Null }
    $used = Get-OpenValueNames $KeyPath
    $name = 'OPEN'
    $n = 0
    while ($used -contains $name) { $n++; $name = "OPEN$n" }
    New-ItemProperty -LiteralPath $KeyPath -Name $name -Value (ConvertTo-OpenValueData $XllPath) -PropertyType String | Out-Null
    $name
}

function Remove-OpenEntry([string]$KeyPath, [string]$XllPath) {
    # Removes the entry for the add-in and renumbers the others so the sequence stays contiguous.
    # Returns $true when an entry was removed.
    if (-not (Find-OpenEntry $KeyPath $XllPath)) { return $false }

    $keep = @()
    foreach ($name in Get-OpenValueNames $KeyPath) {
        $data = [string](Get-ItemProperty -LiteralPath $KeyPath -Name $name).$name
        if (-not (Test-OpenValueMatchesXll $data $XllPath)) { $keep += $data }
    }
    foreach ($name in Get-OpenValueNames $KeyPath) { Remove-ItemProperty -LiteralPath $KeyPath -Name $name }
    for ($i = 0; $i -lt $keep.Count; $i++) {
        $name = if ($i -eq 0) { 'OPEN' } else { "OPEN$i" }
        New-ItemProperty -LiteralPath $KeyPath -Name $name -Value $keep[$i] -PropertyType String | Out-Null
    }
    $true
}

function Get-InstallDirectory { Join-Path $env:LOCALAPPDATA 'FormulaExplorer' }
function Get-InstalledXllPath { Join-Path (Get-InstallDirectory) 'FormulaExplorer64.xll' }

function Assert-ExcelNotRunning {
    if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
        throw 'Excel is running. Close all Excel windows (check Task Manager for EXCEL.EXE) and run this script again.'
    }
}
