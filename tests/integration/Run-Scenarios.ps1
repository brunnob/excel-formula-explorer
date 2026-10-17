# Formula Explorer integration scenario suite (real Excel 365, real mouse/keyboard, screenshots).
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\integration\Run-Scenarios.ps1 [-Only S1,S4] [-Xll path] [-Clean]
# Prints PASS/FAIL/PARTIAL per scenario with the observed values, saves screenshots under tests\integration\out\scenarios
# and a machine-readable summary (scenarios-result.json). Always cleans up: closes the explorer, quits ONLY the Excel
# instance started here, restores the user's real settings.json and removes the test-hook registry gate.
param(
    [string]$Xll,
    [string[]]$Only,
    [switch]$Clean,   # delete the files of a previous run from OutDir first (a full run's evidence must contain no stale file)
    [string]$OutDir = (Join-Path $PSScriptRoot 'out\scenarios')
)
$ErrorActionPreference = 'Stop'
# Excel's IDispatch honors the caller's LCID: keep COM text parsing in English (Range.Formula is always English anyway).
[Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('en-US')
. "$PSScriptRoot\ScenarioLib.ps1"
. "$PSScriptRoot\ScenarioOverlayLib.ps1"
. "$PSScriptRoot\Scenarios.Common.ps1"

New-Item -ItemType Directory -Force $OutDir | Out-Null
if ($Clean) { Get-ChildItem $OutDir -File | Remove-Item -Force }
$script:OutDir = $OutDir
$script:XllPath = $Xll

# Scenario files: each defines functions S1, S2, ... (dot-sourced so they share the helpers).
foreach ($f in Get-ChildItem $PSScriptRoot -Filter 'Scenarios.*.ps1' | Where-Object { $_.Name -ne 'Scenarios.Common.ps1' } | Sort-Object Name) { . $f.FullName }

# Every function named S<number>[<letter>] is a scenario (S12, S12b, S13 ...), run in numeric then alphabetical order.
$all = @(Get-ChildItem function:\ | ForEach-Object { if ($_.Name -match '^S(\d+)([a-z]?)$') { [pscustomobject]@{ Name = $_.Name; N = [int]$Matches[1]; L = $Matches[2] } } } | Sort-Object N, L | ForEach-Object Name)
$run = if ($Only) { $Only | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim().ToUpper() } } else { $all }

$startedAt = Get-Date
try {
    Backup-RealSettings $OutDir
    $kept = Backup-RealLog $OutDir   # the developer's real log is kept (copied) before it is emptied
    Write-Host "Real add-in log backed up to $kept" -ForegroundColor DarkGray
    Clear-RealLog   # unit-test runs append to the same file; start from an empty log
    $script:WorkbookPath = Build-TestWorkbooks
    foreach ($id in $run) {
        if (-not (Get-Command $id -ErrorAction SilentlyContinue)) { Write-Host "No scenario function $id" -ForegroundColor Yellow; continue }
        try {
            & $id
        }
        catch {
            if (-not $script:Cur -or $script:Cur.id -ne $id) { Begin-Scenario $id '(crashed before start)' }
            $script:Cur.error = "$($_.Exception.Message)`n$($_.ScriptStackTrace)"
            Write-Host "  EXCEPTION: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
            try { Save-FeShot $script:fe "$id-exception" | Out-Null } catch { }
        }
        finally {
            if ($script:Cur -and $script:Cur.id -eq $id -and -not ($script:Results | Where-Object { $_.id -eq $id })) { End-Scenario }
            Cleanup-AfterScenario
        }
    }
}
finally {
    try { Close-Explorer $script:fe } catch { }
    try { if ($script:fe) { Stop-FeExcel $script:fe } } catch { Write-Host "Stop-FeExcel failed: $($_.Exception.Message)" }
    try { Restore-RealSettings $OutDir } catch { Write-Host "Restore settings failed: $($_.Exception.Message)" }
    try { Enable-FeTestHooks $false } catch { }
    Remove-Item (Join-Path $OutDir 'settings.current.json') -ErrorAction SilentlyContinue
}

Write-Host ""
Write-Host "================ SUMMARY ($([int]((Get-Date) - $startedAt).TotalSeconds) s) ================"
foreach ($r in $script:Results) {
    $color = switch ($r.status) { 'PASS' { 'Green' } 'PARTIAL' { 'Yellow' } default { 'Red' } }
    $failed = @($r.checks | Where-Object { -not $_.ok })
    Write-Host ("{0,-4} {1,-8} {2}  ({3}/{4} checks)" -f $r.id, $r.status, $r.title, ($r.checks.Count - $failed.Count), $r.checks.Count) -ForegroundColor $color
    foreach ($c in $failed) { Write-Host ("       - {0}: {1}" -f $c.name, $c.observed) -ForegroundColor $color }
    if ($r.error) { Write-Host "       ! $($r.error)" -ForegroundColor Red }
}
$script:Results | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $OutDir 'scenarios-result.json') -Encoding UTF8
$bad = @($script:Results | Where-Object { $_.status -eq 'FAIL' }).Count
exit ([int]($bad -gt 0))
