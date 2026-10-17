# Starts Excel the normal way (outside the MSIX package, like the user does), checks that the INSTALLED add-in loaded
# (real log line naming %LOCALAPPDATA%\FormulaExplorer\FormulaExplorer64.xll), then closes that
# Excel instance only (it refuses to run when another Excel is already open).
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\FeTest.ps1"
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Excel is already running; close it first.' }
$out = Join-Path $PSScriptRoot 'out'
New-Item -ItemType Directory -Force $out | Out-Null
$since = Get-Date
Invoke-Unpackaged 'cmd.exe /c start "" excel.exe /e'
$p = $null
for ($i = 0; $i -lt 60 -and -not $p; $i++) { Start-Sleep -Milliseconds 500; $p = Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -First 1 }
if (-not $p) { throw 'Excel did not start.' }
Start-Sleep -Seconds 12
$log = Join-Path $out 'installed-real-log.txt'
Copy-FeRealLog $log
$lines = Get-Content $log | Where-Object { $_ -match 'loaded' } | Select-Object -Last 3
"Load lines: "; $lines
$ok = $lines | Where-Object { $_ -like '*\AppData\Local\FormulaExplorer\FormulaExplorer64.xll*' }
"INSTALLED ADD-IN LOADED: $([bool]$ok)"
# Close only the instance we started (no workbook is dirty: /e opens no workbook).
$p.CloseMainWindow() | Out-Null
if (-not $p.WaitForExit(15000)) { Stop-Process -Id $p.Id -Force }
"Excel closed."
