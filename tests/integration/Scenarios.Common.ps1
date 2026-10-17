# Environment management shared by the scenarios: workbook builders, Excel lifecycle, per-scenario reset and cleanup.
$script:fe = $null
$script:LogBase = 0

function Build-TestWorkbooks {
    $p = Join-Path $script:OutDir 'fe-test.xlsx'
    & "$PSScriptRoot\New-TestWorkbook.ps1" -Path $p | Out-Null
    # Second workbook for the SDI / cross-workbook scenario.
    $p2 = Join-Path $script:OutDir 'fe-test2.xlsx'
    $xl = New-Object -ComObject Excel.Application
    try {
        $xl.Visible = $false; $xl.DisplayAlerts = $false
        $wb = $xl.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1); $ws.Name = 'Calc'
        [void][System.__ComObject].InvokeMember('Value2', [Reflection.BindingFlags]::SetProperty, $null, $ws.Range('C5'), @(42))
        [void][System.__ComObject].InvokeMember('Formula', [Reflection.BindingFlags]::SetProperty, $null, $ws.Range('C6'), @('=C5*2'))
        if (Test-Path $p2) { Remove-Item $p2 -Force }
        $wb.SaveAs($p2, 51); $wb.Close($false)
    }
    finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
    $p
}

function Get-ExcelAlive {
    if (-not $script:fe -or -not $script:fe.Excel) { return $false }
    try { $null = $script:fe.Excel.Hwnd; return $true } catch { return $false }
}

function Ensure-Excel {
    if (Get-ExcelAlive) { return }
    if ($script:fe) {
        $p = Get-Process -Id $script:fe.Pid -ErrorAction SilentlyContinue
        if ($p -and $p.ProcessName -eq 'EXCEL') { Stop-Process -Id $script:fe.Pid -Force; Start-Sleep -Milliseconds 500 }
    }
    $args0 = @{ OutDir = $script:OutDir }
    if ($script:XllPath) { $args0.Xll = $script:XllPath }
    $script:fe = Start-FeExcel @args0
    Write-Host "  (Excel started, pid $($script:fe.Pid))" -ForegroundColor DarkGray
}

# Closes any message box/dialog of our Excel process (#32770) so the next scenario starts from a clean state.
function Dismiss-Dialogs {
    if (-not $script:fe) { return }
    for ($i = 0; $i -lt 5; $i++) {
        $found = $false
        foreach ($line in [FeWin2]::TopLevel([uint32]$script:fe.Pid)) {
            $parts = $line.Split('|')
            if ($parts[1] -eq '#32770') { $found = $true; [void][FeWin2]::PostMessage([IntPtr][int64]$parts[0], 0x10, [IntPtr]::Zero, [IntPtr]::Zero) }
        }
        if (-not $found) { break }
        Start-Sleep -Milliseconds 300
    }
}

function Open-Workbook([string]$SourcePath, [string]$Name) {
    $copy = Join-Path $script:OutDir $Name
    if (Test-Path $copy) { Remove-Item $copy -Force }
    Copy-Item $SourcePath $copy -Force
    $wb = $script:fe.Excel.Workbooks.Open($copy)
    $wb
}

function Close-AllWorkbooks {
    $xl = $script:fe.Excel
    foreach ($w in @($xl.Workbooks)) { try { $w.Close($false) } catch { } }
}

# Fresh Excel state for a scenario: no explorer, default settings (first-open placement), the test workbook open.
function Reset-Env([switch]$NoWorkbook) {
    Ensure-Excel
    $fe = $script:fe
    Dismiss-Dialogs
    try { Close-Explorer $fe } catch { }
    try { Close-AllWorkbooks } catch { Write-Host "  (close workbooks: $($_.Exception.Message))" -ForegroundColor DarkGray }
    Clear-RealSettings
    $fe.Workbook = $null
    if (-not $NoWorkbook) {
        $wb = Open-Workbook (Join-Path $script:OutDir 'fe-test.xlsx') 'wb-fe-test.xlsx'
        $fe.Workbook = $wb
        $wb.Worksheets.Item('Val_FO').Activate()
        $fe.Excel.ActiveWindow.Zoom = 100
        $fe.Excel.ActiveWindow.ScrollRow = 1; $fe.Excel.ActiveWindow.ScrollColumn = 1
    }
    # FeWin.Foreground restores (SW_RESTORE) a maximized window: bring Excel forward first, then maximize it.
    [FeWin]::Foreground([IntPtr]$fe.Excel.Hwnd)
    Start-Sleep -Milliseconds 300
    if (-not $NoWorkbook) { Ensure-Maximized }
    Start-Sleep -Milliseconds 300
    $script:LogBase = @(Get-RealLogLines $script:OutDir).Count
}

# Maximizes the Excel window that shows the active workbook (SDI: every workbook has its own XLMAIN) and waits for it.
function Ensure-Maximized {
    $xl = $script:fe.Excel
    for ($i = 0; $i -lt 8; $i++) {
        try {
            $h = [IntPtr]$xl.ActiveWindow.Hwnd
            $r = [FeWin2]::Rect($h)
            if ($xl.WindowState -eq -4137 -and ($r[2] - $r[0]) -gt 3000) { return }
            $xl.WindowState = -4137
        } catch { }
        Start-Sleep -Milliseconds 350
    }
}

function Cleanup-AfterScenario {
    try { Dismiss-Dialogs } catch { }
    if (Get-ExcelAlive) { try { Close-Explorer $script:fe } catch { } }
}

# Hard check: nothing in the add-in's real log with level ERROR since Reset-Env.
function Check-LogClean([string]$What = 'add-in log') {
    $lines = Get-RealLogLines $script:OutDir
    $new = @($lines | Select-Object -Skip $script:LogBase | Where-Object { $_ -match '\[ERROR\]' })
    Check "no ERROR entries in the $What" ($new.Count -eq 0) ($new | Select-Object -First 3)
}

function Get-Sheet($sheet) { $script:fe.Excel.ActiveWorkbook.Worksheets.Item($sheet) }

# Value of a node as the tooltip shows it, computed independently: the expression is put into a scratch cell through
# Range.Formula (always English syntax, unlike Worksheet.Evaluate which parses with the pt-BR LCID from PowerShell).
$script:ErrorNames = @{ -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!'; -2146826265 = '#REF!'; -2146826259 = '#NAME?'; -2146826252 = '#NUM!'; -2146826246 = '#N/A' }
function Format-Eval($v) {
    if ($null -eq $v) { return '[Blank]' }
    if ($v -is [int] -and $script:ErrorNames.ContainsKey([int]$v)) { return $script:ErrorNames[[int]$v] }
    if ($v -is [double] -or $v -is [int]) {
        $d = [double]$v
        if ($d -eq 0) { return '0' }
        return $d.ToString('G15', [Globalization.CultureInfo]::InvariantCulture)
    }
    if ($v -is [bool]) { if ($v) { return 'TRUE' } else { return 'FALSE' } }
    if ($v -is [string]) { return '"' + $v + '"' }
    "$v"
}
function Eval-Expected($sheet, [string]$expr) {
    $ws = Get-Sheet $sheet
    $cell = $ws.Range('AZ900')
    try {
        [void][System.__ComObject].InvokeMember('Formula', [Reflection.BindingFlags]::SetProperty, $null, $cell, @('=' + $expr))
        $v = $cell.Value2
    }
    finally { [void]$cell.Clear() }
    Format-Eval $v
}
