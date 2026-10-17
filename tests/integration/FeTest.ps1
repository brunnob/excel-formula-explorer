# Helpers for driving Excel + the Formula Explorer test hooks from PowerShell (dot-source this file).
#   . "$PSScriptRoot\FeTest.ps1"
#   $fe = Start-FeExcel -Xll <path> -Workbook <path>
#   Invoke-Fe $fe 'FE_Test_Open'; $s = Wait-FeState $fe { $_.session.root -ne $null }
#   Save-FeShot $fe 'name'; Stop-FeExcel $fe
# Notes:
# - Late-bound COM from PowerShell uses the user's LCID (pt-BR here): never pass formulas with ',' or '.' through it.
# - Only the Excel instance started here is ever closed/killed (tracked by PID).
. "$PSScriptRoot\WinUtil.ps1"

$script:FeRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

function Enable-FeTestHooks([bool]$On = $true) {
    # Must be written outside the MSIX virtualization (see Invoke-Unpackaged) or Excel will not see it.
    if ($On) { Invoke-Unpackaged 'reg.exe add HKCU\Software\FormulaExplorer /v EnableTestHooks /t REG_DWORD /d 1 /f' }
    else { Invoke-Unpackaged 'reg.exe delete HKCU\Software\FormulaExplorer /v EnableTestHooks /f' }
}

function Start-FeExcel {
    param(
        [string]$Xll = (Join-Path $script:FeRoot 'src\FormulaExplorer.AddIn\bin\Debug\net10.0-windows\FormulaExplorer64.xll'),
        [string]$Workbook,
        [string]$OutDir = (Join-Path $PSScriptRoot 'out')
    )
    New-Item -ItemType Directory -Force $OutDir | Out-Null
    Enable-FeTestHooks $true
    $xl = New-Object -ComObject Excel.Application
    $pid0 = [uint32]0
    [void][FeWin]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$pid0)
    $xl.DisplayAlerts = $false
    $xl.Visible = $true
    $xl.WindowState = -4137   # xlMaximized
    $wb = $null
    if ($Workbook) {
        $copy = Join-Path $OutDir ('wb-' + [IO.Path]::GetFileName($Workbook))
        Copy-Item $Workbook $copy -Force
        $wb = $xl.Workbooks.Open($copy)
    }
    if (-not $xl.RegisterXLL((Resolve-Path $Xll).Path)) { throw "RegisterXLL failed: $Xll" }
    [FeWin]::Foreground([IntPtr]$xl.Hwnd)
    Start-Sleep -Milliseconds 500
    [pscustomobject]@{ Excel = $xl; Pid = $pid0; Workbook = $wb; OutDir = $OutDir; Shot = 0 }
}

function Stop-FeExcel($fe) {
    try { [void](Invoke-Fe $fe 'FE_Test_Close') } catch { }
    try { foreach ($w in @($fe.Excel.Workbooks)) { $w.Close($false) } } catch { }
    try { $fe.Excel.Quit() } catch { }
    if ($fe.Excel) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($fe.Excel) }
    $fe.Excel = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    for ($i = 0; $i -lt 30; $i++) {
        if (-not (Get-Process -Id $fe.Pid -ErrorAction SilentlyContinue)) { break }
        Start-Sleep -Milliseconds 300
    }
    $p = Get-Process -Id $fe.Pid -ErrorAction SilentlyContinue
    if ($p -and $p.ProcessName -eq 'EXCEL') { Stop-Process -Id $fe.Pid -Force }
    Enable-FeTestHooks $false
}

function Invoke-Fe($fe, [string]$Hook, $Arg) {
    $r = if ($PSBoundParameters.ContainsKey('Arg')) { $fe.Excel.Run($Hook, $Arg) } else { $fe.Excel.Run($Hook) }
    if ($r -is [string] -and $r -ne 'ok') { Write-Warning "$Hook -> $r" }
    $r
}

function Get-FeState($fe) {
    $path = Join-Path $fe.OutDir 'state.json'
    if (Test-Path $path) { Remove-Item $path -Force }
    [void]$fe.Excel.Run('FE_Test_Dump', $path)
    Get-Content $path -Raw -Encoding UTF8 | ConvertFrom-Json
}

# Polls the state until $Condition (a scriptblock using $_) is true, pumping time for Excel's message loop.
function Wait-FeState($fe, [scriptblock]$Condition, [int]$TimeoutMs = 5000) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    do {
        $s = Get-FeState $fe
        $ok = $false
        try { $ok = [bool]($s | ForEach-Object $Condition) } catch { }
        if ($ok) { return $s }
        Start-Sleep -Milliseconds 100
    } while ($sw.ElapsedMilliseconds -lt $TimeoutMs)
    Write-Warning "Wait-FeState timed out after $TimeoutMs ms"
    $s
}

function Select-FeCell($fe, [string]$Sheet, [string]$Address) {
    $ws = $fe.Excel.ActiveWorkbook.Worksheets.Item($Sheet)
    $ws.Activate()
    $ws.Range($Address).Select()
}

function Save-FeShot($fe, [string]$Name, [int[]]$Rect) {
    $fe.Shot++
    $file = Join-Path $fe.OutDir ('{0:00}-{1}.png' -f $fe.Shot, $Name)
    if ($Rect) { & "$PSScriptRoot\Screenshot.ps1" -Path $file -Rect $Rect | Out-Null }
    else { & "$PSScriptRoot\Screenshot.ps1" -Path $file | Out-Null }
    $file
}

# Index of the Nth occurrence (1-based) of $Needle in $Text, plus $Offset.
function Find-FeIndex([string]$Text, [string]$Needle, [int]$Occurrence = 1, [int]$Offset = 0) {
    $i = -1
    for ($n = 0; $n -lt $Occurrence; $n++) { $i = $Text.IndexOf($Needle, $i + 1, [StringComparison]::Ordinal); if ($i -lt 0) { throw "'$Needle' #$Occurrence not found in '$Text'" } }
    $i + $Offset
}

# When the test shell runs inside an MSIX-packaged app (some terminals and IDEs), processes started from it see a VIRTUALIZED %LOCALAPPDATA%/%APPDATA% and
# HKCU (copy-on-write), while Excel (started by COM/DCOM) sees the real ones. Invoke-Unpackaged runs a command line
# outside the package (through WMI's Win32_Process.Create) and waits for it, so registry/AppData changes are real.
function Invoke-Unpackaged([string]$CommandLine, [int]$TimeoutMs = 60000) {
    $r = Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{ CommandLine = $CommandLine }
    if ($r.ReturnValue -ne 0) { throw "Win32_Process.Create failed ($($r.ReturnValue)): $CommandLine" }
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while (Get-Process -Id $r.ProcessId -ErrorAction SilentlyContinue) {
        if ($sw.ElapsedMilliseconds -gt $TimeoutMs) { throw "Timed out: $CommandLine" }
        Start-Sleep -Milliseconds 100
    }
}

# Copies the REAL add-in log (as Excel writes it) to $Destination (use a path outside AppData, e.g. under C:\dev).
function Copy-FeRealLog([string]$Destination) {
    Invoke-Unpackaged "cmd.exe /c copy /y `"%LOCALAPPDATA%\FormulaExplorer\log.txt`" `"$Destination`""
}
