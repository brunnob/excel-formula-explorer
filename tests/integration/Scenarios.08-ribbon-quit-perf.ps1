# S17 (ribbon through UI Automation), S18 (Excel quit with the explorer open), S19 (performance).
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes

function Find-UiaElement($root, [string]$Name, $ControlType) {
    $c1 = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::NameProperty, $Name)
    if ($ControlType) {
        $c2 = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::ControlTypeProperty, $ControlType)
        $cond = [System.Windows.Automation.AndCondition]::new($c1, $c2)
    }
    else { $cond = $c1 }
    $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $cond)
}

function Get-UiaRect($el) {
    $r = $el.Current.BoundingRectangle
    @([int]$r.Left, [int]$r.Top, [int]$r.Right, [int]$r.Bottom)
}

function S17 {
    Begin-Scenario 'S17' 'Ribbon: Studio tab and Formula Explorer button (UI Automation), Invoke opens / re-activates, icon and supertip'
    Reset-Env
    $fe = $script:fe
    $xl = $fe.Excel
    $xlmain = [IntPtr]$xl.ActiveWindow.Hwnd
    $xl.ActiveWorkbook.Worksheets.Item('Val_FO').Range('K33').Select() | Out-Null
    Start-Sleep -Milliseconds 400
    $rootEl = [System.Windows.Automation.AutomationElement]::FromHandle($xlmain)
    $tab = $null
    Wait-Until { $script:tab = Find-UiaElement $rootEl 'Studio' ([System.Windows.Automation.ControlType]::TabItem); $null -ne $script:tab } 6000 | Out-Null
    $tab = $script:tab
    Check 'the Studio tab exists in the ribbon' ($null -ne $tab) ''
    if (-not $tab) { Check-LogClean; End-Scenario; return }
    $tabRect = Get-UiaRect $tab
    $xlRect = Get-Rect $xlmain
    # position: after View and before Help/Developer
    $tabs = @()
    foreach ($t in $rootEl.FindAll([System.Windows.Automation.TreeScope]::Descendants, ([System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::TabItem)))) { $tabs += @{ name = $t.Current.Name; left = [int]$t.Current.BoundingRectangle.Left } }
    $ordered = ($tabs | Where-Object { $_.left -gt 0 } | Sort-Object { $_.left } | ForEach-Object { $_.name }) -join ' | '
    $iView = [array]::IndexOf(($ordered -split ' \| '), 'View'); $iStudio = [array]::IndexOf(($ordered -split ' \| '), 'Studio')
    Check 'Studio sits right after View' ($iStudio -eq $iView + 1 -and $iView -ge 0) $ordered
    $sel = $tab.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
    $sel.Select()
    Start-Sleep -Milliseconds 900
    $btn = $null
    Wait-Until { $script:btn = Find-UiaElement $rootEl 'Formula Explorer' ([System.Windows.Automation.ControlType]::Button); $null -ne $script:btn } 4000 | Out-Null
    $btn = $script:btn
    Check 'the "Formula Explorer" button exists on the Studio tab' ($null -ne $btn) ''
    if (-not $btn) { Check-LogClean; End-Scenario; return }
    Check 'the button is enabled' ($btn.Current.IsEnabled) $btn.Current.IsEnabled
    $r = Get-UiaRect $btn
    $nav = Find-UiaElement $rootEl 'Navigation' $null
    Check 'the button is in a group named Navigation' ($null -ne $nav) ''
    # screenshot of the ribbon strip for the icon / layout
    Move-Mouse 100 1700
    Start-Sleep -Milliseconds 300
    $strip = @([math]::Max(0, $r[0] - 200), ($r[1] - 70), 700, 260)
    Save-FeShot $fe 'S17-ribbon' $strip | Out-Null
    $iconBmp = [FeWin2]::Capture($r[0] + 4, $r[1] + 2, [math]::Max(8, $r[2] - $r[0] - 8), [int](($r[3] - $r[1]) * 0.6))
    $blue = [FeWin2]::CountColor($iconBmp, 0, 0, $iconBmp.Width, $iconBmp.Height, 0x1E6FD0, 90)
    $iconBmp.Dispose()
    Check 'the large button shows a blue globe icon (blue pixels present)' ($blue -gt 150) "blue-ish pixels=$blue" -Level soft

    # supertip: hover the real pointer over the button for ~1.5 s
    $cx = [int](($r[0] + $r[2]) / 2); $cy = [int](($r[1] + $r[3]) / 2)
    Move-Mouse ($cx - 3) $cy; Start-Sleep -Milliseconds 100; Move-Mouse $cx $cy
    Start-Sleep -Milliseconds 1800
    Save-FeShot $fe 'S17-supertip' @([math]::Max(0, $r[0] - 250), ($r[1] - 40), 900, 420) | Out-Null
    # (the supertip text is read from the screenshot by eye; also try UI Automation's tooltip window)
    $tip = $null
    foreach ($line in [FeWin2]::TopLevel([uint32]$fe.Pid)) { $p = $line.Split('|'); if ($p[1] -like '*ToolTip*' -or $p[1] -eq 'NetUIHWND' -or $p[1] -like 'XLToolTip*') { $tip = $line } }
    Note "tooltip-like windows: $tip"
    # Invoke
    $inv = $btn.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
    $inv.Invoke()
    $s = Wait-FeState $fe { $_.isOpen -and $_.view -and $_.session.root -and -not $_.view.isLoading -and $_.explorerHwnd -ne 0 } 5000
    Check 'Invoke opens the explorer on the active cell (K33)' ($s.isOpen -and $s.session.root.address -eq '$K$33' -and $s.view.displayText -like '=1/*') @{ open = $s.isOpen; root = $s.session.root.address }
    $h1 = $s.explorerHwnd
    # a second Invoke only activates the existing window
    Click-At ($xlRect[0] + 800) ($xlRect[1] + 14)     # deactivate it
    Start-Sleep -Milliseconds 500
    Check 'precondition: the explorer is not the foreground window' (-not (Test-ExplorerForeground $fe $s)) (Describe-Fg)
    $btn = Find-UiaElement $rootEl 'Formula Explorer' ([System.Windows.Automation.ControlType]::Button)
    $btn.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
    Start-Sleep -Milliseconds 900
    $s2 = Get-FeState $fe
    Check 'second Invoke: same window (no second explorer)' ($s2.isOpen -and $s2.explorerHwnd -eq $h1) @{ first = $h1; second = $s2.explorerHwnd }
    $count = @([FeWin2]::TopLevel([uint32]$fe.Pid) | Where-Object { $_ -like '*|Formula Explorer|*' }).Count
    Check 'exactly one "Formula Explorer" window exists' ($count -eq 1) $count
    Check 'second Invoke activates the existing window (it is the foreground window)' (Get-Foreground-Timed $s2 1200) (Describe-Fg)
    # back to Home tab
    $homeTab = Find-UiaElement $rootEl 'Home' ([System.Windows.Automation.ControlType]::TabItem)
    if ($homeTab) { $homeTab.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select() }
    Check-LogClean
    End-Scenario
}

function S18 {
    Begin-Scenario 'S18' 'Excel quit with the explorer open (window closed like a user would): no hang, no crash dialog'
    Reset-Env
    $fe = $script:fe
    $s = Open-Explorer $fe 'Val_FO' 'K36'
    [FeWin]::Foreground([IntPtr][int64]$s.explorerHwnd); Start-Sleep -Milliseconds 300
    # a gold link into a hidden sheet: the session has something to restore at the end
    Click-Node $fe $s 'Hidden_FO!B2' 1 2
    $r = Wait-Timed $fe { $null -ne $_.view.goldStart } 3000
    Check 'precondition: Hidden_FO was unhidden by the click' ($fe.Excel.ActiveWorkbook.Worksheets.Item('Hidden_FO').Visible -eq -1) ''
    $pid0 = [uint32]$fe.Pid
    $hwnds = @([FeWin2]::TopLevel($pid0) | Where-Object { $_.Split('|')[1] -eq 'XLMAIN' } | ForEach-Object { [int64]$_.Split('|')[0] })
    Check 'the Excel main window and the explorer window exist before quitting' ($hwnds.Count -ge 1 -and ([FeWin2]::TopLevel($pid0) | Where-Object { $_ -like '*|Formula Explorer|*' }).Count -eq 1) $hwnds.Count
    # A well-behaved client lets go of Excel (references held by this script would keep EXCEL.EXE alive after the UI closes).
    foreach ($w in @($fe.Excel.Workbooks)) { $w.Saved = $true }
    $fe.Workbook = $null
    $xlApp = $fe.Excel
    $fe.Excel = $null
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xlApp)
    $xlApp = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    $sw = [Diagnostics.Stopwatch]::StartNew()
    foreach ($h in $hwnds) { [void][FeWin2]::PostMessage([IntPtr]$h, 0x10, [IntPtr]::Zero, [IntPtr]::Zero) }   # WM_CLOSE on every XLMAIN, as the X button does
    $exited = $false; $dialogSeen = $null
    while ($sw.ElapsedMilliseconds -lt 12000) {
        if (-not (Get-Process -Id $pid0 -ErrorAction SilentlyContinue)) { $exited = $true; break }
        foreach ($line in [FeWin2]::TopLevel($pid0)) { $c = $line.Split('|'); if ($c[1] -eq '#32770') { $dialogSeen = $line } }
        Start-Sleep -Milliseconds 100
    }
    $ms = $sw.ElapsedMilliseconds
    Check 'the Excel process exits within 10 s of closing its window with the explorer open' ($exited -and $ms -le 10000) "exited=$exited after $ms ms"
    Check 'no crash / message dialog appeared while quitting' ($null -eq $dialogSeen) $dialogSeen
    if (-not $exited) {
        Note ("windows left: " + (([FeWin2]::TopLevel($pid0)) -join ' ; '))
        $p = Get-Process -Id $pid0 -ErrorAction SilentlyContinue
        if ($p -and $p.ProcessName -eq 'EXCEL') { Stop-Process -Id $pid0 -Force }
    }
    try {
        $since = (Get-Date).AddMinutes(-2)
        $ev = @(Get-WinEvent -FilterHashtable @{ LogName = 'Application'; ProviderName = 'Application Error', 'Application Hang'; StartTime = $since } -ErrorAction SilentlyContinue | Where-Object { $_.Message -match 'EXCEL' })
        Check 'no Windows crash/hang event for EXCEL.EXE was recorded' ($ev.Count -eq 0) ($ev | Select-Object -First 1 | ForEach-Object { $_.Message.Substring(0, [math]::Min(160, $_.Message.Length)) })
    }
    catch { Note "event log not readable: $($_.Exception.Message)" }
    $lines = Get-RealLogLines $script:OutDir
    $new = @($lines | Select-Object -Skip $script:LogBase | Where-Object { $_ -match '\[ERROR\]' })
    Check 'no ERROR entries in the add-in log during the shutdown' ($new.Count -eq 0) ($new | Select-Object -First 3)
    $last = @($lines | Select-Object -Last 4)
    Note "log tail: $($last -join ' // ')"
    # the shared Excel is gone: the next scenario starts a fresh one
    End-Scenario
}

function Measure-CpuPercent([int]$Pid0, [int]$Ms) {
    $p = Get-Process -Id $Pid0
    $t0 = $p.TotalProcessorTime.TotalMilliseconds
    $sw = [Diagnostics.Stopwatch]::StartNew()
    Start-Sleep -Milliseconds $Ms
    $p.Refresh()
    $t1 = $p.TotalProcessorTime.TotalMilliseconds
    [math]::Round(($t1 - $t0) / $sw.ElapsedMilliseconds * 100, 1)
}

function S19 {
    Begin-Scenario 'S19' 'Performance: 36-reference formula (load, hover latency), idle CPU with the explorer open'
    Reset-Env
    $fe = $script:fe
    $ws = $fe.Excel.ActiveSheet
    # 36 references: B40..B75 (plain cells with values)
    foreach ($r in 40..75) { [void][System.__ComObject].InvokeMember('Value2', [Reflection.BindingFlags]::SetProperty, $null, $ws.Range("B$r"), @($r)) }
    $refs = (40..75 | ForEach-Object { "B$_" }) -join '+'
    [void][System.__ComObject].InvokeMember('Formula', [Reflection.BindingFlags]::SetProperty, $null, $ws.Range('T2'), @("=$refs+SUM(B40:B75)*Hundred"))
    Select-FeCell $fe 'Val_FO' 'T2'
    Start-Sleep -Milliseconds 500
    # idle baseline without the explorer
    Move-Mouse 100 1900
    $cpuClosed = Measure-CpuPercent $fe.Pid 5000
    # open and time the load
    $sw = [Diagnostics.Stopwatch]::StartNew()
    [void](Invoke-Fe $fe 'FE_Test_Open')
    $s = Wait-FeState $fe { $_.isOpen -and $_.view -and $_.session.root -and -not $_.view.isLoading -and (Get-Overlay $_).Count -gt 0 } 8000
    $openMs = $sw.ElapsedMilliseconds
    $pal = @(0x0000FF, 0x008000, 0x9900CC, 0x800000, 0x00CC33, 0xFF6600, 0xCC0099)
    $boxes = @(); $k = 0
    foreach ($r in 40..75) { $boxes += @{ Address = "B$r"; Rgb = $pal[$k % 7] }; $k++ }
    $boxes += @{ Address = 'B40:B75'; Rgb = $pal[$k % 7] }
    $oracle = Get-PaneOracle $fe
    $expected = @(Get-ExpectedBoxes $fe $boxes $oracle)
    Check 'a box for every precedent visible on screen (36 cells + the range; the rest is below the window)' ((Get-Overlay $s).Count -eq $expected.Count -and $expected.Count -ge 30) @{ boxes = (Get-Overlay $s).Count; expected = $expected.Count }
    Check 'open + load + first overlay paint within 2 s' ($openMs -le 2000) "$openMs ms"
    $text = $s.view.displayText
    # hover latency: the hook call is synchronous (node lookup + evaluation + highlight)
    $times = @()
    $ixs = @(); $pos = 0
    while ($true) { $i = $text.IndexOf('B', $pos); if ($i -lt 0 -or $ixs.Count -ge 36) { break }; $ixs += $i + 1; $pos = $i + 1 }
    foreach ($i in $ixs) {
        $sw2 = [Diagnostics.Stopwatch]::StartNew()
        [void](Invoke-Fe $fe 'FE_Test_Hover' $i)
        $times += $sw2.ElapsedMilliseconds
    }
    $avg = [math]::Round(($times | Measure-Object -Average).Average, 1); $max = ($times | Measure-Object -Maximum).Maximum
    Check "hover on $($ixs.Count) different references: average call time <= 60 ms and max <= 250 ms (includes the COM round trip of Application.Run)" ($avg -le 60 -and $max -le 250) "avg=$avg ms max=$max ms"
    # the whole-range reference (a range evaluation with 36 cells) and the big SUM
    $iSum = $text.IndexOf('SUM(')
    $sw2 = [Diagnostics.Stopwatch]::StartNew(); [void](Invoke-Fe $fe 'FE_Test_Hover' ($iSum + 1)); $sumMs = $sw2.ElapsedMilliseconds
    $r = Wait-Tooltip $fe 2500
    Check 'hover on SUM(B40:B75)*... node: tooltip present and fast (< 250 ms)' ($sumMs -le 250 -and $r.view.tooltipVisible) @{ ms = $sumMs; tip = $r.view.tooltipText }
    # real mouse sweep across the text: the highlight keeps up (no queue build-up)
    Start-Sleep -Milliseconds 300
    $s = Get-FeState $fe
    $lag = @()
    $keys = @($s.charPoints.PSObject.Properties.Name | ForEach-Object { [int]$_ } | Sort-Object)
    foreach ($i in ($keys[2], $keys[[int]($keys.Count / 3)], $keys[[int]($keys.Count * 2 / 3)], $keys[$keys.Count - 3])) {
        $p = Char-Point $s $i
        Move-Mouse $p.X $p.Y
        $sw3 = [Diagnostics.Stopwatch]::StartNew()
        $r = Wait-Timed $fe { $null -ne $_.view.hoverStart } 1000 10
        $lag += $sw3.ElapsedMilliseconds
    }
    Check 'real pointer: highlight reflected within 150 ms of the move (poll latency included)' (($lag | Measure-Object -Maximum).Maximum -le 150) ($lag -join ',')
    # idle CPU with the explorer open and the pointer away
    Move-Mouse 100 1900
    Start-Sleep -Milliseconds 1500
    $cpuOpen = Measure-CpuPercent $fe.Pid 5000
    Check 'idle CPU of Excel with the explorer open is below 8% of one core' ($cpuOpen -lt 8) "open=$cpuOpen% closed(baseline)=$cpuClosed%"
    Check 'the explorer adds less than 6 percentage points over the baseline' (($cpuOpen - $cpuClosed) -lt 6) "delta=$([math]::Round($cpuOpen - $cpuClosed, 1))"
    Check-LogClean
    End-Scenario
}
