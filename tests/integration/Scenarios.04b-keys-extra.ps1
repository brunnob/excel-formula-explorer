# S12b: keyboard II. Enter twice (Back two levels), Esc after Enter, Esc/Enter while the title bar is being dragged,
# Esc/Enter over a message box owned by the explorer. All keys are real (keybd_event) and go through Excel's message pump.

function Set-Formula($ws, [string]$Address, $Value) {
    $prop = if ($Value -is [string] -and $Value.StartsWith('=')) { 'Formula' } else { 'Value2' }
    [void][System.__ComObject].InvokeMember($prop, [Reflection.BindingFlags]::SetProperty, $null, $ws.Range($Address), @($Value))
}

# Window rectangle polled from Win32 only (usable while a modal loop owns the thread and COM calls into Excel are not served).
function Get-WinRect([int64]$Hwnd) { [FeWin2]::Rect([IntPtr]$Hwnd) }

# M10 = M11+1, M11 = M12+1, ... M14 = 5: a chain four formulas deep (drill-down pushes the stack three times).
function New-DrillChain($ws) {
    foreach ($row in 10..13) { Set-Formula $ws "M$row" "=M$($row + 1)+1" }
    Set-Formula $ws 'M14' 5
}

function Get-DoubleClickTimeMs {
    if (-not ('FeDblClk' -as [type])) {
        Add-Type 'using System.Runtime.InteropServices; public static class FeDblClk { [DllImport("user32.dll")] public static extern uint GetDoubleClickTime(); }'
    }
    [FeDblClk]::GetDoubleClickTime()
}

function Drill-Chain($fe, [int]$Levels) {
    for ($k = 1; $k -le $Levels; $k++) {
        $s = Get-FeState $fe
        Click-Node $fe $s "M$(10 + $k)" 1 1
        $want = '$M$' + (10 + $k)
        $r = Wait-Timed $fe { $_.session.root.address -eq $want -and -not $_.view.isLoading } 4000
        if (-not $r.ok) { return $false }
        # More than the double-click time: a quick second click on the new formula is (by design) ignored.
        Start-Sleep -Milliseconds ([int](Get-DoubleClickTimeMs) + 150)
    }
    $true
}

function S12b {
    Begin-Scenario 'S12b' 'Keyboard II: Enter twice = Back twice, Esc after Enter closes, Esc/Enter while dragging the title bar, Esc/Enter over an explorer-owned message box'
    Reset-Env
    $fe = $script:fe
    $ws = $fe.Excel.ActiveSheet
    New-DrillChain $ws

    # ---- (a) Enter twice goes Back two levels ----
    $s = Open-Explorer $fe 'Val_FO' 'M10'
    Check 'setup: M10 loaded' ($s.view.displayText -ceq '=M11+1') $s.view.displayText
    [FeWin]::Foreground([IntPtr][int64]$s.explorerHwnd); Start-Sleep -Milliseconds 300
    Check 'setup: drilled three levels (root M13, stack 3)' (Drill-Chain $fe 3) ''
    $s = Get-FeState $fe
    Check 'stack is 3 deep' ((Get-Stack $s).Count -eq 3 -and $s.session.root.address -eq '$M$13') @{ root = $s.session.root.address; stack = (Get-Stack $s).Count }
    # a real click on Back pops one level and leaves keyboard focus on the button
    $b = $s.buttons.back
    Click-At ([int](($b[0] + $b[2]) / 2)) ([int](($b[1] + $b[3]) / 2))
    $r = Wait-Timed $fe { $_.session.root.address -eq '$M$12' -and -not $_.view.isLoading } 3000
    Check 'click on Back: root M12, stack 2' ($r.ok -and (Get-Stack $r.state).Count -eq 2) $r.state.session.root.address
    Check 'the Back button has the keyboard focus' ($r.state.buttons.backFocused) $r.state.buttons.backFocused
    Start-Sleep -Milliseconds 300
    Press-Key $script:VK.Enter
    $r = Wait-Timed $fe { $_.session.root.address -eq '$M$11' -and -not $_.view.isLoading } 3000
    Check 'first Enter: Back one level (root M11, stack 1)' ($r.ok -and (Get-Stack $r.state).Count -eq 1) @{ root = $r.state.session.root.address; stack = (Get-Stack $r.state).Count }
    Press-Key $script:VK.Enter
    $r = Wait-Timed $fe { $_.session.root.address -eq '$M$10' -and -not $_.view.isLoading } 3000
    Check 'second Enter (no waiting for a key-up in between): Back again (root M10, stack 0)' ($r.ok -and (Get-Stack $r.state).Count -eq 0 -and $r.state.isOpen) @{ root = $r.state.session.root.address; stack = (Get-Stack $r.state).Count; open = $r.state.isOpen }
    Check 'Excel selection follows (M10)' ((Cell-Addr $fe) -eq 'M10') (Cell-Addr $fe)
    Save-FeShot $fe 'S12b-enter-twice' | Out-Null
    Close-Explorer $fe

    # ---- (b) Esc after Enter ----
    $s = Open-Explorer $fe 'Val_FO' 'M10'
    [FeWin]::Foreground([IntPtr][int64]$s.explorerHwnd); Start-Sleep -Milliseconds 300
    [void](Drill-Chain $fe 2)
    $s = Get-FeState $fe
    $b = $s.buttons.back
    Click-At ([int](($b[0] + $b[2]) / 2)) ([int](($b[1] + $b[3]) / 2))
    $r = Wait-Timed $fe { $_.session.root.address -eq '$M$11' -and -not $_.view.isLoading } 3000
    Press-Key $script:VK.Enter
    $r = Wait-Timed $fe { $_.session.root.address -eq '$M$10' -and -not $_.view.isLoading } 3000
    Check 'setup: Enter went back to M10, explorer open' ($r.ok -and $r.state.isOpen) $r.state.session.root.address
    Press-Key $script:VK.Esc
    $r = Wait-Timed $fe { -not $_.isOpen } 2500
    Check 'Esc right after Enter closes the explorer' $r.ok $r.state.isOpen
    Check 'the explorer window is gone' (-not [FeWin2]::IsWindow([IntPtr][int64]$s.explorerHwnd)) ''

    # ---- (c) Esc / Enter while the title bar is being dragged: the move loop owns the keys ----
    $s = Open-Explorer $fe 'Val_FO' 'K33'
    $eh = [int64]$s.explorerHwnd
    $r0 = Get-WinRect $eh
    $tx = [int](($r0[0] + $r0[2]) / 2) - 40; $ty = $r0[1] + 14
    # Esc: Windows cancels the move; the explorer must neither close nor go Back
    Move-Mouse $tx $ty; Start-Sleep -Milliseconds 150
    [FeWin2]::Down($false); Start-Sleep -Milliseconds 100
    foreach ($i in 1..8) { [FeWin2]::MoveTo($tx - 40 * $i, $ty + 25 * $i); Start-Sleep -Milliseconds 40 }
    $rMoved = Get-WinRect $eh
    Check 'precondition: the window follows the pointer during the drag' ($rMoved[0] -lt $r0[0] - 200) @{ before = $r0[0]; during = $rMoved[0] }
    Press-Key $script:VK.Esc
    Start-Sleep -Milliseconds 400
    $aliveDuring = [FeWin2]::IsWindow([IntPtr]$eh)
    [FeWin2]::Up($false); Start-Sleep -Milliseconds 500
    $r = Wait-Timed $fe { $_.isOpen } 2000
    $r1 = Get-WinRect $eh
    Check 'Esc while dragging does not close the explorer' ($aliveDuring -and $r.state.isOpen) @{ aliveDuringDrag = $aliveDuring; open = $r.state.isOpen }
    Check 'Esc cancelled the move (the window is back where it started)' ([math]::Abs($r1[0] - $r0[0]) -le 3 -and [math]::Abs($r1[1] - $r0[1]) -le 3) @{ start = @($r0[0], $r0[1]); end = @($r1[0], $r1[1]) }
    # Enter: Windows accepts the move
    $tx = [int](($r1[0] + $r1[2]) / 2) - 40; $ty = $r1[1] + 14
    Move-Mouse $tx $ty; Start-Sleep -Milliseconds 150
    [FeWin2]::Down($false); Start-Sleep -Milliseconds 100
    foreach ($i in 1..8) { [FeWin2]::MoveTo($tx - 30 * $i, $ty + 20 * $i); Start-Sleep -Milliseconds 40 }
    Press-Key $script:VK.Enter
    Start-Sleep -Milliseconds 400
    $aliveDuring = [FeWin2]::IsWindow([IntPtr]$eh)
    [FeWin2]::Up($false); Start-Sleep -Milliseconds 500
    $r = Wait-Timed $fe { $_.isOpen } 2000
    $r2 = Get-WinRect $eh
    Check 'Enter while dragging does not close the explorer' ($aliveDuring -and $r.state.isOpen) @{ aliveDuringDrag = $aliveDuring; open = $r.state.isOpen }
    $s = Get-FeState $fe
    $p = Char-Point $s (Ix $s '^')
    Move-Mouse $p.X $p.Y
    $r = Wait-Tooltip $fe 2500
    Check 'the explorer still works after the drags (hover tooltip)' ($r.view.tooltipVisible) $r.view.tooltipText
    Save-FeShot $fe 'S12b-after-drags' | Out-Null
    Close-Explorer $fe

    # ---- (d) Esc / Enter over a message box owned by the explorer ----
    $wbk = $fe.Excel.ActiveWorkbook
    [void]$wbk.Protect([Type]::Missing, $true, $false)
    $s = Open-Explorer $fe 'Val_FO' 'K36'
    foreach ($key in @('Esc', 'Enter')) {
        [FeWin]::Foreground([IntPtr][int64]$s.explorerHwnd); Start-Sleep -Milliseconds 250
        Click-Node $fe $s 'Hidden_FO!B2' 1 2
        $dlg = $null
        Wait-Until { $script:dlg = Get-DialogInfo $fe; $null -ne $script:dlg } 4000 | Out-Null
        $dlg = $script:dlg
        Check "($key) a message box appears (owned by the explorer)" ($null -ne $dlg -and (Get-Owner ([IntPtr]$dlg.hwnd)) -eq [int64]$s.explorerHwnd) @{ dlg = $dlg.text; owner = if ($dlg) { Get-Owner ([IntPtr]$dlg.hwnd) } }
        if ($dlg) {
            Start-Sleep -Milliseconds 600   # a message box that stays up for a while, the hook must not touch its keys
            Press-Key $script:VK[$key]
            $closed = Wait-Until { $null -eq (Get-DialogInfo $fe) } 2500
            if (-not $closed) { [void][FeWin2]::PostMessage([IntPtr]$dlg.hwnd, 0x10, [IntPtr]::Zero, [IntPtr]::Zero); Start-Sleep -Milliseconds 500 }
            Check "$key dismissed the message box" $closed ''
        }
        Start-Sleep -Milliseconds 400
        $s = Get-FeState $fe
        Check "$key over the message box did not close the explorer or go Back" ($s.isOpen -and (Get-Stack $s).Count -eq 0 -and $s.session.root.address -eq '$K$36') @{ open = $s.isOpen; root = $s.session.root.address; stack = (Get-Stack $s).Count }
    }
    [void]$wbk.Unprotect()
    Check-LogClean
    End-Scenario
}
