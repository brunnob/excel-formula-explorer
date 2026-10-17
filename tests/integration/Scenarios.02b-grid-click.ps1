# S4b: a real click on Excel's grid deactivates the explorer (its title goes grey, the keys go to Excel) while it stays on top;
# a link click inside the explorer keeps it active. (Excel's grid answers WM_MOUSEACTIVATE with MA_NOACTIVATE while a window it
# owns is active; the explorer's message hook activates Excel's window after the click.)

function S4b {
    Begin-Scenario 'S4b' 'Click on the Excel grid deactivates the explorer (stays on top, keys go to Excel); link clicks keep it active'
    Reset-Env
    $fe = $script:fe
    $s = Open-Explorer $fe 'Val_FO' 'K33'
    $eh = [int64]$s.explorerHwnd
    $b = $s.view.screenBounds
    # activate the explorer with a REAL click on its title bar (no synthetic Alt tap, which would put it in menu mode)
    Click-At ($b.x + 80) ($b.y + 12); Start-Sleep -Milliseconds 500
    Check 'precondition: the explorer is the foreground window' (Test-ExplorerForeground $fe $s) (Describe-Fg)

    foreach ($cell in @('M40', 'K33')) {
        Click-At ($b.x + 80) ($b.y + 12); Start-Sleep -Milliseconds 400
        Click-Cell $fe $cell
        Start-Sleep -Milliseconds 600
        $fg = [FeWin2]::GetForegroundWindow()
        Check "click on grid cell ${cell}: Excel's window is the foreground window, not the explorer" ([FeWin2]::ClassOf($fg) -eq 'XLMAIN' -and $fg.ToInt64() -ne $eh) (Describe-Fg)
        $st = Get-FeState $fe
        Check "($cell) the explorer reports itself inactive (title greyed)" (-not $st.buttons.windowActive) $st.buttons.windowActive
        $r = Get-WinRect $eh
        Check "($cell) the explorer is still on top of Excel's window" ([FeWin2]::TopWindowAt($r[0] + 60, $r[1] + 14).ToInt64() -eq $eh) ('0x{0:X}' -f [FeWin2]::TopWindowAt($r[0] + 60, $r[1] + 14).ToInt64())
        Save-FeShot $fe "S4b-after-click-$cell" | Out-Null
    }

    # the keyboard now belongs to Excel: typing starts editing the selected cell
    Press-Key 0x31; Start-Sleep -Milliseconds 500
    Check 'typing a character reaches Excel (the cell is in edit mode)' (Test-ExcelEditMode $fe) ''
    Press-Key $script:VK.Esc; Start-Sleep -Milliseconds 700
    Check 'Esc left the edit and did not close the explorer (the keys went to Excel, not to the explorer)' (-not (Test-ExcelEditMode $fe) -and (Get-FeState $fe).isOpen) ''

    # a link click in the explorer still keeps it active
    $s = Get-FeState $fe
    $b = $s.view.screenBounds
    Click-At ($b.x + 80) ($b.y + 12); Start-Sleep -Milliseconds 400
    $s = Get-FeState $fe
    Click-Node $fe $s 'Hundred' 1 2
    $r = Wait-Timed $fe { (Get-GoldText $_) -ceq 'Hundred' } 3000
    Check 'link click: gold on Hundred' $r.ok (Get-GoldText $r.state)
    Check 'the explorer stays the foreground window after the link click' (Get-Foreground-Timed $r.state 900) (Describe-Fg)
    Check-LogClean
    End-Scenario
}
