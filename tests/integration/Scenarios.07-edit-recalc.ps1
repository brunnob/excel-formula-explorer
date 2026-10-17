# S15 (cell edit mode), S16 (recalculation / editing the root).

function Test-ExcelEditMode($fe) {
    try { $null = $fe.Excel.ActiveCell.Address($false, $false); $false } catch { $true }
}

function S15 {
    Begin-Scenario 'S15' 'Cell edit mode (F2) while open: hover/click in the explorer, Esc out; no crash, no dialogs, recovery'
    Reset-Env
    $fe = $script:fe
    $s = Open-Explorer $fe 'Val_FO' 'K33'
    $before = Get-TopLevelSet $fe
    # real click on the grid cell K33 (root: no push), then F2 with the real keyboard
    Click-Cell $fe 'K33'
    Start-Sleep -Milliseconds 500
    $origFormula = $fe.Excel.ActiveSheet.Range('K33').Formula
    Press-Key $script:VK.F2
    Start-Sleep -Milliseconds 500
    Press-Key 0x30    # type a '0' so that there is a pending change that a commit would store
    Start-Sleep -Milliseconds 400
    $edit = Test-ExcelEditMode $fe
    Check 'Excel is in cell edit mode (object model refuses calls)' $edit ''
    Save-FeShot $fe 'S15-edit-mode' | Out-Null
    if (-not $edit) { Check-LogClean; End-Scenario; return }

    # explorer reactions while Excel is busy: hover (tooltip needs an evaluation), click on a link, double click
    $text = $s.view.displayText
    $b = $s.view.screenBounds
    $iHun = Find-FeIndex $text 'Hundred' 1 2
    $iCaret = Find-FeIndex $text '^'
    $pHun = Char-Point $s $iHun; $pCaret = Char-Point $s $iCaret
    Move-Mouse $pCaret.X $pCaret.Y; Start-Sleep -Milliseconds 900
    Move-Mouse $pHun.X $pHun.Y; Start-Sleep -Milliseconds 900
    Save-FeShot $fe 'S15-hover-while-editing' | Out-Null
    Click-At $pHun.X $pHun.Y
    Start-Sleep -Milliseconds 600
    Click-At $pHun.X $pHun.Y -Double
    Start-Sleep -Milliseconds 800
    $dialogs = @([FeWin2]::TopLevel([uint32]$fe.Pid) | Where-Object { $_.Split('|')[1] -eq '#32770' })
    Check 'no dialog / exception box appeared in the Excel process' ($dialogs.Count -eq 0) $dialogs
    Check 'the explorer window is still there (alive) while Excel is editing' ([FeWin2]::IsWindow([IntPtr][int64]$s.explorerHwnd)) ''
    Check 'Excel is still in edit mode (the explorer did not touch Excel)' (Test-ExcelEditMode $fe) ''
    # the Close button also works while busy? (do not press it: we want to come back) - hover the Back/Close buttons
    $c = $s.buttons.close
    Move-Mouse ([int](($c[0] + $c[2]) / 2)) ([int](($c[1] + $c[3]) / 2)); Start-Sleep -Milliseconds 300

    # leave edit mode: click Excel's title bar (activates Excel, edit mode stays), then a real Esc
    $er = Get-Rect (Get-ExcelHwnd $fe)
    Click-At ($er[0] + 800) ($er[1] + 14)
    Start-Sleep -Milliseconds 500
    Press-Key $script:VK.Esc
    Start-Sleep -Milliseconds 900
    $edit2 = Test-ExcelEditMode $fe
    Check 'Esc left the edit mode' (-not $edit2) ''
    if ($edit2) { Press-Key $script:VK.Esc; Start-Sleep -Milliseconds 600 }
    # recovery
    $s2 = Wait-FeState $fe { $_.isOpen -and $_.view -and -not $_.view.isLoading } 5000
    Check 'explorer still open after the edit' $s2.isOpen ''
    Check 'it shows the same formula and root' ($s2.view.displayText -ceq $text -and $s2.session.root.address -eq '$K$33') @{ text = $s2.view.displayText; root = $s2.session.root.address }
    Start-Sleep -Milliseconds 1500
    $s2 = Get-FeState $fe
    Check 'Esc cancelled the pending edit: the formula of K33 is unchanged (the explorer did not commit it)' ($fe.Excel.ActiveSheet.Range('K33').Formula -ceq $origFormula) $fe.Excel.ActiveSheet.Range('K33').Formula
    Check 'the clicks made during the edit were ignored (no late navigation: still on Val_FO!K33, no gold, no push)' ($fe.Excel.ActiveSheet.Name -eq 'Val_FO' -and $fe.Excel.ActiveCell.Address($false, $false) -eq 'K33' -and $null -eq $s2.view.goldStart -and (Get-Stack $s2).Count -eq 0) @{ sheet = $fe.Excel.ActiveSheet.Name; cell = $fe.Excel.ActiveCell.Address($false, $false); gold = (Get-GoldText $s2); stack = (Get-Stack $s2).Count }
    # hover / tooltip work again
    $s2 = Get-FeState $fe
    $p = Char-Point $s2 (Find-FeIndex $s2.view.displayText '^')
    Move-Mouse $p.X $p.Y
    $r = Wait-Timed $fe { $_.view.tooltipVisible } 2500
    Check 'tooltip works again after leaving edit mode' ($r.ok -and $r.state.view.tooltipText -ceq (Eval-Expected 'Val_FO' '(($I27/Hundred)+1)^(K$31)')) $r.state.view.tooltipText
    Check 'overlay is drawn again after the edit' ($r.state.overlayVisible -and (Get-Overlay $r.state).Count -eq 2) @{ visible = $r.state.overlayVisible; rects = (Get-Overlay $r.state).Count }
    # navigating works again
    Click-At $p.X $p.Y   # (the '^' is not a link: nothing)
    $pH = Char-Point $r.state (Find-FeIndex $r.state.view.displayText 'Hundred' 1 2)
    Click-At $pH.X $pH.Y
    $r = Wait-Timed $fe { $null -ne $_.view.goldStart } 3000
    Check 'link click works after the edit (gold on Hundred)' ($r.ok -and (Get-GoldText $r.state) -ceq 'Hundred') (Get-GoldText $r.state)
    $dialogs = @([FeWin2]::TopLevel([uint32]$fe.Pid) | Where-Object { $_.Split('|')[1] -eq '#32770' })
    Check 'still no dialogs at the end' ($dialogs.Count -eq 0) $dialogs
    Check-LogClean 'add-in log (no ERROR lines; busy Excel must not be reported as errors)'
    End-Scenario
}

function S16 {
    Begin-Scenario 'S16' 'Recalculation and editing: I27 6.2 -> 7 refreshes tooltips; editing the root formula reloads the explorer'
    Reset-Env
    $fe = $script:fe
    $ws = $fe.Excel.ActiveSheet
    $s = Open-Explorer $fe 'Val_FO' 'K33'
    $iI27 = Ix $s '$I27' 1 1
    $iCaret = Ix $s '^'
    $iPlus = Ix $s '+1)' 1 2   # the ')' after '+1'
    $text = $s.view.displayText
    $p = Char-Point $s $iI27
    Move-Mouse $p.X $p.Y
    $r = Wait-Tooltip $fe 2500
    Check 'tooltip for $I27 before the change' ($r.view.tooltipText -ceq '6.2') $r.view.tooltipText
    # change I27 while the pointer rests on it: the open tooltip must update by itself
    $ws.Range('I27').Value2 = 7
    $r = Wait-Timed $fe { $_.view.tooltipVisible -and $_.view.tooltipText -ceq '7' } 3000
    Check 'the open tooltip updates to 7 (RefreshHover after recalculation)' $r.ok $r.state.view.tooltipText
    # dependent fragments
    $p = Char-Point $s $iPlus
    Move-Mouse $p.X $p.Y
    $want = Eval-Expected 'Val_FO' '(($I27/Hundred)+1)'
    $r = Wait-Timed $fe { $_.view.tooltipVisible -and $_.view.tooltipText -ceq $want } 3000
    Check "tooltip of (($I27/Hundred)+1) is now $want" $r.ok $r.state.view.tooltipText
    $p = Char-Point $s $iCaret
    Move-Mouse $p.X $p.Y
    $want = Eval-Expected 'Val_FO' '(($I27/Hundred)+1)^(K$31)'
    $r = Wait-Timed $fe { $_.view.tooltipVisible -and $_.view.tooltipText -ceq $want } 3000
    Check "tooltip of the ^ fragment is now $want (was 1.00760985118373)" ($r.ok -and $want -ne '1.00760985118373') $r.state.view.tooltipText
    # the root's own value (hover on '=')
    $p = Char-Point $s 0
    Move-Mouse $p.X $p.Y
    $want = Eval-Expected 'Val_FO' 'K33'
    $r = Wait-Timed $fe { $_.view.tooltipVisible -and $_.view.tooltipText -ceq $want } 3000
    Check "tooltip of '=' is the new value of K33 ($want)" $r.ok $r.state.view.tooltipText
    # recalculation must not reload the formula (no flicker: loading never set); the root is unchanged
    $sNow = Get-FeState $fe
    Check 'recalculation did not reload or push anything' ($sNow.session.root.address -eq '$K$33' -and (Get-Stack $sNow).Count -eq 0 -and $sNow.view.displayText -ceq $text) @{ root = $sNow.session.root.address; stack = (Get-Stack $sNow).Count }

    # edit the root formula: the explorer reloads with the new text
    $newFormula = '=1/(((($I27/Hundred)+2)^(K$31)))'
    [void][System.__ComObject].InvokeMember('Formula', [Reflection.BindingFlags]::SetProperty, $null, $ws.Range('K33'), @($newFormula))
    $r = Wait-FeState $fe { $_.view.displayText -ceq '=1/(((($I27/Hundred)+2)^(K$31)))' -and -not $_.view.isLoading } 4000
    Check 'editing the root formula reloads the explorer with the new text' ($r.view.displayText -ceq $newFormula) $r.view.displayText
    Check 'root and stack unchanged by the edit (stack kept)' ($r.session.root.address -eq '$K$33' -and (Get-Stack $r).Count -eq 0) @{ root = $r.session.root.address; stack = (Get-Stack $r).Count }
    Check 'overlay still shows the two precedents' ((Get-Overlay $r).Count -eq 2) (Get-Overlay $r).Count
    # change the formula to one with a new precedent: overlay gains a box
    $newFormula2 = '=1/(((($I27/Hundred)+2)^(K$31)))+I28'
    [void][System.__ComObject].InvokeMember('Formula', [Reflection.BindingFlags]::SetProperty, $null, $ws.Range('K33'), @($newFormula2))
    $r = Wait-FeState $fe { $_.view.displayText -ceq '=1/(((($I27/Hundred)+2)^(K$31)))+I28' -and -not $_.view.isLoading } 4000
    Start-Sleep -Milliseconds 400
    $r = Get-FeState $fe
    Check 'a new precedent (I28) gets a border after the edit' ((Get-Overlay $r).Count -eq 3) (Get-Overlay $r).Count
    # clearing the root cell: the explorer shows an empty non-interactive text
    $ws.Range('K33').ClearContents() | Out-Null
    $r = Wait-FeState $fe { $_.view.displayText -ceq '' -and -not $_.view.isLoading } 4000
    Check 'clearing the root shows an empty, non-interactive text and removes the borders' ($r.view.displayText -ceq '' -and -not $r.view.interactive -and (Get-Overlay $r).Count -eq 0) @{ text = $r.view.displayText; interactive = $r.view.interactive; rects = (Get-Overlay $r).Count }
    Check-LogClean
    End-Scenario
}
