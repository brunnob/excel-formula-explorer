# S3, S4, S5: real mouse (physical cursor, mouse_event) over the explorer window.

function Get-TopLevelSet($fe) { @([FeWin2]::TopLevel([uint32]$fe.Pid)) }

# Windows that appeared since $Before (lines "hwnd|class|title|l,t,r,b"), excluding the explorer and the overlay.
function Get-NewWindows($fe, $Before, $state) {
    $known = @{}; foreach ($l in $Before) { $known[$l.Split('|')[0]] = $true }
    $res = @()
    foreach ($l in (Get-TopLevelSet $fe)) {
        $p = $l.Split('|')
        if ($known.ContainsKey($p[0])) { continue }
        if ([int64]$p[0] -eq [int64]$state.explorerHwnd -or [int64]$p[0] -eq [int64]$state.overlayHwnd) { continue }
        $res += , @{ hwnd = [int64]$p[0]; cls = $p[1]; title = $p[2]; rect = ($p[3].Split(',') | ForEach-Object { [int]$_ }) }
    }
    $res
}

function Move-ToChar($state, [int]$Index) {
    $p = Char-Point $state $Index
    Move-Mouse $p.X $p.Y
    $p
}

# Polls the dump until $Cond is true; returns @{ ok; ms; state }.
function Wait-Timed($fe, [scriptblock]$Cond, [int]$TimeoutMs = 2500, [int]$PollMs = 25) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    do {
        $s = Get-FeState $fe
        $ok = $false; try { $ok = [bool]($s | ForEach-Object $Cond) } catch { }
        if ($ok) { return @{ ok = $true; ms = $sw.ElapsedMilliseconds; state = $s } }
        Start-Sleep -Milliseconds $PollMs
    } while ($sw.ElapsedMilliseconds -lt $TimeoutMs)
    @{ ok = $false; ms = $sw.ElapsedMilliseconds; state = $s }
}

function S3 {
    Begin-Scenario 'S3' 'Real mouse hover: highlight follows, tooltip timing/placement/persistence, cursors, empty areas'
    Reset-Env
    $fe = $script:fe
    $s = Open-Explorer $fe 'Val_FO' 'K33'
    $text = $s.view.displayText
    $iCaret = Ix $s '^'; $iHun = Ix $s 'Hundred' 1 2; $iPlus = Ix $s '+1' 1 0
    $wantCaret = '(($I27/Hundred)+1)^(K$31)'

    # --- hover the '^' ---
    $before = Get-TopLevelSet $fe
    $pt = Move-ToChar $s $iCaret
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $r = Wait-Timed $fe { $null -ne $_.view.hoverStart } 800
    $s1 = $r.state; $hoverMs = $sw.ElapsedMilliseconds
    Check 'highlight follows the real pointer onto ^ (span)' ((Get-HoverText $s1) -ceq $wantCaret) @{ got = (Get-HoverText $s1); want = $wantCaret; ms = $hoverMs }
    Check 'no tooltip immediately after arriving' (-not $s1.view.tooltipVisible) $s1.view.tooltipVisible
    $r = Wait-Timed $fe { $_.view.tooltipVisible } 2500
    $tipMs = $sw.ElapsedMilliseconds
    Check 'tooltip appears about 500 ms after the pointer stops (430..1000 ms incl. polling latency)' ($r.ok -and $tipMs -ge 430 -and $tipMs -le 1000) "measured $tipMs ms since the move"
    $s2 = $r.state
    Check 'cursor over a non-link is the arrow' ((Get-CursorKind) -eq 'arrow') (Get-CursorKind)
    # the gray #DCDCDC is really painted behind the span
    $span = Get-SpanRect $s1 ([int]$s1.view.hoverStart) ([int]$s1.view.hoverLength)
    $bmp = [FeWin2]::Capture($span.X, $span.Y, $span.W, $span.H)
    $gray = [FeWin2]::CountColor($bmp, 0, 0, $span.W, $span.H, $script:Gray, 3); $bmp.Dispose()
    Check 'hover paints #DCDCDC behind the node' ($gray -gt 200) "gray pixels=$gray"
    Check 'tooltip text is the value of the highlighted node' ($s2.view.tooltipText -ceq (Eval-Expected 'Val_FO' $wantCaret)) $s2.view.tooltipText
    Save-FeShot $fe 'S3-tooltip-caret' | Out-Null
    $new = Get-NewWindows $fe $before $s2
    $cur = [FeWin2]::CursorPos()
    $tip = $new | Where-Object { $_.rect[2] - $_.rect[0] -gt 40 } | Select-Object -First 1
    if ($tip) {
        $dx = $tip.rect[0] - $cur[0]; $dy = $tip.rect[1] - $cur[1]
        Check 'tooltip window is below-right of the pointer (left edge near pointer x, top ~17 DIP below)' ($dx -ge -20 -and $dx -le 20 -and $dy -ge 12 -and $dy -le 70) @{ dx = $dx; dy = $dy; rect = $tip.rect; cursor = $cur }
    }
    else { Check 'tooltip window found among the process windows' $false ($new | ForEach-Object { "$($_.cls) $($_.rect -join ',')" }) -Level soft }

    # --- stays while the pointer is still (> 3 s) ---
    Start-Sleep -Milliseconds 3200
    $s3 = Get-FeState $fe
    Check 'tooltip still visible after 3.2 s without moving' ($s3.view.tooltipVisible -and (Get-HoverText $s3) -ceq $wantCaret) @{ visible = $s3.view.tooltipVisible; hover = (Get-HoverText $s3) }
    # small jitter inside the same node does not make it flicker
    $cur = [FeWin2]::CursorPos(); Move-Mouse ($cur[0] + 1) $cur[1]; Start-Sleep -Milliseconds 150
    $sJ = Get-FeState $fe
    Check 'a 1 px move inside the same char keeps the tooltip' $sJ.view.tooltipVisible $sJ.view.tooltipVisible

    # --- move to another node: highlight changes, old tooltip disappears at once, new one comes later ---
    $pt2 = Move-ToChar $s $iHun
    $r = Wait-Timed $fe { (Get-HoverText $_) -ceq 'Hundred' } 800
    $s4 = $r.state
    Check 'moving to Hundred switches the highlight' ((Get-HoverText $s4) -ceq 'Hundred') (Get-HoverText $s4)
    Check 'the previous tooltip is gone right away' (-not $s4.view.tooltipVisible -or $s4.view.tooltipText -ceq '100') @{ visible = $s4.view.tooltipVisible; text = $s4.view.tooltipText }
    Check 'hover on a link shows the hand cursor' ((Get-CursorKind) -eq 'hand') (Get-CursorKind)
    Check 'hovered link is reported as a link' ($s4.view.hoverIsLink) $s4.view.hoverIsLink
    $underline = $null
    $r = Wait-Timed $fe { $_.view.tooltipVisible } 2000
    Check 'tooltip for Hundred appears with value 100' ($r.ok -and $r.state.view.tooltipText -ceq '100') $r.state.view.tooltipText
    Save-FeShot $fe 'S3-tooltip-hundred' | Out-Null

    # --- empty area right of the text (same line) and below the text inside the box: nothing highlighted ---
    $last = $text.Length - 1
    $lastPt = Char-Point $s $last
    Move-Mouse ($lastPt.X + 140) $lastPt.Y
    $r = Wait-Timed $fe { $null -eq $_.view.hoverStart } 800
    Check 'empty area right of the text highlights nothing' $r.ok (Get-HoverText $r.state)
    Start-Sleep -Milliseconds 700
    $s5 = Get-FeState $fe
    Check 'and shows no tooltip' (-not $s5.view.tooltipVisible) $s5.view.tooltipVisible
    Check 'and the cursor is the arrow' ((Get-CursorKind) -eq 'arrow') (Get-CursorKind)
    $firstPt = Char-Point $s 0
    Move-Mouse ($firstPt.X + 10) ($firstPt.Y + 52)   # two lines below the (single) text line, still inside the 3-line box
    Start-Sleep -Milliseconds 250
    $s6 = Get-FeState $fe
    Check 'empty area below the text highlights nothing' ($null -eq $s6.view.hoverStart) (Get-HoverText $s6)
    # --- back onto a node after the empty area, then leave the window entirely ---
    Move-ToChar $s $iCaret | Out-Null
    $r = Wait-Timed $fe { $_.view.tooltipVisible } 2500
    Check 'tooltip returns when hovering a node again' $r.ok $r.state.view.tooltipText
    Move-Mouse 600 1200
    Start-Sleep -Milliseconds 300
    $s7 = Get-FeState $fe
    Check 'leaving the window clears highlight and tooltip' ($null -eq $s7.view.hoverStart -and -not $s7.view.tooltipVisible) @{ hover = (Get-HoverText $s7); tip = $s7.view.tooltipVisible }

    # --- hover works while the explorer is inactive (click Excel's title bar to deactivate it) ---
    $er = Get-Rect (Get-ExcelHwnd $fe)
    Click-At ($er[0] + 800) ($er[1] + 14)
    Start-Sleep -Milliseconds 500
    $s8 = Get-FeState $fe
    Check 'clicking Excel title bar deactivates the explorer' (-not (Test-ExplorerForeground $fe $s8)) (Describe-Fg)
    Check 'the explorer window stays visible above Excel' ([FeWin2]::TopWindowAt($lastPt.X - 100, $lastPt.Y).ToInt64() -eq [int64]$s8.explorerHwnd) ('0x{0:X}' -f [FeWin2]::TopWindowAt($lastPt.X - 100, $lastPt.Y).ToInt64())
    Move-ToChar $s $iCaret | Out-Null
    $r = Wait-Timed $fe { $_.view.tooltipVisible -and (Get-HoverText $_) -ceq $wantCaret } 2500
    Check 'hover highlight + tooltip work while the explorer is inactive' $r.ok @{ hover = (Get-HoverText $r.state); tip = $r.state.view.tooltipVisible }
    Check 'still inactive (hovering does not activate it)' (-not (Test-ExplorerForeground $fe $r.state)) (Describe-Fg)
    Check-LogClean
    End-Scenario
}

function Get-Foreground-Timed($state, [int]$WaitMs = 700) {
    Wait-Until { [FeWin2]::GetForegroundWindow().ToInt64() -eq [int64]$state.explorerHwnd } $WaitMs 25
}

function S4 {
    Begin-Scenario 'S4' 'Real mouse clicks: link navigation + gold, window stays active, one click drills into a formula cell, a habitual double click drills only once, Back and Close buttons'
    Reset-Env
    $fe = $script:fe
    $s = Open-Explorer $fe 'Val_FO' 'K33'
    $hwnd = [int64]$s.explorerHwnd
    # make the explorer the foreground window first (the user would click its title)
    [FeWin]::Foreground([IntPtr]$hwnd); Start-Sleep -Milliseconds 300
    Check 'explorer is foreground before the click' ([FeWin2]::GetForegroundWindow().ToInt64() -eq $hwnd) (Describe-Fg)

    # --- single click on Hundred (other sheet) ---
    $iHun = Ix $s 'Hundred' 1 2
    $p = Char-Point $s $iHun
    Click-At $p.X $p.Y
    $r = Wait-Timed $fe { $null -ne $_.session.goldTarget } 2500
    $s1 = $r.state
    Check 'click on Hundred sets gold on Hundred' ((Get-GoldText $s1) -ceq 'Hundred') (Get-GoldText $s1)
    Check 'Excel navigated to Names_BA!C5' ($fe.Excel.ActiveSheet.Name -eq 'Names_BA' -and $fe.Excel.ActiveCell.Address($false, $false) -eq 'C5') @{ sheet = $fe.Excel.ActiveSheet.Name; cell = $fe.Excel.ActiveCell.Address($false, $false) }
    Check 'the displayed formula did not change' ($s1.view.displayText -ceq $s.view.displayText) $s1.view.displayText
    Check 'Back is enabled (gold present)' $s1.view.backEnabled $s1.view.backEnabled
    $fgOk = Get-Foreground-Timed $s1 900
    Check 'the explorer stays the foreground window after the click (GetForegroundWindow == explorer)' $fgOk (Describe-Fg)
    Start-Sleep -Milliseconds 700
    Check 'still foreground 700 ms later (no activation tug of war)' ([FeWin2]::GetForegroundWindow().ToInt64() -eq $hwnd) (Describe-Fg)
    # gold is painted, text under the gold keeps its color
    $span = Get-SpanRect $s1 ([int]$s1.view.goldStart) ([int]$s1.view.goldLength)
    # the pointer is still over the gold node: gold only (no gray, no underline, no tooltip)
    $bmp = [FeWin2]::Capture($span.X, $span.Y, $span.W, $span.H)
    $gold = [FeWin2]::CountColor($bmp, 0, 0, $span.W, $span.H, $script:Gold, 3)
    $gray = [FeWin2]::CountColor($bmp, 0, 0, $span.W, $span.H, $script:Gray, 3); $bmp.Dispose()
    Check 'gold #FFD700 is painted behind the clicked node' ($gold -gt 100) "gold pixels=$gold"
    Check 'no gray on the gold node while the pointer is on it' ($gray -lt 15) "gray pixels=$gray"
    Start-Sleep -Milliseconds 700
    $s1b = Get-FeState $fe
    Check 'no tooltip on a freshly gold node' (-not $s1b.view.tooltipVisible) $s1b.view.tooltipVisible
    Save-FeShot $fe 'S4-click-hundred' | Out-Null

    # --- click on a link in the same sheet ($I27) moves the gold ---
    $iI27 = Ix $s1b '$I27'
    $p = Char-Point $s1b $iI27
    Click-At $p.X $p.Y
    $r = Wait-Timed $fe { (Get-GoldText $_) -ceq '$I27' } 2500
    Check 'clicking $I27 moves the gold there' $r.ok (Get-GoldText $r.state)
    Check 'Excel selected Val_FO!I27' ($fe.Excel.ActiveSheet.Name -eq 'Val_FO' -and $fe.Excel.ActiveCell.Address($false, $false) -eq 'I27') $fe.Excel.ActiveCell.Address($false, $false)
    Check 'explorer still foreground after the 2nd click' (Get-Foreground-Timed $r.state 900) (Describe-Fg)
    $s2 = $r.state

    # --- ONE real click on K$31 (a cell with a formula): navigates and drills at once ---
    $iK31 = Ix $s2 'K$31'
    $p = Char-Point $s2 $iK31
    Click-At $p.X $p.Y
    $r = Wait-Timed $fe { $_.session.root.address -eq '$K$31' -and -not $_.view.isLoading -and $_.view.displayText -eq '=J31+K19' } 3000
    $s3 = $r.state
    Check 'single click on K$31 drills down: root K31, text =J31+K19' $r.ok @{ root = $s3.session.root.address; text = $s3.view.displayText }
    Check 'stack has the previous root (K33)' ((Get-Stack $s3).Count -eq 1 -and (Get-Stack $s3)[0].address -eq '$K$33') (Get-Stack $s3)
    Check 'gold cleared after drilling, Back enabled' ($null -eq $s3.view.goldStart -and $s3.view.backEnabled) @{ gold = (Get-GoldText $s3); back = $s3.view.backEnabled }
    Check 'Excel selection is K31' ($fe.Excel.ActiveCell.Address($false, $false) -eq 'K31') $fe.Excel.ActiveCell.Address($false, $false)
    Check 'the explorer STAYED the foreground window through the clicks and the drill' (Get-Foreground-Timed $s3 900) (Describe-Fg)
    Save-FeShot $fe 'S4-after-drill' | Out-Null

    # --- real click on the Back button ---
    $back = $s3.buttons.back
    Click-At ([int](($back[0] + $back[2]) / 2)) ([int](($back[1] + $back[3]) / 2))
    $r = Wait-Timed $fe { $_.session.root.address -eq '$K$33' -and -not $_.view.isLoading -and $_.view.displayText -like '=1/*' } 3000
    $s4 = $r.state
    Check 'real click on Back returns to K33' $r.ok @{ root = $s4.session.root.address; text = $s4.view.displayText }
    Check 'Back disabled again, Excel selected K33' (-not $s4.view.backEnabled -and $fe.Excel.ActiveCell.Address($false, $false) -eq 'K33') @{ back = $s4.view.backEnabled; cell = $fe.Excel.ActiveCell.Address($false, $false) }
    Check 'explorer still foreground after clicking Back' (Get-Foreground-Timed $s4 700) (Describe-Fg)

    # --- a habitual real double click on K$31 drills ONE level only (the 2nd press lands on the new text and is ignored) ---
    Start-Sleep -Milliseconds 400
    $iK31 = Ix $s4 'K$31'
    $p = Char-Point $s4 $iK31
    Click-At $p.X $p.Y -Double -GapMs 110
    $r = Wait-Timed $fe { $_.session.root.address -eq '$K$31' -and -not $_.view.isLoading -and $_.view.displayText -eq '=J31+K19' } 3000
    Start-Sleep -Milliseconds 700
    $sD = Get-FeState $fe
    Check 'double click on K$31 drills exactly one level: root K31, stack [K33], no gold' ($r.ok -and $sD.session.root.address -eq '$K$31' -and (Get-Stack $sD).Count -eq 1 -and $null -eq $sD.view.goldStart) @{ root = $sD.session.root.address; stack = (Get-Stack $sD).Count; gold = (Get-GoldText $sD) }
    $back = $sD.buttons.back
    Click-At ([int](($back[0] + $back[2]) / 2)) ([int](($back[1] + $back[3]) / 2))
    $r = Wait-Timed $fe { $_.session.root.address -eq '$K$33' -and -not $_.view.isLoading -and $_.view.displayText -like '=1/*' } 3000
    $s4 = $r.state
    Check 'Back after the double click returns to K33' $r.ok @{ root = $s4.session.root.address }

    # --- a click on the INACTIVE explorer both activates it and acts ---
    $er = Get-Rect (Get-ExcelHwnd $fe)
    Click-At ($er[0] + 800) ($er[1] + 14)
    Start-Sleep -Milliseconds 500
    $sIn = Get-FeState $fe
    Check 'precondition: explorer inactive after clicking the Excel title bar' (-not (Test-ExplorerForeground $fe $sIn)) (Describe-Fg)
    $iHun = Ix $sIn 'Hundred' 1 2
    $p = Char-Point $sIn $iHun
    Click-At $p.X $p.Y
    $r = Wait-Timed $fe { (Get-GoldText $_) -ceq 'Hundred' } 2500
    Check 'first click on the inactive explorer acts (gold on Hundred, Excel on Names_BA)' ($r.ok -and $fe.Excel.ActiveSheet.Name -eq 'Names_BA') @{ gold = (Get-GoldText $r.state); sheet = $fe.Excel.ActiveSheet.Name }
    Check 'and activates the explorer window' (Get-Foreground-Timed $r.state 900) (Describe-Fg)
    $s4 = $r.state

    # --- real click on the Close button ---
    $close = $s4.buttons.close
    Click-At ([int](($close[0] + $close[2]) / 2)) ([int](($close[1] + $close[3]) / 2))
    $r = Wait-Timed $fe { -not $_.isOpen } 3000
    Check 'real click on Close closes the explorer' $r.ok $r.state.isOpen
    Check 'overlay hidden after close' (-not $r.state.overlayVisible -and (Get-Overlay $r.state).Count -eq 0) @{ visible = $r.state.overlayVisible }
    Check 'the explorer window is gone' ([FeWin]::FindTopLevel([uint32]$fe.Pid, 'Formula Explorer') -eq [IntPtr]::Zero) ''
    Check 'Close leaves the Excel selection where it is (Names_BA!C5 after the last link click)' ($fe.Excel.ActiveSheet.Name -eq 'Names_BA' -and $fe.Excel.ActiveCell.Address($false, $false) -eq 'C5') ($fe.Excel.ActiveSheet.Name + '!' + $fe.Excel.ActiveCell.Address($false, $false))
    Check-LogClean
    End-Scenario
}

# Clicks a node by its text, then waits for gold; returns the state.
function Click-Node($fe, $state, [string]$Needle, [int]$Occ = 1, [int]$Off = 0, [switch]$Double, [string]$WaitGold) {
    $idx = Find-FeIndex $state.view.displayText $Needle $Occ $Off
    $p = Char-Point $state $idx
    Click-At $p.X $p.Y -Double:$Double -GapMs 110
}

function S5 {
    Begin-Scenario 'S5' 'The full video sequence with the real mouse (K33 ... Back x3)'
    Reset-Env
    $fe = $script:fe
    $s = Open-Explorer $fe 'Val_FO' 'K33'
    $hwnd = [int64]$s.explorerHwnd
    [FeWin]::Foreground([IntPtr]$hwnd); Start-Sleep -Milliseconds 300
    $cell = { $fe.Excel.ActiveCell.Address($false, $false) }

    # click Hundred and $I27 (constants: they only turn gold)
    foreach ($n in @(@('Hundred', 1, 2), @('$I27', 1, 0))) {
        Click-Node $fe $s $n[0] $n[1] $n[2]
        $want = $n[0]
        $r = Wait-Timed $fe { (Get-GoldText $_) -ceq $want } 2500
        Check "click $want -> gold on it" $r.ok (Get-GoldText $r.state)
        Check "explorer still foreground after clicking $want" (Get-Foreground-Timed $r.state 800) (Describe-Fg)
        Start-Sleep -Milliseconds 350   # more than a double-click time between separate clicks
    }
    $s = Get-FeState $fe
    Check 'after 2 clicks: root unchanged (K33), stack empty, gold on $I27' ($s.session.root.address -eq '$K$33' -and (Get-Stack $s).Count -eq 0 -and (Get-GoldText $s) -ceq '$I27') @{ root = $s.session.root.address; gold = (Get-GoldText $s) }

    # one click on K$31 -> navigates and drills to K31
    Click-Node $fe $s 'K$31' 1 0
    $r = Wait-Timed $fe { $_.session.root.address -eq '$K$31' -and -not $_.view.isLoading -and $_.view.displayText -eq '=J31+K19' } 3500
    $s = $r.state
    Check 'click K$31 drills: =J31+K19, stack [K33], selection K31' ($r.ok -and (Get-Stack $s).Count -eq 1 -and (& $cell) -eq 'K31') @{ text = $s.view.displayText; stack = (Get-Stack $s).Count; sel = (& $cell) }
    # K19 sits in the collapsed outline group (row 19 hidden): zero-height areas are skipped, so only J31 gets a box.
    Check 'overlay redrawn for the new formula (J31 box; K19 is in a hidden row)' ((Get-Overlay $s).Count -eq 1) (Get-Overlay $s)
    Start-Sleep -Milliseconds 300

    # click K19, J31
    Click-Node $fe $s 'K19'
    $r = Wait-Timed $fe { (Get-GoldText $_) -ceq 'K19' } 2500
    Check 'click K19: gold K19 and Excel selects K19 (row 19 unhidden by the outline rule)' ($r.ok -and (& $cell) -eq 'K19') @{ gold = (Get-GoldText $r.state); sel = (& $cell) }
    Start-Sleep -Milliseconds 350
    # click J31 (a cell with a formula): navigates and drills at once, gold cleared
    Click-Node $fe $r.state 'J31'
    $r = Wait-Timed $fe { $_.session.root.address -eq '$J$31' -and -not $_.view.isLoading -and $_.view.displayText -like '=SUM(*' } 3500
    $s = $r.state
    Check 'click J31 drills: root J31 with the OFFSET formula, stack depth 2, selection J31, no gold' ($r.ok -and (Get-Stack $s).Count -eq 2 -and (& $cell) -eq 'J31' -and $null -eq $s.view.goldStart) @{ root = $s.session.root.address; text = $s.view.displayText; stack = (Get-Stack $s).Count; sel = (& $cell); gold = (Get-GoldText $s) }
    Save-FeShot $fe 'S5-at-J31' | Out-Null
    Start-Sleep -Milliseconds 300

    # click OFFSET
    Click-Node $fe $s 'OFFSET' 1 2
    $r = Wait-Timed $fe { (Get-GoldText $_) -like 'OFFSET(*' } 2500
    $s = $r.state
    $sel = $fe.Excel.Selection.Address($false, $false)
    Check 'click OFFSET: gold on the whole OFFSET(...) call, Excel selects I19:J19' ($r.ok -and $sel -eq 'I19:J19') @{ gold = (Get-GoldText $s); sel = $sel }
    Save-FeShot $fe 'S5-offset-gold' | Out-Null
    $btn = $s.buttons.back

    # Back x3 with the real mouse
    $expect = @(
        @{ root = '$J$31'; sel = 'J31'; back = $true; stack = 2; label = 'Back 1: still J31 (gold removed), selection J31' },
        @{ root = '$K$31'; sel = 'K31'; back = $true; stack = 1; label = 'Back 2: K31 formula, selection K31' },
        @{ root = '$K$33'; sel = 'K33'; back = $false; stack = 0; label = 'Back 3: K33 formula, selection K33, Back disabled' }
    )
    foreach ($e in $expect) {
        $s = Get-FeState $fe
        $b = $s.buttons.back
        Click-At ([int](($b[0] + $b[2]) / 2)) ([int](($b[1] + $b[3]) / 2))
        Start-Sleep -Milliseconds 250
        $r = Wait-Timed $fe { -not $_.view.isLoading -and $_.session.root.address -eq $e.root -and ($null -eq $_.view.goldStart) -and ((Get-Stack $_).Count -eq $e.stack) } 3500
        $s = $r.state
        Check $e.label ($r.ok -and (& $cell) -eq $e.sel -and $s.view.backEnabled -eq $e.back) @{ root = $s.session.root.address; stack = (Get-Stack $s).Count; sel = (& $cell); back = $s.view.backEnabled; gold = (Get-GoldText $s); text = $s.view.displayText }
    }
    Check 'row 19 is hidden again after Back (outline restored)' ($fe.Excel.ActiveSheet.Range('A19').EntireRow.Hidden -eq $true) $fe.Excel.ActiveSheet.Range('A19').EntireRow.Hidden
    Check 'the explorer was foreground throughout the Back sequence' (Get-Foreground-Timed $s 500) (Describe-Fg)
    Save-FeShot $fe 'S5-end' | Out-Null
    Check-LogClean
    End-Scenario
}
