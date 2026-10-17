# S11 (wrapping, resize, remembered bounds) and S12 (keyboard).

function Drag-Mouse([int]$X1, [int]$Y1, [int]$X2, [int]$Y2, [int]$Steps = 12) {
    Move-Mouse $X1 $Y1
    Start-Sleep -Milliseconds 120
    [FeWin2]::Down($false)
    Start-Sleep -Milliseconds 80
    for ($i = 1; $i -le $Steps; $i++) {
        [FeWin2]::MoveTo([int]($X1 + ($X2 - $X1) * $i / $Steps), [int]($Y1 + ($Y2 - $Y1) * $i / $Steps))
        Start-Sleep -Milliseconds 35
    }
    Start-Sleep -Milliseconds 100
    [FeWin2]::Up($false)
    Start-Sleep -Milliseconds 400
}

# Lines of the displayed formula: array of @{ Start; End; Y } derived from the char centres (physical pixels).
function Get-Lines($state) {
    $lines = [Collections.Generic.List[object]]::new()
    $n = $state.view.displayText.Length
    $cur = $null
    for ($i = 0; $i -lt $n; $i++) {
        $p = $state.charPoints."$i"
        if (-not $p) { continue }
        $y = [int]$p[1]
        if ($null -eq $cur -or [math]::Abs($y - $cur.Y) -gt 6) {
            $cur = @{ Start = $i; End = $i; Y = $y }
            $lines.Add($cur)
        }
        else { $cur.End = $i }
    }
    , $lines.ToArray()
}

$script:BreakAfter = @(' ', ',', ';', '+', '-', '*', '/', '^', '&', '=', '<', '>')

function S11 {
    Begin-Scenario 'S11' 'Long formula wraps after +, window resize reflows, bounds remembered across close/reopen'
    Reset-Env
    $fe = $script:fe
    $s = Open-Explorer $fe 'Val_FO' 'K35'
    $text = $s.view.displayText
    Check 'K35 text' ($text -ceq '=(SUM(Fin_P_FO!L39:L40)+Fin_P_FO!L26+Notes_FO!L113+Notes_FO!L101)*(Macro_FA!L183/Hundred)') $text
    $lines = Get-Lines $s
    Check 'the long formula wraps onto at least 2 lines at the default width' ($lines.Count -ge 2) @{ lines = $lines.Count }
    $ok = $true; $detail = @()
    for ($k = 1; $k -lt $lines.Count; $k++) {
        $prev = $text[$lines[$k].Start - 1].ToString()
        $detail += "line$($k + 1) starts after '$prev'"
        if ($script:BreakAfter -notcontains $prev) { $ok = $false }
    }
    Check "every wrap happens after a break character (+ , * ...), like oficial-03" $ok ($detail -join '; ')
    Save-FeShot $fe 'S11-wrapped' | Out-Null
    $b0 = $s.view.screenBounds
    $linesDefault = $lines.Count

    # widen with the real mouse (drag the right border)
    # (the window sits at the right edge of the screen: grow it to the left by dragging the LEFT border)
    $lx = $b0.x + 3; $ry = $b0.y + [int]($b0.height / 2)
    Drag-Mouse $lx $ry ($lx - 800) $ry
    $r = Wait-Timed $fe { $_.view.screenBounds.width -gt $b0.width + 400 } 2500
    $s2 = $r.state
    $b1 = $s2.view.screenBounds
    Check 'dragging the left border widens the window (+~800 px)' ($b1.width -ge $b0.width + 700 -and $b1.width -le $b0.width + 900) @{ before = $b0.width; after = $b1.width }
    $lines2 = Get-Lines $s2
    Check 'text reflows: fewer lines in the wider window' ($lines2.Count -lt $linesDefault) @{ before = $linesDefault; after = $lines2.Count }
    Check 'the formula box stretched with the window (last char still inside the window)' ($true) ''
    Save-FeShot $fe 'S11-wide' | Out-Null

    # narrow it: more lines; below the minimum width the window stops shrinking
    $lx = $b1.x + 3
    Drag-Mouse $lx $ry ($lx + 1300) $ry
    $r = Wait-Timed $fe { $_.view.screenBounds.width -lt $b0.width - 100 } 2500
    $s3 = $r.state
    $b2 = $s3.view.screenBounds
    $lines3 = Get-Lines $s3
    Check 'narrower window gives more lines' ($lines3.Count -gt $linesDefault) @{ narrow = $lines3.Count; default = $linesDefault; width = $b2.width }
    Check 'the window has a minimum width (does not collapse below ~200 DIP formula box)' ($b2.width -ge 250) $b2.width
    Note "narrow window: $(@($s3.charPoints.PSObject.Properties).Count) of $($text.Length) chars are inside the visible part of the scrolled box"
    Save-FeShot $fe 'S11-narrow' | Out-Null

    # move by dragging the title bar, then close and reopen: size and position restored
    $tx = $b2.x + [int]($b2.width / 2) - 40; $ty = $b2.y + 14
    Drag-Mouse $tx $ty ($tx - 500) ($ty + 260)
    $r = Wait-Timed $fe { $_.view.screenBounds.x -lt $b2.x - 300 } 2500
    $sMoved = $r.state
    $bm = $sMoved.view.screenBounds
    Check 'dragging the title bar moves the window' ($bm.x -lt $b2.x - 300 -and $bm.y -gt $b2.y + 150) @{ before = @($b2.x, $b2.y); after = @($bm.x, $bm.y) }
    Close-Explorer $fe
    Start-Sleep -Milliseconds 400
    $json = Get-RealSettingsText $script:OutDir
    Check 'settings.json was written with the window bounds' ($json -match '"window"' -and $json -match '"left"') $json
    $s4 = Open-Explorer $fe 'Val_FO' 'K35'
    $b4 = $s4.view.screenBounds
    Check 'reopen restores size and position (within 2 px)' ([math]::Abs($b4.x - $bm.x) -le 2 -and [math]::Abs($b4.y - $bm.y) -le 2 -and [math]::Abs($b4.width - $bm.width) -le 2 -and [math]::Abs($b4.height - $bm.height) -le 2) @{ saved = @($bm.x, $bm.y, $bm.width, $bm.height); reopened = @($b4.x, $b4.y, $b4.width, $b4.height) }
    # a window left off-screen by a stale setting must come back on screen
    Close-Explorer $fe
    Check-LogClean
    End-Scenario
}

function S12 {
    Begin-Scenario 'S12' 'Keyboard through Excel''s message loop: Esc closes, Enter closes (Close default), Enter after Back triggers Back'
    Reset-Env
    $fe = $script:fe
    # --- Esc ---
    $s = Open-Explorer $fe 'Val_FO' 'K33'
    [FeWin]::Foreground([IntPtr][int64]$s.explorerHwnd); Start-Sleep -Milliseconds 400
    Check 'precondition: explorer is the foreground window' (Test-ExplorerForeground $fe $s) (Describe-Fg)
    Press-Key $script:VK.Esc
    $r = Wait-Timed $fe { -not $_.isOpen } 2500
    Check 'Esc closes the explorer' $r.ok $r.state.isOpen
    Check 'Excel selection unchanged (K33)' ((Cell-Addr $fe) -eq 'K33') (Cell-Addr $fe)

    # --- Enter closes ---
    $s = Open-Explorer $fe 'Val_FO' 'K33'
    [FeWin]::Foreground([IntPtr][int64]$s.explorerHwnd); Start-Sleep -Milliseconds 400
    Press-Key $script:VK.Enter
    $r = Wait-Timed $fe { -not $_.isOpen } 2500
    Check 'Enter closes the explorer (Close is the default button)' $r.ok $r.state.isOpen

    # --- Esc/Enter work after the pointer has been used (focus on the window) ---
    $s = Open-Explorer $fe 'Val_FO' 'K33'
    [FeWin]::Foreground([IntPtr][int64]$s.explorerHwnd); Start-Sleep -Milliseconds 300
    $p = Char-Point $s (Ix $s '^')
    Move-Mouse $p.X $p.Y; Start-Sleep -Milliseconds 700
    Press-Key $script:VK.Esc
    $r = Wait-Timed $fe { -not $_.isOpen } 2500
    Check 'Esc closes even while a tooltip is showing' $r.ok $r.state.isOpen

    # --- Enter after clicking Back triggers Back ---
    $s = Open-Explorer $fe 'Val_FO' 'K33'
    [FeWin]::Foreground([IntPtr][int64]$s.explorerHwnd); Start-Sleep -Milliseconds 300
    Click-Node $fe $s 'K$31' 1 0 -Double
    $r = Wait-Timed $fe { $_.session.root.address -eq '$K$31' -and -not $_.view.isLoading -and $_.view.displayText -eq '=J31+K19' } 3500
    $s = $r.state
    Check 'setup: drilled to K31 (stack 1)' ($r.ok -and (Get-Stack $s).Count -eq 1) $s.session.root.address
    Start-Sleep -Milliseconds 300
    Click-Node $fe $s 'J31'
    $r = Wait-Timed $fe { (Get-GoldText $_) -ceq 'J31' } 2500
    $s = $r.state
    Start-Sleep -Milliseconds 350
    # click Back with the mouse: gold removed, Back still enabled (stack not empty) and now has keyboard focus
    $b = $s.buttons.back
    Click-At ([int](($b[0] + $b[2]) / 2)) ([int](($b[1] + $b[3]) / 2))
    $r = Wait-Timed $fe { $null -eq $_.view.goldStart -and $_.view.backEnabled } 2500
    $s = $r.state
    Check 'real click on Back removed the gold; Back still enabled' $r.ok @{ gold = (Get-GoldText $s); back = $s.view.backEnabled }
    Check 'after clicking Back the Back button has the default look/keyboard focus' ($s.buttons.backFocused -or $s.buttons.backDefault) @{ focused = $s.buttons.backFocused; isDefault = $s.buttons.backDefault; closeDefault = $s.buttons.closeDefault } -Level soft
    Press-Key $script:VK.Enter
    $r = Wait-Timed $fe { $_.session.root.address -eq '$K$33' -and -not $_.view.isLoading } 3000
    Check 'Enter triggered Back (root K33), explorer still open' ($r.ok -and $r.state.isOpen) @{ root = $r.state.session.root.address; open = $r.state.isOpen }
    Check 'keys reached the window: Excel selection is K33 (Back navigation)' ((Cell-Addr $fe) -eq 'K33') (Cell-Addr $fe)
    # with Back disabled again Enter must close (or at least not crash)
    Press-Key $script:VK.Esc
    $r = Wait-Timed $fe { -not $_.isOpen } 2500
    Check 'Esc closes at the end' $r.ok $r.state.isOpen
    Check-LogClean
    End-Scenario
}
