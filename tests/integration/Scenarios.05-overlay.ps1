# S13: overlay geometry under scroll, zoom, frozen/split panes, window moves, sheet switches, z-order and dialogs.

$script:K33Boxes = @(@{ Address = 'I27'; Rgb = 0x0000FF }, @{ Address = 'K31'; Rgb = 0x9900CC })
$script:SWP_FLAGS_MOVE = 0x0004 -bor 0x0010   # NOZORDER | NOACTIVATE

function Set-View($fe, [scriptblock]$Change) {
    & $Change
    Start-Sleep -Milliseconds 450
}

function S13 {
    Begin-Scenario 'S13' 'Overlay geometry: scroll, zoom, frozen and split panes, clipping, sheet/workbook switch, z-order, Find dialog'
    Reset-Env
    $fe = $script:fe
    $xl = $fe.Excel
    $ws = $xl.ActiveSheet
    # multi-pane root: A2 (top-left pane), K2 (top-right), A30 (bottom-left), K30 (bottom-right), I27
    [void][System.__ComObject].InvokeMember('Formula', [Reflection.BindingFlags]::SetProperty, $null, $ws.Range('R2'), @('=A2+K2+A30+K30+I27'))
    $s = Open-Explorer $fe 'Val_FO' 'K33'
    $win = $xl.ActiveWindow

    # ---- baseline, 100% ----
    $null = Verify-Overlay $fe 'S13 100% top-left' $script:K33Boxes -SaveShot

    # ---- scroll: I27 scrolled off the top, K31 still visible ----
    Set-View $fe { $win.ScrollRow = 28 }
    $null = Verify-Overlay $fe 'S13 scrolled: row 27 above the view' $script:K33Boxes
    Set-View $fe { $win.ScrollRow = 33 }
    $r = Verify-Overlay $fe 'S13 scrolled: both precedents above the view' $script:K33Boxes
    Check 'S13 no box at all when both precedents are off-screen' ($r.Expected.Count -eq 0 -and $r.Reported.Count -eq 0) @{ expected = $r.Expected.Count; reported = $r.Reported.Count }
    # partial visibility at the bottom edge: find a ScrollRow where K31's row is cut by the end of the cell area
    $found = $false
    foreach ($row in 14..31) {
        Set-View $fe { $win.ScrollRow = $row }
        $o = Get-PaneOracle $fe
        $raw = Get-RangeScreenRect $fe $ws.Range('K31') $o.Panes[0].Pane
        if ($raw[1] -lt $o.CellArea[3] - 6 -and $raw[3] -gt $o.CellArea[3] + 4) { $found = $true; break }
    }
    if ($found) { Note "ScrollRow=$row cuts K31 at the bottom of the cell area"; $null = Verify-Overlay $fe "S13 partially visible at the bottom (ScrollRow $row)" $script:K33Boxes -SaveShot }
    else { Note 'could not find a ScrollRow that cuts K31 at the bottom edge (window too tall)' }
    # scrolled horizontally past both
    Set-View $fe { $win.ScrollRow = 1; $win.ScrollColumn = 30 }
    $r = Verify-Overlay $fe 'S13 scrolled right: precedents off to the left' $script:K33Boxes
    Check 'S13 no box when scrolled away horizontally' ($r.Reported.Count -eq 0) $r.Reported.Count
    Set-View $fe { $win.ScrollColumn = 10 }   # K visible, I (col 9) hidden to the left
    $null = Verify-Overlay $fe 'S13 scrolled right: I27 left of the view, K31 visible' $script:K33Boxes
    Set-View $fe { $win.ScrollColumn = 1 }

    # ---- zoom ----
    foreach ($z in 75, 150, 200, 100) {
        Set-View $fe { $win.Zoom = $z }
        $null = Verify-Overlay $fe "S13 zoom $z%" $script:K33Boxes -SaveShot:($z -eq 75 -or $z -eq 150)
    }

    # ---- frozen panes at D5 with the 5-precedent root ----
    $fe.Excel.ActiveWorkbook.Worksheets.Item('Val_FO').Activate()
    $ws.Range('R2').Select() | Out-Null
    Start-Sleep -Milliseconds 700
    $s = Get-FeState $fe
    $multi = @(
        @{ Address = 'A2'; Rgb = 0x0000FF }, @{ Address = 'K2'; Rgb = 0x008000 }, @{ Address = 'A30'; Rgb = 0x9900CC },
        @{ Address = 'K30'; Rgb = 0x800000 }, @{ Address = 'I27'; Rgb = 0x00CC33 })
    Check 'S13 root is R2 with 5 precedents' ($s.session.root.address -eq '$R$2' -and $s.view.displayText -eq '=A2+K2+A30+K30+I27') $s.view.displayText
    $null = Verify-Overlay $fe 'S13 R2 unfrozen' $multi
    Set-View $fe { $win.FreezePanes = $false; $win.ScrollRow = 1; $win.ScrollColumn = 1; $win.SplitColumn = 3; $win.SplitRow = 4; $win.FreezePanes = $true }
    Check 'S13 frozen window has 4 panes' ($win.Panes.Count -eq 4) $win.Panes.Count
    $null = Verify-Overlay $fe 'S13 frozen D5, unscrolled' $multi -SaveShot
    Set-View $fe { $win.ScrollRow = 25; $win.ScrollColumn = 6 }
    $null = Verify-Overlay $fe 'S13 frozen D5, scrolled to row 25 / column F' $multi -SaveShot
    Set-View $fe { $win.ScrollRow = 29; $win.ScrollColumn = 11 }
    $null = Verify-Overlay $fe 'S13 frozen D5, scrolled so K30 sits at the pane corner' $multi
    Set-View $fe { $win.Zoom = 150 }
    $null = Verify-Overlay $fe 'S13 frozen D5 at 150%' $multi
    Set-View $fe { $win.Zoom = 100; $win.FreezePanes = $false; $win.SplitColumn = 0; $win.SplitRow = 0 }

    # ---- split (not frozen) ----
    Set-View $fe { $win.ScrollRow = 1; $win.ScrollColumn = 1; $win.FreezePanes = $false; $win.SplitColumn = 6; $win.SplitRow = 12 }
    Check 'S13 split window has 4 panes' ($win.Panes.Count -eq 4) $win.Panes.Count
    $null = Verify-Overlay $fe 'S13 split at G13' $multi -SaveShot
    Set-View $fe { $win.Panes.Item(4).ScrollRow = 25; $win.Panes.Item(4).ScrollColumn = 7 }
    $null = Verify-Overlay $fe 'S13 split, bottom-right pane scrolled' $multi -SaveShot
    Set-View $fe { $win.SplitColumn = 0; $win.SplitRow = 0; $win.ScrollRow = 1; $win.ScrollColumn = 1 }
    $ws.Range('K33').Select() | Out-Null
    Start-Sleep -Milliseconds 700

    # ---- window moved / resized: Excel un-maximized and narrowed so K31 is cut by the right edge ----
    $xlh = [IntPtr]$xl.ActiveWindow.Hwnd
    $xl.WindowState = -4143
    Start-Sleep -Milliseconds 800
    $o = Get-PaneOracle $fe
    $raw = Get-RangeScreenRect $fe $ws.Range('K31') $o.Panes[0].Pane
    $cur = [FeWin2]::Rect($xlh)
    $newRight = $raw[0] + 70 + 60   # K31's left + ~half a column + vertical scrollbar / frame
    [void][FeWin2]::SetWindowPos($xlh, [IntPtr]::Zero, $cur[0], $cur[1], ($newRight - $cur[0]), ($cur[3] - $cur[1]), 0x0004 -bor 0x0010)
    Start-Sleep -Milliseconds 900
    $null = Verify-Overlay $fe 'S13 narrow un-maximized Excel (K31 cut by the right edge)' $script:K33Boxes -SaveShot
    # now also cut K31 by the bottom edge of the cell area (shrink the window height)
    $o = Get-PaneOracle $fe
    $raw = Get-RangeScreenRect $fe $ws.Range('K31') $o.Panes[0].Pane
    $cur = [FeWin2]::Rect($xlh)
    $offset = $cur[3] - $o.CellArea[3]
    $newBottom = $raw[1] + 14 + $offset
    [void][FeWin2]::SetWindowPos($xlh, [IntPtr]::Zero, $cur[0], $cur[1], ($cur[2] - $cur[0]), ($newBottom - $cur[1]), 0x0004 -bor 0x0010)
    Start-Sleep -Milliseconds 900
    $o = Get-PaneOracle $fe
    $raw = Get-RangeScreenRect $fe $ws.Range('K31') $o.Panes[0].Pane
    Note "K31 raw rect [$($raw -join ',')], cell area [$($o.CellArea -join ',')]"
    Check 'S13 precondition: K31 is cut by the bottom (and right) edge of the cell area' ($raw[3] -gt $o.CellArea[3] + 2 -and $raw[1] -lt $o.CellArea[3] - 4) @{ k31 = $raw; area = $o.CellArea }
    $null = Verify-Overlay $fe 'S13 K31 cut by bottom and right edges' $script:K33Boxes -SaveShot
    # move the Excel window: the borders follow
    $cur = [FeWin2]::Rect($xlh)
    [void][FeWin2]::SetWindowPos($xlh, [IntPtr]::Zero, ($cur[0] + 150), ($cur[1] + 60), ($cur[2] - $cur[0]), ($cur[3] - $cur[1]), 0x0004 -bor 0x0010)
    Start-Sleep -Milliseconds 900
    $null = Verify-Overlay $fe 'S13 after moving the Excel window' $script:K33Boxes
    $xl.WindowState = -4137
    Start-Sleep -Milliseconds 900
    $null = Verify-Overlay $fe 'S13 maximized again' $script:K33Boxes

    # ---- another sheet: overlay hidden, then back ----
    $fe.Excel.ActiveWorkbook.Worksheets.Item('Names_BA').Activate() | Out-Null
    Start-Sleep -Milliseconds 700
    $s = Get-FeState $fe
    Check 'S13 other sheet active: overlay hidden' (-not $s.overlayVisible -and (Get-Overlay $s).Count -eq 0) @{ visible = $s.overlayVisible; rects = (Get-Overlay $s).Count }
    $pair = Get-OverlayPair $fe $s
    $stats = [FeWin2]::DiffStats($pair.On, $pair.Off); $pair.On.Dispose(); if ($pair.Off -ne $pair.On) { $pair.Off.Dispose() }
    Check 'S13 other sheet active: no overlay pixels' ($stats[0] -eq 0 -and $s.overlayHwnd -ne 0) "changed=$($stats[0]) hwnd=$($s.overlayHwnd)"
    $fe.Excel.ActiveWorkbook.Worksheets.Item('Val_FO').Activate() | Out-Null
    Start-Sleep -Milliseconds 700
    $null = Verify-Overlay $fe 'S13 back on the root sheet' $script:K33Boxes

    # ---- z-order: the explorer window is above the borders ----
    $s = Get-FeState $fe
    $eh = [IntPtr][int64]$s.explorerHwnd
    $origBounds = $s.view.screenBounds
    $exp = Get-RangeScreenRect $fe $ws.Range('I27')
    $k31 = Get-RangeScreenRect $fe $ws.Range('K31')
    $ew = ([FeWin2]::Rect($eh))[2] - ([FeWin2]::Rect($eh))[0]
    # right edge between the I27 and K31 boxes: covers I27 only
    [void][FeWin2]::SetWindowPos($eh, [IntPtr]::Zero, ([int](($exp[2] + $k31[0]) / 2) - $ew), ($exp[1] - 70), 0, 0, 0x0001 -bor 0x0004 -bor 0x0010)
    Start-Sleep -Milliseconds 700
    $s = Get-FeState $fe
    $er = [FeWin2]::Rect($eh)
    $covers = ($er[0] -lt $exp[0]) -and ($er[2] -gt $exp[2]) -and ($er[1] -lt $exp[1]) -and ($er[3] -gt $exp[3])
    Check 'S13 precondition: the explorer window now covers the I27 box' $covers @{ explorer = $er; i27 = $exp }
    $pair = Get-OverlayPair $fe $s
    $inside = Get-OverlayColorDelta $pair @(($er[0] + 8), ($er[1] + 8), ($er[2] - 8), ($er[3] - 8)) @(0x0000FF, 0x9900CC)
    $total = [FeWin2]::DiffStats($pair.On, $pair.Off)
    $pair.On.Save((Join-Path $fe.OutDir 'S13-explorer-over-box.png'), [Drawing.Imaging.ImageFormat]::Png)
    $pair.On.Dispose(); if ($pair.Off -ne $pair.On) { $pair.Off.Dispose() }
    Check 'S13 the explorer window is above the overlay (no overlay pixels inside the explorer window)' ($inside -eq 0) "overlay pixels inside the explorer=$inside, elsewhere=$($total[0])"
    Check 'S13 the K31 border is still painted (not covered)' ($total[0] -gt 20) "changed=$($total[0])"
    # put it back at the top right
    [void][FeWin2]::SetWindowPos($eh, [IntPtr]::Zero, $origBounds.x, $origBounds.y, 0, 0, 0x0001 -bor 0x0004 -bor 0x0010)
    Start-Sleep -Milliseconds 500

    # ---- Ctrl+H (Find and Replace) stays above the borders ----
    $sBefore = Get-FeState $fe
    Click-Cell $fe 'K33'     # real click: Excel grid gets the keyboard focus (K33 is the root: nothing changes)
    Start-Sleep -Milliseconds 500
    Press-Key $script:VK.H @($script:VK.Ctrl)
    Start-Sleep -Milliseconds 1200
    $dlg = $null
    foreach ($line in [FeWin2]::TopLevel([uint32]$fe.Pid)) { $p = $line.Split('|'); if ($p[1] -like 'bosa_sdm*' -or $p[2] -like 'Find and Replace*') { $dlg = @{ hwnd = [int64]$p[0]; title = $p[2]; rect = ($p[3].Split(',') | ForEach-Object { [int]$_ }) } } }
    Check 'S13 Ctrl+H opened the Find and Replace dialog (real keyboard)' ($null -ne $dlg) $dlg
    if ($dlg) {
        $exp = Get-RangeScreenRect $fe $ws.Range('I27')
        $w = $dlg.rect[2] - $dlg.rect[0]; $h = $dlg.rect[3] - $dlg.rect[1]
        [void][FeWin2]::SetWindowPos([IntPtr]$dlg.hwnd, [IntPtr]::Zero, ($exp[0] - 100), ($exp[1] - 60), 0, 0, 0x0001 -bor 0x0004 -bor 0x0010)
        Start-Sleep -Milliseconds 700
        $dr = [FeWin2]::Rect([IntPtr]$dlg.hwnd)
        $pair = Get-OverlayPair $fe $sBefore
        $box = [FeWin2]::DiffBox($pair.On, $pair.Off, $dr[0] + 20, $dr[1] + 20, $dr[2] - 20, $dr[3] - 20)
        # (a blinking text caret in the dialog can differ between captures; box-colored pixels are what matters)
        $inside = Get-OverlayColorDelta $pair @(($dr[0] + 20), ($dr[1] + 20), ($dr[2] - 20), ($dr[3] - 20)) @(0x0000FF, 0x9900CC)
        Note "pixels changed inside the dialog: $($box[0]) (bbox $($box[1]),$($box[2])-$($box[3]),$($box[4])); box-colored: $inside"
        $total = [FeWin2]::DiffStats($pair.On, $pair.Off)
        $pair.On.Save((Join-Path $fe.OutDir 'S13-find-dialog-over-box.png'), [Drawing.Imaging.ImageFormat]::Png)
        $pair.On.Dispose(); if ($pair.Off -ne $pair.On) { $pair.Off.Dispose() }
        $covers = ($dr[0] -lt $exp[0]) -and ($dr[2] -gt $exp[2]) -and ($dr[1] -lt $exp[1]) -and ($dr[3] -gt $exp[3])
        Check 'S13 precondition: the dialog covers the I27 box' $covers @{ dialog = $dr; i27 = $exp }
        Check 'S13 the Find and Replace dialog is above the borders (no overlay pixels inside it)' ($inside -eq 0) "overlay pixels inside the dialog=$inside, elsewhere=$($total[0])"
        Press-Key $script:VK.Esc
        Start-Sleep -Milliseconds 600
        # the dialog may still be open after one Esc (focus): close it by WM_CLOSE
        [void][FeWin2]::PostMessage([IntPtr]$dlg.hwnd, 0x10, [IntPtr]::Zero, [IntPtr]::Zero)
        Start-Sleep -Milliseconds 500
    }
    $s = Get-FeState $fe
    Check 'S13 the explorer is still open and drawing after the dialog' ($s.isOpen -and $s.overlayVisible) @{ open = $s.isOpen; overlay = $s.overlayVisible }
    Check-LogClean
    End-Scenario
}
