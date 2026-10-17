# Overlay verification helpers: panes, expected boxes and the "overlay on vs. off" screenshot pair (dot-sourced by Run-Scenarios.ps1).

function Get-ScreenBitmap { [FeWin2]::Capture([FeWin2]::GetSystemMetrics(76), [FeWin2]::GetSystemMetrics(77), [FeWin2]::GetSystemMetrics(78), [FeWin2]::GetSystemMetrics(79)) }

# Screenshot pair with the overlay window shown (On) and hidden by hand (Off); the difference is exactly what it paints.
# The pointer is parked away so no hover state differs between the two captures.
function Get-OverlayPair($fe, $state) {
    Move-Mouse 100 1900
    Start-Sleep -Milliseconds 150
    if ($state.overlayHwnd -eq 0 -or -not $state.overlayVisible) { $on = Get-ScreenBitmap; return @{ On = $on; Off = $on; Hwnd = 0 } }
    $h = [IntPtr][int64]$state.overlayHwnd
    $taskbar = @(0, ([FeWin2]::GetSystemMetrics(79) - 70), [FeWin2]::GetSystemMetrics(78), [FeWin2]::GetSystemMetrics(79))
    for ($attempt = 1; $attempt -le 4; $attempt++) {
        $on = Get-ScreenBitmap
        [void][FeWin2]::ShowWindow($h, 0); Start-Sleep -Milliseconds 250
        $off = Get-ScreenBitmap
        [void][FeWin2]::ShowWindow($h, 4); Start-Sleep -Milliseconds 250
        $on2 = Get-ScreenBitmap
        # the screen must be stable (nothing else repainting) between the two "on" captures, otherwise retry
        $unstable = [FeWin2]::DiffOutside($on, $on2, [int[]]$taskbar)
        $on2.Dispose()
        if ($unstable[0] -eq 0) { return @{ On = $on; Off = $off; Hwnd = $h } }
        $on.Dispose(); $off.Dispose()
        Start-Sleep -Milliseconds 600
    }
    $on = Get-ScreenBitmap
    [void][FeWin2]::ShowWindow($h, 0); Start-Sleep -Milliseconds 250
    $off = Get-ScreenBitmap
    [void][FeWin2]::ShowWindow($h, 4); Start-Sleep -Milliseconds 150
    @{ On = $on; Off = $off; Hwnd = $h }
}

# Cell-area rectangles of the panes of the active window (physical pixels) computed from the Excel object model only.
function Get-PaneOracle($fe) {
    $w = $fe.Excel.ActiveWindow
    $dpi = [FeWin2]::GetDpiForWindow([IntPtr]$w.Hwnd)
    $n = $w.Panes.Count
    $list = @()
    for ($i = 1; $i -le $n; $i++) {
        $p = $w.Panes.Item($i); $vr = $p.VisibleRange
        $l = Convert-PointsToPx $p 'X' ([double]$vr.Left); $t = Convert-PointsToPx $p 'Y' ([double]$vr.Top)
        $r = Convert-PointsToPx $p 'X' ([double]$vr.Left + [double]$vr.Width); $b = Convert-PointsToPx $p 'Y' ([double]$vr.Top + [double]$vr.Height)
        $list += [pscustomobject]@{ Index = $i; Pane = $p; Rect = @($l, $t, $r, $b) }
    }
    $x0 = ($list | ForEach-Object { $_.Rect[0] } | Measure-Object -Minimum).Minimum
    $y0 = ($list | ForEach-Object { $_.Rect[1] } | Measure-Object -Minimum).Minimum
    $ur = $x0 + [int][math]::Round([double]$w.UsableWidth * $dpi / 72.0)
    $ub = $y0 + [int][math]::Round([double]$w.UsableHeight * $dpi / 72.0)
    foreach ($e in $list) { $e.Rect = @($e.Rect[0], $e.Rect[1], [math]::Min($e.Rect[2], $ur), [math]::Min($e.Rect[3], $ub)) }
    foreach ($a in $list) {
        foreach ($b in $list) {
            if ($a.Index -eq $b.Index) { continue }
            $vOverlap = $a.Rect[1] -lt $b.Rect[3] -and $b.Rect[1] -lt $a.Rect[3]
            $hOverlap = $a.Rect[0] -lt $b.Rect[2] -and $b.Rect[0] -lt $a.Rect[2]
            if ($b.Rect[0] -gt $a.Rect[0] -and $vOverlap) { $a.Rect[2] = [math]::Min($a.Rect[2], $b.Rect[0]) }
            if ($b.Rect[1] -gt $a.Rect[1] -and $hOverlap) { $a.Rect[3] = [math]::Min($a.Rect[3], $b.Rect[1]) }
        }
    }
    @{ Panes = $list; CellArea = @($x0, $y0, $ur, $ub); Dpi = $dpi }
}

function Intersect-Rect($a, $b) {
    $l = [math]::Max($a[0], $b[0]); $t = [math]::Max($a[1], $b[1]); $r = [math]::Min($a[2], $b[2]); $bt = [math]::Min($a[3], $b[3])
    if ($r -le $l -or $bt -le $t) { return $null }
    , @($l, $t, $r, $bt)
}

# Expected clipped rectangles for the given boxes (list of @{ Address; Rgb }) on the active sheet, one per pane they appear in.
function Get-ExpectedBoxes($fe, $Boxes, $Oracle) {
    $ws = $fe.Excel.ActiveSheet
    $out = @()
    foreach ($box in $Boxes) {
        $range = $ws.Range($box.Address)
        if ($range.EntireRow.Hidden -eq $true) { continue }
        foreach ($pe in $Oracle.Panes) {
            $raw = Get-RangeScreenRect $fe $range $pe.Pane
            $clipped = Intersect-Rect $raw $pe.Rect
            if ($clipped) { $out += [pscustomobject]@{ Address = $box.Address; Rgb = $box.Rgb; Pane = $pe.Index; Rect = $clipped; Raw = $raw } }
        }
    }
    $out
}

# Full geometry verification of the overlay for the current view; records the checks.
function Verify-Overlay($fe, [string]$Label, $Boxes, [switch]$SaveShot, [int]$Tol = 3) {
    $state = Get-FeState $fe
    $oracle = Get-PaneOracle $fe
    $exp = @(Get-ExpectedBoxes $fe $Boxes $oracle)
    $rep = Get-Overlay $state
    $pair = Get-OverlayPair $fe $state
    try {
        $unmatched = @()
        $used = @{}
        foreach ($e in $exp) {
            $found = $false
            for ($i = 0; $i -lt $rep.Count; $i++) {
                if ($used.ContainsKey($i)) { continue }
                if ($rep[$i].rgb -eq $e.Rgb -and (Rect-Near (Overlay-Rect $rep[$i]) $e.Rect $Tol)) { $used[$i] = $true; $found = $true; break }
            }
            if (-not $found) { $unmatched += "$($e.Address)@pane$($e.Pane) want [$($e.Rect -join ',')]" }
        }
        Check "$Label : reported boxes match the expected ones ($($exp.Count) expected)" ($unmatched.Count -eq 0 -and $rep.Count -eq $exp.Count) @{ reported = @($rep | ForEach-Object { "$(Hex $_.rgb) [$((Overlay-Rect $_) -join ',')]" }); unmatched = $unmatched; expected = @($exp | ForEach-Object { "$($_.Address) pane$($_.Pane) [$($_.Rect -join ',')]" }) }
        if ($pair.Hwnd -eq 0) {
            Check "$Label : overlay window is hidden / absent, as expected for $($exp.Count) boxes" ($exp.Count -eq 0) @{ visible = $state.overlayVisible; expected = $exp.Count }
        }
        else {
            $zone = @()
            # the taskbar clock may tick between the two captures: ignore the bottom strip of the screen
            $zone += @(0, ([FeWin2]::GetSystemMetrics(79) - 70), [FeWin2]::GetSystemMetrics(78), [FeWin2]::GetSystemMetrics(79))
            foreach ($e in $exp) {
                # border zone of the box, limited to its pane: nothing may spill onto headers, scroll bars or neighbouring panes
                $pane = ($oracle.Panes | Where-Object { $_.Index -eq $e.Pane }).Rect
                $z = Intersect-Rect @(($e.Rect[0] - 6), ($e.Rect[1] - 6), ($e.Rect[2] + 6), ($e.Rect[3] + 6)) @(($pane[0] - 2), ($pane[1] - 2), ($pane[2] + 2), ($pane[3] + 2))
                if ($z) { $zone += $z }
            }
            $stats = [FeWin2]::DiffStats($pair.On, $pair.Off)
            $far = [FeWin2]::DiffOutside($pair.On, $pair.Off, [int[]]$zone)
            Check "$Label : paints nothing outside the expected box borders (headers, scrollbars, tabs, other cells untouched)" ($far[0] -eq 0) @{ changedTotal = $stats[0]; outside = $far[0]; firstOutsideAt = @($far[1], $far[2]) }
            foreach ($e in $exp) {
                $n = [FeWin2]::DiffInside($pair.On, $pair.Off, $e.Rect[0] - 6, $e.Rect[1] - 6, $e.Rect[2] + 6, $e.Rect[3] + 6)
                Check "$Label : border of $($e.Address) (pane $($e.Pane)) is painted" ($n -ge 6) "changed pixels=$n"
                $cnt = [FeWin2]::CountColor($pair.On, $e.Rect[0] - 6, $e.Rect[1] - 6, ($e.Rect[2] - $e.Rect[0] + 12), ($e.Rect[3] - $e.Rect[1] + 12), [int]$e.Rgb, 6)
                Check "$Label : $($e.Address) border pixels have the box color $(Hex $e.Rgb)" ($cnt -ge 4) "pixels=$cnt"
            }
            if ($exp.Count -eq 0) { Check "$Label : nothing painted at all" ($stats[0] -eq 0) "changed=$($stats[0])" }
        }
        if ($SaveShot) {
            $fe.Shot++
            $f = Join-Path $fe.OutDir ('{0:00}-{1}.png' -f $fe.Shot, ($Label -replace '[^\w-]', '_'))
            $pair.On.Save($f, [Drawing.Imaging.ImageFormat]::Png)
        }
    }
    finally { $pair.On.Dispose(); if ($pair.Off -and $pair.Off -ne $pair.On) { $pair.Off.Dispose() } }
    @{ Expected = $exp; Reported = $rep; Oracle = $oracle; State = $state }
}

# Number of pixels with any of the box colors inside the rectangle that appear only with the overlay shown.
function Get-OverlayColorDelta($pair, $rect, $colors) {
    $sum = 0
    foreach ($c in $colors) {
        $w = $rect[2] - $rect[0]; $h = $rect[3] - $rect[1]
        $a = [FeWin2]::CountColor($pair.On, $rect[0], $rect[1], $w, $h, [int]$c, 6)
        $b = [FeWin2]::CountColor($pair.Off, $rect[0], $rect[1], $w, $h, [int]$c, 6)
        $sum += [math]::Max(0, $a - $b)
    }
    $sum
}
