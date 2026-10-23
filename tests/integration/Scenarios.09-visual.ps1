# S20: visual fidelity against docs/referencias (reference frames scaled x1.5 next to our window at 150%).

function New-Comparison([string]$RefFile, [string]$OursFile, [string]$OutFile, [int[]]$OursRect) {
    Add-Type -AssemblyName System.Drawing
    $ref = [Drawing.Image]::FromFile($RefFile)
    $ours = [Drawing.Image]::FromFile($OursFile)
    $crop = New-Object Drawing.Bitmap $OursRect[2], $OursRect[3]
    $g = [Drawing.Graphics]::FromImage($crop)
    $g.DrawImage($ours, (New-Object Drawing.Rectangle 0, 0, $OursRect[2], $OursRect[3]), (New-Object Drawing.Rectangle $OursRect[0], $OursRect[1], $OursRect[2], $OursRect[3]), [Drawing.GraphicsUnit]::Pixel)
    $g.Dispose()
    $rw = [int]($ref.Width * 1.5); $rh = [int]($ref.Height * 1.5)
    $w = [math]::Max($rw, $crop.Width); $h = $rh + $crop.Height + 24
    $bmp = New-Object Drawing.Bitmap $w, $h
    $g = [Drawing.Graphics]::FromImage($bmp)
    $g.Clear([Drawing.Color]::White)
    $g.InterpolationMode = [Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.DrawImage($ref, 0, 0, $rw, $rh)
    $g.DrawImage($crop, 0, $rh + 24)
    $g.Dispose()
    $bmp.Save($OutFile, [Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose(); $ref.Dispose(); $ours.Dispose(); $crop.Dispose()
}

function Capture-ExplorerComparison($fe, [string]$Name, [string]$RefName, $state, [int]$ExtraBottom = 110) {
    $b = $state.view.screenBounds
    $x = [math]::Max(0, $b.x - 10); $y = [math]::Max(0, $b.y - 6)
    $full = Join-Path $fe.OutDir ("S20-$Name-ours.png")
    $bmp = Get-ScreenBitmap
    $bmp.Save($full, [Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    $rect = @($x, $y, ($b.width + 20), ($b.height + 12 + $ExtraBottom))
    $cmp = Join-Path $fe.OutDir ("S20-$Name-compare.png")
    New-Comparison (Join-Path (Join-Path $script:FeRoot 'docs\referencias') $RefName) $full $cmp $rect
    $cmp
}

function Hover-Real($fe, $state, [int]$Index) {
    $p = Char-Point $state $Index
    Move-Mouse $p.X $p.Y
    $p
}

function Find-TooltipRect($fe, $before, $state) {
    $new = Get-NewWindows $fe $before $state
    $tip = $new | Where-Object { $_.rect[2] - $_.rect[0] -gt 40 -and $_.rect[3] - $_.rect[1] -gt 20 } | Select-Object -First 1
    if ($tip) { $tip.rect } else { $null }
}

function S20 {
    Begin-Scenario 'S20' 'Visual fidelity: window states of docs/referencias 2011-site-01/02/03 and 2011-oficial-08/14/16 (composites saved for review)'
    Reset-Env
    $fe = $script:fe
    # a clean desktop around the window: Excel on the first sheet, pointer parked
    $cases = @()

    # ---- 2011-site-01: hover on 'Inventory Forcast'!C$6 (Stores!F4) ----
    $s = Open-Explorer $fe 'Stores' 'F4'
    $before = Get-TopLevelSet $fe
    $i = Ix $s 'Forcast' 1 2
    Hover-Real $fe $s $i | Out-Null
    $r = Wait-Tooltip $fe 3000; Start-Sleep -Milliseconds 400
    $s = Get-FeState $fe
    $tipRect = Find-TooltipRect $fe $before $s
    Check 'site-01: hover span is the quoted sheet reference, tooltip 25' ((Get-HoverText $s) -ceq "'Inventory Forcast'!C`$6" -and $s.view.tooltipText -ceq '25') @{ hover = (Get-HoverText $s); tip = $s.view.tooltipText }
    # gray background and underline in the node color (purple for 'Inventory...': third color of F4 is C$6? color index 2)
    $sp = Get-SpanRect $s ([int]$s.view.hoverStart) ([int]$s.view.hoverLength)
    $bmp = [FeWin2]::Capture($sp.X, $sp.Y, $sp.W, $sp.H)
    $gray = [FeWin2]::CountColor($bmp, 0, 0, $sp.W, $sp.H, $script:Gray, 3)
    $bmp.Dispose()
    Check 'site-01: #DCDCDC highlight behind the reference' ($gray -gt 300) "gray px=$gray"
    $cmp = Capture-ExplorerComparison $fe 'site-01' '2011-site-01-hyperlink-hover.png' $s
    Note "site-01 -> $cmp"
    if ($tipRect) { Check-Tooltip $fe $tipRect 'site-01' }
    $cases += $cmp

    # ---- 2011-site-02: hover on ^ of K32 ----
    $s = Open-Explorer $fe 'Val_FO' 'K32'
    $before = Get-TopLevelSet $fe
    Hover-Real $fe $s (Ix $s '^') | Out-Null
    $r = Wait-Tooltip $fe 3000; Start-Sleep -Milliseconds 400
    $s = Get-FeState $fe
    $tipRect = Find-TooltipRect $fe $before $s
    Check 'site-02: span (-($I26/Hundred)+1)^(K$31) highlighted' ((Get-HoverText $s) -ceq '(-($I26/Hundred)+1)^(K$31)') (Get-HoverText $s)
    $cmp = Capture-ExplorerComparison $fe 'site-02' '2011-site-02-fragment-solving.png' $s
    if ($tipRect) { Check-Tooltip $fe $tipRect 'site-02' }
    $cases += $cmp

    # ---- 2011-site-03: hover on the ) closing '((365' of K34 ----
    $s = Open-Explorer $fe 'Val_FO' 'K34'
    $before = Get-TopLevelSet $fe
    $inner = '((365*(($I27/Hundred)+1)^(K$31)))'
    $idx = (Find-FeIndex $s.view.displayText $inner 1 0) + $inner.Length - 1
    Hover-Real $fe $s $idx | Out-Null
    $r = Wait-Tooltip $fe 3000; Start-Sleep -Milliseconds 400
    $s = Get-FeState $fe
    $tipRect = Find-TooltipRect $fe $before $s
    Check 'site-03: bracket group highlighted' ((Get-HoverText $s) -ceq $inner) (Get-HoverText $s)
    $cmp = Capture-ExplorerComparison $fe 'site-03' '2011-site-03-bracket-forest.png' $s
    if ($tipRect) { Check-Tooltip $fe $tipRect 'site-03' }
    $cases += $cmp

    # ---- 2011-oficial-08: a reference gold after a click, pointer on it (K$31 in the original; it now drills on one
    #      click, so the gold look is checked on $I27, whose cell has no formula) ----
    $s = Open-Explorer $fe 'Val_FO' 'K33'
    [FeWin]::Foreground([IntPtr][int64]$s.explorerHwnd); Start-Sleep -Milliseconds 300
    $i = Ix $s '$I27' 1 1
    $p = Hover-Real $fe $s $i
    Click-At $p.X $p.Y
    $r = Wait-Timed $fe { (Get-GoldText $_) -ceq '$I27' } 3000
    Start-Sleep -Milliseconds 900
    $s = Get-FeState $fe
    $sp = Get-SpanRect $s ([int]$s.view.goldStart) ([int]$s.view.goldLength)
    $bmp = [FeWin2]::Capture($sp.X, $sp.Y, $sp.W, $sp.H)
    $gold = [FeWin2]::CountColor($bmp, 0, 0, $sp.W, $sp.H, $script:Gold, 3)
    $gray = [FeWin2]::CountColor($bmp, 0, 0, $sp.W, $sp.H, $script:Gray, 3)
    $bmp.Dispose()
    Check 'oficial-08: gold #FFD700 behind $I27, no gray, no tooltip' ($gold -gt 150 -and $gray -lt 15 -and -not $s.view.tooltipVisible) "gold=$gold gray=$gray tip=$($s.view.tooltipVisible)"
    $cmp = Capture-ExplorerComparison $fe 'oficial-08' '2011-oficial-08-1m55-ref-dourada-antes-drill.png' $s 40
    $cases += $cmp

    # ---- 2011-oficial-14: hover on OFFSET in the J31 formula (range tooltip, underline) ----
    $s = Open-Explorer $fe 'Val_FO' 'J31'
    $before = Get-TopLevelSet $fe
    $i = Ix $s 'OFFSET' 1 2
    Hover-Real $fe $s $i | Out-Null
    $r = Wait-Tooltip $fe 3000; Start-Sleep -Milliseconds 500
    $s = Get-FeState $fe
    $tipRect = Find-TooltipRect $fe $before $s
    Check 'oficial-14: OFFSET(...) highlighted, link, range tooltip' ((Get-HoverText $s) -ceq 'OFFSET($I$19,0,0,1,Aqui_Per)' -and $s.view.hoverIsLink -and $s.view.tooltipText -ceq ("I19:J19 containing`n[1x2] , 0.252054794520548")) @{ hover = (Get-HoverText $s); tip = $s.view.tooltipText }
    $cmp = Capture-ExplorerComparison $fe 'oficial-14' '2011-oficial-14-2m19-hover-OFFSET-tooltip-intervalo.png' $s 140
    if ($tipRect) { Check-Tooltip $fe $tipRect 'oficial-14' }
    $cases += $cmp

    # ---- 2011-oficial-16: OFFSET gold, hovering Back shows the button tooltip ----
    [FeWin]::Foreground([IntPtr][int64]$s.explorerHwnd); Start-Sleep -Milliseconds 200
    $p = Char-Point $s $i
    Click-At $p.X $p.Y
    $r = Wait-Timed $fe { (Get-GoldText $_) -like 'OFFSET(*' } 3000
    $s = Get-FeState $fe
    $b = $s.buttons.back
    $before = Get-TopLevelSet $fe
    Move-Mouse ([int](($b[0] + $b[2]) / 2)) ([int](($b[1] + $b[3]) / 2))
    Start-Sleep -Milliseconds 1800
    $s = Get-FeState $fe
    $cmp = Capture-ExplorerComparison $fe 'oficial-16' '2011-oficial-16-2m34-tooltip-back.png' $s 80
    $tip = Get-NewWindows $fe $before $s | Select-Object -First 1
    Check 'oficial-16: the Back button shows a tooltip window ("Retraces your steps")' ($null -ne $tip) ($tip | ForEach-Object { "$($_.cls) $($_.rect -join ',')" })
    $cases += $cmp
    Move-Mouse 100 1900
    Check-LogClean
    Note ("composites: " + ($cases -join ' | '))
    End-Scenario
}

# Windows 7 tooltip look: #767676 border, gradient #FFFFFF -> #E5E6F0, text #575757.
function Check-Tooltip($fe, $rect, [string]$Label) {
    $w = $rect[2] - $rect[0]; $h = $rect[3] - $rect[1]
    $bmp = [FeWin2]::Capture($rect[0], $rect[1], $w, $h)
    try {
        $border = [FeWin2]::CountColor($bmp, 0, 0, $w, $h, 0x767676, 8)
        $bottom = [FeWin2]::CountColor($bmp, 0, 0, $w, $h, 0xE5E6F0, 10)
        $white = [FeWin2]::CountColor($bmp, 0, 0, $w, $h, 0xFFFFFF, 3)
        $text = [FeWin2]::CountColor($bmp, 0, 0, $w, $h, 0x575757, 30)
        Check "$Label : tooltip has the #767676 border, the white-to-#E5E6F0 gradient and #575757 text" ($border -gt 60 -and $bottom -gt 40 -and $white -gt 40 -and $text -gt 40) @{ rect = $rect; border = $border; bottom = $bottom; white = $white; text = $text } -Level soft
    }
    finally { $bmp.Dispose() }
}
