# S1 and S2: opening on Val_FO!K33, and the hover rules / tooltips on the video formulas (through the test hooks).

function S1 {
    Begin-Scenario 'S1' 'Open on Val_FO!K33: text, colors, precedent borders, Back, first-open position'
    Reset-Env
    $fe = $script:fe
    $s = Open-Explorer $fe 'Val_FO' 'K33'
    $t = $s.view.displayText
    Check 'explorer is open and shows the formula text' ($s.isOpen -and $t -ceq '=1/(((($I27/Hundred)+1)^(K$31)))') $t
    Check 'root is Val_FO!$K$33' ($s.session.root.sheetName -eq 'Val_FO' -and $s.session.root.address -eq '$K$33') $s.session.root
    Check 'Back is disabled' (-not $s.view.backEnabled -and -not $s.session.backEnabled) @{ view = $s.view.backEnabled; session = $s.session.backEnabled }
    Check 'view is interactive, not loading' ($s.view.interactive -and -not $s.view.isLoading) @{ interactive = $s.view.interactive; loading = $s.view.isLoading }
    Save-FeShot $fe 'S1-open' | Out-Null

    # Colors painted on screen (independent of the add-in's own color map): blue $I27, green Hundred, purple K$31.
    $iI27 = Ix $s '$I27'; $iHun = Ix $s 'Hundred'; $iK31 = Ix $s 'K$31'
    $blue = Count-SpanColor $s $iI27 4 $script:ExpectedPalette.Blue 40
    $green = Count-SpanColor $s $iHun 7 $script:ExpectedPalette.Green 40
    $purple = Count-SpanColor $s $iK31 4 $script:ExpectedPalette.Purple 40
    Check 'text color of $I27 is blue (#0000FF)' ($blue -ge 5) "pixels=$blue"
    Check 'text color of Hundred is green (#008000)' ($green -ge 5) "pixels=$green"
    Check 'text color of K$31 is purple (#9900CC)' ($purple -ge 5) "pixels=$purple"
    $inkEq = Get-SpanInk $s 0 1   # the '='
    $inkSlash = Get-SpanInk $s 2 1
    Check "'=' and '/' are black" ((Color-Dist $inkEq 0) -le 40 -and (Color-Dist $inkSlash 0) -le 40) @{ eq = (Hex $inkEq); slash = (Hex $inkSlash) }
    # No colored pixels leak into the black parts: the '(((( ' run has no blue/green/purple ink.
    $leak = (Count-SpanColor $s 3 4 $script:ExpectedPalette.Blue 40) + (Count-SpanColor $s 3 4 $script:ExpectedPalette.Green 40) + (Count-SpanColor $s 3 4 $script:ExpectedPalette.Purple 40)
    Check 'parentheses are black (no colored ink)' ($leak -eq 0) "colored pixels=$leak"

    # Precedent borders: I27 (same sheet, blue) and K31 (purple) only. Hundred lives on another sheet.
    $ws = $fe.Excel.ActiveSheet
    $expI27 = Get-RangeScreenRect $fe $ws.Range('I27')
    $expK31 = Get-RangeScreenRect $fe $ws.Range('K31')
    $ov = Get-Overlay $s
    Check 'exactly two precedent boxes (I27, K31); none for Hundred or the root' ($ov.Count -eq 2) $ov.Count
    $oBlue = $ov | Where-Object { $_.rgb -eq $script:ExpectedPalette.Blue }
    $oPurple = $ov | Where-Object { $_.rgb -eq $script:ExpectedPalette.Purple }
    Check 'blue box exists, purple box exists' ($oBlue -and $oPurple) ($ov | ForEach-Object { Hex $_.rgb })
    if ($oBlue) { Check "blue box matches I27's screen rect within 2 px" (Rect-Near (Overlay-Rect $oBlue) $expI27) @{ overlay = (Overlay-Rect $oBlue); expected = $expI27 } }
    if ($oPurple) { Check "purple box matches K31's screen rect within 2 px" (Rect-Near (Overlay-Rect $oPurple) $expK31) @{ overlay = (Overlay-Rect $oPurple); expected = $expK31 } }
    # The pixels really are there: sample a few pixels around the middle of I27's left edge and K31's top edge.
    $bmp = [FeWin2]::Capture($expI27[0] - 4, [int](($expI27[1] + $expI27[3]) / 2) - 3, 9, 7)
    $nBlue = [FeWin2]::CountColor($bmp, 0, 0, 9, 7, $script:ExpectedPalette.Blue, 24); $bmp.Dispose()
    $bmp = [FeWin2]::Capture([int](($expK31[0] + $expK31[2]) / 2) - 3, $expK31[1] - 4, 7, 9)
    $nPurple = [FeWin2]::CountColor($bmp, 0, 0, 7, 9, $script:ExpectedPalette.Purple, 24); $bmp.Dispose()
    Check 'blue border pixels are on screen at I27' ($nBlue -ge 3) "pixels=$nBlue"
    Check 'purple border pixels are on screen at K31' ($nPurple -ge 3) "pixels=$nPurple"
    # No border on cells that are not precedents (e.g. I26, I28, J31).
    foreach ($c in 'I26', 'I28', 'J31') {
        $r = Get-RangeScreenRect $fe $ws.Range($c)
        $b = [FeWin2]::Capture($r[0] - 3, [int](($r[1] + $r[3]) / 2) - 2, 7, 5)
        $n = ([FeWin2]::CountColor($b, 0, 0, 7, 5, $script:ExpectedPalette.Blue, 24)) + ([FeWin2]::CountColor($b, 0, 0, 7, 5, $script:ExpectedPalette.Purple, 24)) + ([FeWin2]::CountColor($b, 0, 0, 7, 5, $script:ExpectedPalette.Green, 24)); $b.Dispose()
        Check "no colored border pixels at $c (not a precedent)" ($n -eq 0) "pixels=$n"
    }
    # Overlay window facts: owned by Excel, visible, click-through, below the explorer.
    Check 'overlay window exists and is visible' ($s.overlayHwnd -ne 0 -and $s.overlayVisible) @{ hwnd = $s.overlayHwnd; visible = $s.overlayVisible }

    # First-open position: top-right inside Excel's window.
    $ex = Get-Rect (Get-ExcelHwnd $fe)
    $b = $s.view.screenBounds
    $right = $b.x + $b.width
    $inside = ($b.x -ge $ex[0]) -and ($right -le $ex[2]) -and ($b.y -ge $ex[1]) -and (($b.y + $b.height) -le $ex[3])
    Check 'window lies inside the Excel window' $inside @{ explorer = @($b.x, $b.y, $right, ($b.y + $b.height)); excel = $ex }
    $dxRight = $ex[2] - $right; $dyTop = $b.y - $ex[1]
    Check 'window is at the top-right (right edge within 90 px of Excel right, top within 40..300 px of Excel top)' (($dxRight -ge 0 -and $dxRight -le 90) -and ($dyTop -ge 40 -and $dyTop -le 300)) @{ fromRight = $dxRight; fromTop = $dyTop }
    $dpiScale = $s.view.pixelsPerDip
    Check 'WPF window is per-monitor DPI aware (pixelsPerDip = 1.5 on this 150% screen)' ($dpiScale -eq 1.5) $dpiScale -Level soft
    Check-LogClean
    End-Scenario
}

# One hover expectation: where to hover, the exact highlighted substring, link flag and tooltip text (null = no tooltip).
function New-Hover($At, $Expect, $Tip, [bool]$Link, [int]$Occ = 1, [int]$Off = 0, [int]$Len = 0) {
    @{ At = $At; Occ = $Occ; Off = $Off; Expect = $Expect; Tip = $Tip; Link = $Link }
}

function Invoke-HoverCases($fe, [string]$Sheet, [string]$Cell, $Cases, [string]$Label) {
    $s = Open-Explorer $fe $Sheet $Cell
    foreach ($c in $Cases) {
        $text = $s.view.displayText
        $occ = 1; if ($c.Occ) { $occ = [int]$c.Occ }
        $off = 0; if ($c.Off) { $off = [int]$c.Off }
        $idx = Find-FeIndex $text $c.At $occ $off
        Hover-Char $fe $idx
        $tipExpected = $c.Tip
        if ($null -ne $tipExpected) { $s2 = Wait-FeState $fe { $_.view.tooltipVisible -and $_.view.tooltipText } 2500 } else { Start-Sleep -Milliseconds 800; $s2 = Get-FeState $fe }
        $hover = Get-HoverText $s2
        $name = "$Label hover '$($c.At)'@$idx"
        $okSpan = ($hover -ceq $c.Expect)
        $okLink = ($s2.view.hoverIsLink -eq $c.Link)
        $tipOk = if ($null -eq $tipExpected) { -not $s2.view.tooltipVisible } else { $s2.view.tooltipVisible -and ($s2.view.tooltipText -ceq $tipExpected) }
        Check "$name highlights the right span" $okSpan @{ got = $hover; want = $c.Expect }
        Check "$name link=$($c.Link)" $okLink $s2.view.hoverIsLink
        Check "$name tooltip" $tipOk @{ got = $(if ($s2.view.tooltipVisible) { $s2.view.tooltipText } else { '(none)' }); want = $tipExpected }
    }
    $s
}

function S2 {
    Begin-Scenario 'S2' 'Hover rules and tooltips on the video formulas and the special cases (test hooks)'
    Reset-Env
    $fe = $script:fe
    $val = 'Val_FO'
    $ev = { param($sheet, $expr) Eval-Expected $sheet $expr }

    # K33 =1/(((($I27/Hundred)+1)^(K$31)))
    $k33 = '1/(((($I27/Hundred)+1)^(K$31)))'
    $cases = @(
        @{ At = '^'; Expect = '(($I27/Hundred)+1)^(K$31)'; Tip = (& $ev $val '(($I27/Hundred)+1)^(K$31)'); Link = $false }
        @{ At = '+1)'; Off = 2; Expect = '(($I27/Hundred)+1)'; Tip = (& $ev $val '(($I27/Hundred)+1)'); Link = $false }
        @{ At = '/'; Expect = $k33; Tip = (& $ev $val $k33); Link = $false }
        @{ At = 'Hundred'; Off = 2; Expect = 'Hundred'; Tip = '100'; Link = $true }
        @{ At = '$I27'; Off = 1; Expect = '$I27'; Tip = '6.2'; Link = $true }
        @{ At = '='; Expect = '=' + $k33; Tip = (& $ev $val 'K33'); Link = $false }
    )
    $null = Invoke-HoverCases $fe $val 'K33' $cases 'K33'
    Save-FeShot $fe 'S2-K33' | Out-Null

    # K32 =IF(K3="On",1/(((-($I26/Hundred)+1)^(K$31))),0)
    $cases = @(
        @{ At = '^'; Expect = '(-($I26/Hundred)+1)^(K$31)'; Tip = (& $ev $val '(-($I26/Hundred)+1)^(K$31)'); Link = $false }
        @{ At = 'IF'; Expect = 'IF(K3="On",1/(((-($I26/Hundred)+1)^(K$31))),0)'; Tip = (& $ev $val 'IF(K3="On",1/(((-($I26/Hundred)+1)^(K$31))),0)'); Link = $false }
        @{ At = '"On"'; Off = 1; Expect = '"On"'; Tip = '"On"'; Link = $false }
        @{ At = 'K3'; Expect = 'K3'; Tip = '"On"'; Link = $true }
    )
    $null = Invoke-HoverCases $fe $val 'K32' $cases 'K32'

    # K34 =IF(J3,(1/((365*(($I27/Hundred)+1)^(K$31)))),0): the ')' that closes '((365...'
    $inner = '((365*(($I27/Hundred)+1)^(K$31)))'
    $cases = @(
        @{ At = $inner; Off = $inner.Length - 1; Expect = $inner; Tip = (& $ev $val $inner); Link = $false }
        @{ At = 'J3'; Off = 0; Expect = 'J3'; Tip = 'TRUE'; Link = $true }
    )
    $null = Invoke-HoverCases $fe $val 'K34' $cases 'K34'

    # K35 long formula: the '*' highlights everything after '='
    $k35 = '(SUM(Fin_P_FO!L39:L40)+Fin_P_FO!L26+Notes_FO!L113+Notes_FO!L101)*(Macro_FA!L183/Hundred)'
    $cases = @(
        @{ At = ')*('; Off = 1; Expect = $k35; Tip = (& $ev $val $k35); Link = $false }
        @{ At = 'Fin_P_FO!L39:L40'; Off = 3; Expect = 'Fin_P_FO!L39:L40'; Tip = ('L39:L40 containing' + "`n" + '[2x1] 1, 2'); Link = $true }
    )
    $null = Invoke-HoverCases $fe $val 'K35' $cases 'K35'

    # K31 =J31+K19 (the drill target of the video) and J31 with OFFSET
    $cases = @(
        @{ At = '+'; Expect = 'J31+K19'; Tip = (& $ev $val 'J31+K19'); Link = $false }
        @{ At = 'J31'; Off = 1; Expect = 'J31'; Tip = (& $ev $val 'J31'); Link = $true }
    )
    $null = Invoke-HoverCases $fe $val 'K31' $cases 'K31'
    $off = 'OFFSET($I$19,0,0,1,Aqui_Per)'
    $cases = @(
        @{ At = 'OFFSET'; Off = 2; Expect = $off; Tip = ('I19:J19 containing' + "`n" + '[1x2] , 0.252054794520548'); Link = $true }
        @{ At = 'I$19:$I$19'; Off = 3; Expect = 'I$19:$I$19'; Tip = '[Blank]'; Link = $true }
        @{ At = 'Aqui_Per'; Off = 2; Expect = 'Aqui_Per'; Tip = '2'; Link = $true }
    )
    $null = Invoke-HoverCases $fe $val 'J31' $cases 'J31'

    # Stores!F3 and F4 (video)
    $f3 = 'Sheet8!E7+(E30/C4)^(D4/D5)+D14-E21*E16+Summary!C5'
    $cases = @(
        @{ At = '/'; Off = 0; Expect = 'E30/C4'; Tip = (& $ev 'Stores' 'E30/C4'); Link = $false }
        @{ At = '+Summary'; Off = 0; Expect = $f3; Tip = (& $ev 'Stores' $f3); Link = $false }
        @{ At = 'Sheet8!E7'; Off = 3; Expect = 'Sheet8!E7'; Tip = '3'; Link = $true }
    )
    $null = Invoke-HoverCases $fe 'Stores' 'F3' $cases 'Stores!F3'
    $cases = @(
        @{ At = 'Forcast'; Off = 1; Expect = "'Inventory Forcast'!C`$6"; Tip = '25'; Link = $true }
        @{ At = 'C3'; Off = 0; Expect = 'C3'; Tip = 'FALSE'; Link = $true }
    )
    $null = Invoke-HoverCases $fe 'Stores' 'F4' $cases 'Stores!F4'
    S2-Special
    Check-LogClean
    End-Scenario
}

# Special cases: blank, ranges "containing", errors, text, bool, defined name that is not a range, structured refs,
# 3D references, INDEX range, IF returning a reference.
function S2-Special {
    $fe = $script:fe
    $val = 'Val_FO'
    # [Blank]: a formula referencing an empty cell (Val_FO!K47 =Q99)
    $ws = Get-Sheet $val
    [void][System.__ComObject].InvokeMember('Formula', [Reflection.BindingFlags]::SetProperty, $null, $ws.Range('K47'), @('=Q99+Q98'))
    [void][System.__ComObject].InvokeMember('Formula', [Reflection.BindingFlags]::SetProperty, $null, $ws.Range('K48'), @('=SUM(I26:I28)'))
    $cases = @(
        @{ At = 'Q99'; Off = 0; Occ = 1; Expect = 'Q99'; Tip = '[Blank]'; Link = $true }
    )
    $null = Invoke-HoverCases $fe $val 'K47' $cases 'K47'
    $cases = @(
        @{ At = 'I26:I28'; Off = 2; Occ = 1; Expect = 'I26:I28'; Tip = ('I26:I28 containing' + "`n" + '[3x1] 9.4, 6.2, 50'); Link = $true }
        @{ At = 'SUM'; Off = 1; Occ = 1; Expect = 'SUM(I26:I28)'; Tip = '65.6'; Link = $false }
    )
    $null = Invoke-HoverCases $fe $val 'K48' $cases 'K48'
    # errors
    $cases = @( @{ At = '/'; Off = 0; Occ = 1; Expect = '1/0'; Tip = '#DIV/0!'; Link = $false } )
    $null = Invoke-HoverCases $fe $val 'K38' $cases 'K38'
    # text (quoted)
    $cases = @(
        @{ At = '"abc"'; Off = 2; Occ = 1; Expect = '"abc"'; Tip = '"abc"'; Link = $false }
        @{ At = '&'; Off = 0; Occ = 1; Expect = '"abc"&"def"'; Tip = '"abcdef"'; Link = $false }
    )
    $null = Invoke-HoverCases $fe $val 'K39' $cases 'K39'
    # bool
    $cases = @(
        @{ At = '>'; Off = 0; Occ = 1; Expect = 'I26>I27'; Tip = 'TRUE'; Link = $false }
        @{ At = 'I26'; Off = 1; Occ = 1; Expect = 'I26'; Tip = '9.4'; Link = $true }
    )
    $null = Invoke-HoverCases $fe $val 'K40' $cases 'K40'
    # defined name that is not a range: no color, no link, tooltip with the value; range on another sheet
    $s = Open-Explorer $fe $val 'K41'
    $iTax = Ix $s 'TaxRate'
    $ink = Get-SpanInk $s $iTax 7
    Check 'TaxRate (a name that is not a range) is black' ((Color-Dist $ink 0) -le 40) (Hex $ink)
    Check 'no overlay box for the Tables!B2:B4 range (other sheet) or TaxRate' ((Get-Overlay $s).Count -eq 0) (Get-Overlay $s).Count
    $cases = @(
        @{ At = 'TaxRate'; Off = 3; Occ = 1; Expect = 'TaxRate'; Tip = '0.3'; Link = $false }
        @{ At = 'Tables!B2:B4'; Off = 3; Occ = 1; Expect = 'Tables!B2:B4'; Tip = ('B2:B4 containing' + "`n" + '[3x1] 10, 20, 30'); Link = $true }
    )
    $null = Invoke-HoverCases $fe $val 'K41' $cases 'K41'
    # structured reference and table
    $cases = @(
        @{ At = 'Table1[Amount]'; Off = 3; Occ = 1; Expect = 'Table1[Amount]'; Tip = ('B2:B4 containing' + "`n" + '[3x1] 10, 20, 30'); Link = $true }
    )
    $null = Invoke-HoverCases $fe 'Tables' 'D2' $cases 'Tables!D2'
    # 3D
    $cases = @(
        @{ At = 'Sheet8:Summary!A1'; Off = 4; Occ = 1; Expect = 'Sheet8:Summary!A1'; Tip = '1'; Link = $true }
    )
    $null = Invoke-HoverCases $fe $val 'K37' $cases 'K37'
    # INDEX returning a range
    $cases = @(
        @{ At = 'INDEX'; Off = 1; Occ = 1; Expect = 'INDEX(J19:L19,1,2)'; Tip = '0.252054794520548'; Link = $true }
        @{ At = '):L19'; Off = 1; Occ = 1; Expect = 'INDEX(J19:L19,1,2):L19'; Tip = ('K19:L19 containing' + "`n" + '[1x2] 0.252054794520548, 0.25'); Link = $true }
    )
    $null = Invoke-HoverCases $fe $val 'K42' $cases 'K42'
    # IF returning a reference
    $cases = @(
        @{ At = 'IF'; Off = 0; Occ = 1; Expect = 'IF(TRUE,I26,I27)'; Tip = '9.4'; Link = $true }
    )
    $null = Invoke-HoverCases $fe $val 'K43' $cases 'K43'
}
