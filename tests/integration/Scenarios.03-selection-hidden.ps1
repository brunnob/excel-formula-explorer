# S6 (manual selection), S7 (hidden sheets), S8 (protected structure), S9 (outline rows), S10 (non-formula cells).

function Cell-Addr($fe) { $fe.Excel.ActiveCell.Address($false, $false) }
function Sheet-Visible($fe, [string]$Name) { $fe.Excel.ActiveWorkbook.Worksheets.Item($Name).Visible }

# Real click on a sheet cell of the active sheet (center of the cell on the active pane).
function Click-Cell($fe, [string]$Address, [switch]$Double) {
    $c = Get-CellCenter $fe $fe.Excel.ActiveSheet.Range($Address)
    Click-At $c.X $c.Y -Double:$Double
}

function Wait-Root($fe, [string]$Address, [int]$TimeoutMs = 4000) {
    Wait-Timed $fe { $_.session.root.address -eq $Address -and -not $_.view.isLoading -and $_.view.displayText.Length -gt 0 } $TimeoutMs
}

function S6 {
    Begin-Scenario 'S6' 'Manual selection while open: root follows, pushes, no push for the root or for explorer-made selections'
    Reset-Env
    $fe = $script:fe
    $s = Open-Explorer $fe 'Val_FO' 'K33'
    # real click on a grid cell (K32)
    Click-Cell $fe 'K32'
    $r = Wait-Root $fe '$K$32'
    $s = $r.state
    Check 'real click on K32: root follows, text is the K32 formula' ($r.ok -and $s.view.displayText -like '=IF(K3="On"*') $s.view.displayText
    Check 'previous root (K33) pushed, Back enabled' ((Get-Stack $s).Count -eq 1 -and (Get-Stack $s)[0].address -eq '$K$33' -and $s.view.backEnabled) @{ stack = (Get-Stack $s); back = $s.view.backEnabled }
    Check 'gold cleared' ($null -eq $s.view.goldStart) (Get-GoldText $s)
    Check 'overlay redrawn for the new root (K3, I26, K31)' ((Get-Overlay $s).Count -eq 3) (Get-Overlay $s).Count
    Save-FeShot $fe 'S6-click-K32' | Out-Null

    # selection through the object model
    $fe.Excel.ActiveSheet.Range('K34').Select() | Out-Null
    $r = Wait-Root $fe '$K$34'
    $s = $r.state
    Check 'Range.Select K34: root follows' $r.ok $s.session.root.address
    Check 'stack is now [K33, K32]' ((Get-Stack $s).Count -eq 2 -and (Get-Stack $s)[1].address -eq '$K$32') (Get-Stack $s | ForEach-Object { $_.address })

    # selecting the root again must not push
    $fe.Excel.ActiveSheet.Range('K34').Select() | Out-Null
    Start-Sleep -Milliseconds 500
    $s = Get-FeState $fe
    Check 'selecting the root again via Range.Select does not push' ((Get-Stack $s).Count -eq 2) (Get-Stack $s).Count
    Click-Cell $fe 'K34'
    Start-Sleep -Milliseconds 600
    $s = Get-FeState $fe
    Check 'clicking the root cell with the mouse does not push either' ((Get-Stack $s).Count -eq 2 -and $s.session.root.address -eq '$K$34') (Get-Stack $s).Count

    # a multi-cell selection uses the active cell
    $fe.Excel.ActiveSheet.Range('K35:K36').Select() | Out-Null
    $r = Wait-Root $fe '$K$35'
    Check 'multi-cell selection K35:K36 uses the active cell K35 as root' $r.ok $r.state.session.root.address
    $depth = (Get-Stack $r.state).Count

    # selections made by the explorer do not push: click a link in the explorer
    $s = $r.state
    $hwnd = [int64]$s.explorerHwnd
    [FeWin]::Foreground([IntPtr]$hwnd); Start-Sleep -Milliseconds 250
    $iHun = Ix $s 'Hundred' 1 2
    $p = Char-Point $s $iHun
    Click-At $p.X $p.Y
    $r = Wait-Timed $fe { (Get-GoldText $_) -ceq 'Hundred' } 2500
    Start-Sleep -Milliseconds 500
    $s = Get-FeState $fe
    Check 'link click selected Names_BA!C5 (made by the explorer)' ($fe.Excel.ActiveSheet.Name -eq 'Names_BA' -and (Cell-Addr $fe) -eq 'C5') ($fe.Excel.ActiveSheet.Name + '!' + (Cell-Addr $fe))
    Check 'that selection did not change the root nor push' ($s.session.root.address -eq '$K$35' -and (Get-Stack $s).Count -eq $depth) @{ root = $s.session.root.address; depth = (Get-Stack $s).Count; expected = $depth }
    Check 'gold stays on Hundred' ((Get-GoldText $s) -ceq 'Hundred') (Get-GoldText $s)
    # selecting the root cell again (it is on another sheet now) clears the gold without pushing
    $fe.Excel.ActiveWorkbook.Worksheets.Item('Val_FO').Activate() | Out-Null
    $fe.Excel.ActiveSheet.Range('K35').Select() | Out-Null
    Start-Sleep -Milliseconds 700
    $s = Get-FeState $fe
    Check 'selecting the root again clears gold, no push' ($null -eq $s.view.goldStart -and (Get-Stack $s).Count -eq $depth -and $s.session.root.address -eq '$K$35') @{ gold = (Get-GoldText $s); depth = (Get-Stack $s).Count }

    # Back walks the manual selections back
    Invoke-Fe $fe 'FE_Test_Back' | Out-Null
    $r = Wait-Root $fe '$K$34'
    Check 'Back after manual selections returns to the previous root (K34) and selects it' ($r.ok -and (Cell-Addr $fe) -eq 'K34') @{ root = $r.state.session.root.address; sel = (Cell-Addr $fe) }
    Check-LogClean
    End-Scenario
}

function Get-DialogInfo($fe) {
    foreach ($line in [FeWin2]::TopLevel([uint32]$fe.Pid)) {
        $p = $line.Split('|')
        if ($p[1] -eq '#32770') {
            $texts = @()
            foreach ($c in [FeWin2]::ChildrenOf([IntPtr][int64]$p[0])) {
                $cp = $c.Split('|')
                if ($cp[1] -eq 'Static') { $t = [FeWin2]::TitleOf([IntPtr][int64]$cp[0]); if ($t) { $texts += $t } }
            }
            return @{ hwnd = [int64]$p[0]; title = $p[2]; text = ($texts -join ' ') }
        }
    }
    $null
}

function S7 {
    Begin-Scenario 'S7' 'Hidden and very hidden sheets: click unhides + navigates, other click/Back re-hides, Close on the destination keeps it'
    Reset-Env
    $fe = $script:fe
    $s = Open-Explorer $fe 'Val_FO' 'K36'
    $wbk = $fe.Excel.ActiveWorkbook
    Check 'precondition: Hidden_FO hidden (0), VeryHidden_FO very hidden (2)' ((Sheet-Visible $fe 'Hidden_FO') -eq 0 -and (Sheet-Visible $fe 'VeryHidden_FO') -eq 2) @{ h = (Sheet-Visible $fe 'Hidden_FO'); vh = (Sheet-Visible $fe 'VeryHidden_FO') }
    [FeWin]::Foreground([IntPtr][int64]$s.explorerHwnd); Start-Sleep -Milliseconds 250
    Click-Node $fe $s 'Hidden_FO!B2' 1 2
    $r = Wait-Timed $fe { (Get-GoldText $_) -ceq 'Hidden_FO!B2' } 3000
    $s = $r.state
    Check 'click Hidden_FO!B2: gold, sheet unhidden (-1) and active, B2 selected' ($r.ok -and (Sheet-Visible $fe 'Hidden_FO') -eq -1 -and $fe.Excel.ActiveSheet.Name -eq 'Hidden_FO' -and (Cell-Addr $fe) -eq 'B2') @{ gold = (Get-GoldText $s); vis = (Sheet-Visible $fe 'Hidden_FO'); sheet = $fe.Excel.ActiveSheet.Name; cell = (Cell-Addr $fe) }
    Check 'explorer stays foreground' (Get-Foreground-Timed $s 900) (Describe-Fg)
    Start-Sleep -Milliseconds 400

    Click-Node $fe $s 'VeryHidden_FO!B3' 1 2
    $r = Wait-Timed $fe { (Get-GoldText $_) -ceq 'VeryHidden_FO!B3' } 3000
    $s = $r.state
    Check 'click VeryHidden_FO!B3: unhidden (-1), active, B3 selected' ($r.ok -and (Sheet-Visible $fe 'VeryHidden_FO') -eq -1 -and $fe.Excel.ActiveSheet.Name -eq 'VeryHidden_FO' -and (Cell-Addr $fe) -eq 'B3') @{ vis = (Sheet-Visible $fe 'VeryHidden_FO'); sheet = $fe.Excel.ActiveSheet.Name }
    Check 'Hidden_FO was re-hidden when the other link was clicked' ((Sheet-Visible $fe 'Hidden_FO') -eq 0) (Sheet-Visible $fe 'Hidden_FO')
    Save-FeShot $fe 'S7-veryhidden' | Out-Null

    # Back (real button) -> root selected, very hidden re-hidden
    $b = $s.buttons.back
    Click-At ([int](($b[0] + $b[2]) / 2)) ([int](($b[1] + $b[3]) / 2))
    $r = Wait-Timed $fe { $null -eq $_.view.goldStart -and -not $_.session.backEnabled } 3000
    Start-Sleep -Milliseconds 300
    Check 'Back: gold gone, K36 selected on Val_FO' ($r.ok -and $fe.Excel.ActiveSheet.Name -eq 'Val_FO' -and (Cell-Addr $fe) -eq 'K36') @{ sheet = $fe.Excel.ActiveSheet.Name; cell = (Cell-Addr $fe) }
    Check 'Back re-hid both sheets (0 and 2)' ((Sheet-Visible $fe 'Hidden_FO') -eq 0 -and (Sheet-Visible $fe 'VeryHidden_FO') -eq 2) @{ h = (Sheet-Visible $fe 'Hidden_FO'); vh = (Sheet-Visible $fe 'VeryHidden_FO') }

    # Close while standing on an unhidden destination leaves it visible; the others are re-hidden
    $s = Get-FeState $fe
    Click-Node $fe $s 'Hidden_FO!B2' 1 2
    $r = Wait-Timed $fe { (Get-GoldText $_) -ceq 'Hidden_FO!B2' } 3000
    Start-Sleep -Milliseconds 400
    Click-Node $fe $r.state 'VeryHidden_FO!B3' 1 2
    $r = Wait-Timed $fe { (Get-GoldText $_) -ceq 'VeryHidden_FO!B3' } 3000
    Close-Explorer $fe
    Start-Sleep -Milliseconds 500
    Check 'Close on VeryHidden_FO!B3 leaves that sheet visible (-1) and still selected' ((Sheet-Visible $fe 'VeryHidden_FO') -eq -1 -and $fe.Excel.ActiveSheet.Name -eq 'VeryHidden_FO') @{ vh = (Sheet-Visible $fe 'VeryHidden_FO'); sheet = $fe.Excel.ActiveSheet.Name }
    Check 'Close re-hid Hidden_FO (0)' ((Sheet-Visible $fe 'Hidden_FO') -eq 0) (Sheet-Visible $fe 'Hidden_FO')
    Check-LogClean
    End-Scenario
}

function S8 {
    Begin-Scenario 'S8' 'Structure-protected workbook: link into a hidden sheet shows the English message and does not navigate'
    Reset-Env
    $fe = $script:fe
    $wbk = $fe.Excel.ActiveWorkbook
    [void]$wbk.Protect([Type]::Missing, $true, $false)
    Check 'precondition: workbook structure is protected' ($wbk.ProtectStructure -eq $true) $wbk.ProtectStructure
    $s = Open-Explorer $fe 'Val_FO' 'K36'
    [FeWin]::Foreground([IntPtr][int64]$s.explorerHwnd); Start-Sleep -Milliseconds 250
    Click-Node $fe $s 'Hidden_FO!B2' 1 2
    # the message box is modal: poll the windows instead of the hooks
    $dlg = $null
    Wait-Until { $script:dlg = Get-DialogInfo $fe; $null -ne $script:dlg } 4000 | Out-Null
    $dlg = $script:dlg
    Check 'a message box appears' ($null -ne $dlg) $dlg
    if ($dlg) {
        Check 'the message is the English text for a hidden sheet in a protected workbook' ($dlg.text -like "*Hidden_FO*hidden*structure is protected*") $dlg.text
        Save-FeShot $fe 'S8-message' | Out-Null
        Press-Key $script:VK.Enter
        $closed = Wait-Until { $null -eq (Get-DialogInfo $fe) } 2000
        if (-not $closed) { [void][FeWin2]::PostMessage([IntPtr]$dlg.hwnd, 0x10, [IntPtr]::Zero, [IntPtr]::Zero); Start-Sleep -Milliseconds 500 }
        Check 'the box can be dismissed' ($null -eq (Get-DialogInfo $fe)) ''
    }
    Start-Sleep -Milliseconds 400
    $s2 = Get-FeState $fe
    Check 'no navigation: still on Val_FO with K36 selected, no gold' ($fe.Excel.ActiveSheet.Name -eq 'Val_FO' -and (Cell-Addr $fe) -eq 'K36' -and $null -eq $s2.view.goldStart) @{ sheet = $fe.Excel.ActiveSheet.Name; cell = (Cell-Addr $fe); gold = (Get-GoldText $s2) }
    Check 'Hidden_FO still hidden' ((Sheet-Visible $fe 'Hidden_FO') -eq 0) (Sheet-Visible $fe 'Hidden_FO')
    Check 'Back stays disabled (nothing happened)' (-not $s2.view.backEnabled) $s2.view.backEnabled
    Hover-Char $fe (Ix $s2 'Hidden_FO!B2' 1 2)
    $r = Wait-Tooltip $fe 2500
    Check 'hover tooltip works after the message box (value of the hidden cell)' ($r.view.tooltipVisible -and $r.view.tooltipText -ceq '11') $r.view.tooltipText
    Check-LogClean
    End-Scenario
}

function S9 {
    Begin-Scenario 'S9' 'Outline: OFFSET target row 19 unhidden alone and re-hidden on Back; manually hidden row 22 only selected'
    Reset-Env
    $fe = $script:fe
    $ws = $fe.Excel.ActiveSheet
    Check 'precondition: rows 8:19 collapsed (hidden), row 22 hidden manually' ($ws.Range('A8').EntireRow.Hidden -eq $true -and $ws.Range('A19').EntireRow.Hidden -eq $true -and $ws.Range('A22').EntireRow.Hidden -eq $true) ''
    $s = Open-Explorer $fe 'Val_FO' 'J31'
    [FeWin]::Foreground([IntPtr][int64]$s.explorerHwnd); Start-Sleep -Milliseconds 250
    Click-Node $fe $s 'OFFSET' 1 2
    $r = Wait-Timed $fe { (Get-GoldText $_) -like 'OFFSET(*' } 3000
    Start-Sleep -Milliseconds 300
    $hidden = 8..19 | ForEach-Object { if ($ws.Range("A$_").EntireRow.Hidden) { $_ } }
    Check 'OFFSET target I19:J19 selected' ($fe.Excel.Selection.Address($false, $false) -eq 'I19:J19') $fe.Excel.Selection.Address($false, $false)
    Check 'only row 19 was unhidden; rows 8..18 stay hidden' (-not $ws.Range('A19').EntireRow.Hidden -and (@($hidden) -join ',') -eq '8,9,10,11,12,13,14,15,16,17,18') (@($hidden) -join ',')
    Check 'overlay follows the unhidden row (box for I$19:$I$19 / J$19 now visible)' ((Get-Overlay (Get-FeState $fe)).Count -ge 2) (Get-Overlay (Get-FeState $fe)).Count
    Save-FeShot $fe 'S9-row19' | Out-Null
    $b = $r.state.buttons.back
    Click-At ([int](($b[0] + $b[2]) / 2)) ([int](($b[1] + $b[3]) / 2))
    Start-Sleep -Milliseconds 800
    Check 'Back re-hides row 19' ($ws.Range('A19').EntireRow.Hidden -eq $true) $ws.Range('A19').EntireRow.Hidden
    $s = Get-FeState $fe
    Check 'overlay boxes for hidden rows disappear again' ((Get-Overlay $s).Count -le 2) (Get-Overlay $s).Count -Level soft
    Close-Explorer $fe

    # manually hidden row: only selected, never unhidden
    $s = Open-Explorer $fe 'Val_FO' 'K46'
    [FeWin]::Foreground([IntPtr][int64]$s.explorerHwnd); Start-Sleep -Milliseconds 250
    Check 'K46 text' ($s.view.displayText -ceq '=J22*2') $s.view.displayText
    Click-Node $fe $s 'J22'
    $r = Wait-Timed $fe { (Get-GoldText $_) -ceq 'J22' } 3000
    Start-Sleep -Milliseconds 300
    Check 'click J22: gold set, J22 selected' ($r.ok -and (Cell-Addr $fe) -eq 'J22') @{ gold = (Get-GoldText $r.state); cell = (Cell-Addr $fe) }
    Check 'row 22 (hidden manually) was NOT unhidden' ($ws.Range('A22').EntireRow.Hidden -eq $true) $ws.Range('A22').EntireRow.Hidden
    Check-LogClean
    End-Scenario
}

function S10 {
    Begin-Scenario 'S10' 'Non-formula cells (number, text, empty): shown in black, no hover/tooltip/clicks'
    Reset-Env
    $fe = $script:fe
    foreach ($case in @(@{ Cell = 'K44'; Text = '123.456' }, @{ Cell = 'K45'; Text = 'plain text' }, @{ Cell = 'Z60'; Text = '' })) {
        $s = Open-Explorer $fe 'Val_FO' $case.Cell
        $text = $s.view.displayText
        Check "$($case.Cell): displays the value '$($case.Text)'" ($text -ceq $case.Text) $text
        Check "$($case.Cell): view is not interactive" (-not $s.view.interactive) $s.view.interactive
        Check "$($case.Cell): no precedent boxes" ((Get-Overlay $s).Count -eq 0) (Get-Overlay $s).Count
        if ($text.Length -gt 0) {
            $ink = Get-SpanInk $s 0 $text.Length
            Check "$($case.Cell): text is black" ((Color-Dist $ink 0) -le 40) (Hex $ink)
            # real hover over several chars: nothing highlights, no tooltip, arrow cursor
            foreach ($i in 0, [int]($text.Length / 2), ($text.Length - 1)) {
                $p = Char-Point $s $i
                Move-Mouse $p.X $p.Y
            }
            Start-Sleep -Milliseconds 800
            $s2 = Get-FeState $fe
            Check "$($case.Cell): no highlight and no tooltip on hover" ($null -eq $s2.view.hoverStart -and -not $s2.view.tooltipVisible) @{ hover = (Get-HoverText $s2); tip = $s2.view.tooltipVisible }
            Check "$($case.Cell): arrow cursor (no hand)" ((Get-CursorKind) -eq 'arrow') (Get-CursorKind)
            $p = Char-Point $s 0
            $selBefore = $fe.Excel.Selection.Address($false, $false); $sheetBefore = $fe.Excel.ActiveSheet.Name
            Click-At $p.X $p.Y
            Click-At $p.X $p.Y -Double
            Start-Sleep -Milliseconds 500
            $s3 = Get-FeState $fe
            Check "$($case.Cell): clicks do nothing (no gold, selection and root unchanged, Back disabled)" ($null -eq $s3.view.goldStart -and $fe.Excel.Selection.Address($false, $false) -eq $selBefore -and $s3.session.root.address -eq $s.session.root.address -and -not $s3.view.backEnabled) @{ gold = (Get-GoldText $s3); sel = $fe.Excel.Selection.Address($false, $false); back = $s3.view.backEnabled }
        }
        else {
            Start-Sleep -Milliseconds 300
            $s2 = Get-FeState $fe
            Check "$($case.Cell): empty cell shows empty text, nothing to hover" ($null -eq $s2.view.hoverStart) ''
        }
        # real hover somewhere in the white box: still nothing
        $b = $s.view.screenBounds
        Move-Mouse ($b.x + 60) ($b.y + 80)
        Start-Sleep -Milliseconds 700
        $s4 = Get-FeState $fe
        Check "$($case.Cell): hover in the empty box shows nothing" ($null -eq $s4.view.hoverStart -and -not $s4.view.tooltipVisible) ''
        Save-FeShot $fe "S10-$($case.Cell)" | Out-Null
    }
    Check-LogClean
    End-Scenario
}
