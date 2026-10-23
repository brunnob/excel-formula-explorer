# S15b: cell edit mode II. Everything the explorer does while Excel is editing a cell must be "nothing": no navigation, no commit of the
# edit, no dialogs, no late action after the edit ends. All input is real (mouse and keys); Excel's object model refuses calls from
# outside while it edits (Test-ExcelEditMode), so states are read before and after, and Win32 (windows, dialogs) in between.

# Puts Excel's keyboard focus on its grid (click the title bar of the workbook window) and presses Esc: leaves any edit mode.
function Leave-EditMode($fe) {
    $er = Get-Rect (Get-ExcelHwnd $fe)
    Click-At ($er[0] + 800) ($er[1] + 14)
    Start-Sleep -Milliseconds 500
    Press-Key $script:VK.Esc
    Start-Sleep -Milliseconds 900
    if (Test-ExcelEditMode $fe) { Press-Key $script:VK.Esc; Start-Sleep -Milliseconds 600 }
}

function Click-FormulaBar($fe) {
    $root = [System.Windows.Automation.AutomationElement]::FromHandle((Get-ExcelHwnd $fe))
    $cond = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::NameProperty, 'Formula Bar')
    $el = $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $cond)
    if (-not $el) { return $false }
    $r = $el.Current.BoundingRectangle
    Click-At ([int]($r.Left + 30)) ([int]($r.Top + $r.Height / 2))
    $true
}

function Count-EditModeLogLines {
    @(Log-Matches 'ComErrors: Excel is in cell edit mode').Count
}

function Assert-NoDialogs($fe, [string]$What) {
    $dialogs = @([FeWin2]::TopLevel([uint32]$fe.Pid) | Where-Object { $_.Split('|')[1] -eq '#32770' })
    Check "no dialog / exception box in the Excel process $What" ($dialogs.Count -eq 0) $dialogs
}

function S15b {
    Begin-Scenario 'S15b' 'Edit mode II: Back, formula-bar editing and Point mode ignored; Close/Esc on the explorer with a gold link into a hidden sheet; deferred restore after the edit'
    Reset-Env
    $fe = $script:fe
    $ws = $fe.Excel.ActiveSheet

    # ---------- (A) Back while a cell is being edited ----------
    $s = Open-Explorer $fe 'Val_FO' 'K33'
    [FeWin]::Foreground([IntPtr][int64]$s.explorerHwnd); Start-Sleep -Milliseconds 300
    Click-Node $fe $s 'K$31' 1 0
    $r = Wait-Timed $fe { $_.session.root.address -eq '$K$31' -and -not $_.view.isLoading } 3500
    Check 'setup: drilled to K31 (stack 1)' ($r.ok -and (Get-Stack $r.state).Count -eq 1) $r.state.session.root.address
    Start-Sleep -Milliseconds 400
    $s = Get-FeState $fe
    $origFormula = $ws.Range('K31').Formula
    $logBefore = Count-EditModeLogLines
    Click-Cell $fe 'K31'                      # root cell: not a new root
    Start-Sleep -Milliseconds 500
    Press-Key $script:VK.F2; Start-Sleep -Milliseconds 400
    Press-Key 0x30; Start-Sleep -Milliseconds 400
    Check 'Excel is editing K31' (Test-ExcelEditMode $fe) ''
    $b = $s.buttons.back
    Click-At ([int](($b[0] + $b[2]) / 2)) ([int](($b[1] + $b[3]) / 2))     # real click on Back
    Start-Sleep -Milliseconds 1200
    Check 'Back was ignored: Excel is still in edit mode (nothing ended or committed the edit)' (Test-ExcelEditMode $fe) ''
    Assert-NoDialogs $fe 'after Back during the edit'
    Leave-EditMode $fe
    Start-Sleep -Milliseconds 2500            # a click must not fire seconds later, when the edit ends
    $s = Get-FeState $fe
    Check 'after the edit: no late Back (root K31, stack 1, Excel on K31, Back enabled)' ($s.session.root.address -eq '$K$31' -and (Get-Stack $s).Count -eq 1 -and (Cell-Addr $fe) -eq 'K31' -and $s.view.backEnabled) @{ root = $s.session.root.address; stack = (Get-Stack $s).Count; cell = (Cell-Addr $fe) }
    Check 'the edit was cancelled, not committed (formula of K31 unchanged)' ($ws.Range('K31').Formula -ceq $origFormula) $ws.Range('K31').Formula
    Check 'the edit-mode detection was logged once for the whole edit' ((Count-EditModeLogLines) - $logBefore -eq 1) @{ lines = (Count-EditModeLogLines) - $logBefore }
    Check 'the ignored Back is in the log' ((Log-Matches "'back' ignored, Excel is in cell edit mode").Count -ge 1) ''
    $b = $s.buttons.back
    Click-At ([int](($b[0] + $b[2]) / 2)) ([int](($b[1] + $b[3]) / 2))
    $r = Wait-Timed $fe { $_.session.root.address -eq '$K$33' -and -not $_.view.isLoading } 3500
    Check 'Back works again once the edit is over (root K33)' $r.ok $r.state.session.root.address
    Close-Explorer $fe

    # ---------- (B) editing in the formula bar ----------
    $s = Open-Explorer $fe 'Val_FO' 'K33'
    $origFormula = $ws.Range('K33').Formula
    $logBefore = Count-EditModeLogLines
    Click-Cell $fe 'K33'; Start-Sleep -Milliseconds 400
    $ok = Click-FormulaBar $fe
    Start-Sleep -Milliseconds 600
    Check 'precondition: the formula bar is being edited' ($ok -and (Test-ExcelEditMode $fe)) $ok
    $text = $s.view.displayText
    $pHun = Char-Point $s (Find-FeIndex $text 'Hundred' 1 2)
    Move-Mouse $pHun.X $pHun.Y; Start-Sleep -Milliseconds 800
    Click-At $pHun.X $pHun.Y; Start-Sleep -Milliseconds 500
    Click-At $pHun.X $pHun.Y -Double; Start-Sleep -Milliseconds 800
    Assert-NoDialogs $fe 'while the formula bar is edited'
    Check 'link clicks were ignored: Excel is still editing the formula bar' (Test-ExcelEditMode $fe) ''
    Check 'the explorer window is still there' ([FeWin2]::IsWindow([IntPtr][int64]$s.explorerHwnd)) ''
    Leave-EditMode $fe
    Start-Sleep -Milliseconds 1500
    $s = Get-FeState $fe
    Check 'after the edit: still on Val_FO!K33, no gold, no push, formula unchanged' ($fe.Excel.ActiveSheet.Name -eq 'Val_FO' -and (Cell-Addr $fe) -eq 'K33' -and $null -eq $s.view.goldStart -and (Get-Stack $s).Count -eq 0 -and $ws.Range('K33').Formula -ceq $origFormula) @{ sheet = $fe.Excel.ActiveSheet.Name; cell = (Cell-Addr $fe); gold = (Get-GoldText $s); stack = (Get-Stack $s).Count }
    Check 'edit mode logged once for this edit (3 ignored clicks)' ((Count-EditModeLogLines) - $logBefore -eq 1) @{ lines = (Count-EditModeLogLines) - $logBefore }
    Save-FeShot $fe 'S15b-after-formula-bar' | Out-Null

    # ---------- (C) Point mode ----------
    $logBefore = Count-EditModeLogLines
    Click-Cell $fe 'K33'; Start-Sleep -Milliseconds 400
    Press-Key 0xBB; Start-Sleep -Milliseconds 300            # '='
    Press-Key $script:VK.Right; Start-Sleep -Milliseconds 400   # Point mode: the arrow picks a reference
    Check 'precondition: Excel is in Point mode (editing)' (Test-ExcelEditMode $fe) ''
    Click-At $pHun.X $pHun.Y; Start-Sleep -Milliseconds 600
    Click-At $pHun.X $pHun.Y -Double; Start-Sleep -Milliseconds 800
    Assert-NoDialogs $fe 'in Point mode'
    Check 'still in Point mode after the clicks' (Test-ExcelEditMode $fe) ''
    Leave-EditMode $fe
    Start-Sleep -Milliseconds 1500
    $s = Get-FeState $fe
    Check 'after Point mode: no navigation, K33 unchanged, explorer on K33' ((Cell-Addr $fe) -eq 'K33' -and $fe.Excel.ActiveSheet.Name -eq 'Val_FO' -and $ws.Range('K33').Formula -ceq $origFormula -and $s.session.root.address -eq '$K$33' -and $null -eq $s.view.goldStart) @{ cell = (Cell-Addr $fe); root = $s.session.root.address }
    Check 'edit mode is not logged per click in Point mode (at most once; the flag resets when a call sees the edit over)' ((Count-EditModeLogLines) - $logBefore -le 1) @{ lines = (Count-EditModeLogLines) - $logBefore }
    Close-Explorer $fe

    # ---------- (D) Close (button) and Esc (key) on the explorer while Excel edits the cell its gold link selected ----------
    foreach ($how in @('Close button', 'Esc key')) {
        $s = Open-Explorer $fe 'Val_FO' 'K36'
        [FeWin]::Foreground([IntPtr][int64]$s.explorerHwnd); Start-Sleep -Milliseconds 300
        Click-Node $fe $s 'Hidden_FO!B2' 1 2
        $r = Wait-Timed $fe { $null -ne $_.view.goldStart } 3000
        Start-Sleep -Milliseconds 500
        Check "($how) setup: Hidden_FO was unhidden and B2 selected (gold)" ($fe.Excel.ActiveWorkbook.Worksheets.Item('Hidden_FO').Visible -eq -1 -and $fe.Excel.ActiveSheet.Name -eq 'Hidden_FO' -and (Cell-Addr $fe) -eq 'B2') @{ sheet = $fe.Excel.ActiveSheet.Name; cell = (Cell-Addr $fe) }
        $er = Get-Rect (Get-ExcelHwnd $fe)
        Click-At ($er[0] + 800) ($er[1] + 14); Start-Sleep -Milliseconds 500     # Excel's window takes the keyboard; the selection stays
        Press-Key $script:VK.F2; Start-Sleep -Milliseconds 400
        Press-Key 0x39; Start-Sleep -Milliseconds 400                             # '9'
        Check "($how) Excel is editing Hidden_FO!B2" (Test-ExcelEditMode $fe) ''
        $s = Get-FeState2 $fe $s
        if ($how -eq 'Close button') {
            $c = $s.buttons.close
            Click-At ([int](($c[0] + $c[2]) / 2)) ([int](($c[1] + $c[3]) / 2))
        }
        else {
            $b = $s.view.screenBounds
            Click-At ($b.x + 80) ($b.y + 12); Start-Sleep -Milliseconds 400       # activate the explorer (title bar); Excel keeps editing
            Press-Key $script:VK.Esc
        }
        Start-Sleep -Milliseconds 1200
        Check "($how) the explorer window closed" (-not [FeWin2]::IsWindow([IntPtr][int64]$s.explorerHwnd)) ''
        Check "($how) Excel is still editing (the explorer closed without ending or committing the edit)" (Test-ExcelEditMode $fe) ''
        Assert-NoDialogs $fe "after closing the explorer ($how)"
        Leave-EditMode $fe
        Start-Sleep -Milliseconds 2500
        $hid = $fe.Excel.ActiveWorkbook.Worksheets.Item('Hidden_FO')
        Check "($how) the edit was cancelled (B2 still 11) and the destination sheet stays visible with the selection (original design)" ($hid.Range('B2').Value2 -eq 11 -and $hid.Visible -eq -1) @{ b2 = $hid.Range('B2').Value2; visible = $hid.Visible }
        Check "($how) the explorer is closed" (-not (Get-FeState $fe).isOpen) ''
        $hid.Visible = 0
        $fe.Excel.ActiveWorkbook.Worksheets.Item('Val_FO').Activate()
    }

    # ---------- (E) the deferred restore itself: a session that ends in edit mode with something left to hide ----------
    foreach ($case in @(@{ Sheet = 'Hidden_FO'; Address = '$B$2'; Original = 0 }, @{ Sheet = 'VeryHidden_FO'; Address = '$B$3'; Original = 2 })) {
        $name = $case.Sheet
        Select-FeCell $fe 'Val_FO' 'M5'
        $logMark = (Get-RealLogLines $script:OutDir).Count
        Check "($name) setup: the gateway unhid the sheet and the selection moved away" ((Invoke-Fe2 $fe 'FE_Test_DeferSetup' $name $case.Address) -eq 'ok' -and $fe.Excel.ActiveWorkbook.Worksheets.Item($name).Visible -eq -1 -and $fe.Excel.ActiveSheet.Name -eq 'Val_FO') ''
        Check "($name) the DeferEnd hook is armed" ((Invoke-Fe2 $fe 'FE_Test_DeferEnd' 3500) -eq 'ok') ''
        Click-Cell $fe 'M5'; Start-Sleep -Milliseconds 300       # real click on the selected cell, then F2: Excel edits before the session ends
        Press-Key $script:VK.F2; Start-Sleep -Milliseconds 300
        Press-Key 0x38; Start-Sleep -Milliseconds 300
        Check "($name) Excel is editing when the session ends" (Test-ExcelEditMode $fe) ''
        Start-Sleep -Milliseconds 5000                             # the session ended at 3.5 s: it could not restore
        Check "($name) while the edit lasts the sheet is still visible (restore deferred) and Excel is still editing" (Test-ExcelEditMode $fe) ''
        Assert-NoDialogs $fe "while the restore is deferred ($name)"
        Press-Key $script:VK.Esc                                   # the keyboard is still on Excel's cell
        Start-Sleep -Milliseconds 700
        if (Test-ExcelEditMode $fe) { Press-Key $script:VK.Esc; Start-Sleep -Milliseconds 500 }
        $restored = Wait-Until { $fe.Excel.ActiveWorkbook.Worksheets.Item($name).Visible -eq $case.Original } 8000
        Check "($name) once the edit ended the deferred restore hid the sheet again (to $($case.Original))" $restored $fe.Excel.ActiveWorkbook.Worksheets.Item($name).Visible
        $lines = Get-RealLogLines $script:OutDir
        $new = @($lines | Select-Object -Skip $logMark)
        Check "($name) the log tells the story: deferred, then finished" (@($new | Where-Object { $_ -match 'deferring' }).Count -ge 1 -and @($new | Where-Object { $_ -match 'deferred restore finished' }).Count -ge 1) (($new | Where-Object { $_ -match 'defer' } | Select-Object -First 3) -join ' // ')
    }
    Check-LogClean
    End-Scenario
}

# Application.Run with two arguments (Invoke-Fe passes at most one).
function Invoke-Fe2($fe, [string]$Hook, $Arg1, $Arg2) {
    $r = if ($PSBoundParameters.ContainsKey('Arg2')) { $fe.Excel.Run($Hook, $Arg1, $Arg2) } else { $fe.Excel.Run($Hook, $Arg1) }
    if ($r -is [string] -and $r -ne 'ok') { Write-Warning "$Hook -> $r" }
    $r
}

# The dump cannot be read while Excel edits: reuse the last one.
function Get-FeState2($fe, $Last) { $Last }
