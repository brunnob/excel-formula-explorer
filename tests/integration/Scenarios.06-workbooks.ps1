# S14: two workbooks (SDI Excel: one XLMAIN per workbook).

function Get-Owner([IntPtr]$h) { [FeWin2]::GetWindowLongPtr($h, -8).ToInt64() }
function Wb-Hwnd($wb) { [int64]$wb.Windows.Item(1).Hwnd }

function S14 {
    Begin-Scenario 'S14' 'Two workbooks: link into workbook B, explorer re-owned and on top, Back to A; closing B, closing the owner A'
    Reset-Env
    $fe = $script:fe
    $xl = $fe.Excel
    $wbA = $fe.Workbook
    $wbB = Open-Workbook (Join-Path $script:OutDir 'fe-test2.xlsx') 'wb-fe-test2.xlsx'
    $wbA.Activate()
    Start-Sleep -Milliseconds 500
    $nameB = $wbB.Name
    $ws = $wbA.Worksheets.Item('Val_FO')
    [void][System.__ComObject].InvokeMember('Formula', [Reflection.BindingFlags]::SetProperty, $null, $ws.Range('K47'), @("='[$nameB]Calc'!`$C`$5*2+K31"))
    Check 'precondition: the cross-workbook formula was accepted' ($ws.Range('K47').Formula -like "*$nameB*Calc*") $ws.Range('K47').Formula
    $hA = Wb-Hwnd $wbA; $hB = Wb-Hwnd $wbB
    Check 'precondition: the two workbooks have different top-level windows (SDI)' ($hA -ne $hB) @{ A = $hA; B = $hB }

    $s = Open-Explorer $fe 'Val_FO' 'K47'
    $eh = [IntPtr][int64]$s.explorerHwnd
    Check 'explorer opened on A!Val_FO!K47' ($s.session.root.workbookName -eq $wbA.Name -and $s.view.displayText -like "=*Calc*") $s.view.displayText
    Check 'explorer is owned by A''s window' ((Get-Owner $eh) -eq $hA) @{ owner = (Get-Owner $eh); A = $hA }
    $text = $s.view.displayText
    # tooltip for the cross-workbook reference
    $iRef = Find-FeIndex $text 'Calc' 1 0
    Hover-Char $fe $iRef
    $r = Wait-Tooltip $fe 2500
    Check 'tooltip of the reference into workbook B shows its value (42)' ($r.view.tooltipVisible -and $r.view.tooltipText -ceq '42') $r.view.tooltipText

    # click the link with the real mouse -> navigates into B
    [FeWin]::Foreground($eh); Start-Sleep -Milliseconds 300
    $p = Char-Point $s $iRef
    Click-At $p.X $p.Y
    $r = Wait-Timed $fe { $null -ne $_.view.goldStart } 3000
    Start-Sleep -Milliseconds 600
    $s = Get-FeState $fe
    Check 'click navigated into workbook B (active workbook B, C5 selected)' ($xl.ActiveWorkbook.Name -eq $nameB -and $xl.ActiveCell.Address($false, $false) -eq 'C5') @{ wb = $xl.ActiveWorkbook.Name; cell = $xl.ActiveCell.Address($false, $false) }
    Check 'gold is set on the reference' ($null -ne $s.view.goldStart) (Get-GoldText $s)
    Check 'explorer was re-owned to B''s window' ((Get-Owner $eh) -eq $hB) @{ owner = (Get-Owner $eh); B = $hB }
    $bnd = $s.view.screenBounds
    $top = [FeWin2]::TopWindowAt($bnd.x + 60, $bnd.y + 60)
    Check 'explorer window is on top of B (hit test at its title)' ($top.ToInt64() -eq [int64]$s.explorerHwnd) ('0x{0:X}' -f $top.ToInt64())
    Check 'explorer is the foreground window' (Get-Foreground-Timed $s 800) (Describe-Fg)
    Check 'overlay is hidden while B is shown (root is in A)' (-not $s.overlayVisible) $s.overlayVisible
    Save-FeShot $fe 'S14-in-B' | Out-Null

    # Back -> returns to A
    $b = $s.buttons.back
    Click-At ([int](($b[0] + $b[2]) / 2)) ([int](($b[1] + $b[3]) / 2))
    $r = Wait-Timed $fe { $null -eq $_.view.goldStart } 3000
    Start-Sleep -Milliseconds 700
    $s = Get-FeState $fe
    Check 'Back returned to workbook A, K47 selected' ($xl.ActiveWorkbook.Name -eq $wbA.Name -and $xl.ActiveCell.Address($false, $false) -eq 'K47') @{ wb = $xl.ActiveWorkbook.Name; cell = $xl.ActiveCell.Address($false, $false) }
    Check 'explorer re-owned to A again, on top' ((Get-Owner $eh) -eq $hA -and [FeWin2]::TopWindowAt($bnd.x + 60, $bnd.y + 60).ToInt64() -eq [int64]$s.explorerHwnd) @{ owner = (Get-Owner $eh); A = $hA }
    Check 'overlay back for A (root K47 has the K31 precedent)' ($s.overlayVisible -and (Get-Overlay $s).Count -ge 1) @{ visible = $s.overlayVisible; rects = (Get-Overlay $s).Count }

    # ---- closing B (not the root's workbook) while the explorer stands on it ----
    Click-Node $fe $s 'Calc' 1 0
    $r = Wait-Timed $fe { $null -ne $_.view.goldStart } 3000
    Start-Sleep -Milliseconds 600
    Check 'again inside B (explorer owned by B)' ($xl.ActiveWorkbook.Name -eq $nameB -and (Get-Owner $eh) -eq $hB) @{ wb = $xl.ActiveWorkbook.Name; owner = (Get-Owner $eh) }
    $wbB.Close($false)
    Start-Sleep -Milliseconds 1500
    $alive = $false
    try { $s = Get-FeState $fe; $alive = $s.isOpen } catch { }
    Check 'closing B while the explorer is on it: no crash; explorer still open (re-owned to A) and Excel responsive' ($alive -and $xl.ActiveWorkbook.Name -eq $wbA.Name) @{ open = $alive; active = $xl.ActiveWorkbook.Name }
    if ($alive) {
        $eh2 = [IntPtr][int64]$s.explorerHwnd
        Check 'the explorer window is alive and owned by A' ([FeWin2]::IsWindow($eh2) -and (Get-Owner $eh2) -eq $hA) @{ owner = (Get-Owner $eh2); A = $hA }
        Check 'no gold left pointing into the closed workbook' ($null -eq $s.view.goldStart) (Get-GoldText $s) -Level soft
        Check 'Excel selection is back in A' ($xl.ActiveWorkbook.Name -eq $wbA.Name) $xl.ActiveWorkbook.Name
    }
    Save-FeShot $fe 'S14-after-closing-B' | Out-Null

    # ---- closing the owner workbook A closes the explorer cleanly; the entry point can reopen it on B ----
    $wbB = Open-Workbook (Join-Path $script:OutDir 'fe-test2.xlsx') 'wb-fe-test2.xlsx'
    $wbA.Activate()
    Start-Sleep -Milliseconds 500
    $wbA.Close($false)
    Start-Sleep -Milliseconds 1500
    $s = $null
    try { $s = Get-FeState $fe } catch { Check 'Excel still answers after closing the owner workbook' $false $_.Exception.Message }
    if ($s) {
        Check 'closing the owner workbook closes the explorer (isOpen false)' (-not $s.isOpen) $s.isOpen
        Check 'no explorer window left for the process' ([FeWin]::FindTopLevel([uint32]$fe.Pid, 'Formula Explorer') -eq [IntPtr]::Zero) ''
        $wbB.Activate(); Start-Sleep -Milliseconds 300
        $xl.ActiveSheet.Range('C5').Select() | Out-Null
        $res = Invoke-Fe $fe 'FE_Test_Open'
        $s = Wait-FeState $fe { $_.isOpen -and $_.view -and $_.session.root -and -not $_.view.isLoading } 5000
        Check 'the explorer can be reopened (entry point) on workbook B' ($s.isOpen -and $s.session.root.workbookName -eq $wbB.Name) @{ res = $res; root = $s.session.root }
        Check 'the reopened explorer is owned by B''s window' ((Get-Owner ([IntPtr][int64]$s.explorerHwnd)) -eq (Wb-Hwnd $wbB)) ''
    }
    Check-LogClean
    End-Scenario
}
