# S14b/c/d: every way a workbook can go away while the explorer is open, driven through the real UI (WM_CLOSE = the X button,
# Excel's own "Do you want to save the changes?" prompt, real mouse clicks on its buttons).
#   S14b  a DIRTY non-root workbook B closed while the explorer stands on it: prompt up for > 2 s then Don't Save; a cancelled close
#   S14c  the ROOT workbook A closed while the explorer is owned by B (cancelled, then confirmed; and a clean close)
#   S14d  one of two windows of the same workbook closed; an unrelated third workbook closed

# Two workbooks, the explorer open on A!K47 whose formula points into B (SDI Excel: one top-level window per workbook).
function Start-CrossBook {
    Reset-Env
    $fe = $script:fe
    $wbA = $fe.Workbook
    $wbB = Open-Workbook (Join-Path $script:OutDir 'fe-test2.xlsx') 'wb-fe-test2.xlsx'
    $wbA.Activate()
    Start-Sleep -Milliseconds 500
    $nameB = $wbB.Name
    Set-Formula $wbA.Worksheets.Item('Val_FO') 'K47' "='[$nameB]Calc'!`$C`$5*2+K31"
    $s = Open-Explorer $fe 'Val_FO' 'K47'
    @{ fe = $fe; wbA = $wbA; wbB = $wbB; nameA = $wbA.Name; nameB = $nameB; hA = (Wb-Hwnd $wbA); hB = (Wb-Hwnd $wbB); eh = [int64]$s.explorerHwnd; state = $s }
}

# Clicks the link into workbook B with the real mouse: B becomes the active workbook and the explorer is re-owned to its window.
function Enter-WorkbookB($c) {
    $fe = $c.fe
    $s = Get-FeState $fe
    [FeWin]::Foreground([IntPtr]$c.eh); Start-Sleep -Milliseconds 300
    Click-Node $fe $s 'Calc' 1 0
    [void](Wait-Timed $fe { $null -ne $_.view.goldStart } 3000)
    Start-Sleep -Milliseconds 700
}

function Explorer-OnTop([int64]$Eh) {
    $r = Get-WinRect $Eh
    [FeWin2]::TopWindowAt($r[0] + 60, $r[1] + 14).ToInt64() -eq $Eh
}

function Log-Matches([string]$Pattern) {
    $lines = Get-RealLogLines $script:OutDir
    @($lines | Select-Object -Skip $script:LogBase | Where-Object { $_ -match $Pattern })
}

function S14b {
    Begin-Scenario 'S14b' 'Dirty non-root workbook B closed through the UI while the explorer stands on it: cancelled close, then save prompt up > 2 s and Don''t Save'
    $c = Start-CrossBook
    $fe = $c.fe; $xl = $fe.Excel; $eh = $c.eh
    try {
        $xl.DisplayAlerts = $true
        Enter-WorkbookB $c
        Check 'setup: the explorer is on top of B and owned by B''s window' ((Get-OwnerOf $eh) -eq $c.hB -and $xl.ActiveWorkbook.Name -eq $c.nameB) @{ owner = (Get-OwnerOf $eh); B = $c.hB }
        Make-Dirty $c.wbB 'Calc' 'C9' 1
        Check 'setup: B has unsaved changes' (-not $c.wbB.Saved) $c.wbB.Saved

        # ---- (1) cancelled close ----
        Post-WindowClose $c.hB
        $prompt = Wait-ExcelPrompt $fe
        Check 'the save prompt appears' ($null -ne $prompt) $prompt
        Start-Sleep -Milliseconds 2000
        $prompt = Get-ExcelPrompt $fe
        Check 'the prompt is still up after 2 s (longer than the old 600 ms timer)' ($null -ne $prompt) ''
        Check 'during the prompt the explorer is alive and owned by the ROOT workbook window, not by the closing B' ([FeWin2]::IsWindow([IntPtr]$eh) -and (Get-OwnerOf $eh) -eq $c.hA) @{ owner = (Get-OwnerOf $eh); A = $c.hA; B = $c.hB }
        Save-FeShot $fe 'S14b-prompt-cancel' | Out-Null
        if ($prompt) { [void](Press-PromptButton $prompt 'Cancel') }
        $gone = Wait-Until { $null -eq (Get-ExcelPrompt $fe) } 4000
        Check 'the prompt was dismissed (Cancel)' $gone ''
        # the cancelled close is recognised by polling (workbook still open, Excel answering): the explorer goes back on B
        $backOnB = Wait-Until { (Get-OwnerOf $eh) -eq $c.hB } 4000
        Check 'after the cancelled close the explorer follows the active window again (owned by B)' $backOnB @{ owner = (Get-OwnerOf $eh); B = $c.hB }
        $s = Get-FeState $fe
        Check 'B is still open and dirty, the explorer is still open and on top' ((Test-WorkbookOpen $fe $c.nameB) -and $s.isOpen -and (Explorer-OnTop $eh)) @{ open = $s.isOpen }
        Check 'the close ended as Cancelled in the log' ((Log-Matches "close of '$($c.nameB)' ended: Cancelled").Count -ge 1) ''

        # ---- (2) Don't Save after a long prompt ----
        Post-WindowClose $c.hB
        $prompt = Wait-ExcelPrompt $fe
        Check 'the save prompt appears again' ($null -ne $prompt) ''
        Start-Sleep -Milliseconds 2500
        $prompt = Get-ExcelPrompt $fe
        $own = Get-OwnerOf $eh
        Check 'the prompt has been up for > 2.5 s and the explorer window is still alive, owned by A' ($null -ne $prompt -and [FeWin2]::IsWindow([IntPtr]$eh) -and $own -eq $c.hA) @{ owner = $own; A = $c.hA; B = $c.hB }
        if ($prompt) { [void](Press-PromptButton $prompt 'DontSave') }
        $closed = Wait-Until { $null -eq (Get-ExcelPrompt $fe) -and (Test-WorkbookOpen $fe $c.nameB) -eq $false } 6000
        Check 'B closed (Don''t Save)' $closed ''
        Start-Sleep -Milliseconds 1200
        $s = Get-FeState $fe
        Check 'the explorer survived: same window, still open, root unchanged' ($s.isOpen -and [FeWin2]::IsWindow([IntPtr]$eh) -and [int64]$s.explorerHwnd -eq $eh -and $s.session.root.workbookName -eq $c.nameA) @{ open = $s.isOpen; alive = [FeWin2]::IsWindow([IntPtr]$eh); root = $s.session.root.address }
        Check 'it is owned by the active window (A) and on top of it' ((Get-OwnerOf $eh) -eq $c.hA -and $xl.ActiveWorkbook.Name -eq $c.nameA -and (Explorer-OnTop $eh)) @{ owner = (Get-OwnerOf $eh); active = $xl.ActiveWorkbook.Name }
        Check 'the close ended as Closed in the log' ((Log-Matches "close of '$($c.nameB)' ended: Closed").Count -ge 1) ''
        Save-FeShot $fe 'S14b-after-dont-save' | Out-Null
    }
    finally { try { $xl.DisplayAlerts = $false } catch { } }
    Check-LogClean
    End-Scenario
}

function S14c {
    Begin-Scenario 'S14c' 'Closing the ROOT workbook A while the explorer is owned by B: a cancelled close keeps it, a confirmed one (dirty + Don''t Save, or clean) closes it explicitly'
    $c = Start-CrossBook
    $fe = $c.fe; $xl = $fe.Excel; $eh = $c.eh
    try {
        $xl.DisplayAlerts = $true
        Enter-WorkbookB $c
        Check 'setup: the explorer is owned by B and the root is in A' ((Get-OwnerOf $eh) -eq $c.hB -and (Get-FeState $fe).session.root.workbookName -eq $c.nameA) @{ owner = (Get-OwnerOf $eh) }
        Make-Dirty $c.wbA 'Val_FO' 'Z1' 1

        # ---- cancelled ----
        Post-WindowClose $c.hA
        $prompt = Wait-ExcelPrompt $fe
        Check 'the save prompt for A appears' ($null -ne $prompt) ''
        Start-Sleep -Milliseconds 2000
        $prompt = Get-ExcelPrompt $fe
        Check 'during the prompt the explorer is alive and still owned by B (it is not owned by a window of the closing workbook)' ([FeWin2]::IsWindow([IntPtr]$eh) -and (Get-OwnerOf $eh) -eq $c.hB) @{ owner = (Get-OwnerOf $eh); A = $c.hA; B = $c.hB }
        if ($prompt) { [void](Press-PromptButton $prompt 'Cancel') }
        [void](Wait-Until { $null -eq (Get-ExcelPrompt $fe) } 4000)
        Start-Sleep -Milliseconds 1500
        $s = Get-FeState $fe
        Check 'cancelled close of the root workbook: explorer still open, root still A!K47' ($s.isOpen -and [FeWin2]::IsWindow([IntPtr]$eh) -and $s.session.root.address -eq '$K$47' -and $s.session.root.workbookName -eq $c.nameA) @{ open = $s.isOpen; root = $s.session.root.address }
        Check 'it follows the active window again (owner = the active workbook window)' ((Get-OwnerOf $eh) -eq [int64]$xl.ActiveWindow.Hwnd) @{ owner = (Get-OwnerOf $eh); active = [int64]$xl.ActiveWindow.Hwnd }

        # ---- confirmed (Don't Save) ----
        Post-WindowClose $c.hA
        $prompt = Wait-ExcelPrompt $fe
        Start-Sleep -Milliseconds 1500
        $prompt = Get-ExcelPrompt $fe
        if ($prompt) { [void](Press-PromptButton $prompt 'DontSave') }
        [void](Wait-Until { $null -eq (Get-ExcelPrompt $fe) -and (Test-WorkbookOpen $fe $c.nameA) -eq $false } 6000)
        Start-Sleep -Milliseconds 1500
        $s = Get-FeState $fe
        Check 'A closed (Don''t Save); the explorer closed explicitly with it' (-not $s.isOpen -and -not [FeWin2]::IsWindow([IntPtr]$eh)) @{ open = $s.isOpen; alive = [FeWin2]::IsWindow([IntPtr]$eh) }
        Check 'no "Formula Explorer" window left in the process' ([FeWin]::FindTopLevel([uint32]$fe.Pid, 'Formula Explorer') -eq [IntPtr]::Zero) ''
        Check 'B is still open and Excel answers' ((Test-WorkbookOpen $fe $c.nameB) -eq $true) ''
        Save-FeShot $fe 'S14c-after-root-closed' | Out-Null
        # the entry point still works afterwards
        $xl.ActiveWorkbook.Worksheets.Item('Calc').Range('C6').Select() | Out-Null
        [void](Invoke-Fe $fe 'FE_Test_Open')
        $s = Wait-FeState $fe { $_.isOpen -and $_.view -and $_.session.root -and -not $_.view.isLoading } 5000
        Check 'the explorer can be opened again on B' ($s.isOpen -and $s.session.root.workbookName -eq $c.nameB) $s.session.root.workbookName
    }
    finally { try { $xl.DisplayAlerts = $false } catch { } }

    # ---- clean root workbook closed through the object model while the explorer is owned by B ----
    $c = Start-CrossBook
    $fe = $c.fe; $xl = $fe.Excel; $eh = $c.eh
    Enter-WorkbookB $c
    Check 'clean variant, setup: owner is B' ((Get-OwnerOf $eh) -eq $c.hB) ''
    $fe.Workbook = $null
    $c.wbA.Close($false)
    Start-Sleep -Milliseconds 1800
    $s = Get-FeState $fe
    Check 'clean close of the root workbook A while the explorer is owned by B closes the explorer' (-not $s.isOpen -and -not [FeWin2]::IsWindow([IntPtr]$eh)) @{ open = $s.isOpen; alive = [FeWin2]::IsWindow([IntPtr]$eh) }
    Check-LogClean
    End-Scenario
}

function S14d {
    Begin-Scenario 'S14d' 'One of two windows of the same workbook closed, and an unrelated third workbook closed, while the explorer is open'
    Reset-Env
    $fe = $script:fe; $xl = $fe.Excel
    $wbA = $fe.Workbook
    $s = Open-Explorer $fe 'Val_FO' 'K33'
    $eh = [int64]$s.explorerHwnd
    $hA = Wb-Hwnd $wbA

    # ---- two windows of A ----
    $w2 = $wbA.NewWindow()
    Start-Sleep -Milliseconds 1000
    $h2 = [int64]$w2.Hwnd
    Check 'setup: the new window has its own top-level window and the explorer is owned by the active one' ($h2 -ne $hA -and (Get-OwnerOf $eh) -eq $h2) @{ A1 = $hA; A2 = $h2; owner = (Get-OwnerOf $eh) }
    $w2.Close()    # Window.Close: only this window goes, the workbook stays open
    Start-Sleep -Milliseconds 1500
    $s = Get-FeState $fe
    Check 'closing the window that owns the explorer (the workbook stays open): the explorer survives' ($s.isOpen -and [FeWin2]::IsWindow([IntPtr]$eh)) @{ open = $s.isOpen; alive = [FeWin2]::IsWindow([IntPtr]$eh) }
    if ($s.isOpen) {
        Check 'it is owned by the remaining window of A and on top of it' ((Get-OwnerOf $eh) -eq $hA -and (Explorer-OnTop $eh)) @{ owner = (Get-OwnerOf $eh); A1 = $hA }
        Check 'the root is unchanged (K33)' ($s.session.root.address -eq '$K$33') $s.session.root.address
    }
    Save-FeShot $fe 'S14d-after-window-close' | Out-Null
    # and with the keyboard
    if (-not $s.isOpen) { $s = Open-Explorer $fe 'Val_FO' 'K33'; $eh = [int64]$s.explorerHwnd }
    $w3 = $wbA.NewWindow()
    Start-Sleep -Milliseconds 1000
    [FeWin]::Foreground([IntPtr]$w3.Hwnd); Start-Sleep -Milliseconds 500
    Press-Key 0x57 -Modifiers 0x11   # Ctrl+W: Close Window (real keys); the workbook has another window, so only this one closes
    Start-Sleep -Milliseconds 1500
    $s = Get-FeState $fe
    Check 'Ctrl+W on the owner window: the explorer survives as well' ($s.isOpen -and [FeWin2]::IsWindow([IntPtr]$s.explorerHwnd)) @{ open = $s.isOpen }

    # ---- an unrelated third workbook ----
    if (-not $s.isOpen) { $s = Open-Explorer $fe 'Val_FO' 'K33' }
    $eh = [int64]$s.explorerHwnd
    $wbC = Open-Workbook (Join-Path $script:OutDir 'fe-test2.xlsx') 'wb-fe-test3.xlsx'
    $hC = Wb-Hwnd $wbC
    $wbA.Activate()
    Start-Sleep -Milliseconds 800
    Check 'setup: the explorer is owned by A, workbook C is unrelated' ((Get-OwnerOf $eh) -eq $hA -and $hC -ne $hA) @{ owner = (Get-OwnerOf $eh); A = $hA; C = $hC }
    $wbC.Close($false)
    Start-Sleep -Milliseconds 1500
    $s = Get-FeState $fe
    Check 'closing the unrelated workbook C changes nothing: explorer open, owned by A, same root' ($s.isOpen -and (Get-OwnerOf $eh) -eq $hA -and $s.session.root.address -eq '$K$33' -and $xl.ActiveWorkbook.Name -eq $wbA.Name) @{ open = $s.isOpen; owner = (Get-OwnerOf $eh) }
    $p = Char-Point $s (Ix $s '^')
    Move-Mouse $p.X $p.Y
    $r = Wait-Tooltip $fe 2500
    Check 'tooltip works after the third workbook closed' $r.view.tooltipVisible $r.view.tooltipText
    Check-LogClean
    End-Scenario
}
