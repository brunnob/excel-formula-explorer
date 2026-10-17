# S14e  the probe of a pending workbook close (ExcelGateway.ProbeWorkbook) classified in real Excel: Open / Gone, also while the save prompt is up
# S14f  the explorer is never handed to a hidden or minimized window when its owner window is deactivated or closed

# Arms the add-in's probe hook for $Name: DelayMs later (from the add-in's own dispatcher) it writes what ExcelGateway.ProbeWorkbook
# answers to a file. With a delay the answer can be taken while the save prompt is up, when Application.Run from outside is refused.
function Arm-WorkbookProbe($fe, [string]$Name, [string]$Tag, [int]$DelayMs) {
    $file = Join-Path $script:OutDir "probe-$Tag.json"
    if (Test-Path $file) { Remove-Item $file -Force }
    $r = $fe.Excel.Run('FE_Test_ProbeWorkbook', $Name, $file, $DelayMs)
    if ($r -ne 'ok') { Write-Warning "FE_Test_ProbeWorkbook -> $r" }
    $file
}
function Read-WorkbookProbe([string]$File, [int]$TimeoutMs = 8000) {
    [void](Wait-Until { Test-Path $File } $TimeoutMs 100)
    if (Test-Path $File) { Get-Content $File -Raw | ConvertFrom-Json } else { $null }
}
function Probe-Workbook($fe, [string]$Name, [string]$Tag) { Read-WorkbookProbe (Arm-WorkbookProbe $fe $Name $Tag 0) 3000 }

function S14e {
    Begin-Scenario 'S14e' 'Workbook probe in real Excel: Open for an open workbook (any case, also while its save prompt is up), Gone only for a name that is not in Application.Workbooks, a renamed workbook'
    $c = Start-CrossBook
    $fe = $c.fe; $xl = $fe.Excel
    try {
        $xl.DisplayAlerts = $true
        $p = Probe-Workbook $fe $c.nameB 'open'
        Check 'an open workbook is Open (and its window is found)' ($p -and $p.presence -eq 'Open' -and $p.windows -ge 1) $p
        $p = Probe-Workbook $fe $c.nameB.ToUpperInvariant() 'open-upper'
        Check 'Excel looks names up case-insensitively: the upper-case name is Open too' ($p -and $p.presence -eq 'Open') $p
        $p = Probe-Workbook $fe 'no-such-workbook.xlsx' 'gone'
        Check 'a name that is not in Application.Workbooks is Gone (confirmed by the walk)' ($p -and $p.presence -eq 'Gone' -and $p.windows -eq 0) $p
        $p = Probe-Workbook $fe ($c.nameB + 'x') 'gone-suffix'
        Check 'a name that only differs by a suffix is Gone, not mistaken for B' ($p -and $p.presence -eq 'Gone') $p

        # ---- while the save prompt is up: the add-in's own calls (made from Excel's main thread by the dispatcher, after the prompt came up) are
        # still answered, so B must be Open (its close is pending, it is not gone). Busy (calls rejected) is covered by the unit tests; the
        # controller also reports Busy while the workbook's windows are disabled by the prompt. ----
        Make-Dirty $c.wbB 'Calc' 'C9' 1
        $fileBusyB = Arm-WorkbookProbe $fe $c.nameB 'busy-B' 4000
        $fileBusyA = Arm-WorkbookProbe $fe $c.nameA 'busy-A' 4200
        $fileBusyX = Arm-WorkbookProbe $fe 'no-such-workbook.xlsx' 'busy-unknown' 4400
        Post-WindowClose $c.hB
        $prompt = Wait-ExcelPrompt $fe
        Check 'the save prompt for B appears' ($null -ne $prompt) ''
        $pB = Read-WorkbookProbe $fileBusyB 9000
        $pA = Read-WorkbookProbe $fileBusyA 3000
        $pX = Read-WorkbookProbe $fileBusyX 3000
        $stillUp = $null -ne (Get-ExcelPrompt $fe)
        Check 'the prompt was still up when the probes ran' $stillUp ''
        Check 'B (close pending, save prompt up) is not Gone: it is still listed, so Open (or Busy if Excel rejected the call)' ($pB -and $pB.presence -in @('Open', 'Busy')) $pB
        Check 'the root workbook A asked during the prompt is not Gone either' ($pA -and $pA.presence -in @('Open', 'Busy')) $pA
        Check 'a name that is not in the list is still Gone during the prompt (the walk completes) or Busy, never Open' ($pX -and $pX.presence -in @('Gone', 'Busy')) $pX
        Save-FeShot $fe 'S14e-prompt-up' | Out-Null
        if ($prompt) { [void](Press-PromptButton $prompt 'Cancel') }
        [void](Wait-Until { $null -eq (Get-ExcelPrompt $fe) } 4000)
        Start-Sleep -Milliseconds 600
        $p = Probe-Workbook $fe $c.nameB 'after-cancel'
        Check 'after Cancel B is Open again' ($p -and $p.presence -eq 'Open') $p

        # ---- Save As renames the workbook: the old name is Gone, the new one Open ----
        $renamed = Join-Path $script:OutDir 'wb-fe-test2-renamed.xlsx'
        if (Test-Path $renamed) { Remove-Item $renamed -Force }
        $xl.DisplayAlerts = $false
        $c.wbB.SaveAs($renamed, 51)
        $xl.DisplayAlerts = $true
        $newName = $c.wbB.Name
        $p1 = Probe-Workbook $fe $c.nameB 'renamed-old'
        $p2 = Probe-Workbook $fe $newName 'renamed-new'
        Check 'after Save As the old name is Gone and the new name Open' ($p1 -and $p2 -and $newName -ne $c.nameB -and $p1.presence -eq 'Gone' -and $p2.presence -eq 'Open') @{ old = $p1.presence; new = $p2.presence; newName = $newName }

        # ---- Don't Save: the workbook is Gone afterwards ----
        Make-Dirty $c.wbB 'Calc' 'C9' 2
        Post-WindowClose (Wb-Hwnd $c.wbB)
        $prompt = Wait-ExcelPrompt $fe
        if ($prompt) { [void](Press-PromptButton $prompt 'DontSave') }
        [void](Wait-Until { $null -eq (Get-ExcelPrompt $fe) -and (Test-WorkbookOpen $fe $newName) -eq $false } 6000)
        Start-Sleep -Milliseconds 600
        $p = Probe-Workbook $fe $newName 'after-close'
        Check 'after Don''t Save the workbook is Gone' ($p -and $p.presence -eq 'Gone') $p
    }
    finally { try { $xl.DisplayAlerts = $false } catch { } }
    Check-LogClean
    End-Scenario
}

# Hand-overs the controller logged since the scenario started: "window 0x.. is being deactivated; the explorer goes to window 0x..".
function Get-HandOvers {
    $lines = Log-Matches 'is being deactivated; the explorer goes to'
    @($lines | ForEach-Object {
        if ($_ -match 'window 0x([0-9A-F]+) is being deactivated; the explorer goes to (?:window 0x([0-9A-F]+)|no other window)') {
            [pscustomobject]@{ From = [Convert]::ToInt64($Matches[1], 16); To = $(if ($Matches[2]) { [Convert]::ToInt64($Matches[2], 16) } else { 0 }) }
        }
    })
}

function S14f {
    Begin-Scenario 'S14f' 'The explorer is never handed to a hidden or minimized window: a hidden workbook (like PERSONAL.XLSB) and a minimized one exist while windows are switched and closed'
    Reset-Env
    $fe = $script:fe; $xl = $fe.Excel
    $wbA = $fe.Workbook
    $hA = Wb-Hwnd $wbA

    # a hidden workbook, a minimized workbook and a normal one (B)
    $wbH = $xl.Workbooks.Add()
    $hH = Wb-Hwnd $wbH
    $wbH.Windows.Item(1).Visible = $false
    $wbM = Open-Workbook (Join-Path $script:OutDir 'fe-test2.xlsx') 'wb-fe-test3.xlsx'
    $hM = Wb-Hwnd $wbM
    $wbM.Windows.Item(1).WindowState = -4140     # xlMinimized
    $wbB = Open-Workbook (Join-Path $script:OutDir 'fe-test2.xlsx') 'wb-fe-test2.xlsx'
    $hB = Wb-Hwnd $wbB
    $wbA.Activate(); Start-Sleep -Milliseconds 800
    Check 'setup: H is hidden, M minimized, A and B are normal windows' (-not [FeWin2]::IsWindowVisible([IntPtr]$hH) -and [FeWin2]::IsIconic([IntPtr]$hM) -and [FeWin2]::IsWindowVisible([IntPtr]$hA) -and [FeWin2]::IsWindowVisible([IntPtr]$hB)) @{ Hvisible = [FeWin2]::IsWindowVisible([IntPtr]$hH); Miconic = [FeWin2]::IsIconic([IntPtr]$hM) }
    $s = Open-Explorer $fe 'Val_FO' 'K33'
    $eh = [int64]$s.explorerHwnd
    Check 'setup: the explorer is owned by A' ((Get-OwnerOf $eh) -eq $hA) @{ owner = (Get-OwnerOf $eh); A = $hA }

    # ---- (1) a plain window switch A -> B and back, several times ----
    foreach ($i in 1..3) {
        $wbB.Activate(); Start-Sleep -Milliseconds 600
        $wbA.Activate(); Start-Sleep -Milliseconds 600
    }
    $moves = @(Get-HandOvers)
    Check 'the controller handed the explorer over at the deactivations (the log has them)' ($moves.Count -ge 3) @{ moves = $moves.Count }
    $bad = @($moves | Where-Object { $_.To -eq $hH -or $_.To -eq $hM })
    Check 'never to the hidden workbook''s window nor to the minimized one' ($bad.Count -eq 0) ($bad | ForEach-Object { "0x{0:X}->0x{1:X}" -f $_.From, $_.To })
    $s = Get-FeState $fe
    Check 'after the switches the explorer is open, owned by A and on top of it' ($s.isOpen -and (Get-OwnerOf $eh) -eq $hA -and (Explorer-OnTop $eh)) @{ open = $s.isOpen; owner = (Get-OwnerOf $eh) }
    Save-FeShot $fe 'S14f-after-switches' | Out-Null

    # ---- (2) the owner window is closed while the only other window of its workbook is minimized ----
    $w1 = $wbA.Windows.Item(1)    # (Item(1) is the ACTIVE window of the workbook: keep a reference to the first one before opening the second)
    $w2 = $wbA.NewWindow()
    Start-Sleep -Milliseconds 1000
    $h2 = [int64]$w2.Hwnd
    $w1.WindowState = -4140       # A's first window (not the owner) is minimized
    Start-Sleep -Milliseconds 600
    Check 'setup: A has a second, active window that owns the explorer (visible); A''s first window is minimized' ((Get-OwnerOf $eh) -eq $h2 -and [FeWin2]::IsIconic([IntPtr]$hA) -and [FeWin2]::IsWindowVisible([IntPtr]$eh)) @{ owner = (Get-OwnerOf $eh); A2 = $h2; iconic = [FeWin2]::IsIconic([IntPtr]$hA); visible = [FeWin2]::IsWindowVisible([IntPtr]$eh) }
    $mark = @(Get-HandOvers).Count
    $w2.Close()
    Start-Sleep -Milliseconds 1800
    $s = Get-FeState $fe
    Check 'closing the owner window (its sibling minimized): the explorer survives' ($s.isOpen -and [FeWin2]::IsWindow([IntPtr]$eh)) @{ open = $s.isOpen; alive = [FeWin2]::IsWindow([IntPtr]$eh) }
    $moves = @(Get-HandOvers | Select-Object -Skip $mark)
    Check 'it was handed to a window that exists and is visible: not the closing one, not A''s minimized one, not H, not M' ($moves.Count -ge 1 -and $moves[0].To -ne 0 -and $moves[0].To -notin @($h2, $hA, $hH, $hM)) ($moves | ForEach-Object { "0x{0:X}->0x{1:X}" -f $_.From, $_.To })
    Check 'the explorer is visible, owned by a visible window' ([FeWin2]::IsWindowVisible([IntPtr]$eh) -and [FeWin2]::IsWindowVisible([IntPtr](Get-OwnerOf $eh))) @{ visible = [FeWin2]::IsWindowVisible([IntPtr]$eh); owner = (Get-OwnerOf $eh) }
    $ownerNow = Get-OwnerOf $eh
    Note ("owner now 0x{0:X} (A1 0x{1:X} iconic={2}, B 0x{3:X}, H 0x{4:X}, M 0x{5:X}); active workbook {6}" -f $ownerNow, $hA, [FeWin2]::IsIconic([IntPtr]$hA), $hB, $hH, $hM, $xl.ActiveWorkbook.Name)
    Save-FeShot $fe 'S14f-after-owner-closed' | Out-Null
    # restore A's window: the explorer follows and is visible again
    $w1.WindowState = -4137
    $wbA.Activate(); Start-Sleep -Milliseconds 1200
    $s = Get-FeState $fe
    Check 'with A''s window restored the explorer is owned by it, visible and on top' ($s.isOpen -and (Get-OwnerOf $eh) -eq $hA -and [FeWin2]::IsWindowVisible([IntPtr]$eh) -and (Explorer-OnTop $eh)) @{ owner = (Get-OwnerOf $eh); visible = [FeWin2]::IsWindowVisible([IntPtr]$eh) }
    Check-LogClean
    End-Scenario
}
