# S18b: Excel quit (WM_CLOSE on its main windows, what a program or the shell does) with a DIRTY workbook while the explorer is open:
# Excel raises Workbook_BeforeClose for every workbook and shows the save prompt, which can stay up as long as the user likes.

function S18b {
    Begin-Scenario 'S18b' 'Excel quit with a dirty workbook (save prompt) while the explorer is open: cancelled quit keeps everything, confirmed quit exits cleanly'
    Reset-Env
    $fe = $script:fe
    $xl = $fe.Excel
    $xl.DisplayAlerts = $true
    $s = Open-Explorer $fe 'Val_FO' 'K36'
    [FeWin]::Foreground([IntPtr][int64]$s.explorerHwnd); Start-Sleep -Milliseconds 300
    # a gold link into a hidden sheet: the session has something to restore at the end
    Click-Node $fe $s 'Hidden_FO!B2' 1 2
    $r = Wait-Timed $fe { $null -ne $_.view.goldStart } 3000
    Check 'precondition: Hidden_FO was unhidden by the click' ($xl.ActiveWorkbook.Worksheets.Item('Hidden_FO').Visible -eq -1) ''
    Make-Dirty $fe.Workbook 'Val_FO' 'Z1' 1
    Check 'precondition: the workbook has unsaved changes' (-not $fe.Workbook.Saved) ''
    $pid0 = [uint32]$fe.Pid
    $eh = [int64]$s.explorerHwnd
    $hwnds = @([FeWin2]::TopLevel($pid0) | Where-Object { $_.Split('|')[1] -eq 'XLMAIN' } | ForEach-Object { [int64]$_.Split('|')[0] })

    # ---- (1) a quit the user cancels ----
    foreach ($h in $hwnds) { Post-AppClose $h }
    $prompt = Wait-ExcelPrompt $fe 6000
    Check 'the save prompt appears' ($null -ne $prompt) $prompt
    Start-Sleep -Milliseconds 2500
    $prompt = Get-ExcelPrompt $fe
    Check 'after 2.5 s the prompt is still up and the explorer window is intact' ($null -ne $prompt -and [FeWin2]::IsWindow([IntPtr]$eh)) ''
    Save-FeShot $fe 'S18b-prompt' | Out-Null
    if ($prompt) { [void](Press-PromptButton $prompt 'Cancel') }
    [void](Wait-Until { $null -eq (Get-ExcelPrompt $fe) } 5000)
    Start-Sleep -Milliseconds 2000
    $alive = Get-ExcelAlive
    Check 'the cancelled quit leaves Excel running and answering' $alive ''
    if ($alive) {
        $s = Get-FeState $fe
        Check 'the explorer is still open and working after the cancelled quit (same window, gold kept)' ($s.isOpen -and [int64]$s.explorerHwnd -eq $eh -and $null -ne $s.view.goldStart) @{ open = $s.isOpen; gold = (Get-GoldText $s) }
        Check 'the session still owns what it unhid (Hidden_FO visible, selection there)' ($xl.ActiveWorkbook.Worksheets.Item('Hidden_FO').Visible -eq -1) ''
        $p = Char-Point $s (Ix $s '+')
        Move-Mouse $p.X $p.Y
        $r = Wait-Tooltip $fe 2500
        Check 'hover still shows the tooltip' $r.view.tooltipVisible $r.view.tooltipText
        $hwnds = @([FeWin2]::TopLevel($pid0) | Where-Object { $_.Split('|')[1] -eq 'XLMAIN' } | ForEach-Object { [int64]$_.Split('|')[0] })
    }

    # ---- (2) the same quit, confirmed with Don't Save after a long wait ----
    # A well-behaved client lets go of Excel (references held by this script would keep EXCEL.EXE alive after the UI closes).
    $fe.Workbook = $null
    $xlApp = $fe.Excel
    $fe.Excel = $null
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xlApp)
    $xlApp = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    $sw = [Diagnostics.Stopwatch]::StartNew()
    foreach ($h in $hwnds) { Post-AppClose $h }
    $prompt = $null
    Wait-Until { $script:q = Get-ExcelPrompt $fe; $null -ne $script:q } 6000 | Out-Null
    $prompt = $script:q
    Check 'the save prompt appears again' ($null -ne $prompt) ''
    Start-Sleep -Milliseconds 2500
    $prompt = Get-ExcelPrompt $fe
    if ($prompt) { [void](Press-PromptButton $prompt 'DontSave') }
    $exited = $false; $dialogSeen = $null
    $sw.Restart()
    while ($sw.ElapsedMilliseconds -lt 12000) {
        if (-not (Get-Process -Id $pid0 -ErrorAction SilentlyContinue)) { $exited = $true; break }
        foreach ($line in [FeWin2]::TopLevel($pid0)) { $c = $line.Split('|'); if ($c[1] -eq '#32770') { $dialogSeen = $line } }
        Start-Sleep -Milliseconds 100
    }
    $ms = $sw.ElapsedMilliseconds
    Check 'after Don''t Save the Excel process exits within 10 s' ($exited -and $ms -le 10000) "exited=$exited after $ms ms"
    Check 'no crash / message dialog appeared while quitting' ($null -eq $dialogSeen) $dialogSeen
    if (-not $exited) {
        Note ("windows left: " + (([FeWin2]::TopLevel($pid0)) -join ' ; '))
        $p = Get-Process -Id $pid0 -ErrorAction SilentlyContinue
        if ($p -and $p.ProcessName -eq 'EXCEL') { Stop-Process -Id $pid0 -Force }
    }
    try {
        $since = (Get-Date).AddMinutes(-3)
        $ev = @(Get-WinEvent -FilterHashtable @{ LogName = 'Application'; ProviderName = 'Application Error', 'Application Hang'; StartTime = $since } -ErrorAction SilentlyContinue | Where-Object { $_.Message -match 'EXCEL' })
        Check 'no Windows crash/hang event for EXCEL.EXE was recorded' ($ev.Count -eq 0) ($ev | Select-Object -First 1 | ForEach-Object { $_.Message.Substring(0, [math]::Min(160, $_.Message.Length)) })
    }
    catch { Note "event log not readable: $($_.Exception.Message)" }
    $lines = Get-RealLogLines $script:OutDir
    $new = @($lines | Select-Object -Skip $script:LogBase | Where-Object { $_ -match '\[ERROR\]' })
    Check 'no ERROR entries in the add-in log during the whole sequence' ($new.Count -eq 0) ($new | Select-Object -First 3)
    # the shared Excel is gone: the next scenario starts a fresh one
    End-Scenario
}
