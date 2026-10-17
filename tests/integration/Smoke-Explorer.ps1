# End-to-end smoke test: opens the explorer on Val_FO!K33 of the test workbook, hovers, clicks, drills and goes Back,
# dumping state and taking screenshots at each step. Output folder: $env:TEMP\fe-integration\smoke
param(
    [string]$Xll,
    [string]$OutDir = (Join-Path $PSScriptRoot 'out\smoke')
)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\FeTest.ps1"

New-Item -ItemType Directory -Force $OutDir | Out-Null
$wbPath = Join-Path $OutDir 'fe-test.xlsx'
& "$PSScriptRoot\New-TestWorkbook.ps1" -Path $wbPath | Out-Null

$args0 = @{ Workbook = $wbPath; OutDir = $OutDir }
if ($Xll) { $args0.Xll = $Xll }
$fe = Start-FeExcel @args0
$log = [Collections.Generic.List[string]]::new()
function Note($m) { $log.Add($m); Write-Host $m }
try {
    Select-FeCell $fe 'Val_FO' 'K33'
    Note "Open: $(Invoke-Fe $fe 'FE_Test_Open')"
    $s = Wait-FeState $fe { $_.view -and -not $_.view.isLoading -and $_.view.displayText.Length -gt 0 }
    Note "root=$($s.session.root | ConvertTo-Json -Compress) text=$($s.view.displayText) back=$($s.view.backEnabled) ppd=$($s.view.pixelsPerDip)"
    Note "overlay=$($s.overlay | ConvertTo-Json -Compress)"
    Note "window=$($s.view.screenBounds | ConvertTo-Json -Compress)"
    Save-FeShot $fe 'open' | Out-Null

    $text = $s.view.displayText
    $iCaret = Find-FeIndex $text '^'
    Note "Hover ^ ($iCaret): $(Invoke-Fe $fe 'FE_Test_Hover' $iCaret)"
    Start-Sleep -Milliseconds 700
    $s = Get-FeState $fe
    Note "hover=[$($s.view.hoverStart),$($s.view.hoverLength)] '$($text.Substring($s.view.hoverStart, $s.view.hoverLength))' tooltip=$($s.view.tooltipVisible) '$($s.view.tooltipText)'"
    Save-FeShot $fe 'hover-caret' | Out-Null

    $iHundred = Find-FeIndex $text 'Hundred' 1 2
    Invoke-Fe $fe 'FE_Test_Hover' $iHundred | Out-Null
    Start-Sleep -Milliseconds 700
    $s = Get-FeState $fe
    Note "hover Hundred link=$($s.view.hoverIsLink) tooltip='$($s.view.tooltipText)'"
    Save-FeShot $fe 'hover-name' | Out-Null

    Note "Click Hundred: $(Invoke-Fe $fe 'FE_Test_Click' $iHundred)"
    $s = Wait-FeState $fe { $_.session.goldTarget -ne $null }
    Note "gold=[$($s.session.goldStart),$($s.session.goldLength)] target=$($s.session.goldTarget | ConvertTo-Json -Compress) activeSheet=$($fe.Excel.ActiveSheet.Name) sel=$($fe.Excel.Selection.Address())"
    Save-FeShot $fe 'click-name' | Out-Null

    $iK31 = Find-FeIndex $text 'K$31' 1 1
    Note "DoubleClick K`$31: $(Invoke-Fe $fe 'FE_Test_DoubleClick' $iK31)"
    $s = Wait-FeState $fe { $_.view.displayText -like '=J31*' -and -not $_.view.isLoading }
    Note "after drill: root=$($s.session.root.address) text=$($s.view.displayText) stack=$(@($s.session.stack).Count) back=$($s.view.backEnabled)"
    Note "overlay=$($s.overlay | ConvertTo-Json -Compress)"
    Save-FeShot $fe 'drill-K31' | Out-Null

    $text = $s.view.displayText
    $iJ31 = Find-FeIndex $text 'J31' 1 1
    Invoke-Fe $fe 'FE_Test_DoubleClick' $iJ31 | Out-Null
    $s = Wait-FeState $fe { $_.view.displayText -like '=SUM(*' -and -not $_.view.isLoading }
    Note "after drill 2: root=$($s.session.root.address) text=$($s.view.displayText)"
    $text = $s.view.displayText
    $iOffset = Find-FeIndex $text 'OFFSET' 1 2
    Invoke-Fe $fe 'FE_Test_Hover' $iOffset | Out-Null
    Start-Sleep -Milliseconds 700
    $s = Get-FeState $fe
    Note "hover OFFSET link=$($s.view.hoverIsLink) tooltip='$($s.view.tooltipText)'"
    Save-FeShot $fe 'hover-offset' | Out-Null
    Invoke-Fe $fe 'FE_Test_Click' $iOffset | Out-Null
    $s = Wait-FeState $fe { $_.session.goldTarget -ne $null }
    Note "click OFFSET target=$($s.session.goldTarget | ConvertTo-Json -Compress) row19Hidden=$($fe.Excel.ActiveSheet.Range('A19').EntireRow.Hidden) row18Hidden=$($fe.Excel.ActiveSheet.Range('A18').EntireRow.Hidden)"
    Save-FeShot $fe 'click-offset' | Out-Null

    foreach ($n in 1..3) {
        Invoke-Fe $fe 'FE_Test_Back' | Out-Null
        Start-Sleep -Milliseconds 600
        $s = Wait-FeState $fe { -not $_.view.isLoading }
        Note "Back $n -> root=$($s.session.root.address) gold=$($s.session.goldStart) text=$($s.view.displayText) sel=$($fe.Excel.ActiveCell.Address()) row19Hidden=$($fe.Excel.ActiveSheet.Range('A19').EntireRow.Hidden) back=$($s.view.backEnabled)"
    }
    Save-FeShot $fe 'after-backs' | Out-Null

    Note "Close: $(Invoke-Fe $fe 'FE_Test_Close')"
    $s = Get-FeState $fe
    Note "isOpen=$($s.isOpen)"
    Save-FeShot $fe 'closed' | Out-Null
}
catch {
    Note "FAILED: $($_.Exception.Message) $($_.ScriptStackTrace)"
    Save-FeShot $fe 'failure' | Out-Null
}
finally {
    Stop-FeExcel $fe
    $log | Set-Content (Join-Path $OutDir 'smoke-log.txt') -Encoding UTF8
}
