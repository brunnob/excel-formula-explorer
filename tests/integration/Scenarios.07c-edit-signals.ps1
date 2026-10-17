# S15c: the edit-mode signals measured from INSIDE real Excel with the shipped build (FE_Test_EditProbe), in every way a cell can be edited.
# For each state it keeps the raw probe output (editprobe-<state>.json in the scenarios folder) and checks that the editor-window
# signal and the production fallback rule (EditModeSignals.Decide: at least 4 commands answered, at least 3/4 of them disabled) both agree
# with Excel's own opinion (the object model refuses calls while a cell is edited).

# The production fallback rule without a window observation.
function Get-VoteSaysEditing([int]$Disabled, [int]$Probed) { $Probed -ge 4 -and ($Disabled * 4 -ge $Probed * 3) }

function Probe-EditSignals($fe, [string]$State, [scriptblock]$Enter, [int]$DelayMs = 2500) {
    $file = Join-Path $script:OutDir "editprobe-$State.json"
    if (Test-Path $file) { Remove-Item $file -Force }
    Select-FeCell $fe 'Val_FO' 'K33'
    Start-Sleep -Milliseconds 300
    # the hook runs DelayMs later from the add-in's dispatcher (Application.Run from outside is refused while a cell is edited)
    [void](Invoke-Fe2 $fe 'FE_Test_EditProbe' $file $DelayMs)
    Start-Sleep -Milliseconds 200
    & $Enter
    [void](Wait-Until { Test-Path $file } 9000 100)
    $editing = Test-ExcelEditMode $fe
    Leave-EditMode $fe
    Start-Sleep -Milliseconds 500
    if (-not (Test-Path $file)) { return $null }
    $sig = (Get-Content $file -Raw | ConvertFrom-Json).signals
    [pscustomobject]@{ State = $State; ExcelEditing = $editing; Window = [bool]$sig.editorWindowVisible; Disabled = [int]$sig.commandsDisabled; Probed = [int]$sig.commandsProbed; Vote = (Get-VoteSaysEditing ([int]$sig.commandsDisabled) ([int]$sig.commandsProbed)) }
}

function S15c {
    Begin-Scenario 'S15c' 'Edit-mode signals measured in real Excel with the shipped command list: editor window and command vote agree in F2, typing, double click, formula bar and Point mode, and both say "not editing" otherwise'
    Reset-Env
    $fe = $script:fe
    Dismiss-Dialogs

    $normal = Probe-EditSignals $fe 'normal' { } 300
    Check 'normal (no edit): probe written' ($null -ne $normal) ''
    if ($normal) {
        Check 'normal: Excel not editing, no editor window, no command disabled by the vote (0 of the probed ones)' (-not $normal.ExcelEditing -and -not $normal.Window -and -not $normal.Vote -and $normal.Disabled -eq 0) $normal
        Check 'normal: enough commands answer for the vote to count' ($normal.Probed -ge 7) $normal.Probed
    }

    $states = [ordered]@{
        'F2'         = { Click-Cell $fe 'K33' | Out-Null; Start-Sleep -Milliseconds 300; Press-Key $script:VK.F2; Start-Sleep -Milliseconds 300 }
        'typing'     = { Click-Cell $fe 'K33' | Out-Null; Start-Sleep -Milliseconds 300; Press-Key 0x30; Start-Sleep -Milliseconds 300 }
        'dblclick'   = { Click-Cell $fe 'K33' -Double | Out-Null; Start-Sleep -Milliseconds 400 }
        'formulabar' = { Click-Cell $fe 'K33' | Out-Null; Start-Sleep -Milliseconds 300; [void](Click-FormulaBar $fe); Start-Sleep -Milliseconds 400 }
        'point'      = { Click-Cell $fe 'K33' | Out-Null; Start-Sleep -Milliseconds 300; Press-Key 0xBB; Start-Sleep -Milliseconds 300; Press-Key $script:VK.Right; Start-Sleep -Milliseconds 400 }
    }
    foreach ($name in $states.Keys) {
        $r = Probe-EditSignals $fe $name $states[$name]
        Check "$name`: probe written" ($null -ne $r) ''
        if (-not $r) { continue }
        Check "$name`: Excel is editing, the editor window is seen, and the command vote agrees ($($r.Disabled) of $($r.Probed) disabled)" ($r.ExcelEditing -and $r.Window -and $r.Vote) $r
        Check "$name`: even with two commands answering 'enabled' the vote would still hold (threshold, not unanimity)" (Get-VoteSaysEditing ([Math]::Max(0, $r.Disabled - 2)) $r.Probed) $r
    }
    Check-LogClean
    End-Scenario
}
