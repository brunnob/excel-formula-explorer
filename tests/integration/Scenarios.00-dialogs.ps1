# Helpers for the scenarios that drive Excel's own modal prompts ("Do you want to save the changes?") and workbook closing
# through the UI. Loaded first (file name order); the shared helpers live in ScenarioLib.ps1 / Scenarios.Common.ps1.
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes

# The first dialog/prompt window of the Excel process that is not the explorer's own message box.
function Get-ExcelPrompt($fe) {
    foreach ($line in [FeWin2]::TopLevel([uint32]$fe.Pid)) {
        $p = $line.Split('|')
        if ($p[1] -in @('NUIDialog', '#32770') -and $p[2] -ne 'Formula Explorer') {
            return @{ hwnd = [int64]$p[0]; cls = $p[1]; title = $p[2] }
        }
    }
    $null
}

function Wait-ExcelPrompt($fe, [int]$TimeoutMs = 6000) {
    $script:waitedPrompt = $null
    [void](Wait-Until { $script:waitedPrompt = Get-ExcelPrompt $fe; $null -ne $script:waitedPrompt } $TimeoutMs 50)
    $script:waitedPrompt
}

# Buttons of a dialog through UI Automation, left to right.
function Get-PromptButtons([int64]$Hwnd) {
    $el = [System.Windows.Automation.AutomationElement]::FromHandle([IntPtr]$Hwnd)
    $cond = [System.Windows.Automation.PropertyCondition]::new([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Button)
    @($el.FindAll([System.Windows.Automation.TreeScope]::Descendants, $cond) | ForEach-Object { $_ } | Sort-Object { $_.Current.BoundingRectangle.Left })
}

# Clicks a button of the "Do you want to save the changes?" prompt with the real mouse. The prompt has Save, (Save All),
# Don't Save, Cancel and sometimes an "Open" link; buttons are found by their English or Portuguese names.
function Press-PromptButton($prompt, [ValidateSet('Save', 'DontSave', 'Cancel')][string]$Which) {
    $buttons = @(Get-PromptButtons $prompt.hwnd)
    $pattern = @{ Save = '^(Save|Salvar)$'; DontSave = "^(Don.?t Save|N[aã]o Salvar)$"; Cancel = '^(Cancel|Cancelar)$' }[$Which]
    $btn = $buttons | Where-Object { $_.Current.Name -match $pattern } | Select-Object -First 1
    if (-not $btn) { throw "No '$Which' button in the prompt: $(($buttons | ForEach-Object { $_.Current.Name }) -join ', ')" }
    $r = $btn.Current.BoundingRectangle
    Click-At ([int]($r.Left + $r.Width / 2)) ([int]($r.Top + $r.Height / 2))
    $btn.Current.Name
}

# $true while the named workbook is open, $false when it is not, $null while Excel rejects calls (a modal prompt is up).
function Test-WorkbookOpen($fe, [string]$Name) {
    try { $null = $fe.Excel.Workbooks.Item($Name); $true }
    catch {
        $h = $_.Exception.HResult; if ($_.Exception.InnerException) { $h = $_.Exception.InnerException.HResult }
        if ($h -in @(-2147418111, -2146777998, -2147417846)) { $null } else { $false }   # $null = Excel is busy (a prompt is up), $false = no such workbook
    }
}

# Closes a workbook window like the X button of its title bar does: WM_SYSCOMMAND/SC_CLOSE (a plain WM_CLOSE posted to an Excel
# main window quits the whole application instead). Works on windows that are not active. Excel raises Workbook_BeforeClose
# and then shows the save prompt when the workbook is dirty and alerts are on.
function Post-WindowClose([int64]$Hwnd) { [void][FeWin2]::PostMessage([IntPtr]$Hwnd, 0x112, [IntPtr]0xF060, [IntPtr]::Zero) }

# What a program does to quit Excel (and what S18 does): WM_CLOSE to every main window.
function Post-AppClose([int64]$Hwnd) { [void][FeWin2]::PostMessage([IntPtr]$Hwnd, 0x10, [IntPtr]::Zero, [IntPtr]::Zero) }

# Marks a workbook as having unsaved changes by writing a value into a scratch cell.
function Make-Dirty($wb, [string]$Sheet, [string]$Address, $Value = 1) {
    [void][System.__ComObject].InvokeMember('Value2', [Reflection.BindingFlags]::SetProperty, $null, $wb.Worksheets.Item($Sheet).Range($Address), @($Value))
}

function Get-OwnerOf([int64]$Hwnd) { [FeWin2]::GetWindowLongPtr([IntPtr]$Hwnd, -8).ToInt64() }

# Closes the workbook shown in the top-level window $Hwnd the way a user does: click its title bar (activates it, leaves
# the explorer), then Ctrl+W (Close Window). Excel raises Workbook_BeforeClose BEFORE its save prompt on this path.
function Close-WorkbookWithKeyboard([int64]$Hwnd) {
    $r = Get-Rect ([IntPtr]$Hwnd)
    Click-At ($r[0] + 800) ($r[1] + 14)
    Start-Sleep -Milliseconds 500
    Press-Key 0x57 -Modifiers 0x11
}
