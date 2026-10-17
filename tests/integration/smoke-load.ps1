# Smoke test: start Excel, load the add-in and report whether it registered.
param([string]$Xll = "$PSScriptRoot\..\..\src\FormulaExplorer.AddIn\bin\Debug\net10.0-windows\FormulaExplorer64.xll")
$ErrorActionPreference = 'Stop'
$xl = New-Object -ComObject Excel.Application
try {
    $xl.Visible = $true
    $wb = $xl.Workbooks.Add()
    $ok = $xl.RegisterXLL((Resolve-Path $Xll).Path)
    "RegisterXLL=$ok"
    foreach ($a in $xl.AddIns2) { if ($a.Name -like 'FormulaExplorer*') { "AddIn: $($a.FullName) installed=$($a.Installed)" } }
    Start-Sleep -Seconds 2
    $wb.Close($false)
} finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
}
