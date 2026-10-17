# Builds the integration-test workbook that reproduces the model used in the original Formula Explorer video
# ("Mine Aquisition.xlsm") plus edge cases. Usage:
#   New-TestWorkbook.ps1 -Path C:\temp\fe-test.xlsx [-Excel $existingApplication]
# When -Excel is omitted, a private invisible Excel instance is started and quit at the end.
param(
    [Parameter(Mandatory)] [string]$Path,
    $Excel
)
$ErrorActionPreference = 'Stop'
# Excel's IDispatch honors the caller's LCID: with pt-BR regional settings '=0.3' would be parsed as a local formula.
[System.Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('en-US')

$ownExcel = $null -eq $Excel
if ($ownExcel) { $Excel = New-Object -ComObject Excel.Application; $Excel.Visible = $false; $Excel.DisplayAlerts = $false }

function Set-Cells($ws, [hashtable]$cells) {
    foreach ($k in $cells.Keys) {
        $v = $cells[$k]
        try {
            # InvokeMember avoids PowerShell's COM call-site caching bug when the value type changes between calls.
            $r = $ws.Range($k)
            $prop = if ($v -is [string] -and $v.StartsWith('=')) { 'Formula' } else { 'Value2' }
            [void][System.__ComObject].InvokeMember($prop, [Reflection.BindingFlags]::SetProperty, $null, $r, @($v))
        } catch { throw "Set-Cells $($ws.Name)!$k = '$v' failed: $($_.Exception.Message)" }
    }
}

try {
    $wb = $Excel.Workbooks.Add()
    while ($wb.Worksheets.Count -lt 12) { $null = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    $names = 'Val_FO','Names_BA','Sheet8','Stores','Summary','Inventory Forcast','Fin_P_FO','Notes_FO','Macro_FA','Hidden_FO','VeryHidden_FO','Tables'
    for ($i = 0; $i -lt $names.Count; $i++) { $wb.Worksheets.Item($i + 1).Name = "tmp_$i" }
    for ($i = 0; $i -lt $names.Count; $i++) { $wb.Worksheets.Item($i + 1).Name = $names[$i] }

    # --- Names_BA: defined names used by the video formulas ---
    $nb = $wb.Worksheets.Item('Names_BA')
    Set-Cells $nb @{ 'B5' = 'Hundred'; 'C5' = 100; 'B6' = 'Aqui_Per'; 'C6' = 2 }
    $null = $wb.Names.Add('Hundred', '=Names_BA!$C$5')
    $null = $wb.Names.Add('Aqui_Per', '=Names_BA!$C$6')
    $null = $wb.Names.Add('TaxRate', '=3/10')             # name that is not a range (no color, no link); no decimal separator (locale-proof)

    # --- Val_FO: the video's root sheet ---
    $v = $wb.Worksheets.Item('Val_FO')
    Set-Cells $v @{
        'B26' = 'Nominal Discount Rate'; 'I26' = 9.4
        'B27' = 'Real Discount Rate';    'I27' = 6.2
        'B28' = 'Cashflow Timing';       'I28' = 50
        'J3'  = $true; 'K3' = 'On'
        'J19' = 0.252054794520548; 'K19' = 0.252054794520548; 'L19' = 0.25
        'J31' = '=SUM(I$19:$I$19)+J$19*$I$28/Hundred-SUM(OFFSET($I$19,0,0,1,Aqui_Per))'
        'K31' = '=J31+K19'
        'K32' = '=IF(K3="On",1/(((-($I26/Hundred)+1)^(K$31))),0)'
        'K33' = '=1/(((($I27/Hundred)+1)^(K$31)))'
        'K34' = '=IF(J3,(1/((365*(($I27/Hundred)+1)^(K$31)))),0)'
        'K35' = '=(SUM(Fin_P_FO!L39:L40)+Fin_P_FO!L26+Notes_FO!L113+Notes_FO!L101)*(Macro_FA!L183/Hundred)'
        'K36' = '=Hidden_FO!B2+VeryHidden_FO!B3'
        'K37' = '=SUM(Sheet8:Summary!A1)'
        'K38' = '=1/0'
        'K39' = '="abc"&"def"'
        'K40' = '=I26>I27'
        'K41' = '=SUM(Tables!B2:B4)*TaxRate'
        'K42' = '=SUM(INDEX(J19:L19,1,2):L19)'
        'K43' = '=IF(TRUE,I26,I27)'
        'K44' = 123.456
        'K45' = 'plain text'
    }
    # Rows 8:19 grouped and collapsed (row 19 holds J19/K19, which OFFSET/K19 point into).
    $null = $v.Range('A8:A19').EntireRow.Group()
    $null = $v.Outline.ShowLevels(1)
    # Row 22 hidden manually (OutlineLevel 1) to test "only select".
    Set-Cells $v @{ 'J22' = 7; 'K46' = '=J22*2' }
    $v.Range('A22').EntireRow.Hidden = $true

    # --- Other sheets ---
    Set-Cells $wb.Worksheets.Item('Sheet8') @{ 'E7' = 3; 'A1' = 1 }
    $s = $wb.Worksheets.Item('Stores')
    Set-Cells $s @{ 'C4' = 1200; 'D4' = 2400; 'D5' = 5200; 'D14' = 4951536; 'E16' = 9.9; 'E21' = 9820800; 'E30' = 3693144;
                    'F3' = '=Sheet8!E7+(E30/C4)^(D4/D5)+D14-E21*E16+Summary!C5'; 'C3' = $false; 'C11' = 81
                    'F4' = "=IF(C3,0,C11*'Inventory Forcast'!C`$6)" }
    Set-Cells $wb.Worksheets.Item('Summary') @{ 'C5' = 10; 'A1' = 5 }
    Set-Cells $wb.Worksheets.Item('Inventory Forcast') @{ 'C6' = 25 }
    Set-Cells $wb.Worksheets.Item('Fin_P_FO') @{ 'L39' = 1; 'L40' = 2; 'L26' = 3 }
    Set-Cells $wb.Worksheets.Item('Notes_FO') @{ 'L113' = 4; 'L101' = 5; 'L150' = 6 }
    Set-Cells $wb.Worksheets.Item('Macro_FA') @{ 'L183' = 50 }
    Set-Cells $wb.Worksheets.Item('Hidden_FO') @{ 'B2' = 11 }
    Set-Cells $wb.Worksheets.Item('VeryHidden_FO') @{ 'B3' = 22 }

    # --- Tables: a ListObject + frozen panes ---
    $t = $wb.Worksheets.Item('Tables')
    Set-Cells $t @{ 'A1' = 'Item'; 'B1' = 'Amount'; 'A2' = 'a'; 'B2' = 10; 'A3' = 'b'; 'B3' = 20; 'A4' = 'c'; 'B4' = 30 }
    $lo = $t.ListObjects.Add(1, $t.Range('A1:B4'), [Type]::Missing, 1)
    $lo.Name = 'Table1'
    Set-Cells $t @{ 'D2' = '=SUM(Table1[Amount])'; 'D3' = '=B50+B2'; 'B50' = 1 }

    $wb.Worksheets.Item('Hidden_FO').Visible = 0        # xlSheetHidden
    $wb.Worksheets.Item('VeryHidden_FO').Visible = 2    # xlSheetVeryHidden

    $v.Activate()
    $v.Range('K33').Select()
    if (Test-Path $Path) { Remove-Item $Path -Force }
    $wb.SaveAs($Path, 51)   # xlOpenXMLWorkbook
    "Saved $Path"
    if ($ownExcel) { $wb.Close($false) }
}
finally {
    if ($ownExcel) { $Excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) }
}
