# Probe of Excel object-model behaviors that ExcelGateway relies on (Worksheet.Evaluate result types and locale
# behavior, Range addressing, outline hide/unhide, sheet visibility, structure protection, Goto).
# Starts its OWN invisible Excel instance, never touches the keyboard/mouse/screen, always quits it and only ever
# kills the EXCEL.EXE whose PID it started. Output: one "name => result" line per probe.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\integration\probe-evaluate.ps1
#
# LOCALE GOTCHA: PowerShell calls Excel late-bound with the user's LCID, so Worksheet.Evaluate (and Worksheet.Range for a
# union) parse the LOCAL separators here (';' arguments and ',' decimals on pt-BR) while the C# PIA (what the add-in uses,
# see GatewayHarness) always passes English syntax to Evaluate. The probe therefore localizes its English expressions with
# Localize() before evaluating. The add-in itself must never localize.
#
# PowerShell gotcha handled here: a script block that emits a COM Range (or a 2-D array) gets ENUMERATED by the
# pipeline, so every probe returns its value through the unary comma (",$value") to keep it a single object.
$ErrorActionPreference = 'Stop'

Add-Type -Namespace Probe -Name Native -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint pid);
'@

$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$pid_ = [uint32]0
[void][Probe.Native]::GetWindowThreadProcessId([System.IntPtr]$xl.Hwnd, [ref]$pid_)
"Excel $($xl.Version) pid=$pid_ (pre-existing instances: $($before -join ','))"

function Show($v) {
    if ($null -eq $v) { return '$null' }
    if ($v -is [System.DBNull]) { return 'DBNull' }
    if ($v -is [System.Reflection.Missing]) { return 'Missing' }
    if ($v -is [System.Array]) {
        $dims = (0..($v.Rank - 1) | ForEach-Object { "$($v.GetLowerBound($_))..$($v.GetUpperBound($_))" }) -join ','
        $idx = [int[]]@(0..($v.Rank - 1) | ForEach-Object { $v.GetLowerBound($_) })
        $first = $v.GetValue($idx)
        $ft = if ($null -eq $first) { 'null' } else { $first.GetType().Name }
        return "$($v.GetType().Name) dims[$dims] first=$first ($ft)"
    }
    if ($v -is [int]) { return "Int32 0x{0:X8} ({1})" -f $v, $v }
    if ($v -is [__ComObject]) {
        try { return "Range $($v.Address($false, $false)) count=$($v.CountLarge) areas=$($v.Areas.Count)" } catch { return "COM object (not a range)" }
    }
    return "$($v.GetType().Name) $v"
}
function Probe($name, [scriptblock]$body) {
    try { "{0} => {1}" -f $name, (Show (& $body)) }
    catch { "{0} => EXCEPTION 0x{1:X8} {2}" -f $name, $_.Exception.HResult, $_.Exception.Message.Split("`n")[0] }
}

try {
    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -lt 4) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    $s1 = $wb.Worksheets.Item(1); $s2 = $wb.Worksheets.Item(2); $s3 = $wb.Worksheets.Item(3); $s4 = $wb.Worksheets.Item(4)
    # Sheet names must not look like cell addresses (S1, Sh1...) or Names.Add refers-to text becomes ambiguous.
    $s1.Name = 'Alpha'; $s2.Name = 'Beta Two'; $s3.Name = 'Gamma'; $s4.Name = 'Delta'
    $s1.Range('A1').Value2 = 2; $s1.Range('B1').Value2 = 'txt'; $s1.Range('C1').Formula = '=A1*2'
    $s1.Range('A2').Value2 = 3; $s1.Range('A3').Value2 = 4
    $s1.Range('D1').Formula = '=1/0'; $s1.Range('E1').Value2 = $true
    $s1.Range('F1').Formula = '=DATE(2020,1,2)'; $s1.Range('F1').NumberFormat = 'yyyy-mm-dd'
    $s1.Range('G1').Value2 = 12.5; $s1.Range('G1').NumberFormat = '$#,##0.00'
    $s1.Range('H1').Formula = '=NA()'
    $s2.Range('A1').Value2 = 10; $s3.Range('A1').Value2 = 100
    [void]$wb.Names.Add('Hundred', '=Gamma!$A$1')
    [void]$wb.Names.Add('Const', '=42')
    # Names.Add takes the refers-to text in the UI list-separator syntax on some locales (';' there), so try both.
    try { [void]$wb.Names.Add('Multi', '=(Alpha!$A$1:$A$2,Alpha!$C$1)') }
    catch { [void]$wb.Names.Add('Multi', '=(Alpha!$A$1:$A$2;Alpha!$C$1)') }

    '--- Locale (late-bound PowerShell): Evaluate takes English function names but the LOCAL list/decimal/array separators ---'
    # (the two lines below prove the late-bound behavior described in the header)
    "Culture list separator '$((Get-Culture).TextInfo.ListSeparator)', decimal '$((Get-Culture).NumberFormat.NumberDecimalSeparator)'"
    foreach ($e in @('SUM(1,2)', 'SUM(1;2)', 'SOMA(1;2)', 'SOMA(1,2)', 'IF(TRUE,1,2)', 'IF(TRUE;1;2)', 'SE(VERDADEIRO;1;2)', '1.5+1', '1,5+1',
                     'MAX(A1:A3)', 'MAX(A1,A2)', 'MAX(A1;A2)', 'TRUE', 'VERDADEIRO')) {
        Probe "Evaluate($e)" { ,$s1.Evaluate($e) }
    }
    "Application.International: list(5)='$($xl.International(5))' decimal(3)='$($xl.International(3))' thousands(4)='$($xl.International(4))' arrayCol(14)='$($xl.International(14))' arrayRow(15)='$($xl.International(15))' useSystemSeparators=$($xl.UseSystemSeparators)"
    $ls = [string]$xl.International(5); $dec = [string]$xl.International(3)
    $ac = [string]$xl.International(14); $ar = [string]$xl.International(15)
    # Localize: English text -> the UI separators (placeholders: ~ decimal point, | array column, ^^ array row). Needed by
    # this late-bound script only; the add-in passes English text through the PIA.
    function Localize([string]$t) { $t.Replace(',', $ls).Replace('~', $dec).Replace('|', $ac).Replace('^^', $ar) }
    foreach ($e in @('IF(TRUE,A1,B1)', 'OFFSET(A1,0,0,1,2)', 'OFFSET(A1,0,0,1,2)+0', 'INDEX(A1:A3,2)', 'A1:INDEX(A1:A3,2)', 'DATE(2020,1,2)', '(A1:A2,C1)',
                     'MAX(A1,A2)', 'SUM(1~5,2)', 'SUM({1|2^^3|4})', '{1|2^^3|4}', '{1|2}', '{1^^2}', 'SUM(A1:A3,SUM(1,2))', 'IF(A1=2,"x,y",1)', 'INDEX(Multi,2,1)')) {
        $loc = Localize $e
        Probe "Evaluate-localized($loc)" { ,$s1.Evaluate($loc) }
    }
    $s1.Range('J1').Formula = '=IF(TRUE,1.5,2)'
    "Range.Formula = $($s1.Range('J1').Formula)   |   Range.FormulaLocal = $($s1.Range('J1').FormulaLocal)"

    '--- Structured references through Evaluate (table Table1 at Alpha!L1:M3, headers Col1/Col2) ---'
    $s1.Range('L1').Value2 = 'Col1'; $s1.Range('M1').Value2 = 'Col2'
    $s1.Range('L2').Value2 = 5; $s1.Range('M2').Value2 = 6; $s1.Range('L3').Value2 = 7; $s1.Range('M3').Value2 = 8
    $lo = $s1.ListObjects.Add(1, $s1.Range('L1:M3'), [Type]::Missing, 1); $lo.Name = 'Table1'
    foreach ($e in @('Table1[Col1]', 'Table1[#Headers]', 'Table1[[#Headers],[Col1]]', 'Table1[[#Headers];[Col1]]', 'Table1[[#All],[Col1]:[Col2]]', 'Table1[[#All];[Col1]:[Col2]]',
                     'Table1[[#Cabecalhos];[Col1]]', ('Table1[[#Cabe' + [char]0xE7 + 'alhos];[Col1]]'), 'SUM(Table1[Col1])', 'Table1')) {
        Probe "Evaluate($e)" { ,$s1.Evaluate($e) }   # these already mix both separators on purpose
    }

    '--- Worksheet.Evaluate result types (root sheet Alpha) ---'
    foreach ($e in @('A1', '(K$31)', 'IF(TRUE,A1,B1)', 'IF(FALSE,A1,B1)', 'OFFSET(A1,0,0,1,2)', 'OFFSET(A1,0,0,1,2)+0', 'Hundred', 'Const', 'Multi',
                     'Gamma!A1', "'Beta Two'!A1", 'Alpha:Gamma!A1', "'Alpha:Gamma'!A1", '$A$1:$A$3', 'A1:A3*2', 'A:A', '1:1', 'SUM(A1:A3)', 'A1+1', '"a"&B1',
                     'D1', 'H1', 'E1', 'F1', 'F1+0', 'DATE(2020,1,2)', 'TODAY()', 'G1', 'G1+0', 'A1=2', '1/0', 'NA()', 'Nope', 'Nope(1)', 'SUM(',
                     '(A1:A2,C1)', 'A1:B2 A2:C3', 'INDEX(A1:A3,2)', 'A1:INDEX(A1:A3,2)', '{1|2^^3|4}', 'ROW(A1:A3)', 'LEN(B1)', 'B1', 'Z99', 'A1#',
                     'Alpha!A1:Alpha!B1', '[Book9]Alpha!A1')) {
        $loc = Localize $e
        Probe "Evaluate($e)" { ,$s1.Evaluate($loc) }
    }
    # Whole-column arithmetic returns a 1,048,576-element array: inspect it without enumerating.
    $big = $s1.Evaluate('A:A*2')
    "Evaluate(A:A*2) => $($big.GetType().Name) rank=$($big.Rank) len=$($big.Length) lb=$($big.GetLowerBound(0)),$($big.GetLowerBound(1))"

    '--- Evaluate string length limit ---'
    $ok2 = 'A1' + ('+1' * 126)     # 254 chars
    Probe "Evaluate(len=$($ok2.Length))" { ,$s1.Evaluate($ok2) }
    $ok3 = 'A1' + ('+1' * 127)     # 256 chars
    Probe "Evaluate(len=$($ok3.Length))" { ,$s1.Evaluate($ok3) }

    '--- Range addressing / Resize ---'
    $r = $s1.Evaluate('Multi')
    Probe 'Multi.Address(false,false)' { $r.Address($false, $false) }
    Probe 'Multi.Address(true,true)' { $r.Address($true, $true) }
    Probe 'Multi.Areas(1).Address' { $r.Areas.Item(1).Address($true, $true) }
    Probe 'Multi.Areas(1).Cells(1,1).Address' { $r.Areas.Item(1).Cells.Item(1, 1).Address($true, $true) }
    Probe 'Multi.CountLarge' { $r.CountLarge }
    $col = $s1.Evaluate('A:A')
    Probe 'A:A.Address(false,false)' { $col.Address($false, $false) }
    Probe 'A:A.Rows.Count' { $col.Rows.Count }
    Probe 'A:A.Resize(20,1).Value2' { ,$col.Resize(20, 1).Value2 }
    Probe 'A:A.Cells(1,1).HasFormula' { $col.Cells.Item(1, 1).HasFormula }
    Probe 'A1:C1.HasFormula (mixed)' { $s1.Range('A1:C1').HasFormula }
    Probe 'A1:B1.HasFormula (none)' { $s1.Range('A1:B1').HasFormula }
    Probe 'Range.Formula empty cell' { $s1.Range('Z50').Formula }
    Probe 'Range.Value2 date cell F1' { $s1.Range('F1').Value2 }
    Probe 'Range.Value2 currency cell G1' { $s1.Range('G1').Value2 }
    Probe 'Range.Value2 error cell D1' { $s1.Range('D1').Value2 }
    Probe 'Range.Value2 empty cell' { $s1.Range('Z50').Value2 }
    Probe 'Range.Value2 A1:A3 (array)' { ,$s1.Range('A1:A3').Value2 }
    Probe 'Range.Value2 A1:A1 (scalar)' { $s1.Range('A1:A1').Value2 }

    '--- Resolving a multi-area address string back to a Range (locale!) ---'
    Probe 'Worksheet.Range("$A$1:$A$2,$C$1")' { ,$s1.Range('$A$1:$A$2,$C$1') }
    Probe 'Worksheet.Range("$A$1:$A$2;$C$1")' { ,$s1.Range('$A$1:$A$2;$C$1') }
    Probe 'Application.Union(areas)' { ,$xl.Union($s1.Range('$A$1:$A$2'), $s1.Range('$C$1')) }
    Probe 'Worksheet.Range("$A$1")' { ,$s1.Range('$A$1') }
    Probe 'Worksheet.Range("1:1")' { ,$s1.Range('1:1') }

    '--- Outline hide / unhide ---'
    $s4.Range('A1:A30').Value2 = 1
    $s4.Range('8:18').Group() | Out-Null
    $s4.Outline.ShowLevels(1) | Out-Null          # collapse
    Probe 'SET Range("25:25").Hidden = true (not an entire-row range)' { $s4.Range('25:25').Hidden = $true; $s4.Rows.Item(25).Hidden }
    $s4.Range('25:25').EntireRow.Hidden = $true   # a manually hidden row (OutlineLevel 1); Hidden can only be SET on EntireRow/EntireColumn
    Probe 'SET Range("25:25").EntireRow.Hidden = true' { $s4.Rows.Item(25).Hidden }
    Probe 'row 7 Hidden/OutlineLevel' { "$($s4.Rows.Item(7).Hidden)/$($s4.Rows.Item(7).OutlineLevel)" }
    Probe 'row 8 Hidden/OutlineLevel (collapsed group)' { "$($s4.Rows.Item(8).Hidden)/$($s4.Rows.Item(8).OutlineLevel)" }
    Probe 'row 19 Hidden/OutlineLevel (summary row)' { "$($s4.Rows.Item(19).Hidden)/$($s4.Rows.Item(19).OutlineLevel)" }
    Probe 'row 25 Hidden/OutlineLevel (manual)' { "$($s4.Rows.Item(25).Hidden)/$($s4.Rows.Item(25).OutlineLevel)" }
    Probe 'EntireRow(A7:A9).Hidden (mixed)' { $s4.Range('A7:A9').EntireRow.Hidden }
    Probe 'EntireRow(A9).Hidden' { $s4.Range('A9').EntireRow.Hidden }
    Probe 'EntireRow(A1).Hidden' { $s4.Range('A1').EntireRow.Hidden }
    $s4.Range('A12').EntireRow.Hidden = $false
    Probe 'after unhide row 12 only: 11/12/13 Hidden' { "$($s4.Rows.Item(11).Hidden)/$($s4.Rows.Item(12).Hidden)/$($s4.Rows.Item(13).Hidden)" }
    Probe 'row 12 OutlineLevel after unhide' { $s4.Rows.Item(12).OutlineLevel }
    $s4.Range('A12').EntireRow.Hidden = $true
    Probe 'row 12 hidden again' { $s4.Rows.Item(12).Hidden }
    Probe 'EntireRow(A8:A9).Hidden (all hidden)' { $s4.Range('A8:A9').EntireRow.Hidden }
    Probe 'EntireRow(A8:A20).Hidden (starts hidden, ends visible)' { $s4.Range('A8:A20').EntireRow.Hidden }
    Probe 'EntireRow(A1:A30).Hidden (mixed, first visible)' { $s4.Range('A1:A30').EntireRow.Hidden }
    Probe 'EntireRow(A19:A26).Hidden (first visible, hidden later)' { $s4.Range('A19:A26').EntireRow.Hidden }
    Probe 'EntireRow(A25:A26).Hidden (starts hidden)' { $s4.Range('A25:A26').EntireRow.Hidden }
    Probe 'EntireRow(A8:A9).Count / Rows.Count' { "$($s4.Range('A8:A9').EntireRow.Count)/$($s4.Range('A8:A9').EntireRow.Rows.Count)" }
    Probe 'SpecialCells(visible) of EntireRow(A7:A9) count' { $s4.Range('A7:A9').EntireRow.SpecialCells(12).Count }
    Probe 'Rows.Item(8).OutlineLevel via area.Rows.Item(2)' { $s4.Range('A7:A9').Rows.Item(2).OutlineLevel }
    Probe 'EntireRow.OutlineLevel on mixed A1:A30' { $s4.Range('A1:A30').EntireRow.OutlineLevel }
    Probe 'Range.Row/Rows.Count of A7:A9' { "$($s4.Range('A7:A9').Row)/$($s4.Range('A7:A9').Rows.Count)" }
    $s4.Range('C:E').Group() | Out-Null
    $s4.Outline.ShowLevels(1, 1) | Out-Null
    Probe 'col D Hidden/OutlineLevel (collapsed group)' { "$($s4.Columns.Item(4).Hidden)/$($s4.Columns.Item(4).OutlineLevel)" }
    Probe 'EntireColumn(D1).Hidden' { $s4.Range('D1').EntireColumn.Hidden }
    $s4.Columns.Item(4).Hidden = $false
    Probe 'col D unhidden -> C/D/E Hidden' { "$($s4.Columns.Item(3).Hidden)/$($s4.Columns.Item(4).Hidden)/$($s4.Columns.Item(5).Hidden)" }

    '--- Sheet visibility and structure protection ---'
    $s2.Visible = 0     # xlSheetHidden
    $s3.Visible = 2     # xlSheetVeryHidden
    Probe 'Beta Two.Visible (hidden)' { $s2.Visible }
    Probe 'Gamma.Visible (very hidden)' { $s3.Visible }
    Probe 'Worksheets("Gamma") by name when very hidden' { $wb.Worksheets.Item('Gamma').Name }
    Probe 'Evaluate Gamma!A1 on very hidden sheet' { ,$s1.Evaluate('Gamma!A1') }
    Probe 'Workbook.ProtectStructure (no)' { $wb.ProtectStructure }
    $s3.Visible = -1
    Probe 'Gamma unhidden Visible' { $s3.Visible }
    $wb.Protect([Type]::Missing, $true, $false)
    Probe 'Workbook.ProtectStructure (yes)' { $wb.ProtectStructure }
    Probe 'hide Gamma when structure protected' { $s3.Visible = 0; $s3.Visible }
    $wb.Unprotect([Type]::Missing)
    $s2.Visible = -1; $s3.Visible = -1

    '--- Goto / selection / window ---'
    Probe 'ActiveWorkbook.Name' { $xl.ActiveWorkbook.Name }
    Probe 'ActiveWindow.Hwnd' { $xl.ActiveWindow.Hwnd }
    Probe 'Application.Hwnd' { $xl.Hwnd }
    Probe 'Goto(Beta Two!A1) from Alpha' { $xl.Goto($s2.Range('A1'), $false); "$($xl.ActiveSheet.Name)!$($xl.ActiveCell.Address($true, $true, 1))" }
    Probe 'Goto multi-area' { $xl.Goto($s1.Evaluate('Multi'), $false); "$($xl.ActiveSheet.Name)!$($xl.ActiveCell.Address($true, $true, 1)) sel=$($xl.Selection.Address($true, $true, 1))" }
    $wb.Close($false)
}
finally {
    try { $xl.Quit() } catch { }
    try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl) } catch { }
    Start-Sleep -Milliseconds 1500
    if ($pid_ -ne 0 -and ($before -notcontains [int]$pid_) -and (Get-Process -Id $pid_ -ErrorAction SilentlyContinue)) {
        Stop-Process -Id $pid_ -Force   # our own instance, left over after Quit
        "killed leftover own EXCEL.EXE pid=$pid_"
    }
}
