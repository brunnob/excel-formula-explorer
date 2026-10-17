using System.Diagnostics;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using FormulaExplorer.AddIn.Excel;
using FormulaExplorer.Core.Values;
using XL = Microsoft.Office.Interop.Excel;

// Live scenarios for ExcelGateway against an invisible Excel instance. See GatewayHarness.csproj.
OfficeAssemblyResolver.Install(); // before any method mentioning an Excel interop type is JIT-compiled

var failures = 0;
var thread = new Thread(() => failures = Scenarios.Run());
thread.SetApartmentState(ApartmentState.STA); // like Excel's main thread
thread.Start();
thread.Join();
Console.WriteLine(failures == 0 ? "ALL CHECKS PASSED" : $"{failures} CHECK(S) FAILED");
return failures;

internal static class Scenarios
{
    private static int _failures;
    private static readonly List<string> Events = new();

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);

    public static int Run()
    {
        XL.Application? app = null;
        uint pid = 0;
        try
        {
            app = (XL.Application)Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!)!;
            app.Visible = false;
            app.DisplayAlerts = false;
            GetWindowThreadProcessId(new IntPtr(app.Hwnd), out pid);
            Console.WriteLine($"Excel {app.Version} pid={pid}");
            Scenario(app);
        }
        catch (Exception ex)
        {
            Fail("scenario crashed: " + ex);
        }
        finally
        {
            try { app?.Quit(); } catch { /* already gone */ }
            Thread.Sleep(1500);
            if (pid != 0 && Process.GetProcesses().Any(p => p.Id == pid && p.ProcessName.Equals("EXCEL", StringComparison.OrdinalIgnoreCase)))
            {
                Console.WriteLine($"killing leftover own EXCEL.EXE pid={pid}");
                Process.GetProcessById((int)pid).Kill();
            }
        }
        return _failures;
    }

    private static void Check(string name, bool ok, string? detail = null)
    {
        if (!ok) _failures++;
        Console.WriteLine($"{(ok ? "PASS" : "FAIL")}  {name}{(ok || detail is null ? "" : "  -> " + detail)}");
    }

    private static void Fail(string message)
    {
        _failures++;
        Console.WriteLine("FAIL  " + message);
    }

    private static void Same<T>(string name, T expected, T actual) =>
        Check(name, EqualityComparer<T>.Default.Equals(expected, actual), $"expected [{expected}] got [{actual}]");

    private static string Describe(object? v) => v switch
    {
        null => "null",
        EvalResult r => r.ToString() ?? "",
        _ => v.ToString() ?? "",
    };

    private static void Scenario(XL.Application app)
    {
        var book = app.Workbooks.Add();
        var main = (XL.Worksheet)book.Worksheets[1];
        main.Name = "Main";
        var beta = Add(book, "Beta Two");
        var gamma = Add(book, "Gamma");
        var delta = Add(book, "Delta");

        main.Range["A1"].Value2 = 2;
        main.Range["B1"].Value2 = "txt";
        main.Range["C1"].Formula = "=A1*2";
        main.Range["D1"].Formula = "=1/0";
        main.Range["E1"].Value2 = true;
        main.Range["F1"].Formula = "=DATE(2020,1,2)";
        main.Range["F1"].NumberFormat = "yyyy-mm-dd";
        main.Range["G1"].Value2 = 12.5;
        main.Range["G1"].NumberFormat = "$#,##0.00";
        main.Range["J1"].Formula = "=IF(TRUE,1.5,2)";
        main.Range["A2"].Value2 = 3;
        main.Range["A3"].Value2 = 4;
        beta.Range["A1"].Value2 = 10;
        gamma.Range["A1"].Value2 = 100;
        book.Names.Add("Hundred", "=Gamma!$A$1");
        book.Names.Add("Const", "=42");
        try { book.Names.Add("Multi", "=(Main!$A$1:$A$2,Main!$C$1)"); }
        catch (COMException) { book.Names.Add("Multi", "=(Main!$A$1:$A$2;Main!$C$1)"); } // UI list separator
        main.Range["L1"].Value2 = "Col1";
        main.Range["M1"].Value2 = "Col2";
        main.Range["L2"].Value2 = 5;
        main.Range["M2"].Value2 = 6;
        var table = main.ListObjects.Add(XL.XlListObjectSourceType.xlSrcRange, main.Range["L1:M2"], Type.Missing, XL.XlYesNoGuess.xlYes);
        table.Name = "Table1";

        delta.Range["A1:A30"].Value2 = 1;
        delta.Range["8:18"].Group();
        delta.Outline.ShowLevels(1);         // collapse rows 8-18; summary row 19 stays visible
        delta.Range["25:25"].EntireRow.Hidden = true;  // hidden by hand (level 1)
        delta.Range["C:E"].Group();
        delta.Outline.ShowLevels(Type.Missing, 1); // collapse columns C:E
        beta.Visible = XL.XlSheetVisibility.xlSheetHidden;
        gamma.Visible = XL.XlSheetVisibility.xlSheetVeryHidden;
        main.Activate();
        main.Range["A1"].Select();

        using var gateway = new ExcelGateway(app);
        gateway.UserSelectionChanged += (_, l) => Events.Add("select " + l);
        gateway.CellsChanged += (_, l) => Events.Add("cells " + l);
        gateway.ValuesInvalidated += (_, _) => Events.Add("invalidated");
        gateway.ActiveWindowChanged += (_, h) => Events.Add("window " + h);

        var mainLoc = (string a) => new CellLocation(book.Name, "Main", a);
        var deltaLoc = (string a) => new CellLocation(book.Name, "Delta", a);

        Console.WriteLine("--- reading cells");
        var active = gateway.GetActiveCell();
        Same("active cell location", mainLoc("$A$1"), active?.Location);
        Same("active cell formula text", "2", active?.DisplayText);
        var c1 = gateway.ReadCell(mainLoc("$C$1"));
        Check("formula cell", c1 is { HasFormula: true, DisplayText: "=A1*2" } && c1.Value is EvalResult.Number { Value: 4 }, Describe(c1));
        Check("error cell", gateway.ReadCell(mainLoc("$D$1"))?.Value is EvalResult.Error { Code: ExcelError.Div0 });
        Check("date cell is its serial number", gateway.ReadCell(mainLoc("$F$1"))?.Value is EvalResult.Number { Value: 43832 });
        Check("currency cell", gateway.ReadCell(mainLoc("$G$1"))?.Value is EvalResult.Number { Value: 12.5 });
        Check("bool cell", gateway.ReadCell(mainLoc("$E$1"))?.Value is EvalResult.Bool { Value: true });
        var empty = gateway.ReadCell(mainLoc("$Z$50"));
        Check("empty cell", empty is { DisplayText: "", HasFormula: false, Value: EvalResult.Blank });
        Same("Range.Formula of a decimal formula is English", "=IF(TRUE,1.5,2)", gateway.ReadCell(mainLoc("$J$1"))?.DisplayText);
        Check("hidden sheet cell readable", gateway.ReadCell(new CellLocation(book.Name, "Beta Two", "$A$1"))?.Value is EvalResult.Number { Value: 10 });
        Check("unknown sheet -> null", gateway.ReadCell(new CellLocation(book.Name, "Nope", "$A$1")) is null);
        Check("unknown workbook -> null", gateway.ReadCell(new CellLocation("Nope.xlsx", "Main", "$A$1")) is null);
        Check("invalid address -> null", gateway.ReadCell(mainLoc("$ZZZZ$0")) is null);

        Console.WriteLine("--- Evaluate");
        var root = mainLoc("$C$1");
        var e = gateway.Evaluate(root, "A1");
        Check("A1 is a range target", e.Target is { IsSingleCell: true, FirstCellHasFormula: false } t && t.Location == mainLoc("$A$1") && e.Result is EvalResult.RangeValue { Address: "A1", Rows: 1, Columns: 1 } rv && rv.Values[0, 0] is 2.0, Describe(e.Result));
        Check("C1 target has a formula", gateway.Evaluate(root, "C1").Target is { FirstCellHasFormula: true });
        Check("(K$31) is a range", gateway.Evaluate(root, "(K$31)").Target?.Location == mainLoc("$K$31"));
        Check("IF(TRUE,A1,B1) is a range", gateway.Evaluate(root, "IF(TRUE,A1,B1)").Target?.Location == mainLoc("$A$1"));
        e = gateway.Evaluate(root, "OFFSET(A1,0,0,1,2)");
        Check("OFFSET range", e.Result is EvalResult.RangeValue { Address: "A1:B1", Rows: 1, Columns: 2 } r2 && r2.Values[0, 0] is 2.0 && r2.Values[0, 1] is "txt"
            && e.Target is { IsSingleCell: false } t2 && t2.FirstCell == mainLoc("$A$1") && t2.Location == mainLoc("$A$1:$B$1"), Describe(e.Result));
        e = gateway.Evaluate(root, "Hundred");
        Check("name -> range on a very hidden sheet", e.Target?.Location == new CellLocation(book.Name, "Gamma", "$A$1") && e.Result is EvalResult.RangeValue, Describe(e.Target));
        Check("constant name", gateway.Evaluate(root, "Const") is { Result: EvalResult.Number { Value: 42 }, Target: null });
        Check("SUM(A1:A3)", gateway.Evaluate(root, "SUM(A1:A3)").Result is EvalResult.Number { Value: 9 });
        Check("decimal literal is localized", gateway.Evaluate(root, "SUM(1.5,2)").Result is EvalResult.Number { Value: 3.5 }, Describe(gateway.Evaluate(root, "SUM(1.5,2)").Result));
        Check("1/0 error", gateway.Evaluate(root, "1/0").Result is EvalResult.Error { Code: ExcelError.Div0 });
        Check("unknown name", gateway.Evaluate(root, "Nope").Result is EvalResult.Error { Code: ExcelError.Name });
        Check("text", gateway.Evaluate(root, "\"a\"&B1").Result is EvalResult.Text { Value: "atxt" });
        Check("string with comma untouched", gateway.Evaluate(root, "IF(A1=2,\"x,y\",1)").Result is EvalResult.Text { Value: "x,y" });
        Check("bool", gateway.Evaluate(root, "A1=2").Result is EvalResult.Bool { Value: true });
        e = gateway.Evaluate(root, "A1:A3*2");
        Check("array", e.Result is EvalResult.ArrayValue { Rows: 3, Columns: 1 } av && av.Values[0, 0] is 4.0 && av.Values[2, 0] is 8.0 && e.Target is null, Describe(e.Result));
        e = gateway.Evaluate(root, "{1,2;3,4.5}");
        Check("array constant", e.Result is EvalResult.ArrayValue { Rows: 2, Columns: 2 } ac && ac.Values[1, 1] is 4.5, Describe(e.Result));
        e = gateway.Evaluate(root, "A:A*2");
        Check("whole-column arithmetic is capped", e.Result is EvalResult.ArrayValue { Rows: 1048576, Columns: 1 } big && big.Values.Length <= ValueFormatter.MaxListedValues + 1 && big.Values[0, 0] is 4.0, Describe(e.Result));
        e = gateway.Evaluate(root, "A:A");
        Check("whole column range reads 20 cells", e.Result is EvalResult.RangeValue { Address: "A:A", Rows: 1048576, Columns: 1 } cv && cv.Values.GetLength(0) == 20 && cv.Values[0, 0] is 2.0
            && e.Target is { IsSingleCell: false } t3 && t3.Location == mainLoc("$A:$A") && t3.FirstCell == mainLoc("$A$1"), Describe(e.Result));
        e = gateway.Evaluate(root, "Multi");
        Check("multi-area range", e.Result is EvalResult.RangeValue { Address: "A1:A2,C1", Rows: 2, Columns: 1 } && e.Target is { IsSingleCell: false } t4 && t4.Location == mainLoc("$A$1:$A$2,$C$1"), Describe(e.Result) + " / " + Describe(e.Target));
        e = gateway.Evaluate(root, "(A1:A2,C1)");
        Check("union expression", e.Target?.Location == mainLoc("$A$1:$A$2,$C$1"), Describe(e.Target));
        e = gateway.Evaluate(root, "Table1[[#Headers],[Col1]]");
        Check("structured reference", e.Target?.Location == mainLoc("$L$1"), Describe(e.Result));
        Check("3-D reference -> #REF!", gateway.Evaluate(root, "Main:Gamma!A1").Result is EvalResult.Error { Code: ExcelError.Ref });
        Check("256 chars -> Unavailable", gateway.Evaluate(root, "A1" + string.Concat(Enumerable.Repeat("+1", 128))).Result is EvalResult.Unavailable { Reason: "Too long to evaluate" });
        Check("255 chars evaluated", gateway.Evaluate(root, "A1" + string.Concat(Enumerable.Repeat("+1", 126)) + "+").Result is not EvalResult.Unavailable);
        Check("unknown root sheet -> Unavailable", gateway.Evaluate(new CellLocation(book.Name, "Nope", "$A$1"), "A1").Result is EvalResult.Unavailable);
        Check("syntax error is an error value, not an exception", gateway.Evaluate(root, "SUM(").Result is EvalResult.Error);

        Console.WriteLine("--- events and selection tracking");
        Events.Clear();
        main.Range["B2"].Select();
        Same("user selection reported", new[] { "select " + mainLoc("$B$2") }, Events.Where(x => x.StartsWith("select")).ToArray(), SequenceEq);
        Events.Clear();
        main.Range["B2"].Select();
        Check("same selection not reported again", !Events.Any(x => x.StartsWith("select")), string.Join(";", Events));
        Events.Clear();
        main.Range["A1"].Value2 = 5;
        Check("SheetChange -> CellsChanged + ValuesInvalidated", Events.Contains("cells " + mainLoc("$A$1")) && Events.Contains("invalidated"), string.Join(";", Events));
        Events.Clear();
        app.Calculate();
        Check("calculation -> ValuesInvalidated", Events.Contains("invalidated"), string.Join(";", Events));
        main.Range["A1"].Value2 = 2;
        Events.Clear();
        gateway.NavigateTo(mainLoc("$D$4"));
        Check("NavigateTo selection is not a user change", !Events.Any(x => x.StartsWith("select")), string.Join(";", Events));
        Same("NavigateTo moved the selection", mainLoc("$D$4"), gateway.GetActiveCell()?.Location);
        Events.Clear();
        main.Range["C3"].Select();
        Same("user selection after a gateway move is reported", new[] { "select " + mainLoc("$C$3") }, Events.Where(x => x.StartsWith("select")).ToArray(), SequenceEq);

        Console.WriteLine("--- hidden sheets");
        Check("Beta Two starts hidden", beta.Visible == XL.XlSheetVisibility.xlSheetHidden);
        var nav = gateway.NavigateTo(new CellLocation(book.Name, "Beta Two", "$A$1"), mainLoc("$C$3"));
        Check("hidden sheet navigated", nav.Status == NavigationStatus.Ok && beta.Visible == XL.XlSheetVisibility.xlSheetVisible && ActiveSheetName(app) == "Beta Two", nav.Message);
        nav = gateway.NavigateTo(new CellLocation(book.Name, "Gamma", "$A$1"), new CellLocation(book.Name, "Beta Two", "$A$1"));
        Check("very hidden sheet navigated, alsoKeep keeps the previous one", nav.Status == NavigationStatus.Ok && gamma.Visible == XL.XlSheetVisibility.xlSheetVisible && beta.Visible == XL.XlSheetVisibility.xlSheetVisible);
        nav = gateway.NavigateTo(mainLoc("$A$1"));
        Check("moving away restores both hidden sheets", nav.Status == NavigationStatus.Ok && beta.Visible == XL.XlSheetVisibility.xlSheetHidden && gamma.Visible == XL.XlSheetVisibility.xlSheetVeryHidden,
            $"{beta.Visible} / {gamma.Visible}");

        book.Protect(Type.Missing, true, false);
        nav = gateway.NavigateTo(new CellLocation(book.Name, "Gamma", "$A$1"));
        Check("protected structure -> Blocked with the exact message", nav.Status == NavigationStatus.Blocked && nav.Message == "The sheet 'Gamma' is hidden and the workbook structure is protected."
            && gamma.Visible == XL.XlSheetVisibility.xlSheetVeryHidden, nav.Message);
        book.Unprotect(Type.Missing);

        Console.WriteLine("--- outline rows and columns");
        Check("rows 8-18 start hidden", RowHidden(delta, 8) && RowHidden(delta, 12) && RowHidden(delta, 18) && !RowHidden(delta, 19));
        nav = gateway.NavigateTo(deltaLoc("$A$12"));
        Check("only the target row is unhidden", nav.Status == NavigationStatus.Ok && !RowHidden(delta, 12) && RowHidden(delta, 11) && RowHidden(delta, 13) && RowHidden(delta, 8), RowStates(delta, 8, 18));
        nav = gateway.NavigateTo(deltaLoc("$A$14"));
        Check("next target restores the previous row", !RowHidden(delta, 14) && RowHidden(delta, 12), RowStates(delta, 8, 18));
        nav = gateway.NavigateTo(deltaLoc("$A$10:$B$13"), deltaLoc("$A$14"));
        Check("multi-row target and alsoKeep", !RowHidden(delta, 10) && !RowHidden(delta, 13) && !RowHidden(delta, 14) && RowHidden(delta, 9) && RowHidden(delta, 15), RowStates(delta, 8, 18));
        gateway.RestoreExcept(Array.Empty<CellLocation>());
        Check("RestoreExcept(empty) re-hides everything", RowStates(delta, 8, 18) == "HHHHHHHHHHH", RowStates(delta, 8, 18));
        nav = gateway.NavigateTo(deltaLoc("$A$25"));
        Check("manually hidden row is left hidden", RowHidden(delta, 25), nav.Status.ToString());
        nav = gateway.NavigateTo(deltaLoc("$D$1"));
        Check("only the target column is unhidden", nav.Status == NavigationStatus.Ok && !ColumnHidden(delta, 4) && ColumnHidden(delta, 3) && ColumnHidden(delta, 5), ColumnStates(delta, 3, 5));
        nav = gateway.NavigateTo(deltaLoc("$A$12,$A$16"));
        Check("multi-area target unhides each area", nav.Status == NavigationStatus.Ok && !RowHidden(delta, 12) && !RowHidden(delta, 16) && ColumnHidden(delta, 4) && RowHidden(delta, 13), RowStates(delta, 8, 18) + " / " + ColumnStates(delta, 3, 5));
        Same("multi-area selection", "$A$12,$A$16", Sel(app));
        nav = gateway.NavigateTo(deltaLoc("$A$1:$A$5000"));
        Check("huge areas are not scanned", nav.Status == NavigationStatus.Ok && RowHidden(delta, 13) && RowHidden(delta, 8), RowStates(delta, 8, 18));
        gateway.RestoreExcept(Array.Empty<CellLocation>());
        Check("restored again", RowStates(delta, 8, 18) == "HHHHHHHHHHH" && ColumnHidden(delta, 4));

        Console.WriteLine("--- cross-workbook and window handle");
        var book2 = app.Workbooks.Add();
        main.Activate();
        Events.Clear();
        var book2Loc = new CellLocation(book2.Name, ((XL.Worksheet)book2.Worksheets[1]).Name, "$B$2");
        nav = gateway.NavigateTo(book2Loc);
        Check("other workbook activated", nav.Status == NavigationStatus.Ok && string.Equals(app.ActiveWorkbook.Name, book2.Name, StringComparison.OrdinalIgnoreCase));
        Check("ActiveWindowChanged reported even while navigating", Events.Any(x => x.StartsWith("window")), string.Join(";", Events));
        Check("cross-workbook navigation is not a user change", !Events.Any(x => x.StartsWith("select")), string.Join(";", Events));
        Same("GetActiveWindowHandle", new IntPtr(app.ActiveWindow.Hwnd), gateway.GetActiveWindowHandle());
        Events.Clear();
        nav = gateway.NavigateTo(mainLoc("$A$1"));
        Check("back to the first workbook", nav.Status == NavigationStatus.Ok && string.Equals(app.ActiveWorkbook.Name, book.Name, StringComparison.OrdinalIgnoreCase) && Events.Any(x => x.StartsWith("window")), string.Join(";", Events));
        book2.Close(false);

        Console.WriteLine("--- failures");
        Check("unknown sheet -> Failed", gateway.NavigateTo(new CellLocation(book.Name, "Nope", "$A$1")).Status == NavigationStatus.Failed);
        Check("unknown workbook -> Failed", gateway.NavigateTo(new CellLocation("Nope.xlsx", "Main", "$A$1")).Status == NavigationStatus.Failed);
        Check("invalid address -> Failed", gateway.NavigateTo(mainLoc("$ZZZZ$0")).Status == NavigationStatus.Failed);

        Console.WriteLine("--- EndSession");
        gateway.NavigateTo(deltaLoc("$A$12"));
        gateway.NavigateTo(new CellLocation(book.Name, "Gamma", "$A$1"), deltaLoc("$A$12"));
        Check("both pending", !RowHidden(delta, 12) && gamma.Visible == XL.XlSheetVisibility.xlSheetVisible);
        delta.Activate();
        delta.Range["A12"].Select(); // the user parks on the unhidden row
        gateway.EndSession();
        Check("EndSession keeps the current selection visible, restores the rest", !RowHidden(delta, 12) && gamma.Visible == XL.XlSheetVisibility.xlSheetVeryHidden,
            $"{RowHidden(delta, 12)} / {gamma.Visible}");
        gateway.RestoreExcept(Array.Empty<CellLocation>());
        Check("ledger forgotten by EndSession", !RowHidden(delta, 12));

        Console.WriteLine("--- Dispose");
        gateway.Dispose();
        Events.Clear();
        main.Activate();
        main.Range["E5"].Select();
        Check("no events after Dispose", Events.Count == 0, string.Join(";", Events));

        book.Close(false);
    }

    private static XL.Worksheet Add(XL.Workbook book, string name)
    {
        var sheet = (XL.Worksheet)book.Worksheets.Add(Type.Missing, book.Worksheets[book.Worksheets.Count]);
        sheet.Name = name;
        return sheet;
    }

    private static string ActiveSheetName(XL.Application app) => ((XL.Worksheet)app.ActiveSheet).Name;

    private static string Sel(XL.Application app) => ((XL.Range)app.Selection).get_Address(true, true, XL.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

    private static bool RowHidden(XL.Worksheet sheet, int row) => sheet.Range[$"{row}:{row}"].EntireRow.Hidden is true;

    private static bool ColumnHidden(XL.Worksheet sheet, int column) => sheet.Range[$"{ExcelAddressMath.ColumnName(column)}:{ExcelAddressMath.ColumnName(column)}"].EntireColumn.Hidden is true;

    private static string RowStates(XL.Worksheet sheet, int first, int last) =>
        string.Concat(Enumerable.Range(first, last - first + 1).Select(r => RowHidden(sheet, r) ? 'H' : '.'));

    private static string ColumnStates(XL.Worksheet sheet, int first, int last) =>
        string.Concat(Enumerable.Range(first, last - first + 1).Select(c => ColumnHidden(sheet, c) ? 'H' : '.'));

    private static bool SequenceEq(string[] a, string[] b) => a.SequenceEqual(b);

    private static void Same(string name, string[] expected, string[] actual, Func<string[], string[], bool> eq) =>
        Check(name, eq(expected, actual), $"expected [{string.Join(";", expected)}] got [{string.Join(";", actual)}]");
}
