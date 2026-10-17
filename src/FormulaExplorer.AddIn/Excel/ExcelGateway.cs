using System.Globalization;
using System.Runtime.InteropServices;
using FormulaExplorer.AddIn.Diagnostics;
using FormulaExplorer.Core.Values;
using XL = Microsoft.Office.Interop.Excel;

namespace FormulaExplorer.AddIn.Excel;

/// <summary>
/// COM implementation of <see cref="IExcelGateway"/>: the only component of the session that talks to the Excel object
/// model. Everything runs on Excel's main thread.
/// <para>Facts verified against Excel 16 (pt-BR) with tests/integration/probe-evaluate.ps1 (late-bound) and
/// tests/integration/GatewayHarness (this class against a live, invisible Excel):</para>
/// <list type="bullet">
/// <item><c>Worksheet.Evaluate</c>, called through the PIA, always parses the ENGLISH syntax of <c>Range.Formula</c>
/// (',' arguments, '.' decimals, "{1,2;3,4}" arrays) whatever the UI culture or thread culture is, so fragments are
/// passed as they are (late-bound callers such as PowerShell use the user's LCID and need the local separators —
/// beware when probing). Its input is limited to 255 characters (256 gives #VALUE!). It returns a Range for
/// references, names, <c>(K$31)</c>, <c>INDEX(...)</c>, <c>OFFSET(...)</c>, <c>IF(c,A1,B1)</c>, a union
/// <c>(A1:A2,C1)</c> (two areas), an intersection and a structured reference; a scalar (Double/String/Boolean) for
/// computations; an Int32 <c>0x800A07xx</c> for errors; a 1-based <c>object[,]</c> — or a 1-based one-dimensional
/// array for a single row — for arrays (whole-column arithmetic gives 1,048,576 elements); #REF! for a 3-D reference
/// (<c>Sheet1:Sheet3!A1</c>) and for a reference to a closed workbook.</item>
/// <item><c>Range.Address</c> separates the areas of a union with ',' in every locale, but <c>Worksheet.Range(address)</c>
/// only accepts the UI list separator (';' in pt-BR; the ',' form fails with 0x800A03EC), so multi-area addresses are
/// resolved area by area and joined with Application.Union.</item>
/// <item><c>Range.Hidden</c> of a multi-row range returns the state of its FIRST row (not null for a mixed range), and
/// <c>OutlineLevel</c> of a range with different levels is DBNull, so hidden rows are found one by one (and only when the
/// span's outline level can be above 1).</item>
/// <item><c>Range.Hidden</c> can only be SET on entire rows/columns: <c>Worksheet.Range("5:5").Hidden = true</c> fails
/// (0x800A03EC) while <c>.EntireRow.Hidden = true</c> works.</item>
/// <item>Hiding a sheet while the workbook structure is protected fails (0x800A03EC); a very hidden sheet is still
/// reachable by name in Worksheets and Evaluate.</item>
/// <item><c>Application.Goto</c> of a multi-area range selects all areas and leaves the LAST cell active.</item>
/// <item>VT_ERROR values arrive as Int32; the PIA events (<see cref="XL.AppEvents_Event"/>) work on .NET and are raised
/// synchronously inside our own COM calls.</item>
/// </list>
/// </summary>
public sealed class ExcelGateway : IExcelGateway
{
    /// <summary>Worksheet.Evaluate refuses longer text.</summary>
    private const int MaxEvaluateLength = 255;

    /// <summary>NavigateTo does not scan areas taller than this many rows for outline-hidden rows.</summary>
    private const int MaxRowsToScan = 1000;

    /// <summary>NavigateTo does not scan areas wider than this many columns for outline-hidden columns.</summary>
    private const int MaxColumnsToScan = 200;

    private const int MaxColumns = 16_384;

    private readonly XL.Application _app;
    private readonly XL.AppEvents_Event _events;
    private readonly RestoreLedger _ledger = new();
    private readonly SelectionTracker _selection;
    private int _lastWindowHwnd;
    private int _suppress;
    private bool _disposed;

    /// <param name="application">Excel Application (ExcelDnaUtil.Application cast to the PIA type). Events are hooked here.</param>
    public ExcelGateway(XL.Application application)
    {
        _app = application ?? throw new ArgumentNullException(nameof(application));
        _events = (XL.AppEvents_Event)application;

        // Initial state; a busy Excel just leaves it unknown until the first event.
        _selection = new SelectionTracker(TryReadActiveLocation(), TryReadSelectionAddress());
        _lastWindowHwnd = TryReadActiveWindowHwnd();

        try
        {
            HookEvents();
        }
        catch
        {
            UnhookEvents();
            throw;
        }
    }

    public event EventHandler<CellLocation>? UserSelectionChanged;
    public event EventHandler<CellLocation>? CellsChanged;
    public event EventHandler? ValuesInvalidated;
    public event EventHandler<IntPtr>? ActiveWindowChanged;

    /// <summary>
    /// A workbook is about to close (the user may still cancel); the argument is its name. Not part of
    /// <see cref="IExcelGateway"/>: only the composition root needs it, to move the explorer window off a workbook window
    /// that is about to be destroyed (an owned window dies with its owner).
    /// </summary>
    public event EventHandler<string>? WorkbookClosing;

    /// <summary>A workbook window (argument: its XLMAIN) is being deactivated; it still exists. See <see cref="OnWindowDeactivate"/>.</summary>
    public event EventHandler<IntPtr>? WindowDeactivating;

    /// <summary>The top-level window of every window of every open workbook; empty when unavailable.</summary>
    public IReadOnlyList<IntPtr> GetAllWindowHandles()
    {
        var handles = new List<IntPtr>();
        XL.Windows? windows = null;
        try
        {
            windows = _app.Windows;
            for (int i = 1; i <= windows.Count; i++)
            {
                XL.Window? window = null;
                try
                {
                    window = windows[i];
                    handles.Add(new IntPtr(window.Hwnd));
                }
                finally { Release(window); }
            }
        }
        catch (Exception ex) when (ex is COMException or ArgumentException or InvalidCastException)
        {
            // busy: what was read so far is the best answer
        }
        finally { Release(windows); }
        return handles;
    }

    /// <summary>The top-level window (XLMAIN) of the first window of the open workbook <paramref name="workbookName"/>; zero when unavailable.</summary>
    public IntPtr GetWindowHandleOf(string workbookName)
    {
        var handles = GetWindowHandlesOf(workbookName, firstOnly: true);
        return handles.Count > 0 ? handles[0] : IntPtr.Zero;
    }

    /// <summary>The top-level windows (XLMAIN) of every window of the open workbook <paramref name="workbookName"/>; empty when unavailable.</summary>
    public IReadOnlyList<IntPtr> GetWindowHandlesOf(string workbookName) => GetWindowHandlesOf(workbookName, firstOnly: false);

    private List<IntPtr> GetWindowHandlesOf(string workbookName, bool firstOnly)
    {
        var handles = new List<IntPtr>();
        XL.Workbooks? books = null;
        XL.Workbook? book = null;
        XL.Windows? windows = null;
        try
        {
            books = _app.Workbooks;
            book = books[workbookName];
            windows = book.Windows;
            int count = firstOnly ? Math.Min(1, windows.Count) : windows.Count;
            for (int i = 1; i <= count; i++)
            {
                XL.Window? window = null;
                try
                {
                    window = windows[i];
                    handles.Add(new IntPtr(window.Hwnd));
                }
                finally { Release(window); }
            }
        }
        catch (Exception ex) when (ex is COMException or ArgumentException or InvalidCastException)
        {
            // not open, or Excel is busy: whatever was read so far is the best answer
        }
        finally
        {
            Release(windows);
            Release(book);
            Release(books);
        }
        return handles;
    }

    /// <summary>
    /// Whether the workbook is still open: <c>Busy</c> when Excel cannot answer (the "Save changes?" prompt of a close in progress
    /// is a modal dialog) or the answer is not conclusive, <c>Gone</c> only when a complete walk over
    /// <c>Application.Workbooks</c> does not find it (see <see cref="WorkbookProbe"/>). The RCWs created by the probe are released:
    /// the poll runs every 250 ms for as long as a close is pending.
    /// </summary>
    public WorkbookPresence ProbeWorkbook(string workbookName)
    {
        XL.Workbooks? books = null;
        try
        {
            books = _app.Workbooks;
            return WorkbookProbe.Probe(
                workbookName,
                () => Release(books[workbookName]),
                () => OpenWorkbookNames(books));
        }
        catch (COMException)
        {
            return WorkbookPresence.Busy;   // Application.Workbooks itself cannot be had: cannot tell
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidCastException)
        {
            return WorkbookPresence.Busy;
        }
        finally { Release(books); }
    }

    private static List<string> OpenWorkbookNames(XL.Workbooks books)
    {
        var names = new List<string>();
        int count = books.Count;
        for (int i = 1; i <= count; i++)
        {
            XL.Workbook? book = null;
            try
            {
                book = books[i];
                names.Add(book.Name);
            }
            finally { Release(book); }
        }
        return names;
    }

    /// <summary>Drops one reference to a COM object obtained from a call (never throws, null is fine).</summary>
    private static void Release(object? com)
    {
        if (com == null || !Marshal.IsComObject(com)) return;
        try
        {
            Marshal.ReleaseComObject(com);
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidComObjectException)
        {
            // already released
        }
    }

    #region Reading cells

    public RootCell? GetActiveCell()
    {
        try
        {
            var cell = ActiveCellOrNull(_app);
            return cell is null ? null : ReadRootCell(cell);
        }
        catch (COMException ex) when (ComErrors.IsBusy(ex)) { throw new ExcelBusyException(ex); }
        catch (Exception ex)
        {
            Log.Error("ExcelGateway.GetActiveCell", ex);
            return null;
        }
    }

    public RootCell? ReadCell(CellLocation cell)
    {
        try
        {
            var range = ResolveRange(cell);
            return range is null ? null : ReadRootCell(range);
        }
        catch (COMException ex) when (ComErrors.IsBusy(ex)) { throw new ExcelBusyException(ex); }
        catch (Exception ex)
        {
            Log.Error($"ExcelGateway.ReadCell {cell}", ex);
            return null;
        }
    }

    private static RootCell ReadRootCell(XL.Range cell)
    {
        var text = cell.Formula as string ?? string.Empty;
        var hasFormula = cell.HasFormula is true; // DBNull for a multi-cell range with mixed content; we read single cells
        return new RootCell(LocationOf(cell), text, hasFormula, ExcelValueConverter.Convert(cell.Value2));
    }

    #endregion

    #region Evaluate

    public NodeEvaluation Evaluate(CellLocation root, string expression)
    {
        if (expression.Length > MaxEvaluateLength)
            return Unavailable("Too long to evaluate");

        try
        {
            var sheet = FindWorksheet(root.WorkbookName, root.SheetName);
            if (sheet is null)
                return Unavailable("The sheet is not available");

            var result = sheet.Evaluate(expression);
            return result is XL.Range range
                ? EvaluateRange(range)
                : new NodeEvaluation(ExcelValueConverter.Convert(result), null);
        }
        catch (COMException ex) when (ComErrors.IsBusy(ex)) { throw new ExcelBusyException(ex); }
        catch (Exception ex)
        {
            // Includes the COM failures Excel reports for fragments it cannot evaluate.
            return Unavailable(ex.Message);
        }
    }

    private static NodeEvaluation Unavailable(string reason) => new(new EvalResult.Unavailable(reason), null);

    private static NodeEvaluation EvaluateRange(XL.Range range)
    {
        var firstArea = (XL.Range)range.Areas[1];
        int rows = firstArea.Rows.Count, columns = firstArea.Columns.Count;

        // Never read Value2 of a whole column: read only the leading block that holds the first 20 cells.
        var (readRows, readColumns) = ExcelValueConverter.LeadingBlock(rows, columns, ValueFormatter.MaxListedValues);
        var values = ExcelValueConverter.ToZeroBasedBlock(firstArea.Resize[readRows, readColumns].Value2, readRows, readColumns);

        var location = LocationOf(range);
        var isSingleCell = System.Convert.ToInt64(range.CountLarge, CultureInfo.InvariantCulture) == 1;
        var firstCell = isSingleCell ? firstArea : (XL.Range)firstArea.Cells[1, 1];
        var firstLocation = isSingleCell ? location : LocationOf(firstCell);
        var target = new RangeTarget(location, firstLocation, isSingleCell, firstCell.HasFormula is true);

        return new NodeEvaluation(new EvalResult.RangeValue(RelativeAddress(range), rows, columns, values), target);
    }

    /// <summary>Converts a raw COM value (Value2 element or Evaluate result that is not a Range) to the COM-free model.</summary>
    public static EvalResult ConvertValue(object? comValue) => ExcelValueConverter.Convert(comValue);

    #endregion

    #region Navigation and restoring

    public NavigationResult NavigateTo(CellLocation target, CellLocation? alsoKeep = null)
    {
        _suppress++;
        try
        {
            // Goto and the unhide calls would silently end (and commit) a cell edit in progress: wait for it to finish.
            ComErrors.ThrowIfEditMode(_app);
            var sheet = FindWorksheet(target.WorkbookName, target.SheetName);
            var workbook = sheet is null ? null : (XL.Workbook)sheet.Parent;

            // Blocked is decided BEFORE anything is restored: it promises that nothing changed, so the previous gold
            // (and the rows/sheets it unhid) must still be valid afterwards.
            if (sheet is not null && workbook!.ProtectStructure && sheet.Visible != XL.XlSheetVisibility.xlSheetVisible)
                return new NavigationResult(NavigationStatus.Blocked,
                    $"The sheet '{target.SheetName}' is hidden and the workbook structure is protected.");

            RestoreCore(alsoKeep is null ? new[] { target } : new[] { target, alsoKeep });

            if (sheet is null)
                return new NavigationResult(NavigationStatus.Failed, $"The sheet '{target.SheetName}' no longer exists.");

            if (sheet.Visible != XL.XlSheetVisibility.xlSheetVisible)
            {
                var original = sheet.Visible;
                sheet.Visible = XL.XlSheetVisibility.xlSheetVisible;
                _ledger.AddSheetVisibility(target.WorkbookName, target.SheetName, (int)original);
            }

            var range = ResolveRange(sheet, target.Address);
            if (range is null)
                return new NavigationResult(NavigationStatus.Failed, $"The range {target.Address} is not valid.");

            UnhideOutlineLines(sheet, target);

            // Never activate a workbook that is already active: that would steal focus from the explorer window.
            var active = _app.ActiveWorkbook;
            if (active is null || !string.Equals(active.Name, target.WorkbookName, StringComparison.OrdinalIgnoreCase))
                workbook!.Activate();
            _app.Goto(range, false);
            return NavigationResult.Ok;
        }
        catch (COMException ex) when (ComErrors.IsBusy(ex)) { throw new ExcelBusyException(ex); }
        catch (Exception ex)
        {
            Log.Error($"ExcelGateway.NavigateTo {target}", ex);
            return new NavigationResult(NavigationStatus.Failed, "Excel could not go to the cell.");
        }
        finally
        {
            _suppress--;
            SyncLastKnown();
        }
    }

    public void RestoreExcept(IReadOnlyCollection<CellLocation> keepVisible)
    {
        _suppress++;
        try
        {
            RestoreCore(keepVisible);
        }
        catch (COMException ex) when (ComErrors.IsBusy(ex)) { throw new ExcelBusyException(ex); }
        finally
        {
            _suppress--;
            SyncLastKnown();
        }
    }

    public void EndSession()
    {
        try
        {
            // Keep what the user is looking at now; if Excel is busy, the last known selection is the best guess.
            CellLocation? current;
            try { current = ReadActiveLocation(); }
            catch (COMException ex) when (ComErrors.IsBusy(ex)) { current = _selection.Location; }

            RestoreExcept(current is null ? Array.Empty<CellLocation>() : new[] { current });
        }
        catch (ExcelBusyException ex)
        {
            // Cell edit mode or a modal dialog: whatever could not be undone (a very hidden sheet made visible, expanded
            // outline rows) is finished later, once Excel takes calls again, instead of being forgotten.
            Log.Info($"ExcelGateway.EndSession: Excel busy, deferring {_ledger.Count} restore record(s): {ex.Message}");
            try { DeferredRestore.Start(_app, _ledger.Detach()); }
            catch (Exception deferEx) { Log.Error("ExcelGateway.EndSession (defer)", deferEx); }
        }
        catch (Exception ex)
        {
            Log.Error("ExcelGateway.EndSession", ex);
        }
        finally
        {
            _ledger.Clear();
        }
    }

    /// <summary>
    /// Undoes (newest first) every record not needed by <paramref name="keepVisible"/>. A record whose workbook or sheet
    /// is gone is dropped silently; if Excel is busy the remaining records stay for the next attempt.
    /// </summary>
    private void RestoreCore(IReadOnlyCollection<CellLocation> keepVisible) => RestoreCore(_app, _ledger, keepVisible);

    /// <param name="checkEditMode">False for the deferred restore, which runs through QueueAsMacro, i.e. only when Excel is not editing (a real busy answer from Excel still defers it again).</param>
    internal static void RestoreCore(XL.Application app, RestoreLedger ledger, IReadOnlyCollection<CellLocation> keepVisible, bool checkEditMode = true)
    {
        var undo = ledger.SelectUndo(keepVisible).ToList();
        if (undo.Count > 0 && checkEditMode) ComErrors.ThrowIfEditMode(app);   // un-hiding/re-hiding would end a cell edit in progress
        foreach (var record in undo)
        {
            try
            {
                Undo(app, record);
            }
            catch (COMException ex) when (ComErrors.IsBusy(ex))
            {
                throw;
            }
            catch (Exception ex)
            {
                Log.Info($"ExcelGateway: dropped restore record ({record.Kind} {record.WorkbookName}/{record.SheetName}): {ex.Message}");
            }
            ledger.Remove(record);
        }
    }

    private static void Undo(XL.Application app, RestoreRecord record)
    {
        var sheet = FindWorksheet(app, record.WorkbookName, record.SheetName);
        if (sheet is null) return;

        switch (record.Kind)
        {
            case RestoreKind.SheetVisibility:
                sheet.Visible = (XL.XlSheetVisibility)record.OriginalVisibility;
                break;
            case RestoreKind.Rows:
                foreach (var (first, last) in ExcelAddressMath.Runs(OutlineRules.LinesToRehide(record.Indexes, line => RowState(sheet, line))))
                    EntireRows(sheet, first, last).Hidden = true;
                break;
            case RestoreKind.Columns:
                foreach (var (first, last) in ExcelAddressMath.Runs(OutlineRules.LinesToRehide(record.Indexes, line => ColumnState(sheet, line))))
                    EntireColumns(sheet, first, last).Hidden = true;
                break;
        }
    }

    private static LineState RowState(XL.Worksheet sheet, int row) =>
        row < 1 || row > ExcelAddressMath.MaxRows ? default : StateOf(EntireRows(sheet, row, row));

    private static LineState ColumnState(XL.Worksheet sheet, int column) =>
        column < 1 || column > MaxColumns ? default : StateOf(EntireColumns(sheet, column, column));

    private static LineState StateOf(XL.Range line) =>
        new(line.Hidden is true, line.OutlineLevel is double level ? (int)level : 0);

    /// <summary>
    /// Unhides ONLY the target's own rows/columns that are hidden by a collapsed outline (OutlineLevel &gt; 1); the rest
    /// of the collapsed group stays hidden, and rows hidden by hand (level 1) are left alone. Each area is scanned only
    /// when it is small enough, and every change is recorded even if a later one fails.
    /// </summary>
    private void UnhideOutlineLines(XL.Worksheet sheet, CellLocation target)
    {
        if (!ExcelAddressMath.TryParseAreas(target.Address, out var areas)) return;

        foreach (var area in areas)
        {
            if (area.RowCount <= MaxRowsToScan)
            {
                var rows = new List<int>();
                try { UnhideHiddenOutlineLines(sheet, area.Row1, area.Row2, EntireRows, rows); }
                finally { _ledger.AddRows(target.WorkbookName, target.SheetName, rows); }
            }
            if (area.ColumnCount <= MaxColumnsToScan)
            {
                var columns = new List<int>();
                try { UnhideHiddenOutlineLines(sheet, area.Col1, area.Col2, EntireColumns, columns); }
                finally { _ledger.AddColumns(target.WorkbookName, target.SheetName, columns); }
            }
        }
    }

    // Hidden can only be SET on a range that Excel knows is entire rows/columns: Worksheet.Range("5:5").Hidden = true
    // fails with 0x800A03EC, Worksheet.Range("5:5").EntireRow.Hidden = true works.

    /// <summary>The entire rows <paramref name="first"/>..<paramref name="last"/>.</summary>
    private static XL.Range EntireRows(XL.Worksheet sheet, int first, int last) => sheet.Range[$"{first}:{last}"].EntireRow;

    /// <summary>The entire columns <paramref name="first"/>..<paramref name="last"/>.</summary>
    private static XL.Range EntireColumns(XL.Worksheet sheet, int first, int last) =>
        sheet.Range[$"{ExcelAddressMath.ColumnName(first)}:{ExcelAddressMath.ColumnName(last)}"].EntireColumn;

    /// <summary>
    /// Unhides the hidden lines with OutlineLevel &gt; 1 among first..last, appending their numbers to <paramref name="unhidden"/>
    /// as they are changed. A protected worksheet refuses to change Hidden: that is logged and ends the pass, the
    /// navigation itself goes on (the target is selected even if it cannot be revealed).
    /// </summary>
    private static void UnhideHiddenOutlineLines(XL.Worksheet sheet, int first, int last, Func<XL.Worksheet, int, int, XL.Range> lines, List<int> unhidden)
    {
        // Range.Hidden only reports the first line of a mixed range, but OutlineLevel is DBNull when the levels differ
        // and a plain level 1 means nothing in the span can be an outline-hidden line (the common case: no outline).
        if (lines(sheet, first, last).OutlineLevel is double spanLevel && spanLevel <= 1) return;

        OutlineRules.UnhideHidden(first, last,
            line => StateOf(lines(sheet, line, line)),
            line => lines(sheet, line, line).Hidden = false,
            unhidden,
            message => Log.Info($"ExcelGateway: {message}"));
    }

    #endregion

    #region Excel events

    private void HookEvents()
    {
        _events.SheetSelectionChange += OnSheetSelectionChange;
        _events.SheetActivate += OnSheetActivate;
        _events.WindowActivate += OnWindowActivate;
        _events.WorkbookActivate += OnWorkbookActivate;
        _events.SheetChange += OnSheetChange;
        _events.SheetCalculate += OnSheetCalculate;
        _events.AfterCalculate += OnAfterCalculate;
        _events.WorkbookBeforeClose += OnWorkbookBeforeClose;
        _events.WindowDeactivate += OnWindowDeactivate;
    }

    private void UnhookEvents()
    {
        // Each removal on its own: the application may already be shutting down.
        Unhook(() => _events.SheetSelectionChange -= OnSheetSelectionChange);
        Unhook(() => _events.SheetActivate -= OnSheetActivate);
        Unhook(() => _events.WindowActivate -= OnWindowActivate);
        Unhook(() => _events.WorkbookActivate -= OnWorkbookActivate);
        Unhook(() => _events.SheetChange -= OnSheetChange);
        Unhook(() => _events.SheetCalculate -= OnSheetCalculate);
        Unhook(() => _events.AfterCalculate -= OnAfterCalculate);
        Unhook(() => _events.WorkbookBeforeClose -= OnWorkbookBeforeClose);
        Unhook(() => _events.WindowDeactivate -= OnWindowDeactivate);
    }

    private static void Unhook(Action unhook)
    {
        try { unhook(); }
        catch (Exception ex) { Log.Info($"ExcelGateway: could not unhook an event: {ex.Message}"); }
    }

    private void OnSheetSelectionChange(object sheet, XL.Range target) => Guard("SheetSelectionChange", TrackSelection);

    private void OnSheetActivate(object sheet) => Guard("SheetActivate", TrackSelection);

    private void OnWorkbookActivate(XL.Workbook workbook) => Guard("WorkbookActivate", TrackSelection);

    private void OnWindowActivate(XL.Workbook workbook, XL.Window window) => Guard("WindowActivate", () =>
    {
        // Reported even while a gateway mutation is in progress: a cross-workbook Goto must re-own the explorer window.
        var hwnd = window.Hwnd;
        if (hwnd != _lastWindowHwnd)
        {
            _lastWindowHwnd = hwnd;
            Raise(ActiveWindowChanged, new IntPtr(hwnd), nameof(ActiveWindowChanged));
        }
        TrackSelection();
    });

    /// <summary>
    /// A workbook window is being deactivated. Excel raises this while the window still exists, both when another window is
    /// about to be activated and when the window itself is being closed (<c>Window.Close</c>, which has no BeforeClose of its
    /// own while the workbook has other windows); the subscriber cannot tell the two apart and must treat both as "this window
    /// may be about to go away".
    /// </summary>
    private void OnWindowDeactivate(XL.Workbook workbook, XL.Window window) => Guard("WindowDeactivate", () =>
        Raise(WindowDeactivating, new IntPtr(window.Hwnd), nameof(WindowDeactivating)));

    private void OnSheetChange(object sheet, XL.Range target) => Guard("SheetChange", () =>
    {
        Raise(CellsChanged, LocationOf(target), nameof(CellsChanged));
        Raise(ValuesInvalidated, nameof(ValuesInvalidated));
    });

    private void OnSheetCalculate(object sheet) => Guard("SheetCalculate", () => Raise(ValuesInvalidated, nameof(ValuesInvalidated)));

    private void OnWorkbookBeforeClose(XL.Workbook workbook, ref bool cancel)
    {
        // The close may still be cancelled (unsaved-changes prompt); the subscriber handles both outcomes.
        string? name = null;
        Guard("WorkbookBeforeClose", () => name = workbook.Name);
        Log.Info($"ExcelGateway: WorkbookBeforeClose raised for {(name is null ? "(name unavailable)" : $"'{name}'")}");
        if (name != null) Raise(WorkbookClosing, name, nameof(WorkbookClosing));
    }

    private void OnAfterCalculate() => Guard("AfterCalculate", () => Raise(ValuesInvalidated, nameof(ValuesInvalidated)));

    /// <summary>
    /// Selection-ish event: while we mutate Excel only remember the new selection; otherwise report the active cell when
    /// the selection differs from the last known one — a different active cell, or the same one with a different selected
    /// address (clicking the first cell of a range the explorer selected).
    /// </summary>
    private void TrackSelection()
    {
        if (_disposed) return;

        var location = ReadActiveLocation();
        if (location is null) return;
        var address = ReadSelectionAddress();

        if (_suppress > 0)
        {
            _selection.Sync(location, address);
            return;
        }
        if (!_selection.Observe(location, address)) return;

        Raise(UserSelectionChanged, location, nameof(UserSelectionChanged));
    }

    private void Raise<T>(EventHandler<T>? handler, T argument, string name)
    {
        try { handler?.Invoke(this, argument); }
        catch (Exception ex) { Log.Error($"ExcelGateway.{name} subscriber", ex); }
    }

    private void Raise(EventHandler? handler, string name)
    {
        try { handler?.Invoke(this, EventArgs.Empty); }
        catch (Exception ex) { Log.Error($"ExcelGateway.{name} subscriber", ex); }
    }

    /// <summary>Runs an Excel event handler body; nothing may escape into Excel. A busy Excel is expected, not an error.</summary>
    private void Guard(string eventName, Action body)
    {
        if (_disposed) return;
        try
        {
            body();
        }
        catch (COMException ex) when (ComErrors.IsBusy(ex))
        {
            Log.Info($"ExcelGateway.{eventName}: Excel busy, event ignored");
        }
        catch (Exception ex)
        {
            Log.Error($"ExcelGateway.{eventName}", ex);
        }
    }

    #endregion

    #region Window, selection and COM helpers

    public IntPtr GetActiveWindowHandle()
    {
        try
        {
            try
            {
                var window = _app.ActiveWindow;
                if (window is not null) return new IntPtr(window.Hwnd);
            }
            catch (COMException ex) when (!ComErrors.IsBusy(ex))
            {
                // No active window (no workbook open): fall back to the application window.
            }
            return new IntPtr(_app.Hwnd);
        }
        catch (COMException ex) when (ComErrors.IsBusy(ex)) { throw new ExcelBusyException(ex); }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        UnhookEvents();
    }

    /// <summary>ActiveCell, or null when there is none (chart sheet, no workbook, protected view). Busy → COMException.</summary>
    private static XL.Range? ActiveCellOrNull(XL.Application app)
    {
        try { return app.ActiveCell; }
        catch (COMException ex) when (!ComErrors.IsBusy(ex)) { return null; }
    }

    internal static CellLocation? ActiveLocation(XL.Application app)
    {
        var cell = ActiveCellOrNull(app);
        return cell is null ? null : LocationOf(cell);
    }

    private CellLocation? ReadActiveLocation() => ActiveLocation(_app);

    /// <summary>Address of the current selection, or null when it is not a range (shape, chart) or there is none. Busy → COMException.</summary>
    private string? ReadSelectionAddress()
    {
        try
        {
            return _app.Selection is XL.Range range
                ? range.get_Address(true, true, XL.XlReferenceStyle.xlA1, Type.Missing, Type.Missing)
                : null;
        }
        catch (COMException ex) when (!ComErrors.IsBusy(ex)) { return null; }
    }

    private CellLocation? TryReadActiveLocation()
    {
        try { return ReadActiveLocation(); }
        catch (Exception ex)
        {
            Log.Info($"ExcelGateway: initial active cell unknown: {ex.Message}");
            return null;
        }
    }

    private string? TryReadSelectionAddress()
    {
        try { return ReadSelectionAddress(); }
        catch (Exception ex)
        {
            Log.Info($"ExcelGateway: initial selection unknown: {ex.Message}");
            return null;
        }
    }

    private int TryReadActiveWindowHwnd()
    {
        try { return _app.ActiveWindow?.Hwnd ?? 0; }
        catch (Exception ex)
        {
            Log.Info($"ExcelGateway: initial active window unknown: {ex.Message}");
            return 0;
        }
    }

    /// <summary>After a mutation: the selection is whatever Excel has now, and must not be reported as a user change.</summary>
    private void SyncLastKnown()
    {
        try
        {
            var location = ReadActiveLocation();
            if (location is not null) _selection.Sync(location, ReadSelectionAddress());
        }
        catch (Exception ex)
        {
            Log.Info($"ExcelGateway: could not refresh the last known selection: {ex.Message}");
        }
    }

    private XL.Worksheet? FindWorksheet(string workbookName, string sheetName) => FindWorksheet(_app, workbookName, sheetName);

    private static XL.Worksheet? FindWorksheet(XL.Application app, string workbookName, string sheetName)
    {
        try
        {
            var workbook = app.Workbooks[workbookName];
            return workbook.Worksheets[sheetName] as XL.Worksheet;
        }
        catch (COMException ex) when (!ComErrors.IsBusy(ex))
        {
            return null; // not open / no such worksheet (chart sheets are not worksheets)
        }
    }

    private XL.Range? ResolveRange(CellLocation location)
    {
        var sheet = FindWorksheet(location.WorkbookName, location.SheetName);
        return sheet is null ? null : ResolveRange(sheet, location.Address);
    }

    /// <summary>
    /// The range for an <c>Address(true, true)</c> string. A multi-area address is resolved area by area: Worksheet.Range
    /// parses it with the UI list separator, not the ',' that Range.Address writes.
    /// </summary>
    private XL.Range? ResolveRange(XL.Worksheet sheet, string address)
    {
        try
        {
            XL.Range? result = null;
            foreach (var part in address.Split(','))
            {
                var area = sheet.Range[part];
                result = result is null ? area : _app.Union(result, area);
            }
            return result;
        }
        catch (COMException ex) when (!ComErrors.IsBusy(ex))
        {
            return null; // invalid address
        }
    }

    private static CellLocation LocationOf(XL.Range range)
    {
        var sheet = range.Worksheet;
        var workbook = (XL.Workbook)sheet.Parent;
        return new CellLocation(workbook.Name, sheet.Name, range.get_Address(true, true, XL.XlReferenceStyle.xlA1, Type.Missing, Type.Missing));
    }

    private static string RelativeAddress(XL.Range range) =>
        range.get_Address(false, false, XL.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);

    #endregion
}
