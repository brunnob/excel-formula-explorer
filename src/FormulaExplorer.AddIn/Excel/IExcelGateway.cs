using FormulaExplorer.Core.Values;

namespace FormulaExplorer.AddIn.Excel;

/// <summary>
/// The ONLY component of the session that talks to the Excel object model (the overlay has its own geometry access).
/// All members must be called on Excel's main thread. Members that need Excel may throw <see cref="ExcelBusyException"/>.
/// A gateway lives for one explorer session (created on open, disposed on close).
/// </summary>
public interface IExcelGateway : IDisposable
{
    /// <summary>The active cell, or null when there is none (chart sheet active, no workbook open, protected view...).</summary>
    RootCell? GetActiveCell();

    /// <summary>Reads a single cell; null if the workbook/sheet no longer exists or the address is invalid.</summary>
    RootCell? ReadCell(CellLocation cell);

    /// <summary>
    /// Evaluates <paramref name="expression"/> (formula text WITHOUT the leading '=') with Worksheet.Evaluate on the
    /// root cell's sheet, so unqualified references resolve on that sheet. Never throws except <see cref="ExcelBusyException"/>.
    /// Expressions longer than 255 characters are not evaluated (Result = Unavailable).
    /// When Excel returns a Range, Result is a <see cref="EvalResult.RangeValue"/> and Target is set.
    /// </summary>
    NodeEvaluation Evaluate(CellLocation root, string expression);

    /// <summary>
    /// Selects <paramref name="target"/> in Excel (activating its workbook window and sheet and scrolling it into view),
    /// first restoring every earlier unhide/expand that is needed neither by <paramref name="target"/> nor by
    /// <paramref name="alsoKeep"/> (the session passes the root cell when the target is a gold link), then unhiding its
    /// sheet (Hidden or VeryHidden) and unhiding its own rows/columns hidden by a collapsed outline (OutlineLevel &gt; 1),
    /// recording each change so it can be undone later. Rows/columns hidden manually (OutlineLevel 1) are left hidden.
    /// Selection events caused by this call are NOT reported as <see cref="UserSelectionChanged"/>.
    /// </summary>
    NavigationResult NavigateTo(CellLocation target, CellLocation? alsoKeep = null);

    /// <summary>
    /// Undoes every recorded unhide/expand that is not needed to keep any of <paramref name="keepVisible"/> visible
    /// (empty = undo all). Does not change the selection. Events caused by this call are suppressed.
    /// </summary>
    void RestoreExcept(IReadOnlyCollection<CellLocation> keepVisible);

    /// <summary>
    /// End of session: undoes recorded changes except those needed by the CURRENT Excel selection (which stays visible),
    /// then forgets all records. Never throws (best effort).
    /// </summary>
    void EndSession();

    /// <summary>Window handle (XLMAIN) of Excel's active workbook window (Window.Hwnd), or Application.Hwnd as fallback.</summary>
    IntPtr GetActiveWindowHandle();

    /// <summary>
    /// The user changed the active cell (clicked a cell, used the keyboard, activated another sheet or workbook window).
    /// Raised only when the new active cell differs from the last known one, and never for changes made by this gateway.
    /// Not raised when there is no active cell (e.g. chart sheet).
    /// </summary>
    event EventHandler<CellLocation>? UserSelectionChanged;

    /// <summary>Cells changed (SheetChange). The argument is the changed range.</summary>
    event EventHandler<CellLocation>? CellsChanged;

    /// <summary>Values may have changed (SheetChange, SheetCalculate, AfterCalculate). Cached evaluations must be dropped.</summary>
    event EventHandler? ValuesInvalidated;

    /// <summary>
    /// The active Excel workbook window changed (WindowActivate) to a DIFFERENT HWND than last reported. Argument = its HWND.
    /// Raised even during gateway-initiated navigation (a cross-workbook Goto must re-own the explorer window).
    /// </summary>
    event EventHandler<IntPtr>? ActiveWindowChanged;

    // NOTE: Excel raises its events synchronously INSIDE object-model calls (Goto, Visible=, Hidden=, Calculate...).
    // Consumers must not call back into the gateway from these event handlers; post the work instead.
}
