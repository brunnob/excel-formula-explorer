using System.Runtime.InteropServices;
using FormulaExplorer.Core.Values;

namespace FormulaExplorer.AddIn.Excel;

/// <summary>
/// COM-free identity of a cell or range. Always built by the gateway from Excel's own names, so plain
/// (ordinal) record equality is reliable.
/// </summary>
/// <param name="WorkbookName">Workbook.Name, e.g. "Book1.xlsx".</param>
/// <param name="SheetName">Worksheet.Name.</param>
/// <param name="Address">Range.Address(true, true) in A1 style, e.g. "$K$31", "$I$19:$J$19", "$A$1:$B$2,$D$4".</param>
public sealed record CellLocation(string WorkbookName, string SheetName, string Address)
{
    public override string ToString() => $"[{WorkbookName}]{SheetName}!{Address}";
}

/// <summary>What a navigable node points to.</summary>
/// <param name="Location">The whole target range.</param>
/// <param name="FirstCell">The first (top-left) cell of the first area.</param>
/// <param name="IsSingleCell">True when the target is exactly one cell.</param>
/// <param name="FirstCellHasFormula">Range.HasFormula of <paramref name="FirstCell"/>.</param>
public sealed record RangeTarget(CellLocation Location, CellLocation FirstCell, bool IsSingleCell, bool FirstCellHasFormula);

/// <summary>A cell whose content can be displayed in the explorer.</summary>
/// <param name="Location">Single-cell location.</param>
/// <param name="DisplayText">Range.Formula of the cell ("" for an empty cell). For a constant, the constant as Range.Formula returns it.</param>
/// <param name="HasFormula">Range.HasFormula.</param>
/// <param name="Value">The cell's current Value2.</param>
public sealed record RootCell(CellLocation Location, string DisplayText, bool HasFormula, EvalResult Value);

/// <summary>Result of evaluating a fragment.</summary>
/// <param name="Result">Value model (a Range result is <see cref="EvalResult.RangeValue"/>).</param>
/// <param name="Target">Non-null exactly when Excel returned a Range object.</param>
public sealed record NodeEvaluation(EvalResult Result, RangeTarget? Target);

public enum NavigationStatus
{
    Ok,
    /// <summary>The target is on a hidden sheet and the workbook structure is protected. Nothing was changed.</summary>
    Blocked,
    /// <summary>The workbook/sheet/range no longer exists, or Excel refused.</summary>
    Failed,
}

/// <param name="Message">English message to show the user (only for Blocked/Failed when worth telling).</param>
public sealed record NavigationResult(NavigationStatus Status, string? Message)
{
    public static NavigationResult Ok { get; } = new(NavigationStatus.Ok, null);
}

/// <summary>
/// Thrown by gateway members when Excel cannot take object-model calls right now
/// (VBA_E_IGNORE 0x800AC472 when Excel is not ready, RPC_E_CALL_REJECTED 0x80010001 behind a modal dialog, or the explicit
/// <see cref="EditModeException"/> raised by <see cref="ComErrors.ThrowIfEditMode"/> while a cell is edited).
/// Callers should retry later.
/// </summary>
public sealed class ExcelBusyException : Exception
{
    public ExcelBusyException(Exception inner) : base("Excel is busy.", inner) { }

    /// <summary>True when the cause is cell edit mode (raised explicitly by ComErrors.ThrowIfEditMode, never inferred from a bare VBA_E_IGNORE): it lasts as long as the user keeps typing, not a moment.</summary>
    public bool IsEditMode => InnerException is EditModeException;

    /// <summary>
    /// True when the cause is a bare VBA_E_IGNORE that Excel raised by itself ("not ready right now": an edit-like state the
    /// edit-mode check does not see, a transient state between commands). It is NOT known to end soon, so a user command
    /// is retried only a couple of times instead of the full budget (see <c>ExplorerSession</c>).
    /// </summary>
    public bool IsNotReady => InnerException is COMException { HResult: ComErrors.VbaIgnore } && !IsEditMode;
}
