using System.Runtime.InteropServices;

namespace FormulaExplorer.AddIn.Excel;

/// <summary>
/// Decides whether a workbook is still open from two lookups the caller supplies (so the logic is testable without Excel).
/// A failed <c>Workbooks[name]</c> is NOT proof that the workbook is gone: Excel also fails it with a generic 0x800A03EC, between
/// two dialogs, or while a new workbook is being renamed by Save As. The workbook is therefore <see cref="WorkbookPresence.Gone"/>
/// only when a complete walk over <c>Application.Workbooks</c> does not find its name; any failure that leaves the question
/// open is <see cref="WorkbookPresence.Busy"/> ("cannot tell"), which a close tracker never mistakes for a finished close.
/// </summary>
internal static class WorkbookProbe
{
    /// <param name="workbookName">The name to look for.</param>
    /// <param name="lookupByName">Performs <c>Workbooks[name]</c> and releases what it got; throws when Excel cannot find it or cannot answer.</param>
    /// <param name="listNames">The names of every open workbook, from a complete walk; throws when Excel cannot answer.</param>
    public static WorkbookPresence Probe(string workbookName, Action lookupByName, Func<IReadOnlyList<string>> listNames)
    {
        try
        {
            lookupByName();
            return WorkbookPresence.Open;
        }
        catch (COMException ex) when (ComErrors.IsBusy(ex))
        {
            return WorkbookPresence.Busy;
        }
        catch (Exception ex) when (ex is COMException or ArgumentException or InvalidCastException)
        {
            // Not found, or a failure of another kind: only the walk can tell which.
        }

        try
        {
            foreach (var name in listNames())
                if (string.Equals(name, workbookName, StringComparison.OrdinalIgnoreCase)) return WorkbookPresence.Open;
            return WorkbookPresence.Gone;
        }
        catch (Exception ex) when (ex is COMException or ArgumentException or InvalidCastException)
        {
            return WorkbookPresence.Busy;   // the walk did not complete: it proves nothing
        }
    }
}
