namespace FormulaExplorer.AddIn.Excel;

/// <summary>What a probe of one workbook found.</summary>
public enum WorkbookPresence
{
    /// <summary>Still in <c>Application.Workbooks</c> and Excel answers.</summary>
    Open,

    /// <summary>Not in <c>Application.Workbooks</c> any more: the close went through.</summary>
    Gone,

    /// <summary>Excel cannot answer (modal dialog such as the "Save changes?" prompt) or the workbook's windows are disabled by one.</summary>
    Busy,
}
