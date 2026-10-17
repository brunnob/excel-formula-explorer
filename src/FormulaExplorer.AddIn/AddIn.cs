using ExcelDna.Integration;
using FormulaExplorer.AddIn.Composition;
using FormulaExplorer.AddIn.Diagnostics;
using FormulaExplorer.AddIn.Testing;

namespace FormulaExplorer.AddIn;

/// <summary>Excel-DNA entry point: add-in load/unload. The UI is started from the ribbon (<see cref="Ribbon.StudioRibbon"/>).</summary>
public sealed class FormulaExplorerAddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        try
        {
            var assembly = typeof(FormulaExplorerAddIn).Assembly.GetName().Version;
            var excelDna = typeof(ExcelDnaUtil).Assembly.GetName().Version;
            Log.Info($"Formula Explorer {assembly} loaded (Excel-DNA {excelDna}, .NET {Environment.Version}, Excel {ExcelDnaUtil.ExcelVersion}, {ExcelDnaUtil.XllPath})");

            TestHooks.RegisterIfEnabled();
        }
        catch (Exception ex)
        {
            Log.Error("FormulaExplorerAddIn.AutoOpen", ex);
        }
    }

    public void AutoClose()
    {
        try
        {
            ExplorerController.CloseAll();
            Log.Info("Formula Explorer unloaded.");
        }
        catch (Exception ex)
        {
            Log.Error("FormulaExplorerAddIn.AutoClose", ex);
        }
    }
}
