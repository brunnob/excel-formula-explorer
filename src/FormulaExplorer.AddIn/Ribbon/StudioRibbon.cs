using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using FormulaExplorer.AddIn.Composition;
using FormulaExplorer.AddIn.Diagnostics;

namespace FormulaExplorer.AddIn.Ribbon;

/// <summary>The "Studio" tab (after View) with the "Navigation" group and the large "Formula Explorer" button.</summary>
[ComVisible(true)]
public class StudioRibbon : ExcelRibbon
{
    /// <summary>Built-in icon used only when the embedded globe cannot be loaded.</summary>
    private const string FallbackImageMso = "HyperlinkInsert";

    /// <summary>Size of the large button's image at 96 DPI.</summary>
    private const int LargeImagePixels = 32;

    public override string GetCustomUI(string ribbonId)
    {
        try
        {
            // getImage is chosen only if the PNG really loads; otherwise a stock icon keeps the button usable.
            var image = RibbonIcon.GetPicture(LargeImagePixels) != null
                ? "getImage=\"GetButtonImage\""
                : $"imageMso=\"{FallbackImageMso}\"";

            return $"""
                <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
                  <ribbon>
                    <tabs>
                      <tab id="StudioTab" label="Studio" insertAfterMso="TabView">
                        <group id="NavigationGroup" label="Navigation">
                          <button id="FormulaExplorerButton" label="Formula Explorer" size="large"
                                  {image}
                                  screentip="Launch Formula Explorer"
                                  supertip="Use the Formula Explorer to easily navigate the most complex formulas."
                                  onAction="OnFormulaExplorer" />
                        </group>
                      </tab>
                    </tabs>
                  </ribbon>
                </customUI>
                """;
        }
        catch (Exception ex)
        {
            Log.Error("StudioRibbon.GetCustomUI", ex);
            return string.Empty;
        }
    }

    /// <summary>getImage callback: the globe, sized for the current DPI (Excel scales it to the button).</summary>
    public object? GetButtonImage(IRibbonControl control)
    {
        try
        {
            var scale = RibbonIcon.GetDpiScale(ExcelDnaUtil.WindowHandle);
            return RibbonIcon.GetPicture((int)Math.Ceiling(LargeImagePixels * scale));
        }
        catch (Exception ex)
        {
            Log.Error("StudioRibbon.GetButtonImage", ex);
            return null;
        }
    }

    /// <summary>onAction callback of the Formula Explorer button.</summary>
    public void OnFormulaExplorer(IRibbonControl control)
    {
        try
        {
            ExplorerController.Open();
        }
        catch (Exception ex)
        {
            // Open() reports its own failures; this is the last line of defence for anything unexpected.
            Log.Error("StudioRibbon.OnFormulaExplorer", ex);
            ExplorerController.ShowError($"Formula Explorer could not start: {ex.Message}");
        }
    }
}
