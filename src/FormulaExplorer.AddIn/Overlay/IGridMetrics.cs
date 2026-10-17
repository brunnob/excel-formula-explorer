using XL = Microsoft.Office.Interop.Excel;

namespace FormulaExplorer.AddIn.Overlay;

/// <summary>
/// One pane of the active workbook window (a window has 1, 2 or 4 panes when split or frozen) as the drawing code
/// sees it: where its visible area is on screen and where a worksheet rectangle lands in it.
/// </summary>
internal interface IPaneMetrics
{
    /// <summary>
    /// The area of the pane where cells are really drawn, in physical screen pixels: the partly visible last row/column
    /// stays included, but the rectangle is already clipped to the cell area (not over scroll bars or the sheet-tab strip)
    /// and to its neighbouring panes (see <see cref="PaneClipping"/>). The caller still intersects it with the grid window.
    /// </summary>
    PixelRect VisibleRect { get; }

    /// <summary>
    /// Maps a worksheet rectangle (points) to physical screen pixels as drawn in this pane (scroll and zoom included).
    /// The result is NOT clipped and may lie entirely outside the pane.
    /// </summary>
    PixelRect ToScreen(PointRect area);
}

/// <summary>
/// The seam that isolates every worksheet-to-screen coordinate conversion. The default implementation is
/// <see cref="PointsToScreenPixelsMetrics"/> (Excel's own <c>Pane.PointsToScreenPixelsX/Y</c>). A "plan B" (EXCEL7 rect +
/// VisibleRange + row heights/column widths x Zoom/100 x DPI/72) can replace it without touching any drawing code.
/// Implementations may throw <see cref="System.Runtime.InteropServices.COMException"/> when Excel is busy.
/// </summary>
internal interface IGridMetrics
{
    /// <summary>
    /// The panes of <paramref name="window"/> (Excel's active window), in <c>Window.Panes</c> order.
    /// </summary>
    /// <param name="window">Excel's active window.</param>
    /// <param name="gridRect">The EXCEL7 window rect in physical screen pixels.</param>
    /// <param name="dpi">DPI of the EXCEL7 window.</param>
    IReadOnlyList<IPaneMetrics> GetPanes(XL.Window window, PixelRect gridRect, int dpi);
}
