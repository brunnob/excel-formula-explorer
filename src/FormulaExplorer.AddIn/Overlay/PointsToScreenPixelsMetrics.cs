using System.Globalization;
using System.Runtime.InteropServices;
using XL = Microsoft.Office.Interop.Excel;

namespace FormulaExplorer.AddIn.Overlay;

/// <summary>
/// Default <see cref="IGridMetrics"/>: Excel's own <c>Pane.PointsToScreenPixelsX/Y</c>.
/// <para>
/// Assumptions (verified by the real-Excel integration tests):
/// <list type="bullet">
/// <item>The conversion maps document coordinates (points from the sheet origin, i.e. <c>Range.Left/Top</c>) to
/// physical screen pixels and already includes the pane's scroll position and the window zoom.</item>
/// <item>The result is in physical pixels like <c>GetWindowRect</c> under per-monitor-v2 DPI awareness (callers run
/// under that context).</item>
/// <item>Each pane converts relative to its own scroll offset, so a range is converted once per pane and clipped to
/// that pane's visible rectangle.</item>
/// <item>The pane's visible rectangle comes from <c>Pane.VisibleRange</c> (partly visible rows/columns count as visible).</item>
/// </list>
/// If any of this is wrong, replace this class (the two conversion functions are <c>ToPixelsX</c>/<c>ToPixelsY</c> of the pane wrapper).
/// <para>
/// <c>PointsToScreenPixelsX/Y</c> take whole points, so fractional coordinates are interpolated
/// (<see cref="PointConversion"/>).
/// </para>
/// </para>
/// </summary>
internal sealed class PointsToScreenPixelsMetrics : IGridMetrics
{
    public IReadOnlyList<IPaneMetrics> GetPanes(XL.Window window, PixelRect gridRect, int dpi)
    {
        var panes = window.Panes;
        var count = panes.Count;
        var raw = new List<PanePoints>(count);
        for (var i = 1; i <= count; i++) raw.Add(new PanePoints(panes[i]));
        if (raw.Count == 0) return raw;

        // VisibleRange includes the partly visible last row/column, which reaches over the scroll bars and sheet tabs
        // (inside the EXCEL7 rect) and, in a split window, over the neighbouring pane: clip to the cell area and to the
        // other panes. Window.UsableWidth/UsableHeight are the window's cell-area size in points (to be confirmed by the
        // integration screenshot; PaneClipping.CellArea falls back to the EXCEL7 rect when they look wrong).
        var cellArea = PaneClipping.CellArea(
            gridRect, raw[0].VisibleRect, ReadPoints(() => window.UsableWidth), ReadPoints(() => window.UsableHeight), dpi);
        var clipped = PaneClipping.Clip(raw.Select(p => p.VisibleRect).ToList(), cellArea);
        var result = new List<IPaneMetrics>(count);
        for (var i = 0; i < count; i++) result.Add(raw[i].WithVisibleRect(clipped[i]));
        return result;
    }

    private static double ReadPoints(Func<double> read)
    {
        try
        {
            return read();
        }
        catch (COMException ex) when (!ExcelViewReader.IsBusy(ex))
        {
            return 0;   // not available: CellArea falls back to the grid window
        }
    }

    private sealed class PanePoints : IPaneMetrics
    {
        private readonly XL.Pane _pane;

        public PanePoints(XL.Pane pane)
        {
            _pane = pane;
            VisibleRect = ToScreen(ReadVisibleRange(pane));
        }

        private PanePoints(XL.Pane pane, PixelRect visibleRect)
        {
            _pane = pane;
            VisibleRect = visibleRect;
        }

        public PixelRect VisibleRect { get; }

        public PanePoints WithVisibleRect(PixelRect visibleRect) => new(_pane, visibleRect);

        public PixelRect ToScreen(PointRect area) => new(
            ToPixelsX(area.Left),
            ToPixelsY(area.Top),
            ToPixelsX(area.Left + area.Width),
            ToPixelsY(area.Top + area.Height));

        private int ToPixelsX(double points) => PointConversion.ToPixels(points, p => _pane.PointsToScreenPixelsX(p));

        private int ToPixelsY(double points) => PointConversion.ToPixels(points, p => _pane.PointsToScreenPixelsY(p));

        private static PointRect ReadVisibleRange(XL.Pane pane)
        {
            var range = pane.VisibleRange;
            return new PointRect(
                Convert.ToDouble(range.Left, CultureInfo.InvariantCulture),
                Convert.ToDouble(range.Top, CultureInfo.InvariantCulture),
                Convert.ToDouble(range.Width, CultureInfo.InvariantCulture),
                Convert.ToDouble(range.Height, CultureInfo.InvariantCulture));
        }
    }
}
