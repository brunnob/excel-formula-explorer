using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using FormulaExplorer.AddIn.Excel;
using static FormulaExplorer.AddIn.Overlay.NativeMethods;
using XL = Microsoft.Office.Interop.Excel;

namespace FormulaExplorer.AddIn.Overlay;

/// <summary>
/// Outcome of looking at Excel's current view.
/// <see cref="Signature"/> is null when the overlay must not be visible; otherwise the COM objects the geometry needs.
/// </summary>
internal readonly record struct ViewReading(
    ViewSignature? Signature,
    IntPtr XlMain,
    XL.Window? Window,
    XL.Worksheet? Sheet)
{
    public static ViewReading Hidden => default;
}

/// <summary>
/// Reads the part of Excel's state that decides whether and where the overlay is drawn. Runs on every tick, so it only
/// does cheap property reads. COM failures (Excel busy in cell-edit mode, modal dialog) propagate as exceptions and the
/// caller treats them as "hide".
/// </summary>
internal sealed class ExcelViewReader
{
    private readonly XL.Application _app;

    public ExcelViewReader(XL.Application app) => _app = app;

    /// <summary>
    /// Visible only when the active window shows the root's workbook and sheet, in a normal/page-break view, is not
    /// minimized or hidden, and its grid window (EXCEL7) exists.
    /// </summary>
    public ViewReading Read(CellLocation root)
    {
        var window = _app.ActiveWindow;
        if (window is null) return ViewReading.Hidden;

        var workbook = _app.ActiveWorkbook;
        if (workbook is null || !string.Equals(workbook.Name, root.WorkbookName, StringComparison.Ordinal))
            return ViewReading.Hidden;

        if (window.ActiveSheet is not XL.Worksheet sheet
            || !string.Equals(sheet.Name, root.SheetName, StringComparison.Ordinal))
            return ViewReading.Hidden;

        // Page Layout view draws pages with margins; Pane.PointsToScreenPixels is not meaningful there.
        if (window.View == XL.XlWindowView.xlPageLayoutView) return ViewReading.Hidden;

        // In SDI Excel (2013+) Window.Hwnd is the workbook window's XLMAIN. Walk to the top-level window to be safe.
        var hwnd = new IntPtr(window.Hwnd);
        var xlMain = GetAncestor(hwnd, GA_ROOT);
        if (xlMain == IntPtr.Zero) xlMain = hwnd;
        // Iconic: the overlay is owned by XLMAIN, so it would minimize with it anyway; hiding keeps our state honest.
        if (IsIconic(xlMain) || !IsWindowVisible(xlMain)) return ViewReading.Hidden;

        var grid = FindGridWindow(xlMain);
        if (grid == IntPtr.Zero || !GetWindowRect(grid, out var rect)) return ViewReading.Hidden;

        var panes = window.Panes;
        var paneCount = panes.Count;
        var paneScroll = ReadPaneLayout(panes, paneCount);

        var signature = new ViewSignature(
            hwnd.ToInt64(),
            sheet.Name,
            window.ScrollRow,
            window.ScrollColumn,
            Convert.ToInt32(window.Zoom, CultureInfo.InvariantCulture),
            window.SplitRow,
            window.SplitColumn,
            window.FreezePanes,
            paneCount,
            paneScroll,
            rect.ToPixelRect(),
            (int)GetDpiForWindow(grid));
        return new ViewReading(signature, xlMain, window, sheet);
    }

    /// <summary>
    /// Per pane: scroll position plus the visible range's address and size in points. The address/size part changes
    /// as soon as rows or columns inside the view are hidden, unhidden (e.g. a navigation expanding an outline row) or
    /// resized, so the borders move at once instead of at the next forced refresh.
    /// </summary>
    private static string ReadPaneLayout(XL.Panes panes, int count)
    {
        var sb = new StringBuilder();
        for (var i = 1; i <= count; i++)
        {
            var pane = panes[i];
            var visible = pane.VisibleRange;
            sb.Append(pane.ScrollRow).Append(',').Append(pane.ScrollColumn).Append(',')
              .Append(visible.get_Address(false, false, XL.XlReferenceStyle.xlA1)).Append(',')
              .Append(Convert.ToDouble(visible.Width, CultureInfo.InvariantCulture).ToString("R", CultureInfo.InvariantCulture)).Append(',')
              .Append(Convert.ToDouble(visible.Height, CultureInfo.InvariantCulture).ToString("R", CultureInfo.InvariantCulture)).Append(',');
            // Where the first visible cell is drawn: smooth (pixel) scrolling by touchpad or by dragging the scroll bar can
            // move the view without changing ScrollRow/ScrollColumn or the visible range.
            int left = (int)Math.Floor(Convert.ToDouble(visible.Left, CultureInfo.InvariantCulture));
            int top = (int)Math.Floor(Convert.ToDouble(visible.Top, CultureInfo.InvariantCulture));
            sb.Append(pane.PointsToScreenPixelsX(left)).Append(',').Append(pane.PointsToScreenPixelsY(top)).Append(';');
        }

        return sb.ToString();
    }

    /// <summary>
    /// True for the two "Excel cannot take object-model calls right now" failures (cell edit mode, modal dialog or
    /// call-rejected) plus RPC "retry later"; expected, not worth logging.
    /// </summary>
    public static bool IsBusy(Exception ex) =>
        ex is COMException { HResult: unchecked((int)0x800AC472) or unchecked((int)0x80010001) or unchecked((int)0x8001010A) };
}
