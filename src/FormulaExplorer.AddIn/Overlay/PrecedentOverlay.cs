using System.Globalization;
using System.Runtime.InteropServices;
using FormulaExplorer.AddIn.Diagnostics;
using FormulaExplorer.AddIn.Excel;
using FormulaExplorer.AddIn.Session;
using XL = Microsoft.Office.Interop.Excel;

namespace FormulaExplorer.AddIn.Overlay;

/// <summary>
/// Win32 layered, click-through window that draws precedent borders over Excel's grid (Range-Finder style:
/// a colored outline plus small square handles on the four corners). Nothing in the workbook is touched.
/// <para>
/// While boxes are shown a 50 ms timer (<see cref="ViewTracker"/>) polls Excel's view. The overlay is visible only
/// while the active window shows the root's workbook and sheet and Excel is neither minimized nor busy; it is recomputed
/// when the view signature changes and at least every 500 ms. Everything runs on Excel's main thread; every entry point
/// catches and logs.
/// </para>
/// </summary>
public sealed class PrecedentOverlay : IPrecedentOverlay
{
    private readonly XL.Application _app;
    private readonly IGridMetrics _metrics;
    private readonly ExcelViewReader _reader;
    private readonly ViewChangeDetector _detector = new();
    private readonly ViewTracker _tracker;

    private readonly OverlayWindowKeeper _windows = new();
    private CellLocation? _root;
    private PrecedentBox[] _boxes = [];
    private IReadOnlyList<OverlayRect> _drawn = Array.Empty<OverlayRect>();
    private bool _disposed;

    // Rate limiting for repeated identical failures (the tick runs 20 times per second).
    private string? _lastFailure;
    private long _lastFailureMs;

    public PrecedentOverlay(XL.Application application)
        : this(application, new PointsToScreenPixelsMetrics())
    {
    }

    /// <summary>Lets the integration phase swap the coordinate conversion (plan B) without touching drawing code.</summary>
    internal PrecedentOverlay(XL.Application application, IGridMetrics metrics)
    {
        _app = application ?? throw new ArgumentNullException(nameof(application));
        _metrics = metrics;
        _reader = new ExcelViewReader(application);
        _tracker = new ViewTracker(Tick);
    }

    /// <inheritdoc />
    public void SetExplorerWindow(IntPtr explorerHwnd)
    {
        // The overlay sits directly above its owner XLMAIN, so every window the explorer shares that owner with is above
        // it already; there is nothing to track, only a z-order that may have been reshuffled by the re-ownership.
        try
        {
            if (_windows.Current is { IsVisible: true } window)
            {
                using (NativeMethods.DpiScope.PerMonitorV2())
                    window.EnsureAboveOwner();
            }
        }
        catch (Exception ex)
        {
            Log.Error("PrecedentOverlay.SetExplorerWindow", ex);
        }
    }

    /// <inheritdoc />
    public void Show(CellLocation root, IReadOnlyList<PrecedentBox> boxes)
    {
        if (_disposed) return;
        try
        {
            if (boxes.Count == 0)
            {
                Clear();
                return;
            }

            _root = root;
            _boxes = boxes.ToArray();
            _detector.Reset();
            _tracker.Start();
            Tick();
        }
        catch (Exception ex)
        {
            Log.Error("PrecedentOverlay.Show", ex);
        }
    }

    /// <inheritdoc />
    public void Clear()
    {
        try
        {
            _tracker.Stop();
            _root = null;
            _boxes = [];
            _detector.Reset();
            HideWindow();
        }
        catch (Exception ex)
        {
            Log.Error("PrecedentOverlay.Clear", ex);
        }
    }

    /// <summary>
    /// The range boundary rectangles of the visible box areas, one per pane they appear in, clipped to the pane, in
    /// physical screen pixels (the outline itself straddles these rectangles). Empty when hidden.
    /// </summary>
    public IReadOnlyList<OverlayRect> GetDrawnRects() => _drawn;

    /// <summary>Native handle of the layer window (0 when none exists); for diagnostics and integration tests.</summary>
    internal IntPtr WindowHandle => _windows.Current is { IsAlive: true } window ? window.Handle : IntPtr.Zero;

    /// <summary>True while the layer window is shown; for diagnostics and integration tests.</summary>
    internal bool WindowVisible => _windows.Current is { IsAlive: true, IsVisible: true };

    /// <inheritdoc />
    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        try
        {
            _tracker.Dispose();
            _boxes = [];
            _root = null;
            _drawn = Array.Empty<OverlayRect>();
            using (NativeMethods.DpiScope.PerMonitorV2())
                _windows.Dispose();
        }
        catch (Exception ex)
        {
            Log.Error("PrecedentOverlay.Dispose", ex);
        }
    }

    // ------------------------------------------------------------------------------------------------------------

    private void Tick()
    {
        if (_disposed || _root is null || _boxes.Length == 0) return;
        try
        {
            // Work in physical pixels no matter how Excel's thread is configured: window rects, window creation and
            // the DPI query all follow this context.
            using (NativeMethods.DpiScope.PerMonitorV2())
                Refresh(_root);
        }
        catch (Exception ex)
        {
            // Excel busy (cell edit mode, modal dialog) or a transient COM failure: hide until it is usable again.
            HideWindow();
            _detector.Reset();   // redraw at the first tick that works, not at the next forced refresh
            if (!ExcelViewReader.IsBusy(ex)) LogFailure("PrecedentOverlay tick", ex);
        }
    }

    private void Refresh(CellLocation root)
    {
        var reading = _reader.Read(root);
        if (!_detector.ShouldRefresh(reading.Signature, ViewTracker.NowMs()))
        {
            // Nothing moved; only keep the layer directly above its owner (Excel's Goto and activation can reshuffle).
            _windows.Current?.EnsureAboveOwner();
            return;
        }

        if (reading.Signature is null || reading.Window is null || reading.Sheet is null)
        {
            HideWindow();
            return;
        }

        var window = _windows.Ensure(reading.XlMain);
        var gridRect = reading.Signature.GridRect;
        window.Place(gridRect, window.AboveOwnerAnchor());

        var dpiScale = window.Dpi / 96.0;
        var panes = _metrics.GetPanes(reading.Window, gridRect, reading.Signature.GridDpi);
        var boxAreas = ReadBoxAreas(reading.Sheet, root);
        var frame = GridGeometry.Compute(panes, gridRect, boxAreas, dpiScale);

        if (frame.Fills.Count == 0)
        {
            HideWindow();
            return;
        }

        window.SetContent(frame.Fills);
        window.Show();
        window.EnsureAboveOwner();
        _drawn = frame.Bounds;
    }

    private void HideWindow()
    {
        _windows.Current?.Hide();
        _drawn = Array.Empty<OverlayRect>();
    }

    /// <summary>Resolves every box to its areas (points) on the root's sheet. Boxes that no longer resolve are skipped.</summary>
    private List<BoxAreas> ReadBoxAreas(XL.Worksheet sheet, CellLocation root)
    {
        var result = new List<BoxAreas>(_boxes.Length);
        foreach (var box in _boxes)
        {
            var location = box.Location;
            if (location.WorkbookName != root.WorkbookName || location.SheetName != root.SheetName) continue;

            try
            {
                // Worksheet.Range does not take the ',' that Address uses between areas on every locale (pt-BR wants ';'):
                // resolve one area at a time.
                var areas = new List<PointRect>();
                foreach (var part in AreaAddress.Split(location.Address))
                    areas.AddRange(ReadAreas(sheet.get_Range(part)));
                if (areas.Count > 0) result.Add(new BoxAreas(OverlayWindow.SafeRgb(box.Rgb), areas));
            }
            catch (Exception ex) when (!ExcelViewReader.IsBusy(ex))
            {
                // A stale or malformed address (sheet changed under us): skip this box, keep the others.
                LogFailure($"PrecedentOverlay: cannot resolve {location}", ex);
            }
        }

        return result;
    }

    private static List<PointRect> ReadAreas(XL.Range range)
    {
        var areas = range.Areas;
        var count = areas.Count;
        var result = new List<PointRect>(count);
        for (var i = 1; i <= count; i++)
        {
            var area = areas[i];
            result.Add(new PointRect(
                ToDouble(area.Left), ToDouble(area.Top), ToDouble(area.Width), ToDouble(area.Height)));
        }

        return result;
    }

    private static double ToDouble(object value) => Convert.ToDouble(value, CultureInfo.InvariantCulture);

    private void LogFailure(string context, Exception ex)
    {
        var key = $"{context}|{ex.GetType().Name}|{(ex as COMException)?.HResult}";
        var now = ViewTracker.NowMs();
        if (key == _lastFailure && now - _lastFailureMs < 10_000) return;
        _lastFailure = key;
        _lastFailureMs = now;
        Log.Error(context, ex);
    }
}
