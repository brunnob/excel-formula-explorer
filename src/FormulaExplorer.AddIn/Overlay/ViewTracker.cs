using System.Diagnostics;
using System.Windows.Threading;
using FormulaExplorer.AddIn.Diagnostics;

namespace FormulaExplorer.AddIn.Overlay;

/// <summary>
/// Everything that changes where the grid draws a given cell. Excel raises no event for scrolling or zooming, so the
/// tracker polls this and redraws when it differs from the previous one. Value equality (record) is what compares them.
/// </summary>
/// <param name="ActiveHwnd">Window.Hwnd of the active window.</param>
/// <param name="SheetName">Active sheet.</param>
/// <param name="ScrollRow">Window.ScrollRow (active pane).</param>
/// <param name="ScrollColumn">Window.ScrollColumn (active pane).</param>
/// <param name="Zoom">Window.Zoom in percent.</param>
/// <param name="SplitRow">Window.SplitRow.</param>
/// <param name="SplitColumn">Window.SplitColumn.</param>
/// <param name="FreezePanes">Window.FreezePanes.</param>
/// <param name="PaneCount">Window.Panes.Count (1, 2 or 4).</param>
/// <param name="PaneScroll">Per pane: ScrollRow/ScrollColumn (the wheel can scroll a pane that is not the active one) and the visible range address and size (hidden/unhidden/resized rows and columns).</param>
/// <param name="GridRect">GetWindowRect of the EXCEL7 window (moves/resizes of the Excel window).</param>
/// <param name="GridDpi">GetDpiForWindow of the EXCEL7 window (the window moved to another monitor / scale changed).</param>
internal sealed record ViewSignature(
    long ActiveHwnd,
    string SheetName,
    int ScrollRow,
    int ScrollColumn,
    int Zoom,
    int SplitRow,
    int SplitColumn,
    bool FreezePanes,
    int PaneCount,
    string PaneScroll,
    PixelRect GridRect,
    int GridDpi);

/// <summary>
/// Decides when the overlay must recompute: whenever the view signature changed, and at least every
/// <see cref="ForcedRefreshMs"/> ms (row heights, hidden rows, cell edits and window state have no cheap signal).
/// A <c>null</c> signature stands for "overlay must be hidden" (wrong window/sheet, minimized, Excel busy), so
/// transitions into and out of that state are changes like any other.
/// </summary>
internal sealed class ViewChangeDetector
{
    public const int ForcedRefreshMs = 500;

    private bool _hasLast;
    private ViewSignature? _last;
    private long _lastRefreshMs;

    /// <summary>Forgets the previous signature so the next call returns true.</summary>
    public void Reset() => _hasLast = false;

    /// <summary>True when the caller must refresh now; records <paramref name="signature"/> as the new baseline when so.</summary>
    public bool ShouldRefresh(ViewSignature? signature, long nowMs)
    {
        var refresh = !_hasLast
                      || !Equals(signature, _last)
                      || nowMs - _lastRefreshMs >= ForcedRefreshMs;
        if (!refresh) return false;

        _hasLast = true;
        _last = signature;
        _lastRefreshMs = nowMs;
        return true;
    }
}

/// <summary>
/// The 50 ms <see cref="DispatcherTimer"/> that drives the overlay while borders are shown. It only invokes
/// <c>tick</c>; the decision logic lives in <see cref="ViewChangeDetector"/>. Exceptions never escape a tick, and a tick
/// is not re-entered if Excel pumps messages while the previous one is still running.
/// </summary>
internal sealed class ViewTracker : IDisposable
{
    public const int IntervalMs = 50;

    private readonly DispatcherTimer _timer;
    private readonly Action _tick;
    private bool _inTick;

    public ViewTracker(Action tick, Dispatcher? dispatcher = null)
    {
        _tick = tick;
        _timer = new DispatcherTimer(DispatcherPriority.Normal, dispatcher ?? Dispatcher.CurrentDispatcher)
        {
            Interval = TimeSpan.FromMilliseconds(IntervalMs),
        };
        _timer.Tick += OnTick;
    }

    public bool IsRunning => _timer.IsEnabled;

    public void Start() => _timer.Start();

    public void Stop() => _timer.Stop();

    public void Dispose()
    {
        _timer.Stop();
        _timer.Tick -= OnTick;
    }

    private void OnTick(object? sender, EventArgs e)
    {
        if (_inTick) return;
        _inTick = true;
        try
        {
            _tick();
        }
        catch (Exception ex)
        {
            Log.Error("Overlay view tracker tick", ex);
        }
        finally
        {
            _inTick = false;
        }
    }

    /// <summary>Monotonic milliseconds for <see cref="ViewChangeDetector"/>.</summary>
    public static long NowMs() => Stopwatch.GetTimestamp() * 1000 / Stopwatch.Frequency;
}
