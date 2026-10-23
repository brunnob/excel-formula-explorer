using System.Windows;

namespace FormulaExplorer.AddIn.UI;

/// <summary>What a left-button press on a link node means.</summary>
internal enum PressAction
{
    /// <summary>Raise NodeClicked.</summary>
    Click,

    /// <summary>Raise NodeDoubleClicked (the first press of the pair already raised NodeClicked).</summary>
    DoubleClick,

    /// <summary>A double click whose first press never raised NodeClicked for this node: raise both, click first.</summary>
    ClickThenDoubleClick,

    /// <summary>
    /// The second press of a double click whose first press already replaced the text (a click drills down into the
    /// precedent's formula): ignore it, so a habitual double click does not drill two levels.
    /// </summary>
    Ignore,
}

/// <summary>
/// Our own double-click detection. WPF's <c>ClickCount</c> is unreliable here: the first click's navigation can make
/// Excel re-activate its own window, which resets the click sequence. A press is the second half of a double click when
/// it hits the same node as the previous link press within the system double-click time and rectangle; WPF's own
/// <c>ClickCount == 2</c> is accepted as well. Presses are timed by their input timestamp (when they happened), not by
/// when they are dispatched: the first click's navigation can keep the second press queued for a long time.
/// </summary>
internal sealed class DoubleClickDetector
{
    private readonly Func<long> _clockMs;
    private readonly Func<double> _intervalMs;
    private readonly Func<(double Width, double Height)> _slop;

    private object? _node;
    private Point _point;
    private int _timeMs;
    private bool _hasPress;   // _point/_timeMs describe the last link press
    private bool _replaced;   // the text was replaced since that press

    /// <param name="clockMs">Monotonic milliseconds.</param>
    /// <param name="intervalMs">Maximum time between the two presses.</param>
    /// <param name="slop">System double-click rectangle (centered on the first press), in the same units as the points.</param>
    public DoubleClickDetector(Func<long> clockMs, Func<double> intervalMs, Func<(double Width, double Height)> slop)
    {
        _clockMs = clockMs;
        _intervalMs = intervalMs;
        _slop = slop;
    }

    /// <summary>Detector wired to the system settings (<c>GetDoubleClickTime</c>, <c>SM_CXDOUBLECLK</c>/<c>SM_CYDOUBLECLK</c>).</summary>
    /// <param name="pixelsToUnits">Converts physical pixels to the unit of the points that will be passed in (DIPs).</param>
    public static DoubleClickDetector CreateForSystem(Func<double> pixelsToUnits) => new(
        () => Environment.TickCount64,
        () => NativeMethods.GetDoubleClickTime(),
        () =>
        {
            var (w, h) = NativeMethods.GetDoubleClickSize();
            double scale = pixelsToUnits();
            return (w * scale, h * scale);
        });

    /// <summary>
    /// Registers a left-button press on a link node and says what to raise. A press that is not on a link must call
    /// <see cref="Reset"/> instead.
    /// </summary>
    /// <param name="node">Identity of the node (reference equality).</param>
    /// <param name="clickCount">WPF's click count for the press.</param>
    /// <param name="timestampMs">
    /// When the press happened, in the <c>Environment.TickCount</c> domain (<c>MouseEventArgs.Timestamp</c>); null for a
    /// simulated press, which is timed by the injected clock when it is processed.
    /// </param>
    public PressAction OnLinkPress(object node, Point point, int clickCount, int? timestampMs = null)
    {
        int now = timestampMs ?? unchecked((int)_clockMs());
        int elapsed = unchecked(now - _timeMs); // wraps correctly across the 49.7-day TickCount rollover
        var (slopWidth, slopHeight) = _slop();
        bool close = _hasPress
            && elapsed >= 0 && elapsed <= _intervalMs()
            && Math.Abs(point.X - _point.X) <= slopWidth / 2
            && Math.Abs(point.Y - _point.Y) <= slopHeight / 2;

        if (_replaced)
        {
            _replaced = false;
            if (close || clickCount == 2)
            {
                Reset();
                return PressAction.Ignore;
            }
        }

        bool sameNode = ReferenceEquals(_node, node);
        bool second = sameNode && close;

        if (sameNode && (second || clickCount == 2))
        {
            Reset();
            return PressAction.DoubleClick;
        }
        if (clickCount == 2)
        {
            Reset();
            return PressAction.ClickThenDoubleClick;
        }

        _node = node;
        _point = point;
        _timeMs = now;
        _hasPress = true;
        return PressAction.Click;
    }

    /// <summary>
    /// The text was replaced (new root or loading). The previous press can no longer pair with a node, but a press
    /// that would have completed a double click with it is swallowed (<see cref="PressAction.Ignore"/>).
    /// </summary>
    public void ContentReplaced()
    {
        _node = null;
        _replaced = _hasPress;
    }

    /// <summary>Forgets the previous press (press elsewhere, window closing).</summary>
    public void Reset()
    {
        _node = null;
        _hasPress = false;
        _replaced = false;
    }
}
