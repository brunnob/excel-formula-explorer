using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using FormulaExplorer.AddIn.Diagnostics;
using FormulaExplorer.AddIn.Session;
using FormulaExplorer.Core.Syntax;

namespace FormulaExplorer.AddIn.UI;

/// <summary>
/// The formula text: own line layout (no text box), per-node hover/gold highlights, link underline, hand cursor,
/// "Returns" tooltip and click / double-click detection. Calibri 11 pt, 20 DIP line pitch. All coordinates the
/// public members take or return are in the view's own DIP space; the text starts at <see cref="TextInset"/>.
/// </summary>
internal sealed class FormulaView : FrameworkElement
{
    /// <summary>
    /// Text origin inside the view: 3 px from the left/right and 3 px from the top, measured on the original's
    /// screenshots (the hover rectangle of the first node starts 3 px right of the box's inner edge).
    /// </summary>
    public static readonly Thickness TextInset = new(3, 3, 3, 0);

    internal const double FontSizeDip = 11 * 96.0 / 72.0;
    internal const int TooltipDelayMs = 500;

    /// <summary>Distance from a line's top to the text baseline (the ink of "I" ends 14 px below the hover rectangle's top).</summary>
    private const double BaselineOffset = 14;

    private static readonly Typeface FormulaTypeface = new(new FontFamily("Calibri, Segoe UI"), FontStyles.Normal, FontWeights.Normal, FontStretches.Normal);
    private static readonly Brush HoverBrush = Frozen(new SolidColorBrush(Color.FromRgb(0xDC, 0xDC, 0xDC)));
    private static readonly Brush GoldBrush = Frozen(new SolidColorBrush(Color.FromRgb(0xFF, 0xD7, 0x00)));
    private static readonly NodeInfo NoInfo = new(false, null);

    private readonly INodeInfoProvider _provider;
    private readonly DoubleClickDetector _doubleClicks;
    private readonly ReturnsToolTip _tooltip;
    private readonly DispatcherTimer _tooltipTimer;
    private readonly Dictionary<uint, Brush> _brushes = new();

    private FormulaDisplay? _display;
    private string _text = string.Empty;
    private bool _interactive;
    private bool _loading;
    private uint?[] _charColors = Array.Empty<uint?>();

    private TextLayout _layout = TextLayout.Empty;
    private FormattedTextMeasurer? _measurer;
    private FormattedText?[] _lineTexts = Array.Empty<FormattedText?>();
    private bool _layoutDirty = true;
    private double? _wrapWidth;
    private double _layoutPpd;
    private double _textTop;
    private double _underlineTop;

    private SyntaxNode? _hoverNode;
    private NodeInfo _hoverInfo = NoInfo;
    private SyntaxNode? _gold;
    private Point? _pointer;
    private SyntaxNode? _tooltipSuppressedFor;

    public FormulaView(INodeInfoProvider provider)
    {
        _provider = provider;
        _tooltip = new ReturnsToolTip(this);
        _doubleClicks = DoubleClickDetector.CreateForSystem(() => 1.0 / Math.Max(VisualTreeHelper.GetDpi(this).PixelsPerDip, 0.01));
        _tooltipTimer = new DispatcherTimer(DispatcherPriority.Normal) { Interval = TimeSpan.FromMilliseconds(TooltipDelayMs) };
        _tooltipTimer.Tick += (_, _) => Guard("FormulaView.TooltipTimer", OnTooltipTimer);

        SnapsToDevicePixels = true;
        UseLayoutRounding = true;
        Focusable = false;
        Cursor = Cursors.Arrow;
        TextOptions.SetTextFormattingMode(this, TextFormattingMode.Display);
    }

    public event EventHandler<SyntaxNode>? NodeClicked;

    public event EventHandler<SyntaxNode>? NodeDoubleClicked;

    /// <summary>Scroll host (set by the panel) used to tell which characters are currently visible.</summary>
    public ScrollViewer? ScrollHost { get; set; }

    // ----- State for snapshots and tests -----

    public string DisplayText => _display?.Formula.Text ?? string.Empty;

    public bool Interactive => _interactive;

    public bool IsLoading => _loading;

    public SyntaxNode? HoverNode => _hoverNode;

    /// <summary>True when the hovered node is shown as a link (underlined, hand cursor): not for the gold node.</summary>
    public bool HoverIsLink => _hoverNode is not null && !HoverIsGold && _hoverInfo.IsLink;

    public SyntaxNode? GoldNode => _gold;

    public bool TooltipVisible => _tooltip.IsOpen;

    public string? TooltipText => _tooltip.Text;

    /// <summary>HWND of the tooltip popup while it is open in a real window (zero otherwise).</summary>
    public IntPtr TooltipHandle => _tooltip.PopupHandle;

    /// <summary>Top-left of the tooltip in this view's coordinates while it is open.</summary>
    public Point? TooltipPlacement => _tooltip.Placement;

    private bool HoverIsGold => _hoverNode is not null && ReferenceEquals(_hoverNode, _gold);

    // ----- Commands from the window -----

    /// <summary>Blank text area (the new root is being loaded).</summary>
    public void ShowLoading()
    {
        _loading = true;
        Reset(null);
    }

    /// <summary>Shows a root: clears hover, gold and tooltip.</summary>
    public void ShowFormula(FormulaDisplay display)
    {
        _loading = false;
        Reset(display);
    }

    public void SetGold(SyntaxNode? node)
    {
        _gold = node;
        if (HoverIsGold) CancelTooltip();
        else if (node is null) ScheduleTooltip(); // the hovered node is an ordinary node again
        UpdateCursor();
        InvalidateVisual();
    }

    /// <summary>Re-asks the provider about the hovered node (values changed).</summary>
    public void RefreshHover()
    {
        if (_hoverNode is null) return;
        _hoverInfo = Query(_hoverNode);
        if (_tooltip.IsOpen)
        {
            if (_hoverInfo.TooltipText is { } text) _tooltip.SetText(text);
            else _tooltip.Hide();
        }
        else if (!CanShowTooltip()) _tooltipTimer.Stop();
        else if (!_tooltipTimer.IsEnabled) StartTooltipTimer();
        // else: a countdown is already running for this node. New values must not restart it (recalculations every
        // few hundred ms would otherwise keep the tooltip from ever appearing); the tick shows the current text.
        UpdateCursor();
        InvalidateVisual();
    }

    /// <summary>Re-evaluates what is under the real mouse pointer (after scrolling or a re-layout).</summary>
    public void RefreshPointer()
    {
        if (IsMouseOver) UpdateHover(Mouse.GetPosition(this));
    }

    /// <summary>Closes the tooltip and stops its timer (window closing).</summary>
    public void Shutdown()
    {
        _tooltipTimer.Stop();
        _tooltip.Hide();
    }

    // ----- Test surface (goes through the same paths as the mouse) -----

    /// <summary>Rectangle of a char in view coordinates, or null when not laid out (or a zero-width char).</summary>
    public Rect? GetCharRect(int charIndex)
    {
        EnsureLayoutForCurrentWidth();
        if (_layout.CharRect(charIndex) is not { } r) return null;
        r.Offset(TextInset.Left, TextInset.Top);
        return r;
    }

    /// <summary>Center of a char in view coordinates when it is inside the scroll viewport.</summary>
    public Point? GetCharCenter(int charIndex)
    {
        if (GetCharRect(charIndex) is not { } r) return null;
        if (ScrollHost is { } host && (r.Top < host.VerticalOffset || r.Bottom > host.VerticalOffset + host.ViewportHeight + 0.5))
            return null;
        return new Point(r.Left + r.Width / 2, r.Top + r.Height / 2);
    }

    public void SimulateMove(Point? point)
    {
        UpdateHover(point);
        if (_tooltipTimer.IsEnabled) OnTooltipTimer(); // "shows the tooltip immediately"
    }

    public void SimulatePress(Point point, int clickCount) => HandlePress(point, clickCount, null);

    // ----- Layout & rendering -----

    protected override Size MeasureOverride(Size availableSize)
    {
        double wrap = double.IsInfinity(availableSize.Width) ? double.PositiveInfinity : Math.Max(0, availableSize.Width - TextInset.Left - TextInset.Right);
        EnsureLayout(wrap);
        double width = double.IsInfinity(wrap)
            ? TextInset.Left + TextInset.Right + (_layout.Lines.Count == 0 ? 0 : _layout.Lines.Max(l => l.Width))
            : availableSize.Width;
        return new Size(Math.Ceiling(width), TextInset.Top + _layout.Height);
    }

    protected override Size ArrangeOverride(Size finalSize)
    {
        EnsureLayout(Math.Max(0, finalSize.Width - TextInset.Left - TextInset.Right));
        return finalSize;
    }

    protected override void OnRender(DrawingContext dc)
    {
        // A transparent fill keeps the empty area hit-testable (mouse leaves nodes, wheel scrolls).
        dc.DrawRectangle(Brushes.Transparent, null, new Rect(RenderSize));
        if (_layout.Lines.Count == 0) return;

        double ppd = VisualTreeHelper.GetDpi(this).PixelsPerDip;
        var origin = new Point(TextInset.Left, TextInset.Top);
        bool hoverIsGold = HoverIsGold;

        if (_hoverNode is not null && !hoverIsGold)
            FillSpan(dc, HoverBrush, _hoverNode, origin, ppd);
        if (_gold is not null)
            FillSpan(dc, GoldBrush, _gold, origin, ppd);

        for (int i = 0; i < _layout.Lines.Count; i++)
            dc.DrawText(LineText(i), new Point(origin.X, origin.Y + i * TextLayout.LinePitch + _textTop));

        if (HoverIsLink && _hoverNode is not null)
            Underline(dc, _hoverNode, origin, ppd);
    }

    private void FillSpan(DrawingContext dc, Brush brush, SyntaxNode node, Point origin, double ppd)
    {
        foreach (var rect in _layout.SpanRects(node.Start, node.Length))
            dc.DrawRectangle(brush, null, Snap(rect, origin, ppd));
    }

    /// <summary>Underlines the whole node, each char in its own color (black outside colored references).</summary>
    private void Underline(DrawingContext dc, SyntaxNode node, Point origin, double ppd)
    {
        double thickness = Math.Max(1, Math.Round(ppd)) / ppd;
        int end = Math.Min(node.End, _text.Length);
        for (int i = Math.Max(0, _layout.LineOf(Math.Max(node.Start, 0))); i >= 0 && i < _layout.Lines.Count && _layout.Lines[i].Start < end; i++)
        {
            var line = _layout.Lines[i];
            int from = Math.Max(node.Start, line.Start);
            int to = Math.Min(end, line.End);
            double y = SnapValue(origin.Y + i * TextLayout.LinePitch + _underlineTop, ppd);
            int runStart = from;
            while (runStart < to)
            {
                uint? color = _charColors[runStart];
                int runEnd = runStart + 1;
                while (runEnd < to && _charColors[runEnd] == color) runEnd++;
                double left = SnapValue(origin.X + line.Offsets[runStart - line.Start], ppd);
                double right = SnapValue(origin.X + line.Offsets[runEnd - line.Start], ppd);
                if (right > left)
                    dc.DrawRectangle(BrushFor(color), null, new Rect(left, y, right - left, thickness));
                runStart = runEnd;
            }
        }
    }

    private FormattedText LineText(int index)
    {
        var cached = _lineTexts[index];
        if (cached is not null) return cached;

        var line = _layout.Lines[index];
        var runs = new List<(int, int, Brush)>();
        int i = 0;
        while (i < line.Length)
        {
            uint? color = _charColors[line.Start + i];
            int j = i + 1;
            while (j < line.Length && _charColors[line.Start + j] == color) j++;
            if (color is not null) runs.Add((i, j - i, BrushFor(color)));
            i = j;
        }
        return _lineTexts[index] = _measurer!.CreateLine(line.Start, line.Length, runs);
    }

    /// <summary>Brings the layout up to date with the last wrap width WPF gave us (no-op before the first layout pass).</summary>
    private void EnsureLayoutForCurrentWidth()
    {
        if (_wrapWidth is { } wrap) EnsureLayout(wrap);
    }

    /// <summary>Rebuilds the line layout when the text, the wrap width or the DPI changed.</summary>
    private void EnsureLayout(double wrapWidth)
    {
        double ppd = VisualTreeHelper.GetDpi(this).PixelsPerDip;
        if (!_layoutDirty && ppd == _layoutPpd && wrapWidth == _wrapWidth) return;

        _layoutDirty = false;
        _wrapWidth = wrapWidth;
        _layoutPpd = ppd;
        _measurer = new FormattedTextMeasurer(_text, FormulaTypeface, FontSizeDip, ppd);
        _layout = _text.Length == 0 ? TextLayout.Empty : TextLayout.Build(_text, wrapWidth, _measurer);
        _lineTexts = new FormattedText?[_layout.Lines.Count];

        // Glyphs sit centered in the 20 DIP line box, which for Calibri 11 pt puts the baseline 14 DIPs below the line
        // top (measured on the original); the underline is 2 DIPs under the baseline.
        var sample = new FormattedTextMeasurer("Ag", FormulaTypeface, FontSizeDip, ppd).CreateLine(0, 2, Array.Empty<(int, int, Brush)>());
        _textTop = SnapValue(BaselineOffset - sample.Baseline, ppd);
        _underlineTop = BaselineOffset + 2;
    }

    // ----- Hover, tooltip, clicks -----

    protected override void OnMouseMove(MouseEventArgs e)
    {
        base.OnMouseMove(e);
        Guard("FormulaView.MouseMove", () => UpdateHover(e.GetPosition(this)));
    }

    protected override void OnMouseLeave(MouseEventArgs e)
    {
        base.OnMouseLeave(e);
        Guard("FormulaView.MouseLeave", () => UpdateHover(null));
    }

    protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e)
    {
        base.OnMouseLeftButtonDown(e);
        Guard("FormulaView.MouseDown", () => HandlePress(e.GetPosition(this), e.ClickCount, e.Timestamp));
    }

    /// <summary>The pointer moved to <paramref name="point"/> (view coordinates) or left the view (null).</summary>
    internal void UpdateHover(Point? point)
    {
        EnsureLayoutForCurrentWidth(); // text may have changed since the last layout pass
        _pointer = point;
        SyntaxNode? node = null;
        if (point is { } p && _interactive && !_loading && _display is not null
            && _layout.HitTest(new Point(p.X - TextInset.Left, p.Y - TextInset.Top)) is int index)
            node = _display.HitMap.NodeAt(index);
        SetHoverNode(node);
    }

    private void SetHoverNode(SyntaxNode? node)
    {
        if (ReferenceEquals(node, _hoverNode)) return;

        CancelTooltip();
        _tooltipSuppressedFor = null;
        _hoverNode = node;
        _hoverInfo = node is null ? NoInfo : Query(node);
        ScheduleTooltip();
        UpdateCursor();
        InvalidateVisual();
    }

    private NodeInfo Query(SyntaxNode node)
    {
        try
        {
            return _provider.GetNodeInfo(node) ?? NoInfo;
        }
        catch (Exception ex)
        {
            Log.Error("FormulaView.GetNodeInfo", ex);
            return NoInfo;
        }
    }

    private void HandlePress(Point point, int clickCount, int? timestampMs)
    {
        UpdateHover(point);
        // A click hides the tooltip, and it stays hidden while the pointer remains on the clicked node.
        CancelTooltip();
        _tooltipSuppressedFor = _hoverNode;

        var node = _hoverNode;
        if (node is null || !_hoverInfo.IsLink)
        {
            _doubleClicks.Reset();
            return;
        }

        switch (_doubleClicks.OnLinkPress(node, point, clickCount, timestampMs))
        {
            case PressAction.Click:
                NodeClicked?.Invoke(this, node);
                break;
            case PressAction.DoubleClick:
                NodeDoubleClicked?.Invoke(this, node);
                break;
            case PressAction.ClickThenDoubleClick:
                NodeClicked?.Invoke(this, node);
                NodeDoubleClicked?.Invoke(this, node);
                break;
        }
    }

    private void ScheduleTooltip()
    {
        _tooltipTimer.Stop();
        if (CanShowTooltip()) StartTooltipTimer();
    }

    private void StartTooltipTimer()
    {
        TooltipTimerStarts++;
        _tooltipTimer.Start();
    }

    /// <summary>How many times the tooltip countdown was started (tests: a refresh must not restart a running one).</summary>
    internal int TooltipTimerStarts { get; private set; }

    /// <summary>True while the tooltip countdown is running.</summary>
    internal bool TooltipTimerRunning => _tooltipTimer.IsEnabled;

    /// <summary>
    /// Tests only: the countdown's length (so a test can rule wall-clock time out by making it very long) and a way to let it
    /// expire at once. Takes effect at the next start of the countdown.
    /// </summary>
    internal TimeSpan TooltipDelay
    {
        get => _tooltipTimer.Interval;
        set => _tooltipTimer.Interval = value;
    }

    /// <summary>Tests only: the countdown expires now (when it is running).</summary>
    internal void ExpireTooltipCountdown()
    {
        if (_tooltipTimer.IsEnabled) Guard("FormulaView.TooltipTimer", OnTooltipTimer);
    }

    private bool CanShowTooltip() =>
        _hoverNode is not null && !HoverIsGold && _hoverInfo.TooltipText is not null
        && !ReferenceEquals(_tooltipSuppressedFor, _hoverNode) && _pointer is not null;

    private void OnTooltipTimer()
    {
        _tooltipTimer.Stop();
        if (CanShowTooltip()) _tooltip.Show(_pointer!.Value, _hoverInfo.TooltipText!);
    }

    private void CancelTooltip()
    {
        _tooltipTimer.Stop();
        _tooltip.Hide();
    }

    private void UpdateCursor() => Cursor = HoverIsLink ? Cursors.Hand : Cursors.Arrow;

    private void Reset(FormulaDisplay? display)
    {
        CancelTooltip();
        _doubleClicks.Reset();
        _hoverNode = null;
        _hoverInfo = NoInfo;
        _gold = null;
        _tooltipSuppressedFor = null;
        _display = display;
        _text = display?.Formula.Text.Replace('\t', ' ') ?? string.Empty;
        _interactive = display is { Interactive: true } && display.Formula.IsFormula;
        _charColors = new uint?[_text.Length];
        if (_interactive)
        {
            foreach (var colored in display!.Colors.Nodes)
                for (int i = Math.Max(0, colored.Node.Start); i < Math.Min(colored.Node.End, _text.Length); i++)
                    _charColors[i] = colored.Rgb;
        }
        _layoutDirty = true;
        UpdateCursor();
        InvalidateMeasure();
        InvalidateVisual();
    }

    // ----- Helpers -----

    private Brush BrushFor(uint? rgb)
    {
        uint value = rgb ?? 0;
        if (_brushes.TryGetValue(value, out var brush)) return brush;
        brush = Frozen(new SolidColorBrush(Color.FromRgb((byte)(value >> 16), (byte)(value >> 8), (byte)value)));
        _brushes[value] = brush;
        return brush;
    }

    private static Rect Snap(Rect rect, Point origin, double ppd)
    {
        double left = SnapValue(rect.Left + origin.X, ppd);
        double top = SnapValue(rect.Top + origin.Y, ppd);
        double right = SnapValue(rect.Right + origin.X, ppd);
        double bottom = SnapValue(rect.Bottom + origin.Y, ppd);
        return new Rect(left, top, right - left, bottom - top);
    }

    private static double SnapValue(double value, double ppd) => Math.Round(value * ppd, MidpointRounding.AwayFromZero) / ppd;

    private static void Guard(string context, Action action)
    {
        try
        {
            action();
        }
        catch (Exception ex)
        {
            Log.Error(context, ex);
        }
    }

    private static T Frozen<T>(T freezable) where T : Freezable
    {
        freezable.Freeze();
        return freezable;
    }
}
