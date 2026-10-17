using System.Globalization;
using System.Windows;
using System.Windows.Controls.Primitives;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Effects;
using FormulaExplorer.AddIn.Diagnostics;

namespace FormulaExplorer.AddIn.UI;

/// <summary>
/// The Windows 7 style "Returns" tooltip: 1 px #767676 border, vertical gradient #FFFFFF → #E5E6F0, bold "Returns"
/// title and the value under it (Segoe UI 9 pt, #575757), small drop shadow. Metrics were measured on the original's
/// screenshots: the first baseline sits 14 px below the outer top edge and lines are 16 px apart.
/// </summary>
internal static class ReturnsToolTipVisual
{
    internal const string Title = "Returns";

    private const double MaxTextWidth = 380;

    /// <summary>Distance of the first text baseline from the inner (padding-box) top edge.</summary>
    private const double FirstBaseline = 13;
    private const double LinePitch = 16;
    private const double BottomPad = 5;
    private const double LeftPad = 5;
    private const double RightPad = 6;
    private const double ShadowRoom = 6;

    private static readonly Brush TextBrush = Freeze(new SolidColorBrush(Color.FromRgb(0x57, 0x57, 0x57)));
    private static readonly Brush BorderBrush = Freeze(new SolidColorBrush(Color.FromRgb(0x76, 0x76, 0x76)));
    private static readonly Brush BackgroundBrush = Freeze(new LinearGradientBrush(
        new GradientStopCollection
        {
            // The original's gradient is flat white over its first rows and flat blue-gray over its last ones.
            new GradientStop(Color.FromRgb(0xFF, 0xFF, 0xFF), 0.09),
            new GradientStop(Color.FromRgb(0xE5, 0xE6, 0xF0), 0.91),
        }, new Point(0, 0), new Point(0, 1)));
    private static readonly FontFamily Font = new("Segoe UI");

    /// <summary>Builds the tooltip visual (including room for the shadow on the right and bottom).</summary>
    public static FrameworkElement Create(string value)
    {
        var border = new System.Windows.Controls.Border
        {
            BorderBrush = BorderBrush,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(2),
            Background = BackgroundBrush,
            Child = new TextElement(value),
            Effect = new DropShadowEffect
            {
                Color = Colors.Black,
                Opacity = 0.5,
                BlurRadius = 4,
                ShadowDepth = 3,
                Direction = 315,
                RenderingBias = RenderingBias.Performance,
            },
            Margin = new Thickness(0, 0, ShadowRoom, ShadowRoom),
        };
        var root = new System.Windows.Controls.Grid { IsHitTestVisible = false, Focusable = false, UseLayoutRounding = true, SnapsToDevicePixels = true };
        root.Children.Add(border);
        TextOptions.SetTextFormattingMode(root, TextFormattingMode.Display);
        return root;
    }

    private static T Freeze<T>(T freezable) where T : Freezable
    {
        freezable.Freeze();
        return freezable;
    }

    /// <summary>Draws the title and the value lines at the measured baselines.</summary>
    private sealed class TextElement : FrameworkElement
    {
        private readonly string _value;
        private FormattedText? _title;
        private FormattedText? _body;
        private double _ppd;

        public TextElement(string value) => _value = value;

        private FormattedText Format(string text, FontWeight weight, double ppd, double maxWidth) =>
            new(text, CultureInfo.InvariantCulture, FlowDirection.LeftToRight,
                new Typeface(Font, FontStyles.Normal, weight, FontStretches.Normal), 12, TextBrush,
                null, TextFormattingMode.Display, ppd)
            {
                MaxTextWidth = maxWidth,
                LineHeight = LinePitch,
                Trimming = TextTrimming.None,
            };

        private void EnsureText()
        {
            double ppd = VisualTreeHelper.GetDpi(this).PixelsPerDip;
            if (_title is not null && ppd == _ppd) return;
            _ppd = ppd;
            _title = Format(Title, FontWeights.Bold, ppd, MaxTextWidth);
            _body = Format(_value.Replace("\r", string.Empty), FontWeights.Normal, ppd, MaxTextWidth);
        }

        protected override Size MeasureOverride(Size availableSize)
        {
            EnsureText();
            double width = Math.Max(_title!.WidthIncludingTrailingWhitespace, _body!.Width) + LeftPad + RightPad;
            int bodyLines = Math.Max(1, _value.Replace("\r", string.Empty).Split('\n').Length);
            double lines = Math.Max(bodyLines, Math.Ceiling(_body.Height / LinePitch));
            return new Size(Math.Ceiling(width), FirstBaseline + LinePitch * lines + BottomPad);
        }

        protected override void OnRender(DrawingContext dc)
        {
            EnsureText();
            dc.DrawText(_title!, new Point(LeftPad, FirstBaseline - _title!.Baseline));
            dc.DrawText(_body!, new Point(LeftPad, FirstBaseline + LinePitch - _body!.Baseline));
        }
    }
}

/// <summary>
/// Shows the "Returns" tooltip as a custom <see cref="Popup"/> (never auto-hides, never takes the mouse). The
/// popup is placed relative to a target element: left edge at the pointer x, top 17 DIPs under the pointer, and it
/// does not move afterwards. When the target is not attached to a window (headless rendering) only the logical state
/// is tracked.
/// </summary>
internal sealed class ReturnsToolTip
{
    /// <summary>Vertical distance between the pointer and the tooltip's top edge.</summary>
    internal const double PointerOffsetY = 17;

    private readonly FrameworkElement _target;
    private Popup? _popup;

    public ReturnsToolTip(FrameworkElement target) => _target = target;

    public bool IsOpen { get; private set; }

    /// <summary>The value text currently shown (null when closed).</summary>
    public string? Text { get; private set; }

    /// <summary>Top-left corner of the tooltip in the target's coordinates (null when closed).</summary>
    public Point? Placement { get; private set; }

    /// <summary>HWND of the open popup (zero when closed or headless); for tests.</summary>
    internal IntPtr PopupHandle { get; private set; }

    /// <summary>Shows the tooltip for <paramref name="pointer"/> (target coordinates) or updates its text when already open.</summary>
    public void Show(Point pointer, string text)
    {
        try
        {
            if (IsOpen)
            {
                if (text != Text) SetText(text);
                return;
            }

            Text = text;
            Placement = new Point(pointer.X, pointer.Y + PointerOffsetY);
            IsOpen = true;
            if (PresentationSource.FromVisual(_target) is null) return;

            var popup = EnsurePopup();
            popup.Child = ReturnsToolTipVisual.Create(text);
            popup.HorizontalOffset = Placement.Value.X;
            popup.VerticalOffset = Placement.Value.Y;
            popup.IsOpen = true;
        }
        catch (Exception ex)
        {
            Log.Error("ReturnsToolTip.Show", ex);
        }
    }

    /// <summary>Replaces the value while the tooltip stays where it is.</summary>
    public void SetText(string text)
    {
        try
        {
            Text = text;
            if (_popup is { IsOpen: true }) _popup.Child = ReturnsToolTipVisual.Create(text);
        }
        catch (Exception ex)
        {
            Log.Error("ReturnsToolTip.SetText", ex);
        }
    }

    public void Hide()
    {
        try
        {
            IsOpen = false;
            Text = null;
            Placement = null;
            PopupHandle = IntPtr.Zero;
            if (_popup is not null)
            {
                _popup.IsOpen = false;
                _popup.Child = null;
            }
        }
        catch (Exception ex)
        {
            Log.Error("ReturnsToolTip.Hide", ex);
        }
    }

    private Popup EnsurePopup()
    {
        if (_popup is not null) return _popup;
        _popup = new Popup
        {
            AllowsTransparency = true,
            StaysOpen = true,
            Placement = PlacementMode.Relative,
            PlacementTarget = _target,
            PopupAnimation = PopupAnimation.None,
            Focusable = false,
            IsHitTestVisible = false,
        };
        _popup.Opened += OnPopupOpened;
        return _popup;
    }

    /// <summary>
    /// Makes the popup's HWND click-through and non-activating: otherwise moving the pointer onto the line below (where
    /// the tooltip sits) would steal the mouse from the view and the tooltip would flicker.
    /// </summary>
    private void OnPopupOpened(object? sender, EventArgs e)
    {
        try
        {
            if (_popup?.Child is { } child && PresentationSource.FromVisual(child) is HwndSource source)
            {
                long style = NativeMethods.GetWindowLong(source.Handle, NativeMethods.GWL_EXSTYLE);
                NativeMethods.SetWindowLong(source.Handle, NativeMethods.GWL_EXSTYLE,
                    style | NativeMethods.WS_EX_TRANSPARENT | NativeMethods.WS_EX_NOACTIVATE);
                PopupHandle = source.Handle;
            }
        }
        catch (Exception ex)
        {
            Log.Error("ReturnsToolTip.OnPopupOpened", ex);
        }
    }
}
