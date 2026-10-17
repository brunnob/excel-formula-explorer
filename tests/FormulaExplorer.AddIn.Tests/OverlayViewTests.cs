using System.Runtime.InteropServices;
using FormulaExplorer.AddIn.Overlay;

namespace FormulaExplorer.AddIn.Tests;

public class ViewSignatureTests
{
    private static ViewSignature Sample() => new(
        ActiveHwnd: 0x1234, SheetName: "Sheet1", ScrollRow: 1, ScrollColumn: 1, Zoom: 100, SplitRow: 0, SplitColumn: 0,
        FreezePanes: false, PaneCount: 1, PaneScroll: "", GridRect: new PixelRect(10, 20, 1000, 700), GridDpi: 144);

    [Fact]
    public void Equal_signatures_compare_equal() => Assert.Equal(Sample(), Sample());

    private static readonly Func<ViewSignature, ViewSignature>[] Changes =
    [
        s => s with { ActiveHwnd = 0x9999 },
        s => s with { SheetName = "Other" },
        s => s with { ScrollRow = 20 },
        s => s with { ScrollColumn = 3 },
        s => s with { Zoom = 150 },
        s => s with { SplitRow = 5 },
        s => s with { SplitColumn = 2 },
        s => s with { FreezePanes = true },
        s => s with { PaneCount = 4 },
        s => s with { PaneScroll = "1,1;5,1;" },
        s => s with { GridRect = new PixelRect(11, 20, 1001, 700) },
        s => s with { GridRect = new PixelRect(10, 20, 1000, 701) },
        s => s with { GridDpi = 96 },
    ];

    [Fact]
    public void Every_component_is_covered() => Assert.Equal(13, Changes.Length);

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    [InlineData(6)]
    [InlineData(7)]
    [InlineData(8)]
    [InlineData(9)]
    [InlineData(10)]
    [InlineData(11)]
    [InlineData(12)]
    public void Any_component_change_makes_signatures_differ(int index) =>
        Assert.NotEqual(Sample(), Changes[index](Sample()));
}

public class ViewChangeDetectorTests
{
    private static ViewSignature Sig(int scrollRow = 1) => new(
        1, "Sheet1", scrollRow, 1, 100, 0, 0, false, 1, "", new PixelRect(0, 0, 100, 100), 96);

    [Fact]
    public void First_call_always_refreshes()
    {
        var d = new ViewChangeDetector();
        Assert.True(d.ShouldRefresh(Sig(), 0));
    }

    [Fact]
    public void Unchanged_signature_does_not_refresh_before_the_forced_interval()
    {
        var d = new ViewChangeDetector();
        d.ShouldRefresh(Sig(), 0);

        Assert.False(d.ShouldRefresh(Sig(), 50));
        Assert.False(d.ShouldRefresh(Sig(), 499));
    }

    [Fact]
    public void Unchanged_signature_refreshes_every_500_ms()
    {
        var d = new ViewChangeDetector();
        d.ShouldRefresh(Sig(), 1000);

        Assert.False(d.ShouldRefresh(Sig(), 1499));
        Assert.True(d.ShouldRefresh(Sig(), 1500));
        // The forced refresh restarts the interval.
        Assert.False(d.ShouldRefresh(Sig(), 1600));
        Assert.True(d.ShouldRefresh(Sig(), 2000));
    }

    [Fact]
    public void Changed_signature_refreshes_immediately_and_becomes_the_baseline()
    {
        var d = new ViewChangeDetector();
        d.ShouldRefresh(Sig(1), 0);

        Assert.True(d.ShouldRefresh(Sig(7), 10));
        Assert.False(d.ShouldRefresh(Sig(7), 60));
    }

    [Fact]
    public void Hidden_state_is_a_signature_of_its_own()
    {
        var d = new ViewChangeDetector();
        d.ShouldRefresh(Sig(), 0);

        Assert.True(d.ShouldRefresh(null, 50));      // became hidden
        Assert.False(d.ShouldRefresh(null, 100));    // still hidden
        Assert.True(d.ShouldRefresh(Sig(), 150));    // visible again
    }

    [Fact]
    public void Reset_forces_the_next_refresh()
    {
        var d = new ViewChangeDetector();
        d.ShouldRefresh(Sig(), 0);
        d.Reset();

        Assert.True(d.ShouldRefresh(Sig(), 10));
    }

    [Fact]
    public void Tracker_constants_match_the_contract()
    {
        Assert.Equal(50, ViewTracker.IntervalMs);
        Assert.Equal(500, ViewChangeDetector.ForcedRefreshMs);
    }
}

/// <summary>
/// Creates the real Win32 overlay window (hidden, unowned) to check styles, hit-testing and lifetime. No Excel needed.
/// </summary>
[Collection(DesktopWindows.Name)]
public class OverlayWindowTests
{
    private const int GWL_EXSTYLE = -20;
    private const uint WS_EX_LAYERED = 0x80000;
    private const uint WS_EX_TRANSPARENT = 0x20;
    private const uint WS_EX_NOACTIVATE = 0x08000000;
    private const uint WS_EX_TOOLWINDOW = 0x80;
    private const uint WM_NCHITTEST = 0x84;
    private const uint WM_MOUSEACTIVATE = 0x21;

    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")]
    private static extern IntPtr GetWindowLongPtr(IntPtr hWnd, int index);

    [DllImport("user32.dll", EntryPoint = "SendMessageW")]
    private static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr GetDC(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern int ReleaseDC(IntPtr hWnd, IntPtr hdc);

    [DllImport("gdi32.dll")]
    private static extern IntPtr CreateCompatibleDC(IntPtr hdc);

    [DllImport("gdi32.dll")]
    private static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int width, int height);

    [DllImport("gdi32.dll")]
    private static extern IntPtr SelectObject(IntPtr hdc, IntPtr obj);

    [DllImport("gdi32.dll")]
    private static extern uint GetPixel(IntPtr hdc, int x, int y);

    [DllImport("gdi32.dll")]
    private static extern bool DeleteObject(IntPtr obj);

    [DllImport("gdi32.dll")]
    private static extern bool DeleteDC(IntPtr hdc);

    private const uint WM_PRINTCLIENT = 0x318;

    [Fact]
    public void Painting_fills_key_color_then_the_rectangles_in_window_local_coordinates()
    {
        using var window = new OverlayWindow(IntPtr.Zero);
        window.Place(new PixelRect(3000, 2000, 3100, 2050), IntPtr.Zero);
        window.SetContent(
        [
            new ColoredRect(new PixelRect(3010, 2010, 3030, 2012), 0x2E75B6),   // blue strip at local (10,10)-(30,12)
            new ColoredRect(new PixelRect(3020, 2011, 3040, 2020), 0xC00000),   // red on top, overlapping the strip
        ]);

        var screen = GetDC(IntPtr.Zero);
        var dc = CreateCompatibleDC(screen);
        var bitmap = CreateCompatibleBitmap(screen, 100, 50);
        var old = SelectObject(dc, bitmap);
        try
        {
            SendMessage(window.Handle, WM_PRINTCLIENT, dc, IntPtr.Zero);

            Assert.Equal(NativeMethods.ToColorRef(OverlayWindow.KeyColorRgb), GetPixel(dc, 0, 0));
            Assert.Equal(NativeMethods.ToColorRef(OverlayWindow.KeyColorRgb), GetPixel(dc, 99, 49));
            Assert.Equal(NativeMethods.ToColorRef(0x2E75B6), GetPixel(dc, 10, 10));
            Assert.Equal(NativeMethods.ToColorRef(0x2E75B6), GetPixel(dc, 19, 11));
            Assert.Equal(NativeMethods.ToColorRef(OverlayWindow.KeyColorRgb), GetPixel(dc, 10, 12));   // end is exclusive
            Assert.Equal(NativeMethods.ToColorRef(0xC00000), GetPixel(dc, 20, 11));                      // later rect on top
            Assert.Equal(NativeMethods.ToColorRef(0xC00000), GetPixel(dc, 39, 19));
            Assert.Equal(NativeMethods.ToColorRef(OverlayWindow.KeyColorRgb), GetPixel(dc, 40, 19));
        }
        finally
        {
            SelectObject(dc, old);
            DeleteObject(bitmap);
            DeleteDC(dc);
            ReleaseDC(IntPtr.Zero, screen);
        }
    }

    [Fact]
    public void Window_is_layered_click_through_and_non_activating()
    {
        using var window = new OverlayWindow(IntPtr.Zero);

        var ex = (uint)GetWindowLongPtr(window.Handle, GWL_EXSTYLE).ToInt64();
        Assert.Equal(WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW,
            ex & (WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW));
        Assert.False(window.IsVisible);
    }

    [Fact]
    public void Hit_test_is_transparent_and_mouse_never_activates()
    {
        using var window = new OverlayWindow(IntPtr.Zero);

        Assert.Equal(new IntPtr(-1), SendMessage(window.Handle, WM_NCHITTEST, IntPtr.Zero, new IntPtr(0x00640064)));
        Assert.Equal(new IntPtr(3), SendMessage(window.Handle, WM_MOUSEACTIVATE, IntPtr.Zero, IntPtr.Zero));
    }

    [Fact]
    public void Place_and_content_work_while_hidden_and_dispose_destroys_the_window()
    {
        var window = new OverlayWindow(IntPtr.Zero);
        var handle = window.Handle;

        window.Place(new PixelRect(5000, 5000, 5200, 5100), IntPtr.Zero);
        window.SetContent([new ColoredRect(new PixelRect(5010, 5010, 5050, 5012), 0xFF0000)]);
        Assert.True(window.Dpi >= 96);

        window.Dispose();
        Assert.False(IsWindow(handle));
        window.Dispose();   // idempotent
    }

    [Fact]
    public void Two_windows_can_coexist_and_a_class_can_be_recreated()
    {
        using var a = new OverlayWindow(IntPtr.Zero);
        using var b = new OverlayWindow(IntPtr.Zero);
        Assert.NotEqual(a.Handle, b.Handle);
    }

    [Fact]
    public void Safe_rgb_never_returns_the_color_key()
    {
        Assert.NotEqual(OverlayWindow.KeyColorRgb, OverlayWindow.SafeRgb(OverlayWindow.KeyColorRgb));
        Assert.Equal(0x2E75B6u, OverlayWindow.SafeRgb(0x2E75B6));
        Assert.Equal(0xFFFFFFu, OverlayWindow.SafeRgb(0xFFFFFFFF));
    }

    [Fact]
    public void Color_ref_swaps_red_and_blue() =>
        Assert.Equal(0x00B6752Eu, NativeMethods.ToColorRef(0x2E75B6));
}
