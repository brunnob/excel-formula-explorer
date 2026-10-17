using System.Runtime.InteropServices;
using FormulaExplorer.AddIn.Overlay;

namespace FormulaExplorer.AddIn.Tests;

public class AreaAddressTests
{
    [Theory]
    [InlineData("$A$1", new[] { "$A$1" })]
    [InlineData("$A$1:$B$3", new[] { "$A$1:$B$3" })]
    [InlineData("$A$1:$A$3,$C$1:$C$3", new[] { "$A$1:$A$3", "$C$1:$C$3" })]
    [InlineData("$A:$A,$3:$3", new[] { "$A:$A", "$3:$3" })]
    [InlineData("$A$1, $C$3 ,$E$5", new[] { "$A$1", "$C$3", "$E$5" })]
    [InlineData("'Sales, 2024'!$A$1,'Sales, 2024'!$B$2", new[] { "'Sales, 2024'!$A$1", "'Sales, 2024'!$B$2" })]
    [InlineData("", new string[0])]
    [InlineData(" , ", new string[0])]
    public void Split_separates_areas_on_commas(string address, string[] expected) =>
        Assert.Equal(expected, AreaAddress.Split(address));
}

public class PaneClippingTests
{
    // 150% DPI: 144.
    private static readonly PixelRect Grid = new(0, 0, 1000, 700);

    [Fact]
    public void Cell_area_is_the_usable_size_in_pixels_from_the_first_pane_origin()
    {
        // Usable 480 x 300 pt at 144 DPI = 960 x 600 px, origin (30, 40).
        var area = PaneClipping.CellArea(Grid, new PixelRect(30, 40, 990, 700), 480, 300, 144);

        Assert.Equal(new PixelRect(30, 40, 990, 640), area);
    }

    [Fact]
    public void Cell_area_never_leaves_the_grid_window()
    {
        var area = PaneClipping.CellArea(Grid, new PixelRect(30, 40, 990, 700), 700, 400, 144);

        Assert.Equal(new PixelRect(30, 40, 1000, 700), area);
    }

    [Theory]
    [InlineData(0.0, 300.0)]
    [InlineData(480.0, 0.0)]
    [InlineData(double.NaN, 300.0)]
    [InlineData(480.0, double.PositiveInfinity)]
    [InlineData(480.0, 20.0)]      // implausibly small: not believed
    [InlineData(20.0, 300.0)]
    public void Cell_area_falls_back_to_the_grid_when_sizes_are_missing_or_implausible(double w, double h) =>
        Assert.Equal(Grid, PaneClipping.CellArea(Grid, new PixelRect(30, 40, 990, 700), w, h, 144));

    [Fact]
    public void Cell_area_falls_back_to_the_grid_for_an_invalid_dpi() =>
        Assert.Equal(Grid, PaneClipping.CellArea(Grid, new PixelRect(30, 40, 990, 700), 480, 300, 0));

    [Fact]
    public void Single_pane_is_cut_at_the_cell_area_so_borders_stay_off_the_scroll_bars()
    {
        // The partly visible last row reaches y = 700 but the cell area ends at 640 (scroll bar below).
        var cell = new PixelRect(30, 40, 990, 640);
        var clipped = PaneClipping.Clip([new PixelRect(30, 40, 990, 700)], cell);

        Assert.Equal(cell, Assert.Single(clipped));
    }

    [Fact]
    public void Split_panes_stop_where_the_neighbour_starts()
    {
        // Upper pane ends mid-row (y = 330) but the lower pane starts at 300 (after the split bar).
        var cell = new PixelRect(30, 40, 990, 640);
        var clipped = PaneClipping.Clip([new PixelRect(30, 40, 990, 330), new PixelRect(30, 300, 990, 700)], cell);

        Assert.Equal(new PixelRect(30, 40, 990, 300), clipped[0]);
        Assert.Equal(new PixelRect(30, 300, 990, 640), clipped[1]);
    }

    [Fact]
    public void Split_columns_stop_where_the_right_pane_starts()
    {
        var cell = new PixelRect(30, 40, 990, 640);
        var clipped = PaneClipping.Clip([new PixelRect(30, 40, 530, 700), new PixelRect(500, 40, 1100, 700)], cell);

        Assert.Equal(new PixelRect(30, 40, 500, 640), clipped[0]);
        Assert.Equal(new PixelRect(500, 40, 990, 640), clipped[1]);
    }

    [Fact]
    public void Four_panes_do_not_overlap_after_clipping()
    {
        var cell = new PixelRect(30, 40, 990, 640);
        PixelRect[] raw =
        [
            new(30, 40, 520, 330),     // TL, reaches into TR and BL
            new(500, 40, 1100, 330),   // TR
            new(30, 300, 520, 700),    // BL
            new(500, 300, 1100, 700),  // BR
        ];

        var c = PaneClipping.Clip(raw, cell);

        for (var i = 0; i < c.Length; i++)
            for (var j = i + 1; j < c.Length; j++)
                Assert.True(c[i].Intersect(c[j]).IsEmpty, $"panes {i} and {j} overlap: {c[i]} {c[j]}");
        Assert.Equal(new PixelRect(30, 40, 500, 300), c[0]);
        Assert.Equal(new PixelRect(500, 300, 990, 640), c[3]);
    }

    [Fact]
    public void Frozen_panes_that_already_abut_are_unchanged()
    {
        var cell = new PixelRect(30, 40, 990, 640);
        PixelRect[] raw = [new(30, 40, 990, 100), new(30, 100, 990, 700)];

        var c = PaneClipping.Clip(raw, cell);

        Assert.Equal(new PixelRect(30, 40, 990, 100), c[0]);
        Assert.Equal(new PixelRect(30, 100, 990, 640), c[1]);
    }

    [Fact]
    public void Empty_panes_stay_empty_and_do_not_clip_others()
    {
        var cell = new PixelRect(0, 0, 100, 100);
        var c = PaneClipping.Clip([PixelRect.Empty, new PixelRect(0, 0, 100, 100)], cell);

        Assert.True(c[0].IsEmpty);
        Assert.Equal(cell, c[1]);
    }
}

/// <summary>Real Win32 windows, no Excel: what happens when the overlay's owner (an XLMAIN) goes away.</summary>
// Real top-level windows share one global z-order and one foreground: every class that creates or shows them runs in this collection, i.e. never in parallel with another.
[Collection(DesktopWindows.Name)]
public class OverlayOwnerLifetimeTests
{
    private const uint WS_POPUP = 0x80000000;
    private const int SW_SHOWNOACTIVATE = 4;

    private static IntPtr CreatePlainWindow(IntPtr owner, bool visible = false)
    {
        var className = Marshal.StringToHGlobalUni("Static");
        try
        {
            var hwnd = NativeMethods.CreateWindowExW(
                0x80 /* WS_EX_TOOLWINDOW */, className, "owner", WS_POPUP, 6000, 6000, 10, 10, owner, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
            Assert.NotEqual(IntPtr.Zero, hwnd);
            if (visible) NativeMethods.ShowWindow(hwnd, SW_SHOWNOACTIVATE);
            return hwnd;
        }
        finally
        {
            Marshal.FreeHGlobal(className);
        }
    }

    [Fact]
    public void Overlay_dies_with_its_owner_and_reports_it()
    {
        var owner = CreatePlainWindow(IntPtr.Zero);
        var overlay = new OverlayWindow(owner);
        try
        {
            Assert.True(overlay.IsAlive);

            NativeMethods.DestroyWindow(owner);

            Assert.False(NativeMethods.IsWindow(owner));
            Assert.False(overlay.IsAlive);
            Assert.False(overlay.IsVisible);
        }
        finally
        {
            overlay.Dispose();   // releases the class and brushes although the HWND is already gone
        }
    }

    [Fact]
    public void Dead_overlay_ignores_calls_and_refuses_to_be_placed()
    {
        var owner = CreatePlainWindow(IntPtr.Zero);
        var overlay = new OverlayWindow(owner);
        NativeMethods.DestroyWindow(owner);

        overlay.SetOwner(new IntPtr(0x1234));
        Assert.Equal(owner, overlay.Owner);          // not recorded on a dead window
        overlay.Hide();
        overlay.Show();
        overlay.EnsureAboveOwner();
        Assert.False(overlay.IsVisible);
        Assert.Throws<InvalidOperationException>(() => overlay.Place(new PixelRect(0, 0, 10, 10), IntPtr.Zero));

        overlay.Dispose();
        overlay.Dispose();
    }

    [Fact]
    public void Keeper_recreates_the_window_for_the_new_owner_after_the_old_owner_was_destroyed()
    {
        var ownerA = CreatePlainWindow(IntPtr.Zero);
        var ownerB = CreatePlainWindow(IntPtr.Zero);
        using var keeper = new OverlayWindowKeeper();
        try
        {
            var first = keeper.Ensure(ownerB);
            Assert.True(first.IsAlive);
            Assert.Same(first, keeper.Ensure(ownerB));   // unchanged owner: same window

            var reowned = keeper.Ensure(ownerA);
            Assert.Same(first, reowned);                  // alive owner change: re-owned in place
            Assert.Equal(ownerA, reowned.Owner);

            NativeMethods.DestroyWindow(ownerA);          // the workbook window is closed
            Assert.False(first.IsAlive);

            var second = keeper.Ensure(ownerB);

            Assert.NotSame(first, second);
            Assert.True(second.IsAlive);
            Assert.Equal(ownerB, second.Owner);
            Assert.Equal(ownerB, new IntPtr(GetOwner(second.Handle)));
        }
        finally
        {
            NativeMethods.DestroyWindow(ownerB);
        }
    }

    [Fact]
    public void Keeper_dispose_destroys_the_window()
    {
        var owner = CreatePlainWindow(IntPtr.Zero);
        var keeper = new OverlayWindowKeeper();
        try
        {
            var handle = keeper.Ensure(owner).Handle;
            keeper.Dispose();

            Assert.False(NativeMethods.IsWindow(handle));
            Assert.Null(keeper.Current);
        }
        finally
        {
            NativeMethods.DestroyWindow(owner);
        }
    }

    [Fact]
    public void Overlay_sits_directly_above_its_owner_and_under_other_owned_windows()
    {
        var owner = CreatePlainWindow(IntPtr.Zero, visible: true);
        var popup = CreatePlainWindow(owner, visible: true);       // e.g. Find & Replace, or the explorer
        OverlayWindow? overlay = null;
        try
        {
            overlay = new OverlayWindow(owner);
            overlay.Place(new PixelRect(6000, 6000, 6010, 6010), overlay.AboveOwnerAnchor());
            overlay.SetContent([new ColoredRect(new PixelRect(6000, 6000, 6005, 6005), 0xFF0000)]);
            overlay.Show();

            // Created after the popup, the overlay is initially above it; the invariant puts it right above the owner.
            overlay.EnsureAboveOwner();

            Assert.Equal(overlay.Handle, NativeMethods.GetWindow(owner, NativeMethods.GW_HWNDPREV));
            Assert.Equal(popup, NativeMethods.GetWindow(overlay.Handle, NativeMethods.GW_HWNDPREV));

            // Idempotent.
            overlay.EnsureAboveOwner();
            Assert.Equal(overlay.Handle, NativeMethods.GetWindow(owner, NativeMethods.GW_HWNDPREV));
        }
        finally
        {
            overlay?.Dispose();
            NativeMethods.DestroyWindow(popup);
            NativeMethods.DestroyWindow(owner);
        }
    }

    [DllImport("user32.dll", EntryPoint = "GetWindow")]
    private static extern IntPtr GetWindowRaw(IntPtr hWnd, uint cmd);

    private static long GetOwner(IntPtr hwnd) => GetWindowRaw(hwnd, 4 /* GW_OWNER */).ToInt64();
}
