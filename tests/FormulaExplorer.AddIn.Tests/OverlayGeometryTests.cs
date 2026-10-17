using FormulaExplorer.AddIn.Overlay;
using FormulaExplorer.AddIn.Session;

namespace FormulaExplorer.AddIn.Tests;

public class OverlayGeometryTests
{
    /// <summary>A pane with a fixed viewport and a trivial linear points-to-pixels mapping (scroll offset and zoom).</summary>
    private sealed class FakePane : IPaneMetrics
    {
        private readonly double _scale;
        private readonly double _originX;
        private readonly double _originY;
        private readonly int _screenX;
        private readonly int _screenY;

        public FakePane(PixelRect visible, double scale = 1.0, double scrollPointsX = 0, double scrollPointsY = 0)
        {
            VisibleRect = visible;
            _scale = scale;
            _originX = scrollPointsX;
            _originY = scrollPointsY;
            _screenX = visible.Left;
            _screenY = visible.Top;
        }

        public PixelRect VisibleRect { get; }

        public PixelRect ToScreen(PointRect a) => new(
            _screenX + (int)Math.Round((a.Left - _originX) * _scale),
            _screenY + (int)Math.Round((a.Top - _originY) * _scale),
            _screenX + (int)Math.Round((a.Left + a.Width - _originX) * _scale),
            _screenY + (int)Math.Round((a.Top + a.Height - _originY) * _scale));
    }

    private static readonly PixelRect Grid = new(100, 200, 1100, 800);

    private static BoxAreas Box(uint rgb, params PointRect[] areas) => new(rgb, areas);

    [Fact]
    public void Single_pane_single_box_gives_outline_and_bounds()
    {
        var pane = new FakePane(new PixelRect(100, 200, 1100, 800));
        var box = Box(0x2E75B6, new PointRect(50, 40, 64, 20));

        var frame = GridGeometry.Compute([pane], Grid, [box], 1.0);

        Assert.Equal(8, frame.Fills.Count);
        Assert.All(frame.Fills, f => Assert.Equal(0x2E75B6u, f.Rgb));
        var bounds = Assert.Single(frame.Bounds);
        Assert.Equal(new OverlayRect(150, 240, 214, 260, 0x2E75B6), bounds);
    }

    [Fact]
    public void Pane_is_clipped_to_the_grid_window()
    {
        // The pane (partly visible last row/column) sticks out past EXCEL7; the box must not either.
        var pane = new FakePane(new PixelRect(100, 200, 1300, 1000));
        var box = Box(1, new PointRect(900, 590, 400, 100));   // crosses right and bottom edge of Grid

        var frame = GridGeometry.Compute([pane], Grid, [box], 1.0);

        Assert.All(frame.Fills, f =>
        {
            Assert.True(f.Rect.Right <= Grid.Right);
            Assert.True(f.Rect.Bottom <= Grid.Bottom);
        });
        var bounds = Assert.Single(frame.Bounds);
        Assert.Equal(1100, bounds.Right);
        Assert.Equal(800, bounds.Bottom);
        // Cut on the right and bottom: left+top edges and the top-left handle only.
        Assert.Equal(3, frame.Fills.Count);
    }

    [Fact]
    public void Pane_narrower_than_the_grid_window_keeps_borders_off_the_scroll_bar_strip()
    {
        // The cell area ends at y = 700; the EXCEL7 rect continues to 800 (tab strip and scroll bar).
        var pane = new FakePane(new PixelRect(100, 200, 1100, 700));
        var box = Box(1, new PointRect(50, 400, 64, 400));   // runs down to y = 800 on screen

        var frame = GridGeometry.Compute([pane], Grid, [box], 1.0);

        Assert.All(frame.Fills, f => Assert.True(f.Rect.Bottom <= 700));
        Assert.Equal(700, Assert.Single(frame.Bounds).Bottom);
    }

    [Fact]
    public void Box_off_screen_yields_nothing()
    {
        var pane = new FakePane(new PixelRect(100, 200, 1100, 800), scrollPointsY: 5000);
        var frame = GridGeometry.Compute([pane], Grid, [Box(1, new PointRect(0, 0, 50, 15))], 1.0);

        Assert.Empty(frame.Fills);
        Assert.Empty(frame.Bounds);
        Assert.Same(OverlayFrame.Empty, frame);
    }

    [Fact]
    public void Hidden_row_area_with_zero_height_is_skipped()
    {
        var pane = new FakePane(new PixelRect(100, 200, 1100, 800));
        var frame = GridGeometry.Compute([pane], Grid,
            [Box(1, new PointRect(10, 10, 50, 0)), Box(2, new PointRect(10, 10, 0, 15))], 1.0);

        Assert.Empty(frame.Fills);
    }

    [Fact]
    public void Area_that_collapses_to_zero_pixels_is_skipped()
    {
        var pane = new FakePane(new PixelRect(100, 200, 1100, 800), scale: 0.01);
        var frame = GridGeometry.Compute([pane], Grid, [Box(1, new PointRect(10, 10, 20, 15))], 1.0);

        Assert.Empty(frame.Fills);
    }

    [Fact]
    public void Zoom_scales_positions()
    {
        var pane = new FakePane(new PixelRect(100, 200, 1100, 800), scale: 2.0);
        var frame = GridGeometry.Compute([pane], Grid, [Box(1, new PointRect(50, 40, 64, 20))], 1.0);

        Assert.Equal(new OverlayRect(200, 280, 328, 320, 1), Assert.Single(frame.Bounds));
    }

    [Fact]
    public void Multi_area_box_yields_one_bound_per_area_in_order()
    {
        var pane = new FakePane(new PixelRect(100, 200, 1100, 800));
        var frame = GridGeometry.Compute([pane], Grid,
            [Box(7, new PointRect(10, 10, 30, 15), new PointRect(200, 100, 30, 15))], 1.0);

        Assert.Equal(2, frame.Bounds.Count);
        Assert.Equal(110, frame.Bounds[0].Left);
        Assert.Equal(300, frame.Bounds[1].Left);
        Assert.Equal(16, frame.Fills.Count);
    }

    [Fact]
    public void Boxes_are_painted_in_list_order()
    {
        var pane = new FakePane(new PixelRect(100, 200, 1100, 800));
        var frame = GridGeometry.Compute([pane], Grid,
            [Box(0xAAAAAA, new PointRect(10, 10, 30, 15)), Box(0xBBBBBB, new PointRect(10, 10, 30, 15))], 1.0);

        Assert.Equal(16, frame.Fills.Count);
        Assert.All(frame.Fills.Take(8), f => Assert.Equal(0xAAAAAAu, f.Rgb));
        Assert.All(frame.Fills.Skip(8), f => Assert.Equal(0xBBBBBBu, f.Rgb));
    }

    [Fact]
    public void Frozen_panes_draw_a_box_straddling_the_split_in_both_panes_without_a_cut_border()
    {
        // Two panes side by side: frozen columns [100,300) and the scrolling part [300,1100).
        var frozen = new FakePane(new PixelRect(100, 200, 300, 800));
        var scrolling = new FakePane(new PixelRect(300, 200, 1100, 800), scrollPointsX: -200);   // origin so that x = 300 + 200 + p
        // A box from sheet x=100..180 pts: in the frozen pane it maps to 200..280 (fully inside),
        // in the scrolling pane to 300+100+200=600..680.
        var box = Box(1, new PointRect(100, 0, 80, 20));

        var frame = GridGeometry.Compute([frozen, scrolling], Grid, [box], 1.0);

        Assert.Equal(2, frame.Bounds.Count);
        Assert.Equal(200, frame.Bounds[0].Left);
        Assert.Equal(600, frame.Bounds[1].Left);
    }

    [Fact]
    public void Box_cut_by_a_pane_boundary_has_no_edge_on_the_cut_side_in_that_pane()
    {
        var left = new FakePane(new PixelRect(100, 200, 300, 800));
        // In the left pane the box spans x = 250..350: its right boundary (350) is beyond the pane (300).
        var box = Box(1, new PointRect(150, 10, 100, 20));

        var frame = GridGeometry.Compute([left], Grid, [box], 1.0);

        Assert.DoesNotContain(frame.Fills, f => f.Rect.Left >= 300);
        Assert.Equal(300, Assert.Single(frame.Bounds).Right);
        // top, bottom, left edges + two left handles
        Assert.Equal(5, frame.Fills.Count);
    }

    [Fact]
    public void Pane_outside_the_grid_is_ignored()
    {
        var pane = new FakePane(new PixelRect(2000, 2000, 2500, 2500));
        Assert.Empty(GridGeometry.Compute([pane], Grid, [Box(1, new PointRect(0, 0, 50, 15))], 1.0).Fills);
    }

    [Fact]
    public void Dpi_scale_changes_thickness_and_handles()
    {
        var pane = new FakePane(new PixelRect(100, 200, 1100, 800));
        var box = Box(1, new PointRect(50, 40, 64, 20));

        var at100 = GridGeometry.Compute([pane], Grid, [box], 1.0);
        var at200 = GridGeometry.Compute([pane], Grid, [box], 2.0);

        Assert.Equal(2, at100.Fills.Min(f => f.Rect.Height));
        Assert.Equal(3, at200.Fills.Min(f => Math.Min(f.Rect.Height, f.Rect.Width)));
        Assert.Equal(10, at200.Fills.Max(f => f.Rect.Width == f.Rect.Height ? f.Rect.Width : 0));
    }
}
