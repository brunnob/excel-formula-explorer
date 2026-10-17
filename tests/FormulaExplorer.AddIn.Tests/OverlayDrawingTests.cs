using FormulaExplorer.AddIn.Overlay;

namespace FormulaExplorer.AddIn.Tests;

public class OverlayDrawingTests
{
    [Theory]
    [InlineData(1.0, 2)]
    [InlineData(1.25, 2)]
    [InlineData(1.5, 2)]
    [InlineData(1.75, 3)]
    [InlineData(2.0, 3)]
    [InlineData(3.0, 5)]   // 4.5 rounds away from zero
    public void Thickness_scales_with_dpi(double scale, int expected) =>
        Assert.Equal(expected, OverlayDrawing.Thickness(scale));

    [Theory]
    [InlineData(1.0, 5)]
    [InlineData(1.25, 6)]
    [InlineData(1.5, 8)]
    [InlineData(2.0, 10)]
    [InlineData(0.5, 4)]   // never below 4
    public void HandleSize_scales_with_dpi(double scale, int expected) =>
        Assert.Equal(expected, OverlayDrawing.HandleSize(scale));

    [Theory]
    [InlineData(0.0)]
    [InlineData(-1.0)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public void Invalid_scale_falls_back_to_100_percent(double scale)
    {
        Assert.Equal(2, OverlayDrawing.Thickness(scale));
        Assert.Equal(5, OverlayDrawing.HandleSize(scale));
    }

    [Fact]
    public void Thickness_is_at_least_one()
    {
        Assert.Equal(1, OverlayDrawing.Thickness(0.1));
    }

    private static List<PixelRect> Outline(PixelRect area, PixelRect viewport, int t = 2, int h = 5)
    {
        var list = new List<PixelRect>();
        OverlayDrawing.BuildOutline(area, viewport, t, h, list);
        return list;
    }

    [Fact]
    public void Box_fully_inside_has_four_edges_and_four_handles()
    {
        var viewport = new PixelRect(0, 0, 1000, 800);
        var area = new PixelRect(100, 100, 200, 140);

        var pieces = Outline(area, viewport);

        Assert.Equal(8, pieces.Count);
        // t = 2: lines cover [edge-1, edge+1); horizontal edges span the corners.
        Assert.Contains(new PixelRect(99, 99, 201, 101), pieces);   // top
        Assert.Contains(new PixelRect(99, 139, 201, 141), pieces);  // bottom
        Assert.Contains(new PixelRect(99, 99, 101, 141), pieces);   // left
        Assert.Contains(new PixelRect(199, 99, 201, 141), pieces);  // right
        // h = 5: [c-2, c+3) around each corner.
        Assert.Contains(new PixelRect(98, 98, 103, 103), pieces);
        Assert.Contains(new PixelRect(198, 98, 203, 103), pieces);
        Assert.Contains(new PixelRect(98, 138, 103, 143), pieces);
        Assert.Contains(new PixelRect(198, 138, 203, 143), pieces);
    }

    [Fact]
    public void Odd_thickness_puts_the_extra_pixel_inside()
    {
        var pieces = Outline(new PixelRect(100, 100, 200, 140), new PixelRect(0, 0, 1000, 800), t: 3, h: 4);

        // outer = 1, inner = 2: [edge-1, edge+2)
        Assert.Contains(new PixelRect(99, 99, 202, 102), pieces);
        Assert.Contains(new PixelRect(99, 99, 102, 142), pieces);
    }

    [Fact]
    public void Edges_and_handles_are_clipped_to_the_viewport()
    {
        var viewport = new PixelRect(100, 100, 300, 300);
        // Box touches the viewport's top-left corner: outer halves and handle quarters are cut off.
        var pieces = Outline(new PixelRect(100, 100, 200, 200), viewport);

        Assert.All(pieces, p => Assert.Equal(p, p.Intersect(viewport)));
        Assert.Contains(new PixelRect(100, 100, 201, 101), pieces);   // top edge, left overhang clipped
        Assert.Contains(new PixelRect(100, 100, 103, 103), pieces);   // top-left handle clipped to 3x3
    }

    [Fact]
    public void Side_cut_by_the_viewport_gets_no_border_and_no_handles()
    {
        var viewport = new PixelRect(0, 0, 150, 800);
        // Right boundary (x = 200) is beyond the viewport: no right edge, no right handles.
        var pieces = Outline(new PixelRect(100, 100, 200, 140), viewport);

        Assert.Equal(2 /* top, bottom */ + 1 /* left */ + 2 /* left handles */, pieces.Count);
        Assert.All(pieces, p => Assert.True(p.Right <= 150));
        Assert.DoesNotContain(pieces, p => p.Left >= 190);
    }

    [Fact]
    public void Top_cut_removes_top_edge_and_top_handles_but_keeps_vertical_edges()
    {
        var viewport = new PixelRect(0, 120, 1000, 800);
        var pieces = Outline(new PixelRect(100, 100, 200, 140), viewport);

        // bottom edge + left + right edges + two bottom handles
        Assert.Equal(5, pieces.Count);
        Assert.Contains(new PixelRect(99, 120, 101, 141), pieces);   // left edge clipped at the viewport top
        Assert.Contains(new PixelRect(199, 120, 201, 141), pieces);
        Assert.DoesNotContain(pieces, p => p.Top < 120);
    }

    [Fact]
    public void Boundary_inclusion_rules_at_the_viewport_edges()
    {
        var viewport = new PixelRect(100, 100, 200, 200);

        // Box whose bottom boundary is exactly the viewport bottom: bottom edge is drawn (half clipped).
        var bottomFlush = Outline(new PixelRect(120, 120, 180, 200), viewport);
        Assert.Contains(new PixelRect(119, 199, 181, 200), bottomFlush);

        // Box that starts exactly at the viewport bottom: its top boundary is outside, so nothing shows at all.
        Assert.Empty(Outline(new PixelRect(120, 200, 180, 240), viewport));

        // Box that ends exactly at the viewport top: bottom boundary is outside (vp.Top < edge fails).
        Assert.Empty(Outline(new PixelRect(120, 60, 180, 100), viewport));
    }

    [Fact]
    public void Box_entirely_outside_produces_nothing()
    {
        Assert.Empty(Outline(new PixelRect(500, 500, 600, 600), new PixelRect(0, 0, 100, 100)));
    }

    [Fact]
    public void Box_larger_than_the_viewport_has_no_edges_and_no_handles()
    {
        // Whole-column style range: every boundary is outside the viewport, so the pane stays clean.
        Assert.Empty(Outline(new PixelRect(-5000, -5000, 5000, 5000), new PixelRect(0, 0, 100, 100)));
    }

    [Fact]
    public void Empty_inputs_produce_nothing()
    {
        Assert.Empty(Outline(PixelRect.Empty, new PixelRect(0, 0, 100, 100)));
        Assert.Empty(Outline(new PixelRect(10, 10, 20, 20), PixelRect.Empty));
    }

    [Fact]
    public void Tiny_box_still_gets_edges_and_handles()
    {
        var pieces = Outline(new PixelRect(50, 50, 52, 52), new PixelRect(0, 0, 100, 100));
        Assert.Equal(8, pieces.Count);
    }
}

public class PixelRectTests
{
    [Fact]
    public void Intersect_returns_overlap()
    {
        var r = new PixelRect(0, 0, 10, 10).Intersect(new PixelRect(5, 5, 20, 20));
        Assert.Equal(new PixelRect(5, 5, 10, 10), r);
    }

    [Fact]
    public void Touching_rectangles_do_not_intersect()
    {
        Assert.True(new PixelRect(0, 0, 10, 10).Intersect(new PixelRect(10, 0, 20, 10)).IsEmpty);
        Assert.True(new PixelRect(0, 0, 10, 10).Intersect(new PixelRect(0, 10, 10, 20)).IsEmpty);
    }

    [Fact]
    public void Empty_detection_and_size()
    {
        Assert.True(PixelRect.Empty.IsEmpty);
        Assert.True(new PixelRect(5, 5, 5, 9).IsEmpty);
        Assert.True(new PixelRect(5, 5, 3, 9).IsEmpty);
        var r = new PixelRect(2, 3, 12, 8);
        Assert.False(r.IsEmpty);
        Assert.Equal(10, r.Width);
        Assert.Equal(5, r.Height);
    }

    [Fact]
    public void Offset_moves_both_corners() =>
        Assert.Equal(new PixelRect(-8, -7, 2, -2), new PixelRect(2, 3, 12, 8).Offset(-10, -10));
}

public class PointConversionTests
{
    // 1 point = 2 pixels, origin at 100 px.
    private static int Linear(int points) => 100 + 2 * points;

    [Fact]
    public void Whole_points_are_converted_directly() =>
        Assert.Equal(Linear(30), PointConversion.ToPixels(30, Linear));

    [Fact]
    public void Fractional_points_are_interpolated_not_truncated()
    {
        Assert.Equal(Linear(30) + 1, PointConversion.ToPixels(30.5, Linear));
        Assert.Equal(Linear(14) + 1, PointConversion.ToPixels(14.25, Linear));   // 0.5 px rounds away from zero
    }

    [Fact]
    public void Negative_points_use_floor()
    {
        Assert.Equal(Linear(-2), PointConversion.ToPixels(-2.0, Linear));
        Assert.Equal(Linear(-3) + 1, PointConversion.ToPixels(-2.5, Linear));
    }

    [Fact]
    public void Only_one_conversion_call_for_whole_values()
    {
        var calls = 0;
        PointConversion.ToPixels(12.0, p => { calls++; return p; });
        Assert.Equal(1, calls);
    }

    [Fact]
    public void Non_finite_input_is_zero() =>
        Assert.Equal(0, PointConversion.ToPixels(double.NaN, Linear));

    [Fact]
    public void Huge_results_are_clamped()
    {
        var px = PointConversion.ToPixels(1e12, p => p);
        Assert.InRange(px, -PointConversion.PixelLimit, PointConversion.PixelLimit);
    }
}
