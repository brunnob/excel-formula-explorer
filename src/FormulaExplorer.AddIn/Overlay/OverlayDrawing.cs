namespace FormulaExplorer.AddIn.Overlay;

/// <summary>
/// Pure drawing math of the overlay: border thickness, handle size and the splitting of a range outline into the
/// pixel-exact rectangles (edge strips and corner handles) that WM_PAINT fills.
/// </summary>
internal static class OverlayDrawing
{
    /// <summary>Border thickness in physical pixels: <c>max(1, round(1.5 x dpiScale))</c> (2 px at 100%, 3 px at 200%).</summary>
    public static int Thickness(double dpiScale) =>
        Math.Max(1, (int)Math.Round(1.5 * Sanitize(dpiScale), MidpointRounding.AwayFromZero));

    /// <summary>Side of the square corner handles: <c>max(4, round(5 x dpiScale))</c>.</summary>
    public static int HandleSize(double dpiScale) =>
        Math.Max(4, (int)Math.Round(5 * Sanitize(dpiScale), MidpointRounding.AwayFromZero));

    private static double Sanitize(double dpiScale) => double.IsFinite(dpiScale) && dpiScale > 0 ? dpiScale : 1.0;

    /// <summary>
    /// Splits the outline of <paramref name="area"/> (the range boundary in screen pixels, NOT yet clipped) into
    /// rectangles clipped to <paramref name="viewport"/> (a pane's visible rectangle), appended to <paramref name="output"/>.
    /// <list type="bullet">
    /// <item>The outline is centered on the range boundary: a line of thickness <c>t</c> covers
    /// <c>[edge - t/2, edge - t/2 + t)</c>, i.e. half outside, half inside (for odd <c>t</c> the extra pixel is inside).</item>
    /// <item>An edge is emitted only when its boundary line lies inside the viewport (top/left: <c>vp.Start &lt;= edge &lt; vp.End</c>,
    /// bottom/right: <c>vp.Start &lt; edge &lt;= vp.End</c>). A side cut by the viewport therefore gets no border.</item>
    /// <item>A corner handle (square of side <paramref name="handleSize"/> centered on the corner) is emitted only when both
    /// boundary lines meeting at that corner are inside the viewport by the same rules.</item>
    /// </list>
    /// Edges come first, handles last (so handles paint over edge ends). Everything is clipped to the viewport.
    /// </summary>
    public static void BuildOutline(PixelRect area, PixelRect viewport, int thickness, int handleSize, List<PixelRect> output)
    {
        if (area.IsEmpty || viewport.IsEmpty) return;

        var outer = thickness / 2;
        var inner = thickness - outer;
        var l = area.Left;
        var t = area.Top;
        var r = area.Right;
        var b = area.Bottom;

        // Horizontal edges span from the outer side of the left line to the inner side of the right line so that corners are filled.
        var spanLeft = l - outer;
        var spanRight = r + inner;
        var spanTop = t - outer;
        var spanBottom = b + inner;

        // Is each boundary line inside the viewport? (start side inclusive, end side exclusive for top/left;
        // the opposite for bottom/right, because the line sits between two pixel rows/columns.)
        var topIn = viewport.Top <= t && t < viewport.Bottom;
        var bottomIn = viewport.Top < b && b <= viewport.Bottom;
        var leftIn = viewport.Left <= l && l < viewport.Right;
        var rightIn = viewport.Left < r && r <= viewport.Right;
        // A line is only drawn when the box also reaches into the viewport along the line's own direction,
        // so a box merely touching the viewport edge leaves no half-thickness sliver behind.
        var overlapsHorizontally = l < viewport.Right && r > viewport.Left;
        var overlapsVertically = t < viewport.Bottom && b > viewport.Top;

        if (topIn && overlapsHorizontally)
            Add(output, new PixelRect(spanLeft, t - outer, spanRight, t + inner), viewport);
        if (bottomIn && overlapsHorizontally)
            Add(output, new PixelRect(spanLeft, b - outer, spanRight, b + inner), viewport);
        if (leftIn && overlapsVertically)
            Add(output, new PixelRect(l - outer, spanTop, l + inner, spanBottom), viewport);
        if (rightIn && overlapsVertically)
            Add(output, new PixelRect(r - outer, spanTop, r + inner, spanBottom), viewport);

        var half = handleSize / 2;
        if (leftIn && topIn) AddHandle(output, l, t, half, handleSize, viewport);
        if (rightIn && topIn) AddHandle(output, r, t, half, handleSize, viewport);
        if (leftIn && bottomIn) AddHandle(output, l, b, half, handleSize, viewport);
        if (rightIn && bottomIn) AddHandle(output, r, b, half, handleSize, viewport);
    }

    private static void AddHandle(List<PixelRect> output, int cx, int cy, int half, int size, PixelRect viewport) =>
        Add(output, new PixelRect(cx - half, cy - half, cx - half + size, cy - half + size), viewport);

    private static void Add(List<PixelRect> output, PixelRect rect, PixelRect viewport)
    {
        var clipped = rect.Intersect(viewport);
        if (!clipped.IsEmpty) output.Add(clipped);
    }
}
