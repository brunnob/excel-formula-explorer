namespace FormulaExplorer.AddIn.Overlay;

/// <summary>
/// Pure clipping rules that turn the raw <c>Pane.VisibleRange</c> rectangles (which include the partly visible last
/// row/column) into the area where cells are really drawn.
/// </summary>
internal static class PaneClipping
{
    /// <summary>A cell area smaller than this fraction of the grid window in either direction is not believed.</summary>
    private const double MinPlausibleFraction = 0.5;

    /// <summary>
    /// The on-screen cell area: from the origin of the first pane, <paramref name="usableWidthPoints"/> wide and
    /// <paramref name="usableHeightPoints"/> tall (window sizes in points, so x dpi/72 and independent of the zoom),
    /// intersected with the EXCEL7 window. The EXCEL7 window also holds the scroll bars and the sheet-tab strip, which the
    /// partly visible last row/column would otherwise overlap. Falls back to <paramref name="gridRect"/> when the sizes
    /// are missing or implausible (smaller than half the grid window), because wrongly cutting borders is worse than
    /// drawing over a scroll bar.
    /// </summary>
    public static PixelRect CellArea(
        PixelRect gridRect, PixelRect firstPane, double usableWidthPoints, double usableHeightPoints, int dpi)
    {
        if (gridRect.IsEmpty || dpi <= 0
            || !(usableWidthPoints > 0) || !(usableHeightPoints > 0)
            || double.IsInfinity(usableWidthPoints) || double.IsInfinity(usableHeightPoints))
            return gridRect;

        var width = (int)Math.Round(usableWidthPoints * dpi / 72.0);
        var height = (int)Math.Round(usableHeightPoints * dpi / 72.0);
        var area = new PixelRect(firstPane.Left, firstPane.Top, firstPane.Left + width, firstPane.Top + height).Intersect(gridRect);
        if (area.IsEmpty
            || area.Width < gridRect.Width * MinPlausibleFraction
            || area.Height < gridRect.Height * MinPlausibleFraction)
            return gridRect;
        return area;
    }

    /// <summary>
    /// Clips every pane to <paramref name="cellArea"/> and to its neighbours. With a split (not frozen) window the top/left
    /// pane usually ends partway through a row/column, so its raw rectangle reaches into the pane below/right of it, which
    /// shows other cells: a pane ends where a pane that starts inside it (and overlaps it across) begins.
    /// Frozen panes already end on a boundary and are not changed by the neighbour rule. Empty panes stay empty.
    /// </summary>
    public static PixelRect[] Clip(IReadOnlyList<PixelRect> panes, PixelRect cellArea)
    {
        var clipped = new PixelRect[panes.Count];
        for (var i = 0; i < clipped.Length; i++) clipped[i] = panes[i].Intersect(cellArea);

        var result = new PixelRect[clipped.Length];
        for (var i = 0; i < clipped.Length; i++)
        {
            var p = clipped[i];
            if (p.IsEmpty)
            {
                result[i] = p;
                continue;
            }

            var right = p.Right;
            var bottom = p.Bottom;
            for (var j = 0; j < clipped.Length; j++)
            {
                var q = clipped[j];
                if (i == j || q.IsEmpty) continue;

                var overlapsAcross = Math.Min(p.Right, q.Right) > Math.Max(p.Left, q.Left);
                if (overlapsAcross && q.Top > p.Top && q.Top < bottom) bottom = q.Top;

                var overlapsDown = Math.Min(p.Bottom, q.Bottom) > Math.Max(p.Top, q.Top);
                if (overlapsDown && q.Left > p.Left && q.Left < right) right = q.Left;
            }

            result[i] = new PixelRect(p.Left, p.Top, right, bottom);
        }

        return result;
    }
}
