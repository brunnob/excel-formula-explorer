using FormulaExplorer.AddIn.Session;

namespace FormulaExplorer.AddIn.Overlay;

/// <summary>The areas (one per range area; multi-area ranges have several) of one precedent box, in worksheet points.</summary>
internal sealed record BoxAreas(uint Rgb, IReadOnlyList<PointRect> Areas);

/// <summary>What the overlay must draw for one view of the sheet, in physical screen pixels.</summary>
/// <param name="Fills">Rectangles to fill, in paint order (later on top).</param>
/// <param name="Bounds">The range boundary rectangle of every visible box area, clipped to its pane (diagnostics/tests).</param>
internal sealed record OverlayFrame(IReadOnlyList<ColoredRect> Fills, IReadOnlyList<OverlayRect> Bounds)
{
    public static OverlayFrame Empty { get; } = new(Array.Empty<ColoredRect>(), Array.Empty<OverlayRect>());
}

/// <summary>
/// Pure geometry: precedent areas + pane metrics + grid window rect -> the rectangles to paint.
/// All COM-dependent conversions are hidden behind <see cref="IPaneMetrics"/>.
/// </summary>
internal static class GridGeometry
{
    /// <param name="panes">Panes of the active window (1, 2 or 4).</param>
    /// <param name="gridRect">The EXCEL7 window rect in screen pixels; every pane is clipped to it.</param>
    /// <param name="boxes">Boxes in draw order (later on top).</param>
    /// <param name="dpiScale">Monitor DPI / 96; scales border thickness and handle size.</param>
    public static OverlayFrame Compute(
        IReadOnlyList<IPaneMetrics> panes, PixelRect gridRect, IReadOnlyList<BoxAreas> boxes, double dpiScale)
    {
        var thickness = OverlayDrawing.Thickness(dpiScale);
        var handle = OverlayDrawing.HandleSize(dpiScale);
        var fills = new List<ColoredRect>();
        var bounds = new List<OverlayRect>();
        var pieces = new List<PixelRect>();

        foreach (var pane in panes)
        {
            var viewport = pane.VisibleRect.Intersect(gridRect);
            if (viewport.IsEmpty) continue;

            foreach (var box in boxes)
            {
                foreach (var points in box.Areas)
                {
                    // Hidden rows/columns give a zero-sized area: nothing to outline.
                    if (!(points.Width > 0) || !(points.Height > 0)) continue;

                    var area = pane.ToScreen(points);
                    if (area.IsEmpty) continue;

                    pieces.Clear();
                    OverlayDrawing.BuildOutline(area, viewport, thickness, handle, pieces);
                    if (pieces.Count == 0) continue;

                    foreach (var piece in pieces) fills.Add(new ColoredRect(piece, box.Rgb));

                    var visible = area.Intersect(viewport);
                    if (!visible.IsEmpty)
                        bounds.Add(new OverlayRect(visible.Left, visible.Top, visible.Right, visible.Bottom, box.Rgb));
                }
            }
        }

        return fills.Count == 0 ? OverlayFrame.Empty : new OverlayFrame(fills, bounds);
    }
}
