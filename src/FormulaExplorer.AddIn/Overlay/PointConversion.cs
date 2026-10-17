namespace FormulaExplorer.AddIn.Overlay;

/// <summary>
/// Excel's <c>Pane.PointsToScreenPixelsX/Y</c> take a whole number of points (<c>Long</c>), while worksheet geometry
/// (row heights such as 14.25 pt) is fractional. This helper converts a fractional point value by linear interpolation
/// between the pixel positions of the two surrounding whole points, instead of truncating (which would be off by up to
/// one point = 1.33 px x zoom x DPI).
/// </summary>
internal static class PointConversion
{
    /// <summary>Largest magnitude (in pixels) handed to the drawing code; keeps all later int arithmetic overflow-free.</summary>
    public const int PixelLimit = 100_000_000;

    public static int ToPixels(double points, Func<int, int> convertWholePoints)
    {
        if (!double.IsFinite(points)) return 0;
        points = Math.Clamp(points, -1_000_000_000d, 1_000_000_000d);
        var floor = (int)Math.Floor(points);
        var fraction = points - floor;
        var p0 = convertWholePoints(floor);
        if (fraction < 1e-6) return Clamp(p0);
        var p1 = convertWholePoints(floor + 1);
        return Clamp((int)Math.Round(p0 + fraction * (p1 - p0), MidpointRounding.AwayFromZero));
    }

    private static int Clamp(int pixels) => Math.Clamp(pixels, -PixelLimit, PixelLimit);
}
