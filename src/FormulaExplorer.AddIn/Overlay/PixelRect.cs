namespace FormulaExplorer.AddIn.Overlay;

/// <summary>
/// Half-open pixel rectangle [Left, Right) x [Top, Bottom) in physical screen pixels (or window-local pixels where noted).
/// A rectangle with Right &lt;= Left or Bottom &lt;= Top is empty.
/// </summary>
internal readonly record struct PixelRect(int Left, int Top, int Right, int Bottom)
{
    public static PixelRect Empty => default;

    public int Width => Right - Left;
    public int Height => Bottom - Top;
    public bool IsEmpty => Right <= Left || Bottom <= Top;

    /// <summary>Intersection with <paramref name="other"/>; <see cref="Empty"/> when they do not overlap.</summary>
    public PixelRect Intersect(PixelRect other)
    {
        var l = Math.Max(Left, other.Left);
        var t = Math.Max(Top, other.Top);
        var r = Math.Min(Right, other.Right);
        var b = Math.Min(Bottom, other.Bottom);
        return r <= l || b <= t ? Empty : new PixelRect(l, t, r, b);
    }

    public PixelRect Offset(int dx, int dy) => new(Left + dx, Top + dy, Right + dx, Bottom + dy);
}

/// <summary>A filled rectangle with an 0xRRGGBB color (screen pixels unless stated otherwise).</summary>
internal readonly record struct ColoredRect(PixelRect Rect, uint Rgb);

/// <summary>A rectangle in worksheet coordinates (points), as Excel's Range.Left/Top/Width/Height report it.</summary>
internal readonly record struct PointRect(double Left, double Top, double Width, double Height);
