using System.Globalization;
using System.Windows;
using System.Windows.Media;

namespace FormulaExplorer.AddIn.UI;

/// <summary>
/// <see cref="ITextMeasurer"/> on WPF <c>FormattedText</c> in <c>TextFormattingMode.Display</c> (GDI-compatible
/// advances, matching the original's WinForms text). Line terminators ('\r', '\n') are removed before formatting and
/// measure as zero width; tabs are formatted as spaces by the caller (the text it is given has none).
/// </summary>
internal sealed class FormattedTextMeasurer : ITextMeasurer
{
    private readonly string _text;
    private readonly Typeface _typeface;
    private readonly double _emSize;
    private readonly double _pixelsPerDip;

    public FormattedTextMeasurer(string text, Typeface typeface, double emSize, double pixelsPerDip)
    {
        _text = text;
        _typeface = typeface;
        _emSize = emSize;
        _pixelsPerDip = pixelsPerDip;
    }

    public double Measure(int start, int length)
    {
        string stripped = Strip(start, length, out _);
        return stripped.Length == 0 ? 0 : Create(stripped, Brushes.Black).WidthIncludingTrailingWhitespace;
    }

    public double[] Offsets(int start, int length)
    {
        var offsets = new double[length + 1];
        string stripped = Strip(start, length, out int[] map);
        if (stripped.Length == 0) return offsets;

        var formatted = Create(stripped, Brushes.Black);
        var origin = new Point(0, 0);
        double previous = 0;
        for (int k = 0; k < length; k++)
        {
            int s = map[k];
            if (map[k + 1] > s)
            {
                // The right edge of the glyph box: neighbours share edges, so offsets stay contiguous.
                double right = formatted.BuildHighlightGeometry(origin, s, 1)?.Bounds.Right ?? previous;
                previous = Math.Max(previous, right);
            }
            offsets[k + 1] = previous;
        }
        return offsets;
    }

    /// <summary>The formatted text of a whole line with its colors applied.</summary>
    /// <param name="colorRuns">Runs as (start index relative to <paramref name="start"/>, length, brush) over the original chars.</param>
    public FormattedText CreateLine(int start, int length, IEnumerable<(int Start, int Length, Brush Brush)> colorRuns)
    {
        string stripped = Strip(start, length, out int[] map);
        var formatted = Create(stripped, Brushes.Black);
        foreach (var (runStart, runLength, brush) in colorRuns)
        {
            int from = map[runStart];
            int to = map[runStart + runLength];
            if (to > from) formatted.SetForegroundBrush(brush, from, to - from);
        }
        return formatted;
    }

    private FormattedText Create(string text, Brush brush) =>
        new(text, CultureInfo.InvariantCulture, FlowDirection.LeftToRight, _typeface, _emSize, brush,
            null, TextFormattingMode.Display, _pixelsPerDip);

    /// <summary>
    /// <c>_text[start, start+length)</c> without '\r'/'\n'. <paramref name="map"/>[k] = index in the result of original
    /// char k (length + 1 entries; a removed char maps to the next kept char's index).
    /// </summary>
    private string Strip(int start, int length, out int[] map)
    {
        map = new int[length + 1];
        var builder = new System.Text.StringBuilder(length);
        for (int k = 0; k < length; k++)
        {
            map[k] = builder.Length;
            char c = _text[start + k];
            if (c is not ('\r' or '\n')) builder.Append(c);
        }
        map[length] = builder.Length;
        return builder.ToString();
    }
}
