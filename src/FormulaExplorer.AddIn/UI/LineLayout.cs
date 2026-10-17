using System.Windows;

namespace FormulaExplorer.AddIn.UI;

/// <summary>Measures runs of the displayed text. Production uses WPF <c>FormattedText</c>; tests use a fixed-width fake.</summary>
internal interface ITextMeasurer
{
    /// <summary>Advance width of <c>text[start, start+length)</c> laid out as one run (line terminators count as zero).</summary>
    double Measure(int start, int length);

    /// <summary>
    /// Cumulative x offsets of the same run: element k is the left edge of char <c>start+k</c> and the last element
    /// (index <paramref name="length"/>) the right edge of the last char. Non-decreasing; zero-width chars repeat the offset.
    /// </summary>
    double[] Offsets(int start, int length);
}

/// <summary>One laid-out line. It owns the chars <c>[Start, Start+Length)</c>, including a trailing '\n' (zero width).</summary>
/// <param name="Offsets">Length + 1 x offsets from the line's left edge (see <see cref="ITextMeasurer.Offsets"/>).</param>
internal sealed record TextLine(int Start, int Length, double[] Offsets)
{
    public int End => Start + Length;

    public double Width => Offsets[^1];
}

/// <summary>
/// The custom wrapping rule of the original Formula Explorer: a line may break after a space or after
/// <c>, ; + - * / ^ &amp; = &lt; &gt;</c> (even when a digit follows); '\n' always breaks; text is filled greedily;
/// a segment wider than a whole line is broken between characters. Trailing spaces may hang past the right edge.
/// </summary>
internal static class LineBreaker
{
    private const double Epsilon = 0.01;

    public static bool IsBreakAfter(char c) => c is ' ' or ',' or ';' or '+' or '-' or '*' or '/' or '^' or '&' or '=' or '<' or '>';

    /// <summary>
    /// Splits <paramref name="text"/> into lines as (start, length) pairs that partition the text.
    /// </summary>
    /// <param name="available">Line width; a line fits when its width without trailing spaces is ≤ this.</param>
    /// <param name="measure">Width of <c>text[start, start+length)</c> as one run.</param>
    public static List<(int Start, int Length)> Break(string text, double available, Func<int, int, double> measure)
    {
        var lines = new List<(int, int)>();
        if (text.Length == 0) return lines;

        int paragraphStart = 0;
        while (true)
        {
            int newline = text.IndexOf('\n', paragraphStart);
            int contentEnd = newline < 0 ? text.Length : newline;
            BreakParagraph(text, paragraphStart, contentEnd, newline < 0 ? 0 : 1, available, measure, lines);
            if (newline < 0) break;
            paragraphStart = newline + 1;
            if (paragraphStart == text.Length)
            {
                lines.Add((paragraphStart, 0)); // text ends with a line break: the caret line after it is empty
                break;
            }
        }
        return lines;
    }

    private static void BreakParagraph(string text, int start, int end, int terminatorLength, double available,
        Func<int, int, double> measure, List<(int, int)> lines)
    {
        bool Fits(int from, int to)
        {
            while (to > from && text[to - 1] == ' ') to--; // trailing spaces hang
            return to == from || measure(from, to - from) <= available + Epsilon;
        }

        int lineStart = start;
        int accepted = start; // end of the text accepted on the current line so far
        int segmentStart = start;
        while (segmentStart < end)
        {
            int segmentEnd = NextBreak(text, segmentStart, end);
            if (Fits(lineStart, segmentEnd))
            {
                accepted = segmentEnd;
                segmentStart = segmentEnd;
            }
            else if (accepted > lineStart)
            {
                lines.Add((lineStart, accepted - lineStart)); // wrap before this segment and retry it on a fresh line
                lineStart = accepted;
            }
            else
            {
                int fit = LargestFit(text, lineStart, segmentEnd, Fits);
                lines.Add((lineStart, fit - lineStart));
                lineStart = accepted = segmentStart = fit;
            }
        }

        if (lineStart < end || start == end)
            lines.Add((lineStart, end - lineStart + terminatorLength));
        else if (terminatorLength > 0)
            lines[^1] = (lines[^1].Item1, lines[^1].Item2 + terminatorLength); // the break belongs to the line that just filled up
    }

    /// <summary>End (exclusive) of the segment that starts at <paramref name="from"/>: just after the next break char.</summary>
    private static int NextBreak(string text, int from, int end)
    {
        for (int i = from; i < end; i++)
            if (IsBreakAfter(text[i])) return i + 1;
        return end;
    }

    /// <summary>Largest end in (from, limit] whose text still fits; at least one character so progress is guaranteed.</summary>
    private static int LargestFit(string text, int from, int limit, Func<int, int, bool> fits)
    {
        int low = from + 1, high = limit; // low always acceptable
        while (low < high)
        {
            int mid = low + (high - low + 1) / 2;
            if (fits(from, mid)) low = mid;
            else high = mid - 1;
        }
        // Never split a surrogate pair.
        if (low < limit && char.IsHighSurrogate(text[low - 1]) && char.IsLowSurrogate(text[low]))
            low = low - 1 > from ? low - 1 : low + 1;
        return low;
    }
}

/// <summary>
/// Lines of wrapped text with per-character geometry (no WPF text objects inside: everything comes from an
/// <see cref="ITextMeasurer"/>). Coordinates are relative to the text origin (top-left of the first line).
/// </summary>
internal sealed class TextLayout
{
    /// <summary>Fixed line pitch (and hover rectangle height) in DIPs.</summary>
    public const double LinePitch = 20;

    private readonly List<TextLine> _lines;

    private TextLayout(string text, List<TextLine> lines, double width)
    {
        Text = text;
        _lines = lines;
        Width = width;
    }

    public static TextLayout Empty { get; } = new(string.Empty, new List<TextLine>(), 0);

    public string Text { get; }

    /// <summary>The wrap width this layout was built for.</summary>
    public double Width { get; }

    public IReadOnlyList<TextLine> Lines => _lines;

    public double Height => _lines.Count * LinePitch;

    public static TextLayout Build(string text, double width, ITextMeasurer measurer)
    {
        var lines = LineBreaker.Break(text, width, measurer.Measure)
            .Select(l => new TextLine(l.Start, l.Length, measurer.Offsets(l.Start, l.Length)))
            .ToList();
        return new TextLayout(text, lines, width);
    }

    /// <summary>Index of the line that contains char <paramref name="charIndex"/>, or -1.</summary>
    public int LineOf(int charIndex)
    {
        if (charIndex < 0 || charIndex >= Text.Length) return -1;
        int low = 0, high = _lines.Count - 1;
        while (low <= high)
        {
            int mid = (low + high) / 2;
            var line = _lines[mid];
            if (charIndex < line.Start) high = mid - 1;
            else if (charIndex >= line.End) low = mid + 1;
            else return mid;
        }
        return -1;
    }

    /// <summary>Rectangle of one char (line height tall), or null when out of range or zero-width (line terminators).</summary>
    public Rect? CharRect(int charIndex)
    {
        int lineIndex = LineOf(charIndex);
        if (lineIndex < 0) return null;
        var line = _lines[lineIndex];
        double left = line.Offsets[charIndex - line.Start];
        double right = line.Offsets[charIndex - line.Start + 1];
        return right > left ? new Rect(left, lineIndex * LinePitch, right - left, LinePitch) : null;
    }

    /// <summary>
    /// The char under a point, or null: outside every char rectangle (right of a line's text, below the text, left of
    /// the origin) nothing is hit.
    /// </summary>
    public int? HitTest(Point point)
    {
        if (point.Y < 0 || point.X < 0) return null;
        int lineIndex = (int)(point.Y / LinePitch);
        if (lineIndex >= _lines.Count) return null;
        var line = _lines[lineIndex];
        if (point.X >= line.Width) return null;

        // Largest k with Offsets[k] <= x; zero-width chars share an offset with their successor so they are skipped.
        int low = 0, high = line.Length - 1, found = -1;
        while (low <= high)
        {
            int mid = (low + high) / 2;
            if (line.Offsets[mid] <= point.X)
            {
                found = mid;
                low = mid + 1;
            }
            else high = mid - 1;
        }
        if (found < 0 || line.Offsets[found + 1] <= point.X) return null;
        return line.Start + found;
    }

    /// <summary>
    /// Highlight rectangles of the span <c>[start, start+length)</c>: one per line it touches, line height tall;
    /// consecutive lines abut. Lines where the span has no width are skipped.
    /// </summary>
    public IEnumerable<Rect> SpanRects(int start, int length)
    {
        int end = start + length;
        if (length <= 0) yield break;
        int first = LineOf(Math.Max(start, 0));
        if (first < 0) first = start < 0 ? 0 : _lines.Count;
        for (int i = first; i < _lines.Count && _lines[i].Start < end; i++)
        {
            var line = _lines[i];
            int from = Math.Max(start, line.Start) - line.Start;
            int to = Math.Min(end, line.End) - line.Start;
            if (to <= from) continue;
            double left = line.Offsets[from];
            double right = line.Offsets[to];
            if (right > left) yield return new Rect(left, i * LinePitch, right - left, LinePitch);
        }
    }
}
