using FormulaExplorer.AddIn.UI;

namespace FormulaExplorer.AddIn.Tests.UI;

public class LineBreakerTests
{
    /// <summary>Breaks with fixed 10-unit chars; <paramref name="columns"/> = how many chars fit in a line.</summary>
    private static List<string> Lines(string text, int columns)
    {
        var measurer = new FixedMeasurer(text);
        return LineBreaker.Break(text, columns * 10, measurer.Measure).Select(l => text.Substring(l.Start, l.Length)).ToList();
    }

    [Fact]
    public void EmptyTextHasNoLines() => Assert.Empty(Lines("", 10));

    [Fact]
    public void ShortTextIsOneLine() => Assert.Equal(new[] { "=A1+B1" }, Lines("=A1+B1", 10));

    [Theory]
    [InlineData(' ')]
    [InlineData(',')]
    [InlineData(';')]
    [InlineData('+')]
    [InlineData('-')]
    [InlineData('*')]
    [InlineData('/')]
    [InlineData('^')]
    [InlineData('&')]
    [InlineData('=')]
    [InlineData('<')]
    [InlineData('>')]
    public void BreaksAfterEachBreakCharacter(char c)
    {
        // "aaaa" + c fits in 5 columns; the following "bbbb" does not fit on the same line.
        string text = "aaaa" + c + "bbbb";
        Assert.Equal(new[] { "aaaa" + c, "bbbb" }, Lines(text, 5));
    }

    [Theory]
    [InlineData('(')]
    [InlineData(')')]
    [InlineData('$')]
    [InlineData(':')]
    [InlineData('!')]
    [InlineData('.')]
    [InlineData('\'')]
    [InlineData('"')]
    [InlineData('{')]
    [InlineData('%')]
    public void DoesNotBreakAfterOtherCharacters(char c)
    {
        // The whole 9-char word is wider than a 5-column line and has no break opportunity: it is split per character,
        // so the first line is exactly full (never "aaaa" + c on one line and the rest on the next because of c).
        string text = "aaaa" + c + "bbbb";
        Assert.Equal(new[] { "aaaa" + c, "bbbb" }, Lines(text, 5));
        // ...and with room for 6 columns the break is not at c either.
        Assert.Equal(new[] { "aaaa" + c + "b", "bbb" }, Lines(text, 6));
    }

    [Fact]
    public void BreaksAfterPlusAndCommaEvenBeforeADigit()
    {
        Assert.Equal(new[] { "A1+", "1,2" }, Lines("A1+1,2", 3));
        Assert.Equal(new[] { "=X+", "100", "+1" }, Lines("=X+100+1", 3));
    }

    [Fact]
    public void FillsGreedily()
    {
        // "SUM(" (4) + "A1," (3) fit in 8 columns; "B1," would make 10.
        Assert.Equal(new[] { "SUM(A1,", "B1,C1)" }, Lines("SUM(A1,B1,C1)", 8));
    }

    [Fact]
    public void TrailingSpacesHangPastTheEdge()
    {
        // 3 columns: "ab" + 3 spaces still fits because trailing spaces do not count.
        Assert.Equal(new[] { "ab   ", "cd" }, Lines("ab   cd", 3));
    }

    [Fact]
    public void WideSegmentIsBrokenPerCharacter()
    {
        Assert.Equal(new[] { "abcd", "efgh", "ij" }, Lines("abcdefghij", 4));
    }

    [Fact]
    public void RemainderOfAWideSegmentContinuesWithTheNextSegments()
    {
        // "abcdefgh+" has no earlier break, so it is cut at 5 columns; "fgh+" then restarts a normal line.
        Assert.Equal(new[] { "abcde", "fgh+", "ij" }, Lines("abcdefgh+ij", 5));
    }

    [Fact]
    public void ZeroWidthAvailableStillMakesProgress()
    {
        Assert.Equal(new[] { "a", "b", "c" }, Lines("abc", 0));
    }

    [Fact]
    public void NewlineAfterAFullLineStaysOnThatLine()
    {
        Assert.Equal(new[] { "abc\n", "d" }, Lines("abc\nd", 3));
        Assert.Equal(new[] { "a", "b", "c\n", "" }, Lines("abc\n", 0)); // the empty caret line after the final break
    }

    [Fact]
    public void NewlineForcesABreakAndBelongsToItsLine()
    {
        Assert.Equal(new[] { "ab\n", "cd" }, Lines("ab\ncd", 10));
    }

    [Fact]
    public void BlankLinesAndTrailingNewline()
    {
        Assert.Equal(new[] { "a\n", "\n", "b\n", "" }, Lines("a\n\nb\n", 10));
    }

    [Fact]
    public void CarriageReturnTakesNoSpace()
    {
        // "\r" is zero width: "abc\r\n" fits in 3 columns.
        Assert.Equal(new[] { "abc\r\n", "d" }, Lines("abc\r\nd", 3));
    }

    [Fact]
    public void LongParagraphsWrapIndependently()
    {
        Assert.Equal(new[] { "aa+", "bb\n", "cc+", "dd" }, Lines("aa+bb\ncc+dd", 3));
    }

    [Fact]
    public void SurrogatePairsAreNeverSplit()
    {
        string text = "a\U0001F600\U0001F600b"; // chars: a, hi, lo, hi, lo, b (2 columns of 10 per char incl. halves)
        var measurer = new FixedMeasurer(text);
        foreach (int columns in new[] { 1, 2, 3, 4 })
        {
            var lines = LineBreaker.Break(text, columns * 10, measurer.Measure);
            foreach (var (start, length) in lines)
            {
                int end = start + length;
                Assert.False(start > 0 && char.IsLowSurrogate(text[start]) && char.IsHighSurrogate(text[start - 1]), $"split at {start}");
                Assert.False(end < text.Length && end > 0 && char.IsHighSurrogate(text[end - 1]) && char.IsLowSurrogate(text[end]), $"split at {end}");
            }
        }
    }

    [Fact]
    public void LinesAlwaysPartitionTheText()
    {
        var random = new Random(42);
        const string alphabet = "ab(),+-*/^&=<> $.\n\r";
        for (int round = 0; round < 500; round++)
        {
            string text = new(Enumerable.Range(0, random.Next(0, 60)).Select(_ => alphabet[random.Next(alphabet.Length)]).ToArray());
            int columns = random.Next(0, 12);
            var measurer = new FixedMeasurer(text);
            var lines = LineBreaker.Break(text, columns * 10, measurer.Measure);

            int expectedStart = 0;
            foreach (var (start, length) in lines)
            {
                Assert.Equal(expectedStart, start);
                Assert.True(length >= 0);
                expectedStart = start + length;
            }
            Assert.Equal(text.Length, expectedStart);
        }
    }

    [Fact]
    public void LinesNeverExceedTheWidthUnlessTheyAreASingleCharacter()
    {
        var random = new Random(7);
        const string alphabet = "ab(),+-*/ ";
        for (int round = 0; round < 300; round++)
        {
            string text = new(Enumerable.Range(0, random.Next(1, 80)).Select(_ => alphabet[random.Next(alphabet.Length)]).ToArray());
            int columns = random.Next(1, 12);
            var measurer = new FixedMeasurer(text);
            foreach (var (start, length) in LineBreaker.Break(text, columns * 10, measurer.Measure))
            {
                string line = text.Substring(start, length).TrimEnd(' ');
                Assert.True(line.Length <= columns, $"'{line}' in {columns} columns of '{text}'");
            }
        }
    }
}
