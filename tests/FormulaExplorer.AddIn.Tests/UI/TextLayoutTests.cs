using System.Windows;
using FormulaExplorer.AddIn.UI;

namespace FormulaExplorer.AddIn.Tests.UI;

public class TextLayoutTests
{
    private static TextLayout Layout(string text, int columns) =>
        TextLayout.Build(text, columns * 10, new FixedMeasurer(text));

    [Fact]
    public void LinesHaveTheFixedPitch()
    {
        var layout = Layout("aaaa+bbbb+cccc", 5);
        Assert.Equal(3, layout.Lines.Count);
        Assert.Equal(3 * TextLayout.LinePitch, layout.Height);
        Assert.Equal(20, TextLayout.LinePitch);
    }

    [Fact]
    public void HitTestFindsTheCharacterUnderThePoint()
    {
        var layout = Layout("abc+def", 4); // "abc+" / "def"
        Assert.Equal(0, layout.HitTest(new Point(0, 0)));
        Assert.Equal(0, layout.HitTest(new Point(9.9, 19.9)));
        Assert.Equal(1, layout.HitTest(new Point(10, 5)));
        Assert.Equal(3, layout.HitTest(new Point(39.9, 10)));
        Assert.Equal(4, layout.HitTest(new Point(0, 20)));
        Assert.Equal(6, layout.HitTest(new Point(25, 39.9)));
    }

    [Fact]
    public void EmptyAreasHitNothing()
    {
        var layout = Layout("abc+def", 4);
        Assert.Null(layout.HitTest(new Point(-0.1, 5)));      // left of the origin
        Assert.Null(layout.HitTest(new Point(5, -0.1)));      // above
        Assert.Null(layout.HitTest(new Point(40, 5)));        // right of the first line's text
        Assert.Null(layout.HitTest(new Point(35, 25)));       // right of the (shorter) second line
        Assert.Null(layout.HitTest(new Point(5, 40)));        // below the text
        Assert.Null(TextLayout.Empty.HitTest(new Point(1, 1)));
    }

    [Fact]
    public void LineTerminatorsAreNotHittable()
    {
        var layout = Layout("ab\r\ncd", 10); // line 0 = "ab\r\n" with two zero-width chars
        Assert.Equal(0, layout.HitTest(new Point(5, 5)));
        Assert.Equal(1, layout.HitTest(new Point(15, 5)));
        Assert.Null(layout.HitTest(new Point(20, 5)));
        Assert.Null(layout.CharRect(2));
        Assert.Null(layout.CharRect(3));
        Assert.Equal(4, layout.HitTest(new Point(5, 25)));
    }

    [Fact]
    public void CharRectsAreOneCharWideAndOneLineTall()
    {
        var layout = Layout("abc+def", 4);
        Assert.Equal(new Rect(10, 0, 10, 20), layout.CharRect(1));
        Assert.Equal(new Rect(20, 20, 10, 20), layout.CharRect(6));
        Assert.Null(layout.CharRect(-1));
        Assert.Null(layout.CharRect(7));
    }

    [Fact]
    public void SpanRectsAreOnePerLineAndAbut()
    {
        var layout = Layout("abc+def+g", 4); // "abc+" / "def+" / "g"
        var rects = layout.SpanRects(2, 7).ToList(); // "c+def+g"
        Assert.Equal(3, rects.Count);
        Assert.Equal(new Rect(20, 0, 20, 20), rects[0]);
        Assert.Equal(new Rect(0, 20, 40, 20), rects[1]);
        Assert.Equal(new Rect(0, 40, 10, 20), rects[2]);
        Assert.Equal(rects[0].Bottom, rects[1].Top);
        Assert.Equal(rects[1].Bottom, rects[2].Top);
    }

    [Fact]
    public void SpanRectsOfAnEmptyOrOutOfRangeSpanAreEmpty()
    {
        var layout = Layout("abc+def", 4);
        Assert.Empty(layout.SpanRects(3, 0));
        Assert.Empty(layout.SpanRects(7, 3));
        Assert.Empty(layout.SpanRects(-5, 0));
    }

    [Fact]
    public void SpanRectsSkipLineTerminatorsWithoutWidth()
    {
        var layout = Layout("ab\ncd", 10);
        var rects = layout.SpanRects(1, 3).ToList(); // "b\nc"
        Assert.Equal(2, rects.Count);
        Assert.Equal(new Rect(10, 0, 10, 20), rects[0]);
        Assert.Equal(new Rect(0, 20, 10, 20), rects[1]);
    }

    [Fact]
    public void LineOfMapsEveryCharToItsLine()
    {
        var layout = Layout("abc+def+g", 4);
        Assert.Equal(new[] { 0, 0, 0, 0, 1, 1, 1, 1, 2 }, Enumerable.Range(0, 9).Select(layout.LineOf));
        Assert.Equal(-1, layout.LineOf(9));
        Assert.Equal(-1, layout.LineOf(-1));
    }

    [Fact]
    public void EveryCharRectContainsItsOwnCenterInHitTest()
    {
        const string text = "=IF(A1>2,SUM(B1:B9)+1,\"x y\")\n&C3";
        var layout = Layout(text, 7);
        for (int i = 0; i < text.Length; i++)
        {
            if (layout.CharRect(i) is not { } rect) continue;
            Assert.Equal(i, layout.HitTest(new Point(rect.Left + rect.Width / 2, rect.Top + rect.Height / 2)));
        }
    }
}
