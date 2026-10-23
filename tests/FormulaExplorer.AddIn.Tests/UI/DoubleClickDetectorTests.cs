using System.Windows;
using FormulaExplorer.AddIn.UI;

namespace FormulaExplorer.AddIn.Tests.UI;

public class DoubleClickDetectorTests
{
    private long _now = 1000;
    private readonly object _a = new();
    private readonly object _b = new();

    private DoubleClickDetector Create(double interval = 500, double slopWidth = 4, double slopHeight = 4) =>
        new(() => _now, () => interval, () => (slopWidth, slopHeight));

    [Fact]
    public void FirstPressIsAClick()
    {
        Assert.Equal(PressAction.Click, Create().OnLinkPress(_a, new Point(10, 10), 1));
    }

    [Fact]
    public void SecondPressOnTheSameNodeWithinTheIntervalIsADoubleClick()
    {
        var detector = Create();
        detector.OnLinkPress(_a, new Point(10, 10), 1);
        _now += 499;
        Assert.Equal(PressAction.DoubleClick, detector.OnLinkPress(_a, new Point(11, 9), 1));
    }

    [Fact]
    public void TooSlowIsTwoClicks()
    {
        var detector = Create();
        detector.OnLinkPress(_a, new Point(10, 10), 1);
        _now += 501;
        Assert.Equal(PressAction.Click, detector.OnLinkPress(_a, new Point(10, 10), 1));
    }

    [Fact]
    public void TooFarIsTwoClicks()
    {
        var detector = Create(slopWidth: 4, slopHeight: 4);
        detector.OnLinkPress(_a, new Point(10, 10), 1);
        _now += 100;
        Assert.Equal(PressAction.Click, detector.OnLinkPress(_a, new Point(12.5, 10), 1)); // > 2 px either side
    }

    [Fact]
    public void ExactlyHalfTheRectangleIsStillInside()
    {
        var detector = Create(slopWidth: 4, slopHeight: 4);
        detector.OnLinkPress(_a, new Point(10, 10), 1);
        _now += 100;
        Assert.Equal(PressAction.DoubleClick, detector.OnLinkPress(_a, new Point(12, 8), 1));
    }

    [Fact]
    public void DifferentNodeIsAClickAndStartsANewPair()
    {
        var detector = Create();
        detector.OnLinkPress(_a, new Point(10, 10), 1);
        _now += 100;
        Assert.Equal(PressAction.Click, detector.OnLinkPress(_b, new Point(10, 10), 1));
        _now += 100;
        Assert.Equal(PressAction.DoubleClick, detector.OnLinkPress(_b, new Point(10, 10), 1));
    }

    [Fact]
    public void ThirdQuickPressStartsANewSequence()
    {
        var detector = Create();
        detector.OnLinkPress(_a, new Point(10, 10), 1);
        _now += 50;
        Assert.Equal(PressAction.DoubleClick, detector.OnLinkPress(_a, new Point(10, 10), 1));
        _now += 50;
        Assert.Equal(PressAction.Click, detector.OnLinkPress(_a, new Point(10, 10), 3));
        _now += 50;
        Assert.Equal(PressAction.DoubleClick, detector.OnLinkPress(_a, new Point(10, 10), 1));
    }

    [Fact]
    public void ResetForgetsThePreviousPress()
    {
        var detector = Create();
        detector.OnLinkPress(_a, new Point(10, 10), 1);
        detector.Reset();
        _now += 10;
        Assert.Equal(PressAction.Click, detector.OnLinkPress(_a, new Point(10, 10), 1));
    }

    [Fact]
    public void WpfClickCountTwoIsAcceptedEvenWhenTheTimingFailed()
    {
        var detector = Create();
        detector.OnLinkPress(_a, new Point(10, 10), 1);
        _now += 5000;
        Assert.Equal(PressAction.DoubleClick, detector.OnLinkPress(_a, new Point(10, 10), 2));
    }

    [Fact]
    public void WpfClickCountTwoWithoutAFirstClickRaisesBoth()
    {
        Assert.Equal(PressAction.ClickThenDoubleClick, Create().OnLinkPress(_a, new Point(10, 10), 2));
    }

    [Fact]
    public void SystemDetectorUsesSystemSettings()
    {
        var detector = DoubleClickDetector.CreateForSystem(() => 1.0);
        Assert.Equal(PressAction.Click, detector.OnLinkPress(_a, new Point(10, 10), 1));
        Assert.Equal(PressAction.DoubleClick, detector.OnLinkPress(_a, new Point(10, 10), 1));
    }

    [Fact]
    public void PressesAreTimedByTheirInputTimestampNotByWhenTheyAreProcessed()
    {
        // The first press's navigation takes 800 ms, so the second press is processed late although the user
        // clicked 250 ms after the first.
        var detector = Create();
        Assert.Equal(PressAction.Click, detector.OnLinkPress(_a, new Point(10, 10), 1, timestampMs: 1000));
        _now += 800;
        Assert.Equal(PressAction.DoubleClick, detector.OnLinkPress(_a, new Point(10, 10), 1, timestampMs: 1250));
    }

    [Fact]
    public void ARealGapBeyondTheIntervalIsStillTwoClicksWhateverTheProcessingTime()
    {
        var detector = Create();
        detector.OnLinkPress(_a, new Point(10, 10), 1, timestampMs: 1000);
        Assert.Equal(PressAction.Click, detector.OnLinkPress(_a, new Point(10, 10), 1, timestampMs: 1501));
    }

    [Fact]
    public void TimestampsSurviveTheTickCountRollover()
    {
        var detector = Create();
        detector.OnLinkPress(_a, new Point(10, 10), 1, timestampMs: int.MaxValue - 100);
        Assert.Equal(PressAction.DoubleClick, detector.OnLinkPress(_a, new Point(10, 10), 1, timestampMs: int.MinValue + 100));
    }

    // ---- ContentReplaced: a click that drilled down swallows the second press of a habitual double click ----------

    [Fact]
    public void ASecondClosePressAfterContentReplacedIsIgnored()
    {
        var detector = Create();
        detector.OnLinkPress(_a, new Point(10, 10), 1);
        detector.ContentReplaced();
        _now += 100;
        Assert.Equal(PressAction.Ignore, detector.OnLinkPress(_b, new Point(11, 10), 1));
    }

    [Fact]
    public void AfterTheIgnoredPressTheNextOneIsAnOrdinaryClick()
    {
        var detector = Create();
        detector.OnLinkPress(_a, new Point(10, 10), 1);
        detector.ContentReplaced();
        _now += 100;
        detector.OnLinkPress(_b, new Point(10, 10), 1);
        _now += 100;
        Assert.Equal(PressAction.Click, detector.OnLinkPress(_b, new Point(10, 10), 1));
    }

    [Fact]
    public void ALatePressAfterContentReplacedIsAClick()
    {
        var detector = Create();
        detector.OnLinkPress(_a, new Point(10, 10), 1);
        detector.ContentReplaced();
        _now += 501;
        Assert.Equal(PressAction.Click, detector.OnLinkPress(_b, new Point(10, 10), 1));
    }

    [Fact]
    public void AFarPressAfterContentReplacedIsAClick()
    {
        var detector = Create(slopWidth: 4, slopHeight: 4);
        detector.OnLinkPress(_a, new Point(10, 10), 1);
        detector.ContentReplaced();
        _now += 100;
        Assert.Equal(PressAction.Click, detector.OnLinkPress(_b, new Point(40, 10), 1));
    }

    [Fact]
    public void ClickCountTwoAfterContentReplacedIsIgnoredEvenWhenLateAndFar()
    {
        var detector = Create();
        detector.OnLinkPress(_a, new Point(10, 10), 1);
        detector.ContentReplaced();
        _now += 900;
        Assert.Equal(PressAction.Ignore, detector.OnLinkPress(_b, new Point(80, 80), 2));
    }

    [Fact]
    public void ContentReplacedWithoutAnyPressChangesNothing()
    {
        var detector = Create();
        detector.ContentReplaced();
        Assert.Equal(PressAction.Click, detector.OnLinkPress(_a, new Point(10, 10), 1));
    }

    [Fact]
    public void ResetClearsThePendingIgnore()
    {
        var detector = Create();
        detector.OnLinkPress(_a, new Point(10, 10), 1);
        detector.ContentReplaced();
        detector.Reset();
        _now += 100;
        Assert.Equal(PressAction.Click, detector.OnLinkPress(_b, new Point(10, 10), 1));
    }

    [Fact]
    public void ContentReplacedDoesNotLetTheOldNodePairWithANewPress()
    {
        var detector = Create();
        detector.OnLinkPress(_a, new Point(10, 10), 1);
        detector.ContentReplaced();
        _now += 100;
        detector.OnLinkPress(_a, new Point(10, 10), 1); // ignored
        _now += 100;
        Assert.Equal(PressAction.Click, detector.OnLinkPress(_a, new Point(10, 10), 1));
    }
}
