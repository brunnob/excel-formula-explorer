using FormulaExplorer.AddIn.Composition;
using FormulaExplorer.AddIn.Excel;

namespace FormulaExplorer.AddIn.Tests;

public class WorkbookCloseTrackerTests
{
    private static readonly IntPtr WindowB = new(0xB0);
    private static readonly IntPtr WindowB2 = new(0xB2);
    private static readonly IntPtr WindowC = new(0xC0);

    private static WorkbookCloseTracker Tracking(string workbook = "B.xlsx", bool isRoot = false)
    {
        var tracker = new WorkbookCloseTracker();
        tracker.Begin(workbook, isRoot, new[] { WindowB, WindowB2 });
        return tracker;
    }

    private static Func<string, WorkbookPresence> Always(WorkbookPresence presence) => _ => presence;

    [Fact]
    public void ABeginPendingCloseIsTrackedAndItsWindowsAreClosingWindows()
    {
        var tracker = Tracking();
        Assert.True(tracker.Any);
        Assert.True(tracker.IsClosing("b.XLSX"));
        Assert.True(tracker.IsClosingWindow(WindowB));
        Assert.True(tracker.IsClosingWindow(WindowB2));
        Assert.False(tracker.IsClosingWindow(WindowC));
        Assert.False(tracker.IsClosingWindow(IntPtr.Zero));
    }

    [Fact]
    public void AWorkbookThatIsGoneEndsTheCloseAsClosed()
    {
        var tracker = Tracking(isRoot: true);
        Assert.Empty(tracker.Poll(Always(WorkbookPresence.Gone)));   // one "gone" is not enough (see the next test)
        var ended = tracker.Poll(Always(WorkbookPresence.Gone));

        var result = Assert.Single(ended);
        Assert.Equal(new CloseResult("B.xlsx", WasRoot: true, CloseOutcome.Closed), result);
        Assert.False(tracker.Any);
        Assert.False(tracker.IsClosingWindow(WindowB));
    }

    [Fact]
    public void ASavePromptThatStaysUpForAMinuteIsNeverTreatedAsCancelled()
    {
        var tracker = Tracking();
        for (int i = 0; i < 240; i++)   // 60 s at 250 ms
            Assert.Empty(tracker.Poll(Always(WorkbookPresence.Busy)));

        Assert.True(tracker.IsClosingWindow(WindowB));   // still protected: the window must not get the explorer
        Assert.Empty(tracker.Poll(Always(WorkbookPresence.Gone)));
        var ended = Assert.Single(tracker.Poll(Always(WorkbookPresence.Gone)));   // user chose Don't Save
        Assert.Equal(CloseOutcome.Closed, ended.Outcome);
    }

    [Fact]
    public void ASingleGonePollBetweenOtherAnswersNeverClosesAnything()
    {
        // A lookup that fails once (Save As renaming, a moment between dialogs) must not end the explorer of the root workbook.
        var tracker = Tracking(isRoot: true);
        var answers = new[]
        {
            WorkbookPresence.Busy, WorkbookPresence.Gone, WorkbookPresence.Busy, WorkbookPresence.Gone,
            WorkbookPresence.Open, WorkbookPresence.Gone, WorkbookPresence.Open,
        };
        foreach (var answer in answers)
            Assert.Empty(tracker.Poll(_ => answer));

        Assert.True(tracker.Any);
        Assert.True(tracker.IsClosingWindow(WindowB));
    }

    [Fact]
    public void ConsecutiveGonePollsCloseEvenAfterEarlierDoubts()
    {
        var tracker = Tracking(isRoot: true);
        Assert.Empty(tracker.Poll(Always(WorkbookPresence.Gone)));
        Assert.Empty(tracker.Poll(Always(WorkbookPresence.Open)));
        Assert.Empty(tracker.Poll(Always(WorkbookPresence.Gone)));
        Assert.Equal(CloseOutcome.Closed, Assert.Single(tracker.Poll(Always(WorkbookPresence.Gone))).Outcome);
    }

    [Fact]
    public void AWorkbookStillOpenAndNotBusyForSeveralPollsMeansTheCloseWasCancelled()
    {
        var tracker = Tracking();
        for (int i = 1; i < WorkbookCloseTracker.CallsBeforeCancelled; i++)
            Assert.Empty(tracker.Poll(Always(WorkbookPresence.Open)));

        var ended = Assert.Single(tracker.Poll(Always(WorkbookPresence.Open)));
        Assert.Equal(CloseOutcome.Cancelled, ended.Outcome);
        Assert.False(tracker.Any);
    }

    [Fact]
    public void ABusyAnswerRestartsTheCancelCountdown()
    {
        var tracker = Tracking();
        var answers = new Queue<WorkbookPresence>(new[]
        {
            WorkbookPresence.Open, WorkbookPresence.Open, WorkbookPresence.Busy,   // a second dialog came up
            WorkbookPresence.Open, WorkbookPresence.Open,
        });
        while (answers.Count > 0)
        {
            var answer = answers.Dequeue();
            Assert.Empty(tracker.Poll(_ => answer));
        }
        Assert.True(tracker.Any);
    }

    [Fact]
    public void TheMomentBetweenThePromptEndingAndTheWorkbookGoingAwayIsNotACancel()
    {
        // Open (Excel answers again) for fewer polls than needed, then the workbook disappears.
        var tracker = Tracking();
        Assert.Empty(tracker.Poll(Always(WorkbookPresence.Busy)));
        Assert.Empty(tracker.Poll(Always(WorkbookPresence.Open)));
        Assert.Empty(tracker.Poll(Always(WorkbookPresence.Gone)));
        Assert.Equal(CloseOutcome.Closed, Assert.Single(tracker.Poll(Always(WorkbookPresence.Gone))).Outcome);
    }

    [Fact]
    public void SeveralPendingClosesEndIndependently()
    {
        var tracker = new WorkbookCloseTracker();
        tracker.Begin("A.xlsx", isRoot: true, new[] { new IntPtr(0xA0) });
        tracker.Begin("B.xlsx", isRoot: false, new[] { WindowB });

        Func<string, WorkbookPresence> probe = name => name == "A.xlsx" ? WorkbookPresence.Gone : WorkbookPresence.Busy;
        Assert.Empty(tracker.Poll(probe));
        var ended = tracker.Poll(probe);

        var result = Assert.Single(ended);
        Assert.Equal("A.xlsx", result.Workbook);
        Assert.True(result.WasRoot);
        Assert.True(tracker.IsClosingWindow(WindowB));
        Assert.False(tracker.IsClosingWindow(new IntPtr(0xA0)));
    }

    [Fact]
    public void ClosingTheSameWorkbookAgainRestartsTracking()
    {
        var tracker = Tracking();
        tracker.Poll(Always(WorkbookPresence.Open));
        tracker.Poll(Always(WorkbookPresence.Open));
        tracker.Begin("B.xlsx", isRoot: false, new[] { WindowC });   // the user tried to close it again

        Assert.False(tracker.IsClosingWindow(WindowB));
        Assert.True(tracker.IsClosingWindow(WindowC));
        Assert.Empty(tracker.Poll(Always(WorkbookPresence.Open)));   // countdown restarted
    }

    [Fact]
    public void AClosedThatNeverResolvesIsEventuallyDropped()
    {
        var tracker = Tracking();
        IReadOnlyList<CloseResult> ended = Array.Empty<CloseResult>();
        for (int i = 0; i < WorkbookCloseTracker.MaxPolls && ended.Count == 0; i++)
            ended = tracker.Poll(Always(WorkbookPresence.Busy));

        Assert.Equal(CloseOutcome.Cancelled, Assert.Single(ended).Outcome);
        Assert.False(tracker.Any);
    }
}
