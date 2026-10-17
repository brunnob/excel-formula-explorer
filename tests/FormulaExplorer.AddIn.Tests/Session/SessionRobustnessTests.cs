using FormulaExplorer.AddIn.Excel;
using FormulaExplorer.AddIn.Session;

namespace FormulaExplorer.AddIn.Tests.Sessions;

/// <summary>Busy Excel retries, unexpected exceptions, versioned loads and Close.</summary>
public class SessionRobustnessTests
{
    private static readonly CellLocation K33 = Rig.Loc("K33");
    private static readonly CellLocation K31 = Rig.Loc("K31");

    // ---- busy Excel --------------------------------------------------------------------------------------------

    [Fact]
    public void StartRetriesEvery200MsWhileExcelIsBusy()
    {
        var rig = Rig.VideoModel();
        rig.Gateway.Busy["GetActiveCell"] = 2;

        rig.Session.Start();
        Assert.Equal(1, rig.Gateway.Calls("GetActiveCell"));
        rig.Scheduler.Advance(199);
        Assert.Equal(1, rig.Gateway.Calls("GetActiveCell"));
        rig.Scheduler.Advance(1);
        Assert.Equal(2, rig.Gateway.Calls("GetActiveCell"));
        rig.Scheduler.Advance(200);
        Assert.Equal(3, rig.Gateway.Calls("GetActiveCell"));
        rig.Settle();

        Assert.Equal(Rig.K33Formula, rig.View.Text);
        Assert.Equal(K33, rig.Snapshot.Root);
    }

    [Fact]
    public void BusyRetriesStopAfter25()
    {
        var rig = Rig.VideoModel();
        rig.Gateway.Busy["GetActiveCell"] = int.MaxValue;

        rig.Session.Start();
        rig.Settle();

        Assert.Equal(26, rig.Gateway.Calls("GetActiveCell")); // the first attempt and 25 retries
        Assert.Equal(TimeSpan.FromMilliseconds(25 * 200), rig.Scheduler.Now);
        Assert.True(rig.Session.IsOpen);
    }

    [Fact]
    public void ClickRetriesNavigationWhenExcelIsBusy()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Gateway.Busy["NavigateTo"] = 2;

        rig.View.Click(rig.View.Node("K$31"));
        Assert.Equal("(none)", rig.View.GoldText);
        rig.Settle();

        Assert.Equal(3, rig.Gateway.Calls("NavigateTo"));
        Assert.Equal("K$31", rig.View.GoldText);
        Assert.Equal(K31, rig.Snapshot.GoldTarget);
    }

    [Fact]
    public void DrillRetriesReadingTheTargetCellWhenExcelIsBusy()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Click("K$31");
        rig.Gateway.Busy["ReadCell"] = 1;

        rig.DoubleClick("K$31");

        Assert.Equal(K31, rig.Snapshot.Root);
        Assert.Equal(new[] { K33 }, rig.Snapshot.Stack);
        Assert.Equal(2, rig.Gateway.NavigateCalls.Count); // both presses navigate; the retry does not navigate again
    }

    [Fact]
    public void BackRetriesWithoutLosingTheStackEntry()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.DoubleClick("K$31");
        rig.Gateway.Busy["ReadCell"] = 1;

        rig.Back();

        Assert.Equal(K33, rig.Snapshot.Root);
        Assert.Empty(rig.Snapshot.Stack);
    }

    [Fact]
    public void ASelectionIsRetriedWhenExcelIsBusyAndPushesOnlyOnce()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Gateway.Busy["ReadCell"] = 2;

        rig.Gateway.UserSelects(K31);
        rig.Settle();

        Assert.Equal(K31, rig.Snapshot.Root);
        Assert.Equal(new[] { K33 }, rig.Snapshot.Stack);
        Assert.Equal(Rig.K31Formula, rig.View.Text);
    }

    [Fact]
    public void ASelectionWhoseRestoreFindsExcelBusyStillLoadsOnce()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Gateway.Busy["RestoreExcept"] = 1;

        rig.Gateway.UserSelects(K31);
        rig.Settle();

        Assert.Equal(K31, rig.Snapshot.Root);
        Assert.Equal(new[] { K33 }, rig.Snapshot.Stack);
        Assert.Equal(2, rig.View.ShowLoadingCount);
        Assert.Equal(new[] { K31 }, Assert.Single(rig.Gateway.RestoreCalls));
    }

    [Fact]
    public void ALoadRetriesWhenEvaluationFindsExcelBusy()
    {
        var rig = Rig.VideoModel();
        rig.Gateway.Busy["Evaluate"] = 3;

        rig.Start();

        Assert.Equal(Rig.K33Formula, rig.View.Text);
        Assert.True(rig.View.Display!.Interactive);
        Assert.Equal(1, rig.View.ShowFormulaCount);
        Assert.Equal(new[] { "Model!$I$27#0000FF", "Model!$K$31#9900CC" }, rig.Overlay.Boxes.Select(Rig.Box));
    }

    [Fact]
    public void ALoadThatNeverGetsThroughFallsBackToPlainText()
    {
        var rig = Rig.VideoModel();
        rig.Gateway.Busy["Evaluate"] = int.MaxValue;

        rig.Start();

        Assert.Equal(26, rig.Gateway.Calls("Evaluate")); // the first attempt and 25 retries
        Assert.False(rig.View.IsLoading);
        Assert.Equal(Rig.K33Formula, rig.View.Text);
        Assert.False(rig.View.Display!.Interactive);
    }

    [Fact]
    public void ALoadThatThrowsUnexpectedlyFallsBackToPlainText()
    {
        var rig = Rig.VideoModel();
        rig.Gateway.Failures["Evaluate"] = new InvalidOperationException("boom");

        rig.Start();

        Assert.False(rig.View.IsLoading);
        Assert.Equal(Rig.K33Formula, rig.View.Text);
        Assert.False(rig.View.Display!.Interactive);
    }

    // ---- busy Excel while gateway events are pending (no hot loop) -----------------------------------------------

    [Theory]
    [InlineData("ReadCell")]
    [InlineData("RestoreExcept")]
    public void ABusyDrainBacksOffOnceEvery200MsAndStopsAfter25Retries(string busyOperation)
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Gateway.Busy[busyOperation] = int.MaxValue;

        rig.Gateway.UserSelects(K31);
        rig.Scheduler.RunPosted(); // would never return if the drain re-armed itself immediately
        Assert.Equal(1, rig.Gateway.Calls(busyOperation));

        rig.Scheduler.Advance(199);
        Assert.Equal(1, rig.Gateway.Calls(busyOperation));
        rig.Scheduler.Advance(1);
        Assert.Equal(2, rig.Gateway.Calls(busyOperation));

        rig.Settle();
        Assert.Equal(26, rig.Gateway.Calls(busyOperation)); // the first attempt and 25 retries
        Assert.True(rig.Session.IsOpen);
    }

    [Fact]
    public void AnEventDuringTheBackOffIsCoveredByThePendingRetry()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Gateway.Busy["ReadCell"] = int.MaxValue;
        rig.Gateway.UserSelects(K31);
        rig.Scheduler.RunPosted();
        Assert.Equal(1, rig.Gateway.Calls("ReadCell"));

        rig.Gateway.RaiseValuesInvalidated();
        rig.Gateway.UserSelects(Rig.Loc("K19"));
        rig.Scheduler.RunPosted();
        Assert.Equal(1, rig.Gateway.Calls("ReadCell")); // nothing immediate

        rig.Gateway.Busy.Clear();
        rig.Settle();
        Assert.Equal(Rig.Loc("K19"), rig.Snapshot.Root); // the retry handled the newest selection
        Assert.Equal(2, rig.Gateway.Calls("ReadCell"));
    }

    [Fact]
    public void AGivenUpDrainKeepsItsItemsButOnlyANewEventRetriesThem()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Gateway.Busy["ReadCell"] = int.MaxValue;
        rig.Gateway.UserSelects(K31);
        rig.Settle();
        Assert.Equal(26, rig.Gateway.Calls("ReadCell"));

        rig.Gateway.Busy.Clear();
        rig.Gateway.RaiseValuesInvalidated(); // a new event starts a fresh drain at attempt 0 and finds Excel free
        rig.Settle();

        Assert.Equal(K31, rig.Snapshot.Root);
    }

    [Fact]
    public void ABusyClickWithEventsPendingDoesNotSpinTheDrain()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Gateway.Busy["NavigateTo"] = int.MaxValue;
        rig.Gateway.Busy["ReadCell"] = int.MaxValue;

        rig.Gateway.UserSelects(K31);
        rig.View.Click(rig.View.Node("K$31"));
        rig.Scheduler.RunPosted();
        Assert.Equal(1, rig.Gateway.Calls("ReadCell"));

        rig.Settle();
        Assert.Equal(26, rig.Gateway.Calls("ReadCell"));
        Assert.Equal(26, rig.Gateway.Calls("NavigateTo"));
    }

    // ---- a user command is never retried for long on a bare VBA_E_IGNORE ----------------------------------------

    private static Exception BareVbaIgnore() => new System.Runtime.InteropServices.COMException("not ready", unchecked((int)0x800AC472));

    private static Exception CallRejected() => new System.Runtime.InteropServices.COMException("rejected", unchecked((int)0x80010001));

    [Fact]
    public void AClickThatExcelRejectsWithABareVbaIgnoreIsRetriedOnlyBrieflyAndNeverFiresLater()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Gateway.BusyCause = BareVbaIgnore;
        rig.Gateway.Busy["NavigateTo"] = int.MaxValue;

        rig.View.Click(rig.View.Node("K$31"));
        rig.Settle();
        Assert.Equal(1 + ExplorerSession.UserCommandNotReadyRetries, rig.Gateway.Calls("NavigateTo"));

        // Excel becomes ready much later (the "edit" ended): the dropped click must not run now.
        rig.Gateway.Busy.Clear();
        rig.Scheduler.Advance(10_000);
        rig.Settle();
        Assert.Empty(rig.Gateway.NavigateCalls);
    }

    [Fact]
    public void AClickThatFindsExcelBusyForAnotherReasonKeepsTheFullRetryBudget()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Gateway.BusyCause = CallRejected;
        rig.Gateway.Busy["NavigateTo"] = int.MaxValue;

        rig.View.Click(rig.View.Node("K$31"));
        rig.Settle();
        Assert.Equal(26, rig.Gateway.Calls("NavigateTo"));
    }

    [Fact]
    public void ABareVbaIgnoreInBackgroundWorkKeepsTheFullRetryBudget()
    {
        var rig = new Rig();
        rig.Gateway.BusyCause = BareVbaIgnore;
        rig.Gateway.Busy["GetActiveCell"] = int.MaxValue;
        rig.Session.Start();
        rig.Settle();

        Assert.Equal(26, rig.Gateway.Calls("GetActiveCell"));
    }

    // ---- the fallback is clean and recoverable ----------------------------------------------------------------

    [Fact]
    public void AFallbackClearsThePreviousRootsBorders()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        Assert.Equal(2, rig.Snapshot.Boxes.Count);
        rig.Gateway.Busy["Evaluate"] = int.MaxValue;

        rig.Gateway.UserSelects(K31);
        rig.Settle();

        Assert.Equal(Rig.K31Formula, rig.View.Text);
        Assert.False(rig.View.Display!.Interactive);
        Assert.Empty(rig.Overlay.Boxes);
        Assert.Empty(rig.Snapshot.Boxes);
    }

    [Fact]
    public void AFallbackBecomesInteractiveAgainOnTheNextInvalidationOnceExcelIsFree()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Gateway.Busy["Evaluate"] = int.MaxValue;
        rig.Gateway.UserSelects(K31);
        rig.Settle();
        Assert.False(rig.View.Display!.Interactive);

        rig.Gateway.Busy.Clear();
        rig.Gateway.RaiseValuesInvalidated();
        rig.Settle();

        Assert.True(rig.View.Display!.Interactive);
        Assert.Equal(new[] { "$J$31", "$K$19" }, rig.Overlay.Boxes.Select(b => b.Location.Address));
        Assert.Equal(rig.Overlay.Boxes.Select(Rig.Box), rig.Snapshot.Boxes.Select(Rig.Box));
    }

    [Fact]
    public void AFallbackBecomesInteractiveAgainWhenTheRootIsReReadUnchanged()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Gateway.Busy["Evaluate"] = int.MaxValue;
        rig.Gateway.UserSelects(K31);
        rig.Settle();
        Assert.False(rig.View.Display!.Interactive);

        rig.Gateway.Busy.Clear();
        rig.Gateway.RaiseCellsChanged(K31);
        rig.Settle();

        Assert.True(rig.View.Display!.Interactive);
    }

    // ---- unexpected exceptions ---------------------------------------------------------------------------------

    [Fact]
    public void ExceptionsInHandlersAreSwallowedAndTheSessionKeepsWorking()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Gateway.Failures["ReadCell"] = new InvalidOperationException("boom");

        rig.Gateway.UserSelects(K31);
        rig.Settle();
        rig.View.PressBack();
        rig.Settle();
        rig.Gateway.Failures.Clear();

        Assert.True(rig.Session.IsOpen);
        Assert.Equal(K33, rig.Snapshot.Root);
        rig.Gateway.UserSelects(K31);
        rig.Settle();
        Assert.Equal(K31, rig.Snapshot.Root);
    }

    [Fact]
    public void ExceptionsOnStartDoNotEscape()
    {
        var rig = Rig.VideoModel();
        rig.Gateway.Failures["GetActiveCell"] = new InvalidOperationException("boom");

        var ex = Record.Exception(() => { rig.Session.Start(); rig.Settle(); });

        Assert.Null(ex);
        Assert.True(rig.Session.IsOpen);
    }

    [Fact]
    public void ExceptionsFromTheViewDoNotEscapeTheClickHandler()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Gateway.Failures["NavigateTo"] = new InvalidOperationException("boom");

        var ex = Record.Exception(() => { rig.View.Click(rig.View.Node("K$31")); rig.Settle(); });

        Assert.Null(ex);
        Assert.Equal("(none)", rig.View.GoldText);
    }

    // ---- close -------------------------------------------------------------------------------------------------

    [Fact]
    public void CloseUnsubscribesClearsTheOverlayAndEndsTheGatewaySession()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        Assert.True(rig.Gateway.HasSubscribers);
        Assert.True(rig.View.HasSubscribers);

        rig.Session.Close();

        Assert.False(rig.Session.IsOpen);
        Assert.False(rig.Gateway.HasSubscribers);
        Assert.False(rig.View.HasSubscribers);
        Assert.Equal(1, rig.Overlay.ClearCalls);
        Assert.Empty(rig.Overlay.Boxes);
        Assert.Equal(1, rig.Gateway.EndSessionCalls);
        Assert.False(rig.Snapshot.IsOpen);
        Assert.Empty(rig.Snapshot.Boxes);
    }

    [Fact]
    public void CloseIsIdempotent()
    {
        var rig = Rig.VideoModel();
        rig.Start();

        rig.Session.Close();
        rig.Session.Close();
        rig.Session.Dispose();

        Assert.Equal(1, rig.Overlay.ClearCalls);
        Assert.Equal(1, rig.Gateway.EndSessionCalls);
    }

    [Fact]
    public void DisposeClosesTheSession()
    {
        var rig = Rig.VideoModel();
        rig.Start();

        rig.Session.Dispose();

        Assert.False(rig.Session.IsOpen);
        Assert.Equal(1, rig.Gateway.EndSessionCalls);
    }

    [Fact]
    public void CloseDoesNotDisposeTheCollaborators()
    {
        var rig = Rig.VideoModel();
        rig.Start();

        rig.Session.Close();

        Assert.Equal(0, rig.Gateway.DisposeCalls);
        Assert.Equal(0, rig.Overlay.DisposeCalls);
    }

    [Fact]
    public void WorkPostedBeforeCloseDoesNothingAfterwards()
    {
        var rig = Rig.VideoModel();
        rig.Session.Start();                    // load pending
        rig.Gateway.UserSelects(K31);           // drain pending
        rig.Gateway.RaiseValuesInvalidated();

        rig.Session.Close();
        rig.Settle();

        Assert.Equal(0, rig.View.ShowFormulaCount);
        Assert.Equal(0, rig.View.RefreshHoverCount);
        Assert.Empty(rig.Overlay.ShowCalls);
        Assert.Empty(rig.Gateway.RestoreCalls);
    }

    [Fact]
    public void EventsAfterCloseAreIgnored()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        var node = rig.View.Node("K$31");
        rig.Session.Close();

        rig.View.Click(node);
        rig.View.PressBack();
        rig.Gateway.UserSelects(K31);
        rig.Gateway.RaiseCellsChanged(K33);
        rig.Gateway.RaiseValuesInvalidated();
        rig.Settle();

        Assert.Empty(rig.Gateway.NavigateCalls);
        Assert.Empty(rig.Gateway.RestoreCalls);
        Assert.Equal(1, rig.View.ShowFormulaCount);
    }

    [Fact]
    public void CloseKeepsTheSelectionWhereItIs()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Click("K$31");

        rig.Session.Close();

        // The session never moves the selection on close; the gateway's EndSession keeps the current one visible.
        Assert.Single(rig.Gateway.NavigateCalls);
        Assert.Equal(K31, rig.Gateway.Active);
        Assert.Equal(1, rig.Gateway.EndSessionCalls);
    }

    [Fact]
    public void ClosingASessionThatNeverStartedDoesNothingAndItCannotStartAfterwards()
    {
        var rig = Rig.VideoModel();

        rig.Session.Close();
        rig.Session.Start();

        Assert.False(rig.Session.IsOpen);
        Assert.Equal(0, rig.Gateway.EndSessionCalls);
        Assert.Equal(0, rig.Overlay.ClearCalls);
        Assert.Equal(0, rig.View.ShowLoadingCount);
    }

    [Fact]
    public void FailuresWhileClosingDoNotEscapeAndTheGatewayIsStillEnded()
    {
        var rig = Rig.VideoModel();
        var overlay = new ThrowingOverlay();
        var session = new ExplorerSession(rig.Gateway, rig.View, overlay, rig.Scheduler);
        session.Start();
        rig.Settle();

        var ex = Record.Exception(session.Close);

        Assert.Null(ex);
        Assert.Equal(1, rig.Gateway.EndSessionCalls);
    }

    private sealed class ThrowingOverlay : IPrecedentOverlay
    {
        public void SetExplorerWindow(IntPtr explorerHwnd) { }
        public void Show(CellLocation root, IReadOnlyList<PrecedentBox> boxes) { }
        public void Clear() => throw new InvalidOperationException("boom");
        public IReadOnlyList<OverlayRect> GetDrawnRects() => Array.Empty<OverlayRect>();
        public void Dispose() { }
    }
}
