using FormulaExplorer.AddIn.Excel;

namespace FormulaExplorer.AddIn.Tests.Sessions;

/// <summary>Gateway events: manual selection, cells changed, values invalidated, and their deferral/coalescing.</summary>
public class SessionEventTests
{
    private static readonly CellLocation K33 = Rig.Loc("K33");
    private static readonly CellLocation K31 = Rig.Loc("K31");
    private static readonly CellLocation J31 = Rig.Loc("J31");

    // ---- manual selection --------------------------------------------------------------------------------------

    [Fact]
    public void SelectingAnotherCellPushesTheRootAndShowsTheNewFormula()
    {
        var rig = Rig.VideoModel();
        rig.Start();

        rig.Gateway.UserSelects(K31);
        rig.Settle();

        var snapshot = rig.Snapshot;
        Assert.Equal(K31, snapshot.Root);
        Assert.Equal(new[] { K33 }, snapshot.Stack);
        Assert.Equal(Rig.K31Formula, rig.View.Text);
        Assert.Equal(new[] { new[] { K31 } }, rig.Gateway.RestoreCalls.Select(c => c));
        Assert.True(rig.View.BackEnabled);
        Assert.Equal(K31, rig.Overlay.ShowCalls[^1].Root);
    }

    [Fact]
    public void SelectionBlanksTheTextAndEnablesBackBeforeTheNewTextShows()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        int mark = rig.View.Log.Count;

        rig.Gateway.UserSelects(K31);
        rig.Scheduler.RunPosted();

        Assert.Equal(new[] { "Back:True", "Loading" }, rig.View.Log.Skip(mark));
        Assert.Null(rig.View.Text);
        rig.Scheduler.Advance(150);
        Assert.Equal(Rig.K31Formula, rig.View.Text);
    }

    [Fact]
    public void SelectionClearsTheGoldAndRestoresExceptTheNewSelection()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Click("K$31");

        rig.Gateway.UserSelects(J31);
        rig.Settle();

        Assert.Null(rig.Snapshot.GoldStart);
        Assert.Equal(J31, rig.Snapshot.Root);
        Assert.Equal(new[] { J31 }, rig.Gateway.RestoreCalls[^1]);
        Assert.Equal(new[] { K33 }, rig.Snapshot.Stack);
    }

    [Fact]
    public void SelectingTheRootCellAgainOnlyClearsTheGold()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Click("K$31");

        rig.Gateway.UserSelects(K33);
        rig.Settle();

        Assert.Equal("(none)", rig.View.GoldText);
        Assert.Equal(K33, rig.Snapshot.Root);
        Assert.Empty(rig.Snapshot.Stack);
        Assert.False(rig.View.BackEnabled);
        Assert.Equal(new[] { K33 }, Assert.Single(rig.Gateway.RestoreCalls));
        Assert.Equal(1, rig.View.ShowLoadingCount);
    }

    [Fact]
    public void SelectingAnUnreadableCellIsIgnored()
    {
        var rig = Rig.VideoModel();
        rig.Start();

        rig.Gateway.UserSelects(Rig.Loc("Z99"));
        rig.Settle();

        Assert.Equal(K33, rig.Snapshot.Root);
        Assert.Empty(rig.Snapshot.Stack);
        Assert.Empty(rig.Gateway.RestoreCalls);
        Assert.Equal(1, rig.View.ShowLoadingCount);
    }

    [Fact]
    public void OnlyTheLatestOfSeveralSelectionsIsHandled()
    {
        var rig = Rig.VideoModel();
        rig.Start();

        rig.Gateway.UserSelects(K31);
        rig.Gateway.UserSelects(J31);
        rig.Gateway.UserSelects(Rig.Loc("K19"));
        rig.Settle();

        Assert.Equal(Rig.Loc("K19"), rig.Snapshot.Root);
        Assert.Equal(new[] { K33 }, rig.Snapshot.Stack);
        Assert.Equal(2, rig.View.ShowLoadingCount);
        Assert.Equal("0.05", rig.View.Text);
    }

    [Fact]
    public void SelectionEventsAreNotHandledInline()
    {
        var rig = Rig.VideoModel();
        rig.Start();

        rig.Gateway.UserSelects(K31);

        Assert.Equal(K33, rig.Snapshot.Root);
        Assert.Empty(rig.Gateway.RestoreCalls);
        rig.Scheduler.RunPosted();
        Assert.Equal(K31, rig.Snapshot.Root);
    }

    [Fact]
    public void ASelectionDuringTheBlankReplacesThePendingLoad()
    {
        var rig = Rig.VideoModel();
        rig.Session.Start(); // loading K33 (150 ms)
        rig.Scheduler.Advance(100);

        rig.Gateway.UserSelects(K31);
        rig.Scheduler.RunPosted();
        rig.Scheduler.Advance(60); // the first load's timer fires here and must be ignored
        Assert.True(rig.View.IsLoading);
        rig.Scheduler.Advance(100);

        Assert.Equal(1, rig.View.ShowFormulaCount);
        Assert.Equal(Rig.K31Formula, rig.View.Text);
        Assert.Equal(new[] { K33 }, rig.Snapshot.Stack);
    }

    // ---- cells changed -----------------------------------------------------------------------------------------

    [Fact]
    public void ChangesOnAnotherSheetAreIgnored()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        int reads = rig.Gateway.Calls("ReadCell");

        rig.Gateway.RaiseCellsChanged(Rig.Loc("B1", Rig.Assumptions));
        rig.Settle();

        Assert.Equal(reads, rig.Gateway.Calls("ReadCell"));
        Assert.Equal(1, rig.View.ShowLoadingCount);
    }

    [Fact]
    public void ChangesWithTheSameTextDoNotReload()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Click("K$31");

        rig.Gateway.RaiseCellsChanged(Rig.Loc("A1"));
        rig.Settle();

        Assert.Equal(1, rig.View.ShowLoadingCount);
        Assert.Equal("K$31", rig.View.GoldText);
    }

    [Fact]
    public void ANewFormulaInTheRootReloadsClearsGoldAndKeepsTheStack()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.DoubleClick("K$31");
        rig.Click("K19");
        Assert.Equal("K19", rig.View.GoldText);

        rig.Gateway.Formula(Rig.Model, "K31", "=J31*2");
        rig.Gateway.RaiseCellsChanged(K31);
        rig.Settle();

        Assert.Equal("=J31*2", rig.View.Text);
        Assert.Equal("(none)", rig.View.GoldText);
        Assert.Null(rig.Snapshot.GoldStart);
        Assert.Equal(new[] { K33 }, rig.Snapshot.Stack);
        Assert.Equal(K31, rig.Snapshot.Root);
    }

    [Fact]
    public void TheFormulaTurningIntoAConstantReloadsAsNonInteractive()
    {
        var rig = Rig.VideoModel();
        rig.Start();

        rig.Gateway.Constant(Rig.Model, "K33", "=1/2"); // text only: HasFormula false
        rig.Gateway.RaiseCellsChanged(K33);
        rig.Settle();

        Assert.False(rig.View.Display!.Interactive);
        Assert.Equal("=1/2", rig.View.Text);
        Assert.Empty(rig.Overlay.Boxes);
    }

    [Fact]
    public void SeveralChangeEventsAreMergedIntoOneReRead()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        int reads = rig.Gateway.Calls("ReadCell");

        rig.Gateway.RaiseCellsChanged(Rig.Loc("A1"));
        rig.Gateway.RaiseCellsChanged(Rig.Loc("B2"));
        rig.Gateway.RaiseCellsChanged(Rig.Loc("C3:D4"));
        rig.Settle();

        Assert.Equal(reads + 1, rig.Gateway.Calls("ReadCell"));
    }

    [Fact]
    public void ARootThatDisappearsShowsEmptyTextClearsTheBordersAndKeepsTheStack()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.DoubleClick("K$31");
        rig.Click("K19");
        int clears = rig.Overlay.ClearCalls;

        rig.Gateway.Cells.Remove(K31);
        rig.Gateway.RaiseCellsChanged(K31);
        rig.Settle();

        var snapshot = rig.Snapshot;
        Assert.Null(snapshot.Root);
        Assert.Equal(string.Empty, rig.View.Text);
        Assert.False(rig.View.Display!.Interactive);
        Assert.Equal(clears + 1, rig.Overlay.ClearCalls);
        Assert.Empty(rig.Overlay.Boxes);
        Assert.Null(snapshot.GoldStart);
        Assert.Equal(new[] { K33 }, snapshot.Stack);
        Assert.True(snapshot.BackEnabled);
    }

    [Fact]
    public void AfterTheRootDisappearedTheNextSelectionBecomesTheRootWithoutPushing()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Gateway.Cells.Remove(K33);
        rig.Gateway.RaiseCellsChanged(K33);
        rig.Settle();

        rig.Gateway.UserSelects(K31);
        rig.Settle();

        Assert.Equal(K31, rig.Snapshot.Root);
        Assert.Empty(rig.Snapshot.Stack);
        Assert.Equal(Rig.K31Formula, rig.View.Text);
    }

    [Fact]
    public void ARootThatDisappearsWhileLoadingCancelsTheLoad()
    {
        var rig = Rig.VideoModel();
        rig.Session.Start();
        rig.Scheduler.Advance(100);

        rig.Gateway.Cells.Remove(K33);
        rig.Gateway.RaiseCellsChanged(K33);
        rig.Scheduler.RunPosted();
        rig.Settle();

        Assert.Equal(string.Empty, rig.View.Text);
        Assert.Equal(1, rig.View.ShowFormulaCount);
    }

    // ---- values invalidated ------------------------------------------------------------------------------------

    [Fact]
    public void ValuesInvalidatedRefreshesTheHoverOnceHoweverManyEventsArrive()
    {
        var rig = Rig.VideoModel();
        rig.Start();

        for (int i = 0; i < 10; i++) rig.Gateway.RaiseValuesInvalidated();
        Assert.Equal(0, rig.View.RefreshHoverCount);
        rig.Settle();

        Assert.Equal(1, rig.View.RefreshHoverCount);
        Assert.Equal(1, rig.View.ShowLoadingCount);
        Assert.Equal(Rig.K33Formula, rig.View.Text);
    }

    [Fact]
    public void InvalidationKeepsGoldAndBoxes()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Click("K$31");
        var boxes = rig.Overlay.Boxes;

        rig.Gateway.RaiseValuesInvalidated();
        rig.Settle();

        Assert.Equal("K$31", rig.View.GoldText);
        Assert.Same(boxes, rig.Overlay.Boxes);
    }

    // ---- deferral while a command runs -------------------------------------------------------------------------

    [Fact]
    public void EventsRaisedInsideAGotoAreHandledAfterTheClickNotDuringIt()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        int refreshesDuringClick = -1;
        rig.Gateway.DuringNavigate = () =>
        {
            rig.Gateway.RaiseValuesInvalidated();
            rig.Gateway.RaiseCellsChanged(K33);
            refreshesDuringClick = rig.View.RefreshHoverCount;
        };

        rig.View.Click(rig.View.Node("K$31"));

        Assert.Equal(0, refreshesDuringClick);
        Assert.Equal(0, rig.View.RefreshHoverCount);
        Assert.Equal("K$31", rig.View.GoldText);

        rig.Scheduler.RunPosted();
        Assert.Equal(1, rig.View.RefreshHoverCount);
        Assert.Equal("K$31", rig.View.GoldText);
    }

    [Fact]
    public void WithASchedulerThatRunsPostsInlineEventsStillWaitForTheEndOfTheCommand()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Scheduler.Inline = true;
        int refreshesDuringClick = -1;
        rig.Gateway.DuringNavigate = () =>
        {
            rig.Gateway.RaiseValuesInvalidated();
            refreshesDuringClick = rig.View.RefreshHoverCount;
        };

        rig.View.Click(rig.View.Node("K$31"));

        Assert.Equal(0, refreshesDuringClick); // not run inside Goto
        Assert.Equal(1, rig.View.RefreshHoverCount); // but right after the command
        Assert.Equal("K$31", rig.View.GoldText);
    }

    [Fact]
    public void TheSelectionEchoOfADrillTargetDoesNotPushAgain()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Click("K$31");

        rig.View.DoubleClick(rig.View.Node("K$31"));
        rig.Gateway.RaiseUserSelectionChanged(K31);
        Assert.Equal(K31, rig.Snapshot.Root); // the drill itself already ran
        rig.Settle();

        // The echo equals the new root: it only clears gold/restores, it does not push again.
        Assert.Equal(K31, rig.Snapshot.Root);
        Assert.Equal(new[] { K33 }, rig.Snapshot.Stack);
    }

    [Fact]
    public void AViewCommandRaisedInsideAnotherCommandIsPostedNotRunInline()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        var hundred = rig.View.Node("Hundred");
        rig.Gateway.DuringNavigate = () =>
        {
            rig.Gateway.DuringNavigate = null;
            rig.View.Click(hundred); // e.g. a message pump inside Goto delivers another click
        };

        rig.View.Click(rig.View.Node("K$31"));

        Assert.Single(rig.Gateway.NavigateCalls); // the nested click did not run yet
        Assert.Equal("K$31", rig.View.GoldText);

        rig.Settle();

        Assert.Equal(2, rig.Gateway.NavigateCalls.Count);
        Assert.Equal("Hundred", rig.View.GoldText);
    }
}
