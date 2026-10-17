using FormulaExplorer.AddIn.Excel;

namespace FormulaExplorer.AddIn.Tests.Sessions;

/// <summary>Back button rules and the full sequence of the original demo video.</summary>
public class SessionBackTests
{
    private static readonly CellLocation K33 = Rig.Loc("K33");
    private static readonly CellLocation K31 = Rig.Loc("K31");
    private static readonly CellLocation J31 = Rig.Loc("J31");

    /// <summary>
    /// 2:00-2:50 of the official video: K33 → click K$31, Hundred, $I27, K$31 → double click (drill) to K31 → click K19,
    /// J31 → drill to J31 → click OFFSET → Back (gold cleared, J31 selected) → Back (K31) → Back (K33).
    /// </summary>
    [Fact]
    public void OriginalVideoSequence()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        Assert.Equal(Rig.K33Formula, rig.Text);
        Assert.False(rig.View.BackEnabled);

        // Clicks only move the gold: nothing is pushed, the formula does not change.
        rig.Click("K$31");
        rig.Click("Hundred");
        rig.Click("$I27");
        rig.Click("K$31");
        Assert.Equal("K$31", rig.View.GoldText);
        Assert.Empty(rig.Snapshot.Stack);
        Assert.True(rig.View.BackEnabled);
        Assert.All(rig.Gateway.NavigateCalls, c => Assert.Equal(K33, c.AlsoKeep));

        // Double click on the gold K$31: drill down.
        rig.DoubleClick("K$31");
        Assert.Equal(K31, rig.Snapshot.Root);
        Assert.Equal(new[] { K33 }, rig.Snapshot.Stack);
        Assert.Equal(Rig.K31Formula, rig.Text);
        Assert.Equal("(none)", rig.View.GoldText);

        rig.Click("K19");
        rig.Click("J31");
        Assert.Equal("J31", rig.View.GoldText);
        Assert.Equal(new[] { K33 }, rig.Snapshot.Stack);

        rig.DoubleClick("J31");
        Assert.Equal(J31, rig.Snapshot.Root);
        Assert.Equal(new[] { K33, K31 }, rig.Snapshot.Stack);
        Assert.Equal(Rig.J31Formula, rig.Text);

        // OFFSET returns a range: the click selects I19:J19 and turns it gold.
        rig.Click("OFFSET($I$19,0,0,1,Aqui_Per)");
        var offsetCall = rig.Gateway.NavigateCalls[^1];
        Assert.Equal(Rig.Loc("I19:J19"), offsetCall.Target);
        Assert.Equal(J31, offsetCall.AlsoKeep);
        Assert.Equal("OFFSET($I$19,0,0,1,Aqui_Per)", rig.View.GoldText);

        // Back 1: gold cleared, J31 selected again; the formula stays.
        int formulas = rig.View.ShowFormulaCount;
        rig.Back();
        Assert.Equal((J31, (CellLocation?)null), rig.Gateway.NavigateCalls[^1]);
        Assert.Equal("(none)", rig.View.GoldText);
        Assert.Equal(J31, rig.Snapshot.Root);
        Assert.Equal(Rig.J31Formula, rig.Text);
        Assert.Equal(formulas, rig.View.ShowFormulaCount);
        Assert.Equal(new[] { K33, K31 }, rig.Snapshot.Stack);
        Assert.True(rig.View.BackEnabled);

        // Back 2: pop to K31.
        rig.Back();
        Assert.Equal((K31, (CellLocation?)null), rig.Gateway.NavigateCalls[^1]);
        Assert.Equal(K31, rig.Snapshot.Root);
        Assert.Equal(Rig.K31Formula, rig.Text);
        Assert.Equal(new[] { K33 }, rig.Snapshot.Stack);
        Assert.True(rig.View.BackEnabled);

        // Back 3: pop to K33, nothing left to go back to.
        rig.Back();
        Assert.Equal((K33, (CellLocation?)null), rig.Gateway.NavigateCalls[^1]);
        Assert.Equal(K33, rig.Snapshot.Root);
        Assert.Equal(Rig.K33Formula, rig.Text);
        Assert.Empty(rig.Snapshot.Stack);
        Assert.False(rig.View.BackEnabled);
        Assert.Equal(new[] { "Model!$I$27#0000FF", "Model!$K$31#9900CC" }, rig.Overlay.Boxes.Select(Rig.Box));

        // Nothing more to do.
        int navigations = rig.Gateway.NavigateCalls.Count;
        rig.Back();
        Assert.Equal(navigations, rig.Gateway.NavigateCalls.Count);
        Assert.False(rig.View.BackEnabled);
        Assert.Equal(0, rig.Overlay.ClearCalls);
    }

    [Fact]
    public void BackIsDisabledAtTheStart()
    {
        var rig = Rig.VideoModel();
        rig.Start();

        Assert.False(rig.View.BackEnabled);
        Assert.False(rig.Snapshot.BackEnabled);
    }

    [Fact]
    public void BackWithGoldReselectsTheRootAndClearsTheGold()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Click("K$31");

        rig.Back();

        Assert.Equal((K33, (CellLocation?)null), rig.Gateway.NavigateCalls[^1]);
        Assert.Equal("(none)", rig.View.GoldText);
        Assert.Null(rig.Snapshot.GoldStart);
        Assert.False(rig.View.BackEnabled); // empty stack
        Assert.Equal(1, rig.View.ShowFormulaCount);
        Assert.Equal(K33, rig.Gateway.Active);
    }

    [Fact]
    public void BackWithGoldKeepsBackEnabledWhenTheStackIsNotEmpty()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.DoubleClick("K$31");
        rig.Click("K19");

        rig.Back();

        Assert.Equal(K31, rig.Snapshot.Root);
        Assert.True(rig.View.BackEnabled);
        Assert.Equal(new[] { K33 }, rig.Snapshot.Stack);
    }

    [Fact]
    public void BackPopsAndBlanksTheTextBeforeShowingTheOldFormula()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.DoubleClick("K$31");
        int mark = rig.View.Log.Count;

        rig.View.PressBack();

        // Back is already updated, text blank; the old formula comes 150 ms later.
        Assert.False(rig.View.BackEnabled);
        Assert.True(rig.View.IsLoading);
        rig.Scheduler.Advance(149);
        Assert.True(rig.View.IsLoading);
        rig.Scheduler.Advance(1);
        Assert.Equal(Rig.K33Formula, rig.View.Text);
        var log = rig.View.Log.Skip(mark).ToArray();
        Assert.Equal(new[] { "Back:False", "Loading", "Formula:" + Rig.K33Formula }, log);
    }

    [Fact]
    public void BackDiscardsEntriesWhoseCellNoLongerExists()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.DoubleClick("K$31");
        rig.DoubleClick("J31");
        Assert.Equal(new[] { K33, K31 }, rig.Snapshot.Stack);
        rig.Gateway.Cells.Remove(K31);

        rig.Back();

        Assert.Equal(K33, rig.Snapshot.Root);
        Assert.Empty(rig.Snapshot.Stack);
        Assert.Equal((K33, (CellLocation?)null), Assert.Single(rig.Gateway.NavigateCalls.Skip(2)));
        Assert.False(rig.View.BackEnabled);
    }

    [Fact]
    public void BackWithOnlyDeadEntriesChangesNothingAndDisablesBack()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.DoubleClick("K$31");
        rig.Gateway.Cells.Remove(K33);

        rig.Back();

        Assert.Equal(K31, rig.Snapshot.Root);
        Assert.Empty(rig.Snapshot.Stack);
        Assert.False(rig.View.BackEnabled);
        Assert.Equal(Rig.K31Formula, rig.View.Text);
    }

    [Fact]
    public void BackToAnEntryThatCannotBeNavigatedToIsDiscarded()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.DoubleClick("K$31");
        rig.DoubleClick("J31");
        rig.Gateway.NavigationResults[K31] = new NavigationResult(NavigationStatus.Failed, null);

        rig.Back();

        Assert.Equal(K33, rig.Snapshot.Root);
        Assert.Empty(rig.Snapshot.Stack);
    }

    [Fact]
    public void BackToABlockedEntryShowsAMessageAndKeepsTheStack()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.DoubleClick("K$31");
        rig.Gateway.NavigationResults[K33] = new NavigationResult(NavigationStatus.Blocked, "blocked!");

        rig.Back();

        Assert.Equal(K31, rig.Snapshot.Root);
        Assert.Equal(new[] { K33 }, rig.Snapshot.Stack);
        Assert.Equal(new[] { "blocked!" }, rig.View.Messages);
        Assert.True(rig.View.BackEnabled);
    }

    [Fact]
    public void BackAfterTheRootWasDeletedPopsTheStack()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.DoubleClick("K$31");
        rig.Gateway.Cells.Remove(K31);
        rig.Gateway.RaiseCellsChanged(K31);
        rig.Settle();
        Assert.Null(rig.Snapshot.Root);
        Assert.True(rig.View.BackEnabled);

        rig.Back();

        Assert.Equal(K33, rig.Snapshot.Root);
        Assert.Equal(Rig.K33Formula, rig.View.Text);
        Assert.False(rig.View.BackEnabled);
    }

    [Fact]
    public void ManualSelectionsPushAndBackRetracesThem()
    {
        var rig = Rig.VideoModel();
        rig.Start();

        rig.Gateway.UserSelects(K31);
        rig.Settle();
        rig.Gateway.UserSelects(J31);
        rig.Settle();
        Assert.Equal(new[] { K33, K31 }, rig.Snapshot.Stack);
        Assert.True(rig.View.BackEnabled);

        rig.Back();
        rig.Back();

        Assert.Equal(K33, rig.Snapshot.Root);
        Assert.Empty(rig.Snapshot.Stack);
        Assert.False(rig.View.BackEnabled);
    }

    [Fact]
    public void DrillAfterBackStartsANewBranch()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.DoubleClick("K$31");
        rig.Back();

        rig.DoubleClick("K$31");

        Assert.Equal(K31, rig.Snapshot.Root);
        Assert.Equal(new[] { K33 }, rig.Snapshot.Stack);
    }

    [Fact]
    public void BackWhenNothingIsOpenAnywhereDoesNothing()
    {
        var rig = Rig.VideoModel();
        rig.Gateway.Active = null;
        rig.Start();

        rig.Back();

        Assert.Empty(rig.Gateway.NavigateCalls);
        Assert.False(rig.View.BackEnabled);
    }
}
