using FormulaExplorer.AddIn.Excel;
using FormulaExplorer.Core.Values;

namespace FormulaExplorer.AddIn.Tests.Sessions;

/// <summary>Single click (gold), double click (drill down), blocked and failed navigation.</summary>
public class SessionNavigationTests
{
    private static readonly CellLocation K33 = Rig.Loc("K33");
    private static readonly CellLocation K31 = Rig.Loc("K31");

    // ---- single click ------------------------------------------------------------------------------------------

    [Fact]
    public void ClickingALinkNavigatesKeepingTheRootAndMakesItGold()
    {
        var rig = Rig.VideoModel();
        rig.Start();

        rig.Click("K$31");

        var call = Assert.Single(rig.Gateway.NavigateCalls);
        Assert.Equal((K31, (CellLocation?)K33), call);
        Assert.Equal("K$31", rig.View.GoldText);
        Assert.Equal(K31, rig.Snapshot.GoldTarget);
        Assert.True(rig.View.BackEnabled);
    }

    [Fact]
    public void AClickNeverChangesTheDisplayedFormulaNorTheStack()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        int formulas = rig.View.ShowFormulaCount;

        rig.Click("K$31");
        rig.Click("Hundred");
        rig.Click("$I27");

        Assert.Equal(formulas, rig.View.ShowFormulaCount);
        Assert.Equal(1, rig.View.ShowLoadingCount);
        Assert.Equal(Rig.K33Formula, rig.View.Text);
        Assert.Equal(K33, rig.Snapshot.Root);
        Assert.Empty(rig.Snapshot.Stack);
    }

    [Fact]
    public void OnlyOneNodeIsGoldAtATime()
    {
        var rig = Rig.VideoModel();
        rig.Start();

        rig.Click("K$31");
        rig.Click("Hundred");

        Assert.Equal("Hundred", rig.View.GoldText);
        Assert.Equal(Rig.Loc("B1", Rig.Assumptions), rig.Snapshot.GoldTarget);
        Assert.Equal(new[] { "Gold:K$31", "Gold:Hundred" }, rig.View.Log.Where(l => l.StartsWith("Gold:")));
    }

    [Fact]
    public void AClickOnANodeThatIsNotALinkIsIgnored()
    {
        var rig = Rig.VideoModel();
        rig.Gateway.Value("(($I27/Hundred)+1)^(K$31)", new EvalResult.Number(1.0076));
        rig.Start();

        rig.View.Click(rig.View.Node("(($I27/Hundred)+1)^(K$31)"));
        rig.View.Click(rig.View.Display!.Formula.Root!);
        rig.Settle();

        Assert.Empty(rig.Gateway.NavigateCalls);
        Assert.Equal("(none)", rig.View.GoldText);
        Assert.False(rig.View.BackEnabled);
    }

    [Fact]
    public void ClickingTheGoldNodeAgainNavigatesAgainAndStaysGold()
    {
        var rig = Rig.VideoModel();
        rig.Start();

        rig.Click("K$31");
        rig.Click("K$31");

        Assert.Equal(2, rig.Gateway.NavigateCalls.Count);
        Assert.Equal("K$31", rig.View.GoldText);
        Assert.Empty(rig.Snapshot.Stack);
    }

    [Fact]
    public void ClickOnAMultiCellRangeFunctionNavigatesToTheRange()
    {
        var rig = Rig.VideoModel();
        rig.Gateway.Active = Rig.Loc("J31");
        rig.Start();

        rig.Click("OFFSET($I$19,0,0,1,Aqui_Per)");

        Assert.Equal((Rig.Loc("I19:J19"), (CellLocation?)Rig.Loc("J31")), Assert.Single(rig.Gateway.NavigateCalls));
        Assert.Equal("OFFSET($I$19,0,0,1,Aqui_Per)", rig.View.GoldText);
    }

    [Fact]
    public void ClickFromADisplayThatWasReplacedIsIgnored()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        var stale = rig.View.Node("K$31");

        rig.Gateway.UserSelects(Rig.Loc("K31"));
        rig.Settle();
        rig.View.Click(stale);
        rig.Settle();

        Assert.Empty(rig.Gateway.NavigateCalls);
        Assert.Equal("(none)", rig.View.GoldText);
    }

    // ---- blocked / failed --------------------------------------------------------------------------------------

    [Fact]
    public void BlockedNavigationShowsAPostedMessageAndNoGold()
    {
        var rig = Rig.VideoModel();
        rig.Gateway.NavigationResults[Rig.Loc("B1", Rig.Assumptions)] = new NavigationResult(
            NavigationStatus.Blocked, "The sheet 'Assumptions' is hidden and the workbook structure is protected.");
        rig.Start();

        rig.View.Click(rig.View.Node("Hundred"));

        // Not inside the click handler.
        Assert.Empty(rig.View.Messages);
        rig.Scheduler.RunPosted();
        Assert.Equal(new[] { "The sheet 'Assumptions' is hidden and the workbook structure is protected." }, rig.View.Messages);
        Assert.Equal("(none)", rig.View.GoldText);
        Assert.Null(rig.Snapshot.GoldStart);
        Assert.False(rig.View.BackEnabled);
    }

    [Fact]
    public void BlockedNavigationKeepsThePreviousGold()
    {
        var rig = Rig.VideoModel();
        rig.Gateway.NavigationResults[Rig.Loc("B1", Rig.Assumptions)] = new NavigationResult(NavigationStatus.Blocked, "blocked");
        rig.Start();
        rig.Click("K$31");

        rig.Click("Hundred");

        Assert.Equal("K$31", rig.View.GoldText);
        Assert.Equal(K31, rig.Snapshot.GoldTarget);
        Assert.True(rig.View.BackEnabled);
    }

    [Fact]
    public void FailedNavigationLeavesNoGold()
    {
        var rig = Rig.VideoModel();
        rig.Gateway.NavigationResults[K31] = new NavigationResult(NavigationStatus.Failed, null);
        rig.Start();

        rig.Click("K$31");

        Assert.Equal("(none)", rig.View.GoldText);
        Assert.Empty(rig.View.Messages);
        Assert.False(rig.View.BackEnabled);
    }

    [Fact]
    public void FailedNavigationClearsAnEarlierGoldAndShowsItsMessage()
    {
        var rig = Rig.VideoModel();
        rig.Gateway.NavigationResults[K31] = new NavigationResult(NavigationStatus.Failed, "The cell no longer exists.");
        rig.Start();
        rig.Click("Hundred");
        Assert.Equal("Hundred", rig.View.GoldText);

        rig.Click("K$31");

        Assert.Equal("(none)", rig.View.GoldText);
        Assert.Null(rig.Snapshot.GoldStart);
        Assert.Equal(new[] { "The cell no longer exists." }, rig.View.Messages);
        Assert.False(rig.View.BackEnabled);
    }

    // ---- double click ------------------------------------------------------------------------------------------

    [Fact]
    public void DoubleClickOnASingleCellWithAFormulaDrillsDown()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Click("K$31");

        rig.DoubleClick("K$31");

        var snapshot = rig.Snapshot;
        Assert.Equal(K31, snapshot.Root);
        Assert.Equal(new[] { K33 }, snapshot.Stack);
        Assert.Null(snapshot.GoldStart);
        Assert.Equal(Rig.K31Formula, rig.View.Text);
        Assert.True(rig.View.BackEnabled);
        Assert.Equal("(none)", rig.View.GoldText);
    }

    [Fact]
    public void DrillDownRedrawsTheBoxesForTheNewRoot()
    {
        var rig = Rig.VideoModel();
        rig.Start();

        rig.DoubleClick("K$31");

        var (root, boxes) = rig.Overlay.ShowCalls[^1];
        Assert.Equal(K31, root);
        Assert.Equal(new[] { "Model!$J$31#0000FF", "Model!$K$19#008000" }, boxes.Select(Rig.Box));
    }

    [Fact]
    public void DrillDownKeepsTheOldBordersUntilTheNewTextAppears()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        var oldBoxes = rig.Overlay.Boxes;
        rig.Click("K$31");

        rig.View.DoubleClick(rig.View.Node("K$31"));
        rig.Scheduler.RunPosted();

        Assert.True(rig.View.IsLoading);
        Assert.Same(oldBoxes, rig.Overlay.Boxes);
        Assert.Equal(0, rig.Overlay.ClearCalls);

        rig.Scheduler.Advance(150);
        Assert.False(rig.View.IsLoading);
        Assert.NotSame(oldBoxes, rig.Overlay.Boxes);
    }

    [Fact]
    public void BackIsEnabledBeforeTheNewTextShows()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        int mark = rig.View.Log.Count;

        rig.View.DoubleClick(rig.View.Node("K$31"));
        rig.Settle();

        var log = rig.View.Log.Skip(mark).ToList();
        int enabled = log.IndexOf("Back:True");
        int loading = log.IndexOf("Loading");
        int shown = log.IndexOf("Formula:" + Rig.K31Formula);
        Assert.True(enabled >= 0 && enabled < loading && loading < shown, string.Join(" | ", log));
        Assert.Equal("Back:True", log[loading - 1]);
    }

    [Fact]
    public void DoubleClickWithoutAPreviousClickNavigatesFirstThenDrills()
    {
        var rig = Rig.VideoModel();
        rig.Start();

        rig.View.RaiseDoubleClickOnly(rig.View.Node("K$31"));
        rig.Settle();

        Assert.Equal((K31, (CellLocation?)K33), Assert.Single(rig.Gateway.NavigateCalls));
        Assert.Equal(K31, rig.Snapshot.Root);
        Assert.Equal(new[] { K33 }, rig.Snapshot.Stack);
    }

    [Fact]
    public void DoubleClickOnAnotherNodeThanTheGoldClicksFirst()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Click("Hundred");

        rig.View.RaiseDoubleClickOnly(rig.View.Node("K$31"));
        rig.Settle();

        Assert.Equal(K31, rig.Snapshot.Root);
        Assert.Equal(2, rig.Gateway.NavigateCalls.Count);
    }

    [Fact]
    public void DoubleClickOnARangeOnlyNavigates()
    {
        var rig = Rig.VideoModel();
        rig.Gateway.Active = Rig.Loc("J31");
        rig.Start();

        rig.DoubleClick("OFFSET($I$19,0,0,1,Aqui_Per)");

        Assert.Equal(Rig.Loc("J31"), rig.Snapshot.Root);
        Assert.Empty(rig.Snapshot.Stack);
        Assert.Equal("OFFSET($I$19,0,0,1,Aqui_Per)", rig.View.GoldText);
        Assert.Equal(1, rig.View.ShowFormulaCount);
        Assert.Equal(Rig.J31Formula, rig.View.Text);
    }

    [Fact]
    public void DoubleClickOnACellWithoutAFormulaOnlyNavigates()
    {
        var rig = Rig.VideoModel();
        rig.Gateway.Active = K31;
        rig.Start();

        rig.DoubleClick("K19");

        Assert.Equal(K31, rig.Snapshot.Root);
        Assert.Empty(rig.Snapshot.Stack);
        Assert.Equal("K19", rig.View.GoldText);
        Assert.Equal(Rig.Loc("K19"), rig.Snapshot.GoldTarget);
        Assert.Equal(1, rig.View.ShowFormulaCount);
    }

    [Fact]
    public void DoubleClickWhenTheTargetCellVanishedStaysOnTheRootWithGold()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Click("K$31");
        rig.Gateway.Cells.Remove(K31);

        rig.DoubleClick("K$31");

        Assert.Equal(K33, rig.Snapshot.Root);
        Assert.Empty(rig.Snapshot.Stack);
        Assert.Equal("K$31", rig.View.GoldText);
    }

    [Fact]
    public void DrillingIntoTheRootItselfDoesNotPushTheSameLocationTwice()
    {
        var rig = new Rig();
        const string formula = "=A1+1";
        rig.Gateway.Formula("S", "A1", formula).Range("A1", "S", "A1", true, 1.0);
        rig.Gateway.Active = FakeGateway.Loc("S", "A1");
        rig.Start();

        rig.DoubleClick("A1");
        rig.DoubleClick("A1");

        Assert.Equal(new[] { FakeGateway.Loc("S", "A1") }, rig.Snapshot.Stack);
    }

    [Fact]
    public void DoubleClickOnAThreeDReferenceDrillsIntoItsFirstSheet()
    {
        var rig = new Rig();
        rig.Gateway.Formula("S", "A1", "=SUM(Sheet2:Sheet4!B2)")
            .Formula("Sheet2", "B2", "=7")
            .Range("'Sheet2'!B2", "Sheet2", "B2", true, 7.0);
        rig.Gateway.Active = FakeGateway.Loc("S", "A1");
        rig.Start();

        rig.DoubleClick("Sheet2:Sheet4!B2");

        Assert.Equal(FakeGateway.Loc("Sheet2", "B2"), rig.Snapshot.Root);
        Assert.Equal(new[] { FakeGateway.Loc("S", "A1") }, rig.Snapshot.Stack);
    }
}
