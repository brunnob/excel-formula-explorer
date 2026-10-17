using FormulaExplorer.AddIn.Excel;
using FormulaExplorer.Core.Syntax;
using FormulaExplorer.Core.Values;

namespace FormulaExplorer.AddIn.Tests.Sessions;

/// <summary>Start(), Load(): blank phase, parsing, colors, precedent boxes, non-formula cells.</summary>
public class SessionLoadTests
{
    [Fact]
    public void StartBlanksTheTextFor150MsThenShowsTheFormula()
    {
        var rig = Rig.VideoModel();

        rig.Session.Start();

        // Immediately: Back updated first, text cleared (wait cursor), nothing else yet.
        Assert.True(rig.Session.IsOpen);
        Assert.True(rig.View.IsLoading);
        Assert.Null(rig.View.Text);
        Assert.Equal(new[] { "Back:False", "Loading" }, rig.View.Log);
        Assert.Empty(rig.Overlay.ShowCalls);

        rig.Scheduler.Advance(149);
        Assert.True(rig.View.IsLoading);

        rig.Scheduler.Advance(1);
        Assert.False(rig.View.IsLoading);
        Assert.Equal(Rig.K33Formula, rig.View.Text);
        Assert.True(rig.View.Display!.Interactive);
        Assert.Equal(new[] { "Back:False", "Loading", "Formula:" + Rig.K33Formula }, rig.View.Log);
    }

    [Fact]
    public void StartDoesNotClearTheOverlayAndShowsBoxesOnlyAfterTheTextAppears()
    {
        var rig = Rig.VideoModel();

        rig.Session.Start();
        rig.Scheduler.Advance(100);
        Assert.Empty(rig.Overlay.ShowCalls);
        Assert.Equal(0, rig.Overlay.ClearCalls);

        rig.Scheduler.Advance(50);
        var (root, boxes) = Assert.Single(rig.Overlay.ShowCalls);
        Assert.Equal(Rig.Loc("K33"), root);
        Assert.NotEmpty(boxes);
        Assert.Equal(0, rig.Overlay.ClearCalls);
    }

    [Fact]
    public void PrecedentBoxesOfTheVideoRootAreOnlyOnTheRootSheet()
    {
        var rig = Rig.VideoModel();
        rig.Start();

        // $I27 blue (0), Hundred green (1, on another sheet: no box), K$31 purple (2).
        Assert.Equal(new[] { "Model!$I$27#0000FF", "Model!$K$31#9900CC" }, rig.Overlay.Boxes.Select(Rig.Box));
        Assert.Equal(rig.Overlay.Boxes, rig.Snapshot.Boxes);
    }

    [Fact]
    public void ReferenceColorsFollowTheRangeFinderPalette()
    {
        var rig = Rig.VideoModel();
        rig.Start();

        var colors = rig.View.Display!.Colors;
        Assert.Equal(new[] { "$I27", "Hundred", "K$31" },
            colors.Nodes.Select(n => n.Node.GetText(Rig.K33Formula)));
        Assert.Equal(new uint[] { 0x0000FF, 0x008000, 0x9900CC }, colors.Nodes.Select(n => n.Rgb));
    }

    [Fact]
    public void NamesAreColoredOnlyWhenTheyResolveToARange()
    {
        var rig = new Rig();
        rig.Gateway.Formula("S", "A1", "=Hundred+Missing+B1")
            .Range("Hundred", "Other", "B1", false, 100.0)
            .Value("Missing", new EvalResult.Error(ExcelError.Name))
            .Range("B1", "S", "B1", false, 5.0);
        rig.Gateway.Active = FakeGateway.Loc("S", "A1");

        rig.Start();

        Assert.Equal(new[] { "Hundred", "B1" },
            rig.View.Display!.Colors.Nodes.Select(n => n.Node.GetText("=Hundred+Missing+B1")));
    }

    [Fact]
    public void ErrorReferencesAreNotColored()
    {
        var rig = new Rig();
        const string formula = "=Sheet1!#REF!+B1";
        rig.Gateway.Formula("S", "A1", formula).Range("B1", "S", "B1", false, 5.0);
        rig.Gateway.Active = FakeGateway.Loc("S", "A1");

        rig.Start();

        Assert.Equal(new[] { "B1" }, rig.View.Display!.Colors.Nodes.Select(n => n.Node.GetText(formula)));
    }

    [Fact]
    public void BoxesAreDedupedByLocationAndSkipOtherSheets()
    {
        var rig = new Rig();
        const string formula = "=A1+MyName+Other!B2+B3+A1";
        rig.Gateway.Formula("S", "C1", formula)
            .Range("A1", "S", "A1", false, 1.0)
            .Range("MyName", "S", "A1", false, 1.0)   // same cell as A1, different color slot
            .Range("Other!B2", "Other", "B2", false, 2.0)
            .Range("B3", "S", "B3", false, 3.0);
        rig.Gateway.Active = FakeGateway.Loc("S", "C1");

        rig.Start();

        // A1 blue(0), MyName green(1) duplicates A1's location, Other!B2 purple(2) is another sheet, B3 maroon(3).
        Assert.Equal(new[] { "S!$A$1#0000FF", "S!$B$3#800000" }, rig.Overlay.Boxes.Select(Rig.Box));
    }

    [Fact]
    public void RangeReferencesProduceOneBoxForTheWholeRange()
    {
        var rig = new Rig();
        const string formula = "=SUM(A1:B2)";
        rig.Gateway.Formula("S", "C1", formula).Range("A1:B2", "S", "A1:B2", false, 1.0, 2.0, 3.0, 4.0);
        rig.Gateway.Active = FakeGateway.Loc("S", "C1");

        rig.Start();

        Assert.Equal(new[] { "S!$A$1:$B$2#0000FF" }, rig.Overlay.Boxes.Select(Rig.Box));
    }

    [Fact]
    public void ThreeDReferencesAreEvaluatedAsTheirFirstSheet()
    {
        var rig = new Rig();
        const string formula = "=SUM(Sheet2:Sheet4!A1)+SUM('Jo''s:Z'!B2:C3)+SUM('[Book 1.xlsx]Sheet 1:Sheet 3'!D4)";
        rig.Gateway.Formula("S", "C1", formula);
        rig.Gateway.Active = FakeGateway.Loc("S", "C1");

        rig.Start();

        var expressions = rig.Gateway.EvaluateCalls.Select(c => c.Expression).ToArray();
        Assert.Contains("'Sheet2'!A1", expressions);
        Assert.Contains("'Jo''s'!B2:C3", expressions);
        Assert.Contains("'[Book 1.xlsx]Sheet 1'!D4", expressions);
    }

    [Fact]
    public void ANonFormulaCellIsShownBlackAndNonInteractive()
    {
        var rig = Rig.VideoModel();
        rig.Gateway.Active = Rig.Loc("K19");

        rig.Start();

        var display = rig.View.Display!;
        Assert.Equal("0.05", rig.View.Text);
        Assert.False(display.Interactive);
        Assert.Empty(display.Colors.Nodes);
        Assert.All(Enumerable.Range(0, 6), i => Assert.Null(display.HitMap.NodeAt(i)));
        Assert.Empty(rig.Overlay.Boxes);
        Assert.Equal(0, rig.Gateway.Calls("Evaluate"));
    }

    [Fact]
    public void TextThatLooksLikeAFormulaButIsAConstantIsNotInteractive()
    {
        var rig = new Rig();
        rig.Gateway.Constant("S", "A1", "=A1+B1"); // e.g. entered with a leading apostrophe
        rig.Gateway.Active = FakeGateway.Loc("S", "A1");

        rig.Start();

        Assert.Equal("=A1+B1", rig.View.Text);
        Assert.False(rig.View.Display!.Interactive);
        Assert.Empty(rig.View.Display.Colors.Nodes);
        Assert.Null(rig.View.Display.HitMap.NodeAt(1));
        Assert.Equal(0, rig.Gateway.Calls("Evaluate"));
    }

    [Fact]
    public void AnEmptyCellShowsEmptyText()
    {
        var rig = Rig.VideoModel();
        rig.Gateway.Active = Rig.Loc("I19");

        rig.Start();

        Assert.Equal(string.Empty, rig.View.Text);
        Assert.False(rig.View.Display!.Interactive);
    }

    [Fact]
    public void StartWithoutActiveCellShowsEmptyNonInteractiveTextImmediately()
    {
        var rig = Rig.VideoModel();
        rig.Gateway.Active = null;

        rig.Session.Start();

        Assert.True(rig.Session.IsOpen);
        Assert.False(rig.View.IsLoading);
        Assert.Equal(string.Empty, rig.View.Text);
        Assert.False(rig.View.Display!.Interactive);
        Assert.Empty(rig.Overlay.ShowCalls);
        Assert.False(rig.View.BackEnabled);
        Assert.Null(rig.Snapshot.Root);
    }

    [Fact]
    public void AfterStartingWithoutAnActiveCellTheFirstSelectionBecomesTheRootWithoutPushing()
    {
        var rig = Rig.VideoModel();
        rig.Gateway.Active = null;
        rig.Start();

        rig.Gateway.UserSelects(Rig.Loc("K33"));
        rig.Settle();

        Assert.Equal(Rig.Loc("K33"), rig.Snapshot.Root);
        Assert.Empty(rig.Snapshot.Stack);
        Assert.False(rig.View.BackEnabled);
        Assert.Equal(Rig.K33Formula, rig.View.Text);
    }

    [Fact]
    public void StartTwiceSubscribesOnce()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Session.Start();
        rig.Settle();

        rig.Click("K$31");

        Assert.Single(rig.Gateway.NavigateCalls);
        Assert.Equal(1, rig.View.ShowFormulaCount);
    }

    [Fact]
    public void SnapshotDescribesTheSession()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Click("K$31");

        var snapshot = rig.Snapshot;

        Assert.True(snapshot.IsOpen);
        Assert.Equal(Rig.Loc("K33"), snapshot.Root);
        Assert.Equal(Rig.K33Formula, snapshot.DisplayText);
        Assert.Empty(snapshot.Stack);
        Assert.Equal(Rig.K33Formula.IndexOf("K$31", StringComparison.Ordinal), snapshot.GoldStart);
        Assert.Equal(4, snapshot.GoldLength);
        Assert.Equal(Rig.Loc("K31"), snapshot.GoldTarget);
        Assert.True(snapshot.BackEnabled);
        Assert.Equal(2, snapshot.Boxes.Count);
    }

    [Fact]
    public void SnapshotOfASessionThatWasNeverStartedIsClosed()
    {
        var rig = Rig.VideoModel();

        var snapshot = rig.Snapshot;

        Assert.False(snapshot.IsOpen);
        Assert.Null(snapshot.Root);
        Assert.Equal(string.Empty, snapshot.DisplayText);
        Assert.Null(snapshot.GoldStart);
        Assert.False(snapshot.BackEnabled);
    }
}
