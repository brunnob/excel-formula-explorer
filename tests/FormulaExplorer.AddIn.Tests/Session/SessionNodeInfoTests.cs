using FormulaExplorer.AddIn.Session;
using FormulaExplorer.Core.Syntax;
using FormulaExplorer.Core.Values;

namespace FormulaExplorer.AddIn.Tests.Sessions;

/// <summary>GetNodeInfo: lazy cached evaluation, link rules, tooltip text, invalidation, busy Excel.</summary>
public class SessionNodeInfoTests
{
    private static readonly NodeInfo Nothing = new(false, null);

    [Fact]
    public void NotOpenReturnsNoInfoBeforeStartAndAfterClose()
    {
        var rig = Rig.VideoModel();
        var stranger = FormulaParser.Parse("=A1").Root!;

        Assert.Equal(Nothing, rig.Session.GetNodeInfo(stranger));

        rig.Start();
        var node = rig.View.Node("K$31");
        Assert.True(rig.Session.GetNodeInfo(node).IsLink);

        rig.Session.Close();
        Assert.Equal(Nothing, rig.Session.GetNodeInfo(node));
    }

    [Fact]
    public void ALetVariableIsNeverEvaluatedAsAWorkbookName()
    {
        var rig = Rig.VideoModel();
        rig.Gateway.Formula(Rig.Model, "K40", "=LET(Hundred,5,Hundred*2)", new EvalResult.Number(10));
        rig.Gateway.Active = Rig.Loc("K40");
        rig.Start();

        var local = rig.View.Display!.Formula.Nodes.OfType<NameNode>().Last();

        Assert.True(local.IsLocal);
        Assert.Equal(Nothing, rig.Session.GetNodeInfo(local));
        Assert.DoesNotContain(rig.Gateway.EvaluateCalls, c => c.Expression == "Hundred");
    }

    [Fact]
    public void ANodeThatIsNotPartOfTheDisplayedFormulaHasNoInfo()
    {
        var rig = Rig.VideoModel();
        rig.Start();

        Assert.Equal(Nothing, rig.Session.GetNodeInfo(FormulaParser.Parse("=K$31").Root!.Body!));
    }

    [Fact]
    public void TheEqualsSignShowsTheValueOfTheRootCell()
    {
        var rig = Rig.VideoModel();
        rig.Start();

        var info = rig.Session.GetNodeInfo(rig.View.Display!.Formula.Root!);

        Assert.Equal(new NodeInfo(false, "1.00760985118373"), info);
    }

    [Fact]
    public void TheRootValueIsReReadAfterAnInvalidation()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        var root = rig.View.Display!.Formula.Root!;
        Assert.Equal("1.00760985118373", rig.Session.GetNodeInfo(root).TooltipText);

        var k33 = Rig.Loc("K33");
        rig.Gateway.Cells[k33] = rig.Gateway.Cells[k33] with { Value = new EvalResult.Number(2.5) };
        rig.Gateway.RaiseValuesInvalidated();
        rig.Settle();

        Assert.Equal("2.5", rig.Session.GetNodeInfo(root).TooltipText);
    }

    [Fact]
    public void AReferenceToACellIsALinkWithItsValue()
    {
        var rig = Rig.VideoModel();
        rig.Start();

        Assert.Equal(new NodeInfo(true, "0.126"), rig.Session.GetNodeInfo(rig.View.Node("K$31")));
        Assert.Equal(new NodeInfo(true, "100"), rig.Session.GetNodeInfo(rig.View.Node("Hundred")));
    }

    [Fact]
    public void ARangeFunctionIsALinkWithTheContainingText()
    {
        var rig = Rig.VideoModel();
        rig.Gateway.Active = Rig.Loc("J31");
        rig.Start();

        var info = rig.Session.GetNodeInfo(rig.View.Node("OFFSET($I$19,0,0,1,Aqui_Per)"));

        Assert.True(info.IsLink);
        Assert.Equal("I19:J19 containing\n[1x2] , 0.252054794520548", info.TooltipText);
    }

    [Fact]
    public void ANodeWithoutARangeIsNotALinkButStillHasATooltip()
    {
        var rig = Rig.VideoModel();
        rig.Gateway.Value("(($I27/Hundred)+1)^(K$31)", new EvalResult.Number(1.00760985118373));
        rig.Start();

        var info = rig.Session.GetNodeInfo(rig.View.Node("(($I27/Hundred)+1)^(K$31)"));

        Assert.Equal(new NodeInfo(false, "1.00760985118373"), info);
    }

    [Fact]
    public void ARangeFromANonLinkNodeIsNotALink()
    {
        // Evaluate("1+1") never returns a Range, but even if a plain operator did, only reference operators are links.
        var rig = new Rig();
        const string formula = "=A1+B1";
        rig.Gateway.Formula("S", "C1", formula).Range("A1+B1", "S", "A1:B1", false, 1.0, 2.0);
        rig.Gateway.Active = FakeGateway.Loc("S", "C1");
        rig.Start();

        var info = rig.Session.GetNodeInfo(rig.View.Node("A1+B1"));

        Assert.False(info.IsLink);
        Assert.NotNull(info.TooltipText);
    }

    [Fact]
    public void UnavailableEvaluationsReadUnableToEvaluate()
    {
        var rig = Rig.VideoModel();
        rig.Start();

        var info = rig.Session.GetNodeInfo(rig.View.Node("1")); // not registered in the fake

        Assert.Equal(new NodeInfo(false, "[Unable to evaluate]"), info);
    }

    [Fact]
    public void EvaluationsDoneForColoringAreReusedByTheHover()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        int before = rig.Gateway.EvaluateCalls.Count;

        rig.Session.GetNodeInfo(rig.View.Node("K$31"));
        rig.Session.GetNodeInfo(rig.View.Node("K$31"));
        rig.Session.GetNodeInfo(rig.View.Node("Hundred"));

        Assert.Equal(before, rig.Gateway.EvaluateCalls.Count);
    }

    [Fact]
    public void ANodeIsEvaluatedOnceAndOnTheRootsSheet()
    {
        var rig = Rig.VideoModel();
        rig.Start();

        rig.Session.GetNodeInfo(rig.View.Node("1"));
        rig.Session.GetNodeInfo(rig.View.Node("1"));

        var calls = rig.Gateway.EvaluateCalls.Where(c => c.Expression == "1").ToArray();
        Assert.Single(calls);
        Assert.Equal(Rig.Loc("K33"), calls[0].Root);
    }

    [Fact]
    public void ValuesInvalidatedClearsTheCacheButKeepsTheDisplay()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Session.GetNodeInfo(rig.View.Node("K$31"));
        int formulas = rig.View.ShowFormulaCount;
        int before = rig.Gateway.EvaluateCalls.Count(c => c.Expression == "K$31");

        rig.Gateway.Range("K$31", Rig.Model, "K31", true, 0.5);
        rig.Gateway.RaiseValuesInvalidated();
        rig.Settle();
        var info = rig.Session.GetNodeInfo(rig.View.Node("K$31"));

        Assert.Equal("0.5", info.TooltipText);
        Assert.Equal(before + 1, rig.Gateway.EvaluateCalls.Count(c => c.Expression == "K$31"));
        Assert.Equal(formulas, rig.View.ShowFormulaCount);
        Assert.Equal(1, rig.View.RefreshHoverCount);
    }

    [Fact]
    public void BusyExcelGivesNoInfoIsNotCachedAndTheHoverIsRefreshedLater()
    {
        var rig = Rig.VideoModel();
        rig.Gateway.Value("1", new EvalResult.Number(1));
        rig.Start();
        var one = rig.View.Node("1");
        rig.Gateway.Busy["Evaluate"] = 1;

        Assert.Equal(Nothing, rig.Session.GetNodeInfo(one));

        Assert.Equal(0, rig.View.RefreshHoverCount);
        rig.Scheduler.Advance(199);
        Assert.Equal(0, rig.View.RefreshHoverCount);
        rig.Scheduler.Advance(1);
        Assert.Equal(1, rig.View.RefreshHoverCount);
        Assert.Equal(new NodeInfo(false, "1"), rig.Session.GetNodeInfo(one)); // not cached as "nothing"
    }

    [Fact]
    public void BusyHoverRetriesDoNotPileUp()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        var one = rig.View.Node("1");
        rig.Gateway.Busy["Evaluate"] = 1000;

        for (int i = 0; i < 5; i++) rig.Session.GetNodeInfo(one);
        rig.Scheduler.Settle();

        Assert.Equal(1, rig.View.RefreshHoverCount);
    }

    [Fact]
    public void ExceptionsInTheGatewayDoNotEscapeTheHover()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        rig.Gateway.Failures["Evaluate"] = new InvalidOperationException("boom");

        var info = rig.Session.GetNodeInfo(rig.View.Node("1"));

        Assert.Equal(Nothing, info);
    }

    [Fact]
    public void InfoIsUnavailableWhileTheNewRootLoads()
    {
        var rig = Rig.VideoModel();
        rig.Start();
        var node = rig.View.Node("K$31");

        rig.Gateway.UserSelects(Rig.Loc("K31"));
        rig.Scheduler.RunPosted(); // handled: the text is blank now

        Assert.True(rig.View.IsLoading);
        Assert.Equal(Nothing, rig.Session.GetNodeInfo(node));
    }
}
