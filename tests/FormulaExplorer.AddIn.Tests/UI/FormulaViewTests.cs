using System.Windows;
using FormulaExplorer.AddIn.Session;
using FormulaExplorer.AddIn.UI;
using FormulaExplorer.Core.Syntax;

namespace FormulaExplorer.AddIn.Tests.UI;

/// <summary>FormulaView behavior with real WPF text layout (no window): hover, gold, tooltip state, click detection.</summary>
[Collection(DesktopWindows.Name)]
public class FormulaViewTests
{
    private const string Formula = "=IF(C3,0,C11*'Inventory Forcast'!C$6)";

    private static (FormulaView View, FakeProvider Provider, FormulaDisplay Display) Create(string text = Formula, Func<SyntaxNode, NodeInfo>? info = null)
    {
        var provider = info is null ? FakeProvider.Standard() : new FakeProvider(info);
        var display = Displays.Of(text);
        return (Displays.NewView(provider, display), provider, display);
    }

    private static Point Center(FormulaView view, int index) => view.GetCharCenter(index)!.Value;

    [Fact]
    public void HoveringACharacterSelectsItsNodeAndAsksTheProviderOnce() => Sta.Run(() =>
    {
        var (view, provider, display) = Create();
        int forcast = Formula.IndexOf("Forcast", StringComparison.Ordinal);
        var reference = Displays.Find(display, "'Inventory Forcast'!C$6", NodeKind.Reference);

        view.SimulateMove(Center(view, forcast));
        view.SimulateMove(Center(view, forcast + 2)); // still the same node: no new query
        Assert.Same(reference, view.HoverNode);
        Assert.Equal(1, provider.Queries.Count(n => ReferenceEquals(n, reference)));
        Assert.True(view.HoverIsLink);
    });

    [Fact]
    public void MovingToAnotherNodeChangesTheHover() => Sta.Run(() =>
    {
        var (view, _, display) = Create();
        view.SimulateMove(Center(view, 4)); // 'C' of C3
        Assert.Same(Displays.Find(display, "C3", NodeKind.Reference), view.HoverNode);
        view.SimulateMove(Center(view, 1)); // 'I' of IF
        Assert.Equal(NodeKind.FunctionCall, view.HoverNode!.Kind);
        Assert.False(view.HoverIsLink); // the fake provider links references only
    });

    [Fact]
    public void EmptyAreaAndLeavingClearTheHover() => Sta.Run(() =>
    {
        var (view, _, _) = Create();
        view.SimulateMove(Center(view, 4));
        Assert.NotNull(view.HoverNode);

        view.SimulateMove(new Point(380, 10)); // right of the text
        Assert.Null(view.HoverNode);

        view.SimulateMove(Center(view, 4));
        view.SimulateMove(null);
        Assert.Null(view.HoverNode);
        Assert.False(view.TooltipVisible);
    });

    [Fact]
    public void CommasHaveNoNode() => Sta.Run(() =>
    {
        var (view, _, _) = Create();
        view.SimulateMove(Center(view, Formula.IndexOf(',', StringComparison.Ordinal)));
        Assert.Null(view.HoverNode);
    });

    [Fact]
    public void SimulatedHoverShowsTheTooltipImmediately() => Sta.Run(() =>
    {
        var (view, _, _) = Create();
        view.SimulateMove(Center(view, 4));
        Assert.True(view.TooltipVisible);
        Assert.Equal("value of Reference", view.TooltipText);
    });

    [Fact]
    public void TooltipAppearsAfterTheDelayAndIsPlacedUnderThePointer() => Sta.Run(() =>
    {
        var (view, _, _) = Create();
        var pointer = Center(view, 4);
        view.UpdateHover(pointer);
        Assert.False(view.TooltipVisible);

        Assert.True(Wait.Until(() => view.TooltipVisible)); // the real countdown, however long a loaded machine takes
        Assert.Equal(new Point(pointer.X, pointer.Y + 17), view.TooltipPlacement);
    });

    [Fact]
    public void LeavingTheNodeBeforeTheDelayCancelsTheTooltip() => Sta.Run(() =>
    {
        var (view, _, _) = Create();
        view.UpdateHover(Center(view, 4));
        view.UpdateHover(null);
        Sta.Pump(FormulaView.TooltipDelayMs + 400);
        Assert.False(view.TooltipVisible);
    });

    [Fact]
    public void TooltipDoesNotFollowThePointerWithinTheNode() => Sta.Run(() =>
    {
        var (view, _, _) = Create();
        int forcast = Formula.IndexOf("Forcast", StringComparison.Ordinal);
        view.SimulateMove(Center(view, forcast));
        var placement = view.TooltipPlacement;
        view.SimulateMove(Center(view, forcast + 3));
        Assert.Equal(placement, view.TooltipPlacement);
        Assert.True(view.TooltipVisible);
    });

    [Fact]
    public void TooltipClosesWhenTheHoveredNodeChanges() => Sta.Run(() =>
    {
        var (view, _, _) = Create();
        view.SimulateMove(Center(view, 4));
        Assert.True(view.TooltipVisible);
        view.UpdateHover(Center(view, 1));
        Assert.False(view.TooltipVisible);
    });

    [Fact]
    public void NodesWithoutTooltipTextShowNone() => Sta.Run(() =>
    {
        var (view, _, _) = Create(info: n => new NodeInfo(true, null));
        view.SimulateMove(Center(view, 4));
        Assert.False(view.TooltipVisible);
        Assert.True(view.HoverIsLink);
    });

    [Fact]
    public void AClickHidesTheTooltipAndItStaysHiddenOnTheSameNode() => Sta.Run(() =>
    {
        var (view, _, _) = Create();
        view.SimulateMove(Center(view, 4));
        Assert.True(view.TooltipVisible);

        view.SimulatePress(Center(view, 4), 1);
        Assert.False(view.TooltipVisible);

        view.SimulateMove(Center(view, 5)); // same node, pointer moved
        Assert.False(view.TooltipVisible);

        view.SimulateMove(Center(view, 1)); // another node...
        view.SimulateMove(Center(view, 4)); // ...and back: the tooltip is allowed again
        Assert.True(view.TooltipVisible);
    });

    [Fact]
    public void RefreshHoverUpdatesTheOpenTooltipAndLinkState() => Sta.Run(() =>
    {
        string text = "first";
        bool link = false;
        var (view, provider, _) = Create(info: n => new NodeInfo(link, text));
        view.SimulateMove(Center(view, 4));
        Assert.Equal("first", view.TooltipText);
        Assert.False(view.HoverIsLink);

        text = "second";
        link = true;
        view.RefreshHover();
        Assert.Equal("second", view.TooltipText);
        Assert.True(view.HoverIsLink);

        text = null!;
        view.RefreshHover();
        Assert.False(view.TooltipVisible);
    });

    [Fact]
    public void RefreshHoverDoesNotRestartARunningTooltipCountdown() => Sta.Run(() =>
    {
        string? text = "v1";
        var (view, _, _) = Create(info: n => new NodeInfo(true, text));
        view.TooltipDelay = TimeSpan.FromHours(1);   // no wall-clock race: the countdown only ends when the test says so
        view.UpdateHover(Center(view, 4));
        Assert.True(view.TooltipTimerRunning);
        Assert.Equal(1, view.TooltipTimerStarts);

        text = "v2";
        view.RefreshHover(); // a recalculation while the countdown runs
        Assert.Equal(1, view.TooltipTimerStarts);
        Assert.True(view.TooltipTimerRunning);
        Assert.False(view.TooltipVisible);

        view.ExpireTooltipCountdown();
        Assert.True(view.TooltipVisible);
        Assert.Equal("v2", view.TooltipText);
    });

    [Fact]
    public void RefreshHoverStartsTheCountdownWhenTheTooltipTextAppearsAndStopsItWhenItGoes() => Sta.Run(() =>
    {
        string? text = null;
        var (view, _, _) = Create(info: n => new NodeInfo(true, text));
        view.UpdateHover(Center(view, 4));
        Assert.False(view.TooltipTimerRunning);

        text = "now";
        view.RefreshHover();
        Assert.True(view.TooltipTimerRunning);
        Assert.Equal(1, view.TooltipTimerStarts);

        text = null;
        view.RefreshHover();
        Assert.False(view.TooltipTimerRunning);
    });

    [Fact]
    public void ProviderExceptionsAreContained() => Sta.Run(() =>
    {
        var (view, _, _) = Create(info: _ => throw new InvalidOperationException("boom"));
        view.SimulateMove(Center(view, 4));
        Assert.NotNull(view.HoverNode);
        Assert.False(view.HoverIsLink);
        Assert.False(view.TooltipVisible);
    });

    // ----- gold -----

    [Fact]
    public void HoveringTheGoldNodeShowsOnlyGold() => Sta.Run(() =>
    {
        var (view, _, display) = Create();
        var c3 = Displays.Find(display, "C3", NodeKind.Reference);
        view.SimulateMove(Center(view, 4));
        Assert.True(view.TooltipVisible);
        Assert.True(view.HoverIsLink);

        view.SetGold(c3); // the click made the hovered node gold
        Assert.False(view.TooltipVisible);
        Assert.False(view.HoverIsLink); // no underline / hand cursor
        Assert.Same(c3, view.GoldNode);

        view.SimulateMove(Center(view, 1));
        view.SimulateMove(Center(view, 4)); // coming back onto the gold node: still no tooltip
        Assert.False(view.TooltipVisible);
    });

    [Fact]
    public void OtherNodesBehaveNormallyWhileSomethingIsGold() => Sta.Run(() =>
    {
        var (view, _, display) = Create();
        view.SetGold(Displays.Find(display, "C3", NodeKind.Reference));
        view.SimulateMove(Center(view, Formula.IndexOf("C11", StringComparison.Ordinal)));
        Assert.True(view.HoverIsLink);
        Assert.True(view.TooltipVisible);
    });

    [Fact]
    public void ANodeContainingTheGoldNodeIsHoveredNormally() => Sta.Run(() =>
    {
        var (view, _, display) = Create();
        view.SetGold(Displays.Find(display, "C3", NodeKind.Reference));
        view.SimulateMove(Center(view, 1)); // IF( ... ) contains C3
        Assert.Equal(NodeKind.FunctionCall, view.HoverNode!.Kind);
        Assert.True(view.TooltipVisible);
    });

    [Fact]
    public void RemovingTheGoldWhileHoveringRestoresTheNormalLook() => Sta.Run(() =>
    {
        var (view, _, display) = Create();
        view.SetGold(Displays.Find(display, "C3", NodeKind.Reference));
        view.SimulateMove(Center(view, 4));
        Assert.False(view.HoverIsLink);

        view.SetGold(null);
        Assert.True(view.HoverIsLink);
        Assert.Null(view.GoldNode);
    });

    // ----- clicks -----

    [Fact]
    public void ClickOnALinkRaisesNodeClicked() => Sta.Run(() =>
    {
        var (view, _, display) = Create();
        var clicked = new List<SyntaxNode>();
        view.NodeClicked += (_, n) => clicked.Add(n);

        view.SimulatePress(Center(view, 4), 1);
        Assert.Equal(new[] { Displays.Find(display, "C3", NodeKind.Reference) }, clicked);
    });

    [Fact]
    public void ClickOnANonLinkDoesNothing() => Sta.Run(() =>
    {
        var (view, _, _) = Create();
        int raised = 0;
        view.NodeClicked += (_, _) => raised++;
        view.NodeDoubleClicked += (_, _) => raised++;

        view.SimulatePress(Center(view, 1), 1); // the IF function: not a link for the fake provider
        view.SimulatePress(Center(view, 1), 2);
        view.SimulatePress(new Point(380, 10), 1); // empty area
        Assert.Equal(0, raised);
    });

    [Fact]
    public void TwoQuickPressesOnTheSameNodeRaiseClickThenDoubleClick() => Sta.Run(() =>
    {
        var (view, _, display) = Create();
        var events = new List<string>();
        view.NodeClicked += (_, n) => events.Add("click:" + n.GetText(display.Formula.Text));
        view.NodeDoubleClicked += (_, n) => events.Add("double:" + n.GetText(display.Formula.Text));

        view.SimulatePress(Center(view, 4), 1);
        view.SimulatePress(Center(view, 4), 1); // WPF would say ClickCount 1 again after Excel re-activated
        Assert.Equal(new[] { "click:C3", "double:C3" }, events);
    });

    [Fact]
    public void WpfClickCountTwoWithoutAnEarlierPressRaisesBoth() => Sta.Run(() =>
    {
        var (view, _, _) = Create();
        var events = new List<string>();
        view.NodeClicked += (_, _) => events.Add("click");
        view.NodeDoubleClicked += (_, _) => events.Add("double");

        view.SimulatePress(Center(view, 4), 2);
        Assert.Equal(new[] { "click", "double" }, events);
    });

    [Fact]
    public void PressesOnDifferentNodesAreTwoClicks() => Sta.Run(() =>
    {
        var (view, _, _) = Create();
        var events = new List<string>();
        view.NodeClicked += (_, _) => events.Add("click");
        view.NodeDoubleClicked += (_, _) => events.Add("double");

        view.SimulatePress(Center(view, 4), 1);
        view.SimulatePress(Center(view, Formula.IndexOf("C11", StringComparison.Ordinal)), 1);
        Assert.Equal(new[] { "click", "click" }, events);
    });

    // ----- text, loading, non-formulas -----

    [Fact]
    public void ShowFormulaClearsHoverGoldAndTooltip() => Sta.Run(() =>
    {
        var (view, _, display) = Create();
        view.SetGold(Displays.Find(display, "C3", NodeKind.Reference));
        view.SimulateMove(Center(view, Formula.IndexOf("C11", StringComparison.Ordinal)));
        Assert.True(view.TooltipVisible);

        view.ShowFormula(Displays.Of("=A1+B1"));
        Assert.Null(view.HoverNode);
        Assert.Null(view.GoldNode);
        Assert.False(view.TooltipVisible);
        Assert.Equal("=A1+B1", view.DisplayText);
        Assert.False(view.IsLoading);
    });

    [Fact]
    public void ShowLoadingBlanksTheText() => Sta.Run(() =>
    {
        var (view, _, _) = Create();
        view.SimulateMove(Center(view, 4));
        view.ShowLoading();
        Assert.True(view.IsLoading);
        Assert.Equal(string.Empty, view.DisplayText);
        Assert.False(view.TooltipVisible);
        Assert.Null(view.HoverNode);
        Assert.False(view.Interactive);
    });

    [Fact]
    public void NonFormulaTextIsNotInteractive() => Sta.Run(() =>
    {
        var provider = FakeProvider.Standard();
        var view = Displays.NewView(provider, Displays.Of("Quarterly forecast"));
        Assert.False(view.Interactive);
        view.SimulateMove(Center(view, 3));
        view.SimulatePress(Center(view, 3), 1);
        Assert.Null(view.HoverNode);
        Assert.False(view.TooltipVisible);
        Assert.Empty(provider.Queries);
    });

    [Fact]
    public void LongFormulasWrapAndKeepHitTestingConsistent() => Sta.Run(() =>
    {
        const string longFormula = "=SUM(I$19:$I$19)+J$19*$I$28/Hundred-SUM(OFFSET($I$19,0,0,1,Aqui_Per))+Sheet8!E7+(E30/C4)^(D4/D5)+D14-E21*E16+Summary!C5";
        var (view, _, display) = Create(longFormula);
        Assert.True(view.DesiredSize.Height > TextLayoutHeightOfOneLine);

        for (int i = 0; i < longFormula.Length; i++)
        {
            var rect = view.GetCharRect(i)!.Value;
            view.UpdateHover(new Point(rect.Left + rect.Width / 2, rect.Top + rect.Height / 2));
            Assert.Same(display.HitMap.NodeAt(i), view.HoverNode);
        }
    });

    private const double TextLayoutHeightOfOneLine = 3 + 20;

    [Fact]
    public void LineHeightIsTwentyDipsAndLinesAbut() => Sta.Run(() =>
    {
        const string longFormula = "=SUM(I$19:$I$19)+J$19*$I$28/Hundred-SUM(OFFSET($I$19,0,0,1,Aqui_Per))+Sheet8!E7+(E30/C4)^(D4/D5)";
        var (view, _, _) = Create(longFormula);
        var first = view.GetCharRect(0)!.Value;
        Assert.Equal(20, first.Height);
        Assert.Equal(FormulaView.TextInset.Left, first.Left);
        var lastChar = view.GetCharRect(longFormula.Length - 1)!.Value;
        Assert.True(lastChar.Top >= 20);
        Assert.Equal(0, (lastChar.Top - FormulaView.TextInset.Top) % 20);
    });
}
