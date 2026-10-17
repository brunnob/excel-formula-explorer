using FormulaExplorer.AddIn.Excel;
using FormulaExplorer.AddIn.Session;
using FormulaExplorer.Core.Syntax;
using FormulaExplorer.Core.Values;

namespace FormulaExplorer.AddIn.Tests.Sessions;

/// <summary>A session wired to fakes, plus the workbook of the original demo video.</summary>
internal sealed class Rig
{
    public const string Model = "Model";
    public const string Assumptions = "Assumptions";

    public const string K33Formula = "=1/(((($I27/Hundred)+1)^(K$31)))";
    public const string K31Formula = "=J31+K19";
    public const string J31Formula = "=SUM(I$19:$I$19)+J$19*$I$28/Hundred-SUM(OFFSET($I$19,0,0,1,Aqui_Per))";

    public FakeGateway Gateway { get; } = new();
    public FakeView View { get; } = new();
    public FakeOverlay Overlay { get; } = new();
    public ManualScheduler Scheduler { get; } = new();
    public ExplorerSession Session { get; }

    public Rig()
    {
        Session = new ExplorerSession(Gateway, View, Overlay, Scheduler);
    }

    public static CellLocation Loc(string address, string sheet = Model) => FakeGateway.Loc(sheet, address);

    /// <summary>
    /// The model of the demo: K33 = 1/((($I27/Hundred)+1)^(K$31)), K31 = J31+K19, J31 = SUM(...)+OFFSET(...), with the
    /// name Hundred (=100) on another sheet. K19, I27, I28, J19 are constants, I19 is blank.
    /// </summary>
    public static Rig VideoModel()
    {
        var rig = new Rig();
        var g = rig.Gateway;
        g.Formula(Model, "K33", K33Formula, new EvalResult.Number(1.00760985118373))
         .Formula(Model, "K31", K31Formula, new EvalResult.Number(0.126))
         .Formula(Model, "J31", J31Formula, new EvalResult.Number(0.461538461538462))
         .Constant(Model, "K19", "0.05", new EvalResult.Number(0.05))
         .Constant(Model, "I27", "7", new EvalResult.Number(7))
         .Constant(Model, "I28", "3", new EvalResult.Number(3))
         .Constant(Model, "I19", string.Empty)
         .Constant(Model, "J19", "0.252054794520548", new EvalResult.Number(0.252054794520548))
         .Constant(Assumptions, "B1", "100", new EvalResult.Number(100))
         .Constant(Assumptions, "B2", "2", new EvalResult.Number(2));

        // Fragments as the session sends them to Evaluate (the node text).
        g.Range("K$31", Model, "K31", firstCellHasFormula: true, 0.126)
         .Range("$I27", Model, "I27", false, 7.0)
         .Range("Hundred", Assumptions, "B1", false, 100.0)
         .Range("K19", Model, "K19", false, 0.05)
         .Range("J31", Model, "J31", true, 0.461538461538462)
         .Range("I$19:$I$19", Model, "I19", false, new object?[] { null })
         .Range("J$19", Model, "J19", false, 0.252054794520548)
         .Range("$I$28", Model, "I28", false, 3.0)
         .Range("$I$19", Model, "I19", false, new object?[] { null })
         .Range("Aqui_Per", Assumptions, "B2", false, 2.0)
         .Range("OFFSET($I$19,0,0,1,Aqui_Per)", Model, "I19:J19", false, null, 0.252054794520548);

        g.Active = Loc("K33");
        return rig;
    }

    public ExplorerSession Start()
    {
        Session.Start();
        Settle();
        return Session;
    }

    /// <summary>Runs everything the session posted, including the 150 ms loads and busy retries.</summary>
    public void Settle() => Scheduler.Settle();

    public SessionSnapshot Snapshot => Session.GetSnapshot();

    public string Text => View.Text ?? "(loading)";

    public void Click(string nodeText)
    {
        View.Click(View.Node(nodeText));
        Settle();
    }

    public void DoubleClick(string nodeText)
    {
        View.DoubleClick(View.Node(nodeText));
        Settle();
    }

    public void Back()
    {
        View.PressBack();
        Settle();
    }

    public static IReadOnlyList<string> Addresses(IEnumerable<CellLocation> locations) =>
        locations.Select(l => l.Address).ToArray();

    public static string Box(PrecedentBox box) => $"{box.Location.SheetName}!{box.Location.Address}#{box.Rgb:X6}";
}
