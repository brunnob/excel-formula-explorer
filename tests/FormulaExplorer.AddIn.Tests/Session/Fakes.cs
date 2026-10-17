using System.Text.RegularExpressions;
using FormulaExplorer.AddIn.Excel;
using FormulaExplorer.AddIn.Session;
using FormulaExplorer.Core.Syntax;
using FormulaExplorer.Core.Values;

namespace FormulaExplorer.AddIn.Tests.Sessions;

/// <summary>
/// Deterministic <see cref="IUiScheduler"/> with a virtual clock. Nothing runs until the test says so:
/// <see cref="RunPosted"/> runs what is due now, <see cref="Advance"/> moves the clock, <see cref="Settle"/> runs everything.
/// </summary>
internal sealed class ManualScheduler : IUiScheduler
{
    private sealed record Item(long Seq, TimeSpan Due, Action Action);

    private readonly List<Item> _items = new();
    private long _seq;

    public TimeSpan Now { get; private set; }

    public int PendingCount => _items.Count;

    /// <summary>When true, <see cref="Post"/> runs the action immediately (the worst case for re-entrancy). Delayed work still waits.</summary>
    public bool Inline { get; set; }

    public void Post(Action action)
    {
        if (Inline) action();
        else _items.Add(new Item(_seq++, Now, action));
    }

    public void PostDelayed(TimeSpan delay, Action action) => _items.Add(new Item(_seq++, Now + delay, action));

    /// <summary>Runs everything that is due at the current time (including work posted meanwhile).</summary>
    public void RunPosted() => RunUntil(Now);

    /// <summary>Moves the clock forward, running everything that falls due on the way, in time order.</summary>
    public void Advance(int milliseconds) => RunUntil(Now + TimeSpan.FromMilliseconds(milliseconds));

    /// <summary>Runs all pending work, however far in the future, until nothing is left.</summary>
    public void Settle()
    {
        for (int guard = 0; _items.Count > 0; guard++)
        {
            if (guard > 10_000) throw new InvalidOperationException("Scheduler did not settle.");
            var next = _items.MinBy(i => (i.Due, i.Seq))!;
            _items.Remove(next);
            if (next.Due > Now) Now = next.Due;
            next.Action();
        }
    }

    private void RunUntil(TimeSpan target)
    {
        while (true)
        {
            var next = _items.Where(i => i.Due <= target).MinBy(i => (i.Due, i.Seq));
            if (next is null) break;
            _items.Remove(next);
            if (next.Due > Now) Now = next.Due;
            next.Action();
        }
        if (target > Now) Now = target;
    }
}

/// <summary>In-memory workbook model behind <see cref="IExcelGateway"/>, with call recording and busy injection.</summary>
internal sealed class FakeGateway : IExcelGateway
{
    public const string Book = "Book1.xlsx";

    private static readonly Regex CellPattern = new(@"^\$?([A-Za-z]+)\$?(\d+)$");

    public Dictionary<CellLocation, RootCell> Cells { get; } = new();
    public Dictionary<string, NodeEvaluation> Evaluations { get; } = new();

    public CellLocation? Active { get; set; }

    public List<(CellLocation Target, CellLocation? AlsoKeep)> NavigateCalls { get; } = new();
    public List<CellLocation[]> RestoreCalls { get; } = new();
    public List<(CellLocation Root, string Expression)> EvaluateCalls { get; } = new();
    public Dictionary<string, int> CallCounts { get; } = new();
    public int EndSessionCalls { get; private set; }
    public int DisposeCalls { get; private set; }

    /// <summary>Overrides the result of NavigateTo for a target (default Ok).</summary>
    public Dictionary<CellLocation, NavigationResult> NavigationResults { get; } = new();

    /// <summary>Runs inside NavigateTo, like Excel raising events synchronously inside Goto.</summary>
    public Action? DuringNavigate { get; set; }

    /// <summary>Operation name → how many calls still throw ExcelBusyException before succeeding.</summary>
    public Dictionary<string, int> Busy { get; } = new();

    /// <summary>What Excel said (the inner exception of ExcelBusyException) when <see cref="Busy"/> applies; a neutral failure by default.</summary>
    public Func<Exception> BusyCause { get; set; } = () => new InvalidOperationException("busy");

    /// <summary>Operation name → exception thrown by every call (an unexpected failure, not busy).</summary>
    public Dictionary<string, Exception> Failures { get; } = new();

    public event EventHandler<CellLocation>? UserSelectionChanged;
    public event EventHandler<CellLocation>? CellsChanged;
    public event EventHandler? ValuesInvalidated;
    public event EventHandler<IntPtr>? ActiveWindowChanged;

    public bool HasSubscribers =>
        UserSelectionChanged != null || CellsChanged != null || ValuesInvalidated != null;

    // ---- model building --------------------------------------------------------------------------------------

    private static string Absolute(string cell)
    {
        var m = CellPattern.Match(cell);
        if (!m.Success) throw new ArgumentException("Bad address " + cell);
        return $"${m.Groups[1].Value.ToUpperInvariant()}${m.Groups[2].Value}";
    }

    /// <summary>"K31" → "$K$31", "I19:J19" → "$I$19:$J$19".</summary>
    public static CellLocation Loc(string sheet, string address) =>
        new(Book, sheet, string.Join(":", address.Split(':').Select(Absolute)));

    public FakeGateway Formula(string sheet, string address, string formula, EvalResult? value = null)
    {
        var loc = Loc(sheet, address);
        Cells[loc] = new RootCell(loc, formula, true, value ?? new EvalResult.Number(1));
        return this;
    }

    public FakeGateway Constant(string sheet, string address, string text, EvalResult? value = null)
    {
        var loc = Loc(sheet, address);
        Cells[loc] = new RootCell(loc, text, false, value ?? (text.Length == 0 ? EvalResult.BlankValue : new EvalResult.Text(text)));
        return this;
    }

    /// <summary>Evaluate(expression) returns a Range on <paramref name="sheet"/>.</summary>
    public FakeGateway Range(string expression, string sheet, string address, bool firstCellHasFormula = false, params object?[] values)
    {
        var location = Loc(sheet, address);
        var corners = address.Split(':').Select(p => CellPattern.Match(p)).ToArray();
        bool single = corners.Length == 1 || Absolute(address.Split(':')[0]) == Absolute(address.Split(':')[1]);
        int rows = 1, columns = 1;
        if (corners.Length == 2)
        {
            columns = ColumnNumber(corners[1].Groups[1].Value) - ColumnNumber(corners[0].Groups[1].Value) + 1;
            rows = int.Parse(corners[1].Groups[2].Value) - int.Parse(corners[0].Groups[2].Value) + 1;
        }
        var grid = new object?[rows, columns];
        for (int i = 0; i < values.Length && i < rows * columns; i++)
            grid[i / columns, i % columns] = values[i];
        var relative = address.Replace("$", string.Empty).ToUpperInvariant();
        var first = Loc(sheet, address.Split(':')[0]);
        Evaluations[expression] = new NodeEvaluation(
            new EvalResult.RangeValue(relative, rows, columns, grid),
            new RangeTarget(location, first, single, firstCellHasFormula));
        return this;
    }

    public FakeGateway Value(string expression, EvalResult result)
    {
        Evaluations[expression] = new NodeEvaluation(result, null);
        return this;
    }

    private static int ColumnNumber(string letters) =>
        letters.ToUpperInvariant().Aggregate(0, (n, c) => n * 26 + (c - 'A' + 1));

    // ---- IExcelGateway ----------------------------------------------------------------------------------------

    private void Enter(string operation)
    {
        CallCounts[operation] = CallCounts.GetValueOrDefault(operation) + 1;
        if (Failures.TryGetValue(operation, out var failure)) throw failure;
        if (Busy.TryGetValue(operation, out int remaining) && remaining > 0)
        {
            Busy[operation] = remaining - 1;
            throw new ExcelBusyException(BusyCause());
        }
    }

    public int Calls(string operation) => CallCounts.GetValueOrDefault(operation);

    public RootCell? GetActiveCell()
    {
        Enter(nameof(GetActiveCell));
        return Active is { } a && Cells.TryGetValue(a, out var cell) ? cell : null;
    }

    public RootCell? ReadCell(CellLocation cell)
    {
        Enter(nameof(ReadCell));
        return Cells.GetValueOrDefault(cell);
    }

    public NodeEvaluation Evaluate(CellLocation root, string expression)
    {
        Enter(nameof(Evaluate));
        EvaluateCalls.Add((root, expression));
        return Evaluations.TryGetValue(expression, out var e)
            ? e
            : new NodeEvaluation(new EvalResult.Unavailable("not registered in the fake: " + expression), null);
    }

    public NavigationResult NavigateTo(CellLocation target, CellLocation? alsoKeep = null)
    {
        Enter(nameof(NavigateTo));
        NavigateCalls.Add((target, alsoKeep));
        DuringNavigate?.Invoke();
        if (NavigationResults.TryGetValue(target, out var result)) return result;
        // Selecting the target: its first cell becomes the active cell. Excel's own selection event is not reported.
        Active = Loc(target.SheetName, target.Address.Split(':')[0]);
        return NavigationResult.Ok;
    }

    public void RestoreExcept(IReadOnlyCollection<CellLocation> keepVisible)
    {
        Enter(nameof(RestoreExcept));
        RestoreCalls.Add(keepVisible.ToArray());
    }

    public void EndSession() => EndSessionCalls++;

    public IntPtr GetActiveWindowHandle() => IntPtr.Zero;

    public void Dispose() => DisposeCalls++;

    // ---- events as raised by Excel -----------------------------------------------------------------------------

    /// <summary>The user clicked a cell: the active cell changes and the gateway reports it.</summary>
    public void UserSelects(CellLocation location)
    {
        Active = location;
        UserSelectionChanged?.Invoke(this, location);
    }

    public void RaiseUserSelectionChanged(CellLocation location) => UserSelectionChanged?.Invoke(this, location);

    public void RaiseCellsChanged(CellLocation range) => CellsChanged?.Invoke(this, range);

    public void RaiseValuesInvalidated() => ValuesInvalidated?.Invoke(this, EventArgs.Empty);

    public void RaiseActiveWindowChanged() => ActiveWindowChanged?.Invoke(this, IntPtr.Zero);
}

/// <summary>Records everything the session tells the view, and keeps the visible state like the real window would.</summary>
internal sealed class FakeView : IExplorerView
{
    /// <summary>Chronological log: "Loading", "Formula:&lt;text&gt;", "Gold:&lt;text|none&gt;", "Back:True", "Message:...", "Refresh".</summary>
    public List<string> Log { get; } = new();

    public FormulaDisplay? Display { get; private set; }
    public bool IsLoading { get; private set; }
    public SyntaxNode? Gold { get; private set; }
    public bool BackEnabled { get; private set; }
    public List<string> Messages { get; } = new();
    public int RefreshHoverCount { get; private set; }
    public int ShowFormulaCount { get; private set; }
    public int ShowLoadingCount { get; private set; }

    /// <summary>The text on screen: null while loading (blank).</summary>
    public string? Text => IsLoading ? null : Display?.Formula.Text;

    public event EventHandler<SyntaxNode>? NodeClicked;
    public event EventHandler<SyntaxNode>? NodeDoubleClicked;
    public event EventHandler? BackRequested;
    public event EventHandler? CloseRequested;

    public bool HasSubscribers => NodeClicked != null || NodeDoubleClicked != null || BackRequested != null;

    public void ShowLoading()
    {
        IsLoading = true;
        ShowLoadingCount++;
        Gold = null; // the text is gone, and its gold with it
        Log.Add("Loading");
    }

    public void ShowFormula(FormulaDisplay display)
    {
        IsLoading = false;
        Display = display;
        Gold = null; // documented: ShowFormula clears hover, gold and tooltip
        ShowFormulaCount++;
        Log.Add("Formula:" + display.Formula.Text);
    }

    public void SetGold(SyntaxNode? node)
    {
        Gold = node;
        Log.Add("Gold:" + (node is null || Display is null ? "none" : node.GetText(Display.Formula.Text)));
    }

    public void SetBackEnabled(bool enabled)
    {
        BackEnabled = enabled;
        Log.Add("Back:" + enabled);
    }

    public void ShowMessage(string message)
    {
        Messages.Add(message);
        Log.Add("Message:" + message);
    }

    public void RefreshHover()
    {
        RefreshHoverCount++;
        Log.Add("Refresh");
    }

    // ---- user input ------------------------------------------------------------------------------------------

    public void Click(SyntaxNode node) => NodeClicked?.Invoke(this, node);

    /// <summary>What the real view raises for a double click: a click for the first press, then the double click.</summary>
    public void DoubleClick(SyntaxNode node)
    {
        NodeClicked?.Invoke(this, node);
        NodeDoubleClicked?.Invoke(this, node);
    }

    /// <summary>Only the second event (the session must cope with a missing first click).</summary>
    public void RaiseDoubleClickOnly(SyntaxNode node) => NodeDoubleClicked?.Invoke(this, node);

    public void PressBack() => BackRequested?.Invoke(this, EventArgs.Empty);

    public void PressClose() => CloseRequested?.Invoke(this, EventArgs.Empty);

    /// <summary>The node whose text is <paramref name="text"/> in the displayed formula (not the formula itself).</summary>
    public SyntaxNode Node(string text)
    {
        var display = Display ?? throw new InvalidOperationException("Nothing displayed.");
        return display.Formula.Nodes.First(n => n is not FormulaNode && n.GetText(display.Formula.Text) == text);
    }

    public string GoldText => Gold is null || Display is null ? "(none)" : Gold.GetText(Display.Formula.Text);
}

internal sealed class FakeOverlay : IPrecedentOverlay
{
    public List<(CellLocation Root, IReadOnlyList<PrecedentBox> Boxes)> ShowCalls { get; } = new();
    public int ClearCalls { get; private set; }
    public int DisposeCalls { get; private set; }

    /// <summary>What the overlay draws right now.</summary>
    public IReadOnlyList<PrecedentBox> Boxes { get; private set; } = Array.Empty<PrecedentBox>();

    public void SetExplorerWindow(IntPtr explorerHwnd) { }

    public void Show(CellLocation root, IReadOnlyList<PrecedentBox> boxes)
    {
        ShowCalls.Add((root, boxes));
        Boxes = boxes;
    }

    public void Clear()
    {
        ClearCalls++;
        Boxes = Array.Empty<PrecedentBox>();
    }

    public IReadOnlyList<OverlayRect> GetDrawnRects() => Array.Empty<OverlayRect>();

    public void Dispose() => DisposeCalls++;
}
