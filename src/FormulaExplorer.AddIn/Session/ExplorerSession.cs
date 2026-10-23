using FormulaExplorer.AddIn.Diagnostics;
using FormulaExplorer.AddIn.Excel;
using FormulaExplorer.Core.Analysis;
using FormulaExplorer.Core.Syntax;
using FormulaExplorer.Core.Values;

namespace FormulaExplorer.AddIn.Session;

/// <summary>
/// The Formula Explorer state machine: the <em>root</em> (the cell whose formula is shown), the <em>gold</em> link
/// (the node the user last clicked, with the range it resolved to), the Back <em>stack</em> of earlier roots, and the
/// per-root evaluation cache. Pure logic against <see cref="IExcelGateway"/>, <see cref="IExplorerView"/>,
/// <see cref="IPrecedentOverlay"/> and <see cref="IUiScheduler"/>; see docs/PESQUISA-FORMULA-EXPLORER.md section 4.
/// </summary>
/// <remarks>
/// Everything runs on Excel's main thread. Two rules keep the machine safe against Excel's synchronous events:
/// <list type="bullet">
/// <item>Gateway events (selection, cells changed, values invalidated) are never handled inline. A handler only records
/// the event (coalescing it with earlier ones) and posts a drain; the drain is also deferred while a view command or
/// another drain is running (<c>_inCommand</c>) and is rescheduled when that one ends.</item>
/// <item>Anything that talks to Excel may throw <see cref="ExcelBusyException"/> (cell edit mode, modal dialog). Such an
/// operation is retried every 200 ms, at most 25 times. State is therefore mutated only after the last call that can
/// throw, or in an order that makes a second attempt harmless.</item>
/// </list>
/// </remarks>
public sealed class ExplorerSession : INodeInfoProvider, IDisposable
{
    /// <summary>The short, visible blank between "text cleared" and "new text shown" (the original did the same).</summary>
    private static readonly TimeSpan LoadDelay = TimeSpan.FromMilliseconds(150);

    private static readonly TimeSpan BusyRetryDelay = TimeSpan.FromMilliseconds(200);
    private const int MaxBusyRetries = 25;

    /// <summary>
    /// Retries (about 400 ms) of a USER command (click, double click, Back) that Excel rejects with a bare VBA_E_IGNORE, i.e.
    /// "not ready" without the edit-mode check knowing why. That state may be an edit-like mode that lasts as long as the user
    /// keeps typing (a shape, chart title or note being edited), and a click must never fire seconds later, when it ends.
    /// Background work (loads, the drain) and the other busy answers keep the full <see cref="MaxBusyRetries"/>.
    /// </summary>
    internal const int UserCommandNotReadyRetries = 2;

    private static readonly NodeInfo NoInfo = new(false, null);

    /// <summary>Shown when there is no root cell: empty text, nothing interactive.</summary>
    private static readonly FormulaDisplay EmptyDisplay = MakePlainDisplay(string.Empty);

    private enum LifeCycle { New, Open, Closed }

    /// <summary>The clicked node and where it points.</summary>
    private sealed record Gold(SyntaxNode Node, RangeTarget Target);

    private readonly IExcelGateway _gateway;
    private readonly IExplorerView _view;
    private readonly IPrecedentOverlay _overlay;
    private readonly IUiScheduler _scheduler;

    private LifeCycle _life = LifeCycle.New;

    // --- state machine ---
    private RootCell? _root;
    private Gold? _gold;
    private readonly List<CellLocation> _stack = new(); // oldest first, top = last
    private FormulaDisplay? _display;                    // null while a new root loads (blank text) or before the first load
    private HashSet<SyntaxNode> _displayNodes = new(ReferenceEqualityComparer.Instance);
    private IReadOnlyList<PrecedentBox> _boxes = Array.Empty<PrecedentBox>();
    private readonly Dictionary<SyntaxNode, NodeEvaluation> _cache = new(ReferenceEqualityComparer.Instance);
    private EvalResult? _rootValue; // fresh Value2 of the root cell; dropped on invalidation
    private int _loadVersion;
    private bool _fallbackShown; // the plain-text fallback is on screen for a formula root; the next invalidation retries

    // --- re-entrancy ---
    private bool _inCommand;
    private bool _drainScheduled;
    private bool _drainBackingOff; // a delayed busy retry of the drain is pending; it covers every event recorded meanwhile
    private bool _pendingInvalidated;
    private CellLocation? _pendingSelection;
    private readonly HashSet<(string Workbook, string Sheet)> _pendingCells = new();

    // --- hover refresh after a busy evaluation ---
    private bool _hoverRetryPending;
    private int _hoverRetries;

    public ExplorerSession(IExcelGateway gateway, IExplorerView view, IPrecedentOverlay overlay, IUiScheduler scheduler)
    {
        _gateway = gateway ?? throw new ArgumentNullException(nameof(gateway));
        _view = view ?? throw new ArgumentNullException(nameof(view));
        _overlay = overlay ?? throw new ArgumentNullException(nameof(overlay));
        _scheduler = scheduler ?? throw new ArgumentNullException(nameof(scheduler));
    }

    /// <summary>True between <see cref="Start"/> and <see cref="Close"/>.</summary>
    public bool IsOpen => _life == LifeCycle.Open;

    private bool BackEnabled => _gold != null || _stack.Count > 0;

    // ------------------------------------------------------------------------------------------------------------
    // Lifecycle
    // ------------------------------------------------------------------------------------------------------------

    /// <summary>Starts the session: root := active cell, empty stack, no gold; subscribes to view and gateway events.</summary>
    public void Start()
    {
        if (_life != LifeCycle.New) return;
        _life = LifeCycle.Open;

        // Subscribe first so that a selection change between the subscription and the read is not lost.
        _view.NodeClicked += OnNodeClicked;
        _view.NodeDoubleClicked += OnNodeDoubleClicked;
        _view.BackRequested += OnBackRequested;
        _gateway.UserSelectionChanged += OnUserSelectionChanged;
        _gateway.CellsChanged += OnCellsChanged;
        _gateway.ValuesInvalidated += OnValuesInvalidated;

        Execute("start", () =>
        {
            var cell = _gateway.GetActiveCell();
            if (cell is null)
            {
                ShowNoRoot();
                return;
            }
            _root = cell;
            Load();
        });
    }

    /// <summary>
    /// Ends the session: unsubscribes from all events, clears the overlay and calls gateway.EndSession().
    /// Idempotent. Does not close the window (the controller does).
    /// </summary>
    public void Close()
    {
        if (_life != LifeCycle.Open)
        {
            _life = LifeCycle.Closed; // Close before Start: there is nothing to undo
            return;
        }

        _life = LifeCycle.Closed; // first: every posted callback checks IsOpen and becomes a no-op
        _loadVersion++;

        _view.NodeClicked -= OnNodeClicked;
        _view.NodeDoubleClicked -= OnNodeDoubleClicked;
        _view.BackRequested -= OnBackRequested;
        _gateway.UserSelectionChanged -= OnUserSelectionChanged;
        _gateway.CellsChanged -= OnCellsChanged;
        _gateway.ValuesInvalidated -= OnValuesInvalidated;

        _gold = null;
        _display = null;
        _fallbackShown = false;
        _drainBackingOff = false;
        _displayNodes = new HashSet<SyntaxNode>(ReferenceEqualityComparer.Instance);
        _cache.Clear();
        _pendingSelection = null;
        _pendingCells.Clear();
        _pendingInvalidated = false;

        try
        {
            _overlay.Clear();
            _boxes = Array.Empty<PrecedentBox>();
        }
        catch (Exception ex)
        {
            Log.Error("ExplorerSession.Close: overlay.Clear", ex);
        }

        try
        {
            _gateway.EndSession();
        }
        catch (Exception ex)
        {
            Log.Error("ExplorerSession.Close: gateway.EndSession", ex);
        }
    }

    public void Dispose() => Close();

    // ------------------------------------------------------------------------------------------------------------
    // INodeInfoProvider
    // ------------------------------------------------------------------------------------------------------------

    /// <summary>
    /// What the view shows for a hovered node. The evaluation is lazy and cached per root; it is the node's own text
    /// (a 3D reference is evaluated as its first sheet so that it can navigate). Excel busy → no info, not cached.
    /// </summary>
    public NodeInfo GetNodeInfo(SyntaxNode node)
    {
        if (!IsOpen) return NoInfo;
        try
        {
            var display = _display;
            if (display is null || _root is null || !_displayNodes.Contains(node)) return NoInfo;

            NodeInfo info;
            if (node is NameNode { IsLocal: true })
            {
                // A LET/LAMBDA variable only has a value inside its formula: evaluating the bare name would report the
                // workbook name of the same spelling (or #NAME?). Highlight it, but no link and no tooltip.
                info = NoInfo;
            }
            else if (node is FormulaNode)
            {
                info = new NodeInfo(false, ValueFormatter.Format(GetRootValue()));
            }
            else
            {
                var evaluation = Evaluate(display.Formula, node);
                info = new NodeInfo(IsLink(node, evaluation), ValueFormatter.Format(evaluation.Result));
            }
            _hoverRetries = 0;
            return info;
        }
        catch (ExcelBusyException)
        {
            ScheduleHoverRetry();
            return NoInfo;
        }
        catch (Exception ex)
        {
            Log.Error("ExplorerSession.GetNodeInfo", ex);
            return NoInfo;
        }
    }

    /// <summary>
    /// Session state for tests and diagnostics. <see cref="SessionSnapshot.Stack"/> is oldest first (the last entry is
    /// where Back goes next); <see cref="SessionSnapshot.Boxes"/> are the boxes the overlay shows right now.
    /// </summary>
    public SessionSnapshot GetSnapshot() => new(
        IsOpen,
        _root?.Location,
        _root?.DisplayText ?? string.Empty,
        _stack.ToArray(),
        _gold?.Node.Start,
        _gold?.Node.Length,
        _gold?.Target.Location,
        BackEnabled,
        _boxes);

    // ------------------------------------------------------------------------------------------------------------
    // View events (commands)
    // ------------------------------------------------------------------------------------------------------------

    private void OnNodeClicked(object? sender, SyntaxNode node)
    {
        var display = _display;
        if (display is null) return;
        var attempt = new ClickAttempt(); // shared by the busy retries of this one click
        Execute("click", () => ClickAndDrill(display, node, attempt), userCommand: true);
    }

    /// <summary>
    /// A double click needs nothing more: its first click already navigated and, for a cell holding a formula, drilled
    /// down (the view swallows the second press when the text was replaced in between).
    /// </summary>
    private void OnNodeDoubleClicked(object? sender, SyntaxNode node)
    {
    }

    /// <summary>State of one user click across its busy retries.</summary>
    private sealed class ClickAttempt
    {
        public bool Navigated;
    }

    private void OnBackRequested(object? sender, EventArgs e) => Execute("back", Back, userCommand: true);

    /// <summary>
    /// Single click on a link: select its target in Excel (keeping the root's sheet visible) and make the node gold.
    /// Returns true when the node is gold afterwards.
    /// </summary>
    private bool Click(FormulaDisplay display, SyntaxNode node)
    {
        // The display may have been replaced while a busy retry was waiting.
        if (!ReferenceEquals(display, _display) || _root is null || !_displayNodes.Contains(node)) return false;

        var evaluation = Evaluate(display.Formula, node);
        if (!IsLink(node, evaluation)) return false;
        var target = evaluation.Target!;

        var result = _gateway.NavigateTo(target.Location, alsoKeep: _root.Location);
        switch (result.Status)
        {
            case NavigationStatus.Ok:
                SetGold(new Gold(node, target));
                return true;

            case NavigationStatus.Failed:
                // The gateway already restored what the previous gold needed; do not leave a stale gold behind.
                SetGold(null);
                PostMessage(result.Message);
                return false;

            default: // Blocked: the gateway changed nothing, the previous gold (if any) is still valid
                PostMessage(result.Message);
                return false;
        }
    }

    /// <summary>
    /// Single click on a link: navigate to the target (the node becomes gold). When the target is a single cell holding
    /// a formula, drill down into it at once: the window shows that cell's formula and the old root goes on the Back
    /// stack. Otherwise (range, constant, empty cell) the node stays gold. A busy retry does not navigate twice.
    /// </summary>
    private void ClickAndDrill(FormulaDisplay display, SyntaxNode node, ClickAttempt attempt)
    {
        if (!ReferenceEquals(display, _display) || !_displayNodes.Contains(node)) return;
        // A fresh click always navigates (also on the gold node: the user may have moved away in Excel); only a busy
        // retry of the same click skips the navigation it already did.
        if (!(attempt.Navigated && _gold?.Node == node))
        {
            if (!Click(display, node)) return;
            attempt.Navigated = true;
        }

        var target = _gold!.Target;
        if (!target.IsSingleCell || !target.FirstCellHasFormula) return;

        var cell = _gateway.ReadCell(target.FirstCell);
        if (cell is null) return; // the cell vanished: stay on the current root

        PushRoot();
        _root = cell;
        SetGold(null);
        Load();
    }

    /// <summary>Back: first undo the gold (re-select the root), then retrace the stack.</summary>
    private void Back()
    {
        if (_gold != null && _root != null)
        {
            var result = _gateway.NavigateTo(_root.Location);
            SetGold(null);
            if (result.Status != NavigationStatus.Ok) PostMessage(result.Message);
            return;
        }

        // Peek, don't pop, until the entry is known to be alive: a busy retry must not lose entries.
        while (_stack.Count > 0)
        {
            var entry = _stack[^1];
            var cell = _gateway.ReadCell(entry);
            if (cell is null)
            {
                _stack.RemoveAt(_stack.Count - 1); // workbook closed or sheet deleted
                continue;
            }

            var result = _gateway.NavigateTo(entry);
            if (result.Status == NavigationStatus.Blocked)
            {
                PostMessage(result.Message);
                break;
            }

            _stack.RemoveAt(_stack.Count - 1);
            if (result.Status == NavigationStatus.Failed) continue;

            _root = cell;
            Load();
            return;
        }
        UpdateBack();
    }

    // ------------------------------------------------------------------------------------------------------------
    // Gateway events (recorded, coalesced, handled later)
    // ------------------------------------------------------------------------------------------------------------

    private void OnUserSelectionChanged(object? sender, CellLocation location)
    {
        if (!IsOpen) return;
        _pendingSelection = location; // only the latest selection matters
        ScheduleDrain();
    }

    private void OnCellsChanged(object? sender, CellLocation range)
    {
        if (!IsOpen) return;
        _pendingCells.Add((range.WorkbookName, range.SheetName)); // merged: one re-read of the root per drain
        ScheduleDrain();
    }

    private void OnValuesInvalidated(object? sender, EventArgs e)
    {
        if (!IsOpen) return;
        _pendingInvalidated = true;
        ScheduleDrain();
    }

    private bool HasPendingEvents => _pendingInvalidated || _pendingSelection != null || _pendingCells.Count > 0;

    private void ScheduleDrain()
    {
        // While a busy retry of the drain is waiting, new events are only recorded: that retry handles them, and an
        // immediate drain here would hit the same busy Excel again (a hot loop).
        if (_drainScheduled || _drainBackingOff || !IsOpen) return;
        _drainScheduled = true;
        try
        {
            _scheduler.Post(() =>
            {
                _drainScheduled = false;
                // A running command reschedules the drain when it ends; running now would interleave with it.
                if (!IsOpen || _inCommand) return;
                Execute("gateway events", DrainPending, isDrain: true);
            });
        }
        catch (Exception ex)
        {
            _drainScheduled = false;
            Log.Error("ExplorerSession.ScheduleDrain", ex);
        }
    }

    /// <summary>Handles everything recorded so far. Items hit by a busy Excel are put back for the retry.</summary>
    private void DrainPending()
    {
        if (_pendingInvalidated)
        {
            _pendingInvalidated = false;
            HandleValuesInvalidated();
        }

        if (_pendingSelection is { } selection)
        {
            _pendingSelection = null;
            try
            {
                HandleUserSelection(selection);
            }
            catch (ExcelBusyException)
            {
                _pendingSelection ??= selection; // unless a newer selection arrived meanwhile
                throw;
            }
        }

        if (_pendingCells.Count > 0)
        {
            var changed = _pendingCells.ToArray();
            _pendingCells.Clear();
            try
            {
                HandleCellsChanged(changed);
            }
            catch (ExcelBusyException)
            {
                _pendingCells.UnionWith(changed);
                throw;
            }
        }
    }

    private void HandleUserSelection(CellLocation location)
    {
        if (_root != null && location == _root.Location)
        {
            // Back on the root cell: the link is no longer where the user is.
            SetGold(null);
            _gateway.RestoreExcept(new[] { location });
            return;
        }

        var cell = _gateway.ReadCell(location);
        if (cell is null) return;

        PushRoot();
        _root = cell;
        SetGold(null);
        Load();
        // Last, because it may find Excel busy: the retry then takes the "selected the root" branch above.
        _gateway.RestoreExcept(new[] { location });
    }

    private void HandleCellsChanged(IReadOnlyList<(string Workbook, string Sheet)> changed)
    {
        var root = _root;
        if (root is null) return;
        if (!changed.Contains((root.Location.WorkbookName, root.Location.SheetName))) return;

        var fresh = _gateway.ReadCell(root.Location);
        if (fresh is null)
        {
            // The sheet or workbook of the root is gone. Keep the stack so that Back still works.
            _root = null;
            SetGold(null);
            ShowNoRoot();
            return;
        }

        _root = fresh;
        _rootValue = null;
        if (fresh.DisplayText != root.DisplayText || fresh.HasFormula != root.HasFormula || _fallbackShown)
        {
            SetGold(null);
            Load();
        }
    }

    private void HandleValuesInvalidated()
    {
        _cache.Clear();
        _rootValue = null;
        if (_fallbackShown && _root is { HasFormula: true })
        {
            Load(); // Excel may be free again: try to make the formula interactive (bounded by the usual 25 retries)
            return;
        }
        _view.RefreshHover();
    }

    // ------------------------------------------------------------------------------------------------------------
    // Loading a root
    // ------------------------------------------------------------------------------------------------------------

    /// <summary>
    /// Starts showing <see cref="_root"/>: Back is updated first (so it is enabled before the new text shows), the text
    /// area goes blank with a wait cursor, and 150 ms later the parsed, colored text replaces it. The overlay is not
    /// cleared: the old precedent borders stay until the new text appears.
    /// </summary>
    private void Load()
    {
        int version = ++_loadVersion;
        _fallbackShown = false;
        _display = null;
        _displayNodes = new HashSet<SyntaxNode>(ReferenceEqualityComparer.Instance);
        _cache.Clear();
        _rootValue = null;

        UpdateBack();
        _view.ShowLoading();
        _scheduler.PostDelayed(LoadDelay, () => Execute("load", () => CompleteLoad(version), loadVersion: version));
    }

    private void CompleteLoad(int version)
    {
        if (version != _loadVersion) return; // superseded by a newer load
        var root = _root;
        if (root is null)
        {
            ShowNoRoot();
            return;
        }

        var parsed = FormulaParser.Parse(root.DisplayText);
        FormulaDisplay display;
        var boxes = new List<PrecedentBox>();

        if (root.HasFormula && parsed.IsFormula)
        {
            // Names are colored only when they resolve to a range; the evaluation is cached for the hover that follows.
            var colors = RefColorizer.Assign(parsed, root.Location.SheetName, node => node switch
            {
                ReferenceNode reference => reference.Form != ReferenceForm.RefError,
                StructuredReferenceNode => true,
                NameNode => Evaluate(parsed, node).Target is not null,
                _ => false,
            });
            display = new FormulaDisplay(parsed, HitMap.Build(parsed), colors, Interactive: true);

            // Precedent boxes: only ranges on the root's sheet, one per distinct location (first color wins).
            var seen = new HashSet<CellLocation>();
            foreach (var colored in colors.Nodes)
            {
                var target = Evaluate(parsed, colored.Node).Target;
                if (target is not null && OnSameSheet(target.Location, root.Location) && seen.Add(target.Location))
                    boxes.Add(new PrecedentBox(target.Location, colored.Rgb));
            }
        }
        else
        {
            // A constant (or text that merely looks like a formula): shown in black, no hover, no tooltip, no clicks.
            display = MakePlainDisplay(root.DisplayText);
        }

        // Every call that can find Excel busy is behind us: from here on, publish the result.
        _display = display;
        _displayNodes = display.Interactive
            ? new HashSet<SyntaxNode>(display.Formula.Nodes, ReferenceEqualityComparer.Instance)
            : new HashSet<SyntaxNode>(ReferenceEqualityComparer.Instance);
        _view.ShowFormula(display);
        _overlay.Show(root.Location, boxes);
        _boxes = boxes;
        _fallbackShown = false;
    }

    /// <summary>
    /// Last resort when a load cannot complete (Excel stayed busy, unexpected error): show the plain text so the
    /// window is not left blank.
    /// </summary>
    private void ShowFallback(int version)
    {
        if (version != _loadVersion || _root is null) return;
        _fallbackShown = _root.HasFormula; // recoverable: HandleValuesInvalidated / HandleCellsChanged reload it
        _cache.Clear();
        _rootValue = null;
        _displayNodes = new HashSet<SyntaxNode>(ReferenceEqualityComparer.Instance);
        try
        {
            var display = MakePlainDisplay(_root.DisplayText);
            _display = display;
            _view.ShowFormula(display);
        }
        catch (Exception ex)
        {
            Log.Error("ExplorerSession.ShowFallback: view", ex);
        }

        // The old root's borders were kept only until the new text appears; it has appeared, so they must go.
        try
        {
            _overlay.Clear();
            _boxes = Array.Empty<PrecedentBox>();
        }
        catch (Exception ex)
        {
            Log.Error("ExplorerSession.ShowFallback: overlay", ex);
        }
    }

    /// <summary>No root cell (no active cell at start, or the root's sheet/workbook was deleted): empty text, no borders.</summary>
    private void ShowNoRoot()
    {
        _loadVersion++; // cancels a pending load of the deleted root
        _fallbackShown = false;
        _display = EmptyDisplay;
        _displayNodes = new HashSet<SyntaxNode>(ReferenceEqualityComparer.Instance);
        _cache.Clear();
        _rootValue = null;
        UpdateBack();
        _view.ShowFormula(EmptyDisplay);
        _overlay.Clear();
        _boxes = Array.Empty<PrecedentBox>();
    }

    private static FormulaDisplay MakePlainDisplay(string text)
    {
        var parsed = FormulaParser.Parse(text);
        // An empty HitMap (built from an empty text): no node is ever hoverable.
        return new FormulaDisplay(parsed, HitMap.Build(FormulaParser.Parse(string.Empty)), ColorMap.Empty, Interactive: false);
    }

    // ------------------------------------------------------------------------------------------------------------
    // Evaluation
    // ------------------------------------------------------------------------------------------------------------

    /// <summary>Cached <see cref="IExcelGateway.Evaluate"/> of a node on the root's sheet. Busy is not cached and propagates.</summary>
    private NodeEvaluation Evaluate(ParsedFormula formula, SyntaxNode node)
    {
        if (_cache.TryGetValue(node, out var cached)) return cached;
        var evaluation = _gateway.Evaluate(_root!.Location, ExpressionOf(formula.Text, node));
        _cache[node] = evaluation;
        return evaluation;
    }

    /// <summary>
    /// The text handed to Worksheet.Evaluate: the node's own text, except for a 3D reference, which is evaluated as its
    /// first sheet (<c>'[Workbook]Sheet'!Address</c> or <c>'Sheet'!Address</c>, quotes doubled) because Excel cannot
    /// navigate to a multi-sheet range.
    /// </summary>
    private static string ExpressionOf(string formulaText, SyntaxNode node)
    {
        if (node is ReferenceNode { Is3D: true, Sheet: { } firstSheet } reference)
        {
            string qualifier = reference.Workbook is null ? firstSheet : $"[{reference.Workbook}]{firstSheet}";
            return $"'{qualifier.Replace("'", "''")}'!{reference.Address}";
        }
        return node.GetText(formulaText);
    }

    private static bool IsLink(SyntaxNode node, NodeEvaluation evaluation) =>
        evaluation.Target is not null && HitMap.IsLinkCandidate(node);

    /// <summary>The root cell's current value (re-read after an invalidation) for the tooltip of the '=' sign.</summary>
    private EvalResult GetRootValue()
    {
        if (_rootValue is not null) return _rootValue;
        var root = _root!;
        var value = _gateway.ReadCell(root.Location)?.Value ?? root.Value;
        _rootValue = value;
        return value;
    }

    private static bool OnSameSheet(CellLocation a, CellLocation b) =>
        a.WorkbookName == b.WorkbookName && a.SheetName == b.SheetName;

    // ------------------------------------------------------------------------------------------------------------
    // Small state helpers
    // ------------------------------------------------------------------------------------------------------------

    /// <summary>Pushes the current root on the Back stack (never twice in a row, never when there is no root).</summary>
    private void PushRoot()
    {
        if (_root is null) return;
        if (_stack.Count > 0 && _stack[^1] == _root.Location) return;
        _stack.Add(_root.Location);
    }

    /// <summary>Sets or removes the gold node (the view paints it) and refreshes Back.</summary>
    private void SetGold(Gold? gold)
    {
        if (_gold is null && gold is null) return;
        _gold = gold;
        _view.SetGold(gold?.Node);
        UpdateBack();
    }

    private void UpdateBack() => _view.SetBackEnabled(BackEnabled);

    /// <summary>Shows a message box later, never from inside the click handler that caused it.</summary>
    private void PostMessage(string? message)
    {
        if (string.IsNullOrEmpty(message)) return;
        _scheduler.Post(() =>
        {
            if (!IsOpen) return;
            try
            {
                _view.ShowMessage(message);
            }
            catch (Exception ex)
            {
                Log.Error("ExplorerSession.ShowMessage", ex);
            }
        });
    }

    /// <summary>After a busy evaluation in <see cref="GetNodeInfo"/>, ask the view to re-query the hovered node in 200 ms.</summary>
    private void ScheduleHoverRetry()
    {
        if (_hoverRetryPending || _hoverRetries >= MaxBusyRetries) return;
        _hoverRetryPending = true;
        _hoverRetries++;
        _scheduler.PostDelayed(BusyRetryDelay, () =>
        {
            _hoverRetryPending = false;
            if (!IsOpen) return;
            try
            {
                _view.RefreshHover();
            }
            catch (Exception ex)
            {
                Log.Error("ExplorerSession.RefreshHover", ex);
            }
        });
    }

    // ------------------------------------------------------------------------------------------------------------
    // Command runner
    // ------------------------------------------------------------------------------------------------------------

    /// <summary>
    /// Runs one handler as a "command": no other handler runs inline while it executes (a command that arrives meanwhile
    /// is posted, recorded gateway events wait for the drain that the end of the command schedules), an
    /// <see cref="ExcelBusyException"/> is retried every 200 ms at most 25 times (a user command that Excel rejects with a bare VBA_E_IGNORE: at most
    /// <see cref="UserCommandNotReadyRetries"/> times, so that it cannot fire seconds later), any other exception is logged.
    /// </summary>
    /// <param name="loadVersion">For a load: its version, so that giving up can fall back to plain text.</param>
    private void Execute(string context, Action action, int attempt = 0, int? loadVersion = null, bool isDrain = false, bool userCommand = false)
    {
        if (!IsOpen) return;
        if (_inCommand)
        {
            _scheduler.Post(() => Execute(context, action, attempt, loadVersion, isDrain, userCommand));
            return;
        }

        _inCommand = true;
        bool holdDrain = false;
        try
        {
            // The drain (this one or its retry) is running now and takes whatever is recorded.
            if (isDrain) _drainBackingOff = false;
            action();
        }
        catch (ExcelBusyException ex) when (userCommand && ex.IsEditMode)
        {
            // The user clicked while a cell is being edited: Excel cannot navigate without ending the edit. Do nothing
            // (as the original did) instead of retrying: a click must not fire seconds later, when the edit ends.
            Log.Info($"ExplorerSession: '{context}' ignored, Excel is in cell edit mode.");
        }
        catch (ExcelBusyException ex)
        {
            int budget = userCommand && ex.IsNotReady ? UserCommandNotReadyRetries : MaxBusyRetries;
            if (attempt < budget)
            {
                // A drain keeps its items and backs off once: no immediate drain may be posted meanwhile (see finally).
                if (isDrain) _drainBackingOff = true;
                _scheduler.PostDelayed(BusyRetryDelay, () => Execute(context, action, attempt + 1, loadVersion, isDrain, userCommand));
            }
            else
            {
                Log.Info($"ExplorerSession: Excel stayed busy, giving up on '{context}'.");
                if (loadVersion is { } version) ShowFallback(version);
                // The items stay recorded, but only a new event or command may try them again; draining right now would
                // start the same busy loop from attempt 0.
                holdDrain = isDrain;
            }
        }
        catch (Exception ex)
        {
            Log.Error($"ExplorerSession.{context}", ex);
            if (loadVersion is { } version) ShowFallback(version);
        }
        finally
        {
            _inCommand = false;
            if (HasPendingEvents && !holdDrain) ScheduleDrain();
        }
    }
}
