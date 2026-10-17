using System.Runtime.InteropServices;
using FormulaExplorer.AddIn.Excel;

namespace FormulaExplorer.AddIn.Tests;

/// <summary>Pure tests (no Excel) of the decisions the gateway takes around outline lines, selection tracking and the deferred restore.</summary>
public class ExcelGatewayRulesTests
{
    private static readonly int ProtectedSheetHr = unchecked((int)0x800A03EC);
    private static readonly int VbaIgnoreHr = unchecked((int)0x800AC472);

    private static Func<int, LineState> Sheet(Dictionary<int, LineState> lines) =>
        line => lines.TryGetValue(line, out var state) ? state : new LineState(false, 1);

    private static Dictionary<int, LineState> CollapsedGroup(int first, int last, int level = 2)
    {
        var lines = new Dictionary<int, LineState>();
        for (var line = first; line <= last; line++) lines[line] = new LineState(true, level);
        return lines;
    }

    // ---- UnhideHidden (gateway-2) ----------------------------------------------------------------------------

    [Fact]
    public void UnhideHidden_unhides_only_hidden_outline_lines_and_records_them()
    {
        var lines = CollapsedGroup(8, 18);
        lines[10] = new LineState(true, 1); // hidden by hand
        var unhidden = new List<int>();

        OutlineRules.UnhideHidden(9, 12, Sheet(lines), line => lines[line] = new LineState(false, 2), unhidden, _ => { });

        Assert.Equal(new[] { 9, 11, 12 }, unhidden);
        Assert.True(lines[10].Hidden);
        Assert.True(lines[13].Hidden);
    }

    [Fact]
    public void UnhideHidden_stops_quietly_when_the_sheet_is_protected()
    {
        var lines = CollapsedGroup(8, 18);
        var unhidden = new List<int>();
        var attempts = 0;
        var logged = new List<string>();

        OutlineRules.UnhideHidden(8, 18, Sheet(lines), _ => { attempts++; throw new COMException("Unable to set the Hidden property", ProtectedSheetHr); },
            unhidden, logged.Add);

        Assert.Equal(1, attempts); // one COM exception, not one per line
        Assert.Empty(unhidden);
        Assert.Single(logged);
    }

    [Fact]
    public void UnhideHidden_keeps_the_lines_changed_before_a_refusal()
    {
        var lines = CollapsedGroup(8, 18);
        var unhidden = new List<int>();

        OutlineRules.UnhideHidden(8, 18, Sheet(lines), line =>
        {
            if (line == 10) throw new COMException("refused", ProtectedSheetHr);
            lines[line] = new LineState(false, 2);
        }, unhidden, _ => { });

        Assert.Equal(new[] { 8, 9 }, unhidden);
    }

    [Fact]
    public void UnhideHidden_lets_a_busy_failure_propagate()
    {
        var lines = CollapsedGroup(8, 18);
        var unhidden = new List<int>();

        Assert.Throws<COMException>(() => OutlineRules.UnhideHidden(8, 18, Sheet(lines),
            _ => throw new COMException("busy", VbaIgnoreHr), unhidden, _ => { }));
    }

    // ---- LinesToRehide (gateway-3) ---------------------------------------------------------------------------

    [Fact]
    public void Rehide_hides_a_line_that_is_still_inside_a_collapsed_group()
    {
        var lines = CollapsedGroup(8, 18);
        lines[12] = new LineState(false, 2); // the one we unhid

        Assert.Equal(new[] { 12 }, OutlineRules.LinesToRehide(new[] { 12 }, Sheet(lines)));
    }

    [Fact]
    public void Rehide_hides_the_whole_record_when_the_target_covered_the_group()
    {
        var lines = new Dictionary<int, LineState>();
        for (var line = 8; line <= 18; line++) lines[line] = new LineState(false, 2);
        var recorded = Enumerable.Range(8, 11).ToArray();

        Assert.Equal(recorded, OutlineRules.LinesToRehide(recorded, Sheet(lines)));
    }

    [Fact]
    public void Rehide_skips_a_line_of_a_group_the_user_expanded()
    {
        var lines = new Dictionary<int, LineState>();
        for (var line = 8; line <= 18; line++) lines[line] = new LineState(false, 2); // expanded with '+'

        Assert.Empty(OutlineRules.LinesToRehide(new[] { 12 }, Sheet(lines)));
    }

    [Fact]
    public void Rehide_skips_a_line_that_is_not_an_outline_line_any_more()
    {
        // Rows were deleted: the recorded number now points at a plain visible row (e.g. the summary row).
        var lines = CollapsedGroup(8, 17);
        lines[19] = new LineState(false, 1);

        Assert.Empty(OutlineRules.LinesToRehide(new[] { 19 }, Sheet(lines)));
    }

    [Fact]
    public void Rehide_skips_a_line_that_is_hidden_already()
    {
        var lines = CollapsedGroup(8, 18);

        Assert.Empty(OutlineRules.LinesToRehide(new[] { 12 }, Sheet(lines)));
    }

    [Fact]
    public void Rehide_ignores_a_hidden_neighbour_of_a_shallower_level()
    {
        var lines = new Dictionary<int, LineState>
        {
            [11] = new LineState(true, 1), // hidden by hand, not part of the collapse
            [12] = new LineState(false, 2),
            [13] = new LineState(false, 2),
        };

        Assert.Empty(OutlineRules.LinesToRehide(new[] { 12 }, Sheet(lines)));
    }

    // ---- SelectionTracker (session-4) -----------------------------------------------------------------------

    private static CellLocation Cell(string address) => new("Book1.xlsx", "Sheet1", address);

    [Fact]
    public void Collapsing_a_selected_range_to_its_active_cell_is_a_user_selection()
    {
        var tracker = new SelectionTracker(Cell("$J$31"), "$J$31");
        tracker.Sync(Cell("$I$19"), "$I$19:$J$19"); // NavigateTo selected the OFFSET result

        Assert.True(tracker.Observe(Cell("$I$19"), "$I$19"));
    }

    [Fact]
    public void The_same_selection_is_not_reported_again()
    {
        var tracker = new SelectionTracker(Cell("$A$1"), "$A$1");

        Assert.False(tracker.Observe(Cell("$A$1"), "$A$1"));
        Assert.True(tracker.Observe(Cell("$B$2"), "$B$2"));
        Assert.False(tracker.Observe(Cell("$B$2"), "$B$2"));
    }

    [Fact]
    public void A_different_active_cell_is_reported_even_when_the_address_is_unknown()
    {
        var tracker = new SelectionTracker(Cell("$A$1"), "$A$1");

        Assert.True(tracker.Observe(Cell("$B$2"), null));
    }

    [Fact]
    public void A_non_range_selection_with_the_same_active_cell_is_unchanged()
    {
        var tracker = new SelectionTracker(Cell("$A$1"), "$A$1");

        Assert.False(tracker.Observe(Cell("$A$1"), null)); // a shape was selected
        Assert.False(tracker.Observe(Cell("$A$1"), "$A$1"));
    }

    [Fact]
    public void Activating_a_window_with_the_same_selection_is_not_a_change()
    {
        var tracker = new SelectionTracker(Cell("$I$19"), "$I$19:$J$19");

        Assert.False(tracker.Observe(Cell("$I$19"), "$I$19:$J$19"));
    }

    [Fact]
    public void Extending_the_selection_from_the_active_cell_is_a_user_selection()
    {
        var tracker = new SelectionTracker(Cell("$A$1"), "$A$1");

        Assert.True(tracker.Observe(Cell("$A$1"), "$A$1:$A$10"));
    }

    [Fact]
    public void Sync_does_not_report_and_keeps_the_known_part_when_one_is_unavailable()
    {
        var tracker = new SelectionTracker(null, null);
        Assert.Null(tracker.Location);

        tracker.Sync(Cell("$C$3"), "$C$3");
        tracker.Sync(null, null);

        Assert.Equal(Cell("$C$3"), tracker.Location);
        Assert.False(tracker.Observe(Cell("$C$3"), "$C$3"));
    }

    // ---- RetryLoop (gateway-4) ------------------------------------------------------------------------------

    private sealed class ManualScheduler
    {
        public readonly Queue<Action> Pending = new();
        public void Schedule(Action step) => Pending.Enqueue(step);
        public void RunAll() { while (Pending.Count > 0) Pending.Dequeue()(); }
    }

    [Fact]
    public void RetryLoop_never_runs_inline()
    {
        var scheduler = new ManualScheduler();
        var runs = 0;

        RetryLoop.Start(() => { runs++; return true; }, 5, scheduler.Schedule);

        Assert.Equal(0, runs);
        scheduler.RunAll();
        Assert.Equal(1, runs);
    }

    [Fact]
    public void RetryLoop_retries_until_the_attempt_succeeds()
    {
        var scheduler = new ManualScheduler();
        var runs = 0;

        RetryLoop.Start(() => ++runs == 3, 10, scheduler.Schedule);
        scheduler.RunAll();

        Assert.Equal(3, runs);
    }

    [Fact]
    public void RetryLoop_gives_up_after_the_maximum_number_of_attempts()
    {
        var scheduler = new ManualScheduler();
        var runs = 0;

        RetryLoop.Start(() => { runs++; return false; }, 4, scheduler.Schedule);
        scheduler.RunAll();

        Assert.Equal(4, runs);
    }

    [Fact]
    public void RetryLoop_stops_and_does_not_throw_when_the_attempt_throws()
    {
        var scheduler = new ManualScheduler();
        var runs = 0;

        RetryLoop.Start(() => { runs++; throw new InvalidOperationException("boom"); }, 4, scheduler.Schedule);
        scheduler.RunAll();

        Assert.Equal(1, runs);
    }

    // ---- RestoreLedger.Detach (gateway-4) -------------------------------------------------------------------

    [Fact]
    public void Detach_moves_the_records_in_order_and_empties_the_ledger()
    {
        var ledger = new RestoreLedger();
        ledger.AddSheetVisibility("B.xlsx", "Hidden", 0);
        ledger.AddRows("B.xlsx", "Hidden", new[] { 5, 6 });

        var detached = ledger.Detach();

        Assert.Equal(0, ledger.Count);
        Assert.Equal(new[] { RestoreKind.SheetVisibility, RestoreKind.Rows }, detached.Records.Select(r => r.Kind));
        // The detached ledger still decides what to undo: with nothing to keep, everything, newest first.
        Assert.Equal(new[] { RestoreKind.Rows, RestoreKind.SheetVisibility }, detached.SelectUndo(Array.Empty<CellLocation>()).Select(r => r.Kind));
    }
}
