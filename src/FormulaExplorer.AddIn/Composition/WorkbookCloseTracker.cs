using FormulaExplorer.AddIn.Excel;

namespace FormulaExplorer.AddIn.Composition;

/// <summary>How a pending close ended.</summary>
internal enum CloseOutcome
{
    /// <summary>The workbook is gone.</summary>
    Closed,

    /// <summary>The workbook is still open and Excel is no longer busy: the user (or a handler) cancelled the close.</summary>
    Cancelled,
}

/// <summary>A pending close that has ended.</summary>
/// <param name="WasRoot">True when the closing workbook is the one the explorer's root cell belongs to.</param>
internal readonly record struct CloseResult(string Workbook, bool WasRoot, CloseOutcome Outcome);

/// <summary>
/// Workbooks for which <c>WorkbookBeforeClose</c> was raised and whose close has not ended yet. That event cannot say
/// whether the close goes through: Excel may show "Save changes?" for as long as the user likes (every Excel call is
/// rejected meanwhile) and the user may cancel. A close is therefore only judged by polling: it is over when the workbook
/// is gone (for <see cref="GonePollsBeforeClosed"/> consecutive polls), and it was cancelled only when the workbook is still open although Excel has been answering (not busy) for
/// <see cref="CallsBeforeCancelled"/> consecutive polls (so the moment between the end of the prompt and Excel actually
/// closing the workbook is not mistaken for a cancel). While a close is pending, nothing may hand the explorer window to
/// one of that workbook's windows: Windows destroys owned windows together with their owner.
/// Pure logic: the probe and the clock are supplied by the caller.
/// </summary>
internal sealed class WorkbookCloseTracker
{
    /// <summary>Consecutive "open and not busy" polls that mean the close was cancelled.</summary>
    internal const int CallsBeforeCancelled = 3;

    /// <summary>
    /// Consecutive "gone" polls that mean the close went through. Two, not one: the probe already confirms an absence with a
    /// complete walk over <c>Application.Workbooks</c>, but a close of the ROOT workbook ends the whole explorer and cannot be
    /// taken back, so a lookup that happens to fail for one poll (a workbook renamed by Save As, a moment between dialogs) must
    /// not be enough.
    /// </summary>
    internal const int GonePollsBeforeClosed = 2;

    /// <summary>Polls after which a close that never resolves is dropped (at 250 ms: about 25 minutes).</summary>
    internal const int MaxPolls = 6000;

    private sealed class Entry(string workbook, bool isRoot, IReadOnlyCollection<IntPtr> windows)
    {
        public string Workbook { get; } = workbook;
        public bool IsRoot { get; } = isRoot;
        public IReadOnlyCollection<IntPtr> Windows { get; } = windows;
        public int OpenAndIdlePolls { get; set; }
        public int GonePolls { get; set; }
        public int Polls { get; set; }
    }

    private readonly List<Entry> _entries = new();

    /// <summary>True while at least one close is pending.</summary>
    public bool Any => _entries.Count > 0;

    /// <summary>Starts (or restarts, when the same workbook is closed again) tracking a close.</summary>
    /// <param name="windows">Every top-level window (XLMAIN) of the closing workbook.</param>
    public void Begin(string workbook, bool isRoot, IReadOnlyCollection<IntPtr> windows)
    {
        _entries.RemoveAll(e => string.Equals(e.Workbook, workbook, StringComparison.OrdinalIgnoreCase));
        _entries.Add(new Entry(workbook, isRoot, windows));
    }

    /// <summary>True when <paramref name="hwnd"/> is a window of a workbook whose close is pending.</summary>
    public bool IsClosingWindow(IntPtr hwnd) => hwnd != IntPtr.Zero && _entries.Any(e => e.Windows.Contains(hwnd));

    /// <summary>True when the workbook's close is pending.</summary>
    public bool IsClosing(string workbook) =>
        _entries.Any(e => string.Equals(e.Workbook, workbook, StringComparison.OrdinalIgnoreCase));

    /// <summary>Runs one poll over every pending close and returns (and forgets) the ones that ended.</summary>
    public IReadOnlyList<CloseResult> Poll(Func<string, WorkbookPresence> probe)
    {
        var ended = new List<CloseResult>();
        foreach (var entry in _entries.ToList())
        {
            entry.Polls++;
            var presence = probe(entry.Workbook);
            if (presence == WorkbookPresence.Gone)
            {
                entry.OpenAndIdlePolls = 0;
                if (++entry.GonePolls >= GonePollsBeforeClosed)
                {
                    ended.Add(new CloseResult(entry.Workbook, entry.IsRoot, CloseOutcome.Closed));
                    _entries.Remove(entry);
                }
                continue;
            }
            entry.GonePolls = 0;

            // Busy: the prompt (or another modal) is up and the user has not decided yet.
            entry.OpenAndIdlePolls = presence == WorkbookPresence.Busy ? 0 : entry.OpenAndIdlePolls + 1;
            if (entry.OpenAndIdlePolls >= CallsBeforeCancelled || entry.Polls >= MaxPolls)
            {
                ended.Add(new CloseResult(entry.Workbook, entry.IsRoot, CloseOutcome.Cancelled));
                _entries.Remove(entry);
            }
        }
        return ended;
    }
}
