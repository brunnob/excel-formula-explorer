using System.Runtime.InteropServices;
using ExcelDna.Integration;
using FormulaExplorer.AddIn.Diagnostics;
using XL = Microsoft.Office.Interop.Excel;

namespace FormulaExplorer.AddIn.Excel;

/// <summary>Bounded retry loop: runs <c>attempt</c> until it reports success, one scheduled step at a time. Never throws.</summary>
internal static class RetryLoop
{
    /// <param name="attempt">Returns true when there is nothing left to do; false to try again later.</param>
    /// <param name="maxAttempts">Total number of attempts.</param>
    /// <param name="scheduleLater">Runs the given step later (never inline).</param>
    public static void Start(Func<bool> attempt, int maxAttempts, Action<Action> scheduleLater)
    {
        void Step(int number)
        {
            bool done;
            try
            {
                done = attempt();
            }
            catch (Exception ex)
            {
                Log.Error("RetryLoop attempt", ex);
                return;
            }
            if (done || number >= maxAttempts) return;
            scheduleLater(() => Step(number + 1));
        }

        scheduleLater(() => Step(1));
    }
}

/// <summary>
/// Finishes the restore of a session that ended while Excel was busy (cell edit mode, a modal dialog): the records that
/// could not be undone are retried later, on Excel's main thread, so a sheet that was hidden or very hidden does not stay
/// visible (and get saved that way). Independent of the gateway's event hooks, which are gone by then.
/// </summary>
internal sealed class DeferredRestore
{
    private const int MaxAttempts = 40;
    private static readonly TimeSpan Interval = TimeSpan.FromMilliseconds(250);

    private readonly XL.Application _app;
    private readonly RestoreLedger _pending;

    private DeferredRestore(XL.Application app, RestoreLedger pending)
    {
        _app = app;
        _pending = pending;
    }

    /// <summary>Takes over the records of <paramref name="pending"/> and keeps retrying their undo in the background.</summary>
    public static void Start(XL.Application app, RestoreLedger pending)
    {
        if (pending.Count == 0) return;
        var restore = new DeferredRestore(app, pending);
        RetryLoop.Start(restore.Attempt, MaxAttempts, ScheduleOnExcel);
    }

    /// <summary>One try; true when finished (done, or nothing more can be done), false when Excel is still busy.</summary>
    private bool Attempt()
    {
        try
        {
            var active = ExcelGateway.ActiveLocation(_app);
            // QueueAsMacro runs this only when Excel accepts macros (not editing a cell, no modal dialog): no edit-mode check here, a
            // real busy answer from the object model is what defers it again.
            ExcelGateway.RestoreCore(_app, _pending, active is null ? Array.Empty<CellLocation>() : new[] { active }, checkEditMode: false);
            Log.Info("ExcelGateway: deferred restore finished");
            return true;
        }
        catch (COMException ex) when (ComErrors.IsBusy(ex))
        {
            return false;
        }
        catch (Exception ex)
        {
            Log.Error("DeferredRestore.Attempt", ex);
            return true;
        }
    }

    private static void ScheduleOnExcel(Action step)
    {
        System.Threading.Timer? timer = null;
        timer = new System.Threading.Timer(_ =>
        {
            timer?.Dispose();
            try
            {
                // Runs on Excel's main thread once Excel accepts macros (not in edit mode, no modal dialog).
                ExcelAsyncUtil.QueueAsMacro(() => step());
            }
            catch (Exception ex)
            {
                Log.Info($"ExcelGateway: deferred restore could not be queued: {ex.Message}");
            }
        }, null, Interval, Timeout.InfiniteTimeSpan);
    }
}
