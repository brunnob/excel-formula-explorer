using System.Windows;
using System.Windows.Threading;
using ExcelDna.Integration;
using FormulaExplorer.AddIn.Diagnostics;
using FormulaExplorer.AddIn.Excel;
using FormulaExplorer.AddIn.Overlay;
using FormulaExplorer.AddIn.Session;
using FormulaExplorer.AddIn.UI;
using XL = Microsoft.Office.Interop.Excel;

namespace FormulaExplorer.AddIn.Composition;

/// <summary>
/// Composition root and lifecycle owner: creates the gateway, overlay, window and session when the ribbon button is
/// pressed and tears everything down again when the explorer closes. At most one explorer is open at a time.
/// Everything runs on Excel's main thread.
/// </summary>
public static class ExplorerController
{
    private const string Title = "Formula Explorer";

    /// <summary>How often a pending workbook close is polled.</summary>
    private static readonly TimeSpan ClosePollInterval = TimeSpan.FromMilliseconds(250);

    /// <summary>The objects of the open explorer; null when closed. Only touched on Excel's main thread.</summary>
    private static Instance? _current;

    /// <summary>True while an explorer is open (or being opened).</summary>
    public static bool IsOpen => _current != null;

    public static ExplorerSession? CurrentSession => _current?.Session;
    public static ExplorerWindow? CurrentWindow => _current?.Window;
    public static PrecedentOverlay? CurrentOverlay => _current?.Overlay;

    /// <summary>
    /// Opens the explorer, or brings the open one to the front. Never throws: a failure is logged, the partly built
    /// explorer is torn down and the user sees "Formula Explorer could not start: ...".
    /// </summary>
    public static void Open()
    {
        try
        {
            OpenOrThrow();
        }
        catch (Exception ex)
        {
            ShowError($"Formula Explorer could not start: {ex.Message}");
        }
    }

    /// <summary>
    /// Same as <see cref="Open"/> but reports the failure to the caller (test hooks) instead of showing a message box.
    /// The failure is logged and everything built so far is torn down before the exception is rethrown.
    /// </summary>
    internal static void OpenOrThrow()
    {
        if (_current != null)
        {
            if (_current.Window is { IsAlive: true } window)
            {
                window.Activate();
                return;
            }

            // The window died without a clean close (e.g. destroyed together with its owner workbook window):
            // tear the stale instance down and open a fresh one instead of leaving the ribbon button dead.
            CloseAll();
        }

        var instance = new Instance();
        _current = instance;
        try
        {
            instance.Start();
        }
        catch (Exception ex)
        {
            Log.Error("ExplorerController.Open", ex);
            CloseAll();
            throw;
        }
    }

    /// <summary>
    /// Closes session, window, overlay and gateway (in that order). Idempotent, safe to call from inside the window's
    /// Closing event or from <c>AutoClose</c>; every step is attempted even if an earlier one fails.
    /// </summary>
    public static void CloseAll()
    {
        // Detach first: later calls (including re-entrant ones from the window's Closing) see "closed".
        var instance = _current;
        _current = null;
        instance?.Dispose();
    }

    /// <summary>Shows an error message box (and logs it); used by entry points that cannot propagate exceptions.</summary>
    internal static void ShowError(string message)
    {
        Log.Info($"Showing error to the user: {message}");
        try
        {
            MessageBox.Show(message, Title, MessageBoxButton.OK, MessageBoxImage.Error);
        }
        catch (Exception ex)
        {
            Log.Error("ExplorerController.ShowError", ex);
        }
    }

    /// <summary>One open explorer: the objects that live and die together.</summary>
    private sealed class Instance : IDisposable
    {
        private ExcelGateway? _gateway;
        private PrecedentOverlay? _overlay;
        private ExplorerWindow? _window;
        private ExplorerSession? _session;
        private DispatcherUiScheduler? _scheduler;
        private bool _disposed;
        private readonly WorkbookCloseTracker _closing = new();
        private bool _closePollArmed;

        // Null until Start() has built them (and after a failed Start).
        public ExplorerSession? Session => _session;
        public PrecedentOverlay? Overlay => _overlay;
        public ExplorerWindow? Window => _window;

        public void Start()
        {
            var app = (XL.Application)ExcelDnaUtil.Application;
            _gateway = new ExcelGateway(app);
            _overlay = new PrecedentOverlay(app);
            var relay = new NodeInfoRelay();
            _window = new ExplorerWindow(_gateway.GetActiveWindowHandle(), relay);
            _scheduler = new DispatcherUiScheduler(Dispatcher.CurrentDispatcher);
            _session = new ExplorerSession(_gateway, _window, _overlay, _scheduler);
            relay.Target = _session;

            _window.CloseRequested += OnCloseRequested;
            _gateway.ActiveWindowChanged += OnActiveWindowChanged;
            _gateway.WorkbookClosing += OnWorkbookClosing;
            _gateway.WindowDeactivating += OnWindowDeactivating;

            _window.Show();
            _overlay.SetExplorerWindow(_window.Handle);
            _session.Start();
        }

        private void OnCloseRequested(object? sender, EventArgs e)
        {
            try
            {
                CloseAll();
            }
            catch (Exception ex)
            {
                Log.Error("ExplorerController.OnCloseRequested", ex);
            }
        }

        private void OnActiveWindowChanged(object? sender, IntPtr hwnd)
        {
            try
            {
                if (_disposed) return;
                // A window of a workbook whose close is pending (its "Save changes?" prompt, for one, activates it again)
                // must not get the explorer: it would be destroyed with that window when the close goes through.
                if (_closing.IsClosingWindow(hwnd)) return;
                _window?.SetOwnerWindow(hwnd);
            }
            catch (Exception ex)
            {
                Log.Error("ExplorerController.OnActiveWindowChanged", ex);
            }
        }

        /// <summary>
        /// A workbook window is being deactivated while it still exists. That is also how Excel announces that ONE window of a
        /// workbook that has several is being closed (<c>Window.Close</c>: no BeforeClose, the window is destroyed right after and
        /// would take the explorer with it). So when the explorer is owned by the deactivating window, it is handed to another
        /// window at once (the root workbook's other window if any, else any other workbook window); the activation that
        /// follows a plain window switch re-owns it to the right window anyway.
        /// </summary>
        private void OnWindowDeactivating(object? sender, IntPtr hwnd)
        {
            try
            {
                if (_disposed || _gateway == null || _window == null || _window.OwnerWindow != hwnd) return;

                var rootBook = _session?.GetSnapshot().Root?.WorkbookName;
                var rootWindows = rootBook != null ? _gateway.GetWindowHandlesOf(rootBook) : (IReadOnlyList<IntPtr>)Array.Empty<IntPtr>();
                var allWindows = _gateway.GetAllWindowHandles();

                // First choice: a window the user can see, the root workbook's other window first, else the window Excel reports as
                // active (once the deactivating one is no longer it: the window the switch goes to), else any other. A hidden window
                // (PERSONAL.XLSB) or a minimized one would leave the explorer owned by a window that hides it if no activation follows.
                var alternative = rootWindows.FirstOrDefault(h => Usable(h, hwnd, visibleOnly: true));
                if (alternative == IntPtr.Zero)
                {
                    var active = ActiveWindowOrZero();
                    if (Usable(active, hwnd, visibleOnly: true)) alternative = active;
                }
                if (alternative == IntPtr.Zero) alternative = allWindows.FirstOrDefault(h => Usable(h, hwnd, visibleOnly: true));

                // Last resort: an existing window nobody sees (hidden or minimized). Owned by a minimized window the explorer is hidden
                // until that window is restored (a hidden owner does not hide it), but it stays alive, and the activation that follows
                // the closing of a window re-owns it; owned by the window that is about to be destroyed it would die with it.
                if (alternative == IntPtr.Zero) alternative = rootWindows.Concat(allWindows).FirstOrDefault(h => Usable(h, hwnd, visibleOnly: false));
                Log.Info($"ExplorerController: window 0x{hwnd.ToInt64():X} is being deactivated; the explorer goes to " +
                         (alternative == IntPtr.Zero ? "no other window" : $"window 0x{alternative.ToInt64():X}"));
                if (alternative != IntPtr.Zero) _window.SetOwnerWindow(alternative);
            }
            catch (ExcelBusyException)
            {
                // Excel cannot answer right now; the activation that follows re-owns the explorer anyway.
            }
            catch (Exception ex)
            {
                Log.Error("ExplorerController.OnWindowDeactivating", ex);
            }
        }

        /// <summary>
        /// The active workbook window, or zero. A busy Excel answers zero instead of throwing, so that one unanswered call
        /// does not cancel a whole hand-over (the visible and last-resort candidates still apply).
        /// </summary>
        private IntPtr ActiveWindowOrZero()
        {
            try
            {
                return _gateway?.GetActiveWindowHandle() ?? IntPtr.Zero;
            }
            catch (ExcelBusyException)
            {
                return IntPtr.Zero;
            }
        }

        /// <summary>
        /// Whether the explorer may be handed to <paramref name="candidate"/>: another existing window, not one of a pending close,
        /// and (<paramref name="visibleOnly"/>) one the user can see: neither hidden nor minimized.
        /// </summary>
        private bool Usable(IntPtr candidate, IntPtr deactivating, bool visibleOnly) =>
            candidate != IntPtr.Zero && candidate != deactivating && !_closing.IsClosingWindow(candidate) && UI.NativeMethods.IsWindow(candidate)
            && (!visibleOnly || (UI.NativeMethods.IsWindowVisible(candidate) && !UI.NativeMethods.IsIconic(candidate)));

        /// <summary>
        /// A workbook is closing (the event comes BEFORE Excel's "Save changes?" prompt and cannot say whether the close goes
        /// through). The explorer window is owned by the active workbook's window (so that it stays on top of it) and
        /// Windows destroys owned windows together with their owner, so:
        /// <list type="bullet">
        /// <item>when the explorer is currently owned by a window of a closing workbook other than the root's, it is handed to
        /// the root workbook's window at once;</item>
        /// <item>the close is then polled (<see cref="WorkbookCloseTracker"/>) until it ends. While it is pending nothing
        /// is ever re-owned to the closing workbook's windows, however long the prompt stays up; when it ended, the explorer
        /// follows the active window again (a cancelled close leaves the old workbook active; a confirmed one activates another);</item>
        /// <item>when the root's workbook is the one that closed, the explorer is closed explicitly once the close is
        /// confirmed (whatever window owned it), since the session has nothing left to explore.</item>
        /// </list>
        /// </summary>
        private void OnWorkbookClosing(object? sender, string workbookName)
        {
            try
            {
                if (_disposed || _gateway == null || _window == null) return;
                var rootBook = _session?.GetSnapshot().Root?.WorkbookName;
                bool isRoot = rootBook != null && string.Equals(rootBook, workbookName, StringComparison.OrdinalIgnoreCase);
                var windows = _gateway.GetWindowHandlesOf(workbookName);
                _closing.Begin(workbookName, isRoot, windows);
                Log.Info($"ExplorerController: '{workbookName}' is closing (root: {isRoot}, explorer owned by one of its windows: {windows.Contains(_window.OwnerWindow)})");

                if (!isRoot && rootBook != null && windows.Contains(_window.OwnerWindow))
                {
                    var keep = _gateway.GetWindowHandleOf(rootBook);
                    if (keep != IntPtr.Zero && !_closing.IsClosingWindow(keep)) _window.SetOwnerWindow(keep);
                }
                ArmClosePoll();
            }
            catch (Exception ex)
            {
                Log.Error("ExplorerController.OnWorkbookClosing", ex);
            }
        }

        private void ArmClosePoll()
        {
            if (_closePollArmed || _disposed || !_closing.Any) return;
            _closePollArmed = true;
            _scheduler?.PostDelayed(ClosePollInterval, PollClosing);
        }

        private void PollClosing()
        {
            _closePollArmed = false;
            try
            {
                if (_disposed || _gateway == null) return;
                bool follow = false, closeExplorer = false;
                foreach (var result in _closing.Poll(ProbeClosingWorkbook))
                {
                    Log.Info($"ExplorerController: close of '{result.Workbook}' ended: {result.Outcome}");
                    if (result.Outcome == CloseOutcome.Closed && result.WasRoot) closeExplorer = true;
                    follow = true;
                }

                if (closeExplorer)
                {
                    if (ReferenceEquals(_current, this)) CloseAll();
                    return;
                }
                if (follow) FollowActiveWindow();
                ArmClosePoll();
            }
            catch (Exception ex)
            {
                Log.Error("ExplorerController.PollClosing", ex);
                ArmClosePoll();
            }
        }

        /// <summary>Open / gone / busy for a pending close; busy also when a modal dialog disabled the workbook's windows.</summary>
        private WorkbookPresence ProbeClosingWorkbook(string workbook)
        {
            var presence = _gateway!.ProbeWorkbook(workbook);
            if (presence != WorkbookPresence.Open) return presence;
            foreach (var hwnd in _gateway.GetWindowHandlesOf(workbook))
                if (UI.NativeMethods.IsWindow(hwnd) && !UI.NativeMethods.IsWindowEnabled(hwnd)) return WorkbookPresence.Busy;
            return WorkbookPresence.Open;
        }

        /// <summary>After a close ended the active window is the one to sit on top of again (never one of a still-closing workbook).</summary>
        private void FollowActiveWindow()
        {
            try
            {
                if (_disposed || _gateway == null) return;
                var active = _gateway.GetActiveWindowHandle();
                if (!_closing.IsClosingWindow(active)) _window?.SetOwnerWindow(active);
            }
            catch (Exception ex)
            {
                Log.Error("ExplorerController.FollowActiveWindow", ex);
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            if (_window != null) _window.CloseRequested -= OnCloseRequested;
            if (_gateway != null)
            {
                _gateway.ActiveWindowChanged -= OnActiveWindowChanged;
                _gateway.WorkbookClosing -= OnWorkbookClosing;
                _gateway.WindowDeactivating -= OnWindowDeactivating;
            }

            // Session first: it clears the overlay and restores whatever the gateway unhid/expanded.
            Attempt("session.Close", () => _session?.Close());
            Attempt("window.CloseWindow", () => _window?.CloseWindow());
            Attempt("overlay.Dispose", () => _overlay?.Dispose());
            Attempt("gateway.Dispose", () => _gateway?.Dispose());
        }

        private static void Attempt(string step, Action action)
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                Log.Error($"ExplorerController.CloseAll ({step})", ex);
            }
        }
    }
}
