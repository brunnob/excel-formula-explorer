using System.Runtime.InteropServices;
using FormulaExplorer.AddIn.Diagnostics;

namespace FormulaExplorer.AddIn.UI;

/// <summary>
/// Gives Esc and Enter to the explorer window even though it lives inside Excel's message loop.
/// <para>
/// Excel's pump runs every keyboard message through its own accelerator table (<c>TranslateAccelerator</c> ignores which
/// window the message is for), so a real Esc or Enter typed into a foreign window such as ours is swallowed (it turns into a
/// WM_COMMAND for Excel) before <c>DispatchMessage</c> and never reaches WPF's window procedure; posting the same message
/// directly to the window works, which is why this only shows with real keystrokes. A thread-level <c>WH_GETMESSAGE</c>
/// hook runs inside <c>GetMessage</c>/<c>PeekMessage</c>, i.e. before Excel's pump sees the message, so we take the key
/// there and turn the message into <c>WM_NULL</c>.
/// </para>
/// The hook lives only while the window does, runs on Excel's main thread and ignores every message whose target is not
/// this window (Excel's own windows, message boxes, the Find dialog...).
/// <para>
/// The hook only CLAIMS the key (rewrites it to WM_NULL) and queues the action with
/// <c>Dispatcher.BeginInvoke(DispatcherPriority.Input)</c>: it runs inside whichever <c>PeekMessage</c>/<c>GetMessage</c>
/// removed the message (Excel's interrupt checks during a long calculation, a modal loop, the window's own move/size
/// loop...), where running the action (a navigation through the object model, closing the window) would re-enter code that
/// is in the middle of something. Auto-repeats are claimed too (so they never reach Excel's accelerators) but do nothing.
/// While the window is being moved or sized (WM_ENTERSIZEMOVE..WM_EXITSIZEMOVE) nothing is intercepted: Esc cancels the
/// move there, as Windows defines it.
/// </para>
/// <para>
/// The same hook also sees real mouse clicks. Excel's own windows answer WM_MOUSEACTIVATE with MA_NOACTIVATE when the active
/// window is one of the windows it owns (its grid, <c>EXCEL7</c>, does so for our explorer exactly as for its own modeless
/// Find dialog), so a click on a cell would select it but leave the explorer active: its title stays blue and the keys the
/// user types next go to the explorer, not to the cell. When the explorer is the active window and a click lands on a child
/// of an Excel main window, Excel's window is activated right after the click (queued like the keys); the explorer stays on
/// top because it is owned by that window, with its title greyed, as the original did.
/// </para>
/// </summary>
internal sealed class WindowKeyHook : IDisposable
{
    private const int WH_GETMESSAGE = 3;
    private const int PM_REMOVE = 1;
    private const uint GA_ROOT = 2;
    private const uint WM_LBUTTONDOWN = 0x0201;
    private const uint WM_RBUTTONDOWN = 0x0204;

    private delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    private struct MSG
    {
        public IntPtr Hwnd;
        public uint Message;
        public IntPtr WParam;
        public IntPtr LParam;
        public uint Time;
        public int PtX;
        public int PtY;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookExW(int idHook, HookProc proc, IntPtr module, uint threadId);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hook);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);

    [DllImport("user32.dll")]
    private static extern IntPtr GetActiveWindow();

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetForegroundWindow(IntPtr hwnd);

    [DllImport("user32.dll", EntryPoint = "GetClassNameW", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hwnd, System.Text.StringBuilder buffer, int maxCount);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();

    private readonly HookProc _proc;          // kept in a field: the native hook holds a pointer to it
    private readonly Func<IntPtr> _windowHandle;
    private readonly Func<bool> _suspended;
    private readonly Action<int, uint> _onKey;
    private readonly System.Windows.Threading.Dispatcher _dispatcher;
    private IntPtr _hook;
    private bool _disposed;

    /// <param name="windowHandle">The explorer window's HWND (read on every message).</param>
    /// <param name="suspended">True while the hook must leave every key alone (the window is being moved or sized).</param>
    /// <param name="onKey">Acts on VK_ESCAPE / VK_RETURN (key and the message time); runs queued on <paramref name="dispatcher"/>, never inside the hook.</param>
    public WindowKeyHook(Func<IntPtr> windowHandle, Func<bool> suspended, Action<int, uint> onKey, System.Windows.Threading.Dispatcher dispatcher)
    {
        _windowHandle = windowHandle;
        _suspended = suspended;
        _onKey = onKey;
        _dispatcher = dispatcher;
        _proc = HookCallback;
        _hook = SetWindowsHookExW(WH_GETMESSAGE, _proc, IntPtr.Zero, GetCurrentThreadId());
        if (_hook == IntPtr.Zero)
            Log.Info($"WindowKeyHook: SetWindowsHookEx failed (error {Marshal.GetLastWin32Error()}); Esc/Enter may not reach the window.");
    }

    /// <summary>True for the messages the hook takes over: key-down of Esc or Enter.</summary>
    internal static bool IsInterceptedKey(uint message, long virtualKey) =>
        message == NativeMethods.WM_KEYDOWN && virtualKey is NativeMethods.VK_ESCAPE or NativeMethods.VK_RETURN;

    /// <summary>
    /// True for an auto-repeat WM_KEYDOWN: bit 30 of lParam is the previous key state (set when the key was already down).
    /// Unlike "ignore until the key-up arrives" this does not depend on the matching WM_KEYUP getting through Excel's pump.
    /// </summary>
    internal static bool IsAutoRepeat(long lParam) => (lParam & (1L << 30)) != 0;

    private IntPtr HookCallback(int code, IntPtr wParam, IntPtr lParam)
    {
        try
        {
            if (code >= 0 && wParam.ToInt64() == PM_REMOVE && lParam != IntPtr.Zero && !_disposed)
            {
                var msg = Marshal.PtrToStructure<MSG>(lParam);
                int virtualKey = (int)(msg.WParam.ToInt64() & 0xFFFF);
                if (msg.Message is WM_LBUTTONDOWN or WM_RBUTTONDOWN)
                {
                    ActivateExcelAfterClick(msg.Hwnd);
                }
                else if (IsInterceptedKey(msg.Message, virtualKey) && IsOurs(msg.Hwnd) && !_suspended())
                {
                    // Excel's pump gets a harmless WM_NULL instead of the key; the action runs later, outside this PeekMessage.
                    Marshal.WriteInt32(lParam, IntPtr.Size, 0);
                    if (!IsAutoRepeat(msg.LParam.ToInt64()))
                    {
                        uint time = msg.Time;
                        _dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Input, new Action(() => Deliver(virtualKey, time)));
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Log.Error("WindowKeyHook", ex);
        }

        return CallNextHookEx(_hook, code, wParam, lParam);
    }

    /// <summary>
    /// A mouse button went down on <paramref name="target"/>. When the explorer is the active window and the target is a child of
    /// an Excel main window, queues the activation of that main window (it runs after Excel handled the click).
    /// </summary>
    private void ActivateExcelAfterClick(IntPtr target)
    {
        var window = _windowHandle();
        if (window == IntPtr.Zero || target == IntPtr.Zero || target == window || GetActiveWindow() != window) return;

        var root = GetAncestor(target, GA_ROOT);
        if (root == IntPtr.Zero || root == window || !IsExcelMainWindow(root)) return;

        _dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Input, new Action(() =>
        {
            try
            {
                // Still the explorer (Excel did not activate itself) and Excel's window still there.
                if (!_disposed && GetActiveWindow() == _windowHandle() && NativeMethods.IsWindow(root)) SetForegroundWindow(root);
            }
            catch (Exception ex)
            {
                Log.Error("WindowKeyHook.ActivateExcel", ex);
            }
        }));
    }

    private static bool IsExcelMainWindow(IntPtr hwnd)
    {
        var name = new System.Text.StringBuilder(32);
        return GetClassName(hwnd, name, name.Capacity) > 0 && name.ToString() == "XLMAIN";
    }

    private void Deliver(int virtualKey, uint time)
    {
        if (_disposed) return;
        try
        {
            _onKey(virtualKey, time);
        }
        catch (Exception ex)
        {
            Log.Error("WindowKeyHook.Deliver", ex);
        }
    }

    private bool IsOurs(IntPtr target)
    {
        var window = _windowHandle();
        return window != IntPtr.Zero && target != IntPtr.Zero && (target == window || GetAncestor(target, GA_ROOT) == window);
    }

    public void Dispose()
    {
        _disposed = true;
        var hook = _hook;
        _hook = IntPtr.Zero;
        if (hook != IntPtr.Zero) UnhookWindowsHookEx(hook);
    }
}
