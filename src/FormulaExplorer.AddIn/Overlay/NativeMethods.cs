using System.Runtime.InteropServices;
using System.Text;

namespace FormulaExplorer.AddIn.Overlay;

/// <summary>
/// Win32 declarations used by the overlay. Every handle is an <see cref="IntPtr"/> and every pointer-sized
/// parameter is pointer-sized, so the declarations are correct on 64-bit Excel (the only supported bitness).
/// </summary>
internal static class NativeMethods
{
    private const string User32 = "user32.dll";
    private const string Gdi32 = "gdi32.dll";
    private const string Kernel32 = "kernel32.dll";

    // Window styles.
    public const uint WS_POPUP = 0x80000000;
    public const uint WS_EX_TRANSPARENT = 0x00000020;
    public const uint WS_EX_TOOLWINDOW = 0x00000080;
    public const uint WS_EX_LAYERED = 0x00080000;
    public const uint WS_EX_NOACTIVATE = 0x08000000;

    // Messages.
    public const uint WM_ERASEBKGND = 0x0014;
    public const uint WM_PAINT = 0x000F;
    public const uint WM_MOUSEACTIVATE = 0x0021;
    public const uint WM_NCDESTROY = 0x0082;
    public const uint WM_NCHITTEST = 0x0084;
    public const uint WM_PRINTCLIENT = 0x0318;

    public static readonly IntPtr HTTRANSPARENT = new(-1);
    public static readonly IntPtr MA_NOACTIVATE = new(3);

    // SetWindowPos.
    public const uint SWP_NOSIZE = 0x0001;
    public const uint SWP_NOMOVE = 0x0002;
    public const uint SWP_NOZORDER = 0x0004;
    public const uint SWP_NOACTIVATE = 0x0010;
    public static readonly IntPtr HWND_TOP = IntPtr.Zero;

    public const int SW_HIDE = 0;
    public const int SW_SHOWNOACTIVATE = 4;

    public const uint LWA_COLORKEY = 0x00000001;

    public const uint GW_HWNDNEXT = 2;
    public const uint GW_HWNDPREV = 3;
    public const uint GA_ROOT = 2;

    public const int GWLP_HWNDPARENT = -8;

    /// <summary>DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2.</summary>
    public static readonly IntPtr DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = new(-4);

    public delegate IntPtr WndProc(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left, Top, Right, Bottom;

        public RECT(PixelRect r)
        {
            Left = r.Left;
            Top = r.Top;
            Right = r.Right;
            Bottom = r.Bottom;
        }

        public readonly PixelRect ToPixelRect() => new(Left, Top, Right, Bottom);
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct WNDCLASSEX
    {
        public uint cbSize;
        public uint style;
        public IntPtr lpfnWndProc;
        public int cbClsExtra;
        public int cbWndExtra;
        public IntPtr hInstance;
        public IntPtr hIcon;
        public IntPtr hCursor;
        public IntPtr hbrBackground;
        public IntPtr lpszMenuName;
        public IntPtr lpszClassName;
        public IntPtr hIconSm;
    }

    /// <summary>PAINTSTRUCT; the 32 reserved bytes are declared as eight ints to keep the struct blittable.</summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct PAINTSTRUCT
    {
        public IntPtr hdc;
        public int fErase;
        public RECT rcPaint;
        public int fRestore;
        public int fIncUpdate;
        public int reserved0, reserved1, reserved2, reserved3, reserved4, reserved5, reserved6, reserved7;
    }

    [DllImport(Kernel32, CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr GetModuleHandleW(string? moduleName);

    [DllImport(User32, SetLastError = true)]
    public static extern ushort RegisterClassExW(ref WNDCLASSEX windowClass);

    [DllImport(User32, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool UnregisterClassW(IntPtr className, IntPtr hInstance);

    [DllImport(User32, CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr CreateWindowExW(
        uint exStyle, IntPtr className, string? windowName, uint style,
        int x, int y, int width, int height,
        IntPtr parentOrOwner, IntPtr menu, IntPtr hInstance, IntPtr param);

    [DllImport(User32, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool DestroyWindow(IntPtr hWnd);

    [DllImport(User32)]
    public static extern IntPtr DefWindowProcW(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport(User32, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetLayeredWindowAttributes(IntPtr hWnd, uint colorKey, byte alpha, uint flags);

    [DllImport(User32, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint flags);

    [DllImport(User32)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ShowWindow(IntPtr hWnd, int command);

    [DllImport(User32, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

    [DllImport(User32)]
    public static extern IntPtr GetWindow(IntPtr hWnd, uint command);

    [DllImport(User32)]
    public static extern IntPtr GetAncestor(IntPtr hWnd, uint flags);

    [DllImport(User32)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsIconic(IntPtr hWnd);

    [DllImport(User32)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsWindow(IntPtr hWnd);

    [DllImport(User32)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport(User32, CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr FindWindowExW(IntPtr parent, IntPtr childAfter, string? className, string? windowTitle);

    [DllImport(User32, CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int GetClassNameW(IntPtr hWnd, StringBuilder className, int maxCount);

    [DllImport(User32, EntryPoint = "SetWindowLongPtrW", SetLastError = true)]
    private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int index, IntPtr newValue);

    [DllImport(User32, EntryPoint = "SetWindowLongW", SetLastError = true)]
    private static extern int SetWindowLong32(IntPtr hWnd, int index, int newValue);

    /// <summary>SetWindowLongPtr is only exported on 64-bit Windows; on 32-bit it is a macro over SetWindowLong.</summary>
    public static IntPtr SetWindowLongPtr(IntPtr hWnd, int index, IntPtr newValue) =>
        IntPtr.Size == 8
            ? SetWindowLongPtr64(hWnd, index, newValue)
            : new IntPtr(SetWindowLong32(hWnd, index, newValue.ToInt32()));

    [DllImport(User32)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool InvalidateRect(IntPtr hWnd, IntPtr rect, [MarshalAs(UnmanagedType.Bool)] bool erase);

    [DllImport(User32)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool UpdateWindow(IntPtr hWnd);

    [DllImport(User32)]
    public static extern IntPtr BeginPaint(IntPtr hWnd, out PAINTSTRUCT paint);

    [DllImport(User32)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EndPaint(IntPtr hWnd, ref PAINTSTRUCT paint);

    [DllImport(User32)]
    public static extern int FillRect(IntPtr hdc, ref RECT rect, IntPtr brush);

    [DllImport(Gdi32)]
    public static extern IntPtr CreateSolidBrush(uint colorRef);

    [DllImport(Gdi32)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool DeleteObject(IntPtr obj);

    [DllImport(User32, SetLastError = true)]
    public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);

    [DllImport(User32)]
    public static extern uint GetDpiForWindow(IntPtr hWnd);

    /// <summary>0xRRGGBB to a GDI COLORREF (0x00BBGGRR).</summary>
    public static uint ToColorRef(uint rgb) => ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 0xFF);

    /// <summary>
    /// The grid window ("EXCEL7") of the workbook window whose top-level window is <paramref name="xlMain"/>:
    /// XLMAIN > XLDESK > EXCEL7 (Excel 2013+ is SDI: one XLMAIN per workbook window). Returns zero when not found.
    /// </summary>
    public static IntPtr FindGridWindow(IntPtr xlMain)
    {
        if (xlMain == IntPtr.Zero) return IntPtr.Zero;
        var desk = FindWindowExW(xlMain, IntPtr.Zero, "XLDESK", null);
        if (desk == IntPtr.Zero) return IntPtr.Zero;

        var child = IntPtr.Zero;
        while ((child = FindWindowExW(desk, child, "EXCEL7", null)) != IntPtr.Zero)
        {
            if (IsWindowVisible(child)) return child;
        }

        return IntPtr.Zero;
    }

    /// <summary>
    /// Switches the calling thread to per-monitor-v2 DPI awareness for the lifetime of the scope so that every window
    /// rectangle and every created window works in physical pixels (like Excel's COM coordinates), then restores it.
    /// </summary>
    public readonly struct DpiScope : IDisposable
    {
        private readonly IntPtr _previous;

        private DpiScope(IntPtr previous) => _previous = previous;

        public static DpiScope PerMonitorV2()
        {
            try
            {
                return new DpiScope(SetThreadDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2));
            }
            catch (EntryPointNotFoundException)
            {
                // Windows older than 10 1703: nothing to switch.
                return new DpiScope(IntPtr.Zero);
            }
        }

        public void Dispose()
        {
            if (_previous == IntPtr.Zero) return;
            try
            {
                SetThreadDpiAwarenessContext(_previous);
            }
            catch (EntryPointNotFoundException)
            {
                // Cannot happen after a successful switch; nothing to restore otherwise.
            }
        }
    }
}
