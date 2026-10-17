using System.Runtime.InteropServices;

namespace FormulaExplorer.AddIn.UI;

/// <summary>A rectangle in physical screen pixels (left/top inclusive, right/bottom exclusive).</summary>
internal readonly record struct PixelRect(int Left, int Top, int Right, int Bottom)
{
    public int Width => Right - Left;

    public int Height => Bottom - Top;
}

/// <summary>The few Win32 calls the explorer window needs. Everything is on the UI thread.</summary>
internal static class NativeMethods
{
    internal const int GWL_STYLE = -16;
    internal const int GWL_EXSTYLE = -20;
    internal const int GWLP_HWNDPARENT = -8;

    internal const long WS_MAXIMIZEBOX = 0x00010000;
    internal const long WS_MINIMIZEBOX = 0x00020000;
    internal const long WS_EX_DLGMODALFRAME = 0x00000001;
    internal const long WS_EX_TRANSPARENT = 0x00000020;
    internal const long WS_EX_NOACTIVATE = 0x08000000;

    internal const uint SWP_NOSIZE = 0x0001;
    internal const uint SWP_NOMOVE = 0x0002;
    internal const uint SWP_NOZORDER = 0x0004;
    internal const uint SWP_NOACTIVATE = 0x0010;
    internal const uint SWP_FRAMECHANGED = 0x0020;

    internal const int WM_DESTROY = 0x0002;
    internal const int WM_KEYDOWN = 0x0100;
    internal const int WM_KEYUP = 0x0101;
    internal const int WM_ENTERSIZEMOVE = 0x0231;
    internal const int WM_EXITSIZEMOVE = 0x0232;

    internal const int VK_RETURN = 0x0D;
    internal const int VK_ESCAPE = 0x1B;

    private const int SM_CXDOUBLECLK = 36;
    private const int SM_CYDOUBLECLK = 37;

    internal static readonly IntPtr HWND_TOP = IntPtr.Zero;

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left, Top, Right, Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MONITORINFO
    {
        public int Size;
        public RECT Monitor;
        public RECT Work;
        public uint Flags;
    }

    private delegate bool MonitorEnumProc(IntPtr monitor, IntPtr hdc, IntPtr rect, IntPtr data);

    internal const int WM_SETICON = 0x0080;
    internal const int ICON_SMALL = 0;
    internal const int ICON_BIG = 1;

    [DllImport("user32.dll", EntryPoint = "SendMessageW")]
    internal static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")]
    private static extern IntPtr GetWindowLongPtr64(IntPtr hwnd, int index);

    [DllImport("user32.dll", EntryPoint = "GetWindowLongW")]
    private static extern int GetWindowLong32(IntPtr hwnd, int index);

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW")]
    private static extern IntPtr SetWindowLongPtr64(IntPtr hwnd, int index, IntPtr value);

    [DllImport("user32.dll", EntryPoint = "SetWindowLongW")]
    private static extern int SetWindowLong32(IntPtr hwnd, int index, int value);

    [StructLayout(LayoutKind.Sequential)]
    private struct WINDOWPLACEMENT
    {
        public int Length;
        public int Flags;
        public int ShowCmd;
        public int MinX, MinY, MaxX, MaxY;
        public RECT NormalPosition;
    }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetWindowPlacement(IntPtr hwnd, ref WINDOWPLACEMENT placement);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool IsWindow(IntPtr hwnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool IsWindowEnabled(IntPtr hwnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool IsWindowVisible(IntPtr hwnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool IsIconic(IntPtr hwnd);

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool SetWindowPos(IntPtr hwnd, IntPtr insertAfter, int x, int y, int cx, int cy, uint flags);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool SetForegroundWindow(IntPtr hwnd);

    [DllImport("user32.dll")]
    internal static extern uint GetDoubleClickTime();

    /// <summary>Time (ms, same clock as MSG.time) of the message the calling thread last retrieved.</summary>
    [DllImport("user32.dll")]
    internal static extern int GetMessageTime();

    [DllImport("user32.dll")]
    private static extern int GetSystemMetrics(int index);

    [DllImport("user32.dll")]
    private static extern uint GetDpiForWindow(IntPtr hwnd);

    [DllImport("user32.dll")]
    private static extern bool EnumDisplayMonitors(IntPtr hdc, IntPtr clip, MonitorEnumProc callback, IntPtr data);

    [DllImport("user32.dll", EntryPoint = "GetMonitorInfoW")]
    private static extern bool GetMonitorInfo(IntPtr monitor, ref MONITORINFO info);

    internal static long GetWindowLong(IntPtr hwnd, int index) =>
        IntPtr.Size == 8 ? GetWindowLongPtr64(hwnd, index).ToInt64() : GetWindowLong32(hwnd, index);

    internal static void SetWindowLong(IntPtr hwnd, int index, long value)
    {
        if (IntPtr.Size == 8) SetWindowLongPtr64(hwnd, index, new IntPtr(value));
        else SetWindowLong32(hwnd, index, unchecked((int)value));
    }

    /// <summary>Current owner (GWLP_HWNDPARENT) of a top-level window.</summary>
    internal static IntPtr GetOwner(IntPtr hwnd) => new(GetWindowLong(hwnd, GWLP_HWNDPARENT));

    internal static void SetOwner(IntPtr hwnd, IntPtr owner) => SetWindowLong(hwnd, GWLP_HWNDPARENT, owner.ToInt64());

    /// <summary>Window rectangle in physical screen pixels, or null when the handle is not valid.</summary>
    internal static PixelRect? GetWindowBounds(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero || !GetWindowRect(hwnd, out var r)) return null;
        return new PixelRect(r.Left, r.Top, r.Right, r.Bottom);
    }

    /// <summary>Restored (non-maximized/minimized) window rectangle in physical pixels, or null when unavailable.</summary>
    internal static PixelRect? GetNormalPlacement(IntPtr hwnd)
    {
        var placement = new WINDOWPLACEMENT { Length = Marshal.SizeOf<WINDOWPLACEMENT>() };
        if (hwnd == IntPtr.Zero || !GetWindowPlacement(hwnd, ref placement)) return null;
        var r = placement.NormalPosition; // workspace coordinates; equal to screen coordinates for non-toolwindow top-levels
        return new PixelRect(r.Left, r.Top, r.Right, r.Bottom);
    }

    /// <summary>DPI scale (1.0 = 96 DPI) of a window; 1.0 when unknown.</summary>
    internal static double GetDpiScale(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero) return 1.0;
        uint dpi = GetDpiForWindow(hwnd);
        return dpi == 0 ? 1.0 : dpi / 96.0;
    }

    /// <summary>Work areas of all monitors in physical pixels.</summary>
    internal static IReadOnlyList<PixelRect> GetWorkAreas()
    {
        var areas = new List<PixelRect>();
        EnumDisplayMonitors(IntPtr.Zero, IntPtr.Zero, (monitor, _, _, _) =>
        {
            var info = new MONITORINFO { Size = Marshal.SizeOf<MONITORINFO>() };
            if (GetMonitorInfo(monitor, ref info))
                areas.Add(new PixelRect(info.Work.Left, info.Work.Top, info.Work.Right, info.Work.Bottom));
            return true;
        }, IntPtr.Zero);
        return areas;
    }

    /// <summary>System double-click rectangle (width, height) in pixels.</summary>
    internal static (int Width, int Height) GetDoubleClickSize() =>
        (Math.Max(1, GetSystemMetrics(SM_CXDOUBLECLK)), Math.Max(1, GetSystemMetrics(SM_CYDOUBLECLK)));
}
