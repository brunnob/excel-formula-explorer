using System.Runtime.InteropServices;
using FormulaExplorer.AddIn.Diagnostics;
using static FormulaExplorer.AddIn.Overlay.NativeMethods;

namespace FormulaExplorer.AddIn.Overlay;

/// <summary>
/// The Win32 window that carries the precedent borders: a popup owned by a workbook's XLMAIN (kept directly above it in z-order), layered with a color key
/// (every pixel painted in <see cref="KeyColorRgb"/> is see-through), click-through (WS_EX_TRANSPARENT plus
/// WM_NCHITTEST = HTTRANSPARENT), never activating and hidden from Alt-Tab (WS_EX_TOOLWINDOW).
/// WM_PAINT fills the key color and then the border rectangles with <c>FillRect</c> (pixel-exact, no antialiasing).
/// <para>
/// The window is created under the per-monitor-v2 DPI awareness context so its size and position are in physical pixels,
/// like Excel's COM coordinates. All members must be called on the thread that created it (Excel's main thread).
/// </para>
/// </summary>
internal sealed class OverlayWindow : IDisposable
{
    /// <summary>Color key as 0xRRGGBB; no border may use it (see <see cref="SafeRgb"/>).</summary>
    public const uint KeyColorRgb = 0x010203;

    private static int _classCounter;

    private readonly WndProc _wndProc;               // kept in a field: the native class holds a pointer to it
    private readonly IntPtr _className;              // unmanaged copy of the class name (WNDCLASSEX holds a pointer)
    private readonly IntPtr _hInstance;
    private readonly IntPtr _keyBrush;
    private readonly Dictionary<uint, IntPtr> _brushes = new();

    private IntPtr _hwnd;
    private IntPtr _owner;
    private PixelRect _bounds;
    private PaintItem[] _items = [];
    private bool _visible;
    private bool _dirty = true;
    private bool _disposed;

    private readonly record struct PaintItem(PixelRect Local, uint ColorRef);

    /// <summary>Creates the (hidden) window owned by <paramref name="owner"/>. Throws when Win32 refuses.</summary>
    public OverlayWindow(IntPtr owner)
    {
        _owner = owner;
        _wndProc = HandleMessage;
        _hInstance = GetModuleHandleW(null);
        _className = Marshal.StringToHGlobalUni($"FormulaExplorerOverlay.{Environment.ProcessId}.{Interlocked.Increment(ref _classCounter)}");
        _keyBrush = CreateSolidBrush(ToColorRef(KeyColorRgb));

        var registered = false;
        try
        {
            using (DpiScope.PerMonitorV2())
            {
                var wc = new WNDCLASSEX
                {
                    cbSize = (uint)Marshal.SizeOf<WNDCLASSEX>(),
                    lpfnWndProc = Marshal.GetFunctionPointerForDelegate(_wndProc),
                    hInstance = _hInstance,
                    lpszClassName = _className,
                };
                if (RegisterClassExW(ref wc) == 0)
                    throw new InvalidOperationException($"RegisterClassEx failed (error {Marshal.GetLastWin32Error()}).");
                registered = true;

                _hwnd = CreateWindowExW(
                    WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW,
                    _className, "Formula Explorer precedents", WS_POPUP,
                    0, 0, 1, 1, owner, IntPtr.Zero, _hInstance, IntPtr.Zero);
                if (_hwnd == IntPtr.Zero)
                    throw new InvalidOperationException($"CreateWindowEx failed (error {Marshal.GetLastWin32Error()}).");

                if (!SetLayeredWindowAttributes(_hwnd, ToColorRef(KeyColorRgb), 0, LWA_COLORKEY))
                    throw new InvalidOperationException($"SetLayeredWindowAttributes failed (error {Marshal.GetLastWin32Error()}).");
            }
        }
        catch
        {
            ReleaseNative(registered);
            throw;
        }
    }

    public IntPtr Handle => _hwnd;

    public IntPtr Owner => _owner;

    public bool IsVisible => _visible;

    /// <summary>
    /// False once the native window is gone: destroyed together with its owner (closing a workbook window in SDI Excel
    /// destroys its XLMAIN and every window it owns) or disposed. A dead window cannot be reused; create a new one.
    /// </summary>
    public bool IsAlive => _hwnd != IntPtr.Zero && IsWindow(_hwnd);

    /// <summary>The window's current DPI (the DPI of the monitor it is on; physical, because the window is per-monitor-v2).</summary>
    public int Dpi
    {
        get
        {
            var dpi = (int)GetDpiForWindow(_hwnd);
            return dpi > 0 ? dpi : 96;
        }
    }

    /// <summary>A color that cannot be mistaken for the color key.</summary>
    public static uint SafeRgb(uint rgb)
    {
        rgb &= 0xFFFFFF;
        return rgb == KeyColorRgb ? KeyColorRgb + 1 : rgb;
    }

    /// <summary>
    /// Makes <paramref name="owner"/> the owner window. Changing the owner alone does not re-sort the z-order, so the
    /// caller follows up with <see cref="Place"/> or <see cref="EnsureAboveOwner"/>.
    /// </summary>
    public void SetOwner(IntPtr owner)
    {
        if (owner == _owner || !IsAlive) return;
        // GWLP_HWNDPARENT on a top-level window sets its owner (SetParent would make it a child).
        SetWindowLongPtr(_hwnd, GWLP_HWNDPARENT, owner);
        _owner = owner;
    }

    /// <summary>
    /// Positions the window over <paramref name="screenBounds"/> (the grid window rect) and puts it directly below
    /// <paramref name="insertAfter"/> in z-order (normally <see cref="AboveOwnerAnchor"/>; <see cref="IntPtr.Zero"/> = top;
    /// the window itself = leave the z-order alone). Never activates. Works while hidden.
    /// </summary>
    public void Place(PixelRect screenBounds, IntPtr insertAfter)
    {
        if (screenBounds.IsEmpty) return;
        if (!IsAlive) throw new InvalidOperationException("The overlay window no longer exists.");
        var flags = SWP_NOACTIVATE;
        if (screenBounds == _bounds) flags |= SWP_NOMOVE | SWP_NOSIZE;
        if (insertAfter == _hwnd) flags |= SWP_NOZORDER;
        var ok = SetWindowPos(_hwnd, insertAfter, screenBounds.Left, screenBounds.Top, screenBounds.Width, screenBounds.Height, flags);
        // The explorer window may be gone or unrelated (e.g. mid-close): fall back to the top of the owner's windows.
        if (!ok && insertAfter != HWND_TOP)
            ok = SetWindowPos(_hwnd, HWND_TOP, screenBounds.Left, screenBounds.Top, screenBounds.Width, screenBounds.Height, flags);
        if (!ok)
            throw new InvalidOperationException($"SetWindowPos failed (error {Marshal.GetLastWin32Error()}).");
        if (screenBounds.Width != _bounds.Width || screenBounds.Height != _bounds.Height) _dirty = true;
        _bounds = screenBounds;
    }

    /// <summary>
    /// Sets what to paint. <paramref name="screenFills"/> are in screen pixels and must already be clipped; call after
    /// <see cref="Place"/> (they are stored relative to the window origin). Repaints only when something changed.
    /// </summary>
    public void SetContent(IReadOnlyList<ColoredRect> screenFills)
    {
        var items = new PaintItem[screenFills.Count];
        for (var i = 0; i < items.Length; i++)
        {
            var f = screenFills[i];
            items[i] = new PaintItem(f.Rect.Offset(-_bounds.Left, -_bounds.Top), ToColorRef(SafeRgb(f.Rgb)));
        }

        if (!items.AsSpan().SequenceEqual(_items)) _dirty = true;
        _items = items;
    }

    /// <summary>Shows the window without activating it (when hidden) and paints synchronously when the content changed.</summary>
    public void Show()
    {
        if (!IsAlive) return;
        if (!_visible)
        {
            ShowWindow(_hwnd, SW_SHOWNOACTIVATE);
            _visible = true;
            _dirty = true;
        }

        if (!_dirty) return;
        _dirty = false;
        // Paint right away: a freshly shown layered window must not get a frame before its first WM_PAINT.
        InvalidateRect(_hwnd, IntPtr.Zero, false);
        UpdateWindow(_hwnd);
    }

    public void Hide()
    {
        if (!_visible || _hwnd == IntPtr.Zero) return;
        ShowWindow(_hwnd, SW_HIDE);
        _visible = false;
    }

    /// <summary>
    /// The window to pass as <c>insertAfter</c> so this window sits directly above its owner (XLMAIN): the window that is
    /// currently just above the owner (this window itself when it already is; <see cref="IntPtr.Zero"/> when the owner is topmost).
    /// Owned windows always stay above their owner, so the explorer (same owner) and every other Excel popup (Find &amp;
    /// Replace, floating panes, dialogs) end up above the borders, as with Range Finder.
    /// </summary>
    public IntPtr AboveOwnerAnchor() =>
        _owner != IntPtr.Zero && IsWindow(_owner) ? GetWindow(_owner, GW_HWNDPREV) : HWND_TOP;

    /// <summary>Keeps the window directly above its owner (see <see cref="AboveOwnerAnchor"/>); a no-op while hidden or dead.</summary>
    public void EnsureAboveOwner()
    {
        if (!_visible || !IsAlive || _owner == IntPtr.Zero || !IsWindow(_owner)) return;
        var anchor = AboveOwnerAnchor();
        if (anchor == _hwnd) return;
        SetWindowPos(_hwnd, anchor, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        ReleaseNative(classRegistered: true);
    }

    private void ReleaseNative(bool classRegistered)
    {
        try
        {
            if (_hwnd != IntPtr.Zero)
            {
                if (!DestroyWindow(_hwnd))
                    Log.Info($"Overlay: DestroyWindow failed (error {Marshal.GetLastWin32Error()}).");
                _hwnd = IntPtr.Zero;
            }

            if (classRegistered && !UnregisterClassW(_className, _hInstance))
                Log.Info($"Overlay: UnregisterClass failed (error {Marshal.GetLastWin32Error()}).");
        }
        finally
        {
            _visible = false;
            foreach (var brush in _brushes.Values) DeleteObject(brush);
            _brushes.Clear();
            DeleteObject(_keyBrush);
            Marshal.FreeHGlobal(_className);
        }
    }

    private IntPtr HandleMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam)
    {
        try
        {
            switch (msg)
            {
                case WM_NCDESTROY:
                    // Destroyed by DestroyWindow or together with the owner: remember it, so the handle is never reused.
                    _hwnd = IntPtr.Zero;
                    _visible = false;
                    break;
                case WM_NCHITTEST:
                    return HTTRANSPARENT;
                case WM_MOUSEACTIVATE:
                    return MA_NOACTIVATE;
                case WM_ERASEBKGND:
                    return new IntPtr(1);          // WM_PAINT paints every pixel
                case WM_PAINT:
                    Paint(hWnd);
                    return IntPtr.Zero;
                case WM_PRINTCLIENT:
                    // PrintWindow / AnimateWindow style requests: draw the whole client area into the caller's DC.
                    var client = new RECT(new PixelRect(0, 0, _bounds.Width, _bounds.Height));
                    Draw(wParam, ref client);
                    return IntPtr.Zero;
            }
        }
        catch (Exception ex)
        {
            Log.Error("Overlay window procedure", ex);
        }

        return DefWindowProcW(hWnd, msg, wParam, lParam);
    }

    private void Paint(IntPtr hWnd)
    {
        var hdc = BeginPaint(hWnd, out var ps);
        try
        {
            Draw(hdc, ref ps.rcPaint);
        }
        finally
        {
            EndPaint(hWnd, ref ps);
        }
    }

    /// <summary>Fills <paramref name="clip"/> with the key color (= see-through) and the border rectangles on top, in list order.</summary>
    private void Draw(IntPtr hdc, ref RECT clip)
    {
        FillRect(hdc, ref clip, _keyBrush);
        foreach (var item in _items)
        {
            if (!_brushes.TryGetValue(item.ColorRef, out var brush))
                _brushes[item.ColorRef] = brush = CreateSolidBrush(item.ColorRef);
            var rect = new RECT(item.Local);
            FillRect(hdc, ref rect, brush);
        }
    }
}
