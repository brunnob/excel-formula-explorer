using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;
using FormulaExplorer.AddIn.Diagnostics;
using FormulaExplorer.AddIn.Session;
using FormulaExplorer.Core.Syntax;

[assembly: InternalsVisibleTo("FormulaExplorer.AddIn.Tests")]
[assembly: InternalsVisibleTo("FormulaExplorer.UiSnapshot")]

namespace FormulaExplorer.AddIn.UI;

/// <summary>State of the view for tests and diagnostics.</summary>
/// <param name="ScreenBounds">Window bounds in physical screen pixels.</param>
/// <param name="PixelsPerDip">VisualTreeHelper.GetDpi(window).PixelsPerDip (1.5 at 150% scaling).</param>
public sealed record ViewSnapshot(
    string DisplayText,
    bool Interactive,
    int? HoverStart,
    int? HoverLength,
    bool HoverIsLink,
    bool TooltipVisible,
    string? TooltipText,
    int? GoldStart,
    int? GoldLength,
    bool BackEnabled,
    bool IsLoading,
    double PixelsPerDip,
    Int32Rect ScreenBounds);

/// <summary>
/// The non-modal "Formula Explorer" window (code-only WPF, no XAML).
/// Owned by Excel's window (never in the taskbar), resizable, only the X in the title bar, remembers its bounds.
/// </summary>
public sealed class ExplorerWindow : Window, IExplorerView
{
    /// <summary>First-open placement: right edge this far inside the owner's right edge, top this far below its top.</summary>
    private const double OwnerInsetRight = 24;
    private const double OwnerInsetTop = 100;

    private readonly ExplorerPanel _panel;
    private readonly string _settingsPath;
    private IntPtr _owner;
    private HwndSource? _hwndSource;
    private WindowKeyHook? _keyHook;
    private bool _placeInOwner;
    private bool _closeForReal;
    private bool _inClosing;
    private bool _closed;
    private bool _destroyNotified;
    private bool _inSizeMove;
    private int _lastKey;
    private uint _lastKeyTime;

    /// <param name="ownerHwnd">Excel workbook window (XLMAIN) that owns this window.</param>
    /// <param name="infoProvider">Asked about hovered nodes (link? tooltip text?).</param>
    public ExplorerWindow(IntPtr ownerHwnd, INodeInfoProvider infoProvider)
        : this(ownerHwnd, infoProvider, WindowSettings.DefaultPath)
    {
    }

    internal ExplorerWindow(IntPtr ownerHwnd, INodeInfoProvider infoProvider, string settingsPath)
    {
        _owner = ownerHwnd;
        _settingsPath = settingsPath;

        Title = "Formula Explorer";
        ShowInTaskbar = false;
        WindowStyle = WindowStyle.SingleBorderWindow;
        ResizeMode = ResizeMode.CanResize;
        Background = new SolidColorBrush(Color.FromRgb(0xF0, 0xF0, 0xF0));
        FontFamily = new FontFamily("Segoe UI");
        FontSize = 12;
        UseLayoutRounding = true;
        SnapsToDevicePixels = true;
        TextOptions.SetTextFormattingMode(this, TextFormattingMode.Display);
        if (ownerHwnd != IntPtr.Zero) new WindowInteropHelper(this).Owner = ownerHwnd;

        _panel = new ExplorerPanel(infoProvider);
        Content = _panel;
        _panel.BackButton.Click += (_, _) => RaiseKeepingActive("ExplorerWindow.Back", () => BackRequested?.Invoke(this, EventArgs.Empty));
        _panel.CloseButton.Click += (_, _) => Raise("ExplorerWindow.Close", () => CloseRequested?.Invoke(this, EventArgs.Empty));
        _panel.View.NodeClicked += (_, node) => RaiseKeepingActive("ExplorerWindow.Click", () => NodeClicked?.Invoke(this, node));
        _panel.View.NodeDoubleClicked += (_, node) => RaiseKeepingActive("ExplorerWindow.DoubleClick", () => NodeDoubleClicked?.Invoke(this, node));

        ApplyInitialBounds();
        Loaded += (_, _) => Raise("ExplorerWindow.Loaded", OnFirstLoaded);
    }

    public event EventHandler<SyntaxNode>? NodeClicked;
    public event EventHandler<SyntaxNode>? NodeDoubleClicked;
    public event EventHandler? BackRequested;
    public event EventHandler? CloseRequested;

    /// <summary>HWND of this window (valid after Show).</summary>
    public IntPtr Handle => new WindowInteropHelper(this).Handle;

    /// <summary>
    /// False once the window is gone, however that happened (closed by us, or destroyed together with its owner).
    /// A window that has not been shown yet is alive.
    /// </summary>
    public bool IsAlive
    {
        get
        {
            if (_closed) return false;
            IntPtr hwnd = Handle;
            return hwnd == IntPtr.Zero || NativeMethods.IsWindow(hwnd);
        }
    }

    internal ExplorerPanel Panel => _panel;

    // ----- Owner / z-order -----

    /// <summary>The Excel workbook window (XLMAIN) that currently owns this window.</summary>
    public IntPtr OwnerWindow => _owner;

    /// <summary>Re-owns the window to another Excel workbook window (Excel is SDI: one XLMAIN per workbook).</summary>
    public void SetOwnerWindow(IntPtr ownerHwnd)
    {
        try
        {
            if (ownerHwnd == IntPtr.Zero || ownerHwnd == _owner || _closed) return;

            IntPtr hwnd = Handle;
            if (hwnd == IntPtr.Zero)
            {
                new WindowInteropHelper(this).Owner = ownerHwnd; // not shown yet: WPF applies it at creation
                _owner = ownerHwnd;
                return;
            }
            if (NativeMethods.GetOwner(hwnd) == ownerHwnd)
            {
                _owner = ownerHwnd;
                return;
            }

            NativeMethods.SetOwner(hwnd, ownerHwnd);
            _owner = ownerHwnd;
            // Changing the owner alone does not re-sort the z-order: bring the window above its new owner.
            NativeMethods.SetWindowPos(hwnd, NativeMethods.HWND_TOP, 0, 0, 0, 0,
                NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOACTIVATE);
        }
        catch (Exception ex)
        {
            Log.Error("ExplorerWindow.SetOwnerWindow", ex);
        }
    }

    // ----- Closing -----

    /// <summary>Closes the window for real (saves position/size). Does not raise CloseRequested.</summary>
    public void CloseWindow()
    {
        try
        {
            if (_closed) return;
            _closeForReal = true;
            if (!_inClosing) Close(); // inside Closing the pending close completes by itself (see OnClosing)
        }
        catch (Exception ex)
        {
            Log.Error("ExplorerWindow.CloseWindow", ex);
        }
    }

    protected override void OnClosing(CancelEventArgs e)
    {
        base.OnClosing(e);
        try
        {
            if (_closeForReal)
            {
                SaveBounds();
                return;
            }

            // X or Alt+F4: only ask the controller; it closes us through CloseWindow().
            e.Cancel = true;
            _inClosing = true;
            try
            {
                CloseRequested?.Invoke(this, EventArgs.Empty);
            }
            finally
            {
                _inClosing = false;
            }
            if (_closeForReal)
            {
                e.Cancel = false;
                SaveBounds();
            }
        }
        catch (Exception ex)
        {
            Log.Error("ExplorerWindow.OnClosing", ex);
        }
    }

    /// <summary>
    /// The window is being destroyed without having been closed through <see cref="CloseWindow"/> or the X: Win32
    /// destroys owned windows together with their owner (the Excel workbook window, e.g. when the workbook closes or
    /// Excel quits), and WPF then sees WM_DESTROY without WM_CLOSE, so neither Closing nor the controller hear of it.
    /// Called while the HWND still exists (position can be read) or, as a fallback, from Closed. Saves the bounds and
    /// tells the controller (posted, so it runs after the destruction) to tear the session down.
    /// </summary>
    private void OnDestroyedWithoutClose(bool hwndStillExists)
    {
        if (_destroyNotified || _closeForReal) return;
        _destroyNotified = true;
        ReleaseKeyHook();
        try
        {
            if (hwndStillExists) SaveBounds();
        }
        catch (Exception ex)
        {
            Log.Error("ExplorerWindow.SaveBoundsOnDestroy", ex);
        }
        _closeForReal = true; // nothing can be vetoed any more; also stops Reactivate

        Dispatcher.BeginInvoke(DispatcherPriority.Normal,
            new Action(() => Raise("ExplorerWindow.Destroyed", () => CloseRequested?.Invoke(this, EventArgs.Empty))));
    }

    protected override void OnClosed(EventArgs e)
    {
        _closed = true;
        OnDestroyedWithoutClose(hwndStillExists: false); // safety net when the WM_DESTROY hook did not see it
        try
        {
            ReleaseKeyHook();
            _panel.View.Shutdown();
            _hwndSource?.RemoveHook(WndProc);
            _hwndSource = null;
        }
        catch (Exception ex)
        {
            Log.Error("ExplorerWindow.OnClosed", ex);
        }
        base.OnClosed(e);
    }

    private void ReleaseKeyHook()
    {
        _keyHook?.Dispose();
        _keyHook = null;
    }

    private void SaveBounds()
    {
        var bounds = RestoreBounds;
        if (bounds.IsEmpty && !TryReadPlacement(out bounds)) return;
        WindowSettings.Save(new WindowBounds(bounds.Left, bounds.Top, bounds.Width, bounds.Height), _settingsPath);
    }

    /// <summary>Normal (restored) bounds in DIPs from the live HWND, for when WPF no longer knows them.</summary>
    private bool TryReadPlacement(out Rect bounds)
    {
        bounds = Rect.Empty;
        IntPtr hwnd = Handle;
        if (hwnd == IntPtr.Zero || NativeMethods.GetNormalPlacement(hwnd) is not { } placement) return false;

        double scale = NativeMethods.GetDpiScale(hwnd);
        bounds = new Rect(placement.Left / scale, placement.Top / scale, placement.Width / scale, placement.Height / scale);
        return true;
    }

    // ----- Initial bounds, chrome -----

    /// <summary>
    /// Restores the remembered bounds when they are still on a monitor; otherwise the window sizes itself to the
    /// default 414×119 client (fixed box) and is placed at the top right of its owner when it opens.
    /// </summary>
    private void ApplyInitialBounds()
    {
        WindowBounds? saved = WindowSettings.Load(_settingsPath);
        if (saved is { } bounds && WindowSettings.IsOnScreen(bounds, NativeMethods.GetDpiScale(_owner), NativeMethods.GetWorkAreas()))
        {
            WindowStartupLocation = WindowStartupLocation.Manual;
            SizeToContent = SizeToContent.Manual;
            Left = bounds.Left;
            Top = bounds.Top;
            Width = bounds.Width;
            Height = bounds.Height;
            _panel.ReleaseBoxSize();
            return;
        }

        WindowStartupLocation = WindowStartupLocation.Manual;
        SizeToContent = SizeToContent.WidthAndHeight;
        _placeInOwner = true;
    }

    private static void RemoveIcon(IntPtr hwnd)
    {
        NativeMethods.SendMessage(hwnd, NativeMethods.WM_SETICON, (IntPtr)NativeMethods.ICON_SMALL, IntPtr.Zero);
        NativeMethods.SendMessage(hwnd, NativeMethods.WM_SETICON, (IntPtr)NativeMethods.ICON_BIG, IntPtr.Zero);
    }

    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);
        try
        {
            IntPtr hwnd = Handle;
            // Title bar with only the X: no minimize/maximize boxes and no icon.
            long style = NativeMethods.GetWindowLong(hwnd, NativeMethods.GWL_STYLE);
            NativeMethods.SetWindowLong(hwnd, NativeMethods.GWL_STYLE,
                style & ~(NativeMethods.WS_MINIMIZEBOX | NativeMethods.WS_MAXIMIZEBOX));
            long exStyle = NativeMethods.GetWindowLong(hwnd, NativeMethods.GWL_EXSTYLE);
            NativeMethods.SetWindowLong(hwnd, NativeMethods.GWL_EXSTYLE, exStyle | NativeMethods.WS_EX_DLGMODALFRAME);
            // Inside Excel WPF gives the window EXCEL.EXE's icon; the original title bar had none.
            RemoveIcon(hwnd);
            NativeMethods.SetWindowPos(hwnd, IntPtr.Zero, 0, 0, 0, 0,
                NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOZORDER | NativeMethods.SWP_NOACTIVATE
                | NativeMethods.SWP_FRAMECHANGED);

            _hwndSource = HwndSource.FromHwnd(hwnd);
            _hwndSource?.AddHook(WndProc);
            // Excel's pump swallows real Esc/Enter keystrokes (accelerators) before they reach us: take them in GetMessage.
            _keyHook = new WindowKeyHook(() => Handle, () => _inSizeMove, (key, time) => HandleKeyDown(key, false, time), Dispatcher);

            if (_placeInOwner) PlaceAtTopRightOfOwner(hwnd);
        }
        catch (Exception ex)
        {
            Log.Error("ExplorerWindow.OnSourceInitialized", ex);
        }
    }

    /// <summary>First open: over the ribbon area, top right inside the owner window (physical pixels).</summary>
    private void PlaceAtTopRightOfOwner(IntPtr hwnd)
    {
        var own = NativeMethods.GetWindowBounds(hwnd);
        if (own is not { } size) return;

        double scale = NativeMethods.GetDpiScale(_owner != IntPtr.Zero ? _owner : hwnd);
        int x, y;
        if (NativeMethods.GetWindowBounds(_owner) is { } owner)
        {
            x = owner.Right - (int)Math.Round(OwnerInsetRight * scale) - size.Width;
            y = owner.Top + (int)Math.Round(OwnerInsetTop * scale);
        }
        else
        {
            var area = NativeMethods.GetWorkAreas().FirstOrDefault();
            x = area.Right - (int)Math.Round(OwnerInsetRight * scale) - size.Width;
            y = area.Top + (int)Math.Round(OwnerInsetTop * scale);
        }
        NativeMethods.SetWindowPos(hwnd, IntPtr.Zero, x, y, 0, 0,
            NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOZORDER | NativeMethods.SWP_NOACTIVATE);
    }

    /// <summary>
    /// After the first layout: derive the chrome size, set the minimum size (a 200×20 formula box) and, when the window
    /// sized itself to the fixed box, let the box stretch from now on.
    /// </summary>
    private void OnFirstLoaded()
    {
        RemoveIcon(Handle); // WPF may (re)apply the default icon after SourceInitialized
        double chromeWidth = ActualWidth - _panel.ActualWidth;
        double chromeHeight = ActualHeight - _panel.ActualHeight;
        if (SizeToContent != SizeToContent.Manual)
        {
            double width = ActualWidth, height = ActualHeight;
            SizeToContent = SizeToContent.Manual;
            Width = width;
            Height = height;
            _panel.ReleaseBoxSize();
        }
        MinWidth = chromeWidth + ExplorerPanel.MinClientSize.Width;
        MinHeight = chromeHeight + ExplorerPanel.MinClientSize.Height;
    }

    // ----- IExplorerView -----

    public void ShowLoading() => Raise("ExplorerWindow.ShowLoading", _panel.ShowLoading);

    public void ShowFormula(FormulaDisplay display) => Raise("ExplorerWindow.ShowFormula", () => _panel.ShowFormula(display));

    public void SetGold(SyntaxNode? node) => Raise("ExplorerWindow.SetGold", () => _panel.View.SetGold(node));

    public void SetBackEnabled(bool enabled) => Raise("ExplorerWindow.SetBackEnabled", () => _panel.BackButton.IsEnabled = enabled);

    public void ShowMessage(string message) => Raise("ExplorerWindow.ShowMessage", () =>
    {
        if (IsLoaded) MessageBox.Show(this, message, Title, MessageBoxButton.OK, MessageBoxImage.Information);
        else MessageBox.Show(message, Title, MessageBoxButton.OK, MessageBoxImage.Information);
    });

    public void RefreshHover() => Raise("ExplorerWindow.RefreshHover", _panel.View.RefreshHover);

    // ----- Link clicks keep the window active -----

    /// <summary>
    /// Runs a click / Back handler (both navigate). Excel's Goto may activate the workbook window; the original stayed active (which also lets
    /// the second click of a double click reach us), so the window re-activates itself afterwards when it was active.
    /// </summary>
    private void RaiseKeepingActive(string context, Action raise)
    {
        bool wasActive = IsActive;
        try
        {
            Raise(context, raise);
        }
        finally
        {
            if (wasActive) Dispatcher.BeginInvoke(DispatcherPriority.Input, new Action(Reactivate));
        }
    }

    private void Reactivate()
    {
        try
        {
            if (_closed || _closeForReal || IsActive || !IsVisible) return;
            Activate();
            if (!IsActive) NativeMethods.SetForegroundWindow(Handle);
        }
        catch (Exception ex)
        {
            Log.Error("ExplorerWindow.Reactivate", ex);
        }
    }

    // ----- Keyboard -----

    protected override void OnPreviewKeyDown(KeyEventArgs e)
    {
        base.OnPreviewKeyDown(e);
        try
        {
            int vk = e.Key switch
            {
                Key.Escape => NativeMethods.VK_ESCAPE,
                Key.Enter => NativeMethods.VK_RETURN,
                _ => 0,
            };
            if (vk != 0 && HandleKeyDown(vk, e.IsRepeat, unchecked((uint)e.Timestamp))) e.Handled = true;
        }
        catch (Exception ex)
        {
            Log.Error("ExplorerWindow.PreviewKeyDown", ex);
        }
    }

    /// <summary>
    /// Excel's message loop may not route keys to WPF, so WM_KEYDOWN also arrives here through the HwndSource hook (when the
    /// <see cref="WindowKeyHook"/> is not active); WM_ENTERSIZEMOVE/WM_EXITSIZEMOVE tell the key hook to stay out of the
    /// window's own move/size loop.
    /// </summary>
    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        try
        {
            switch (msg)
            {
                case NativeMethods.WM_DESTROY:
                    _inSizeMove = false;
                    OnDestroyedWithoutClose(hwndStillExists: true);
                    return IntPtr.Zero;
                case NativeMethods.WM_ENTERSIZEMOVE:
                    _inSizeMove = true;
                    break;
                case NativeMethods.WM_EXITSIZEMOVE:
                    _inSizeMove = false;
                    break;
                case NativeMethods.WM_KEYDOWN:
                    int vk = (int)(wParam.ToInt64() & 0xFFFF);
                    if (HandleKeyDown(vk, WindowKeyHook.IsAutoRepeat(lParam.ToInt64()), unchecked((uint)NativeMethods.GetMessageTime())))
                        handled = true;
                    break;
            }
        }
        catch (Exception ex)
        {
            Log.Error("ExplorerWindow.WndProc", ex);
        }
        return IntPtr.Zero;
    }

    /// <summary>
    /// Esc closes; Enter presses Back when it has the keyboard focus, otherwise Close. Returns true when the key was ours.
    /// An auto-repeat (<paramref name="isRepeat"/>, from the key-down message itself) does nothing, and neither does the same
    /// press (same key, same message <paramref name="time"/>) arriving through a second input path. Nothing waits for a
    /// key-up: Excel's pump can swallow it just as it swallows the key-down, which would leave the next press ignored.
    /// </summary>
    internal bool HandleKeyDown(int virtualKey, bool isRepeat = false, uint time = 0)
    {
        if (virtualKey is not (NativeMethods.VK_ESCAPE or NativeMethods.VK_RETURN)) return false;
        if (isRepeat) return true;
        if (time != 0 && virtualKey == _lastKey && time == _lastKeyTime) return true;
        _lastKey = virtualKey;
        _lastKeyTime = time;

        bool back = virtualKey == NativeMethods.VK_RETURN && EnterPressesBack(_panel.BackButton.IsEnabled, _panel.BackButton.IsKeyboardFocused);
        if (back) RaiseKeepingActive("ExplorerWindow.Enter", () => BackRequested?.Invoke(this, EventArgs.Empty));
        else Raise("ExplorerWindow.Key", () => CloseRequested?.Invoke(this, EventArgs.Empty));
        return true;
    }

    /// <summary>Enter acts on the focused button: Back only when it is enabled and focused (Close is the default otherwise).</summary>
    internal static bool EnterPressesBack(bool backEnabled, bool backFocused) => backEnabled && backFocused;

    // ----- Test / diagnostic surface -----

    /// <summary>As if the mouse moved over char <paramref name="charIndex"/> (null = mouse left). Shows the tooltip immediately.</summary>
    public void SimulateHover(int? charIndex)
    {
        Raise("ExplorerWindow.SimulateHover", () =>
        {
            if (charIndex is not { } index)
            {
                _panel.View.SimulateMove(null);
                return;
            }
            if (CharCenter(index) is { } point) _panel.View.SimulateMove(point);
        });
    }

    /// <summary>As if the user clicked char <paramref name="charIndex"/> once.</summary>
    public void SimulateClick(int charIndex)
    {
        Raise("ExplorerWindow.SimulateClick", () =>
        {
            if (CharCenter(charIndex) is { } point) _panel.View.SimulatePress(point, 1);
        });
    }

    /// <summary>As if the user double-clicked char <paramref name="charIndex"/> (raises NodeClicked then NodeDoubleClicked, like a real double click).</summary>
    public void SimulateDoubleClick(int charIndex)
    {
        Raise("ExplorerWindow.SimulateDoubleClick", () =>
        {
            if (CharCenter(charIndex) is not { } point) return;
            _panel.View.SimulatePress(point, 1);
            _panel.View.SimulatePress(point, 2);
        });
    }

    /// <summary>Center of char <paramref name="charIndex"/> in physical screen pixels, or null if not laid out/visible.</summary>
    public System.Drawing.Point? GetCharScreenPoint(int charIndex)
    {
        try
        {
            if (_panel.View.GetCharCenter(charIndex) is not { } center) return null;
            var screen = _panel.View.PointToScreen(center); // device pixels
            return new System.Drawing.Point((int)Math.Round(screen.X), (int)Math.Round(screen.Y));
        }
        catch (InvalidOperationException)
        {
            return null; // not attached to a window
        }
    }

    public ViewSnapshot GetSnapshot()
    {
        var view = _panel.View;
        var hover = view.HoverNode;
        var gold = view.GoldNode;
        var bounds = NativeMethods.GetWindowBounds(Handle);
        return new ViewSnapshot(
            view.DisplayText,
            view.Interactive,
            hover?.Start,
            hover?.Length,
            view.HoverIsLink,
            view.TooltipVisible,
            view.TooltipText,
            gold?.Start,
            gold?.Length,
            _panel.BackButton.IsEnabled,
            view.IsLoading,
            VisualTreeHelper.GetDpi(this).PixelsPerDip,
            bounds is { } b ? new Int32Rect(b.Left, b.Top, b.Width, b.Height) : default);
    }

    /// <summary>Char center in view DIPs after scrolling it into view.</summary>
    private Point? CharCenter(int charIndex)
    {
        _panel.EnsureCharVisible(charIndex);
        return _panel.View.GetCharCenter(charIndex);
    }

    /// <summary>Runs an entry point; nothing may escape into Excel.</summary>
    private static void Raise(string context, Action action)
    {
        try
        {
            action();
        }
        catch (Exception ex)
        {
            Log.Error(context, ex);
        }
    }
}
