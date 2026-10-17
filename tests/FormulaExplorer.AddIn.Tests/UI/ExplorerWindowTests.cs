using System.IO;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;
using FormulaExplorer.AddIn.Session;
using FormulaExplorer.AddIn.UI;
using FormulaExplorer.Core.Syntax;

namespace FormulaExplorer.AddIn.Tests.UI;

/// <summary>
/// Window-level behavior. Most tests only create the HWND (<c>EnsureHandle</c>) and never show the window, so nothing
/// flashes on screen; one test shows it off-screen-sized to check the default metrics.
/// </summary>
[Collection(DesktopWindows.Name)]
public sealed class ExplorerWindowTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "fe-window-" + Guid.NewGuid().ToString("N"));

    private string SettingsPath => Path.Combine(_dir, "settings.json");

    public void Dispose()
    {
        if (Directory.Exists(_dir)) Directory.Delete(_dir, recursive: true);
    }

    private ExplorerWindow NewWindow(IntPtr owner = default, INodeInfoProvider? provider = null)
    {
        var window = new ExplorerWindow(owner, provider ?? FakeProvider.Standard(), SettingsPath);
        new WindowInteropHelper(window).EnsureHandle();
        return window;
    }

    private static IntPtr NewOwner()
    {
        var owner = new Window();
        return new WindowInteropHelper(owner).EnsureHandle();
    }

    // ----- keyboard -----

    [Fact]
    public void EscapeRequestsCloseOncePerKeyPress() => Sta.Run(() =>
    {
        var window = NewWindow();
        int closes = 0;
        window.CloseRequested += (_, _) => closes++;

        Assert.True(window.HandleKeyDown(NativeMethods.VK_ESCAPE, isRepeat: false, time: 1000));
        Assert.True(window.HandleKeyDown(NativeMethods.VK_ESCAPE, isRepeat: false, time: 1000)); // the same press seen by a second input path
        Assert.Equal(1, closes);

        Assert.True(window.HandleKeyDown(NativeMethods.VK_ESCAPE, isRepeat: true, time: 1500)); // auto-repeat: claimed, does nothing
        Assert.True(window.HandleKeyDown(NativeMethods.VK_ESCAPE, isRepeat: true, time: 1530));
        Assert.Equal(1, closes);
        window.CloseWindow();
    });

    [Fact]
    public void ASecondPressActsEvenIfTheKeyUpNeverArrived() => Sta.Run(() =>
    {
        // Excel's pump can swallow the key-up like the key-down: nothing may wait for it (Enter twice must act twice).
        var window = NewWindow();
        int closes = 0;
        window.CloseRequested += (_, _) => closes++;

        window.HandleKeyDown(NativeMethods.VK_RETURN, isRepeat: false, time: 1000);
        window.HandleKeyDown(NativeMethods.VK_RETURN, isRepeat: false, time: 1350);
        window.HandleKeyDown(NativeMethods.VK_ESCAPE, isRepeat: false, time: 1700);
        Assert.Equal(3, closes);
        window.CloseWindow();
    });

    [Fact]
    public void EnterPressesCloseWhenBackIsNotFocused() => Sta.Run(() =>
    {
        var window = NewWindow();
        int closes = 0, backs = 0;
        window.CloseRequested += (_, _) => closes++;
        window.BackRequested += (_, _) => backs++;

        window.SetBackEnabled(true); // enabled but without the keyboard focus (the window is not even active)
        Assert.True(window.HandleKeyDown(NativeMethods.VK_RETURN));
        Assert.Equal((1, 0), (closes, backs));
        window.CloseWindow();
    });

    [Theory]
    [InlineData(true, true, true)]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(false, false, false)]
    public void EnterPressesBackOnlyWhenItIsEnabledAndFocused(bool enabled, bool focused, bool expected) =>
        Assert.Equal(expected, ExplorerWindow.EnterPressesBack(enabled, focused));

    [Fact]
    public void OtherKeysAreNotHandled() => Sta.Run(() =>
    {
        var window = NewWindow();
        Assert.False(window.HandleKeyDown(0x41));
        Assert.False(window.HandleKeyDown(0x20));
        window.CloseWindow();
    });

    // ----- closing -----

    [Fact]
    public void TitleBarCloseOnlyRequestsAndTheControllerClosesForReal() => Sta.Run(() =>
    {
        var window = NewWindow();
        bool closed = false;
        window.Closed += (_, _) => closed = true;
        int requests = 0;
        window.CloseRequested += (_, _) => requests++;

        window.Close(); // the X button / Alt+F4
        Assert.Equal(1, requests);
        Assert.False(closed);

        window.CloseWindow(); // what the controller does in response
        Assert.True(closed);
    });

    [Fact]
    public void CloseWindowFromInsideTheClosingHandlerDoesNotReenter() => Sta.Run(() =>
    {
        var window = NewWindow();
        bool closed = false;
        window.Closed += (_, _) => closed = true;
        window.CloseRequested += (_, _) => window.CloseWindow(); // controller: CloseAll() → window.CloseWindow() inside Closing

        window.Close();
        Assert.True(closed);
    });

    [Fact]
    public void CloseWindowIsIdempotentAndRaisesNoCloseRequested() => Sta.Run(() =>
    {
        var window = NewWindow();
        int requests = 0;
        window.CloseRequested += (_, _) => requests++;

        window.CloseWindow();
        window.CloseWindow();
        Assert.Equal(0, requests);
    });

    [Fact]
    public void ExceptionsFromCloseRequestedHandlersDoNotEscape() => Sta.Run(() =>
    {
        var window = NewWindow();
        window.CloseRequested += (_, _) => throw new InvalidOperationException("boom");

        window.Close(); // must not throw; the window stays open
        Assert.True(window.HandleKeyDown(NativeMethods.VK_ESCAPE));
        window.CloseWindow();
    });

    // ----- owner -----

    [Fact]
    public void SetOwnerWindowIsANoOpForTheCurrentOwnerAndReownsOtherwise() => Sta.Run(() =>
    {
        IntPtr first = NewOwner(), second = NewOwner();
        var window = NewWindow(first);
        IntPtr hwnd = window.Handle;
        Assert.Equal(first, NativeMethods.GetOwner(hwnd));

        window.SetOwnerWindow(first);
        Assert.Equal(first, NativeMethods.GetOwner(hwnd));

        window.SetOwnerWindow(second);
        Assert.Equal(second, NativeMethods.GetOwner(hwnd));

        window.SetOwnerWindow(IntPtr.Zero); // ignored
        Assert.Equal(second, NativeMethods.GetOwner(hwnd));
        window.CloseWindow();
    });

    [Fact]
    public void SetOwnerWindowBeforeTheHandleExistsWorks() => Sta.Run(() =>
    {
        IntPtr first = NewOwner(), second = NewOwner();
        var window = new ExplorerWindow(first, FakeProvider.Standard(), SettingsPath);
        window.SetOwnerWindow(second);
        new WindowInteropHelper(window).EnsureHandle();
        Assert.Equal(second, NativeMethods.GetOwner(window.Handle));
        window.CloseWindow();
    });

    // ----- chrome and metrics -----

    [Fact]
    public void ClientAreaIs414By119WithA390By70Box() => Sta.Run(() =>
    {
        var window = new ExplorerWindow(IntPtr.Zero, FakeProvider.Standard(), SettingsPath) { ShowActivated = false };
        var loaded = false;
        window.Loaded += (_, _) => loaded = true;
        window.Show();
        try
        {
            Assert.True(Wait.Until(() => loaded));
            // 119 DIPs is not a whole number of pixels at 125%/150% scaling, hence the half-pixel tolerance.
            Assert.Equal(414, window.Panel.ActualWidth, 0.7);
            Assert.Equal(119, window.Panel.ActualHeight, 0.7);
            // After the first layout the box stretches instead of being fixed.
            Assert.True(double.IsNaN(window.Panel.Box.Width));
            Assert.Equal(390, window.Panel.Box.ActualWidth, 0.7);
            Assert.Equal(70, window.Panel.Box.ActualHeight, 0.7);
            Assert.Equal(SizeToContent.Manual, window.SizeToContent);

            Assert.Equal(75, window.Panel.BackButton.ActualWidth, 0.7);
            Assert.Equal(23, window.Panel.CloseButton.ActualHeight, 0.7);
            Assert.False(window.Panel.BackButton.IsEnabled);
            Assert.True(window.Panel.CloseButton.IsDefault);
            Assert.Equal("Retraces your steps", window.Panel.BackButton.ToolTip);
            Assert.Equal("Exits Formula Explorer", window.Panel.CloseButton.ToolTip);
            Assert.Equal("Formula Explorer", window.Title);
            Assert.False(window.ShowInTaskbar);
            Assert.Equal(ResizeMode.CanResize, window.ResizeMode);

            // Min size = chrome + a 200x20 box.
            double chromeW = window.ActualWidth - window.Panel.ActualWidth;
            double chromeH = window.ActualHeight - window.Panel.ActualHeight;
            Assert.Equal(chromeW + 224, window.MinWidth, 3);
            Assert.Equal(chromeH + 69, window.MinHeight, 3);

            // Only the X in the title bar.
            long style = NativeMethods.GetWindowLong(window.Handle, NativeMethods.GWL_STYLE);
            Assert.Equal(0, style & (NativeMethods.WS_MINIMIZEBOX | NativeMethods.WS_MAXIMIZEBOX));
        }
        finally
        {
            window.CloseWindow();
        }
    });

    [Fact]
    public void FirstOpenIsAtTheTopRightInsideTheOwner() => Sta.Run(() =>
    {
        var ownerWindow = new Window
        {
            WindowStartupLocation = WindowStartupLocation.Manual, Left = 120, Top = 90, Width = 1000, Height = 640, ShowActivated = false,
        };
        ownerWindow.Show();
        var window = new ExplorerWindow(new WindowInteropHelper(ownerWindow).Handle, FakeProvider.Standard(), SettingsPath) { ShowActivated = false };
        window.Show();
        try
        {
            Sta.Pump(200);
            IntPtr ownerHwnd = new WindowInteropHelper(ownerWindow).Handle;
            var owner = NativeMethods.GetWindowBounds(ownerHwnd)!.Value;
            var own = NativeMethods.GetWindowBounds(window.Handle)!.Value;
            double scale = NativeMethods.GetDpiScale(ownerHwnd);

            Assert.Equal(owner.Right - Math.Round(24 * scale), own.Right);
            Assert.Equal(owner.Top + Math.Round(100 * scale), own.Top);
            Assert.Equal(ownerHwnd, NativeMethods.GetOwner(window.Handle));
        }
        finally
        {
            window.CloseWindow();
            ownerWindow.Close();
        }
    });

    [Fact]
    public void TooltipPopupIsClickThroughAndNonActivating() => Sta.Run(() =>
    {
        var window = new ExplorerWindow(IntPtr.Zero, FakeProvider.Standard(), SettingsPath) { ShowActivated = false };
        window.Show();
        window.ShowFormula(Displays.Of("=IF(C3,0,C11)"));
        try
        {
            Sta.Pump(200);
            var view = window.Panel.View;
            Assert.Equal(IntPtr.Zero, view.TooltipHandle);

            window.SimulateHover(4);
            Assert.True(Wait.Until(() => window.GetSnapshot().TooltipVisible && view.TooltipHandle != IntPtr.Zero));
            IntPtr popup = view.TooltipHandle;
            Assert.NotEqual(IntPtr.Zero, popup);
            long style = NativeMethods.GetWindowLong(popup, NativeMethods.GWL_EXSTYLE);
            Assert.Equal(NativeMethods.WS_EX_TRANSPARENT | NativeMethods.WS_EX_NOACTIVATE,
                style & (NativeMethods.WS_EX_TRANSPARENT | NativeMethods.WS_EX_NOACTIVATE));

            window.SimulateClick(4); // a click hides it...
            Assert.False(window.GetSnapshot().TooltipVisible);
            window.SimulateHover(4); // ...and it stays hidden on the same node
            Assert.False(window.GetSnapshot().TooltipVisible);

            window.SimulateHover(1); // another node shows its own
            Assert.True(window.GetSnapshot().TooltipVisible);
            window.SimulateHover(null);
            Assert.False(window.GetSnapshot().TooltipVisible);
            Assert.Equal(IntPtr.Zero, view.TooltipHandle);
        }
        finally
        {
            window.CloseWindow();
        }
    });

    [Fact]
    public void BoundsAreSavedOnCloseAndRestoredOnTheNextOpen() => Sta.Run(() =>
    {
        var first = new ExplorerWindow(IntPtr.Zero, FakeProvider.Standard(), SettingsPath) { ShowActivated = false };
        first.Show();
        Sta.Pump(200);
        first.Left = 50;
        first.Top = 60;
        first.Width = 500;
        first.Height = 260;
        Assert.True(Wait.Until(() => Math.Abs(first.ActualWidth - 500) < 1 && Math.Abs(first.ActualHeight - 260) < 1));
        first.CloseWindow();

        var saved = WindowSettings.Load(SettingsPath);
        Assert.NotNull(saved);
        Assert.Equal(500, saved!.Value.Width, 1);
        Assert.Equal(260, saved.Value.Height, 1);

        var second = new ExplorerWindow(IntPtr.Zero, FakeProvider.Standard(), SettingsPath) { ShowActivated = false };
        second.Show();
        try
        {
            Assert.True(Wait.Until(() => second.IsLoaded && Math.Abs(second.ActualWidth - 500) < 1));
            Assert.Equal(saved.Value.Left, second.Left, 1);
            Assert.Equal(saved.Value.Top, second.Top, 1);
            Assert.Equal(500, second.ActualWidth, 1);
            Assert.Equal(260, second.ActualHeight, 1);
            Assert.True(second.MinWidth > 200);
        }
        finally
        {
            second.CloseWindow();
        }
    });

    // ----- view commands and test surface -----

    [Fact]
    public void ViewCommandsReachTheView() => Sta.Run(() =>
    {
        var window = NewWindow();
        var display = Displays.Of("=IF(C3,0,C11)");
        window.ShowLoading();
        Assert.True(window.GetSnapshot().IsLoading);
        Assert.Equal(string.Empty, window.GetSnapshot().DisplayText);

        window.ShowFormula(display);
        window.SetBackEnabled(true);
        window.SetGold(Displays.Find(display, "C3", NodeKind.Reference));

        var snapshot = window.GetSnapshot();
        Assert.Equal("=IF(C3,0,C11)", snapshot.DisplayText);
        Assert.True(snapshot.Interactive);
        Assert.False(snapshot.IsLoading);
        Assert.True(snapshot.BackEnabled);
        Assert.Equal((4, 2), (snapshot.GoldStart, snapshot.GoldLength));
        Assert.True(snapshot.PixelsPerDip > 0);
        Assert.Null(snapshot.HoverStart);
        window.CloseWindow();
    });

    [Fact]
    public void SimulatedInputGoesThroughTheViewAndRaisesTheEvents() => Sta.Run(() =>
    {
        var window = NewWindow();
        var display = Displays.Of("=IF(C3,0,C11)");
        window.ShowFormula(display);
        window.Panel.Measure(new Size(414, 119));
        window.Panel.Arrange(new Rect(0, 0, 414, 119));
        window.Panel.UpdateLayout();

        window.SimulateHover(4);
        var hover = window.GetSnapshot();
        Assert.Equal((4, 2), (hover.HoverStart, hover.HoverLength));
        Assert.True(hover.HoverIsLink);
        Assert.True(hover.TooltipVisible);
        Assert.Equal("value of Reference", hover.TooltipText);

        window.SimulateHover(null);
        Assert.Null(window.GetSnapshot().HoverStart);

        var events = new List<string>();
        window.NodeClicked += (_, n) => events.Add("click:" + n.GetText(display.Formula.Text));
        window.NodeDoubleClicked += (_, n) => events.Add("double:" + n.GetText(display.Formula.Text));

        window.SimulateClick(4);
        Assert.Equal(new[] { "click:C3" }, events);

        events.Clear();
        window.SimulateDoubleClick(9); // C11 (first press, then the second one)
        Assert.Equal(new[] { "click:C11", "double:C11" }, events);
        window.CloseWindow();
    });

    [Fact]
    public void BackAndCloseButtonsRaiseTheirEvents() => Sta.Run(() =>
    {
        var window = NewWindow();
        int backs = 0, closes = 0;
        window.BackRequested += (_, _) => backs++;
        window.CloseRequested += (_, _) => closes++;

        window.Panel.BackButton.RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent));
        window.Panel.CloseButton.RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent));
        Assert.Equal((1, 1), (backs, closes));
        window.CloseWindow();
    });

    [Fact]
    public void CharScreenPointIsNullUntilLaidOutAndOnScreenAfterShow() => Sta.Run(() =>
    {
        var window = new ExplorerWindow(IntPtr.Zero, FakeProvider.Standard(), SettingsPath) { ShowActivated = false };
        window.ShowFormula(Displays.Of("=IF(C3,0,C11)"));
        Assert.Null(window.GetCharScreenPoint(4)); // not attached to a window yet

        window.Show();
        try
        {
            Assert.True(Wait.Until(() => window.GetCharScreenPoint(4) is not null));
            var point = window.GetCharScreenPoint(4);
            Assert.NotNull(point);
            var bounds = window.GetSnapshot().ScreenBounds;
            Assert.InRange(point!.Value.X, bounds.X, bounds.X + bounds.Width);
            Assert.InRange(point.Value.Y, bounds.Y, bounds.Y + bounds.Height);
            Assert.Null(window.GetCharScreenPoint(1000));
        }
        finally
        {
            window.CloseWindow();
        }
    });

    [Fact]
    public void ClickingALinkKeepsTheWindowActive() => Sta.Run(() =>
    {
        // The handler stands in for Excel's Goto activating another window; the explorer must take the activation back.
        var window = new ExplorerWindow(IntPtr.Zero, FakeProvider.Standard(), SettingsPath);
        var display = Displays.Of("=IF(C3,0,C11)");
        var other = new Window { ShowActivated = true, Width = 100, Height = 100, WindowStartupLocation = WindowStartupLocation.Manual, Left = 0, Top = 0 };
        window.NodeClicked += (_, _) => other.Activate();
        window.Show();
        window.ShowFormula(display);
        other.Show();
        window.Activate();
        Wait.Until(() => window.IsActive, 2000);
        try
        {
            if (!window.IsActive) return; // the desktop refused foreground activation (e.g. a locked session): nothing to verify

            window.SimulateClick(4);
            Assert.True(Wait.Until(() => window.IsActive), "the window took its activation back"); // queued at Input priority
        }
        finally
        {
            other.Close();
            window.CloseWindow();
        }
    });

    [Fact]
    public void BackKeepsTheWindowActive() => Sta.Run(() =>
    {
        // Back navigates (Goto / Workbook.Activate may activate Excel): the explorer must take the activation back.
        var window = new ExplorerWindow(IntPtr.Zero, FakeProvider.Standard(), SettingsPath);
        var other = new Window { ShowActivated = true, Width = 100, Height = 100, WindowStartupLocation = WindowStartupLocation.Manual, Left = 0, Top = 0 };
        window.BackRequested += (_, _) => other.Activate();
        window.Show();
        window.SetBackEnabled(true);
        other.Show();
        window.Activate();
        Wait.Until(() => window.IsActive, 2000);
        try
        {
            if (!window.IsActive) return; // the desktop refused foreground activation: nothing to verify

            window.Panel.BackButton.RaiseEvent(new RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent));
            Assert.True(Wait.Until(() => window.IsActive), "the window took its activation back"); // queued at Input priority
        }
        finally
        {
            other.Close();
            window.CloseWindow();
        }
    });

    [Fact]
    public void EnterOnBackKeepsTheWindowActive() => Sta.Run(() =>
    {
        var window = new ExplorerWindow(IntPtr.Zero, FakeProvider.Standard(), SettingsPath);
        var other = new Window { ShowActivated = true, Width = 100, Height = 100, WindowStartupLocation = WindowStartupLocation.Manual, Left = 0, Top = 0 };
        int backs = 0;
        window.BackRequested += (_, _) =>
        {
            backs++;
            other.Activate();
        };
        window.Show();
        window.SetBackEnabled(true);
        other.Show();
        window.Activate();
        Wait.Until(() => window.IsActive, 2000);
        try
        {
            if (!window.IsActive) return;
            window.Panel.BackButton.Focus();
            if (!window.Panel.BackButton.IsKeyboardFocused) return;

            Assert.True(window.HandleKeyDown(NativeMethods.VK_RETURN));
            Assert.Equal(1, backs);
            Assert.True(Wait.Until(() => window.IsActive), "the window took its activation back"); // queued at Input priority
        }
        finally
        {
            other.Close();
            window.CloseWindow();
        }
    });

    // ----- owner destroyed -----

    private static HwndSource NewRawOwner() =>
        new(new HwndSourceParameters("FeTestOwner") { WindowStyle = unchecked((int)0x80000000), Width = 300, Height = 200 }); // WS_POPUP

    [Fact]
    public void DestroyingTheOwnerSavesTheBoundsAndAsksTheControllerToClose() => Sta.Run(() =>
    {
        var owner = NewRawOwner();
        var window = new ExplorerWindow(owner.Handle, FakeProvider.Standard(), SettingsPath) { ShowActivated = false };
        int closeRequests = 0;
        window.CloseRequested += (_, _) => closeRequests++;
        window.Show();
        Assert.True(Wait.Until(() => window.IsLoaded));
        window.Left = 70;
        window.Top = 80;
        window.Width = 520;
        window.Height = 270;
        Assert.True(Wait.Until(() => Math.Abs(window.ActualWidth - 520) < 1 && Math.Abs(window.ActualHeight - 270) < 1));
        Assert.True(window.IsAlive);

        owner.Dispose(); // Win32 destroys the owned window without WM_CLOSE
        Assert.Equal(0, closeRequests); // posted, not raised from inside the destruction
        Assert.True(Wait.Until(() => closeRequests == 1));

        Assert.Equal(1, closeRequests);
        Assert.False(window.IsAlive);
        var saved = WindowSettings.Load(SettingsPath);
        Assert.NotNull(saved);
        Assert.Equal(520, saved!.Value.Width, 1);
        Assert.Equal(270, saved.Value.Height, 1);
        Assert.Equal(70, saved.Value.Left, 1);
        Assert.Equal(80, saved.Value.Top, 1);

        window.CloseWindow(); // what the controller does next: harmless on the dead window
        Sta.Pump(50);
        Assert.Equal(1, closeRequests);
    });

    [Fact]
    public void ANormalCloseDoesNotLookLikeADestroyedOwner() => Sta.Run(() =>
    {
        var window = new ExplorerWindow(IntPtr.Zero, FakeProvider.Standard(), SettingsPath) { ShowActivated = false };
        int closeRequests = 0;
        window.CloseRequested += (_, _) => closeRequests++;
        window.Show();
        Sta.Pump(100);
        window.CloseWindow();
        Sta.Pump(200);
        Assert.Equal(0, closeRequests);
        Assert.False(window.IsAlive);
    });
}
