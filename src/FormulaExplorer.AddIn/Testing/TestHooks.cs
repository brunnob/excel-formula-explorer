using System.IO;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using ExcelDna.Integration;
using FormulaExplorer.AddIn.Composition;
using FormulaExplorer.AddIn.Diagnostics;
using FormulaExplorer.AddIn.Excel;
using FormulaExplorer.AddIn.UI;
using Microsoft.Win32;
using XL = Microsoft.Office.Interop.Excel;

namespace FormulaExplorer.AddIn.Testing;

/// <summary>
/// Entry points for the integration tests, called from outside as <c>$xl.Run("FE_Test_Hover", 12)</c>.
/// They are plain static methods WITHOUT Excel-DNA attributes (with explicit exports an attributed method would be
/// registered at load time, bypassing the gate) and are registered at run time by <see cref="RegisterIfEnabled"/>,
/// only when <c>HKCU\Software\FormulaExplorer</c> has the DWORD <c>EnableTestHooks</c> = 1.
/// Every hook returns <c>"ok"</c> or <c>"error: ..."</c> and never throws into Excel. Hooks that make the explorer
/// post work (loads are delayed 150 ms) return immediately; poll <c>FE_Test_Dump</c> for the expected state.
/// </summary>
public static class TestHooks
{
    private const string RegistryKeyPath = @"Software\FormulaExplorer";
    private const string RegistryValueName = "EnableTestHooks";
    private const string Ok = "ok";

    private static readonly JsonSerializerOptions DumpOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        Converters = { new JsonStringEnumConverter() },
    };

    /// <summary>True when the registry gate that enables the hooks is on.</summary>
    public static bool IsEnabled()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath);
            return key?.GetValue(RegistryValueName) is int value && value == 1;
        }
        catch (Exception ex)
        {
            Log.Error("TestHooks.IsEnabled", ex);
            return false;
        }
    }

    /// <summary>Registers the FE_Test_* macro functions with Excel if (and only if) the registry gate is on.</summary>
    public static void RegisterIfEnabled()
    {
        if (!IsEnabled()) return;

        // (hook method, Excel name, argument names)
        var hooks = new (string Method, string Name, string[] Args)[]
        {
            (nameof(Open), "FE_Test_Open", []),
            (nameof(Hover), "FE_Test_Hover", ["charIndex"]),
            (nameof(Click), "FE_Test_Click", ["charIndex"]),
            (nameof(DoubleClick), "FE_Test_DoubleClick", ["charIndex"]),
            (nameof(Back), "FE_Test_Back", []),
            (nameof(Close), "FE_Test_Close", []),
            (nameof(Dump), "FE_Test_Dump", ["path"]),
            (nameof(EditProbe), "FE_Test_EditProbe", ["path", "delayMs"]),
            (nameof(DeferSetup), "FE_Test_DeferSetup", ["sheet", "address"]),
            (nameof(DeferEnd), "FE_Test_DeferEnd", ["delayMs"]),
            (nameof(ProbeWorkbook), "FE_Test_ProbeWorkbook", ["name", "path", "delayMs"]),
        };

        var methods = new List<MethodInfo>();
        var functionAttributes = new List<object>();
        var argumentAttributes = new List<List<object>>();
        foreach (var (method, name, args) in hooks)
        {
            methods.Add(typeof(TestHooks).GetMethod(method, BindingFlags.Public | BindingFlags.Static)!);
            // IsMacroType gives the function the macro context that Application.Run needs; hidden from the function wizard.
            functionAttributes.Add(new ExcelFunctionAttribute { Name = name, IsMacroType = true, IsHidden = true });
            argumentAttributes.Add(args.Select(a => (object)new ExcelArgumentAttribute { Name = a }).ToList());
        }

        ExcelIntegration.RegisterMethods(methods, functionAttributes, argumentAttributes);
        Log.Info($"Test hooks registered ({hooks.Length} functions).");
    }

    /// <summary>Opens the explorer (or brings the open one to the front).</summary>
    public static object Open() => Run(nameof(Open), ExplorerController.OpenOrThrow);

    /// <summary>Mouse over character <paramref name="charIndex"/> of the formula; -1 = the mouse left the text.</summary>
    public static object Hover(object charIndex) => Run(nameof(Hover), () =>
    {
        var index = ToInt(charIndex);
        RequireWindow().SimulateHover(index < 0 ? null : index);
    });

    /// <summary>One click on character <paramref name="charIndex"/>.</summary>
    public static object Click(object charIndex) => Run(nameof(Click), () => RequireWindow().SimulateClick(ToInt(charIndex)));

    /// <summary>Double click on character <paramref name="charIndex"/>.</summary>
    public static object DoubleClick(object charIndex) => Run(nameof(DoubleClick), () => RequireWindow().SimulateDoubleClick(ToInt(charIndex)));

    /// <summary>Presses the explorer's Back button (a real Click event on the button, so the normal handler runs).</summary>
    public static object Back() => Run(nameof(Back), () =>
    {
        var button = FindButton(RequireWindow(), "Back") ?? throw new InvalidOperationException("Back button not found.");
        if (!button.IsEnabled) throw new InvalidOperationException("Back button is disabled.");
        button.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent, button));
    });

    /// <summary>Closes the explorer exactly as the Close button does.</summary>
    public static object Close() => Run(nameof(Close), ExplorerController.CloseAll);

    /// <summary>
    /// Writes the current state as UTF-8 indented JSON:
    /// <c>{ isOpen, session, view, overlay: [...], charPoints: { "index": [x, y] } }</c>.
    /// <c>charPoints</c> maps each character index of the displayed formula to its centre in physical screen pixels.
    /// </summary>
    public static object Dump(object path) => Run(nameof(Dump), () =>
    {
        var target = path as string;
        if (string.IsNullOrWhiteSpace(target)) throw new ArgumentException("A file path is required.");

        var window = ExplorerController.CurrentWindow;
        var session = ExplorerController.CurrentSession?.GetSnapshot();
        var view = window?.GetSnapshot();
        var overlay = ExplorerController.CurrentOverlay?.GetDrawnRects() ?? [];

        var charPoints = new SortedDictionary<int, int[]>();
        if (window != null && view != null)
        {
            for (var i = 0; i < view.DisplayText.Length; i++)
            {
                var point = window.GetCharScreenPoint(i);
                if (point is { } p) charPoints[i] = [p.X, p.Y];
            }
        }

        var overlayWindow = ExplorerController.CurrentOverlay;
        var json = JsonSerializer.Serialize(new
        {
            isOpen = ExplorerController.IsOpen,
            explorerHwnd = window?.Handle.ToInt64() ?? 0,
            overlayHwnd = overlayWindow?.WindowHandle.ToInt64() ?? 0,
            overlayVisible = overlayWindow?.WindowVisible ?? false,
            buttons = window == null ? null : new
            {
                back = ScreenRect(FindButton(window, "Back")),
                close = ScreenRect(FindButton(window, "Close")),
                backFocused = FindButton(window, "Back")?.IsKeyboardFocused ?? false,
                closeFocused = FindButton(window, "Close")?.IsKeyboardFocused ?? false,
                backDefault = FindButton(window, "Back")?.IsDefault ?? false,
                closeDefault = FindButton(window, "Close")?.IsDefault ?? false,
                backEnabled = FindButton(window, "Back")?.IsEnabled ?? false,
                windowActive = window.IsActive,
            },
            editMode = ReadEditModeSignals(),
            session, view, overlay, charPoints,
        }, DumpOptions);
        var directory = Path.GetDirectoryName(Path.GetFullPath(target));
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        File.WriteAllText(target, json, new UTF8Encoding(false));
    });

    /// <summary>The raw edit-mode observations, for diagnostics (does not log and does not decide).</summary>
    private static object ReadEditModeSignals()
    {
        var (disabled, probed) = EditModeSignals.CommandVote(ExcelDnaUtil.Application);
        return new { editorWindowVisible = EditModeSignals.EditorWindowVisible(), commandsDisabled = disabled, commandsProbed = probed };
    }

    /// <summary>Writes the enabled state of a wide list of built-in commands (seen from inside Excel) as JSON, to find edit-only ones.</summary>
    public static object EditProbe(object path, object delayMs) => Run(nameof(EditProbe), () => RunLater(ToInt(delayMs), () =>
    {
        var target = path as string ?? throw new ArgumentException("A file path is required.");
        dynamic bars = ((dynamic)ExcelDnaUtil.Application).CommandBars;
        var ids = new[]
        {
            "FileNewDefault", "FileOpen", "FileSave", "FileSaveAs", "FilePrint", "FileClose", "PasteSpecialDialog", "SheetInsert", "Undo",
            "Copy", "Cut", "Paste", "GoTo", "FormatCellsDialog", "SortDialog", "WindowNew", "ZoomDialog", "ReviewNewComment", "NameManager",
            "HyperlinkInsert", "SymbolInsert", "AutoSum", "MacroPlay", "SheetDelete", "FileProperties", "FilePrintPreview", "InsertFunction",
            "Find", "Replace", "PageSetupDialog", "PivotTableInsert", "TableInsert", "SpellingCheck", "ProtectSheet", "ViewFreezePanes",
        };
        var states = new SortedDictionary<string, string>();
        foreach (var id in ids)
        {
            try { states[id] = ((bool)bars.GetEnabledMso(id)).ToString(); }
            catch (Exception ex) { states[id] = "err:" + ex.GetType().Name + ":" + ex.Message; }
        }
        File.WriteAllText(target, JsonSerializer.Serialize(new { signals = ReadEditModeSignals(), states }, DumpOptions), new UTF8Encoding(false));
    }));

    /// <summary>
    /// What <see cref="ExcelGateway.ProbeWorkbook"/> (the probe of a pending workbook close) answers for <paramref name="name"/>, and
    /// how many windows the gateway finds for it, written as JSON to <paramref name="path"/> after <paramref name="delayMs"/>. The
    /// delay is what lets the answer be taken while a modal dialog (the "Save changes?" prompt) is up: Excel rejects Application.Run
    /// from outside then, but the dispatcher keeps running and the call is made from Excel's own main thread.
    /// </summary>
    public static object ProbeWorkbook(object name, object path, object delayMs) => Run(nameof(ProbeWorkbook), () => RunLater(ToInt(delayMs), () =>
    {
        var target = path as string ?? throw new ArgumentException("A file path is required.");
        var workbook = name as string ?? throw new ArgumentException("A workbook name is required.");
        var gateway = new ExcelGateway((XL.Application)ExcelDnaUtil.Application);
        try
        {
            var presence = gateway.ProbeWorkbook(workbook);
            var windows = gateway.GetWindowHandlesOf(workbook).Count;
            File.WriteAllText(target, JsonSerializer.Serialize(new { workbook, presence, windows }, DumpOptions), new UTF8Encoding(false));
        }
        finally
        {
            gateway.Dispose();
        }
    }));

    private static ExcelGateway? _deferGateway;

    /// <summary>
    /// Sets up what a session leaves behind when it ends while Excel cannot take calls: a gateway of its own unhides the (hidden or
    /// very hidden) <paramref name="sheet"/> of the active workbook and selects <paramref name="address"/> there, then the user's
    /// selection moves to Val_FO!M5 behind the gateway's back, so its undo record is no longer needed by the selection and is still
    /// pending. <see cref="DeferEnd"/> then ends the "session" (while the test script has Excel in cell edit mode).
    /// </summary>
    public static object DeferSetup(object sheet, object address) => Run(nameof(DeferSetup), () =>
    {
        var app = (XL.Application)ExcelDnaUtil.Application;
        _deferGateway?.Dispose();
        _deferGateway = new ExcelGateway(app);
        var workbook = app.ActiveWorkbook.Name;
        var result = _deferGateway.NavigateTo(new CellLocation(workbook, (string)sheet, (string)address));
        if (result.Status != NavigationStatus.Ok) throw new InvalidOperationException($"NavigateTo failed: {result.Status} {result.Message}");

        var home = (XL.Worksheet)app.ActiveWorkbook.Worksheets["Val_FO"];
        home.Activate();
        home.Range["M5"].Select();
    });

    /// <summary>Ends the gateway session started by <see cref="DeferSetup"/> after <paramref name="delayMs"/> (use it to let the script enter edit mode first).</summary>
    public static object DeferEnd(object delayMs) => Run(nameof(DeferEnd), () => RunLater(ToInt(delayMs), () =>
    {
        var gateway = _deferGateway ?? throw new InvalidOperationException("DeferSetup was not called.");
        _deferGateway = null;
        gateway.EndSession();
        gateway.Dispose();
    }));

    /// <summary>
    /// Runs <paramref name="body"/> on Excel's main thread after a delay, from the dispatcher (which keeps running while a cell
    /// is being edited, unlike <c>Application.Run</c> from outside, which Excel rejects then); 0 = now.
    /// </summary>
    private static void RunLater(int delayMs, Action body)
    {
        if (delayMs <= 0)
        {
            body();
            return;
        }
        var timer = new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromMilliseconds(delayMs) };
        timer.Tick += (_, _) =>
        {
            timer.Stop();
            try { body(); }
            catch (Exception ex) { Log.Error("TestHooks.RunLater", ex); }
        };
        timer.Start();
    }

    /// <summary>Runs a hook body, turning any exception into the returned error text (and the log).</summary>
    private static object Run(string hook, Action body)
    {
        try
        {
            body();
            return Ok;
        }
        catch (Exception ex)
        {
            Log.Error($"TestHooks.{hook}", ex);
            return $"error: {ex.Message}";
        }
    }

    private static ExplorerWindow RequireWindow()
        => ExplorerController.CurrentWindow ?? throw new InvalidOperationException("The explorer is not open.");

    /// <summary>Excel passes numbers as double (Application.Run from COM may also hand over text).</summary>
    private static int ToInt(object value) => value switch
    {
        double d => (int)Math.Round(d),
        int i => i,
        string s when double.TryParse(s, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var d) => (int)Math.Round(d),
        _ => throw new ArgumentException($"A character index is required, got {value?.GetType().Name ?? "null"}."),
    };

    /// <summary>Screen rectangle [left, top, right, bottom] in physical pixels of a visible element; null when unavailable.</summary>
    private static int[]? ScreenRect(FrameworkElement? element)
    {
        if (element == null || !element.IsVisible) return null;
        var scale = VisualTreeHelper.GetDpi(element);
        var topLeft = element.PointToScreen(new System.Windows.Point(0, 0));
        return
        [
            (int)Math.Round(topLeft.X), (int)Math.Round(topLeft.Y),
            (int)Math.Round(topLeft.X + element.ActualWidth * scale.DpiScaleX), (int)Math.Round(topLeft.Y + element.ActualHeight * scale.DpiScaleY),
        ];
    }

    private static Button? FindButton(DependencyObject root, string caption)
    {
        if (root is Button button && string.Equals((button.Content as string)?.Replace("_", ""), caption, StringComparison.OrdinalIgnoreCase))
            return button;
        for (var i = 0; i < VisualTreeHelper.GetChildrenCount(root); i++)
        {
            var found = FindButton(VisualTreeHelper.GetChild(root, i), caption);
            if (found != null) return found;
        }
        return null;
    }
}
