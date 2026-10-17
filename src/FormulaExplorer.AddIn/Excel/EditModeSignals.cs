using System.Runtime.InteropServices;
using System.Text;

namespace FormulaExplorer.AddIn.Excel;

/// <summary>
/// The independent observations <see cref="ComErrors.IsEditMode"/> combines to tell whether Excel is editing a cell
/// (in-cell, F2, formula bar, Point mode). None of them is a safe test alone:
/// a built-in command can be disabled by policy, a template or another add-in; and the editor's class name is an
/// implementation detail of Excel.
/// </summary>
internal static class EditModeSignals
{
    /// <summary>Class of the in-cell editor window Excel shows (as a child of the workbook window) only while a cell is being edited.</summary>
    internal const string EditorClass = "EXCEL6";

    /// <summary>
    /// Built-in commands from unrelated families that are enabled with a workbook open and disabled by Excel while a cell
    /// is being edited (measured with this exact list: tests/integration/Probe-EditMode.ps1, results in
    /// tests/integration/out/editprobe-*.json). A policy or another add-in disables one or two of them, practically never most.
    /// </summary>
    internal static readonly string[] EditDisabledCommands =
    [
        "FileNewDefault", "FileOpen", "FilePrint", "SheetInsert", "SortDialog", "WindowNew", "NameManager", "MacroPlay", "ZoomDialog",
    ];

    /// <summary>Commands that must answer for the vote to count at all.</summary>
    internal const int MinimumProbed = 4;

    /// <summary>
    /// Combines the signals. A conclusive window observation decides; only when the window could not be looked at
    /// (<paramref name="editorWindowVisible"/> null) is the command vote used: at least <paramref name="minimumProbed"/>
    /// commands must have answered and at least three quarters of those must be disabled (7 of the 9 shipped commands). The
    /// vote is a threshold, not unanimity, so that one or two commands disabled for good by a policy, a template or another
    /// add-in cannot turn the fallback into "never in edit mode" (navigation would then commit the user's edit), while the
    /// same few commands disabled while NOT editing cannot make it report "editing" either.
    /// </summary>
    internal static bool Decide(bool? editorWindowVisible, int commandsDisabled, int commandsProbed, int minimumProbed = MinimumProbed)
    {
        if (editorWindowVisible is { } visible) return visible;
        return commandsProbed >= minimumProbed && commandsDisabled * 4 >= commandsProbed * 3;
    }

    /// <summary>
    /// Whether a visible in-cell editor window exists under any Excel main window (XLMAIN) of this thread; null when the
    /// windows could not be enumerated at all. Must run on Excel's main thread. Reads window state only (no COM).
    /// </summary>
    internal static bool? EditorWindowVisible()
    {
        try
        {
            bool sawMainWindow = false, found = false;
            EnumThreadWindows(GetCurrentThreadId(), (top, _) =>
            {
                if (ClassOf(top) != "XLMAIN") return true;
                sawMainWindow = true;
                EnumChildWindows(top, (child, _) =>
                {
                    if (ClassOf(child) != EditorClass || !IsWindowVisible(child)) return true;
                    found = true;
                    return false;
                }, IntPtr.Zero);
                return !found;
            }, IntPtr.Zero);
            return sawMainWindow ? found : null;
        }
        catch (Exception ex) when (ex is DllNotFoundException or EntryPointNotFoundException or MarshalDirectiveException)
        {
            return null;
        }
    }

    /// <summary>How many of <see cref="EditDisabledCommands"/> are disabled and how many answered. Never throws.</summary>
    internal static (int Disabled, int Probed) CommandVote(object application)
    {
        int disabled = 0, probed = 0;
        try
        {
            dynamic bars = ((dynamic)application).CommandBars;
            foreach (var id in EditDisabledCommands)
            {
                try
                {
                    bool enabled = bars.GetEnabledMso(id);
                    probed++;
                    if (!enabled) disabled++;
                }
                catch (Exception ex) when (ex is COMException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException or InvalidCastException)
                {
                    // this command cannot be asked: it takes no part in the vote
                }
            }
        }
        catch (Exception ex) when (ex is COMException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException or InvalidCastException)
        {
            // no command bars at all: no vote
        }
        return (disabled, probed);
    }

    private static string ClassOf(IntPtr hwnd)
    {
        var buffer = new StringBuilder(64);
        return GetClassName(hwnd, buffer, buffer.Capacity) > 0 ? buffer.ToString() : string.Empty;
    }

    private delegate bool EnumProc(IntPtr hwnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumThreadWindows(uint threadId, EnumProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr parent, EnumProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowVisible(IntPtr hwnd);

    [DllImport("user32.dll", EntryPoint = "GetClassNameW", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hwnd, StringBuilder buffer, int maxCount);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();
}
