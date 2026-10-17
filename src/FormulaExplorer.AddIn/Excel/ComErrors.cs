using System.Runtime.InteropServices;
using FormulaExplorer.AddIn.Diagnostics;

namespace FormulaExplorer.AddIn.Excel;

/// <summary>
/// The "Excel is in cell edit mode" condition raised by <see cref="ComErrors.ThrowIfEditMode"/>. It carries the same
/// HRESULT Excel itself uses (<see cref="ComErrors.VbaIgnore"/>), so every busy handler treats it as busy, but unlike the
/// raw HRESULT it says WHY: only this exception makes <see cref="ExcelBusyException.IsEditMode"/> true. A
/// VBA_E_IGNORE that Excel raised by itself is "not ready right now" and is retried like any other busy answer.
/// </summary>
public sealed class EditModeException : COMException
{
    public EditModeException() : base("Excel is in cell edit mode.", ComErrors.VbaIgnore) { }
}

/// <summary>Classification of the COM failures Excel produces when it cannot take object-model calls.</summary>
public static class ComErrors
{
    /// <summary>VBA_E_IGNORE: Excel is in cell-edit mode (or otherwise not ready for the call).</summary>
    internal const int VbaIgnore = unchecked((int)0x800AC472);

    /// <summary>RPC_E_CALL_REJECTED: a modal dialog (or similar) is rejecting incoming calls.</summary>
    private const int CallRejected = unchecked((int)0x80010001);

    /// <summary>RPC_E_SERVERCALL_RETRYLATER: the same condition reported through the retry-later path.</summary>
    private const int RetryLater = unchecked((int)0x8001010A);

    private static bool _editModeLogged;

    /// <summary>True when the failure only means "Excel is busy right now, retry later".</summary>
    public static bool IsBusy(COMException ex) => ex.HResult is VbaIgnore or CallRejected or RetryLater;

    /// <summary>
    /// True while Excel is in cell edit mode (in-cell, F2, formula bar, Point mode). Called from inside Excel the object
    /// model does NOT refuse most calls then (reads work; <c>Application.Goto</c> or <c>Hidden = ...</c> silently end the
    /// edit and commit it), so state-changing calls must check this first and treat edit mode as "busy". (A modal dialog
    /// does not matter here: those make Excel reject the calls with a real busy HRESULT.)
    /// <para>
    /// The decision rests on the in-cell editor window that Excel shows only while a cell is being edited (see
    /// <see cref="EditModeSignals"/>); when the windows cannot be inspected it falls back to a vote (at least three quarters)
    /// of several unrelated built-in commands that Excel disables while editing. A single command is never enough: policies,
    /// templates and other add-ins disable commands such as <c>FileNewDefault</c> permanently.
    /// </para>
    /// The first detection of each edit session is logged once.
    /// </summary>
    public static bool IsEditMode(object application) => IsEditMode(application, EditModeSignals.EditorWindowVisible);

    internal static bool IsEditMode(object application, Func<bool?> editorWindowVisible)
    {
        bool? editor = editorWindowVisible();
        bool editing;
        string signal;
        if (editor is { } visible)
        {
            editing = visible;
            signal = "editor window";
        }
        else
        {
            var (disabled, probed) = EditModeSignals.CommandVote(application);
            editing = EditModeSignals.Decide(null, disabled, probed);
            signal = $"{disabled}/{probed} edit-only commands disabled";
        }

        NoteEditMode(editing, signal);
        return editing;
    }

    /// <summary>Logs the transition into edit mode once (not on every poll while the edit lasts).</summary>
    private static void NoteEditMode(bool editing, string signal)
    {
        if (editing)
        {
            if (_editModeLogged) return;
            _editModeLogged = true;
            Log.Info($"ComErrors: Excel is in cell edit mode ({signal}); navigation and restores wait for it to end.");
        }
        else
        {
            _editModeLogged = false;
        }
    }

    /// <summary>Throws <see cref="EditModeException"/> (a busy answer that is explicitly edit mode) when <see cref="IsEditMode"/> is true.</summary>
    public static void ThrowIfEditMode(object application)
    {
        if (IsEditMode(application)) throw new EditModeException();
    }
}
