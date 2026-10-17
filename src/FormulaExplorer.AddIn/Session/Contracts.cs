using FormulaExplorer.AddIn.Excel;
using FormulaExplorer.Core.Analysis;
using FormulaExplorer.Core.Syntax;

namespace FormulaExplorer.AddIn.Session;

/// <summary>Everything the view needs to render one root cell.</summary>
/// <param name="Formula">Parsed text (Range.Formula). For a non-formula cell, IsFormula is false.</param>
/// <param name="HitMap">Hover map (all-null for non-formulas).</param>
/// <param name="Colors">Reference colors (empty for non-formulas).</param>
/// <param name="Interactive">False for non-formula cells: text in black, no hover, no tooltip, no clicks.</param>
public sealed record FormulaDisplay(ParsedFormula Formula, HitMap HitMap, ColorMap Colors, bool Interactive);

/// <summary>What the view shows for a hovered node.</summary>
/// <param name="IsLink">Underline + hand cursor; single click navigates; double click drills down.</param>
/// <param name="TooltipText">Text under the bold "Returns" title; null = no tooltip.</param>
public sealed record NodeInfo(bool IsLink, string? TooltipText);

/// <summary>Implemented by the session; the view asks it about the node under the mouse (evaluation is lazy and cached).</summary>
public interface INodeInfoProvider
{
    NodeInfo GetNodeInfo(SyntaxNode node);
}

/// <summary>
/// Forwards to <see cref="Target"/>; lets the controller create the window (which needs a provider) before the session
/// (which needs the window) exists.
/// </summary>
public sealed class NodeInfoRelay : INodeInfoProvider
{
    private static readonly NodeInfo Nothing = new(false, null);

    public INodeInfoProvider? Target { get; set; }

    public NodeInfo GetNodeInfo(SyntaxNode node) => Target?.GetNodeInfo(node) ?? Nothing;
}

/// <summary>The explorer window as seen by the session. All calls on Excel's main thread.</summary>
public interface IExplorerView
{
    /// <summary>Blank text area + wait cursor (shown briefly while a new root loads).</summary>
    void ShowLoading();

    /// <summary>Shows a new root. Clears hover, gold and tooltip; restores the normal cursor.</summary>
    void ShowFormula(FormulaDisplay display);

    /// <summary>Paints <paramref name="node"/> with the gold (#FFD700) background, or removes the gold when null. Only one at a time.</summary>
    void SetGold(SyntaxNode? node);

    void SetBackEnabled(bool enabled);

    /// <summary>Shows an informational message box (English) owned by the explorer window.</summary>
    void ShowMessage(string message);

    /// <summary>Values changed: re-query the provider for the hovered node (underline/cursor/tooltip text).</summary>
    void RefreshHover();

    /// <summary>Single click on a node for which the provider said IsLink.</summary>
    event EventHandler<SyntaxNode>? NodeClicked;

    /// <summary>Double click on a link node. Always preceded by a NodeClicked for the same node (first click of the pair).</summary>
    event EventHandler<SyntaxNode>? NodeDoubleClicked;

    /// <summary>Back button.</summary>
    event EventHandler? BackRequested;

    /// <summary>Close button, title-bar X, or Esc. The view does NOT close itself; the controller does.</summary>
    event EventHandler? CloseRequested;
}

/// <summary>A colored border to draw around a precedent on the root's sheet.</summary>
public sealed record PrecedentBox(CellLocation Location, uint Rgb);

/// <summary>A rectangle currently drawn by the overlay, in physical screen pixels (diagnostics/tests).</summary>
public sealed record OverlayRect(int Left, int Top, int Right, int Bottom, uint Rgb);

/// <summary>Transparent click-through layer drawing Range-Finder-style borders on the root's precedents.</summary>
public interface IPrecedentOverlay : IDisposable
{
    /// <summary>HWND of the explorer window; the overlay keeps itself just below it in z-order.</summary>
    void SetExplorerWindow(IntPtr explorerHwnd);

    /// <summary>
    /// Draws <paramref name="boxes"/> (all on the root's sheet) whenever the root's workbook window and sheet are active,
    /// following scroll, zoom, frozen/split panes, window moves and resizes. Replaces any previous boxes.
    /// </summary>
    void Show(CellLocation root, IReadOnlyList<PrecedentBox> boxes);

    /// <summary>Removes all borders (hides the layer).</summary>
    void Clear();

    /// <summary>What is drawn right now (empty when hidden).</summary>
    IReadOnlyList<OverlayRect> GetDrawnRects();
}

/// <summary>Posts work to Excel's main thread message loop (WPF Dispatcher in production, immediate in tests).</summary>
public interface IUiScheduler
{
    /// <summary>Runs <paramref name="action"/> later on the UI thread, after pending input/render work (Background priority).</summary>
    void Post(Action action);

    /// <summary>Runs <paramref name="action"/> on the UI thread after <paramref name="delay"/>.</summary>
    void PostDelayed(TimeSpan delay, Action action);
}

/// <summary>Session state for tests and diagnostics.</summary>
public sealed record SessionSnapshot(
    bool IsOpen,
    CellLocation? Root,
    string DisplayText,
    IReadOnlyList<CellLocation> Stack,
    int? GoldStart,
    int? GoldLength,
    CellLocation? GoldTarget,
    bool BackEnabled,
    IReadOnlyList<PrecedentBox> Boxes);
