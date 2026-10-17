namespace FormulaExplorer.AddIn.Excel;

/// <summary>
/// Remembers the selection Excel had at the last look so that a change made by the user can be told from what the gateway
/// itself did. The selection is the ACTIVE CELL plus the selected address: collapsing a range the explorer selected down
/// to its own active cell leaves the active cell unchanged but is still a manual selection.
/// </summary>
internal sealed class SelectionTracker
{
    private CellLocation? _location;
    private string? _address;

    public SelectionTracker(CellLocation? location, string? address)
    {
        Sync(location, address);
    }

    /// <summary>The active cell at the last look.</summary>
    public CellLocation? Location => _location;

    /// <summary>
    /// Records the current selection without reporting it (initial state, or after the gateway moved the selection).
    /// A null location or address (no active cell; a shape or chart is selected) leaves that part as it was.
    /// </summary>
    public void Sync(CellLocation? location, string? address)
    {
        if (location is not null) _location = location;
        if (address is not null) _address = address;
    }

    /// <summary>
    /// Records the current selection and tells whether it differs from the last one. A non-range selection (null
    /// address) counts as unchanged so that selecting a shape does not move the root.
    /// </summary>
    public bool Observe(CellLocation location, string? address)
    {
        var changed = location != _location || (address is not null && address != _address);
        Sync(location, address);
        return changed;
    }
}
