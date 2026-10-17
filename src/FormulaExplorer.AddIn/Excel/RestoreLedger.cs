namespace FormulaExplorer.AddIn.Excel;

public enum RestoreKind
{
    /// <summary>A hidden/very hidden sheet that was made visible; <see cref="RestoreRecord.OriginalVisibility"/> is its old state.</summary>
    SheetVisibility,
    /// <summary>Rows (1-based numbers) hidden by a collapsed outline that were unhidden.</summary>
    Rows,
    /// <summary>Columns (1-based numbers) hidden by a collapsed outline that were unhidden.</summary>
    Columns,
}

/// <summary>One change the gateway made to the workbook and may have to undo. Identity (not value) equality.</summary>
public sealed class RestoreRecord
{
    internal RestoreRecord(RestoreKind kind, string workbookName, string sheetName, int originalVisibility, IReadOnlyList<int> indexes)
    {
        Kind = kind;
        WorkbookName = workbookName;
        SheetName = sheetName;
        OriginalVisibility = originalVisibility;
        Indexes = indexes;
    }

    public RestoreKind Kind { get; }
    public string WorkbookName { get; }
    public string SheetName { get; }

    /// <summary>XlSheetVisibility value before the change (only for <see cref="RestoreKind.SheetVisibility"/>).</summary>
    public int OriginalVisibility { get; }

    /// <summary>Row or column numbers (only for <see cref="RestoreKind.Rows"/> / <see cref="RestoreKind.Columns"/>).</summary>
    public IReadOnlyList<int> Indexes { get; }

    public bool IsOnSameSheetAs(CellLocation location) =>
        string.Equals(WorkbookName, location.WorkbookName, StringComparison.OrdinalIgnoreCase) &&
        string.Equals(SheetName, location.SheetName, StringComparison.OrdinalIgnoreCase);
}

/// <summary>
/// Bookkeeping of the unhide/expand operations done by <see cref="ExcelGateway.NavigateTo"/> and the decision of which
/// ones are still needed. Pure (no Excel access) so the rules can be unit-tested.
/// </summary>
public sealed class RestoreLedger
{
    private readonly List<RestoreRecord> _records = new();

    /// <summary>All records, oldest first.</summary>
    public IReadOnlyList<RestoreRecord> Records => _records;

    public int Count => _records.Count;

    public void AddSheetVisibility(string workbookName, string sheetName, int originalVisibility) =>
        _records.Add(new RestoreRecord(RestoreKind.SheetVisibility, workbookName, sheetName, originalVisibility, Array.Empty<int>()));

    public void AddRows(string workbookName, string sheetName, IEnumerable<int> rows) => AddIndexes(RestoreKind.Rows, workbookName, sheetName, rows);

    public void AddColumns(string workbookName, string sheetName, IEnumerable<int> columns) => AddIndexes(RestoreKind.Columns, workbookName, sheetName, columns);

    public void Remove(RestoreRecord record) => _records.Remove(record);

    public void Clear() => _records.Clear();

    /// <summary>Moves every record into a new ledger (same order) and leaves this one empty.</summary>
    public RestoreLedger Detach()
    {
        var detached = new RestoreLedger();
        detached._records.AddRange(_records);
        _records.Clear();
        return detached;
    }

    private void AddIndexes(RestoreKind kind, string workbookName, string sheetName, IEnumerable<int> indexes)
    {
        var list = indexes.Distinct().Order().ToArray();
        if (list.Length > 0)
            _records.Add(new RestoreRecord(kind, workbookName, sheetName, 0, list));
    }

    /// <summary>
    /// The records to undo so that everything in <paramref name="keepVisible"/> stays visible, NEWEST FIRST (rows and
    /// columns are re-hidden before the sheet that contains them is hidden again). A record is needed when some kept
    /// location is on the same workbook+sheet and — for rows/columns — its range touches one of the record's
    /// rows/columns; a sheet record is needed by any kept location on that sheet. An empty collection undoes everything.
    /// </summary>
    public IReadOnlyList<RestoreRecord> SelectUndo(IReadOnlyCollection<CellLocation> keepVisible)
    {
        var keep = keepVisible.Select(Resolve).ToList();
        var needed = new HashSet<RestoreRecord>();

        foreach (var record in _records.Where(r => r.Kind != RestoreKind.SheetVisibility))
        {
            if (keep.Any(k => record.IsOnSameSheetAs(k.Location) && k.Touches(record)))
                needed.Add(record);
        }

        foreach (var record in _records.Where(r => r.Kind == RestoreKind.SheetVisibility))
        {
            // A kept location on the sheet needs it visible; so does any kept row/column record of that sheet
            // (implied by the first rule, stated explicitly because undoing it would hide the kept rows with it).
            if (keep.Any(k => record.IsOnSameSheetAs(k.Location)) || needed.Any(n => SameSheet(n, record)))
                needed.Add(record);
        }

        return _records.Where(r => !needed.Contains(r)).Reverse().ToList();
    }

    private static bool SameSheet(RestoreRecord a, RestoreRecord b) =>
        string.Equals(a.WorkbookName, b.WorkbookName, StringComparison.OrdinalIgnoreCase) &&
        string.Equals(a.SheetName, b.SheetName, StringComparison.OrdinalIgnoreCase);

    private static KeptLocation Resolve(CellLocation location) =>
        new(location, ExcelAddressMath.TryParseAreas(location.Address, out var areas) ? areas : null);

    /// <summary>A kept location with its parsed areas (null = unparseable address, conservatively treated as the whole sheet).</summary>
    private sealed record KeptLocation(CellLocation Location, IReadOnlyList<AreaSpan>? Areas)
    {
        public bool Touches(RestoreRecord record)
        {
            if (Areas is null) return true;
            foreach (var area in Areas)
            {
                foreach (var index in record.Indexes)
                {
                    var (from, to) = record.Kind == RestoreKind.Rows ? (area.Row1, area.Row2) : (area.Col1, area.Col2);
                    if (index >= from && index <= to) return true;
                }
            }
            return false;
        }
    }
}
