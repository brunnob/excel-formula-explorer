namespace FormulaExplorer.AddIn.Excel;

/// <summary>A rectangular area of a sheet, 1-based inclusive row/column numbers.</summary>
public readonly record struct AreaSpan(int Row1, int Row2, int Col1, int Col2)
{
    public int RowCount => Row2 - Row1 + 1;
    public int ColumnCount => Col2 - Col1 + 1;
}

/// <summary>
/// Pure A1-address arithmetic on the addresses the gateway produces (<c>Range.Address(true, true)</c>, no sheet
/// prefix): "$K$31", "$I$19:$J$19", "$A:$C", "$3:$5", "$A$1:$B$2,$D$4". No Excel access.
/// </summary>
public static class ExcelAddressMath
{
    public const int MaxRows = 1_048_576;
    public const int MaxColumns = 16_384;

    /// <summary>
    /// Parses an address (one or more areas separated by ',' — or ';' as some locales write the union) into areas.
    /// Whole-column areas span all rows and whole-row areas span all columns. Returns false for anything that is not a
    /// plain sheet-less A1 address.
    /// </summary>
    public static bool TryParseAreas(string? address, out IReadOnlyList<AreaSpan> areas)
    {
        var result = new List<AreaSpan>();
        areas = result;
        if (string.IsNullOrWhiteSpace(address)) return false;

        foreach (var part in address.Split(',', ';'))
        {
            var ends = part.Split(':');
            if (ends.Length is < 1 or > 2) return false;
            if (!TryParseReference(ends[0], out var first)) return false;
            var second = first;
            if (ends.Length == 2 && !TryParseReference(ends[1], out second)) return false;

            // Both ends must be of the same kind (cell:cell, column:column or row:row).
            if (first.Kind != second.Kind) return false;
            var (row1, row2) = first.Kind == RefKind.Column ? (1, MaxRows) : (Math.Min(first.Row, second.Row), Math.Max(first.Row, second.Row));
            var (col1, col2) = first.Kind == RefKind.Row ? (1, MaxColumns) : (Math.Min(first.Column, second.Column), Math.Max(first.Column, second.Column));
            result.Add(new AreaSpan(row1, row2, col1, col2));
        }
        return result.Count > 0;
    }

    /// <summary>"A" for 1, "Z" for 26, "AA" for 27, "XFD" for 16384.</summary>
    public static string ColumnName(int column)
    {
        if (column < 1 || column > MaxColumns) throw new ArgumentOutOfRangeException(nameof(column));
        Span<char> buffer = stackalloc char[3];
        var pos = buffer.Length;
        while (column > 0)
        {
            column--;
            buffer[--pos] = (char)('A' + column % 26);
            column /= 26;
        }
        return new string(buffer[pos..]);
    }

    /// <summary>Column number of a column name (case-insensitive), or 0 when it is not a valid column (1-3 letters, at most XFD).</summary>
    public static int ColumnNumber(string name)
    {
        if (name.Length is < 1 or > 3) return 0;
        var number = 0;
        foreach (var ch in name)
        {
            var upper = char.ToUpperInvariant(ch);
            if (upper is < 'A' or > 'Z') return 0;
            number = number * 26 + (upper - 'A' + 1);
        }
        return number <= MaxColumns ? number : 0;
    }

    /// <summary>Groups ascending-or-not numbers into runs of consecutive values, e.g. 3,4,5,9,11,12 → (3,5) (9,9) (11,12).</summary>
    public static IReadOnlyList<(int First, int Last)> Runs(IEnumerable<int> numbers)
    {
        var runs = new List<(int First, int Last)>();
        foreach (var n in numbers.Distinct().Order())
        {
            if (runs.Count > 0 && runs[^1].Last == n - 1)
                runs[^1] = (runs[^1].First, n);
            else
                runs.Add((n, n));
        }
        return runs;
    }

    private enum RefKind { Cell, Column, Row }

    private readonly record struct Reference(RefKind Kind, int Row, int Column);

    /// <summary>Parses "$A$1", "A1", "$A", "C", "$12", "12" (a '$' only directly before the letters and before the digits).</summary>
    private static bool TryParseReference(string text, out Reference reference)
    {
        reference = default;
        var s = text.AsSpan().Trim();
        var i = 0;
        if (i < s.Length && s[i] == '$') i++;

        var lettersStart = i;
        while (i < s.Length && char.IsAsciiLetter(s[i])) i++;
        var letters = s[lettersStart..i];

        if (letters.Length > 0 && i < s.Length && s[i] == '$') i++; // "$A$1": the second '$' needs letters before it and digits after

        var digits = s[i..];
        if (letters.Length == 0 && digits.Length == 0) return false;
        if (digits.Length > 0 && !AllDigits(digits)) return false;
        if (digits.Length == 0 && s[^1] == '$') return false; // "A$"

        var column = letters.Length > 0 ? ColumnNumber(letters.ToString()) : 0;
        if (letters.Length > 0 && column == 0) return false;
        var row = 0;
        if (digits.Length > 0 && (!int.TryParse(digits, out row) || row < 1 || row > MaxRows)) return false;

        reference = (letters.Length > 0, digits.Length > 0) switch
        {
            (true, true) => new Reference(RefKind.Cell, row, column),
            (true, false) => new Reference(RefKind.Column, 0, column),
            _ => new Reference(RefKind.Row, row, 0),
        };
        return true;
    }

    private static bool AllDigits(ReadOnlySpan<char> text)
    {
        foreach (var ch in text)
            if (!char.IsAsciiDigit(ch)) return false;
        return true;
    }
}
