namespace FormulaExplorer.AddIn.Overlay;

/// <summary>
/// Helpers for <c>Range.Address(true, true)</c> strings. That address separates the areas of a multi-area range with ','
/// in every locale, but <c>Worksheet.Range(address)</c> only accepts the UI list separator (';' in pt-BR) and fails with
/// 0x800A03EC on the ',' form, so callers resolve one area at a time.
/// </summary>
internal static class AreaAddress
{
    /// <summary>
    /// Splits <paramref name="address"/> on the ',' that separates areas. A ',' inside a single-quoted sheet name is not
    /// a separator (addresses without a sheet prefix never contain one; the quote handling is defensive). Parts are
    /// trimmed; empty parts are dropped.
    /// </summary>
    public static IReadOnlyList<string> Split(string address)
    {
        var parts = new List<string>();
        if (string.IsNullOrWhiteSpace(address)) return parts;

        var inQuotes = false;
        var start = 0;
        for (var i = 0; i < address.Length; i++)
        {
            var c = address[i];
            if (c == '\'') inQuotes = !inQuotes;
            else if (c == ',' && !inQuotes)
            {
                Add(parts, address.AsSpan(start, i - start));
                start = i + 1;
            }
        }

        Add(parts, address.AsSpan(start));
        return parts;
    }

    private static void Add(List<string> parts, ReadOnlySpan<char> part)
    {
        var trimmed = part.Trim();
        if (!trimmed.IsEmpty) parts.Add(trimmed.ToString());
    }
}
