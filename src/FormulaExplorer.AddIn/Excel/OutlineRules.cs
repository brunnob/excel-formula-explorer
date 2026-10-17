using System.Runtime.InteropServices;

namespace FormulaExplorer.AddIn.Excel;

/// <summary>What the gateway reads about one row/column when deciding whether it is an outline-hidden line.</summary>
/// <param name="Hidden">The line is currently hidden.</param>
/// <param name="Level">Its OutlineLevel (1 = not part of any group; 0 = unknown / outside the sheet).</param>
internal readonly record struct LineState(bool Hidden, int Level);

/// <summary>
/// The decisions of the outline unhide/re-hide steps of <see cref="ExcelGateway"/>, kept free of Excel access (the
/// object-model reads and writes arrive as delegates) so they can be unit-tested.
/// </summary>
internal static class OutlineRules
{
    /// <summary>
    /// Unhides, one line at a time, the hidden lines in first..last whose outline level is above 1, appending each line
    /// to <paramref name="unhidden"/> as soon as it is changed. A non-busy COM failure from <paramref name="unhide"/>
    /// (a protected worksheet refuses to change Hidden) stops the pass: revealing is cosmetic, the caller can still
    /// select the target. A busy failure propagates.
    /// </summary>
    public static void UnhideHidden(int first, int last, Func<int, LineState> probe, Action<int> unhide, List<int> unhidden, Action<string> log)
    {
        for (var line = first; line <= last; line++)
        {
            var state = probe(line);
            if (!state.Hidden || state.Level <= 1) continue;

            try
            {
                unhide(line);
            }
            catch (COMException ex) when (!ComErrors.IsBusy(ex))
            {
                log($"could not unhide line {line} ({ex.Message}); leaving the rest of the span as it is");
                return;
            }
            unhidden.Add(line);
        }
    }

    /// <summary>
    /// Of the lines the gateway unhid earlier, the ones that are still worth hiding again: a line that is visible, still
    /// part of an outline (level above 1) and whose group still looks collapsed, i.e. a neighbouring line of the same or a
    /// deeper level is hidden — or is another line of the record, which was hidden by that same collapse. Anything else
    /// means the user moved, ungrouped or expanded the group since, and hiding the stored number would hit the wrong line.
    /// </summary>
    public static IReadOnlyList<int> LinesToRehide(IReadOnlyList<int> recorded, Func<int, LineState> probe)
    {
        var inRecord = new HashSet<int>(recorded);
        var result = new List<int>();

        // Everything is read before the caller hides anything, so the lines of one record do not affect each other.
        foreach (var line in recorded.Distinct().Order())
        {
            var state = probe(line);
            if (state.Hidden || state.Level <= 1) continue;

            if (IsCollapsedNeighbour(line - 1) || IsCollapsedNeighbour(line + 1))
                result.Add(line);

            bool IsCollapsedNeighbour(int other)
            {
                if (inRecord.Contains(other)) return true;
                var neighbour = probe(other);
                return neighbour.Hidden && neighbour.Level >= state.Level;
            }
        }
        return result;
    }
}
