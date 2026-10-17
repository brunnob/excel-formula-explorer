using FormulaExplorer.Core.Syntax;

namespace FormulaExplorer.Core.Analysis;

/// <summary>Result of color assignment: which nodes are colored and with which RGB color.</summary>
public sealed class ColorMap
{
    public ColorMap(IReadOnlyList<ColoredNode> nodes) => Nodes = nodes;

    public static ColorMap Empty { get; } = new(Array.Empty<ColoredNode>());

    /// <summary>Colored nodes in order of appearance in the text.</summary>
    public IReadOnlyList<ColoredNode> Nodes { get; }

    /// <summary>RGB (0xRRGGBB) of a colored node, or null if the node is not colored.</summary>
    public uint? ColorOf(SyntaxNode node)
    {
        foreach (var n in Nodes)
            if (ReferenceEquals(n.Node, node)) return n.Rgb;
        return null;
    }
}

/// <param name="Node">A Reference, Name or StructuredReference node.</param>
/// <param name="PaletteIndex">Index into <see cref="RefColorizer.Palette"/>.</param>
/// <param name="Rgb">0xRRGGBB.</param>
/// <param name="Key">Normalized identity: occurrences with the same key share a color.</param>
public sealed record ColoredNode(SyntaxNode Node, int PaletteIndex, uint Rgb, string Key);

/// <summary>
/// Assigns Excel Range-Finder colors to references, in order of first appearance, cycling through 7 colors.
/// </summary>
public static class RefColorizer
{
    /// <summary>Range Finder palette (0xRRGGBB): blue, green, purple, maroon, light green, orange, magenta.</summary>
    public static IReadOnlyList<uint> Palette { get; } = new uint[]
    {
        0x0000FF, 0x008000, 0x9900CC, 0x800000, 0x00CC33, 0xFF6600, 0xCC0099,
    };

    /// <summary>
    /// Colors every Reference, StructuredReference and Name node for which <paramref name="isColorable"/> returns true
    /// (the add-in passes a predicate that excludes names not referring to a range, and #REF! references).
    /// Order = order of appearance (by Start). A repeated reference (same normalized key) reuses its first color
    /// and does not consume a palette slot. Key normalization: case-insensitive; '$' removed; a reference without a sheet
    /// prefix is keyed with <paramref name="rootSheetName"/> (so A1 and Sheet1!A1 share a color when the root is on Sheet1);
    /// names keyed by prefix+name; structured refs by their whole text.
    /// </summary>
    public static ColorMap Assign(ParsedFormula formula, string? rootSheetName, Func<SyntaxNode, bool> isColorable)
    {
        if (!formula.IsFormula || formula.Root is null) return ColorMap.Empty;

        var colored = new List<ColoredNode>();
        var paletteIndexByKey = new Dictionary<string, int>(StringComparer.Ordinal);

        // Nodes come in pre-order and the leaves are already in text order; sorting by Start keeps the contract
        // independent of how error recovery nested them.
        var candidates = formula.Nodes
            .Where(n => n.Kind is NodeKind.Reference or NodeKind.Name or NodeKind.StructuredReference)
            .OrderBy(n => n.Start);

        foreach (var node in candidates)
        {
            if (node is NameNode { IsLocal: true }) continue; // LET/LAMBDA variable, not a workbook name
            if (!isColorable(node)) continue;

            string key = KeyOf(node, formula.Text, rootSheetName);
            if (!paletteIndexByKey.TryGetValue(key, out int index))
            {
                index = paletteIndexByKey.Count % Palette.Count; // the 8th distinct reference wraps to the 1st color
                paletteIndexByKey.Add(key, index);
            }
            colored.Add(new ColoredNode(node, index, Palette[index], key));
        }
        return colored.Count == 0 ? ColorMap.Empty : new ColorMap(colored);
    }

    /// <summary>Identity of a reference for coloring: case-insensitive, no '$', no quotes, missing sheet = root sheet.</summary>
    private static string KeyOf(SyntaxNode node, string text, string? rootSheetName)
    {
        switch (node)
        {
            case ReferenceNode r:
                {
                    string book = r.Workbook is null ? string.Empty : "[" + r.Workbook + "]";
                    string sheet = r.Sheet ?? (r.Workbook is null ? rootSheetName ?? string.Empty : string.Empty);
                    string last = r.LastSheet is null ? string.Empty : ":" + r.LastSheet;
                    return ("R:" + book + sheet + last + "!" + r.Address.Replace("$", string.Empty)).ToUpperInvariant();
                }
            case NameNode name:
                {
                    string book = name.Workbook is null ? string.Empty : "[" + name.Workbook + "]";
                    return ("N:" + book + (name.Sheet ?? string.Empty) + "!" + name.Name).ToUpperInvariant();
                }
            default:
                return "S:" + node.GetText(text).ToUpperInvariant();
        }
    }
}
