namespace FormulaExplorer.Core.Syntax;

/// <summary>Result of parsing a formula text. Produced by <see cref="FormulaParser.Parse"/>; never null, never throws.</summary>
public sealed class ParsedFormula
{
    public ParsedFormula(string text, IReadOnlyList<Token> tokens, FormulaNode? root, bool hasErrors)
    {
        Text = text;
        Tokens = tokens;
        Root = root;
        HasErrors = hasErrors;
        Nodes = root is null ? Array.Empty<SyntaxNode>() : root.DescendantsAndSelf().ToArray();
    }

    /// <summary>The exact text that is displayed (Range.Formula).</summary>
    public string Text { get; }

    /// <summary>True when <see cref="Text"/> starts with '='. When false, <see cref="Root"/> is null and nothing is interactive.</summary>
    public bool IsFormula => Text.Length > 0 && Text[0] == '=';

    public IReadOnlyList<Token> Tokens { get; }

    /// <summary>
    /// Root of the tree (span = whole text). Null only when <see cref="IsFormula"/> is false.
    /// When the parser hits a syntax error it still returns a root: the parseable parts become
    /// proper nodes and the rest becomes <see cref="UnparsedNode"/>s; reference/name/structured-reference
    /// tokens inside unparsed regions are still emitted as leaf nodes so that coloring and links keep working.
    /// </summary>
    public FormulaNode? Root { get; }

    /// <summary>True when error recovery was needed.</summary>
    public bool HasErrors { get; }

    /// <summary>All nodes in pre-order (root first). Empty when <see cref="Root"/> is null.</summary>
    public IReadOnlyList<SyntaxNode> Nodes { get; }
}
