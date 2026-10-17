using FormulaExplorer.Core.Syntax;

namespace FormulaExplorer.Core.Analysis;

/// <summary>
/// Maps each character index of the formula text to the node that is highlighted when the mouse is over that character.
/// Rules:
/// <list type="bullet">
/// <item>'=' at index 0 → the <see cref="FormulaNode"/> (whole formula).</item>
/// <item>any char of a Reference / Name / StructuredReference / Literal → that node.</item>
/// <item>operator chars of a Binary node → the Binary node (incl. ':' ' ' ',' reference operators).</item>
/// <item>operator char of a Prefix / Postfix node → that node.</item>
/// <item>'(' or ')' of a Paren node → the Paren node.</item>
/// <item>function-name chars, '(' or ')' of a FunctionCall → the FunctionCall node.</item>
/// <item>'{', '}', ',' and ';' of an array constant → the ArrayConstant node.</item>
/// <item>argument-separator ',', insignificant whitespace, Unknown chars, Unparsed regions → null.</item>
/// </list>
/// </summary>
public sealed class HitMap
{
    private readonly SyntaxNode?[] _owners;

    private HitMap(SyntaxNode?[] owners) => _owners = owners;

    /// <summary>Builds the map. For a non-formula (<see cref="ParsedFormula.IsFormula"/> false) every index maps to null.</summary>
    public static HitMap Build(ParsedFormula formula)
    {
        if (!formula.IsFormula || formula.Root is null)
            return new HitMap(Array.Empty<SyntaxNode?>());

        var owners = new SyntaxNode?[formula.Text.Length];
        // Pre-order: a child always overwrites what its parent marked on the same characters.
        foreach (var node in formula.Nodes)
            Mark(node, formula, owners);
        return new HitMap(owners);
    }

    /// <summary>The highlightable node for the char at <paramref name="charIndex"/>, or null. Out-of-range index → null.</summary>
    public SyntaxNode? NodeAt(int charIndex) =>
        (uint)charIndex < (uint)_owners.Length ? _owners[charIndex] : null;

    /// <summary>
    /// True for node kinds that may become hyperlinks when their evaluation returns a Range:
    /// Reference, Name, StructuredReference, FunctionCall, and Binary nodes whose operator is a reference operator (':' ' ' ',').
    /// </summary>
    public static bool IsLinkCandidate(SyntaxNode node) => node switch
    {
        NameNode name => !name.IsLocal, // a LET/LAMBDA variable is not a workbook name
        ReferenceNode or StructuredReferenceNode or FunctionCallNode => true,
        BinaryNode binary => binary.IsReferenceOperator,
        _ => false,
    };

    /// <summary>Marks the characters that belong to the node itself (not to its children).</summary>
    private static void Mark(SyntaxNode node, ParsedFormula formula, SyntaxNode?[] owners)
    {
        switch (node)
        {
            case FormulaNode:
                Fill(owners, 0, 1, node); // the leading '='
                break;
            case LiteralNode or ReferenceNode or NameNode or StructuredReferenceNode:
                Fill(owners, node.Start, node.Length, node);
                break;
            case BinaryNode binary:
                Fill(owners, binary.OperatorStart, binary.OperatorLength, node);
                break;
            case PrefixNode prefix:
                Fill(owners, prefix.OperatorStart, 1, node);
                break;
            case PostfixNode postfix:
                Fill(owners, postfix.OperatorStart, 1, node);
                break;
            case ParenNode paren:
                Fill(owners, paren.OpenParen, 1, node);
                if (paren.CloseParen >= 0) Fill(owners, paren.CloseParen, 1, node);
                break;
            case FunctionCallNode call:
                // An invocation's "name" is its callee, which marks its own characters.
                if (call.Callee is null) Fill(owners, call.Start, call.NameLength, node);
                Fill(owners, call.OpenParen, 1, node);
                if (call.CloseParen >= 0) Fill(owners, call.CloseParen, 1, node);
                break;
            case ArrayConstantNode array:
                MarkArrayPunctuation(array, formula.Tokens, owners);
                break;
            // MissingArgument (zero length) and Unparsed regions are never highlighted.
        }
    }

    /// <summary>'{', '}' and the ',' / ';' separators of an array constant belong to the constant itself.</summary>
    private static void MarkArrayPunctuation(ArrayConstantNode array, IReadOnlyList<Token> tokens, SyntaxNode?[] owners)
    {
        int lo = 0, hi = tokens.Count;
        while (lo < hi) // first token that starts at or after the constant
        {
            int mid = (lo + hi) / 2;
            if (tokens[mid].Start < array.Start) lo = mid + 1; else hi = mid;
        }
        for (int i = lo; i < tokens.Count && tokens[i].Start < array.End; i++)
        {
            if (tokens[i].Kind is TokenKind.ArrayOpen or TokenKind.ArrayClose
                or TokenKind.ArrayColumnSeparator or TokenKind.ArrayRowSeparator)
                Fill(owners, tokens[i].Start, tokens[i].Length, array);
        }
    }

    private static void Fill(SyntaxNode?[] owners, int start, int length, SyntaxNode node)
    {
        int end = Math.Min(start + length, owners.Length);
        for (int i = Math.Max(start, 0); i < end; i++)
            owners[i] = node;
    }
}
