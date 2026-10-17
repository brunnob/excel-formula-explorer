using System.Text;
using FormulaExplorer.Core.Analysis;
using FormulaExplorer.Core.Syntax;

namespace FormulaExplorer.Core.Tests;

/// <summary>Compact textual renderings of tokens and trees so expectations fit on one line.</summary>
internal static class TestHelpers
{
    private static string Code(TokenKind kind) => kind switch
    {
        TokenKind.FormulaStart => "Start",
        TokenKind.Number => "N",
        TokenKind.String => "S",
        TokenKind.Bool => "B",
        TokenKind.Error => "E",
        TokenKind.Reference => "R",
        TokenKind.Name => "Nm",
        TokenKind.StructuredReference => "SR",
        TokenKind.Function => "F",
        TokenKind.BinaryOperator => "Bin",
        TokenKind.PrefixOperator => "Pre",
        TokenKind.PostfixOperator => "Post",
        TokenKind.IntersectionOperator => "Isect",
        TokenKind.UnionOperator => "Union",
        TokenKind.ArgumentSeparator => "Sep",
        TokenKind.OpenParen => "Open",
        TokenKind.CloseParen => "Close",
        TokenKind.ArrayOpen => "ArrOpen",
        TokenKind.ArrayClose => "ArrClose",
        TokenKind.ArrayColumnSeparator => "ColSep",
        TokenKind.ArrayRowSeparator => "RowSep",
        TokenKind.Whitespace => "Ws",
        _ => "U",
    };

    /// <summary>Tokens as <c>Kind[text]</c> separated by spaces, leaving out the leading '=' token.</summary>
    public static string Tokens(string text)
    {
        var tokens = FormulaTokenizer.Tokenize(text);
        return string.Join(" ", tokens.Where(t => t.Kind != TokenKind.FormulaStart).Select(t => $"{Code(t.Kind)}[{t.Text}]"));
    }

    /// <summary>The tree as an S-expression, see the cases in <c>ParserTests</c>.</summary>
    public static string Tree(string text)
    {
        var parsed = FormulaParser.Parse(text);
        return parsed.Root is null ? "<none>" : Dump(parsed.Root, text);
    }

    public static string Dump(SyntaxNode? node, string text)
    {
        switch (node)
        {
            case null:
                return string.Empty;
            case FormulaNode f:
                return Dump(f.Body, text);
            case LiteralNode or ReferenceNode or NameNode or StructuredReferenceNode:
                return node.GetText(text);
            case MissingArgumentNode:
                return "_";
            case BinaryNode b:
                {
                    string op = b.Operator switch { " " => "isect", "," => "union", _ => b.Operator };
                    return $"({op} {Dump(b.Left, text)} {Dump(b.Right, text)})";
                }
            case PrefixNode p:
                return $"(pre{p.Operator} {Dump(p.Operand, text)})";
            case PostfixNode p:
                return $"(post{p.Operator} {Dump(p.Operand, text)})";
            case ParenNode p:
                return "[" + Dump(p.Inner, text) + "]";
            case FunctionCallNode c:
                return c.Name + "(" + string.Join(";", c.Arguments.Select(a => Dump(a, text))) + ")";
            case ArrayConstantNode a:
                return "{" + string.Join(";", a.Rows.Select(r => string.Join(",", r.Select(e => Dump(e, text))))) + "}";
            case UnparsedNode u:
                return "?<" + u.GetText(text) + ">(" + string.Join(",", u.Children.Select(c => Dump(c, text))) + ")";
            default:
                return node.ToString()!;
        }
    }

    /// <summary>The text of the node that would be highlighted for the character at <paramref name="index"/>, or null.</summary>
    public static string? HighlightAt(string text, int index)
    {
        var node = HitMap.Build(FormulaParser.Parse(text)).NodeAt(index);
        return node?.GetText(text);
    }

    /// <summary>Same as <see cref="HighlightAt"/> for the first character of the <paramref name="occurrence"/>-th match of <paramref name="hover"/>.</summary>
    public static string? HighlightOn(string text, string hover, int occurrence = 0, int offset = 0)
    {
        int index = -1;
        for (int i = 0; i <= occurrence; i++)
        {
            index = text.IndexOf(hover, index + 1, StringComparison.Ordinal);
            if (index < 0) throw new InvalidOperationException($"'{hover}' (#{occurrence}) not found in '{text}'");
        }
        return HighlightAt(text, index + offset);
    }

    /// <summary>All nodes in the tree, pre-order.</summary>
    public static IEnumerable<SyntaxNode> AllNodes(string text) => FormulaParser.Parse(text).Nodes;

    public static string Describe(IEnumerable<SyntaxNode> nodes, string text) =>
        string.Join(" ", nodes.Select(n => $"{n.Kind}[{n.GetText(text)}]"));

    public static string JoinTokenTexts(IReadOnlyList<Token> tokens)
    {
        var sb = new StringBuilder();
        foreach (var t in tokens) sb.Append(t.Text);
        return sb.ToString();
    }
}
