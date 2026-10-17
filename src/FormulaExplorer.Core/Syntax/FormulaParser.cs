namespace FormulaExplorer.Core.Syntax;

/// <summary>
/// Builds the syntax tree with Excel's real operator precedence (highest first): reference operators
/// (<c>:</c>, intersection, union), prefix <c>-</c>/<c>+</c>/<c>@</c>, postfix <c>%</c>/<c>#</c>, <c>^</c>,
/// <c>* /</c>, <c>+ -</c>, <c>&amp;</c>, comparisons. All binary operators are left-associative.
/// </summary>
public static class FormulaParser
{
    /// <summary>Parses <paramref name="text"/> (Range.Formula). Never throws.</summary>
    public static ParsedFormula Parse(string text)
    {
        text ??= string.Empty;
        var tokens = FormulaTokenizer.Tokenize(text);
        if (text.Length == 0 || text[0] != '=')
            return new ParsedFormula(text, tokens, null, false);

        var parser = new Parser(text, tokens);
        var root = parser.ParseRoot();
        LocalNames.Mark(root);
        return new ParsedFormula(text, tokens, root, parser.HasErrors);
    }

    private sealed class Parser
    {
        /// <summary>Nesting limit (parentheses, calls, arrays, prefix chains). Excel itself stops at 64 levels;
        /// the limit keeps recursion shallow on hostile input.</summary>
        private const int MaxDepth = 128;

        private readonly string _text;

        /// <summary>Significant tokens: everything but the leading '=' and insignificant whitespace.</summary>
        private readonly List<Token> _sig = new();

        private int _pos;
        private int _depth;

        public bool HasErrors { get; private set; }

        public Parser(string text, IReadOnlyList<Token> tokens)
        {
            _text = text;
            foreach (var t in tokens)
                if (t.Kind != TokenKind.Whitespace && t.Kind != TokenKind.FormulaStart)
                    _sig.Add(t);
        }

        private Token? Peek => _pos < _sig.Count ? _sig[_pos] : null;

        /// <summary>End of the last consumed token (1 = just after '=' when nothing was consumed).</summary>
        private int LastEnd => _pos > 0 ? _sig[_pos - 1].End : 1;

        private bool PeekIs(TokenKind kind) => _pos < _sig.Count && _sig[_pos].Kind == kind;

        public FormulaNode ParseRoot()
        {
            var body = ParseExpr(1);
            if (Peek is not null)
                body = Recover(body, _ => false);
            return new FormulaNode(_text.Length, body);
        }

        // ------------------------------------------------------------------------------------------
        // Expressions
        // ------------------------------------------------------------------------------------------

        private static int Precedence(string op) => op switch
        {
            "=" or "<>" or "<" or ">" or "<=" or ">=" => 1,
            "&" => 2,
            "+" or "-" => 3,
            "*" or "/" => 4,
            "^" => 5,
            _ => 0,
        };

        /// <summary>Precedence climbing over the arithmetic/comparison operators. Null (nothing consumed) when no operand starts here.</summary>
        private SyntaxNode? ParseExpr(int minPrecedence)
        {
            var left = ParsePostfix();
            if (left is null) return null;

            while (Peek is { Kind: TokenKind.BinaryOperator } op)
            {
                int precedence = Precedence(op.Text);
                if (precedence == 0 || precedence < minPrecedence) break;

                int save = _pos;
                _pos++;
                var right = ParseExpr(precedence + 1);
                if (right is null)
                {
                    _pos = save; // dangling operator: leave it for error recovery
                    break;
                }
                left = MakeBinary(left, op, right);
            }
            return left;
        }

        private static BinaryNode MakeBinary(SyntaxNode left, Token op, SyntaxNode right) =>
            new(left.Start, right.End - left.Start, op.Kind == TokenKind.IntersectionOperator ? " " : op.Text,
                op.Start, op.Length, left, right);

        /// <summary>Postfix % and # (binds looser than prefix operators, tighter than ^).</summary>
        private SyntaxNode? ParsePostfix()
        {
            var node = ParsePrefix();
            if (node is null) return null;
            while (Peek is { Kind: TokenKind.PostfixOperator } op)
            {
                _pos++;
                node = new PostfixNode(node.Start, op.End - node.Start, op.Text, op.Start, node);
            }
            return node;
        }

        private SyntaxNode? ParsePrefix()
        {
            if (Peek is { Kind: TokenKind.PrefixOperator } op)
            {
                if (_depth >= MaxDepth) return null;
                int save = _pos;
                _pos++;
                _depth++;
                var operand = ParsePrefix();
                _depth--;
                if (operand is null)
                {
                    _pos = save;
                    return null;
                }
                return new PrefixNode(op.Start, operand.End - op.Start, op.Text, operand);
            }
            return ParseUnion();
        }

        // Reference operators, loosest to tightest: union ',', intersection ' ', range ':'.

        private SyntaxNode? ParseUnion()
        {
            var left = ParseIntersection();
            if (left is null) return null;
            while (Peek is { Kind: TokenKind.UnionOperator } op)
            {
                int save = _pos;
                _pos++;
                var right = ParseIntersection();
                if (right is null)
                {
                    _pos = save;
                    break;
                }
                left = MakeBinary(left, op, right);
            }
            return left;
        }

        private SyntaxNode? ParseIntersection()
        {
            var left = ParseRange();
            if (left is null) return null;
            while (Peek is { Kind: TokenKind.IntersectionOperator } op)
            {
                int save = _pos;
                _pos++;
                var right = ParseRange();
                if (right is null)
                {
                    _pos = save;
                    break;
                }
                left = MakeBinary(left, op, right);
            }
            return left;
        }

        private SyntaxNode? ParseRange()
        {
            var left = ParseSpillOperand();
            if (left is null) return null;
            while (Peek is { Kind: TokenKind.BinaryOperator, Text: ":" } op)
            {
                int save = _pos;
                _pos++;
                var right = ParseSpillOperand();
                if (right is null)
                {
                    _pos = save;
                    break;
                }
                left = MakeBinary(left, op, right);
            }
            return left;
        }

        /// <summary>A primary with its spill operators (<c>A1#</c>), which bind tighter than the reference operators.</summary>
        private SyntaxNode? ParseSpillOperand()
        {
            var node = ParsePrimary();
            if (node is null) return null;
            while (Peek is { Kind: TokenKind.PostfixOperator, Text: "#" } op)
            {
                _pos++;
                node = new PostfixNode(node.Start, op.End - node.Start, op.Text, op.Start, node);
            }
            return node;
        }

        // ------------------------------------------------------------------------------------------
        // Primaries
        // ------------------------------------------------------------------------------------------

        private SyntaxNode? ParsePrimary()
        {
            if (Peek is not { } t) return null;
            switch (t.Kind)
            {
                case TokenKind.Number:
                    _pos++;
                    return new LiteralNode(t.Start, t.Length, LiteralKind.Number);
                case TokenKind.String:
                    _pos++;
                    return new LiteralNode(t.Start, t.Length, LiteralKind.String);
                case TokenKind.Bool:
                    _pos++;
                    return new LiteralNode(t.Start, t.Length, LiteralKind.Bool);
                case TokenKind.Error:
                    _pos++;
                    return new LiteralNode(t.Start, t.Length, LiteralKind.Error);
                case TokenKind.Reference:
                case TokenKind.Name:
                case TokenKind.StructuredReference:
                    _pos++;
                    return MakeLeaf(t);
                case TokenKind.Function:
                    return ParseFunction(t);
                case TokenKind.OpenParen:
                    return ParseGroup(t);
                case TokenKind.ArrayOpen:
                    return ParseArray(t);
                default:
                    return null;
            }
        }

        private SyntaxNode MakeLeaf(Token t)
        {
            var prefix = ReferenceSyntax.GetPrefix(t.Text);
            string rest = t.Text.Substring(prefix.End);
            switch (t.Kind)
            {
                case TokenKind.Reference:
                    return new ReferenceNode(t.Start, t.Length, ReferenceSyntax.ClassifyAddress(rest), rest,
                        prefix.Workbook, prefix.Sheet, prefix.LastSheet);
                case TokenKind.Name:
                    return new NameNode(t.Start, t.Length, rest, prefix.Workbook, prefix.Sheet);
                default:
                    {
                        int bracket = rest.IndexOf('[');
                        string? table = bracket > 0 ? rest.Substring(0, bracket) : null;
                        return new StructuredReferenceNode(t.Start, t.Length, table, bracket >= 0 ? rest.Substring(bracket) : rest);
                    }
            }
        }

        private SyntaxNode? ParseFunction(Token nameToken)
        {
            if (_depth >= MaxDepth || _pos + 1 >= _sig.Count || _sig[_pos + 1].Kind != TokenKind.OpenParen)
                return null;

            _pos++; // the name; the '(' is next
            SyntaxNode call = ParseArgumentList(nameToken.Start, nameToken.Text, nameToken.Length, null);

            // Immediate invocation of the result: LAMBDA(x,x)(2), INDEX(fns,1)(x). Only a '(' glued to the ')' gets here
            // (whitespace before it would be an intersection), and the tokenizer already made its commas argument separators.
            int chain = 0;
            while (PeekIs(TokenKind.OpenParen) && _depth < MaxDepth && ++chain < MaxDepth)
                call = ParseArgumentList(call.Start, _text.Substring(call.Start, call.Length), call.Length, call);
            return call;
        }

        /// <summary>
        /// Parses <c>( args )</c> with the current token on the '('. <paramref name="callee"/> is null for a named call;
        /// for an invocation it is the call whose result is invoked (it becomes the first child).
        /// </summary>
        private FunctionCallNode ParseArgumentList(int start, string name, int nameLength, SyntaxNode? callee)
        {
            _depth++;
            _pos++;
            int openParen = _sig[_pos - 1].Start;
            var args = new List<SyntaxNode>();
            var separators = new List<int>();
            int closeParen = -1;

            bool first = true;
            while (true)
            {
                var t = Peek;
                if (first && t is { Kind: TokenKind.CloseParen })
                {
                    closeParen = t.Start;
                    _pos++;
                    break;
                }
                if (t is null)
                {
                    if (!first) args.Add(new MissingArgumentNode(LastEnd)); // "F(A1," with nothing after the comma
                    break;
                }

                SyntaxNode arg = t.Kind is TokenKind.ArgumentSeparator or TokenKind.CloseParen
                    ? new MissingArgumentNode(t.Start)
                    : ParseExpressionUntil(IsArgumentTerminator);
                args.Add(arg);
                first = false;

                t = Peek;
                if (t is null) break;
                _pos++;
                if (t.Kind == TokenKind.CloseParen)
                {
                    closeParen = t.Start;
                    break;
                }
                separators.Add(t.Start); // ArgumentSeparator: ParseExpressionUntil only stops at ',' or ')'
            }
            _depth--;

            if (closeParen < 0) HasErrors = true;
            int end = closeParen >= 0 ? closeParen + 1 : LastEnd;
            return new FunctionCallNode(start, end - start, name, nameLength, openParen, closeParen, args, separators, callee);
        }

        private SyntaxNode? ParseGroup(Token open)
        {
            if (_depth >= MaxDepth) return null;
            _depth++;
            _pos++;

            var inner = ParseExpr(1);
            if (Peek is not null && !PeekIs(TokenKind.CloseParen))
                inner = Recover(inner, t => t.Kind == TokenKind.CloseParen);
            else if (inner is null) HasErrors = true; // "()" is not valid Excel

            int closeParen = -1;
            if (Peek is { Kind: TokenKind.CloseParen } close)
            {
                closeParen = close.Start;
                _pos++;
            }
            else
            {
                HasErrors = true;
            }
            _depth--;

            int end = closeParen >= 0 ? closeParen + 1 : LastEnd;
            return new ParenNode(open.Start, end - open.Start, closeParen, inner);
        }

        private SyntaxNode? ParseArray(Token open)
        {
            if (_depth >= MaxDepth) return null;
            _depth++;
            _pos++;

            var rows = new List<IReadOnlyList<SyntaxNode>>();
            var row = new List<SyntaxNode>();
            int closeBrace = -1;

            bool first = true;
            while (true)
            {
                var t = Peek;
                if (first && t is { Kind: TokenKind.ArrayClose })
                {
                    closeBrace = t.Start;
                    _pos++;
                    break;
                }
                if (t is null)
                {
                    if (!first) row.Add(new MissingArgumentNode(LastEnd));
                    break;
                }

                SyntaxNode element = t.Kind is TokenKind.ArrayColumnSeparator or TokenKind.ArrayRowSeparator or TokenKind.ArrayClose
                    ? new MissingArgumentNode(t.Start)
                    : ParseExpressionUntil(IsArrayTerminator);
                row.Add(element);
                first = false;

                t = Peek;
                if (t is null) break;
                _pos++;
                if (t.Kind == TokenKind.ArrayClose)
                {
                    closeBrace = t.Start;
                    break;
                }
                if (t.Kind == TokenKind.ArrayRowSeparator)
                {
                    rows.Add(row);
                    row = new List<SyntaxNode>();
                }
            }
            if (row.Count > 0) rows.Add(row);
            _depth--;

            if (closeBrace < 0) HasErrors = true;
            int end = closeBrace >= 0 ? closeBrace + 1 : LastEnd;
            return new ArrayConstantNode(open.Start, end - open.Start, rows);
        }

        private static bool IsArgumentTerminator(Token t) => t.Kind is TokenKind.ArgumentSeparator or TokenKind.CloseParen;

        private static bool IsArrayTerminator(Token t) =>
            t.Kind is TokenKind.ArrayColumnSeparator or TokenKind.ArrayRowSeparator or TokenKind.ArrayClose;

        /// <summary>
        /// Parses one list element (function argument, array element) and guarantees that the parser then stands
        /// on a terminator or at the end of the text: anything else is wrapped by <see cref="Recover"/>.
        /// Never returns null and always consumes at least one token.
        /// </summary>
        private SyntaxNode ParseExpressionUntil(Func<Token, bool> isTerminator)
        {
            var expr = ParseExpr(1);
            if (expr is not null && (Peek is null || isTerminator(Peek)))
                return expr;
            return Recover(expr, isTerminator);
        }

        // ------------------------------------------------------------------------------------------
        // Error recovery
        // ------------------------------------------------------------------------------------------

        /// <summary>
        /// Wraps <paramref name="parsed"/> (if any) and the tokens from the current one up to the next terminator at
        /// nesting level 0 into an <see cref="UnparsedNode"/>. The current token is always consumed (callers only call
        /// this when it cannot start or continue an expression and is not a terminator), which guarantees progress.
        /// Reference, name and structured-reference tokens inside the region become leaf children so coloring and
        /// links keep working.
        /// </summary>
        private UnparsedNode Recover(SyntaxNode? parsed, Func<Token, bool> isTerminator)
        {
            HasErrors = true;

            int start = parsed?.Start ?? Peek?.Start ?? LastEnd;
            int end = parsed?.End ?? start;
            var children = new List<SyntaxNode>();
            if (parsed is not null) children.Add(parsed);

            int nesting = 0;
            bool first = true;
            while (Peek is { } t)
            {
                if (!first && nesting == 0 && isTerminator(t)) break;
                first = false;

                _pos++;
                end = t.End;
                switch (t.Kind)
                {
                    case TokenKind.OpenParen:
                    case TokenKind.ArrayOpen:
                        nesting++;
                        break;
                    case TokenKind.CloseParen:
                    case TokenKind.ArrayClose:
                        if (nesting > 0) nesting--;
                        break;
                    case TokenKind.Reference:
                    case TokenKind.Name:
                    case TokenKind.StructuredReference:
                        children.Add(MakeLeaf(t));
                        break;
                }
            }

            var unparsed = new UnparsedNode(start, end - start);
            foreach (var child in children) unparsed.AddChild(child);
            return unparsed;
        }
    }
}
