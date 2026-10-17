namespace FormulaExplorer.Core.Syntax;

/// <summary>
/// Splits an Excel formula (English, A1 style, as returned by Range.Formula) into tokens with exact spans.
/// A hand-written state machine in the spirit of E. W. Bachtal's ExcelFormulaTokenizer, extended with exact spans,
/// complete references (sheet prefixes, 3D, external, areas, whole columns/rows), structured references and the
/// context rules for ',' (argument separator vs. union vs. array column separator) and whitespace (intersection).
/// </summary>
public static class FormulaTokenizer
{
    private static readonly string[] ErrorLiterals =
    {
        "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "#SPILL!",
        "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#UNKNOWN!", "#BUSY!",
    };

    /// <summary>
    /// Tokenizes <paramref name="text"/>. Never throws. The concatenation of all token texts equals the input.
    /// If <paramref name="text"/> does not start with '=', returns a single Unknown token covering everything (or nothing for "").
    /// </summary>
    public static IReadOnlyList<Token> Tokenize(string text)
    {
        if (string.IsNullOrEmpty(text)) return Array.Empty<Token>();
        if (text[0] != '=') return new[] { new Token(TokenKind.Unknown, 0, text.Length, text) };
        return new Scanner(text).Run();
    }

    /// <summary>Kind of the innermost open bracket, used to decide what ',' and ';' mean.</summary>
    private enum Context { Function, Group, Array }

    private sealed class Scanner
    {
        private readonly string _s;
        private readonly int _n;
        private readonly List<Token> _tokens = new();
        private readonly List<Context> _stack = new();

        /// <summary>Kind of the last non-whitespace token (decides operand position vs. operand end).</summary>
        private TokenKind _prev = TokenKind.FormulaStart;

        /// <summary>Index in <see cref="_tokens"/> of the last ')' that closed a function call (-1 if none).
        /// A '(' glued to it is an invocation of the call's result (<c>LAMBDA(x,x)(2)</c>), not a plain group.</summary>
        private int _callCloseIndex = -1;

        public Scanner(string s)
        {
            _s = s;
            _n = s.Length;
        }

        public IReadOnlyList<Token> Run()
        {
            Add(TokenKind.FormulaStart, 0, 1);
            int i = 1;
            while (i < _n) i = Step(i);
            return _tokens;
        }

        private void Add(TokenKind kind, int start, int end)
        {
            _tokens.Add(new Token(kind, start, end - start, _s.Substring(start, end - start)));
            if (kind != TokenKind.Whitespace) _prev = kind;
        }

        /// <summary>True when the previous significant token ends an operand (so +/- are binary and % is postfix).</summary>
        private bool PrevEndsOperand() => _prev is TokenKind.Number or TokenKind.String or TokenKind.Bool
            or TokenKind.Error or TokenKind.Reference or TokenKind.Name or TokenKind.StructuredReference
            or TokenKind.CloseParen or TokenKind.ArrayClose or TokenKind.PostfixOperator;

        /// <summary>True when the previous token can be the left side of an intersection (a reference-valued operand).</summary>
        private bool PrevIsReferenceOperand()
        {
            var last = _tokens[^1];
            return last.Kind is TokenKind.Reference or TokenKind.Name or TokenKind.StructuredReference or TokenKind.CloseParen
                || (last.Kind == TokenKind.PostfixOperator && last.Text == "#")
                || (last.Kind == TokenKind.Error && IsRefError(last.Start)); // A1:A3 #REF! (a deleted range)
        }

        private static bool IsWhitespace(char c) => c is ' ' or '\t' or '\r' or '\n' || char.IsWhiteSpace(c);

        /// <summary>Scans one token at <paramref name="i"/> and returns the index after it. Always advances.</summary>
        private int Step(int i)
        {
            char c = _s[i];
            if (IsWhitespace(c)) return ScanWhitespace(i);

            switch (c)
            {
                case '"':
                    return ScanString(i);
                case '#':
                    return ScanHash(i);
                case '(':
                    {
                        bool invocation = _prev == TokenKind.CloseParen && _tokens.Count - 1 == _callCloseIndex;
                        _stack.Add(_prev == TokenKind.Function || invocation ? Context.Function : Context.Group);
                    }
                    Add(TokenKind.OpenParen, i, i + 1);
                    return i + 1;
                case ')':
                    if (_stack.Count > 0 && _stack[^1] != Context.Array)
                    {
                        if (_stack[^1] == Context.Function) _callCloseIndex = _tokens.Count;
                        _stack.RemoveAt(_stack.Count - 1);
                    }
                    Add(TokenKind.CloseParen, i, i + 1);
                    return i + 1;
                case '{':
                    _stack.Add(Context.Array);
                    Add(TokenKind.ArrayOpen, i, i + 1);
                    return i + 1;
                case '}':
                    if (_stack.Count > 0 && _stack[^1] == Context.Array) _stack.RemoveAt(_stack.Count - 1);
                    Add(TokenKind.ArrayClose, i, i + 1);
                    return i + 1;
                case ',':
                    Add(CommaKind(), i, i + 1);
                    return i + 1;
                case ';':
                    Add(_stack.Count > 0 && _stack[^1] == Context.Array ? TokenKind.ArrayRowSeparator : TokenKind.Unknown, i, i + 1);
                    return i + 1;
                case '+':
                case '-':
                    Add(PrevEndsOperand() ? TokenKind.BinaryOperator : TokenKind.PrefixOperator, i, i + 1);
                    return i + 1;
                case '@':
                    Add(PrevEndsOperand() ? TokenKind.Unknown : TokenKind.PrefixOperator, i, i + 1);
                    return i + 1;
                case '%':
                    Add(PrevEndsOperand() ? TokenKind.PostfixOperator : TokenKind.Unknown, i, i + 1);
                    return i + 1;
                case '*':
                case '/':
                case '^':
                case '&':
                case '=':
                case ':':
                    Add(TokenKind.BinaryOperator, i, i + 1);
                    return i + 1;
                case '<':
                    {
                        int e = i + 1 < _n && _s[i + 1] is '=' or '>' ? i + 2 : i + 1;
                        Add(TokenKind.BinaryOperator, i, e);
                        return e;
                    }
                case '>':
                    {
                        int e = i + 1 < _n && _s[i + 1] == '=' ? i + 2 : i + 1;
                        Add(TokenKind.BinaryOperator, i, e);
                        return e;
                    }
            }

            if (TryScanOperand(i, out var kind, out int end))
            {
                Add(kind, i, end);
                return end;
            }

            Add(TokenKind.Unknown, i, i + 1);
            return i + 1;
        }

        private TokenKind CommaKind()
        {
            if (_stack.Count == 0) return TokenKind.UnionOperator;
            return _stack[^1] switch
            {
                Context.Function => TokenKind.ArgumentSeparator,
                Context.Array => TokenKind.ArrayColumnSeparator,
                _ => TokenKind.UnionOperator,
            };
        }

        private int ScanWhitespace(int i)
        {
            int j = i;
            while (j < _n && IsWhitespace(_s[j])) j++;
            bool intersection = j < _n && PrevIsReferenceOperand() && StartsReferenceOperand(j);
            Add(intersection ? TokenKind.IntersectionOperator : TokenKind.Whitespace, i, j);
            return j;
        }

        /// <summary>True when a reference-valued operand (reference, name, structured reference, function call, group) starts at <paramref name="j"/>.</summary>
        private bool StartsReferenceOperand(int j)
        {
            if (_s[j] == '(') return true;
            if (_s[j] == '#') return IsRefError(j);
            return TryScanOperand(j, out var kind, out _)
                && kind is TokenKind.Reference or TokenKind.Name or TokenKind.StructuredReference or TokenKind.Function;
        }

        /// <summary>True when the error literal <c>#REF!</c> (any case) starts at <paramref name="i"/>.</summary>
        private bool IsRefError(int i) => string.Compare(_s, i, "#REF!", 0, 5, StringComparison.OrdinalIgnoreCase) == 0;

        private int ScanString(int i)
        {
            int j = i + 1;
            while (j < _n)
            {
                if (_s[j] == '"')
                {
                    if (j + 1 < _n && _s[j + 1] == '"') { j += 2; continue; }
                    j++;
                    break;
                }
                j++;
            }
            Add(TokenKind.String, i, j);
            return j;
        }

        /// <summary>'#': an error literal, the spill operator after a reference, or unknown.</summary>
        private int ScanHash(int i)
        {
            foreach (var literal in ErrorLiterals)
            {
                if (string.Compare(_s, i, literal, 0, literal.Length, StringComparison.OrdinalIgnoreCase) != 0) continue;
                // A deleted sheet turns Sheet2!A1 into #REF!A1: one dead reference, not an error plus a live A1.
                if (literal == "#REF!" && TryScanDeletedSheetTail(i + literal.Length, out int refEnd))
                {
                    Add(TokenKind.Reference, i, refEnd);
                    return refEnd;
                }
                Add(TokenKind.Error, i, i + literal.Length);
                return i + literal.Length;
            }

            var last = _tokens[^1];
            bool spill = last.End == i && last.Kind is TokenKind.Reference or TokenKind.Name or TokenKind.StructuredReference or TokenKind.CloseParen;
            Add(spill ? TokenKind.PostfixOperator : TokenKind.Unknown, i, i + 1);
            return i + 1;
        }

        /// <summary>The address or name glued to a <c>#REF!</c> sheet prefix (<c>#REF!A1:B2</c>, <c>#REF!Name</c>).</summary>
        private bool TryScanDeletedSheetTail(int j, out int end)
        {
            end = j;
            if (j >= _n) return false;
            if (ReferenceSyntax.TryScanAddress(_s, j, out end)) return true;
            if (!ReferenceSyntax.IsIdentStart(_s[j])) return false;
            int k = ReferenceSyntax.ScanIdent(_s, j);
            if (k < _n && _s[k] == '(') return false; // a call, not a name
            end = k;
            return true;
        }

        /// <summary>
        /// Scans an operand token (number, reference, name, structured reference, function name or boolean)
        /// that starts at <paramref name="i"/>. Pure lookahead: does not add tokens.
        /// </summary>
        private bool TryScanOperand(int i, out TokenKind kind, out int end)
        {
            kind = TokenKind.Unknown;
            end = i;
            char c = _s[i];

            if (c == '$')
            {
                if (!ReferenceSyntax.TryScanAddress(_s, i, out end)) return false;
                kind = TokenKind.Reference;
                return true;
            }

            if (c >= '0' && c <= '9')
            {
                if (ReferenceSyntax.TryScanAddress(_s, i, out end)) // 1:1, $3:$5
                {
                    kind = TokenKind.Reference;
                    return true;
                }
                end = ScanNumber(i);
                kind = TokenKind.Number;
                return true;
            }

            if (c == '.')
            {
                if (i + 1 >= _n || !(_s[i + 1] >= '0' && _s[i + 1] <= '9')) return false;
                end = ScanNumber(i);
                kind = TokenKind.Number;
                return true;
            }

            if (c == '\'') return TryScanQualified(i, out kind, out end);

            if (c == '[')
            {
                if (TryScanQualified(i, out kind, out end)) return true; // [Book.xlsx]Sheet1!A1
                if (!ReferenceSyntax.TryScanBrackets(_s, i, out end)) return false;
                kind = TokenKind.StructuredReference; // [@Col]
                return true;
            }

            if (!ReferenceSyntax.IsIdentStart(c)) return false;

            if (TryScanQualified(i, out kind, out end)) return true;

            int k = ReferenceSyntax.ScanIdent(_s, i);
            if (k < _n && _s[k] == '(')
            {
                kind = TokenKind.Function;
                end = k;
                return true;
            }
            if (k < _n && _s[k] == '[' && ReferenceSyntax.TryScanBrackets(_s, k, out int bracketEnd))
            {
                kind = TokenKind.StructuredReference;
                end = bracketEnd;
                return true;
            }
            if (ReferenceSyntax.TryScanAddress(_s, i, out int addressEnd))
            {
                kind = TokenKind.Reference;
                end = addressEnd;
                return true;
            }

            end = k;
            int length = k - i;
            bool isBool = (length == 4 && string.Compare(_s, i, "TRUE", 0, 4, StringComparison.OrdinalIgnoreCase) == 0)
                || (length == 5 && string.Compare(_s, i, "FALSE", 0, 5, StringComparison.OrdinalIgnoreCase) == 0);
            kind = isBool ? TokenKind.Bool : TokenKind.Name;
            return true;
        }

        /// <summary>A sheet/workbook prefix followed by an address, #REF!, a name or a structured reference.</summary>
        private bool TryScanQualified(int i, out TokenKind kind, out int end)
        {
            kind = TokenKind.Unknown;
            end = i;
            if (!ReferenceSyntax.TryScanPrefix(_s, i, out var prefix)) return false;

            int j = prefix.End;
            if (j >= _n) return false;

            if (ReferenceSyntax.TryScanAddress(_s, j, out end))
            {
                kind = TokenKind.Reference;
                return true;
            }
            if (string.Compare(_s, j, "#REF!", 0, 5, StringComparison.OrdinalIgnoreCase) == 0)
            {
                kind = TokenKind.Reference;
                end = j + 5;
                return true;
            }
            if (ReferenceSyntax.IsIdentStart(_s[j]))
            {
                int k = ReferenceSyntax.ScanIdent(_s, j);
                if (k < _n && _s[k] == '(') // PERSONAL.XLSB!Fn(...), Sheet1!MyLambda(...): the prefix belongs to the call
                {
                    kind = TokenKind.Function;
                    end = k;
                    return true;
                }
                if (k < _n && _s[k] == '[' && ReferenceSyntax.TryScanBrackets(_s, k, out int bracketEnd))
                {
                    kind = TokenKind.StructuredReference;
                    end = bracketEnd;
                    return true;
                }
                kind = TokenKind.Name;
                end = k;
                return true;
            }
            return false;
        }

        /// <summary>digits [. digits] [E [+-] digits], or . digits [E ...]. The exponent only counts when digits follow.</summary>
        private int ScanNumber(int i)
        {
            int j = i;
            while (j < _n && _s[j] >= '0' && _s[j] <= '9') j++;
            if (j < _n && _s[j] == '.')
            {
                j++;
                while (j < _n && _s[j] >= '0' && _s[j] <= '9') j++;
            }
            if (j < _n && _s[j] is 'e' or 'E')
            {
                int k = j + 1;
                if (k < _n && _s[k] is '+' or '-') k++;
                if (k < _n && _s[k] >= '0' && _s[k] <= '9')
                {
                    while (k < _n && _s[k] >= '0' && _s[k] <= '9') k++;
                    j = k;
                }
            }
            return j;
        }
    }
}
