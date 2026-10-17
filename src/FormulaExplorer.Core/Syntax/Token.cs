namespace FormulaExplorer.Core.Syntax;

/// <summary>Kinds of lexical tokens produced by <see cref="FormulaTokenizer"/>.</summary>
public enum TokenKind
{
    /// <summary>The leading '=' of a formula.</summary>
    FormulaStart,
    Number,
    /// <summary>A string literal including its quotes; inner quotes are doubled ("a""b").</summary>
    String,
    /// <summary>TRUE or FALSE (case-insensitive) when not followed by '('.</summary>
    Bool,
    /// <summary>#DIV/0!, #N/A, #NAME?, #NULL!, #NUM!, #REF!, #VALUE!, #SPILL!, #CALC!, #GETTING_DATA, ...</summary>
    Error,
    /// <summary>
    /// A cell/range reference, including any sheet/workbook prefix, e.g. K$31, I$19:$I$19, A:A, 1:1,
    /// Sheet8!E7, 'Inventory Forcast'!C$6, Sheet1:Sheet3!A1, [Book.xlsx]Sheet!A1, Sheet1!#REF!.
    /// A range written as two cell refs joined by ':' (A1:B2) is ONE token.
    /// </summary>
    Reference,
    /// <summary>A defined name (optionally prefixed by a sheet or workbook, e.g. Sheet1!MyName).</summary>
    Name,
    /// <summary>A structured reference, e.g. Table1[Col], Table1[[#Headers],[Col]], [@Col].</summary>
    StructuredReference,
    /// <summary>A function name immediately followed by '(' (the '(' is a separate OpenParen token).</summary>
    Function,
    /// <summary>A binary operator: + - * / ^ &amp; = &lt;&gt; &lt; &gt; &lt;= &gt;= and the range operator ':' between non-trivial operands.</summary>
    BinaryOperator,
    /// <summary>A prefix operator: unary - or + or the implicit-intersection '@'.</summary>
    PrefixOperator,
    /// <summary>A postfix operator: '%' or the spill operator '#'.</summary>
    PostfixOperator,
    /// <summary>Whitespace that acts as the intersection operator (between two reference operands).</summary>
    IntersectionOperator,
    /// <summary>',' that acts as the union operator (inside parentheses, not separating function arguments).</summary>
    UnionOperator,
    /// <summary>',' separating function arguments.</summary>
    ArgumentSeparator,
    OpenParen,
    CloseParen,
    ArrayOpen,
    ArrayClose,
    /// <summary>',' between columns of an array constant.</summary>
    ArrayColumnSeparator,
    /// <summary>';' between rows of an array constant.</summary>
    ArrayRowSeparator,
    /// <summary>Insignificant whitespace, including line breaks (\n, \r\n).</summary>
    Whitespace,
    /// <summary>Anything the tokenizer could not classify.</summary>
    Unknown,
}

/// <summary>A token with its exact span in the formula text.</summary>
public sealed record Token(TokenKind Kind, int Start, int Length, string Text)
{
    public int End => Start + Length;
}
