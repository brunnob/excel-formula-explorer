namespace FormulaExplorer.Core.Syntax;

public enum NodeKind
{
    /// <summary>The whole formula, including the leading '='. Span = entire text.</summary>
    Formula,
    Literal,
    ArrayConstant,
    Reference,
    Name,
    StructuredReference,
    FunctionCall,
    Binary,
    Prefix,
    Postfix,
    Paren,
    MissingArgument,
    /// <summary>A span the parser could not understand (error recovery). Never highlighted.</summary>
    Unparsed,
}

public enum LiteralKind { Number, String, Bool, Error }

/// <summary>
/// Base class of the syntax tree. Every node knows its exact character span in the displayed text
/// (which is exactly <c>Range.Formula</c>). Reference equality only.
/// </summary>
public abstract class SyntaxNode
{
    private readonly List<SyntaxNode> _children = new();

    protected SyntaxNode(int start, int length)
    {
        Start = start;
        Length = length;
    }

    public abstract NodeKind Kind { get; }

    /// <summary>Index of the first character of the node in the formula text.</summary>
    public int Start { get; }

    /// <summary>Number of characters covered by the node.</summary>
    public int Length { get; }

    public int End => Start + Length;

    public SyntaxNode? Parent { get; private set; }

    public IReadOnlyList<SyntaxNode> Children => _children;

    /// <summary>The source text covered by this node.</summary>
    public string GetText(string formulaText) => formulaText.Substring(Start, Length);

    /// <summary>Adds a child and sets its <see cref="Parent"/>. For use by the parser.</summary>
    protected internal void AddChild(SyntaxNode child)
    {
        child.Parent = this;
        _children.Add(child);
    }

    /// <summary>Pre-order enumeration of this node and all descendants.</summary>
    /// <remarks>Iterative (explicit stack): a left-deep chain such as =A1+A2+...+A4000 is thousands of levels deep,
    /// and Excel's main thread must not risk a stack overflow.</remarks>
    public IEnumerable<SyntaxNode> DescendantsAndSelf()
    {
        var stack = new Stack<SyntaxNode>();
        stack.Push(this);
        while (stack.Count > 0)
        {
            var node = stack.Pop();
            yield return node;
            for (int i = node._children.Count - 1; i >= 0; i--)
                stack.Push(node._children[i]);
        }
    }

    public override string ToString() => $"{Kind}[{Start},{Length}]";
}

/// <summary>Root node. Span covers the whole formula text including '='.</summary>
public sealed class FormulaNode : SyntaxNode
{
    public FormulaNode(int length, SyntaxNode? body) : base(0, length)
    {
        Body = body;
        if (body is not null) AddChild(body);
    }

    public override NodeKind Kind => NodeKind.Formula;

    /// <summary>The expression after '='. Null when the formula is just "=" or could not be parsed at all.</summary>
    public SyntaxNode? Body { get; }
}

public sealed class LiteralNode : SyntaxNode
{
    public LiteralNode(int start, int length, LiteralKind literalKind) : base(start, length) => LiteralKind = literalKind;

    public override NodeKind Kind => NodeKind.Literal;

    public LiteralKind LiteralKind { get; }
}

/// <summary>An array constant such as {1,2;3,4}. Span includes the braces.</summary>
public sealed class ArrayConstantNode : SyntaxNode
{
    public ArrayConstantNode(int start, int length, IReadOnlyList<IReadOnlyList<SyntaxNode>> rows) : base(start, length)
    {
        Rows = rows;
        foreach (var row in rows)
            foreach (var item in row)
                AddChild(item);
    }

    public override NodeKind Kind => NodeKind.ArrayConstant;

    public IReadOnlyList<IReadOnlyList<SyntaxNode>> Rows { get; }
}

public enum ReferenceForm
{
    /// <summary>A1, $A$1</summary>
    Cell,
    /// <summary>A1:B2 (two cells joined by ':' inside one token)</summary>
    Area,
    /// <summary>A:A, $A:$C</summary>
    Columns,
    /// <summary>1:1, $3:$5</summary>
    Rows,
    /// <summary>#REF! (optionally with a sheet prefix)</summary>
    RefError,
}

/// <summary>
/// A cell/range reference token, e.g. <c>'Inventory Forcast'!C$6</c>. The span includes the sheet/workbook prefix.
/// </summary>
public sealed class ReferenceNode : SyntaxNode
{
    public ReferenceNode(int start, int length, ReferenceForm form, string address,
        string? workbook, string? sheet, string? lastSheet) : base(start, length)
    {
        Form = form;
        Address = address;
        Workbook = workbook;
        Sheet = sheet;
        LastSheet = lastSheet;
    }

    public override NodeKind Kind => NodeKind.Reference;

    public ReferenceForm Form { get; }

    /// <summary>The address part as written, without prefix, e.g. "I$19:$I$19".</summary>
    public string Address { get; }

    /// <summary>Workbook name from an external prefix ([Book.xlsx]), unquoted, without brackets; else null.</summary>
    public string? Workbook { get; }

    /// <summary>Sheet name, unquoted and with '' unescaped; null when the reference has no sheet prefix.</summary>
    public string? Sheet { get; }

    /// <summary>For 3D references (Sheet1:Sheet3!A1), the last sheet name; else null.</summary>
    public string? LastSheet { get; }

    public bool Is3D => LastSheet is not null;
}

/// <summary>A defined name, e.g. <c>Hundred</c> or <c>Sheet1!LocalName</c>. Span includes any prefix.</summary>
public sealed class NameNode : SyntaxNode
{
    public NameNode(int start, int length, string name, string? workbook, string? sheet) : base(start, length)
    {
        Name = name;
        Workbook = workbook;
        Sheet = sheet;
    }

    public override NodeKind Kind => NodeKind.Name;

    /// <summary>The name without prefix, as written.</summary>
    public string Name { get; }

    public string? Workbook { get; }

    public string? Sheet { get; }

    /// <summary>
    /// True when this is a variable bound by an enclosing LET/LAMBDA (declaration or use), not a workbook name:
    /// it must never be evaluated, colored or linked as if it were the defined name of the same spelling.
    /// </summary>
    public bool IsLocal { get; internal set; }
}

/// <summary>A structured (table) reference, e.g. <c>Table1[Col]</c> or <c>[@Col]</c>.</summary>
public sealed class StructuredReferenceNode : SyntaxNode
{
    public StructuredReferenceNode(int start, int length, string? table, string specifier) : base(start, length)
    {
        Table = table;
        Specifier = specifier;
    }

    public override NodeKind Kind => NodeKind.StructuredReference;

    /// <summary>Table name, or null for the table-less form ([@Col]).</summary>
    public string? Table { get; }

    /// <summary>Everything from the first '[' to the matching ']', inclusive.</summary>
    public string Specifier { get; }
}

/// <summary>
/// A function call <c>NAME(arg, ...)</c>. Span runs from the first char of the name to the closing ')'
/// (or to the end of the last argument if the ')' is missing).
/// </summary>
public sealed class FunctionCallNode : SyntaxNode
{
    public FunctionCallNode(int start, int length, string name, int nameLength, int openParen, int closeParen,
        IReadOnlyList<SyntaxNode> arguments, IReadOnlyList<int> separators, SyntaxNode? callee = null) : base(start, length)
    {
        Callee = callee;
        if (callee is not null) AddChild(callee);
        Name = name;
        NameLength = nameLength;
        OpenParen = openParen;
        CloseParen = closeParen;
        Arguments = arguments;
        Separators = separators;
        foreach (var a in arguments) AddChild(a);
    }

    public override NodeKind Kind => NodeKind.FunctionCall;

    /// <summary>Function name as written (may include a prefix such as _xlfn., or a workbook/sheet qualifier such as PERSONAL.XLSB!).</summary>
    public string Name { get; }

    /// <summary>Length of the name (the name starts at <see cref="SyntaxNode.Start"/>). For an invocation, the length of the callee.</summary>
    public int NameLength { get; }

    /// <summary>
    /// For an immediate invocation such as <c>LAMBDA(x,y,x+y)(1,2)</c>: the call whose result is invoked (the first child;
    /// <see cref="Name"/> is its text and the node spans callee..')'). Null for an ordinary named call.
    /// </summary>
    public SyntaxNode? Callee { get; }

    /// <summary>Index of '('.</summary>
    public int OpenParen { get; }

    /// <summary>Index of ')', or -1 if missing.</summary>
    public int CloseParen { get; }

    /// <summary>Arguments in order. An empty argument is a <see cref="MissingArgumentNode"/>. No arguments => empty list.</summary>
    public IReadOnlyList<SyntaxNode> Arguments { get; }

    /// <summary>Indexes of the ',' argument separators.</summary>
    public IReadOnlyList<int> Separators { get; }
}

/// <summary>
/// A binary operation. Span runs from the start of <see cref="Left"/> to the end of <see cref="Right"/>.
/// Operator is one of + - * / ^ &amp; = &lt;&gt; &lt; &gt; &lt;= &gt;= : (range) " " (intersection) "," (union).
/// </summary>
public sealed class BinaryNode : SyntaxNode
{
    public BinaryNode(int start, int length, string op, int operatorStart, int operatorLength, SyntaxNode left, SyntaxNode right)
        : base(start, length)
    {
        Operator = op;
        OperatorStart = operatorStart;
        OperatorLength = operatorLength;
        Left = left;
        Right = right;
        AddChild(left);
        AddChild(right);
    }

    public override NodeKind Kind => NodeKind.Binary;

    /// <summary>The operator text. For intersection the operator is a single space " " even if several spaces were written.</summary>
    public string Operator { get; }

    public int OperatorStart { get; }

    public int OperatorLength { get; }

    public SyntaxNode Left { get; }

    public SyntaxNode Right { get; }

    /// <summary>True for the reference operators ':' ' ' ','.</summary>
    public bool IsReferenceOperator => Operator is ":" or " " or ",";
}

/// <summary>A prefix operation: unary '-', unary '+', or implicit intersection '@'. Span = operator + operand.</summary>
public sealed class PrefixNode : SyntaxNode
{
    public PrefixNode(int start, int length, string op, SyntaxNode operand) : base(start, length)
    {
        Operator = op;
        Operand = operand;
        AddChild(operand);
    }

    public override NodeKind Kind => NodeKind.Prefix;

    public string Operator { get; }

    /// <summary>The operator is at <see cref="SyntaxNode.Start"/> and is 1 char long.</summary>
    public int OperatorStart => Start;

    public SyntaxNode Operand { get; }
}

/// <summary>A postfix operation: percent '%' or spill '#'. Span = operand + operator.</summary>
public sealed class PostfixNode : SyntaxNode
{
    public PostfixNode(int start, int length, string op, int operatorStart, SyntaxNode operand) : base(start, length)
    {
        Operator = op;
        OperatorStart = operatorStart;
        Operand = operand;
        AddChild(operand);
    }

    public override NodeKind Kind => NodeKind.Postfix;

    public string Operator { get; }

    public int OperatorStart { get; }

    public SyntaxNode Operand { get; }
}

/// <summary>A parenthesized group <c>( inner )</c>. Span includes both parentheses (or up to the end if ')' is missing).</summary>
public sealed class ParenNode : SyntaxNode
{
    public ParenNode(int start, int length, int closeParen, SyntaxNode? inner) : base(start, length)
    {
        CloseParen = closeParen;
        Inner = inner;
        if (inner is not null) AddChild(inner);
    }

    public override NodeKind Kind => NodeKind.Paren;

    /// <summary>Index of '(' (same as Start).</summary>
    public int OpenParen => Start;

    /// <summary>Index of ')', or -1 if missing.</summary>
    public int CloseParen { get; }

    /// <summary>The inner expression; null for "()".</summary>
    public SyntaxNode? Inner { get; }
}

/// <summary>An omitted function argument, e.g. the middle one in IF(A1,,2). Zero length at the position where it would be.</summary>
public sealed class MissingArgumentNode : SyntaxNode
{
    public MissingArgumentNode(int position) : base(position, 0) { }

    public override NodeKind Kind => NodeKind.MissingArgument;
}

/// <summary>Error-recovery node for text the parser could not understand.</summary>
public sealed class UnparsedNode : SyntaxNode
{
    public UnparsedNode(int start, int length) : base(start, length) { }

    public override NodeKind Kind => NodeKind.Unparsed;
}
