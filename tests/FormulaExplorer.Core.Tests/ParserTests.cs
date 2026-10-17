using FormulaExplorer.Core.Syntax;

namespace FormulaExplorer.Core.Tests;

public class ParserTests
{
    private static ParsedFormula Parse(string text) => FormulaParser.Parse(text);

    // ---- non-formulas ---------------------------------------------------------------------------

    [Theory]
    [InlineData("")]
    [InlineData("hello")]
    [InlineData("42")]
    [InlineData("A1+B1")]
    public void NonFormulaHasNoRoot(string text)
    {
        var parsed = Parse(text);
        Assert.False(parsed.IsFormula);
        Assert.Null(parsed.Root);
        Assert.False(parsed.HasErrors);
        Assert.Empty(parsed.Nodes);
        Assert.Equal(text, parsed.Text);
    }

    [Fact]
    public void NullTextIsTreatedAsEmpty()
    {
        var parsed = FormulaParser.Parse(null!);
        Assert.False(parsed.IsFormula);
        Assert.Equal("", parsed.Text);
    }

    [Theory]
    [InlineData("=")]
    [InlineData("=   ")]
    public void EqualsAloneHasRootWithoutBody(string text)
    {
        var parsed = Parse(text);
        Assert.True(parsed.IsFormula);
        Assert.NotNull(parsed.Root);
        Assert.Null(parsed.Root!.Body);
        Assert.False(parsed.HasErrors);
        Assert.Equal(0, parsed.Root.Start);
        Assert.Equal(text.Length, parsed.Root.Length);
    }

    // ---- precedence and associativity -------------------------------------------------------------

    [Theory]
    [InlineData("=-2^2", "(^ (pre- 2) 2)")]
    [InlineData("=2^3^2", "(^ (^ 2 3) 2)")]
    [InlineData("=1+2*3", "(+ 1 (* 2 3))")]
    [InlineData("=1-2-3", "(- (- 1 2) 3)")]
    [InlineData("=1*2/3", "(/ (* 1 2) 3)")]
    [InlineData("=a&b=c", "(= (& a b) c)")]
    [InlineData("=-A1%", "(post% (pre- A1))")]
    [InlineData("=2^-2", "(^ 2 (pre- 2))")]
    [InlineData("=--A1", "(pre- (pre- A1))")]
    [InlineData("=+A1", "(pre+ A1)")]
    [InlineData("=@A1", "(pre@ A1)")]
    [InlineData("=1+2=3", "(= (+ 1 2) 3)")]
    [InlineData("=1&2+3", "(& 1 (+ 2 3))")]
    [InlineData("=1<2&3", "(< 1 (& 2 3))")]
    [InlineData("=1=2<3", "(< (= 1 2) 3)")]
    [InlineData("=1<>2>=3", "(>= (<> 1 2) 3)")]
    [InlineData("=2*3^2", "(* 2 (^ 3 2))")]
    [InlineData("=-A1^2", "(^ (pre- A1) 2)")]
    [InlineData("=A1^2%", "(^ A1 (post% 2))")]
    [InlineData("=50%*2", "(* (post% 50) 2)")]
    [InlineData("=5%%", "(post% (post% 5))")]
    [InlineData("=1+2*3-4/5", "(- (+ 1 (* 2 3)) (/ 4 5))")]
    [InlineData("=2^3*4^5", "(* (^ 2 3) (^ 4 5))")]
    [InlineData("=A1*-B1", "(* A1 (pre- B1))")]
    [InlineData("=A1--B1", "(- A1 (pre- B1))")]
    [InlineData("=-(1+2)", "(pre- [(+ 1 2)])")]
    [InlineData("=(1+2)*3", "(* [(+ 1 2)] 3)")]
    [InlineData("=1 + 2 * 3", "(+ 1 (* 2 3))")]
    [InlineData("=1+\n2", "(+ 1 2)")]
    [InlineData("=A1&\"x\"&B1", "(& (& A1 \"x\") B1)")]
    public void Precedence(string formula, string expected) => Assert.Equal(expected, TestHelpers.Tree(formula));

    [Theory]
    [InlineData("=A1:B2 C3:D4", "(isect A1:B2 C3:D4)")]
    [InlineData("=A1:INDEX(B:B,3)", "(: A1 INDEX(B:B;3))")]
    [InlineData("=(A1,B2)", "[(union A1 B2)]")]
    [InlineData("=A1:B1 B1:C1,D1", "(union (isect A1:B1 B1:C1) D1)")]
    [InlineData("=Name1:Name2 C1", "(isect (: Name1 Name2) C1)")]
    [InlineData("=A1 B1 C1", "(isect (isect A1 B1) C1)")]
    [InlineData("=-Name1:Name2", "(pre- (: Name1 Name2))")]
    [InlineData("=A1+Name1:Name2", "(+ A1 (: Name1 Name2))")]
    [InlineData("=Name1:Name2:Name3", "(: (: Name1 Name2) Name3)")]
    [InlineData("=(A1):B2", "(: [A1] B2)")]
    [InlineData("=Sheet1!A1:Sheet1!B2", "(: Sheet1!A1 Sheet1!B2)")]
    [InlineData("=SUM((A1,B1),C1)", "SUM([(union A1 B1)];C1)")]
    [InlineData("=A1,B1+1", "(+ (union A1 B1) 1)")]
    [InlineData("=A1# B1", "(isect (post# A1) B1)")]
    public void ReferenceOperators(string formula, string expected) => Assert.Equal(expected, TestHelpers.Tree(formula));

    [Theory]
    [InlineData("=1/(((($I27/Hundred)+1)^(K$31)))", "(/ 1 [[(^ [(+ [(/ $I27 Hundred)] 1)] [K$31])]])")]
    [InlineData("=IF(K3=\"On\",1/(((-($I26/Hundred)+1)^(K$31))),0)", "IF((= K3 \"On\");(/ 1 [[(^ [(+ (pre- [(/ $I26 Hundred)]) 1)] [K$31])]]);0)")]
    [InlineData("=Sheet8!E7+(E30/C4)^(D4/D5)+D14-E21*E16+Summary!C5",
        "(+ (- (+ (+ Sheet8!E7 (^ [(/ E30 C4)] [(/ D4 D5)])) D14) (* E21 E16)) Summary!C5)")]
    [InlineData("=SUM(I$19:$I$19)+J$19*$I$28/Hundred-SUM(OFFSET($I$19,0,0,1,Aqui_Per))",
        "(- (+ SUM(I$19:$I$19) (/ (* J$19 $I$28) Hundred)) SUM(OFFSET($I$19;0;0;1;Aqui_Per)))")]
    public void VideoFormulasHaveExcelsTree(string formula, string expected) => Assert.Equal(expected, TestHelpers.Tree(formula));

    // ---- functions, arguments, arrays -----------------------------------------------------------------

    [Theory]
    [InlineData("=SUM(A1,B1)", "SUM(A1;B1)")]
    [InlineData("=SUM()", "SUM()")]
    [InlineData("=SUM( )", "SUM()")]
    [InlineData("=IF(A1,,2)", "IF(A1;_;2)")]
    [InlineData("=SUM(A1,)", "SUM(A1;_)")]
    [InlineData("=F(,)", "F(_;_)")]
    [InlineData("=IF(,1)", "IF(_;1)")]
    [InlineData("=IF(A1, ,2)", "IF(A1;_;2)")]
    [InlineData("=SUM(1,2,3,4)", "SUM(1;2;3;4)")]
    [InlineData("=SUM(SUM(1,2),3)", "SUM(SUM(1;2);3)")]
    [InlineData("=SUM({1,2;3,4})", "SUM({1,2;3,4})")]
    [InlineData("={-1,2}", "{(pre- 1),2}")]
    [InlineData("={\"a\",TRUE,#N/A}", "{\"a\",TRUE,#N/A}")]
    [InlineData("={1,2,3}", "{1,2,3}")]
    [InlineData("={1;2;3}", "{1;2;3}")]
    [InlineData("=NOT(TRUE)", "NOT(TRUE)")]
    [InlineData("=SUM(Table1[Col],[@Col])", "SUM(Table1[Col];[@Col])")]
    [InlineData("=LOG10(100)+_xlfn.XLOOKUP(1,A:A,B:B)", "(+ LOG10(100) _xlfn.XLOOKUP(1;A:A;B:B))")]
    [InlineData("=IF(A1>0,\"pos\",IF(A1<0,\"neg\",\"zero\"))", "IF((> A1 0);\"pos\";IF((< A1 0);\"neg\";\"zero\"))")]
    public void FunctionsAndArrays(string formula, string expected) => Assert.Equal(expected, TestHelpers.Tree(formula));

    [Fact]
    public void FunctionCallRecordsPositions()
    {
        const string text = "=IF(A1,,SUM(B1))";
        var call = (FunctionCallNode)Parse(text).Root!.Body!;
        Assert.Equal("IF", call.Name);
        Assert.Equal(1, call.Start);
        Assert.Equal(2, call.NameLength);
        Assert.Equal(3, call.OpenParen);
        Assert.Equal(text.Length - 1, call.CloseParen);
        Assert.Equal(text.Length, call.End);
        Assert.Equal(new[] { 6, 7 }, call.Separators);
        Assert.Equal(3, call.Arguments.Count);
        var missing = Assert.IsType<MissingArgumentNode>(call.Arguments[1]);
        Assert.Equal(7, missing.Start);
        Assert.Equal(0, missing.Length);
    }

    [Fact]
    public void MissingArgumentAtTheEndIsZeroLengthBeforeTheParen()
    {
        const string text = "=SUM(A1,)";
        var call = (FunctionCallNode)Parse(text).Root!.Body!;
        var missing = Assert.IsType<MissingArgumentNode>(call.Arguments[1]);
        Assert.Equal(8, missing.Start);
        Assert.Equal(0, missing.Length);
        Assert.False(Parse(text).HasErrors);
    }

    [Fact]
    public void FunctionWithoutArgumentsHasEmptyList()
    {
        var call = (FunctionCallNode)Parse("=NOW()").Root!.Body!;
        Assert.Empty(call.Arguments);
        Assert.Empty(call.Separators);
        Assert.Equal(5, call.CloseParen);
    }

    [Fact]
    public void ArrayConstantKeepsRows()
    {
        var array = (ArrayConstantNode)Parse("={1,2;3,4}").Root!.Body!;
        Assert.Equal(2, array.Rows.Count);
        Assert.All(array.Rows, row => Assert.Equal(2, row.Count));
        Assert.Equal(1, array.Start);
        Assert.Equal(9, array.Length);
        Assert.Equal(4, array.Children.Count);
    }

    // ---- node spans --------------------------------------------------------------------------------------

    [Theory]
    [InlineData("=-2^2", "Binary[-2^2] Prefix[-2] Literal[2] Literal[2]")]
    [InlineData("=A1%", "Postfix[A1%] Reference[A1]")]
    [InlineData("=(1+2)", "Paren[(1+2)] Binary[1+2] Literal[1] Literal[2]")]
    [InlineData("=A1 B1", "Binary[A1 B1] Reference[A1] Reference[B1]")]
    [InlineData("=A1   B1", "Binary[A1   B1] Reference[A1] Reference[B1]")]
    [InlineData("=SUM(A1,B1)", "FunctionCall[SUM(A1,B1)] Reference[A1] Reference[B1]")]
    [InlineData("=IF(A1,,2)", "FunctionCall[IF(A1,,2)] Reference[A1] MissingArgument[] Literal[2]")]
    [InlineData("={1,2;3,4}", "ArrayConstant[{1,2;3,4}] Literal[1] Literal[2] Literal[3] Literal[4]")]
    [InlineData("= A1 + B1 ", "Binary[A1 + B1] Reference[A1] Reference[B1]")]
    [InlineData("=@A1:A3", "Prefix[@A1:A3] Reference[A1:A3]")]
    [InlineData("=Sheet1!A1+'My Sheet'!B2", "Binary[Sheet1!A1+'My Sheet'!B2] Reference[Sheet1!A1] Reference['My Sheet'!B2]")]
    [InlineData("=-(A1)", "Prefix[-(A1)] Paren[(A1)] Reference[A1]")]
    public void NodeSpans(string formula, string expected) =>
        Assert.Equal(expected, TestHelpers.Describe(Parse(formula).Nodes.Skip(1), formula));

    [Fact]
    public void RootCoversWholeTextIncludingEquals()
    {
        var parsed = Parse("=A1 + B1 ");
        Assert.Equal(0, parsed.Root!.Start);
        Assert.Equal(9, parsed.Root.Length);
        Assert.Same(parsed.Root, parsed.Nodes[0]);
    }

    [Fact]
    public void BinaryNodeRecordsOperator()
    {
        var binary = (BinaryNode)Parse("=A1   *  B1").Root!.Body!;
        Assert.Equal("*", binary.Operator);
        Assert.Equal(6, binary.OperatorStart);
        Assert.Equal(1, binary.OperatorLength);
        Assert.False(binary.IsReferenceOperator);

        var isect = (BinaryNode)Parse("=A1   B1").Root!.Body!;
        Assert.Equal(" ", isect.Operator);
        Assert.Equal(3, isect.OperatorStart);
        Assert.Equal(3, isect.OperatorLength);
        Assert.True(isect.IsReferenceOperator);

        Assert.True(((BinaryNode)Parse("=Name1:Name2").Root!.Body!).IsReferenceOperator);
        Assert.True(((BinaryNode)Parse("=(A1,B1)").Root!.Body!.Children[0]).IsReferenceOperator);
    }

    [Fact]
    public void ParentLinksAreSet()
    {
        var parsed = Parse("=SUM(A1,2)");
        foreach (var node in parsed.Nodes.Skip(1))
            Assert.NotNull(node.Parent);
        Assert.Null(parsed.Root!.Parent);
        Assert.Same(parsed.Root, parsed.Root.Body!.Parent);
    }

    // ---- reference / name / structured-reference details -------------------------------------------------

    private static string RefInfo(string address)
    {
        var node = (ReferenceNode)Parse("=" + address).Root!.Body!;
        return $"{node.Form}|{node.Workbook}|{node.Sheet}|{node.LastSheet}|{node.Address}|{node.Is3D}";
    }

    [Theory]
    [InlineData("A1", "Cell||||A1|False")]
    [InlineData("$A$1", "Cell||||$A$1|False")]
    [InlineData("I$19:$I$19", "Area||||I$19:$I$19|False")]
    [InlineData("A:A", "Columns||||A:A|False")]
    [InlineData("$A:$C", "Columns||||$A:$C|False")]
    [InlineData("1:1", "Rows||||1:1|False")]
    [InlineData("$3:$5", "Rows||||$3:$5|False")]
    [InlineData("Sheet8!E7", "Cell||Sheet8||E7|False")]
    [InlineData("'Inventory Forcast'!C$6", "Cell||Inventory Forcast||C$6|False")]
    [InlineData("'Jo''s Sheet'!A1", "Cell||Jo's Sheet||A1|False")]
    [InlineData("Sheet1:Sheet3!A1", "Cell||Sheet1|Sheet3|A1|True")]
    [InlineData("'Sheet 1:Sheet 3'!A1:B2", "Area||Sheet 1|Sheet 3|A1:B2|True")]
    [InlineData("[Book.xlsx]Sheet1!A1", "Cell|Book.xlsx|Sheet1||A1|False")]
    [InlineData("'[Book 1.xlsx]Sheet 1'!A1", "Cell|Book 1.xlsx|Sheet 1||A1|False")]
    [InlineData("'C:\\Dir\\[Book.xlsx]Sheet1'!A1", "Cell|Book.xlsx|Sheet1||A1|False")]
    [InlineData("Sheet1!#REF!", "RefError||Sheet1||#REF!|False")]
    [InlineData("Sheet1!A:A", "Columns||Sheet1||A:A|False")]
    [InlineData("Fin_P_FO!L39:L40", "Area||Fin_P_FO||L39:L40|False")]
    public void ReferenceParts(string address, string expected) => Assert.Equal(expected, RefInfo(address));

    [Theory]
    [InlineData("Hundred", "Hundred||")]
    [InlineData("Sheet1!MyName", "MyName||Sheet1")]
    [InlineData("'My Sheet'!MyName", "MyName||My Sheet")]
    [InlineData("Book1.xlsx!Name", "Name|Book1.xlsx|")]
    [InlineData("[Book.xlsx]Sheet!Name", "Name|Book.xlsx|Sheet")]
    [InlineData("[Book.xlsx]!Name", "Name|Book.xlsx|")]
    public void NameParts(string name, string expected)
    {
        var node = (NameNode)Parse("=" + name).Root!.Body!;
        Assert.Equal(expected, $"{node.Name}|{node.Workbook}|{node.Sheet}");
        Assert.Equal(1, node.Start);
        Assert.Equal(name.Length, node.Length);
    }

    [Theory]
    [InlineData("Table1[Col]", "Table1|[Col]")]
    [InlineData("Table1[[#Headers],[Col]]", "Table1|[[#Headers],[Col]]")]
    [InlineData("Table1[@Col]", "Table1|[@Col]")]
    [InlineData("Table1[]", "Table1|[]")]
    [InlineData("[@Col]", "|[@Col]")]
    [InlineData("[@[Col 1]]", "|[@[Col 1]]")]
    public void StructuredReferenceParts(string reference, string expected)
    {
        var node = (StructuredReferenceNode)Parse("=" + reference).Root!.Body!;
        Assert.Equal(expected, $"{node.Table}|{node.Specifier}");
    }

    [Theory]
    [InlineData("=1", LiteralKind.Number)]
    [InlineData("=\"x\"", LiteralKind.String)]
    [InlineData("=TRUE", LiteralKind.Bool)]
    [InlineData("=#N/A", LiteralKind.Error)]
    public void LiteralKinds(string formula, LiteralKind expected) =>
        Assert.Equal(expected, ((LiteralNode)Parse(formula).Root!.Body!).LiteralKind);

    [Theory]
    [InlineData("=SUM(A1,2)+\"x\"&{1}*(B1)")]
    [InlineData("=IF(A1,,2)")]
    [InlineData("=-2^2")]
    public void ValidFormulasHaveNoErrors(string formula) => Assert.False(Parse(formula).HasErrors);

    // ---- error recovery ----------------------------------------------------------------------------------

    [Theory]
    [InlineData("=A1+", "?<A1+>(A1)")]
    [InlineData("=A1+*B1", "?<A1+*B1>(A1,B1)")]
    [InlineData("=SUM(A1", "SUM(A1)")]
    [InlineData("=SUM(A1,B1", "SUM(A1;B1)")]
    [InlineData("=SUM(A1,", "SUM(A1;_)")]
    [InlineData("=SUM(", "SUM()")]
    [InlineData("=IF(A1,B1,", "IF(A1;B1;_)")]
    [InlineData("=(", "[]")]
    [InlineData("=()", "[]")]
    [InlineData("=(A1", "[A1]")]
    [InlineData("=(A1+", "[?<A1+>(A1)]")]
    [InlineData("=)", "?<)>()")]
    [InlineData("=A1)", "?<A1)>(A1)")]
    [InlineData("=1+2)+3", "?<1+2)+3>((+ 1 2))")]
    [InlineData("=@#$", "?<@#$>()")]
    [InlineData("=*", "?<*>()")]
    [InlineData("=+", "?<+>()")]
    [InlineData("=-", "?<->()")]
    [InlineData("=A1 2", "?<A1 2>(A1)")]
    [InlineData("=SUM(*)", "SUM(?<*>())")]
    [InlineData("=SUM(A1 B1 2)", "SUM(?<A1 B1 2>((isect A1 B1)))")]
    [InlineData("=SUM(A1;B1)", "SUM(?<A1;B1>(A1,B1))")]
    [InlineData("=SUM(A1,;B1,C1)", "SUM(A1;?<;B1>(B1);C1)")]
    [InlineData("=SUM((A1+)),B1)", "?<SUM((A1+)),B1)>((union SUM([?<A1+>(A1)]) B1))")]
    [InlineData("={1,2", "{1,2}")]
    [InlineData("={1,2)", "{1,?<2)>(2)}")]
    [InlineData("={", "{}")]
    [InlineData("=:", "?<:>()")]
    [InlineData("=A1:", "?<A1:>(A1)")]
    [InlineData("=SUM(A1,B1))+C1", "?<SUM(A1,B1))+C1>(SUM(A1;B1),C1)")]
    public void ErrorRecovery(string formula, string expected)
    {
        var parsed = Parse(formula);
        Assert.True(parsed.HasErrors);
        Assert.Equal(expected, TestHelpers.Dump(parsed.Root, formula));
    }

    [Fact]
    public void MissingCloseParenOnFunction()
    {
        const string text = "=SUM(A1,B1";
        var parsed = Parse(text);
        var call = (FunctionCallNode)parsed.Root!.Body!;
        Assert.Equal(-1, call.CloseParen);
        Assert.Equal(1, call.Start);
        Assert.Equal(text.Length, call.End);
        Assert.True(parsed.HasErrors);
    }

    [Fact]
    public void MissingCloseParenOnGroup()
    {
        const string text = "=(A1+B1";
        var paren = (ParenNode)Parse(text).Root!.Body!;
        Assert.Equal(-1, paren.CloseParen);
        Assert.Equal(1, paren.Start);
        Assert.Equal(text.Length, paren.End);
    }

    [Fact]
    public void TrailingWhitespaceIsNotPartOfAnUnclosedNode()
    {
        var call = (FunctionCallNode)Parse("=SUM(A1,B1   ").Root!.Body!;
        Assert.Equal(10, call.End);
    }

    [Fact]
    public void ReferencesInsideUnparsedRegionsStillBecomeNodes()
    {
        const string text = "=A1+*B1+Sheet2!C3*Hundred*Table1[Col]";
        var parsed = Parse(text);
        var leaves = parsed.Nodes.Where(n => n.Kind is NodeKind.Reference or NodeKind.Name or NodeKind.StructuredReference)
            .Select(n => n.GetText(text)).ToArray();
        Assert.Equal(new[] { "A1", "B1", "Sheet2!C3", "Hundred", "Table1[Col]" }, leaves);
    }

    [Fact]
    public void ReferencesInsideBrokenFunctionArgumentsStillBecomeNodes()
    {
        const string text = "=SUM(A1 ;B1:C2, ?, D1)";
        var parsed = Parse(text);
        var leaves = parsed.Nodes.Where(n => n.Kind == NodeKind.Reference).Select(n => n.GetText(text)).ToArray();
        Assert.Equal(new[] { "A1", "B1:C2", "D1" }, leaves);
    }

    [Fact]
    public void ErrorRecoveryNeverThrowsOnHostileInput()
    {
        foreach (var text in new[]
                 {
                     "=", "=)", "=((", "=))", "=,", "=,,,", "=;;", "={;}", "={,}", "={{", "=}}", "=F(,,", "=F(((",
                     "=A1:::B1", "=A1 ,,, B1", "=\"", "='", "=[", "=]", "=#", "=%%%", "=@@@", "=---", "=^^", "=<>><",
                 })
        {
            var parsed = Parse(text);
            Assert.NotNull(parsed.Root);
            Assert.Equal(text.Length, parsed.Root!.Length);
        }
    }

    [Fact]
    public void DeepNestingIsGuardedAgainstStackOverflow()
    {
        string opens = new('(', 5000);
        var parsed = Parse("=" + opens + "A1" + new string(')', 5000));
        Assert.True(parsed.HasErrors);
        Assert.Contains(parsed.Nodes, n => n is ReferenceNode);

        var prefixes = Parse("=" + new string('-', 5000) + "A1");
        Assert.Contains(prefixes.Nodes, n => n is ReferenceNode);

        var calls = Parse("=" + string.Concat(Enumerable.Repeat("SUM(", 3000)) + "1");
        Assert.True(calls.HasErrors);
    }

    [Fact]
    public void RealisticDepthIsParsedWithoutErrors()
    {
        // Excel allows 64 nested levels.
        string text = "=" + string.Concat(Enumerable.Repeat("SUM(", 60)) + "A1" + new string(')', 60);
        var parsed = Parse(text);
        Assert.False(parsed.HasErrors);
        Assert.Equal(60, parsed.Nodes.OfType<FunctionCallNode>().Count());
    }

    [Fact]
    public void LongChainOfTermsParses()
    {
        string text = "=" + string.Join("+", Enumerable.Range(1, 800).Select(i => "A" + i));
        var parsed = Parse(text);
        Assert.False(parsed.HasErrors);
        Assert.Equal(800, parsed.Nodes.OfType<ReferenceNode>().Count());
        Assert.Equal(799, parsed.Nodes.OfType<BinaryNode>().Count());
    }
}
