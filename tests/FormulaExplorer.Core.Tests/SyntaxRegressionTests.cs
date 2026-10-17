using FormulaExplorer.Core.Analysis;
using FormulaExplorer.Core.Syntax;
using static FormulaExplorer.Core.Tests.TestHelpers;

namespace FormulaExplorer.Core.Tests;

/// <summary>Regression tests for review findings core-1 .. core-4.</summary>
public class SyntaxRegressionTests
{
    // ---- core-1: #REF! used as the sheet prefix of a deleted sheet ----------------------------------

    [Theory]
    [InlineData("=#REF!A1+1", "R[#REF!A1] Bin[+] N[1]")]
    [InlineData("=SUM(#REF!A1:B2)", "F[SUM] Open[(] R[#REF!A1:B2] Close[)]")]
    [InlineData("=#REF!Name*2", "R[#REF!Name] Bin[*] N[2]")]
    [InlineData("=#REF!$A$1", "R[#REF!$A$1]")]
    public void DeletedSheetPrefixIsOneReferenceToken(string text, string expected) => Assert.Equal(expected, Tokens(text));

    [Theory]
    [InlineData("=#REF!", "E[#REF!]")]
    [InlineData("=#REF!+1", "E[#REF!] Bin[+] N[1]")]
    [InlineData("=Sheet1!#REF!", "R[Sheet1!#REF!]")]
    [InlineData("=#REF!#REF!", "E[#REF!] E[#REF!]")]
    [InlineData("=#REF!(", "E[#REF!] Open[(]")]
    public void PlainRefErrorsTokenizeAsBefore(string text, string expected) => Assert.Equal(expected, Tokens(text));

    [Fact]
    public void DeletedSheetReferenceParsesCleanlyAsRefError()
    {
        foreach (var text in new[] { "=#REF!A1+1", "=SUM(#REF!A1:B2)" })
        {
            var parsed = FormulaParser.Parse(text);
            Assert.False(parsed.HasErrors, text);
            var reference = Assert.Single(parsed.Nodes.OfType<ReferenceNode>());
            Assert.Equal(ReferenceForm.RefError, reference.Form);
            Assert.Null(reference.Sheet);
        }
        Assert.Equal("(+ #REF!A1 1)", Tree("=#REF!A1+1"));
        Assert.Equal("SUM(#REF!A1:B2)", Tree("=SUM(#REF!A1:B2)"));
    }

    [Fact]
    public void DeletedSheetReferenceIsNotColoredWithTheSessionPredicate()
    {
        var parsed = FormulaParser.Parse("=#REF!A1+B1");
        var colors = RefColorizer.Assign(parsed, "Sheet1", n => n is not ReferenceNode { Form: ReferenceForm.RefError });
        var only = Assert.Single(colors.Nodes);
        Assert.Equal("B1", only.Node.GetText(parsed.Text));
    }

    [Fact]
    public void HoveringThePlusAfterADeletedSheetReferenceHighlightsTheBinary()
    {
        Assert.Equal("#REF!A1+1", HighlightOn("=#REF!A1+1", "+"));
        Assert.Equal("#REF!A1", HighlightOn("=#REF!A1+1", "A1"));
    }

    [Fact]
    public void IntersectionWithADeletedRangeIsRecognised()
    {
        Assert.Equal("SUM((isect A1:A3 #REF!))", Tree("=SUM(A1:A3 #REF!)"));
        Assert.Equal("SUM((isect #REF! A1:A3))", Tree("=SUM(#REF! A1:A3)"));
        Assert.False(FormulaParser.Parse("=SUM(A1:A3 #REF!)").HasErrors);
    }

    // ---- core-2: qualified function calls ------------------------------------------------------------

    [Theory]
    [InlineData("=PERSONAL.XLSB!Fn(A1,B1)", "F[PERSONAL.XLSB!Fn] Open[(] R[A1] Sep[,] R[B1] Close[)]")]
    [InlineData("=Sheet1!MyLambda(2)", "F[Sheet1!MyLambda] Open[(] N[2] Close[)]")]
    [InlineData("='C:\\x\\Book.xlsm'!Fn(A1)", "F['C:\\x\\Book.xlsm'!Fn] Open[(] R[A1] Close[)]")]
    [InlineData("=[Book.xlsx]Sheet1!Fn(1)", "F[[Book.xlsx]Sheet1!Fn] Open[(] N[1] Close[)]")]
    public void QualifiedFunctionCallsAreFunctionTokens(string text, string expected) => Assert.Equal(expected, Tokens(text));

    [Fact]
    public void QualifiedNameWithoutParenthesisIsStillAName() =>
        Assert.Equal("Nm[Sheet1!Name] Bin[+] N[1]", Tokens("=Sheet1!Name+1"));

    [Fact]
    public void QualifiedCallParsesAsFunctionCallInsideAnExpression()
    {
        const string text = "=1+PERSONAL.XLSB!Fn(A1,B1)*2";
        var parsed = FormulaParser.Parse(text);
        Assert.False(parsed.HasErrors);
        Assert.Equal("(+ 1 (* PERSONAL.XLSB!Fn(A1;B1) 2))", Tree(text));
        Assert.Equal("1+PERSONAL.XLSB!Fn(A1,B1)*2", HighlightOn(text, "+"));
        Assert.Equal("PERSONAL.XLSB!Fn(A1,B1)*2", HighlightOn(text, "*"));
        Assert.Equal("PERSONAL.XLSB!Fn(A1,B1)", HighlightOn(text, "Fn"));
        Assert.Equal("PERSONAL.XLSB!Fn(A1,B1)", HighlightOn(text, "("));
        Assert.Equal("PERSONAL.XLSB!Fn(A1,B1)", HighlightOn(text, ")"));
        Assert.Null(HighlightOn(text, ","));
    }

    [Theory]
    [InlineData("=Sheet1!MyLambda(2)")]
    [InlineData("='C:\\x\\Book.xlsm'!Fn(A1)")]
    [InlineData("=[Book.xlsx]Sheet1!Fn(1)")]
    public void OtherQualifiedCallsParseWithoutErrors(string text) => Assert.False(FormulaParser.Parse(text).HasErrors);

    // ---- core-3: immediately invoked LAMBDA ----------------------------------------------------------

    [Fact]
    public void InvocationCommasAreArgumentSeparators() =>
        Assert.Equal("F[LAMBDA] Open[(] Nm[x] Sep[,] Nm[y] Sep[,] Nm[x] Bin[+] Nm[y] Close[)] Open[(] N[1] Sep[,] N[2] Close[)]",
            Tokens("=LAMBDA(x,y,x+y)(1,2)"));

    [Fact]
    public void ImmediateInvocationIsACallOfTheCallee()
    {
        const string text = "=LAMBDA(x,y,x+y)(1,2)";
        var parsed = FormulaParser.Parse(text);
        Assert.False(parsed.HasErrors);
        var invocation = Assert.IsType<FunctionCallNode>(parsed.Root!.Body);
        Assert.NotNull(invocation.Callee);
        Assert.Equal(1, invocation.Start);
        Assert.Equal(text.Length, invocation.End);
        Assert.Equal(2, invocation.Arguments.Count);
        Assert.Equal(text.IndexOf(")(", StringComparison.Ordinal) + 1, invocation.OpenParen);
        Assert.Equal("LAMBDA(x,y,x+y)", invocation.Callee!.GetText(text));
        Assert.Equal("LAMBDA(x,y,x+y)", invocation.Name);
        // the invocation's own parentheses highlight the invocation, the callee keeps its own characters
        Assert.Equal(text[1..], HighlightAt(text, invocation.OpenParen));
        Assert.Equal(text[1..], HighlightAt(text, text.Length - 1));
        Assert.Equal("LAMBDA(x,y,x+y)", HighlightOn(text, "LAMBDA"));
        Assert.Equal("x+y", HighlightOn(text, "+"));
        Assert.Null(HighlightAt(text, text.IndexOf("1,2", StringComparison.Ordinal) + 1)); // the ',' between 1 and 2
    }

    [Fact]
    public void OperatorsAfterAnInvocationStillParse()
    {
        const string text = "=1+LAMBDA(x,x)(2)*3";
        Assert.False(FormulaParser.Parse(text).HasErrors);
        Assert.Equal("LAMBDA(x,x)(2)*3", HighlightOn(text, "*"));
        Assert.Equal("1+LAMBDA(x,x)(2)*3", HighlightOn(text, "+"));
    }

    [Fact]
    public void InvocationsChainAndWorkOnOtherCalls()
    {
        Assert.False(FormulaParser.Parse("=LAMBDA(x,LAMBDA(y,x+y))(1)(2)").HasErrors);
        var parsed = FormulaParser.Parse("=INDEX(fns,1)(A1,B1)");
        Assert.False(parsed.HasErrors);
        Assert.Equal(2, ((FunctionCallNode)parsed.Root!.Body!).Arguments.Count);
    }

    [Fact]
    public void ParenthesisAfterWhitespaceIsStillAnIntersection() =>
        Assert.Equal("(isect SUM(A1:A2) [B1:B2])", Tree("=SUM(A1:A2) (B1:B2)"));

    [Fact]
    public void GroupInsideACallKeepsUnionCommas() =>
        Assert.Equal("SUM([(union A1 B1)])", Tree("=SUM((A1,B1))"));

    // ---- core-4: LET / LAMBDA variables --------------------------------------------------------------

    private static string[] Locals(string text) =>
        FormulaParser.Parse(text).Nodes.OfType<NameNode>().Where(n => n.IsLocal).Select(n => n.GetText(text) + "@" + n.Start).ToArray();

    [Fact]
    public void LetVariableDeclarationAndUsesAreLocal() =>
        Assert.Equal(new[] { "rate@5", "rate@18" }, Locals("=LET(rate,0.05,A1*rate)"));

    [Fact]
    public void LocalNamesAreNotColoredNorLinkCandidates()
    {
        const string text = "=LET(rate,0.05,A1*rate)";
        var parsed = FormulaParser.Parse(text);
        var colors = RefColorizer.Assign(parsed, "Sheet1", _ => true);
        var only = Assert.Single(colors.Nodes);
        Assert.Equal("A1", only.Node.GetText(text));
        foreach (var name in parsed.Nodes.OfType<NameNode>())
            Assert.False(HitMap.IsLinkCandidate(name));
        Assert.True(HitMap.IsLinkCandidate(parsed.Nodes.OfType<ReferenceNode>().Single()));
    }

    [Fact]
    public void LambdaParametersInsideLetAreLocal() =>
        Assert.Equal(new[] { "f@5", "x@14", "x@16" }, Locals("=LET(f,LAMBDA(x,x*2),f(3))"));

    [Fact]
    public void LetValueExpressionSeesTheOuterName() =>
        // the 'rate' inside the value of 'rate' is the workbook name; only the later uses are the variable
        Assert.Equal(new[] { "rate@5", "rate@17" }, Locals("=LET(rate,rate+1,rate*2)"));

    [Fact]
    public void QualifiedAndUnrelatedNamesAreNotLocal()
    {
        var parsed = FormulaParser.Parse("=LET(rate,1,Sheet1!rate+other+rate)");
        var locals = parsed.Nodes.OfType<NameNode>().Where(n => n.IsLocal).ToArray();
        Assert.Equal(2, locals.Length);
        Assert.All(locals, n => Assert.Null(n.Sheet));
        Assert.DoesNotContain(parsed.Nodes.OfType<NameNode>(), n => n.IsLocal && (n.Name == "other" || n.Sheet is not null));
    }

    [Fact]
    public void PrefixedFunctionAndParameterNamesAreRecognised()
    {
        Assert.Equal(2, Locals("=_xlfn.LET(_xlpm.rate,0.05,A1*_xlpm.rate)").Length);
        Assert.Equal(2, Locals("=_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)").Length);
    }

    [Fact]
    public void NamesOutsideLetOrLambdaAreNeverLocal()
    {
        Assert.Empty(Locals("=SUM(rate,x)+rate"));
        var parsed = FormulaParser.Parse("=LET(a,1,a)+a");
        Assert.Equal(new[] { 5, 9 }, parsed.Nodes.OfType<NameNode>().Where(n => n.IsLocal).Select(n => n.Start).ToArray());
        Assert.False(parsed.Nodes.OfType<NameNode>().Last().IsLocal);
    }
}
