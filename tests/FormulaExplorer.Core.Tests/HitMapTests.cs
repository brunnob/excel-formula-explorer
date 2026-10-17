using FormulaExplorer.Core.Analysis;
using FormulaExplorer.Core.Syntax;

namespace FormulaExplorer.Core.Tests;

public class HitMapTests
{
    // ---- one case per rule --------------------------------------------------------------------------

    [Theory]
    // '=' -> whole formula
    [InlineData("=A1+B1", 0, "=A1+B1")]
    [InlineData("=SUM(A1)", 0, "=SUM(A1)")]
    // reference characters -> the reference
    [InlineData("=A1+B1", 1, "A1")]
    [InlineData("=A1+B1", 2, "A1")]
    [InlineData("=A1+B1", 4, "B1")]
    [InlineData("=A1+B1", 5, "B1")]
    [InlineData("=I$19:$I$19", 5, "I$19:$I$19")]
    [InlineData("='Inventory Forcast'!C$6", 1, "'Inventory Forcast'!C$6")]
    [InlineData("='Inventory Forcast'!C$6", 12, "'Inventory Forcast'!C$6")]
    [InlineData("=Sheet8!E7+1", 8, "Sheet8!E7")]
    // names, structured references
    [InlineData("=Hundred*2", 4, "Hundred")]
    [InlineData("=Table1[Col]+1", 8, "Table1[Col]")]
    // literals
    [InlineData("=12+\"ab\"", 1, "12")]
    [InlineData("=12+\"ab\"", 5, "\"ab\"")]
    [InlineData("=12+\"ab\"", 6, "\"ab\"")]
    [InlineData("=TRUE", 2, "TRUE")]
    [InlineData("=#N/A", 3, "#N/A")]
    // binary operator -> the whole binary expression
    [InlineData("=A1+B1", 3, "A1+B1")]
    [InlineData("=1+2*3", 2, "1+2*3")]
    [InlineData("=1+2*3", 4, "2*3")]
    [InlineData("=A1<=B1", 3, "A1<=B1")]
    [InlineData("=A1<=B1", 4, "A1<=B1")]
    [InlineData("=A1<>B1", 4, "A1<>B1")]
    // prefix / postfix
    [InlineData("=-A1", 1, "-A1")]
    [InlineData("=-A1", 2, "A1")]
    [InlineData("=--A1", 1, "--A1")]
    [InlineData("=--A1", 2, "-A1")]
    [InlineData("=A1%", 3, "A1%")]
    [InlineData("=A1%", 1, "A1")]
    [InlineData("=@A1", 1, "@A1")]
    [InlineData("=A1#", 3, "A1#")]
    // parentheses -> the group
    [InlineData("=(A1)", 1, "(A1)")]
    [InlineData("=(A1)", 4, "(A1)")]
    [InlineData("=(A1)", 2, "A1")]
    [InlineData("=((A1))", 1, "((A1))")]
    [InlineData("=((A1))", 2, "(A1)")]
    [InlineData("=((A1))", 5, "(A1)")]
    [InlineData("=((A1))", 6, "((A1))")]
    // function name, '(' and ')' -> the whole call
    [InlineData("=SUM(A1,B1)", 1, "SUM(A1,B1)")]
    [InlineData("=SUM(A1,B1)", 3, "SUM(A1,B1)")]
    [InlineData("=SUM(A1,B1)", 4, "SUM(A1,B1)")]
    [InlineData("=SUM(A1,B1)", 10, "SUM(A1,B1)")]
    [InlineData("=SUM()", 4, "SUM()")]
    [InlineData("=SUM()", 5, "SUM()")]
    [InlineData("=_xlfn.XLOOKUP(1)", 8, "_xlfn.XLOOKUP(1)")]
    // argument separator -> nothing
    [InlineData("=SUM(A1,B1)", 7, null)]
    [InlineData("=IF(A1,,2)", 6, null)]
    [InlineData("=IF(A1,,2)", 7, null)]
    // insignificant whitespace -> nothing
    [InlineData("= A1", 1, null)]
    [InlineData("=SUM(A1, B1)", 8, null)]
    [InlineData("=A1 + B1", 3, null)]
    [InlineData("=A1 + B1", 5, null)]
    [InlineData("=A1 ", 3, null)]
    [InlineData("=SUM(A1,\nB1)", 8, null)]
    // reference operators -> the binary node
    [InlineData("=A1:INDEX(B:B,3)", 3, "A1:INDEX(B:B,3)")]
    [InlineData("=Name1:Name2", 6, "Name1:Name2")]
    [InlineData("=A1  B1", 3, "A1  B1")]
    [InlineData("=A1  B1", 4, "A1  B1")]
    [InlineData("=(A1,B1)", 4, "A1,B1")]
    [InlineData("=A1:B2 B2:C3", 6, "A1:B2 B2:C3")]
    // array constants
    [InlineData("={1,2;3,4}", 1, "{1,2;3,4}")]
    [InlineData("={1,2;3,4}", 3, "{1,2;3,4}")]
    [InlineData("={1,2;3,4}", 5, "{1,2;3,4}")]
    [InlineData("={1,2;3,4}", 9, "{1,2;3,4}")]
    [InlineData("={1,2;3,4}", 2, "1")]
    [InlineData("={1,2;3,4}", 8, "4")]
    [InlineData("=SUM({1,2},3)", 5, "{1,2}")]
    [InlineData("=SUM({1,2},3)", 9, "{1,2}")]
    [InlineData("=SUM({1,2},3)", 10, null)]
    // unknown characters
    [InlineData("=A1~B1", 3, null)]
    [InlineData("=A1;B1", 3, null)]
    // out of range
    [InlineData("=A1", -1, null)]
    [InlineData("=A1", 3, null)]
    [InlineData("=A1", 1000, null)]
    public void NodeAtFollowsTheRules(string formula, int index, string? expected) =>
        Assert.Equal(expected, TestHelpers.HighlightAt(formula, index));

    [Fact]
    public void UnparsedRegionsAreNotHighlightedButTheirReferencesAre()
    {
        const string text = "=A1+*B1";
        Assert.Equal("=A1+*B1", TestHelpers.HighlightAt(text, 0));
        Assert.Equal("A1", TestHelpers.HighlightAt(text, 1));
        Assert.Null(TestHelpers.HighlightAt(text, 3)); // +
        Assert.Null(TestHelpers.HighlightAt(text, 4)); // *
        Assert.Equal("B1", TestHelpers.HighlightAt(text, 5));
    }

    [Fact]
    public void UnclosedFunctionStillHighlightsNameAndParen()
    {
        const string text = "=SUM(A1,B1";
        Assert.Equal("SUM(A1,B1", TestHelpers.HighlightAt(text, 2));
        Assert.Equal("SUM(A1,B1", TestHelpers.HighlightAt(text, 4));
        Assert.Equal("A1", TestHelpers.HighlightAt(text, 5));
        Assert.Null(TestHelpers.HighlightAt(text, 7));
    }

    [Fact]
    public void EveryCharacterOfAReferenceMapsToTheSameNodeInstance()
    {
        var parsed = FormulaParser.Parse("=1+'My Sheet'!B2:C3*2");
        var map = HitMap.Build(parsed);
        var reference = parsed.Nodes.OfType<ReferenceNode>().Single();
        for (int i = reference.Start; i < reference.End; i++)
            Assert.Same(reference, map.NodeAt(i));
    }

    [Fact]
    public void NodeAtReturnsNodesOfTheSameParsedTree()
    {
        var parsed = FormulaParser.Parse("=SUM(A1,B1)*(C1+D1)");
        var map = HitMap.Build(parsed);
        for (int i = 0; i < parsed.Text.Length; i++)
        {
            var node = map.NodeAt(i);
            if (node is not null) Assert.Contains(node, parsed.Nodes);
        }
    }

    [Theory]
    [InlineData("")]
    [InlineData("hello")]
    [InlineData("42")]
    [InlineData("A1+B1")]
    public void NonFormulaMapsEverythingToNull(string text)
    {
        var map = HitMap.Build(FormulaParser.Parse(text));
        for (int i = -1; i <= text.Length + 1; i++)
            Assert.Null(map.NodeAt(i));
    }

    // ---- IsLinkCandidate -------------------------------------------------------------------------------

    [Theory]
    [InlineData("=A1", true)]
    [InlineData("=A1:B2", true)]
    [InlineData("=Sheet1!A1", true)]
    [InlineData("=Hundred", true)]
    [InlineData("=Table1[Col]", true)]
    [InlineData("=SUM(1)", true)]
    [InlineData("=OFFSET(A1,0,0,1,2)", true)]
    [InlineData("=Name1:Name2", true)]
    [InlineData("=A1 B1", true)]
    [InlineData("=1+2", false)]
    [InlineData("=1", false)]
    [InlineData("=\"x\"", false)]
    [InlineData("=(A1)", false)]
    [InlineData("=-A1", false)]
    [InlineData("=A1%", false)]
    [InlineData("={1}", false)]
    [InlineData("=A1&B1", false)]
    public void IsLinkCandidateOnTheBody(string formula, bool expected) =>
        Assert.Equal(expected, HitMap.IsLinkCandidate(FormulaParser.Parse(formula).Root!.Body!));

    [Fact]
    public void UnionBinaryIsALinkCandidate()
    {
        var inner = ((ParenNode)FormulaParser.Parse("=(A1,B1)").Root!.Body!).Inner!;
        Assert.True(HitMap.IsLinkCandidate(inner));
    }

    [Fact]
    public void FormulaMissingAndUnparsedNodesAreNotLinkCandidates()
    {
        var parsed = FormulaParser.Parse("=SUM(A1,,*)");
        Assert.False(HitMap.IsLinkCandidate(parsed.Root!));
        Assert.False(HitMap.IsLinkCandidate(parsed.Nodes.OfType<MissingArgumentNode>().Single()));
        Assert.False(HitMap.IsLinkCandidate(parsed.Nodes.OfType<UnparsedNode>().Single()));
    }
}
