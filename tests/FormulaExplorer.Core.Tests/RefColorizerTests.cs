using FormulaExplorer.Core.Analysis;
using FormulaExplorer.Core.Syntax;

namespace FormulaExplorer.Core.Tests;

public class RefColorizerTests
{
    private static ColorMap Assign(string formula, string? root = "Sheet1", Func<SyntaxNode, bool>? colorable = null) =>
        RefColorizer.Assign(FormulaParser.Parse(formula), root, colorable ?? (_ => true));

    private static int[] Indexes(string formula, string? root = "Sheet1", Func<SyntaxNode, bool>? colorable = null) =>
        Assign(formula, root, colorable).Nodes.Select(n => n.PaletteIndex).ToArray();

    [Fact]
    public void PaletteIsTheRangeFinderPalette()
    {
        Assert.Equal(new uint[] { 0x0000FF, 0x008000, 0x9900CC, 0x800000, 0x00CC33, 0xFF6600, 0xCC0099 }, RefColorizer.Palette);
    }

    [Theory]
    [InlineData("=A1+B1+C1", new[] { 0, 1, 2 })]
    [InlineData("=A1+A1+B1", new[] { 0, 0, 1 })]
    [InlineData("=A1+B1+A1", new[] { 0, 1, 0 })]
    [InlineData("=A1+B1+C1+D1+E1+F1+G1", new[] { 0, 1, 2, 3, 4, 5, 6 })]
    [InlineData("=A1+B1+C1+D1+E1+F1+G1+H1", new[] { 0, 1, 2, 3, 4, 5, 6, 0 })]
    [InlineData("=A1+B1+C1+D1+E1+F1+G1+H1+I1+A1+H1", new[] { 0, 1, 2, 3, 4, 5, 6, 0, 1, 0, 0 })]
    [InlineData("=SUM(A1:A3,B1:B3)", new[] { 0, 1 })]
    [InlineData("=1+2", new int[0])]
    [InlineData("=SUM(1,2)", new int[0])]
    public void PaletteIndexesFollowAppearance(string formula, int[] expected) => Assert.Equal(expected, Indexes(formula));

    [Fact]
    public void RgbMatchesThePaletteIndex()
    {
        var colors = Assign("=A1+B1+C1+D1+E1+F1+G1+H1");
        Assert.All(colors.Nodes, n => Assert.Equal(RefColorizer.Palette[n.PaletteIndex], n.Rgb));
        Assert.Equal(0x0000FFu, colors.Nodes[7].Rgb);
    }

    [Fact]
    public void SameCellWrittenDifferentlyShareAColor()
    {
        var colors = Assign("=A1+$A$1+a1+Sheet1!A1+'Sheet1'!A$1+sheet1!$a$1");
        Assert.Equal(6, colors.Nodes.Count);
        Assert.All(colors.Nodes, n => Assert.Equal(0, n.PaletteIndex));
        Assert.Single(colors.Nodes.Select(n => n.Key).Distinct());
    }

    [Fact]
    public void AreaOfOneCellIsNotTheSameAsTheCell()
    {
        Assert.Equal(new[] { 0, 1, 0, 1 }, Indexes("=I19+I19:I19+$I$19+$I$19:$I$19"));
    }

    [Fact]
    public void DifferentSheetsDifferentColors()
    {
        Assert.Equal(new[] { 0, 1 }, Indexes("=Sheet2!A1+A1", "Sheet1"));
        Assert.Equal(new[] { 0, 0 }, Indexes("=Sheet2!A1+A1", "Sheet2"));
    }

    [Fact]
    public void QuotesAreIgnoredInSheetNames()
    {
        Assert.Equal(new[] { 0, 0 }, Indexes("='Sheet2'!A1+Sheet2!A1"));
        Assert.Equal(new[] { 0, 0 }, Indexes("='My Sheet'!A1+'my sheet'!A1"));
        Assert.Equal(new[] { 0, 1 }, Indexes("='Jo''s'!A1+'Jos'!A1"));
    }

    [Fact]
    public void UnknownRootSheetKeysUnqualifiedReferencesSeparately()
    {
        Assert.Equal(new[] { 0, 1 }, Indexes("=A1+Sheet1!A1", root: null));
    }

    [Fact]
    public void ThreeDAndExternalReferencesAreDistinct()
    {
        Assert.Equal(new[] { 0, 1, 2, 3 }, Indexes("=Sheet1:Sheet3!A1+Sheet1!A1+[Book.xlsx]Sheet1!A1+[Other.xlsx]Sheet1!A1"));
        Assert.Equal(new[] { 0, 0 }, Indexes("=[Book.xlsx]Sheet1!A1+'[Book.xlsx]Sheet1'!$A$1"));
    }

    [Fact]
    public void WholeColumnsAndRowsAreColored()
    {
        Assert.Equal(new[] { 0, 1, 0 }, Indexes("=SUM(A:A)+SUM(1:1)+SUM($A:$A)"));
    }

    [Fact]
    public void NamesAreColoredAndCaseInsensitive()
    {
        var colors = Assign("=Hundred+A1+hundred+HUNDRED");
        Assert.Equal(new[] { 0, 1, 0, 0 }, colors.Nodes.Select(n => n.PaletteIndex).ToArray());
        Assert.Equal(NodeKind.Name, colors.Nodes[0].Node.Kind);
    }

    [Fact]
    public void NamesWithDifferentScopesAreDifferent()
    {
        Assert.Equal(new[] { 0, 1, 2 }, Indexes("=Total+Sheet1!Total+Sheet2!Total"));
    }

    [Fact]
    public void StructuredReferencesKeyByWholeText()
    {
        Assert.Equal(new[] { 0, 0, 1, 2 }, Indexes("=Table1[Col]+table1[COL]+Table1[Other]+[@Col]"));
    }

    [Fact]
    public void NotColorableNodesAreSkippedAndDoNotConsumeSlots()
    {
        var colors = Assign("=Hundred+A1+Tax+B1", colorable: n => n is not NameNode);
        Assert.Equal(new[] { 0, 1 }, colors.Nodes.Select(n => n.PaletteIndex).ToArray());
        Assert.Equal(new[] { "A1", "B1" }, colors.Nodes.Select(n => n.Node.GetText("=Hundred+A1+Tax+B1")).ToArray());
    }

    [Fact]
    public void RefErrorsCanBeExcludedByThePredicate()
    {
        var colors = Assign("=Sheet1!#REF!+A1+B1", colorable: n => n is not ReferenceNode { Form: ReferenceForm.RefError });
        Assert.Equal(new[] { 0, 1 }, colors.Nodes.Select(n => n.PaletteIndex).ToArray());
    }

    [Fact]
    public void PredicateReceivesOnlyReferenceLikeNodes()
    {
        var seen = new List<NodeKind>();
        Assign("=SUM(A1,Hundred,Table1[C])+1&\"x\"", colorable: n => { seen.Add(n.Kind); return true; });
        Assert.Equal(new[] { NodeKind.Reference, NodeKind.Name, NodeKind.StructuredReference }, seen);
    }

    [Fact]
    public void ColorOfLooksUpByNodeInstance()
    {
        const string text = "=A1+B1";
        var parsed = FormulaParser.Parse(text);
        var colors = RefColorizer.Assign(parsed, "Sheet1", _ => true);
        var a1 = parsed.Nodes.OfType<ReferenceNode>().First();
        var b1 = parsed.Nodes.OfType<ReferenceNode>().Last();
        Assert.Equal(0x0000FFu, colors.ColorOf(a1));
        Assert.Equal(0x008000u, colors.ColorOf(b1));
        Assert.Null(colors.ColorOf(parsed.Root!));
        Assert.Null(colors.ColorOf(parsed.Root!.Body!));
    }

    [Fact]
    public void NodesAreInTextOrder()
    {
        const string text = "=SUM(C3,A1)+B2*Hundred";
        var starts = Assign(text).Nodes.Select(n => n.Node.Start).ToArray();
        Assert.Equal(starts.OrderBy(s => s), starts);
        Assert.Equal(4, starts.Length);
    }

    [Fact]
    public void ReferencesInsideUnparsedRegionsAreColored()
    {
        Assert.Equal(new[] { 0, 1, 0 }, Indexes("=A1+*B1+A1"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("hello")]
    [InlineData("A1+B1")]
    public void NonFormulaHasNoColors(string text) => Assert.Empty(Assign(text).Nodes);

    [Fact]
    public void NoColorableNodesReturnsTheEmptyMap() => Assert.Same(ColorMap.Empty, Assign("=1+2"));
}
