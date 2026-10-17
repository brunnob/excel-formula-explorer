using FormulaExplorer.Core.Analysis;
using FormulaExplorer.Core.Syntax;

namespace FormulaExplorer.Core.Tests;

/// <summary>
/// The formulas shown in the original product's video and screenshots (docs/referencias) with the exact
/// fragment highlighted when hovering a given character.
/// </summary>
public class VideoFixtureTests
{
    private const string F1 = "=1/(((($I27/Hundred)+1)^(K$31)))";
    private const string F2 = "=IF(K3=\"On\",1/(((-($I26/Hundred)+1)^(K$31))),0)";
    private const string F3 = "=IF(J3,(1/((365*(($I27/Hundred)+1)^(K$31)))),0)";
    private const string F4 = "=Sheet8!E7+(E30/C4)^(D4/D5)+D14-E21*E16+Summary!C5";
    private const string F5 = "=SUM(I$19:$I$19)+J$19*$I$28/Hundred-SUM(OFFSET($I$19,0,0,1,Aqui_Per))";
    private const string F6 = "=IF(C3,0,C11*'Inventory Forcast'!C$6)";
    private const string F7 = "=(SUM(Fin_P_FO!L39:L40)+Fin_P_FO!L26+Notes_FO!L113+Notes_FO!L101)*(Macro_FA!L183/Hundred)";

    private static string Body(string formula) => formula.Substring(1);

    // ---- F1 -----------------------------------------------------------------------------------------

    [Fact]
    public void F1_Caret_HighlightsThePowerExpression() =>
        Assert.Equal("(($I27/Hundred)+1)^(K$31)", TestHelpers.HighlightOn(F1, "^"));

    [Fact]
    public void F1_CloseParenAfterPlusOne_HighlightsItsGroup() =>
        Assert.Equal("(($I27/Hundred)+1)", TestHelpers.HighlightAt(F1, F1.IndexOf("+1)", StringComparison.Ordinal) + 2));

    [Fact]
    public void F1_FirstSlash_HighlightsTheWholeBody()
    {
        Assert.Equal('/', F1[2]);
        Assert.Equal("1/(((($I27/Hundred)+1)^(K$31)))", TestHelpers.HighlightAt(F1, 2));
    }

    [Fact]
    public void F1_InnerSlash_HighlightsOnlyItsOperands() =>
        Assert.Equal("$I27/Hundred", TestHelpers.HighlightOn(F1, "/", occurrence: 1));

    [Fact]
    public void F1_Hundred_HighlightsTheName() =>
        Assert.Equal("Hundred", TestHelpers.HighlightOn(F1, "Hundred", offset: 3));

    [Fact]
    public void F1_Equals_HighlightsTheWholeText() => Assert.Equal(F1, TestHelpers.HighlightAt(F1, 0));

    [Fact]
    public void F1_ParenGroupsAreNestedForests()
    {
        // The four '(' after "1/" and their partners.
        Assert.Equal("(((($I27/Hundred)+1)^(K$31)))", TestHelpers.HighlightAt(F1, 3));
        Assert.Equal("((($I27/Hundred)+1)^(K$31))", TestHelpers.HighlightAt(F1, 4));
        Assert.Equal("(($I27/Hundred)+1)", TestHelpers.HighlightAt(F1, 5));
        Assert.Equal("($I27/Hundred)", TestHelpers.HighlightAt(F1, 6));
        Assert.Equal("(K$31)", TestHelpers.HighlightOn(F1, "(K$31)"));
        Assert.Equal("(K$31)", TestHelpers.HighlightAt(F1, F1.IndexOf("(K$31)", StringComparison.Ordinal) + 5));
        Assert.Equal("K$31", TestHelpers.HighlightOn(F1, "K$31", offset: 2));
    }

    // ---- F2 -----------------------------------------------------------------------------------------

    [Fact]
    public void F2_Caret_HighlightsThePowerExpression() =>
        Assert.Equal("(-($I26/Hundred)+1)^(K$31)", TestHelpers.HighlightOn(F2, "^"));

    [Fact]
    public void F2_IfName_HighlightsTheWholeCall()
    {
        Assert.Equal(Body(F2), TestHelpers.HighlightOn(F2, "IF"));
        Assert.Equal(Body(F2), TestHelpers.HighlightOn(F2, "IF", offset: 1));
    }

    [Fact]
    public void F2_OnString_HighlightsTheString() => Assert.Equal("\"On\"", TestHelpers.HighlightOn(F2, "\"On\"", offset: 2));

    [Fact]
    public void F2_Minus_HighlightsTheNegatedGroup() =>
        Assert.Equal("-($I26/Hundred)", TestHelpers.HighlightOn(F2, "-($I26"));

    [Fact]
    public void F2_ArgumentCommasHighlightNothing()
    {
        Assert.Null(TestHelpers.HighlightOn(F2, ",1/"));
        Assert.Null(TestHelpers.HighlightAt(F2, F2.LastIndexOf(',')));
    }

    [Fact]
    public void F2_ComparisonHighlightsTheComparison() => Assert.Equal("K3=\"On\"", TestHelpers.HighlightOn(F2, "=", occurrence: 1));

    // ---- F3 -----------------------------------------------------------------------------------------

    [Fact]
    public void F3_CloseParenOfTheOuterGroup_HighlightsTheGroupStartingAtDoubleParen()
    {
        int closing = F3.IndexOf("(K$31)))", StringComparison.Ordinal) + 7; // third ')' of "(K$31)))"
        Assert.Equal(')', F3[closing]);
        Assert.Equal("((365*(($I27/Hundred)+1)^(K$31)))", TestHelpers.HighlightAt(F3, closing));
    }

    [Fact]
    public void F3_OpeningDoubleParen_HighlightsTheSameGroup()
    {
        int open = F3.IndexOf("((365", StringComparison.Ordinal);
        Assert.Equal("((365*(($I27/Hundred)+1)^(K$31)))", TestHelpers.HighlightAt(F3, open));
        Assert.Equal("(365*(($I27/Hundred)+1)^(K$31))", TestHelpers.HighlightAt(F3, open + 1));
    }

    [Fact]
    public void F3_Star_HighlightsTheProduct() =>
        Assert.Equal("365*(($I27/Hundred)+1)^(K$31)", TestHelpers.HighlightOn(F3, "*"));

    // ---- F4 -----------------------------------------------------------------------------------------

    [Fact]
    public void F4_SlashInE30C4_HighlightsE30C4() => Assert.Equal("E30/C4", TestHelpers.HighlightOn(F4, "/"));

    [Fact]
    public void F4_LastPlus_HighlightsTheWholeBody() =>
        Assert.Equal(Body(F4), TestHelpers.HighlightAt(F4, F4.LastIndexOf('+')));

    [Fact]
    public void F4_Operators()
    {
        Assert.Equal("Sheet8!E7+(E30/C4)^(D4/D5)", TestHelpers.HighlightOn(F4, "+", occurrence: 0));
        Assert.Equal("Sheet8!E7+(E30/C4)^(D4/D5)+D14", TestHelpers.HighlightOn(F4, "+", occurrence: 1));
        Assert.Equal("Sheet8!E7+(E30/C4)^(D4/D5)+D14-E21*E16", TestHelpers.HighlightOn(F4, "-"));
        Assert.Equal("E21*E16", TestHelpers.HighlightOn(F4, "*"));
        Assert.Equal("(E30/C4)^(D4/D5)", TestHelpers.HighlightOn(F4, "^"));
    }

    [Fact]
    public void F4_SheetPrefixHighlightsTheWholeReference()
    {
        Assert.Equal("Sheet8!E7", TestHelpers.HighlightOn(F4, "Sheet8"));
        Assert.Equal("Summary!C5", TestHelpers.HighlightOn(F4, "Summary", offset: 4));
    }

    // ---- F5 -----------------------------------------------------------------------------------------

    [Fact]
    public void F5_Offset_HighlightsTheWholeCall() =>
        Assert.Equal("OFFSET($I$19,0,0,1,Aqui_Per)", TestHelpers.HighlightOn(F5, "OFFSET"));

    [Fact]
    public void F5_AreaIsOneReference()
    {
        Assert.Equal("I$19:$I$19", TestHelpers.HighlightOn(F5, "I$19:$I$19"));
        Assert.Equal("I$19:$I$19", TestHelpers.HighlightOn(F5, "I$19:$I$19", offset: 4)); // the ':'
        var node = HitMap.Build(FormulaParser.Parse(F5)).NodeAt(F5.IndexOf("I$19:", StringComparison.Ordinal) + 4);
        var reference = Assert.IsType<ReferenceNode>(node);
        Assert.Equal(ReferenceForm.Area, reference.Form);
    }

    [Fact]
    public void F5_NamesAndOtherOperators()
    {
        Assert.Equal("Aqui_Per", TestHelpers.HighlightOn(F5, "Aqui_Per"));
        Assert.Equal("J$19*$I$28/Hundred", TestHelpers.HighlightOn(F5, "/"));
        Assert.Equal("J$19*$I$28", TestHelpers.HighlightOn(F5, "*"));
        Assert.Equal("SUM(I$19:$I$19)", TestHelpers.HighlightOn(F5, "SUM"));
        Assert.Equal(Body(F5), TestHelpers.HighlightAt(F5, F5.LastIndexOf('-')));
    }

    [Fact]
    public void F5_LinkCandidates()
    {
        var parsed = FormulaParser.Parse(F5);
        var map = HitMap.Build(parsed);
        Assert.True(HitMap.IsLinkCandidate(map.NodeAt(F5.IndexOf("OFFSET", StringComparison.Ordinal))!));
        Assert.True(HitMap.IsLinkCandidate(map.NodeAt(F5.IndexOf("Hundred", StringComparison.Ordinal))!));
        Assert.False(HitMap.IsLinkCandidate(map.NodeAt(F5.IndexOf('*'))!));
    }

    // ---- F6 -----------------------------------------------------------------------------------------

    [Fact]
    public void F6_QuotedSheetReference_HighlightsTheWholeReference()
    {
        Assert.Equal("'Inventory Forcast'!C$6", TestHelpers.HighlightOn(F6, "Forcast"));
        Assert.Equal("'Inventory Forcast'!C$6", TestHelpers.HighlightOn(F6, "'Inventory"));
        Assert.Equal("'Inventory Forcast'!C$6", TestHelpers.HighlightOn(F6, "C$6"));
    }

    [Fact]
    public void F6_ProductAndCall()
    {
        Assert.Equal("C11*'Inventory Forcast'!C$6", TestHelpers.HighlightOn(F6, "*"));
        Assert.Equal(Body(F6), TestHelpers.HighlightOn(F6, "IF"));
    }

    // ---- F7 -----------------------------------------------------------------------------------------

    [Fact]
    public void F7_Star_HighlightsEverythingAfterEquals() => Assert.Equal(Body(F7), TestHelpers.HighlightOn(F7, ")*(", offset: 1));

    [Fact]
    public void F7_GroupsAndSums()
    {
        Assert.Equal("SUM(Fin_P_FO!L39:L40)", TestHelpers.HighlightOn(F7, "SUM"));
        Assert.Equal("(Macro_FA!L183/Hundred)", TestHelpers.HighlightOn(F7, "(Macro_FA"));
        Assert.Equal("Macro_FA!L183/Hundred", TestHelpers.HighlightOn(F7, "/"));
        Assert.Equal("(SUM(Fin_P_FO!L39:L40)+Fin_P_FO!L26+Notes_FO!L113+Notes_FO!L101)", TestHelpers.HighlightAt(F7, 1));
        Assert.Equal("SUM(Fin_P_FO!L39:L40)+Fin_P_FO!L26+Notes_FO!L113+Notes_FO!L101",
            TestHelpers.HighlightAt(F7, F7.LastIndexOf('+')));
    }

    // ---- Parsing the video formulas is clean ------------------------------------------------------------

    [Theory]
    [InlineData(F1)]
    [InlineData(F2)]
    [InlineData(F3)]
    [InlineData(F4)]
    [InlineData(F5)]
    [InlineData(F6)]
    [InlineData(F7)]
    public void VideoFormulasParseWithoutErrors(string formula)
    {
        var parsed = FormulaParser.Parse(formula);
        Assert.False(parsed.HasErrors);
        Assert.Equal(Body(formula), parsed.Root!.Body!.GetText(formula));
    }

    // ---- colors (palette indexes in order of appearance) -------------------------------------------------

    private static int[] ColorIndexes(string formula, string root = "Sheet1") =>
        RefColorizer.Assign(FormulaParser.Parse(formula), root, _ => true).Nodes.Select(n => n.PaletteIndex).ToArray();

    private static string[] ColoredTexts(string formula) =>
        RefColorizer.Assign(FormulaParser.Parse(formula), "Sheet1", _ => true).Nodes.Select(n => n.Node.GetText(formula)).ToArray();

    [Fact]
    public void Colors_F4_WrapAroundAfterSeven() =>
        Assert.Equal(new[] { 0, 1, 2, 3, 4, 5, 6, 0, 1 }, ColorIndexes(F4));

    [Fact]
    public void Colors_F5_ByAppearance()
    {
        Assert.Equal(new[] { "I$19:$I$19", "J$19", "$I$28", "Hundred", "$I$19", "Aqui_Per" }, ColoredTexts(F5));
        Assert.Equal(new[] { 0, 1, 2, 3, 4, 5 }, ColorIndexes(F5));
    }

    [Fact]
    public void Colors_RepeatedReferenceReusesColor() => Assert.Equal(new[] { 0, 1, 0 }, ColorIndexes("=A1+B1+A1"));

    [Fact]
    public void Colors_F1_HundredAndK31()
    {
        Assert.Equal(new[] { "$I27", "Hundred", "K$31" }, ColoredTexts(F1));
        Assert.Equal(new[] { 0, 1, 2 }, ColorIndexes(F1));
    }

    [Fact]
    public void Colors_F2_RepeatedAcrossFormulaPartsKeepsFirstColor()
    {
        Assert.Equal(new[] { "K3", "$I26", "Hundred", "K$31" }, ColoredTexts(F2));
        Assert.Equal(new[] { 0, 1, 2, 3 }, ColorIndexes(F2));
    }

    [Fact]
    public void Colors_F6_SheetPrefixIsPartOfTheColoredNode()
    {
        var parsed = FormulaParser.Parse(F6);
        var colors = RefColorizer.Assign(parsed, "Sheet1", _ => true);
        var last = colors.Nodes[^1];
        Assert.Equal("'Inventory Forcast'!C$6", last.Node.GetText(F6));
        Assert.Equal(2, last.PaletteIndex);
        Assert.Equal(0x9900CCu, last.Rgb);
    }

    [Fact]
    public void Colors_F7_RepeatedHundredAndSheets()
    {
        Assert.Equal(new[] { 0, 1, 2, 3, 4, 5 }, ColorIndexes(F7));
    }
}
