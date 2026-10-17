using FormulaExplorer.Core.Syntax;

namespace FormulaExplorer.Core.Tests;

public class TokenizerTests
{
    // ---- non-formulas and the leading '=' ---------------------------------------------------------

    [Fact]
    public void EmptyTextHasNoTokens() => Assert.Empty(FormulaTokenizer.Tokenize(""));

    [Theory]
    [InlineData("abc")]
    [InlineData("123")]
    [InlineData(" =A1")]
    [InlineData("A1+B1")]
    [InlineData("'=A1")]
    public void TextWithoutLeadingEqualsIsOneUnknownToken(string text)
    {
        var tokens = FormulaTokenizer.Tokenize(text);
        var token = Assert.Single(tokens);
        Assert.Equal(TokenKind.Unknown, token.Kind);
        Assert.Equal(0, token.Start);
        Assert.Equal(text.Length, token.Length);
        Assert.Equal(text, token.Text);
    }

    [Fact]
    public void LeadingEqualsIsFormulaStart()
    {
        var tokens = FormulaTokenizer.Tokenize("=A1");
        Assert.Equal(new Token(TokenKind.FormulaStart, 0, 1, "="), tokens[0]);
        Assert.Equal(2, tokens.Count);
    }

    [Fact]
    public void EqualsAloneIsOnlyFormulaStart()
    {
        var tokens = FormulaTokenizer.Tokenize("=");
        Assert.Equal(TokenKind.FormulaStart, Assert.Single(tokens).Kind);
    }

    [Fact]
    public void SpansAreExact()
    {
        var tokens = FormulaTokenizer.Tokenize("=Sheet8!E7+(E30/C4)");
        Assert.Collection(tokens,
            t => Assert.Equal(new Token(TokenKind.FormulaStart, 0, 1, "="), t),
            t => Assert.Equal(new Token(TokenKind.Reference, 1, 9, "Sheet8!E7"), t),
            t => Assert.Equal(new Token(TokenKind.BinaryOperator, 10, 1, "+"), t),
            t => Assert.Equal(new Token(TokenKind.OpenParen, 11, 1, "("), t),
            t => Assert.Equal(new Token(TokenKind.Reference, 12, 3, "E30"), t),
            t => Assert.Equal(new Token(TokenKind.BinaryOperator, 15, 1, "/"), t),
            t => Assert.Equal(new Token(TokenKind.Reference, 16, 2, "C4"), t),
            t => Assert.Equal(new Token(TokenKind.CloseParen, 18, 1, ")"), t));
        Assert.Equal(19, tokens[^1].End);
    }

    // ---- numbers, strings, booleans, errors -------------------------------------------------------

    [Theory]
    [InlineData("=1", "N[1]")]
    [InlineData("=1.062", "N[1.062]")]
    [InlineData("=.5", "N[.5]")]
    [InlineData("=1E+20", "N[1E+20]")]
    [InlineData("=1.5e-3", "N[1.5e-3]")]
    [InlineData("=1E5", "N[1E5]")]
    [InlineData("=2E", "N[2] Nm[E]")]
    [InlineData("=1.+2", "N[1.] Bin[+] N[2]")]
    [InlineData("=-1", "Pre[-] N[1]")]
    [InlineData("=-.5E-3", "Pre[-] N[.5E-3]")]
    [InlineData("=123456789012345678", "N[123456789012345678]")]
    public void Numbers(string formula, string expected) => Assert.Equal(expected, TestHelpers.Tokens(formula));

    [Theory]
    [InlineData("=\"On\"", "S[\"On\"]")]
    [InlineData("=\"a\"\"b\"", "S[\"a\"\"b\"]")]
    [InlineData("=\"\"", "S[\"\"]")]
    [InlineData("=\"\"\"\"", "S[\"\"\"\"]")]
    [InlineData("=\"abc", "S[\"abc]")]
    [InlineData("=\"a,b(c)+\"", "S[\"a,b(c)+\"]")]
    [InlineData("=\"x\"&\"y\"", "S[\"x\"] Bin[&] S[\"y\"]")]
    [InlineData("=\"line1\nline2\"", "S[\"line1\nline2\"]")]
    public void Strings(string formula, string expected) => Assert.Equal(expected, TestHelpers.Tokens(formula));

    [Theory]
    [InlineData("=TRUE", "B[TRUE]")]
    [InlineData("=false", "B[false]")]
    [InlineData("=True", "B[True]")]
    [InlineData("=TRUE()", "F[TRUE] Open[(] Close[)]")]
    [InlineData("=TRUE!A1", "R[TRUE!A1]")]
    [InlineData("=TRUE[Col]", "SR[TRUE[Col]]")]
    [InlineData("=TRUEX", "Nm[TRUEX]")]
    [InlineData("=NOT(TRUE)", "F[NOT] Open[(] B[TRUE] Close[)]")]
    public void Booleans(string formula, string expected) => Assert.Equal(expected, TestHelpers.Tokens(formula));

    [Theory]
    [InlineData("#NULL!")]
    [InlineData("#DIV/0!")]
    [InlineData("#VALUE!")]
    [InlineData("#REF!")]
    [InlineData("#NAME?")]
    [InlineData("#NUM!")]
    [InlineData("#N/A")]
    [InlineData("#GETTING_DATA")]
    [InlineData("#SPILL!")]
    [InlineData("#CALC!")]
    [InlineData("#FIELD!")]
    [InlineData("#BLOCKED!")]
    [InlineData("#CONNECT!")]
    [InlineData("#UNKNOWN!")]
    [InlineData("#BUSY!")]
    [InlineData("#div/0!")]
    [InlineData("#n/a")]
    public void ErrorLiterals(string error) => Assert.Equal($"E[{error}]", TestHelpers.Tokens("=" + error));

    [Theory]
    [InlineData("=#FOO", "U[#] Nm[FOO]")]
    [InlineData("=#", "U[#]")]
    [InlineData("=IFERROR(A1,#N/A)", "F[IFERROR] Open[(] R[A1] Sep[,] E[#N/A] Close[)]")]
    public void NotAnErrorLiteral(string formula, string expected) => Assert.Equal(expected, TestHelpers.Tokens(formula));

    // ---- references ---------------------------------------------------------------------------------

    [Theory]
    [InlineData("A1")]
    [InlineData("$A$1")]
    [InlineData("a1")]
    [InlineData("$a1")]
    [InlineData("A$1")]
    [InlineData("K$31")]
    [InlineData("AA10")]
    [InlineData("XFD1")]
    [InlineData("XFD1048576")]
    [InlineData("A1:B2")]
    [InlineData("$A$1:$B$2")]
    [InlineData("I$19:$I$19")]
    [InlineData("a1:b2")]
    [InlineData("A:A")]
    [InlineData("$A:$C")]
    [InlineData("A:XFD")]
    [InlineData("1:1")]
    [InlineData("$3:$5")]
    [InlineData("1:1048576")]
    public void SingleReferenceToken(string address) => Assert.Equal($"R[{address}]", TestHelpers.Tokens("=" + address));

    [Theory]
    [InlineData("=XFE1", "Nm[XFE1]")]
    [InlineData("=A1048577", "Nm[A1048577]")]
    [InlineData("=ABCD1", "Nm[ABCD1]")]
    [InlineData("=TAX", "Nm[TAX]")]
    [InlineData("=A0", "Nm[A0]")]
    [InlineData("=R2D2", "Nm[R2D2]")]
    [InlineData("=A1.5", "Nm[A1.5]")]
    [InlineData("=A1B", "Nm[A1B]")]
    [InlineData("=XFE:XFE", "Nm[XFE] Bin[:] Nm[XFE]")]
    [InlineData("=1:1048577", "N[1] Bin[:] N[1048577]")]
    [InlineData("=A1:B2:C3", "R[A1:B2] Bin[:] R[C3]")]
    [InlineData("=A1:A", "R[A1] Bin[:] Nm[A]")]
    public void NamesThatLookLikeReferences(string formula, string expected) => Assert.Equal(expected, TestHelpers.Tokens(formula));

    [Theory]
    [InlineData("=Sheet8!E7")]
    [InlineData("='Inventory Forcast'!C$6")]
    [InlineData("='Jo''s Sheet'!A1")]
    [InlineData("=Sheet1:Sheet3!A1")]
    [InlineData("='Sheet 1:Sheet 3'!A1")]
    [InlineData("=[Book.xlsx]Sheet1!A1")]
    [InlineData("='[Book 1.xlsx]Sheet 1'!A1")]
    [InlineData("='C:\\Dir\\[Book.xlsx]Sheet1'!A1")]
    [InlineData("=Sheet1!#REF!")]
    [InlineData("=Sheet1!A1:B2")]
    [InlineData("=Sheet1!A:A")]
    [InlineData("=Sheet1!1:1")]
    [InlineData("=Fin_P_FO!L39:L40")]
    [InlineData("='1Sheet'!A1")]
    public void ReferenceWithPrefixIsOneToken(string formula) =>
        Assert.Equal($"R[{formula[1..]}]", TestHelpers.Tokens(formula));

    [Theory]
    [InlineData("=Sheet1!A1:Sheet1!B2", "R[Sheet1!A1] Bin[:] R[Sheet1!B2]")]
    [InlineData("=A1:INDEX(B:B,3)", "R[A1] Bin[:] F[INDEX] Open[(] R[B:B] Sep[,] N[3] Close[)]")]
    [InlineData("=Name1:Name2", "Nm[Name1] Bin[:] Nm[Name2]")]
    [InlineData("=(A1):B2", "Open[(] R[A1] Close[)] Bin[:] R[B2]")]
    [InlineData("=OFFSET(A1,0,0):C3", "F[OFFSET] Open[(] R[A1] Sep[,] N[0] Sep[,] N[0] Close[)] Bin[:] R[C3]")]
    public void ColonBetweenOperandsThatAreNotOneToken(string formula, string expected) =>
        Assert.Equal(expected, TestHelpers.Tokens(formula));

    // ---- names -----------------------------------------------------------------------------------------

    [Theory]
    [InlineData("=Hundred", "Nm[Hundred]")]
    [InlineData("=Aqui_Per", "Nm[Aqui_Per]")]
    [InlineData("=_x", "Nm[_x]")]
    [InlineData("=\\Name", "Nm[\\Name]")]
    [InlineData("=Is.Valid?", "Nm[Is.Valid?]")]
    [InlineData("=Sheet1!MyName", "Nm[Sheet1!MyName]")]
    [InlineData("='My Sheet'!MyName", "Nm['My Sheet'!MyName]")]
    [InlineData("=Book1.xlsx!Name", "Nm[Book1.xlsx!Name]")]
    [InlineData("=[Book.xlsx]Sheet!Name", "Nm[[Book.xlsx]Sheet!Name]")]
    [InlineData("=[Book.xlsx]!Name", "Nm[[Book.xlsx]!Name]")]
    [InlineData("=SUM(Hundred)", "F[SUM] Open[(] Nm[Hundred] Close[)]")]
    [InlineData("=Sheet1!", "Nm[Sheet1] U[!]")]
    [InlineData("='abc", "U['] Nm[abc]")]
    public void Names(string formula, string expected) => Assert.Equal(expected, TestHelpers.Tokens(formula));

    // ---- structured references -------------------------------------------------------------------------

    [Theory]
    [InlineData("Table1[Col]")]
    [InlineData("Table1[[#Headers],[Col]]")]
    [InlineData("Table1[@Col]")]
    [InlineData("Table1[]")]
    [InlineData("[@Col]")]
    [InlineData("[@[Col 1]]")]
    [InlineData("Table1[[Col1]:[Col2]]")]
    [InlineData("Table1[Col'[1']]")]
    [InlineData("Table1['#Col]")]
    [InlineData("Table1[#This Row]")]
    [InlineData("tbl.Sales_1[Net Amount]")]
    public void StructuredReferenceIsOneToken(string reference) =>
        Assert.Equal($"SR[{reference}]", TestHelpers.Tokens("=" + reference));

    [Theory]
    [InlineData("=SUM(Table1[Col])", "F[SUM] Open[(] SR[Table1[Col]] Close[)]")]
    [InlineData("=[@A]*[@B]", "SR[[@A]] Bin[*] SR[[@B]]")]
    [InlineData("=Table1[Col", "Nm[Table1] U[[] Nm[Col]")]
    [InlineData("=SUM(Table1[[#Totals],[A]:[B]],1)", "F[SUM] Open[(] SR[Table1[[#Totals],[A]:[B]]] Sep[,] N[1] Close[)]")]
    public void StructuredReferenceInContext(string formula, string expected) =>
        Assert.Equal(expected, TestHelpers.Tokens(formula));

    // ---- functions -------------------------------------------------------------------------------------

    [Theory]
    [InlineData("=SUM(A1)", "F[SUM] Open[(] R[A1] Close[)]")]
    [InlineData("=NORM.S.DIST(0)", "F[NORM.S.DIST] Open[(] N[0] Close[)]")]
    [InlineData("=_xlfn.XLOOKUP(1)", "F[_xlfn.XLOOKUP] Open[(] N[1] Close[)]")]
    [InlineData("=_xll.MyFunc()", "F[_xll.MyFunc] Open[(] Close[)]")]
    [InlineData("=LOG10(100)", "F[LOG10] Open[(] N[100] Close[)]")]
    [InlineData("=LOG10", "R[LOG10]")]
    [InlineData("=sum(a1)", "F[sum] Open[(] R[a1] Close[)]")]
    [InlineData("=IF(A1,SUM(B1),0)", "F[IF] Open[(] R[A1] Sep[,] F[SUM] Open[(] R[B1] Close[)] Sep[,] N[0] Close[)]")]
    [InlineData("=SUM (A1)", "Nm[SUM] Isect[ ] Open[(] R[A1] Close[)]")]
    public void Functions(string formula, string expected) => Assert.Equal(expected, TestHelpers.Tokens(formula));

    // ---- operators -------------------------------------------------------------------------------------

    [Theory]
    [InlineData("+")]
    [InlineData("-")]
    [InlineData("*")]
    [InlineData("/")]
    [InlineData("^")]
    [InlineData("&")]
    [InlineData("=")]
    [InlineData("<>")]
    [InlineData("<")]
    [InlineData(">")]
    [InlineData("<=")]
    [InlineData(">=")]
    public void BinaryOperators(string op) => Assert.Equal($"R[A1] Bin[{op}] R[B1]", TestHelpers.Tokens($"=A1{op}B1"));

    [Theory]
    [InlineData("=-A1", "Pre[-] R[A1]")]
    [InlineData("=+A1", "Pre[+] R[A1]")]
    [InlineData("=--A1", "Pre[-] Pre[-] R[A1]")]
    [InlineData("=1+-2", "N[1] Bin[+] Pre[-] N[2]")]
    [InlineData("=1--2", "N[1] Bin[-] Pre[-] N[2]")]
    [InlineData("=(-1)", "Open[(] Pre[-] N[1] Close[)]")]
    [InlineData("=SUM(-1,+2)", "F[SUM] Open[(] Pre[-] N[1] Sep[,] Pre[+] N[2] Close[)]")]
    [InlineData("=2^-2", "N[2] Bin[^] Pre[-] N[2]")]
    [InlineData("=1>-1", "N[1] Bin[>] Pre[-] N[1]")]
    [InlineData("=-(-A1)", "Pre[-] Open[(] Pre[-] R[A1] Close[)]")]
    [InlineData("={-1,2}", "ArrOpen[{] Pre[-] N[1] ColSep[,] N[2] ArrClose[}]")]
    [InlineData("=@A1", "Pre[@] R[A1]")]
    [InlineData("=SUM(@A1:A3)", "F[SUM] Open[(] Pre[@] R[A1:A3] Close[)]")]
    [InlineData("=A1@", "R[A1] U[@]")]
    [InlineData("=A1)-1", "R[A1] Close[)] Bin[-] N[1]")]
    [InlineData("=5%-1", "N[5] Post[%] Bin[-] N[1]")]
    public void PrefixVersusBinary(string formula, string expected) => Assert.Equal(expected, TestHelpers.Tokens(formula));

    [Theory]
    [InlineData("=A1%", "R[A1] Post[%]")]
    [InlineData("=50%", "N[50] Post[%]")]
    [InlineData("=A1%%", "R[A1] Post[%] Post[%]")]
    [InlineData("=1%+2", "N[1] Post[%] Bin[+] N[2]")]
    [InlineData("=-A1%", "Pre[-] R[A1] Post[%]")]
    [InlineData("=(A1)%", "Open[(] R[A1] Close[)] Post[%]")]
    [InlineData("=A1 %", "R[A1] Ws[ ] Post[%]")]
    [InlineData("=%", "U[%]")]
    [InlineData("=1+%", "N[1] Bin[+] U[%]")]
    [InlineData("=A1#", "R[A1] Post[#]")]
    [InlineData("=SUM(A1#)", "F[SUM] Open[(] R[A1] Post[#] Close[)]")]
    [InlineData("=Name#", "Nm[Name] Post[#]")]
    [InlineData("=(A1)#", "Open[(] R[A1] Close[)] Post[#]")]
    [InlineData("=A1#+1", "R[A1] Post[#] Bin[+] N[1]")]
    [InlineData("=A1 #", "R[A1] Ws[ ] U[#]")]
    public void PostfixOperators(string formula, string expected) => Assert.Equal(expected, TestHelpers.Tokens(formula));

    // ---- whitespace and intersection -------------------------------------------------------------------

    [Theory]
    [InlineData("= A1", "Ws[ ] R[A1]")]
    [InlineData("=A1 + B1", "R[A1] Ws[ ] Bin[+] Ws[ ] R[B1]")]
    [InlineData("=A1 ", "R[A1] Ws[ ]")]
    [InlineData("=A1\t-B1", "R[A1] Ws[\t] Bin[-] R[B1]")]
    [InlineData("=IF(A1, B1)", "F[IF] Open[(] R[A1] Sep[,] Ws[ ] R[B1] Close[)]")]
    [InlineData("=SUM(A1 )", "F[SUM] Open[(] R[A1] Ws[ ] Close[)]")]
    [InlineData("= 1", "Ws[ ] N[1]")]
    [InlineData("=A1 1", "R[A1] Ws[ ] N[1]")]
    [InlineData("=A1 TRUE", "R[A1] Ws[ ] B[TRUE]")]
    [InlineData("=A1 \"x\"", "R[A1] Ws[ ] S[\"x\"]")]
    [InlineData("=1 2", "N[1] Ws[ ] N[2]")]
    [InlineData("=SUM(\n1,\r\n2)", "F[SUM] Open[(] Ws[\n] N[1] Sep[,] Ws[\r\n] N[2] Close[)]")]
    public void InsignificantWhitespace(string formula, string expected) => Assert.Equal(expected, TestHelpers.Tokens(formula));

    [Theory]
    [InlineData("=A1 B1", "R[A1] Isect[ ] R[B1]")]
    [InlineData("=A1   B1", "R[A1] Isect[   ] R[B1]")]
    [InlineData("=A1\r\nB1", "R[A1] Isect[\r\n] R[B1]")]
    [InlineData("=A1\tB1", "R[A1] Isect[\t] R[B1]")]
    [InlineData("=A1:B2 B2:C3", "R[A1:B2] Isect[ ] R[B2:C3]")]
    [InlineData("=SUM(A1:B2 B2:C3)", "F[SUM] Open[(] R[A1:B2] Isect[ ] R[B2:C3] Close[)]")]
    [InlineData("=(A1:B2) (B2:C3)", "Open[(] R[A1:B2] Close[)] Isect[ ] Open[(] R[B2:C3] Close[)]")]
    [InlineData("=A1 SUM(B1)", "R[A1] Isect[ ] F[SUM] Open[(] R[B1] Close[)]")]
    [InlineData("=Name1 Name2", "Nm[Name1] Isect[ ] Nm[Name2]")]
    [InlineData("=Table1[Col] A1", "SR[Table1[Col]] Isect[ ] R[A1]")]
    [InlineData("=A1 B1 C1", "R[A1] Isect[ ] R[B1] Isect[ ] R[C1]")]
    [InlineData("=A1# B1", "R[A1] Post[#] Isect[ ] R[B1]")]
    [InlineData("=A:A 1:1", "R[A:A] Isect[ ] R[1:1]")]
    [InlineData("=SUM(A1) B1", "F[SUM] Open[(] R[A1] Close[)] Isect[ ] R[B1]")]
    public void IntersectionOperator(string formula, string expected) => Assert.Equal(expected, TestHelpers.Tokens(formula));

    // ---- commas, semicolons, braces --------------------------------------------------------------------

    [Theory]
    [InlineData("=SUM(A1,B1)", "F[SUM] Open[(] R[A1] Sep[,] R[B1] Close[)]")]
    [InlineData("=(A1,B1)", "Open[(] R[A1] Union[,] R[B1] Close[)]")]
    [InlineData("=A1,B1", "R[A1] Union[,] R[B1]")]
    [InlineData("=SUM((A1,B1),C1)", "F[SUM] Open[(] Open[(] R[A1] Union[,] R[B1] Close[)] Sep[,] R[C1] Close[)]")]
    [InlineData("=INDEX((A1:B2,C1:D2),1)", "F[INDEX] Open[(] Open[(] R[A1:B2] Union[,] R[C1:D2] Close[)] Sep[,] N[1] Close[)]")]
    [InlineData("=IF(A1,,2)", "F[IF] Open[(] R[A1] Sep[,] Sep[,] N[2] Close[)]")]
    [InlineData("=SUM(A1,)", "F[SUM] Open[(] R[A1] Sep[,] Close[)]")]
    [InlineData("=SUM(,)", "F[SUM] Open[(] Sep[,] Close[)]")]
    [InlineData("=SUM(A1))", "F[SUM] Open[(] R[A1] Close[)] Close[)]")]
    [InlineData("=A1;B1", "R[A1] U[;] R[B1]")]
    public void CommaMeaningDependsOnContext(string formula, string expected) =>
        Assert.Equal(expected, TestHelpers.Tokens(formula));

    [Theory]
    [InlineData("={1,2;3,4}", "ArrOpen[{] N[1] ColSep[,] N[2] RowSep[;] N[3] ColSep[,] N[4] ArrClose[}]")]
    [InlineData("=SUM({1,2},3)", "F[SUM] Open[(] ArrOpen[{] N[1] ColSep[,] N[2] ArrClose[}] Sep[,] N[3] Close[)]")]
    [InlineData("={\"a\",\"b\"}", "ArrOpen[{] S[\"a\"] ColSep[,] S[\"b\"] ArrClose[}]")]
    [InlineData("={TRUE,#N/A}", "ArrOpen[{] B[TRUE] ColSep[,] E[#N/A] ArrClose[}]")]
    [InlineData("={1,2)", "ArrOpen[{] N[1] ColSep[,] N[2] Close[)]")]
    [InlineData("={1}+1", "ArrOpen[{] N[1] ArrClose[}] Bin[+] N[1]")]
    [InlineData("={1;2}", "ArrOpen[{] N[1] RowSep[;] N[2] ArrClose[}]")]
    public void ArrayConstants(string formula, string expected) => Assert.Equal(expected, TestHelpers.Tokens(formula));

    // ---- unknown characters and robustness ---------------------------------------------------------------

    [Theory]
    [InlineData("=~", "U[~]")]
    [InlineData("=A1~B1", "R[A1] U[~] R[B1]")]
    [InlineData("=$", "U[$]")]
    [InlineData("=$A", "U[$] Nm[A]")]
    [InlineData("=?", "U[?]")]
    [InlineData("=]", "U[]]")]
    [InlineData("=!", "U[!]")]
    [InlineData("=[", "U[[]")]
    public void UnknownCharactersAreSingleTokens(string formula, string expected) =>
        Assert.Equal(expected, TestHelpers.Tokens(formula));

    [Theory]
    [InlineData("=")]
    [InlineData("=1/(((($I27/Hundred)+1)^(K$31)))")]
    [InlineData("=IF(K3=\"On\",1/(((-($I26/Hundred)+1)^(K$31))),0)")]
    [InlineData("=SUM(I$19:$I$19)+J$19*$I$28/Hundred-SUM(OFFSET($I$19,0,0,1,Aqui_Per))")]
    [InlineData("=IF(C3,0,C11*'Inventory Forcast'!C$6)")]
    [InlineData("=SUM(Table1[[#Headers],[A]:[B]])\n+ 'It''s'!A1:B2 C3 ")]
    [InlineData("=\"unterminated")]
    [InlineData("='unterminated!A1")]
    [InlineData("=A1 \u00A0 B1 \u2003")]
    [InlineData("=\uD83D\uDE00+A1")]
    [InlineData("=ñ+Çá!A1")]
    [InlineData("=}{)(,;:%#@!$&")]
    public void TokenTextsConcatenateToTheInputAndAreContiguous(string formula)
    {
        var tokens = FormulaTokenizer.Tokenize(formula);
        Assert.Equal(formula, TestHelpers.JoinTokenTexts(tokens));
        int expectedStart = 0;
        foreach (var token in tokens)
        {
            Assert.Equal(expectedStart, token.Start);
            Assert.True(token.Length > 0);
            Assert.Equal(formula.Substring(token.Start, token.Length), token.Text);
            expectedStart = token.End;
        }
        Assert.Equal(formula.Length, expectedStart);
    }

    [Fact]
    public void UnicodeSpacesAreWhitespace()
    {
        Assert.Equal("R[A1] Isect[\u00A0] R[B1]", TestHelpers.Tokens("=A1\u00A0B1"));
    }

    [Fact]
    public void UnicodeLettersFormNames()
    {
        Assert.Equal("Nm[Preço_Total]", TestHelpers.Tokens("=Preço_Total"));
    }

    [Fact]
    public void LongFormulaDoesNotRecurseOrSlowDown()
    {
        string formula = "=" + string.Join("+", Enumerable.Range(1, 1500).Select(i => "A" + i));
        var tokens = FormulaTokenizer.Tokenize(formula);
        Assert.Equal(1 + 1500 + 1499, tokens.Count);
    }
}
