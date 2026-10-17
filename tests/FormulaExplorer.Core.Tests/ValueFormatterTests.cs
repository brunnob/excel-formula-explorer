using System.Globalization;
using FormulaExplorer.Core.Values;

namespace FormulaExplorer.Core.Tests;

public class ValueFormatterTests
{
    private static string Num(double value) => ValueFormatter.Format(new EvalResult.Number(value));

    // ---- numbers -----------------------------------------------------------------------------------

    [Theory]
    [InlineData(0.461538461538462, "0.461538461538462")]
    [InlineData(6.0 / 13.0, "0.461538461538462")]
    [InlineData(1.00760985118373, "1.00760985118373")]
    [InlineData(1.062, "1.062")]
    [InlineData(100.0, "100")]
    [InlineData(0.126, "0.126")]
    [InlineData(1e20, "1E+20")]
    [InlineData(1.0 / 3.0, "0.333333333333333")]
    [InlineData(0.1 + 0.2, "0.3")]
    [InlineData(-1.5, "-1.5")]
    [InlineData(0.0, "0")]
    [InlineData(1e-5, "1E-05")]
    [InlineData(123456789012345678.0, "1.23456789012346E+17")]
    [InlineData(123456789012345.0, "123456789012345")]
    [InlineData(1234567.891, "1234567.891")]
    [InlineData(-100.0, "-100")]
    public void NumbersUseFifteenSignificantDigits(double value, string expected) => Assert.Equal(expected, Num(value));

    [Fact]
    public void NegativeZeroIsZero() => Assert.Equal("0", Num(-0.0));

    [Fact]
    public void NumbersIgnoreTheCurrentCulture()
    {
        var saved = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = new CultureInfo("pt-BR");
            Assert.Equal("1.062", Num(1.062));
            Assert.Equal("1.5E+20", Num(1.5e20));
            Assert.Equal("[2x1] 1.5, 2.5", ValueFormatter.Format(new EvalResult.ArrayValue(2, 1, new object?[,] { { 1.5 }, { 2.5 } })));
        }
        finally
        {
            CultureInfo.CurrentCulture = saved;
        }
    }

    // ---- text, bool, error, blank, unavailable -----------------------------------------------------------

    [Theory]
    [InlineData("On", "\"On\"")]
    [InlineData("", "\"\"")]
    [InlineData("a b", "\"a b\"")]
    [InlineData("say \"hi\"", "\"say \"hi\"\"")]
    public void TextIsQuoted(string value, string expected) =>
        Assert.Equal(expected, ValueFormatter.Format(new EvalResult.Text(value)));

    [Fact]
    public void TextOf255CharactersIsNotCut()
    {
        string value = new('x', 255);
        Assert.Equal($"\"{value}\"", ValueFormatter.Format(new EvalResult.Text(value)));
    }

    [Fact]
    public void LongerTextIsCutAt255AndEndsWithEllipsis()
    {
        string value = new string('x', 255) + "yz";
        string formatted = ValueFormatter.Format(new EvalResult.Text(value));
        Assert.Equal($"\"{new string('x', 255)}...\"", formatted);
        Assert.Equal(255 + 5, formatted.Length);
    }

    [Fact]
    public void ConstantsMatchTheContract()
    {
        Assert.Equal(20, ValueFormatter.MaxListedValues);
        Assert.Equal(255, ValueFormatter.MaxStringLength);
    }

    [Theory]
    [InlineData(true, "TRUE")]
    [InlineData(false, "FALSE")]
    public void Booleans(bool value, string expected) => Assert.Equal(expected, ValueFormatter.Format(new EvalResult.Bool(value)));

    [Theory]
    [InlineData(ExcelError.Null, "#NULL!")]
    [InlineData(ExcelError.Div0, "#DIV/0!")]
    [InlineData(ExcelError.Value, "#VALUE!")]
    [InlineData(ExcelError.Ref, "#REF!")]
    [InlineData(ExcelError.Name, "#NAME?")]
    [InlineData(ExcelError.Num, "#NUM!")]
    [InlineData(ExcelError.NA, "#N/A")]
    [InlineData(ExcelError.GettingData, "#GETTING_DATA")]
    [InlineData(ExcelError.Spill, "#SPILL!")]
    [InlineData(ExcelError.Connect, "#CONNECT!")]
    [InlineData(ExcelError.Blocked, "#BLOCKED!")]
    [InlineData(ExcelError.Unknown, "#UNKNOWN!")]
    [InlineData(ExcelError.Field, "#FIELD!")]
    [InlineData(ExcelError.Calc, "#CALC!")]
    [InlineData(ExcelError.Busy, "#BUSY!")]
    public void Errors(ExcelError error, string expected)
    {
        Assert.Equal(expected, ValueFormatter.ErrorText(error));
        Assert.Equal(expected, ValueFormatter.Format(new EvalResult.Error(error)));
    }

    [Fact]
    public void UndefinedErrorCodeFallsBackToValueError() => Assert.Equal("#VALUE!", ValueFormatter.ErrorText((ExcelError)12345));

    [Fact]
    public void BlankAndUnavailable()
    {
        Assert.Equal("[Blank]", ValueFormatter.Format(EvalResult.BlankValue));
        Assert.Equal("[Blank]", ValueFormatter.Format(new EvalResult.Blank()));
        Assert.Equal("[Unable to evaluate]", ValueFormatter.Format(new EvalResult.Unavailable("Too long to evaluate")));
    }

    // ---- scalars -------------------------------------------------------------------------------------------

    [Fact]
    public void BlankScalarDependsOnList()
    {
        Assert.Equal("[Blank]", ValueFormatter.FormatScalar(null, inList: false));
        Assert.Equal("", ValueFormatter.FormatScalar(null, inList: true));
        Assert.Equal("[Blank]", ValueFormatter.FormatScalar(DBNull.Value, inList: false));
        Assert.Equal("", ValueFormatter.FormatScalar(DBNull.Value, inList: true));
    }

    [Theory]
    [InlineData(1.062, "1.062")]
    [InlineData("x", "\"x\"")]
    [InlineData(true, "TRUE")]
    [InlineData(ExcelError.NA, "#N/A")]
    [InlineData(5, "5")]
    [InlineData(5L, "5")]
    [InlineData(2.5f, "2.5")]
    public void ScalarsOfAnyKind(object value, string expected)
    {
        Assert.Equal(expected, ValueFormatter.FormatScalar(value, inList: false));
        Assert.Equal(expected, ValueFormatter.FormatScalar(value, inList: true));
    }

    [Fact]
    public void DecimalAndDateScalars()
    {
        Assert.Equal("12.5", ValueFormatter.FormatScalar(12.5m, inList: false));
        Assert.Equal("44927", ValueFormatter.FormatScalar(new DateTime(2023, 1, 1), inList: false));
    }

    // ---- arrays ------------------------------------------------------------------------------------------------

    [Fact]
    public void ArrayIsListedRowMajor()
    {
        var array = new EvalResult.ArrayValue(2, 2, new object?[,] { { 1.0, 2.0 }, { 3.0, 4.0 } });
        Assert.Equal("[2x2] 1, 2, 3, 4", ValueFormatter.Format(array));
    }

    [Fact]
    public void ArrayColumn()
    {
        var array = new EvalResult.ArrayValue(3, 1, new object?[,] { { 1.0 }, { 2.0 }, { 3.0 } });
        Assert.Equal("[3x1] 1, 2, 3", ValueFormatter.Format(array));
    }

    [Fact]
    public void ArrayShowsEveryKindOfValue()
    {
        var array = new EvalResult.ArrayValue(1, 5, new object?[,] { { 1.5, "a", true, ExcelError.Div0, null } });
        Assert.Equal("[1x5] 1.5, \"a\", TRUE, #DIV/0!, ", ValueFormatter.Format(array));
    }

    [Fact]
    public void ArrayWithExactlyTwentyValuesHasNoEllipsis()
    {
        var values = new object?[1, 20];
        for (int i = 0; i < 20; i++) values[0, i] = (double)(i + 1);
        string expected = "[1x20] " + string.Join(", ", Enumerable.Range(1, 20));
        Assert.Equal(expected, ValueFormatter.Format(new EvalResult.ArrayValue(1, 20, values)));
    }

    [Fact]
    public void ArrayWithMoreThanTwentyValuesIsCut()
    {
        var values = new object?[1, 21];
        for (int i = 0; i < 21; i++) values[0, i] = (double)(i + 1);
        string expected = "[1x25] " + string.Join(", ", Enumerable.Range(1, 20)) + ", ...";
        Assert.Equal(expected, ValueFormatter.Format(new EvalResult.ArrayValue(1, 25, values)));
    }

    [Fact]
    public void HugeArrayWithTruncatedValuesKeepsRealDimensions()
    {
        // Evaluate("A:A*2") has 1,048,576 rows; the gateway copies only a few more than MaxListedValues cells.
        var values = new object?[21, 1];
        for (int i = 0; i < 21; i++) values[i, 0] = (double)(i * 2);
        string expected = "[1048576x1] " + string.Join(", ", Enumerable.Range(0, 20).Select(i => (i * 2).ToString())) + ", ...";
        Assert.Equal(expected, ValueFormatter.Format(new EvalResult.ArrayValue(1048576, 1, values)));
    }

    // ---- ranges -------------------------------------------------------------------------------------------------

    [Fact]
    public void RangeExampleFromTheVideo()
    {
        var range = new EvalResult.RangeValue("I19:J19", 1, 2, new object?[,] { { null, 0.252054794520548 } });
        Assert.Equal("I19:J19 containing\n[1x2] , 0.252054794520548", ValueFormatter.Format(range));
    }

    [Theory]
    [InlineData(1.062, "1.062")]
    [InlineData("x", "\"x\"")]
    [InlineData(true, "TRUE")]
    [InlineData(ExcelError.Ref, "#REF!")]
    [InlineData(null, "[Blank]")]
    public void SingleCellRangeIsTheScalarValue(object? value, string expected) =>
        Assert.Equal(expected, ValueFormatter.Format(new EvalResult.RangeValue("K31", 1, 1, new object?[,] { { value } })));

    [Fact]
    public void SingleCellRangeWithoutValuesIsBlank() =>
        Assert.Equal("[Blank]", ValueFormatter.Format(new EvalResult.RangeValue("K31", 1, 1, new object?[0, 0])));

    [Fact]
    public void MultiAreaRangeOfOneCellIsListedNotScalar()
    {
        var range = new EvalResult.RangeValue("A1,B2", 1, 1, new object?[,] { { 5.0 } });
        Assert.Equal("A1,B2 containing\n[1x1] 5, ...", ValueFormatter.Format(range));
    }

    [Fact]
    public void MultiAreaRangeAlwaysEndsWithEllipsis()
    {
        var range = new EvalResult.RangeValue("A1:A3,C1:C5", 3, 1, new object?[,] { { 1.0 }, { 2.0 }, { 3.0 } });
        Assert.Equal("A1:A3,C1:C5 containing\n[3x1] 1, 2, 3, ...", ValueFormatter.Format(range));
    }

    [Fact]
    public void RangeIsListedRowMajor()
    {
        var range = new EvalResult.RangeValue("A1:B2", 2, 2, new object?[,] { { 1.0, "b" }, { null, 4.0 } });
        Assert.Equal("A1:B2 containing\n[2x2] 1, \"b\", , 4", ValueFormatter.Format(range));
    }

    [Fact]
    public void LargeRangeIsCutAfterTwentyValues()
    {
        var values = new object?[20, 1];
        for (int i = 0; i < 20; i++) values[i, 0] = (double)(i + 1);
        var range = new EvalResult.RangeValue("A1:A30", 30, 1, values);
        string expected = "A1:A30 containing\n[30x1] " + string.Join(", ", Enumerable.Range(1, 20)) + ", ...";
        Assert.Equal(expected, ValueFormatter.Format(range));
    }

    [Fact]
    public void RangeWithExactlyTwentyCellsHasNoEllipsis()
    {
        var values = new object?[4, 5];
        for (int i = 0; i < 20; i++) values[i / 5, i % 5] = (double)(i + 1);
        var range = new EvalResult.RangeValue("A1:E4", 4, 5, values);
        Assert.Equal("A1:E4 containing\n[4x5] " + string.Join(", ", Enumerable.Range(1, 20)), ValueFormatter.Format(range));
    }

    [Fact]
    public void WholeColumnRangeUsesRealDimensions()
    {
        var values = new object?[21, 1];
        for (int i = 0; i < 21; i++) values[i, 0] = i % 2 == 0 ? null : (double)i;
        var range = new EvalResult.RangeValue("A1:A1048576", 1048576, 1, values);
        string text = ValueFormatter.Format(range);
        Assert.StartsWith("A1:A1048576 containing\n[1048576x1] , 1, , 3,", text);
        Assert.EndsWith(", ...", text);
    }
}
