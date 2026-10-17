using System.Runtime.InteropServices;
using FormulaExplorer.AddIn.Excel;
using FormulaExplorer.Core.Values;

namespace FormulaExplorer.AddIn.Tests;

/// <summary>Pure tests (no Excel) of ExcelGateway.ConvertValue and of the helpers the gateway is built from.</summary>
public class ExcelGatewayConvertTests
{
    private static int ErrorInt(int code) => unchecked((int)(0x800A0000u | (uint)code));

    private static Array OneBased2D(int rows, int columns, Func<int, int, object?> cell)
    {
        var array = Array.CreateInstance(typeof(object), new[] { rows, columns }, new[] { 1, 1 });
        for (var r = 1; r <= rows; r++)
            for (var c = 1; c <= columns; c++)
                array.SetValue(cell(r, c), r, c);
        return array;
    }

    // ---- scalars ---------------------------------------------------------------------------------------------

    [Fact]
    public void Null_DBNull_and_Missing_are_blank()
    {
        Assert.Same(EvalResult.BlankValue, ExcelGateway.ConvertValue(null));
        Assert.Same(EvalResult.BlankValue, ExcelGateway.ConvertValue(DBNull.Value));
        Assert.Same(EvalResult.BlankValue, ExcelGateway.ConvertValue(System.Reflection.Missing.Value));
    }

    [Fact]
    public void Double_string_and_bool_map_directly()
    {
        Assert.Equal(new EvalResult.Number(1.062), ExcelGateway.ConvertValue(1.062));
        Assert.Equal(new EvalResult.Text("On"), ExcelGateway.ConvertValue("On"));
        Assert.Equal(new EvalResult.Text(""), ExcelGateway.ConvertValue(""));
        Assert.Equal(new EvalResult.Bool(true), ExcelGateway.ConvertValue(true));
        Assert.Equal(new EvalResult.Bool(false), ExcelGateway.ConvertValue(false));
    }

    [Theory]
    [InlineData(2000, ExcelError.Null)]
    [InlineData(2007, ExcelError.Div0)]
    [InlineData(2015, ExcelError.Value)]
    [InlineData(2023, ExcelError.Ref)]
    [InlineData(2029, ExcelError.Name)]
    [InlineData(2036, ExcelError.Num)]
    [InlineData(2042, ExcelError.NA)]
    [InlineData(2043, ExcelError.GettingData)]
    [InlineData(2045, ExcelError.Spill)]
    [InlineData(2050, ExcelError.Calc)]
    public void Int_with_the_0x800A_high_word_is_an_error(int code, ExcelError expected)
    {
        Assert.Equal(new EvalResult.Error(expected), ExcelGateway.ConvertValue(ErrorInt(code)));
    }

    [Fact]
    public void Error_int_matches_the_value_Excel_returns_for_div0()
    {
        // 0x800A07D7 = -2146826281, what Worksheet.Evaluate("1/0") returns.
        Assert.Equal(new EvalResult.Error(ExcelError.Div0), ExcelGateway.ConvertValue(-2146826281));
    }

    [Fact]
    public void Unknown_error_code_is_the_unknown_error()
    {
        Assert.Equal(new EvalResult.Error(ExcelError.Unknown), ExcelGateway.ConvertValue(ErrorInt(2999)));
    }

    [Fact]
    public void Ordinary_integers_are_numbers_even_when_negative()
    {
        Assert.Equal(new EvalResult.Number(42), ExcelGateway.ConvertValue(42));
        Assert.Equal(new EvalResult.Number(-5), ExcelGateway.ConvertValue(-5));
        Assert.Equal(new EvalResult.Number(0), ExcelGateway.ConvertValue(0));
        Assert.Equal(new EvalResult.Number(unchecked((int)0x7FFFFFFF)), ExcelGateway.ConvertValue(int.MaxValue));
        Assert.Equal(new EvalResult.Number(int.MinValue), ExcelGateway.ConvertValue(int.MinValue));
    }

    [Fact]
    public void Other_numeric_types_are_numbers()
    {
        Assert.Equal(new EvalResult.Number(7), ExcelGateway.ConvertValue((short)7));
        Assert.Equal(new EvalResult.Number(8_000_000_000), ExcelGateway.ConvertValue(8_000_000_000L));
        Assert.Equal(new EvalResult.Number(1.5), ExcelGateway.ConvertValue(1.5f));
        Assert.Equal(new EvalResult.Number(9), ExcelGateway.ConvertValue((byte)9));
        Assert.Equal(new EvalResult.Number(4_000_000_000), ExcelGateway.ConvertValue(4_000_000_000u));
    }

    [Fact]
    public void Decimal_currency_is_a_number()
    {
        Assert.Equal(new EvalResult.Number(12.5), ExcelGateway.ConvertValue(12.5m));
    }

    [Fact]
    public void DateTime_is_its_OLE_automation_serial_number()
    {
        var date = new DateTime(2020, 1, 2);
        Assert.Equal(new EvalResult.Number(43832), ExcelGateway.ConvertValue(date));
        Assert.Equal(new EvalResult.Number(date.AddHours(12).ToOADate()), ExcelGateway.ConvertValue(date.AddHours(12)));
    }

    [Fact]
    public void Unsupported_types_are_unavailable_not_exceptions()
    {
        var result = ExcelGateway.ConvertValue(new object());
        Assert.IsType<EvalResult.Unavailable>(result);
    }

    // ---- arrays ----------------------------------------------------------------------------------------------

    [Fact]
    public void Small_one_based_2D_array_keeps_dimensions_and_is_copied_zero_based()
    {
        var com = OneBased2D(2, 3, (r, c) => (r, c) switch
        {
            (1, 1) => 1.5,
            (1, 2) => "x",
            (1, 3) => true,
            (2, 1) => ErrorInt(2007),
            (2, 2) => null,
            _ => (object)DBNull.Value,
        });

        var result = Assert.IsType<EvalResult.ArrayValue>(ExcelGateway.ConvertValue(com));

        Assert.Equal(2, result.Rows);
        Assert.Equal(3, result.Columns);
        Assert.Equal(0, result.Values.GetLowerBound(0));
        Assert.Equal(1.5, result.Values[0, 0]);
        Assert.Equal("x", result.Values[0, 1]);
        Assert.Equal(true, result.Values[0, 2]);
        Assert.Equal(ExcelError.Div0, result.Values[1, 0]);
        Assert.Null(result.Values[1, 1]);
        Assert.Null(result.Values[1, 2]);
    }

    [Fact]
    public void One_dimensional_array_is_a_single_row_whatever_its_lower_bound()
    {
        // Worksheet.Evaluate("OFFSET(A1,0,0,1,2)+0") returns a 1-based object[*] of length 2.
        var com = Array.CreateInstance(typeof(object), new[] { 2 }, new[] { 1 });
        com.SetValue(2.0, 1);
        com.SetValue("txt", 2);

        var result = Assert.IsType<EvalResult.ArrayValue>(ExcelGateway.ConvertValue(com));

        Assert.Equal(1, result.Rows);
        Assert.Equal(2, result.Columns);
        Assert.Equal(2.0, result.Values[0, 0]);
        Assert.Equal("txt", result.Values[0, 1]);
    }

    [Fact]
    public void Zero_based_arrays_work_too()
    {
        var result = Assert.IsType<EvalResult.ArrayValue>(ExcelGateway.ConvertValue(new object?[,] { { 1.0, 2.0 }, { 3.0, 4.0 } }));
        Assert.Equal((2, 2), (result.Rows, result.Columns));
        Assert.Equal(4.0, result.Values[1, 1]);
    }

    [Fact]
    public void Huge_column_array_copies_only_the_leading_cells()
    {
        // Evaluate("A:A*2") is a 1,048,576 x 1 array.
        var com = OneBased2D(1_048_576, 1, (r, _) => (double)r);

        var result = Assert.IsType<EvalResult.ArrayValue>(ExcelGateway.ConvertValue(com));

        Assert.Equal(1_048_576, result.Rows);
        Assert.Equal(1, result.Columns);
        Assert.Equal(ValueFormatter.MaxListedValues + 1, result.Values.Length);
        Assert.Equal(1.0, result.Values[0, 0]);
        Assert.Equal((double)(ValueFormatter.MaxListedValues + 1), result.Values[ValueFormatter.MaxListedValues, 0]);
    }

    [Fact]
    public void Wide_row_array_copies_the_leading_cells_of_the_row()
    {
        var com = OneBased2D(1, 5_000, (_, c) => (double)c);

        var result = Assert.IsType<EvalResult.ArrayValue>(ExcelGateway.ConvertValue(com));

        Assert.Equal((1, 5_000), (result.Rows, result.Columns));
        Assert.Equal((1, ValueFormatter.MaxListedValues + 1), (result.Values.GetLength(0), result.Values.GetLength(1)));
        Assert.Equal(21.0, result.Values[0, 20]);
    }

    [Fact]
    public void Wide_block_keeps_whole_rows_so_it_can_be_indexed_with_the_real_column_count()
    {
        var com = OneBased2D(100, 10, (r, c) => (double)((r - 1) * 10 + c));

        var result = Assert.IsType<EvalResult.ArrayValue>(ExcelGateway.ConvertValue(com));

        Assert.Equal((100, 10), (result.Rows, result.Columns));
        Assert.Equal(10, result.Values.GetLength(1));
        Assert.True(result.Values.Length >= ValueFormatter.MaxListedValues + 1);
        for (var i = 0; i < ValueFormatter.MaxListedValues + 1; i++)
            Assert.Equal((double)(i + 1), result.Values[i / 10, i % 10]);
    }

    [Fact]
    public void Empty_array_has_no_values()
    {
        var result = Assert.IsType<EvalResult.ArrayValue>(ExcelGateway.ConvertValue(new object?[0, 0]));
        Assert.Equal((0, 0), (result.Rows, result.Columns));
        Assert.Empty(result.Values);
    }

    [Fact]
    public void Three_dimensional_array_is_unavailable()
    {
        Assert.IsType<EvalResult.Unavailable>(ExcelGateway.ConvertValue(new object[1, 1, 1]));
    }

    [Theory]
    [InlineData(1, 1, 20, 1, 1)]
    [InlineData(3, 1, 20, 3, 1)]
    [InlineData(1_048_576, 1, 20, 20, 1)]
    [InlineData(1, 1_000, 20, 1, 20)]
    [InlineData(100, 3, 20, 7, 3)]
    [InlineData(100, 20, 20, 1, 20)]
    [InlineData(100, 19, 20, 2, 19)]
    [InlineData(2, 2, 20, 2, 2)]
    [InlineData(0, 5, 20, 0, 0)]
    public void LeadingBlock_holds_the_first_cells_in_row_major_order(int rows, int columns, int cap, int expectedRows, int expectedColumns)
    {
        Assert.Equal((expectedRows, expectedColumns), ExcelValueConverter.LeadingBlock(rows, columns, cap));
    }

    [Fact]
    public void ToZeroBasedBlock_converts_a_Value2_read()
    {
        var single = ExcelValueConverter.ToZeroBasedBlock(2.5, 1, 1);
        Assert.Equal(2.5, single[0, 0]);

        var block = ExcelValueConverter.ToZeroBasedBlock(OneBased2D(2, 2, (r, c) => r == 2 && c == 2 ? ErrorInt(2042) : (object)(double)(r * 10 + c)), 2, 2);
        Assert.Equal(11.0, block[0, 0]);
        Assert.Equal(ExcelError.NA, block[1, 1]);

        Assert.Null(ExcelValueConverter.ToZeroBasedBlock(null, 1, 1)[0, 0]);
    }

    // ---- COM failure classification ---------------------------------------------------------------------------

    [Theory]
    [InlineData(unchecked((int)0x800AC472), true)] // VBA_E_IGNORE: cell edit mode
    [InlineData(unchecked((int)0x80010001), true)] // RPC_E_CALL_REJECTED
    [InlineData(unchecked((int)0x8001010A), true)] // RPC_E_SERVERCALL_RETRYLATER
    [InlineData(unchecked((int)0x800A03EC), false)] // generic "Exception from HRESULT"
    [InlineData(unchecked((int)0x8002000B), false)] // DISP_E_BADINDEX
    public void Busy_detection(int hresult, bool busy)
    {
        Assert.Equal(busy, ComErrors.IsBusy(new COMException("x", hresult)));
    }

    // ---- address math -----------------------------------------------------------------------------------------

    [Fact]
    public void Parses_cells_areas_columns_rows_and_unions()
    {
        Assert.True(ExcelAddressMath.TryParseAreas("$K$31", out var cell));
        Assert.Equal(new[] { new AreaSpan(31, 31, 11, 11) }, cell);

        Assert.True(ExcelAddressMath.TryParseAreas("$I$19:$J$19", out var area));
        Assert.Equal(new[] { new AreaSpan(19, 19, 9, 10) }, area);

        Assert.True(ExcelAddressMath.TryParseAreas("$B$5:$A$2", out var reversed));
        Assert.Equal(new[] { new AreaSpan(2, 5, 1, 2) }, reversed);

        Assert.True(ExcelAddressMath.TryParseAreas("$A:$C", out var columns));
        Assert.Equal(new[] { new AreaSpan(1, ExcelAddressMath.MaxRows, 1, 3) }, columns);

        Assert.True(ExcelAddressMath.TryParseAreas("$3:$5", out var rows));
        Assert.Equal(new[] { new AreaSpan(3, 5, 1, ExcelAddressMath.MaxColumns) }, rows);

        Assert.True(ExcelAddressMath.TryParseAreas("$A$1:$B$2,$D$4", out var union));
        Assert.Equal(new[] { new AreaSpan(1, 2, 1, 2), new AreaSpan(4, 4, 4, 4) }, union);

        Assert.True(ExcelAddressMath.TryParseAreas("$A$1:$B$2;$D$4", out var localizedUnion));
        Assert.Equal(2, localizedUnion.Count);

        Assert.True(ExcelAddressMath.TryParseAreas("XFD1048576", out var last));
        Assert.Equal(new[] { new AreaSpan(1_048_576, 1_048_576, 16_384, 16_384) }, last);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("Sheet1!A1")]
    [InlineData("A0")]
    [InlineData("A1048577")]
    [InlineData("XFE1")]
    [InlineData("ABCD1")]
    [InlineData("A1:B")]
    [InlineData("A:5")]
    [InlineData("A1:B2:C3")]
    [InlineData("$$A1")]
    [InlineData("A1B")]
    public void Rejects_malformed_addresses(string address)
    {
        Assert.False(ExcelAddressMath.TryParseAreas(address, out _));
    }

    [Fact]
    public void Column_names_round_trip()
    {
        Assert.Equal("A", ExcelAddressMath.ColumnName(1));
        Assert.Equal("Z", ExcelAddressMath.ColumnName(26));
        Assert.Equal("AA", ExcelAddressMath.ColumnName(27));
        Assert.Equal("AZ", ExcelAddressMath.ColumnName(52));
        Assert.Equal("BA", ExcelAddressMath.ColumnName(53));
        Assert.Equal("XFD", ExcelAddressMath.ColumnName(16_384));
        for (var c = 1; c <= ExcelAddressMath.MaxColumns; c += 37)
            Assert.Equal(c, ExcelAddressMath.ColumnNumber(ExcelAddressMath.ColumnName(c)));
        Assert.Equal(0, ExcelAddressMath.ColumnNumber("XFE"));
        Assert.Equal(0, ExcelAddressMath.ColumnNumber("A1"));
        Assert.Throws<ArgumentOutOfRangeException>(() => ExcelAddressMath.ColumnName(0));
        Assert.Throws<ArgumentOutOfRangeException>(() => ExcelAddressMath.ColumnName(16_385));
    }

    [Fact]
    public void Runs_group_consecutive_numbers()
    {
        Assert.Equal(new[] { (3, 5), (9, 9), (11, 12) }, ExcelAddressMath.Runs(new[] { 12, 3, 4, 11, 5, 9, 9 }));
        Assert.Empty(ExcelAddressMath.Runs(Array.Empty<int>()));
    }

    // ---- restore bookkeeping ----------------------------------------------------------------------------------

    private static CellLocation Loc(string sheet, string address, string book = "B.xlsx") => new(book, sheet, address);

    [Fact]
    public void Ledger_without_keep_undoes_everything_newest_first()
    {
        var ledger = new RestoreLedger();
        ledger.AddSheetVisibility("B.xlsx", "Hidden1", 0);
        ledger.AddRows("B.xlsx", "Hidden1", new[] { 19 });
        ledger.AddColumns("B.xlsx", "S2", new[] { 4 });

        var undo = ledger.SelectUndo(Array.Empty<CellLocation>());

        Assert.Equal(new[] { RestoreKind.Columns, RestoreKind.Rows, RestoreKind.SheetVisibility }, undo.Select(r => r.Kind));
    }

    [Fact]
    public void Row_record_is_kept_while_a_kept_range_touches_one_of_its_rows()
    {
        var ledger = new RestoreLedger();
        ledger.AddRows("B.xlsx", "S", new[] { 19 });

        Assert.Empty(ledger.SelectUndo(new[] { Loc("S", "$I$19:$J$19") }));
        Assert.Empty(ledger.SelectUndo(new[] { Loc("S", "$A$1:$A$30") }));
        Assert.Empty(ledger.SelectUndo(new[] { Loc("S", "$A:$A") }));        // whole column spans every row
        Assert.Empty(ledger.SelectUndo(new[] { Loc("S", "$A$1,$B$19") }));   // second area
        Assert.Single(ledger.SelectUndo(new[] { Loc("S", "$I$20") }));
        Assert.Single(ledger.SelectUndo(new[] { Loc("OtherSheet", "$I$19") }));
        Assert.Single(ledger.SelectUndo(new[] { Loc("S", "$I$19", "Other.xlsx") }));
    }

    [Fact]
    public void Column_record_uses_column_numbers()
    {
        var ledger = new RestoreLedger();
        ledger.AddColumns("B.xlsx", "S", new[] { 4 });

        Assert.Empty(ledger.SelectUndo(new[] { Loc("S", "$D$1") }));
        Assert.Empty(ledger.SelectUndo(new[] { Loc("S", "$C$1:$E$1") }));
        Assert.Empty(ledger.SelectUndo(new[] { Loc("S", "$5:$5") }));        // whole row spans every column
        Assert.Single(ledger.SelectUndo(new[] { Loc("S", "$E$1") }));
    }

    [Fact]
    public void Sheet_record_is_needed_by_any_kept_location_on_that_sheet()
    {
        var ledger = new RestoreLedger();
        ledger.AddSheetVisibility("B.xlsx", "Hidden1", 2);

        Assert.Empty(ledger.SelectUndo(new[] { Loc("Hidden1", "$Z$99") }));
        var undo = Assert.Single(ledger.SelectUndo(new[] { Loc("Main", "$A$1") }));
        Assert.Equal(2, undo.OriginalVisibility);
    }

    [Fact]
    public void Sheet_stays_visible_while_one_of_its_rows_is_kept_but_unneeded_rows_are_undone()
    {
        var ledger = new RestoreLedger();
        ledger.AddSheetVisibility("B.xlsx", "S", 0);
        ledger.AddRows("B.xlsx", "S", new[] { 8 });
        ledger.AddRows("B.xlsx", "S", new[] { 19 });

        var undo = ledger.SelectUndo(new[] { Loc("S", "$A$19") });

        var only = Assert.Single(undo);
        Assert.Equal(new[] { 8 }, only.Indexes);
    }

    [Fact]
    public void Several_kept_locations_each_keep_what_they_need()
    {
        var ledger = new RestoreLedger();
        ledger.AddRows("B.xlsx", "S", new[] { 12 });
        ledger.AddRows("B.xlsx", "S", new[] { 14 });
        ledger.AddRows("B.xlsx", "S", new[] { 16 });

        var undo = ledger.SelectUndo(new[] { Loc("S", "$A$12"), Loc("S", "$A$16") });

        Assert.Equal(new[] { 14 }, Assert.Single(undo).Indexes);
    }

    [Fact]
    public void Unparseable_keep_address_conservatively_keeps_the_whole_sheet()
    {
        var ledger = new RestoreLedger();
        ledger.AddRows("B.xlsx", "S", new[] { 12 });

        Assert.Empty(ledger.SelectUndo(new[] { Loc("S", "garbage") }));
    }

    [Fact]
    public void Names_are_compared_ignoring_case_like_Excel_does()
    {
        var ledger = new RestoreLedger();
        ledger.AddRows("Book1.xlsx", "Sheet", new[] { 3 });

        Assert.Empty(ledger.SelectUndo(new[] { Loc("SHEET", "$A$3", "BOOK1.XLSX") }));
    }

    [Fact]
    public void Remove_and_Clear_forget_records_and_empty_row_lists_are_ignored()
    {
        var ledger = new RestoreLedger();
        ledger.AddRows("B.xlsx", "S", Array.Empty<int>());
        Assert.Equal(0, ledger.Count);

        ledger.AddRows("B.xlsx", "S", new[] { 5, 5, 3 });
        ledger.AddColumns("B.xlsx", "S", new[] { 2 });
        Assert.Equal(new[] { 3, 5 }, ledger.Records[0].Indexes);

        ledger.Remove(ledger.Records[0]);
        Assert.Equal(1, ledger.Count);
        ledger.Clear();
        Assert.Equal(0, ledger.Count);
    }
}
