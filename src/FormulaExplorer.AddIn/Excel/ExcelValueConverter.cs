using System.Reflection;
using FormulaExplorer.Core.Values;

namespace FormulaExplorer.AddIn.Excel;

/// <summary>
/// Pure conversion of raw COM values (Range.Value2 elements, Worksheet.Evaluate results) to the COM-free
/// <see cref="EvalResult"/> model. No Excel access, so it is unit-testable; exposed through
/// <see cref="ExcelGateway.ConvertValue"/>.
/// </summary>
public static class ExcelValueConverter
{
    /// <summary>High word of an Excel error returned as Int32: <c>0x800A0000 | CVErr code</c>.</summary>
    private const uint ErrorHighWord = 0x800A0000;

    public static EvalResult Convert(object? comValue)
    {
        if (comValue is Array array) return ConvertArray(array);
        if (!TryConvertScalar(comValue, out var scalar))
            return new EvalResult.Unavailable($"Unsupported value type {comValue!.GetType().Name}");

        return scalar switch
        {
            null => EvalResult.BlankValue,
            double d => new EvalResult.Number(d),
            string s => new EvalResult.Text(s),
            bool b => new EvalResult.Bool(b),
            ExcelError e => new EvalResult.Error(e),
            _ => throw new InvalidOperationException("TryConvertScalar produced an unexpected type."),
        };
    }

    /// <summary>
    /// Converts one cell value to the scalar vocabulary of <see cref="EvalResult.ArrayValue"/>:
    /// double, string, bool, <see cref="ExcelError"/> or null (blank). False for types that are not cell values.
    /// </summary>
    public static bool TryConvertScalar(object? value, out object? scalar)
    {
        switch (value)
        {
            case null:
            case DBNull:
            case Missing:
                scalar = null;
                return true;
            case double or string or bool:
                scalar = value;
                return true;
            case int i when (unchecked((uint)i) & 0xFFFF0000) == ErrorHighWord:
                var code = i & 0xFFFF;
                scalar = Enum.IsDefined((ExcelError)code) ? (ExcelError)code : ExcelError.Unknown;
                return true;
            case int i:
                scalar = (double)i;
                return true;
            case short s:
                scalar = (double)s;
                return true;
            case long l:
                scalar = (double)l;
                return true;
            case float f:
                scalar = (double)f;
                return true;
            case byte or sbyte or ushort or uint or ulong:
                scalar = System.Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture);
                return true;
            case decimal m: // currency-formatted results
                scalar = (double)m;
                return true;
            case DateTime dt: // date-formatted results
                scalar = dt.ToOADate();
                return true;
            default:
                scalar = null;
                return false;
        }
    }

    /// <summary>
    /// The size of the leading block of a rows x columns grid that contains its first <paramref name="cap"/> cells in
    /// row-major order and keeps the real column count whenever possible, so the block can be indexed [i / columns, i % columns]:
    /// whole rows when the grid has at most <paramref name="cap"/> columns, otherwise a single partial row.
    /// The block can hold up to <c>columns - 1</c> cells more than <paramref name="cap"/> (the rest of its last row).
    /// </summary>
    public static (int Rows, int Columns) LeadingBlock(int rows, int columns, int cap)
    {
        if (rows <= 0 || columns <= 0 || cap <= 0) return (0, 0);
        var cells = (long)rows * columns;
        var wanted = (int)Math.Min(cells, cap);
        if (columns > wanted) return (1, wanted);
        return (Math.Min(rows, (wanted + columns - 1) / columns), columns);
    }

    /// <summary>
    /// Converts an Excel array (always 1-based when it comes from COM; a one-dimensional array is a single row) to
    /// <see cref="EvalResult.ArrayValue"/> keeping the real dimensions but copying only the leading block of at most
    /// <see cref="ValueFormatter.MaxListedValues"/> + 1 cells (plus the remainder of its last row), 0-based.
    /// </summary>
    private static EvalResult ConvertArray(Array array)
    {
        int rows, columns;
        switch (array.Rank)
        {
            case 1:
                (rows, columns) = (1, array.GetLength(0));
                break;
            case 2:
                (rows, columns) = (array.GetLength(0), array.GetLength(1));
                break;
            default:
                return new EvalResult.Unavailable($"Unsupported array rank {array.Rank}");
        }
        if (rows == 0 || columns == 0) return new EvalResult.ArrayValue(rows, columns, new object?[0, 0]);

        var (copyRows, copyColumns) = LeadingBlock(rows, columns, ValueFormatter.MaxListedValues + 1);
        var values = new object?[copyRows, copyColumns];
        var rowBase = array.GetLowerBound(0);
        var columnBase = array.Rank == 2 ? array.GetLowerBound(1) : 0;
        for (var r = 0; r < copyRows; r++)
        {
            for (var c = 0; c < copyColumns; c++)
            {
                var raw = array.Rank == 1 ? array.GetValue(rowBase + c) : array.GetValue(rowBase + r, columnBase + c);
                values[r, c] = TryConvertScalar(raw, out var scalar) ? scalar : raw!.ToString();
            }
        }
        return new EvalResult.ArrayValue(rows, columns, values);
    }

    /// <summary>
    /// A <c>Range.Value2</c> read of a rectangular block (scalar for one cell, 1-based object[,] otherwise) as a
    /// 0-based array of scalars (see <see cref="TryConvertScalar"/>).
    /// </summary>
    public static object?[,] ToZeroBasedBlock(object? value2, int rows, int columns)
    {
        var block = new object?[rows, columns];
        if (value2 is Array array && array.Rank == 2)
        {
            var rowBase = array.GetLowerBound(0);
            var columnBase = array.GetLowerBound(1);
            for (var r = 0; r < Math.Min(rows, array.GetLength(0)); r++)
                for (var c = 0; c < Math.Min(columns, array.GetLength(1)); c++)
                {
                    var raw = array.GetValue(rowBase + r, columnBase + c);
                    block[r, c] = TryConvertScalar(raw, out var scalar) ? scalar : raw!.ToString();
                }
        }
        else if (rows > 0 && columns > 0)
        {
            block[0, 0] = TryConvertScalar(value2, out var scalar) ? scalar : value2!.ToString();
        }
        return block;
    }
}
