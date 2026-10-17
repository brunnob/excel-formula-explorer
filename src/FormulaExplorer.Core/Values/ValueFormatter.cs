using System.Globalization;
using System.Text;

namespace FormulaExplorer.Core.Values;

/// <summary>
/// Formats evaluation results for the "Returns" tooltip. Always invariant culture.
/// Formats follow the original add-in (see docs/PESQUISA-FORMULA-EXPLORER.md section 3.4).
/// </summary>
public static class ValueFormatter
{
    /// <summary>Maximum number of cell values listed for ranges/arrays before "...".</summary>
    public const int MaxListedValues = 20;

    /// <summary>Maximum characters of a string value before it is cut with "...".</summary>
    public const int MaxStringLength = 255;

    /// <summary>Text shown in the tooltip for <paramref name="result"/>.</summary>
    public static string Format(EvalResult result)
    {
        switch (result)
        {
            case EvalResult.Number n:
                return FormatNumber(n.Value);
            case EvalResult.Text t:
                return FormatText(t.Value);
            case EvalResult.Bool b:
                return b.Value ? "TRUE" : "FALSE";
            case EvalResult.Error e:
                return ErrorText(e.Code);
            case EvalResult.Blank:
                return "[Blank]";
            case EvalResult.ArrayValue a:
                return "[" + Dimensions(a.Rows, a.Columns) + "] " + ListValues(a.Values, (long)a.Rows * a.Columns);
            case EvalResult.RangeValue r:
                return FormatRange(r);
            case EvalResult.Unavailable:
                return "[Unable to evaluate]";
            default:
                return "[Unable to evaluate]";
        }
    }

    /// <summary>
    /// Formats a single scalar (double, string, bool, <see cref="ExcelError"/>, or null).
    /// When <paramref name="inList"/> is true a blank is "" instead of "[Blank]".
    /// </summary>
    public static string FormatScalar(object? value, bool inList)
    {
        switch (value)
        {
            case null:
            case DBNull:
                return inList ? string.Empty : "[Blank]";
            case double d:
                return FormatNumber(d);
            case string s:
                return FormatText(s);
            case bool b:
                return b ? "TRUE" : "FALSE";
            case ExcelError e:
                return ErrorText(e);
            case DateTime dt:
                return FormatNumber(dt.ToOADate());
            case IConvertible c when value is not char:
                // int, float, decimal, ... (the gateway normally converts to double already)
                return FormatNumber(c.ToDouble(CultureInfo.InvariantCulture));
            default:
                return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        }
    }

    /// <summary>"#DIV/0!", "#N/A", ... for an error code.</summary>
    public static string ErrorText(ExcelError error) => error switch
    {
        ExcelError.Null => "#NULL!",
        ExcelError.Div0 => "#DIV/0!",
        ExcelError.Value => "#VALUE!",
        ExcelError.Ref => "#REF!",
        ExcelError.Name => "#NAME?",
        ExcelError.Num => "#NUM!",
        ExcelError.NA => "#N/A",
        ExcelError.GettingData => "#GETTING_DATA",
        ExcelError.Spill => "#SPILL!",
        ExcelError.Connect => "#CONNECT!",
        ExcelError.Blocked => "#BLOCKED!",
        ExcelError.Unknown => "#UNKNOWN!",
        ExcelError.Field => "#FIELD!",
        ExcelError.Calc => "#CALC!",
        ExcelError.Busy => "#BUSY!",
        _ => "#VALUE!",
    };

    /// <summary>Up to 15 significant digits like Excel shows them; negative zero is "0".</summary>
    private static string FormatNumber(double value)
    {
        if (value == 0) return "0"; // also -0.0
        return value.ToString("G15", CultureInfo.InvariantCulture);
    }

    private static string FormatText(string value)
    {
        var sb = new StringBuilder(Math.Min(value.Length, MaxStringLength) + 5);
        sb.Append('"');
        if (value.Length > MaxStringLength)
            sb.Append(value, 0, MaxStringLength).Append("...");
        else
            sb.Append(value);
        sb.Append('"');
        return sb.ToString();
    }

    private static string Dimensions(int rows, int columns) =>
        rows.ToString(CultureInfo.InvariantCulture) + "x" + columns.ToString(CultureInfo.InvariantCulture);

    private static string FormatRange(EvalResult.RangeValue range)
    {
        if (range.Rows == 1 && range.Columns == 1 && !range.Address.Contains(','))
            return FormatScalar(range.Values.Length > 0 ? range.Values[0, 0] : null, inList: false);

        // The gateway fills Rows/Columns/Values from the first area only; a multi-area range always has more
        // cells than that, so it always ends with ", ...".
        long totalCells = range.Address.Contains(',') ? long.MaxValue : (long)range.Rows * range.Columns;
        return range.Address + " containing\n[" + Dimensions(range.Rows, range.Columns) + "] "
            + ListValues(range.Values, totalCells);
    }

    /// <summary>
    /// The first <see cref="MaxListedValues"/> values in row-major order of <paramref name="values"/> (which may hold
    /// fewer cells than <paramref name="totalCells"/>), blanks as empty strings, joined by ", ", plus ", ..." if more exist.
    /// </summary>
    private static string ListValues(object?[,] values, long totalCells)
    {
        int rows = values.GetLength(0), columns = values.GetLength(1);
        long available = (long)rows * columns;
        int listed = (int)Math.Min(MaxListedValues, available);

        var sb = new StringBuilder();
        for (int k = 0; k < listed; k++)
        {
            if (k > 0) sb.Append(", ");
            sb.Append(FormatScalar(values[k / columns, k % columns], inList: true));
        }
        if (totalCells > listed) sb.Append(", ...");
        return sb.ToString();
    }
}
