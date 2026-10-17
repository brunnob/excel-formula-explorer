namespace FormulaExplorer.Core.Values;

/// <summary>Excel error values (the CVErr numbers). COM returns them as Int32 = unchecked((int)0x800A0000) | code.</summary>
public enum ExcelError
{
    Null = 2000,
    Div0 = 2007,
    Value = 2015,
    Ref = 2023,
    Name = 2029,
    Num = 2036,
    NA = 2042,
    GettingData = 2043,
    Spill = 2045,
    Connect = 2046,
    Blocked = 2047,
    Unknown = 2048,
    Field = 2049,
    Calc = 2050,
    Busy = 2051,
}

/// <summary>
/// A COM-free model of what Excel returned for a value or an Evaluate call.
/// Scalar cell values inside <see cref="ArrayValue"/>/<see cref="RangeValue"/> are: double, string, bool,
/// <see cref="ExcelError"/>, or null (blank cell / empty).
/// </summary>
public abstract record EvalResult
{
    private EvalResult() { }

    public sealed record Number(double Value) : EvalResult;

    public sealed record Text(string Value) : EvalResult;

    public sealed record Bool(bool Value) : EvalResult;

    public sealed record Error(ExcelError Code) : EvalResult;

    /// <summary>An empty cell / empty value.</summary>
    public sealed record Blank : EvalResult;

    /// <summary>
    /// An in-memory array result (e.g. Evaluate("A1:A3*2")). <paramref name="Rows"/>/<paramref name="Columns"/> are the
    /// real dimensions; <paramref name="Values"/> (0-based [row, col]) may hold only the first cells (row-major) when the
    /// array is large — the gateway copies at most ValueFormatter.MaxListedValues + 1 cells.
    /// </summary>
    public sealed record ArrayValue(int Rows, int Columns, object?[,] Values) : EvalResult;

    /// <summary>
    /// A Range result. <paramref name="Address"/> is the relative A1 address without '$' and without sheet (e.g. "I19:J19");
    /// multi-area ranges use a comma list ("A1:B2,D4"). <paramref name="Rows"/>/<paramref name="Columns"/> describe the first area.
    /// <paramref name="Values"/> holds at most the first N cells of the first area (row-major), 0-based [row, col];
    /// it may be smaller than Rows x Columns when truncated.
    /// </summary>
    public sealed record RangeValue(string Address, int Rows, int Columns, object?[,] Values) : EvalResult;

    /// <summary>Could not evaluate (fragment longer than 255 chars, Excel busy, exception...).</summary>
    public sealed record Unavailable(string Reason) : EvalResult;

    public static Blank BlankValue { get; } = new();
}
