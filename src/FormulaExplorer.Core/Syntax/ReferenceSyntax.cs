using System.Text;

namespace FormulaExplorer.Core.Syntax;

/// <summary>The sheet/workbook prefix of a reference or name (everything up to and including the '!').</summary>
/// <param name="End">Index just after the '!' (0 when there is no prefix).</param>
/// <param name="Workbook">Workbook name from [Book.xlsx] (without brackets), or null.</param>
/// <param name="Sheet">First sheet name, unquoted and unescaped, or null.</param>
/// <param name="LastSheet">Last sheet of a 3D prefix (Sheet1:Sheet3!), or null.</param>
internal readonly record struct SheetPrefix(int End, string? Workbook, string? Sheet, string? LastSheet);

/// <summary>
/// Lexical helpers for references (cell addresses, sheet prefixes, bracketed specifiers) shared by the
/// tokenizer, which finds the token boundaries, and the parser, which splits a token back into its parts.
/// </summary>
internal static class ReferenceSyntax
{
    internal const int MaxColumn = 16384; // XFD
    internal const int MaxRow = 1048576;

    private static readonly string[] WorkbookExtensions = { ".xlsx", ".xlsm", ".xlsb", ".xls", ".xlam", ".xla", ".xltx", ".xltm", ".xlt" };

    internal static bool IsIdentStart(char c) => char.IsLetter(c) || c == '_' || c == '\\';

    /// <summary>Characters allowed inside a defined name or function name.</summary>
    internal static bool IsIdentPart(char c) => char.IsLetterOrDigit(c) || c is '_' or '.' or '\\' or '?';

    /// <summary>Characters allowed in an unquoted sheet or workbook name.</summary>
    private static bool IsSheetPart(char c) => char.IsLetterOrDigit(c) || c is '_' or '.';

    internal static int ScanIdent(string s, int i)
    {
        while (i < s.Length && IsIdentPart(s[i])) i++;
        return i;
    }

    // ------------------------------------------------------------------------------------------------
    // Addresses
    // ------------------------------------------------------------------------------------------------

    /// <summary>
    /// True when a cell, area, column range or row range starts at <paramref name="i"/>. Fails when the address is
    /// glued to more identifier characters (ABCD1, TAX1.5, A1B) — those are names.
    /// </summary>
    internal static bool TryScanAddress(string s, int i, out int end)
    {
        end = i;
        int n = s.Length;
        if (i >= n) return false;
        char c = s[i];
        if (IsAsciiDigit(c) || (c == '$' && i + 1 < n && IsAsciiDigit(s[i + 1])))
            return TryScanRows(s, i, out end);

        if (TryScanCell(s, i, out int q))
        {
            if (Continues(s, q)) return false;
            end = q;
            if (q < n && s[q] == ':' && TryScanCell(s, q + 1, out int r) && !Continues(s, r)) end = r;
            return true;
        }

        return TryScanColumns(s, i, out end);
    }

    /// <summary>True if the character at <paramref name="i"/> would extend an identifier (so a candidate address is really a name).</summary>
    private static bool Continues(string s, int i) => i < s.Length && (IsIdentPart(s[i]) || s[i] == '(');

    private static bool IsAsciiDigit(char c) => c >= '0' && c <= '9';

    private static bool IsAsciiLetter(char c) => (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');

    /// <summary>$?COL (1-3 letters, at most XFD) then $?ROW (1..1048576).</summary>
    private static bool TryScanCell(string s, int i, out int end)
    {
        end = i;
        int n = s.Length, j = i;
        if (j < n && s[j] == '$') j++;
        if (!TryScanColumnLabel(s, j, out j)) return false;
        if (j < n && s[j] == '$') j++;
        if (!TryScanRowNumber(s, j, out j)) return false;
        end = j;
        return true;
    }

    private static bool TryScanColumnLabel(string s, int i, out int end)
    {
        end = i;
        int j = i, value = 0;
        while (j < s.Length && IsAsciiLetter(s[j]))
        {
            if (j - i == 3) return false; // a fourth letter: not a column
            value = value * 26 + (char.ToUpperInvariant(s[j]) - 'A' + 1);
            j++;
        }
        if (j == i || value > MaxColumn) return false;
        end = j;
        return true;
    }

    private static bool TryScanRowNumber(string s, int i, out int end)
    {
        end = i;
        int j = i;
        long value = 0;
        while (j < s.Length && IsAsciiDigit(s[j]))
        {
            if (value <= MaxRow) value = value * 10 + (s[j] - '0'); // saturates above MaxRow
            j++;
        }
        if (j == i || value < 1 || value > MaxRow) return false;
        end = j;
        return true;
    }

    /// <summary>$?COL:$?COL</summary>
    private static bool TryScanColumns(string s, int i, out int end)
    {
        end = i;
        int j = i;
        if (j < s.Length && s[j] == '$') j++;
        if (!TryScanColumnLabel(s, j, out j)) return false;
        if (j >= s.Length || s[j] != ':') return false;
        j++;
        if (j < s.Length && s[j] == '$') j++;
        if (!TryScanColumnLabel(s, j, out j) || Continues(s, j)) return false;
        end = j;
        return true;
    }

    /// <summary>$?ROW:$?ROW</summary>
    private static bool TryScanRows(string s, int i, out int end)
    {
        end = i;
        int j = i;
        if (j < s.Length && s[j] == '$') j++;
        if (!TryScanRowNumber(s, j, out j)) return false;
        if (j >= s.Length || s[j] != ':') return false;
        j++;
        if (j < s.Length && s[j] == '$') j++;
        if (!TryScanRowNumber(s, j, out j) || Continues(s, j)) return false;
        end = j;
        return true;
    }

    /// <summary>Classifies an address (the part after the prefix) of a Reference token.</summary>
    internal static ReferenceForm ClassifyAddress(string address)
    {
        if (address.Length > 0 && address[0] == '#') return ReferenceForm.RefError;
        int colon = address.IndexOf(':');
        if (colon < 0) return ReferenceForm.Cell;
        string left = address.Substring(0, colon).TrimStart('$');
        string right = address.Substring(colon + 1).TrimStart('$');
        if (left.Length > 0 && right.Length > 0)
        {
            if (left.All(IsAsciiLetter) && right.All(IsAsciiLetter)) return ReferenceForm.Columns;
            if (left.All(IsAsciiDigit) && right.All(IsAsciiDigit)) return ReferenceForm.Rows;
        }
        return ReferenceForm.Area;
    }

    // ------------------------------------------------------------------------------------------------
    // Brackets (structured references)
    // ------------------------------------------------------------------------------------------------

    /// <summary>
    /// Scans a bracketed specifier starting at the '[' at <paramref name="i"/> to its matching ']'. Brackets nest and
    /// the apostrophe escapes the next character (<c>'[</c>, <c>']</c>, <c>'#</c>, <c>''</c>). False when unterminated.
    /// </summary>
    internal static bool TryScanBrackets(string s, int i, out int end)
    {
        end = i;
        int depth = 0;
        for (int j = i; j < s.Length; j++)
        {
            char c = s[j];
            if (c == '\'') { j++; continue; }
            if (c == '[') depth++;
            else if (c == ']' && --depth == 0)
            {
                end = j + 1;
                return true;
            }
        }
        return false;
    }

    // ------------------------------------------------------------------------------------------------
    // Prefixes
    // ------------------------------------------------------------------------------------------------

    /// <summary>
    /// Scans a sheet/workbook prefix that starts at <paramref name="i"/>: <c>Sheet1!</c>, <c>'My Sheet'!</c>,
    /// <c>Sheet1:Sheet3!</c>, <c>'Sheet 1:Sheet 3'!</c>, <c>[Book.xlsx]Sheet1!</c>, <c>'C:\Dir\[Book.xlsx]Sheet1'!</c>,
    /// <c>[Book.xlsx]!</c>. The '!' is part of the prefix.
    /// </summary>
    internal static bool TryScanPrefix(string s, int i, out SheetPrefix prefix)
    {
        prefix = default;
        int n = s.Length;
        if (i >= n) return false;
        char c = s[i];

        if (c == '\'')
        {
            var sb = new StringBuilder();
            int j = i + 1;
            while (true)
            {
                if (j >= n) return false;
                if (s[j] == '\'')
                {
                    if (j + 1 < n && s[j + 1] == '\'') { sb.Append('\''); j += 2; continue; }
                    break;
                }
                sb.Append(s[j]);
                j++;
            }
            if (j + 1 >= n || s[j + 1] != '!') return false;
            prefix = FromQuotedContent(sb.ToString(), j + 2);
            return true;
        }

        string? workbook = null;
        int p = i;
        if (c == '[')
        {
            int close = s.IndexOf(']', i + 1);
            if (close < 0) return false;
            workbook = s.Substring(i + 1, close - i - 1);
            p = close + 1;
            if (p < n && s[p] == '!')
            {
                prefix = new SheetPrefix(p + 1, workbook, null, null);
                return true;
            }
        }
        else if (!IsIdentStart(c) || c == '\\')
        {
            return false;
        }

        int e = p;
        while (e < n && IsSheetPart(s[e])) e++;
        if (e == p) return false;
        string sheet = s.Substring(p, e - p);
        string? last = null;
        if (e < n && s[e] == ':')
        {
            int f = e + 1, g = f;
            while (g < n && IsSheetPart(s[g])) g++;
            if (g == f) return false;
            last = s.Substring(f, g - f);
            e = g;
        }
        if (e >= n || s[e] != '!') return false;

        if (workbook is null && last is null && HasWorkbookExtension(sheet))
            prefix = new SheetPrefix(e + 1, sheet, null, null); // Book1.xlsx!Name
        else
            prefix = new SheetPrefix(e + 1, workbook, sheet, last);
        return true;
    }

    private static SheetPrefix FromQuotedContent(string content, int end)
    {
        string? workbook = null;
        string sheetPart = content;
        int open = content.IndexOf('[');
        if (open >= 0)
        {
            int close = content.IndexOf(']', open + 1);
            if (close > open)
            {
                workbook = content.Substring(open + 1, close - open - 1);
                sheetPart = content.Substring(close + 1);
            }
        }

        string? sheet = sheetPart, last = null;
        int colon = sheetPart.IndexOf(':');
        if (colon >= 0)
        {
            sheet = sheetPart.Substring(0, colon);
            last = sheetPart.Substring(colon + 1);
        }
        if (sheet.Length == 0) sheet = null;
        if (string.IsNullOrEmpty(last)) last = null;

        if (workbook is null && last is null && sheet is not null && HasWorkbookExtension(sheet))
            return new SheetPrefix(end, sheet, null, null); // 'Book 1.xlsx'!Name
        return new SheetPrefix(end, workbook, sheet, last);
    }

    private static bool HasWorkbookExtension(string name)
    {
        foreach (var ext in WorkbookExtensions)
            if (name.Length > ext.Length && name.EndsWith(ext, StringComparison.OrdinalIgnoreCase))
                return true;
        return false;
    }

    /// <summary>The prefix at the start of a Reference/Name/StructuredReference token text; <c>End == 0</c> when absent.</summary>
    internal static SheetPrefix GetPrefix(string tokenText) =>
        TryScanPrefix(tokenText, 0, out var prefix) ? prefix : default;
}
