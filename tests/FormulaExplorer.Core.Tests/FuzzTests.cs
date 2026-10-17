using System.Text;
using FormulaExplorer.Core.Analysis;
using FormulaExplorer.Core.Syntax;
using FormulaExplorer.Core.Values;

namespace FormulaExplorer.Core.Tests;

/// <summary>Random and mutated formulas: nothing may throw, every character stays covered, every span stays consistent.</summary>
public class FuzzTests
{
    private static readonly string[] Fragments =
    {
        "=", "=", "A1", "$B$2", "C", "1", "12.5", ".5", "1E+20", "\"txt\"", "\"a\"\"b\"", "\"", "TRUE", "false", "#N/A", "#REF!", "#",
        "Sheet1!A1", "'My Sheet'!B2", "'It''s'!C3", "'", "Sheet1:Sheet3!A1", "[Book.xlsx]S!A1", "A1:B2", "A:A", "1:1", "Hundred",
        "Table1[Col]", "[@Col]", "Table1[[#Headers],[Col]]", "[", "]", "SUM(", "IF(", "OFFSET(", "_xlfn.F(", "(", ")", "(", ")", "{", "}",
        ",", ",", ";", ":", "+", "-", "*", "/", "^", "&", "=", "<", ">", "<=", ">=", "<>", "%", "@", "!", "$", "?", "~", " ", " ", "  ",
        "\n", "\t", "\r\n", "\u00A0", "é", "\uD83D\uDE00", "\\", "_", ".",
    };

    private static readonly string[] Seeds =
    {
        "=1/(((($I27/Hundred)+1)^(K$31)))",
        "=IF(K3=\"On\",1/(((-($I26/Hundred)+1)^(K$31))),0)",
        "=Sheet8!E7+(E30/C4)^(D4/D5)+D14-E21*E16+Summary!C5",
        "=SUM(I$19:$I$19)+J$19*$I$28/Hundred-SUM(OFFSET($I$19,0,0,1,Aqui_Per))",
        "=IF(C3,0,C11*'Inventory Forcast'!C$6)",
        "=(SUM(Fin_P_FO!L39:L40)+Fin_P_FO!L26+Notes_FO!L113+Notes_FO!L101)*(Macro_FA!L183/Hundred)",
        "=SUM(Table1[[#Headers],[A]:[B]],{1,2;3,4})&\"x\"",
        "=INDEX((A1:B2,C1:D2),1)+A1:B2 B2:C3-@A1#",
        "=-A1%^2<>--B1",
        "=IF(A1,,2)+F(,)",
    };

    // ---- the checks -------------------------------------------------------------------------------------

    private static void CheckInvariants(string text)
    {
        try
        {
            var tokens = FormulaTokenizer.Tokenize(text);

            // tokens: contiguous, non-empty, cover the whole text
            var rebuilt = new StringBuilder();
            int expectedStart = 0;
            foreach (var token in tokens)
            {
                Assert.Equal(expectedStart, token.Start);
                Assert.True(token.Length > 0, "empty token");
                Assert.Equal(text.Substring(token.Start, token.Length), token.Text);
                rebuilt.Append(token.Text);
                expectedStart = token.End;
            }
            Assert.Equal(text, rebuilt.ToString());

            var parsed = FormulaParser.Parse(text);
            if (!parsed.IsFormula)
            {
                Assert.Null(parsed.Root);
                return;
            }

            // tree: root spans everything, children inside parents and in order
            var root = parsed.Root!;
            Assert.Equal(0, root.Start);
            Assert.Equal(text.Length, root.Length);
            foreach (var node in parsed.Nodes)
            {
                Assert.True(node.Start >= 0 && node.Length >= 0 && node.End <= text.Length, $"span of {node}");
                int previousEnd = node.Start;
                foreach (var child in node.Children)
                {
                    Assert.Same(node, child.Parent);
                    Assert.True(child.Start >= previousEnd, $"child {child} overlaps previous sibling or starts before parent {node}");
                    Assert.True(child.End <= node.End, $"child {child} ends after parent {node}");
                    previousEnd = child.End;
                }
                if (node is not FormulaNode) Assert.NotNull(node.Parent);
            }

            // every reference-like token is exactly one leaf node, in order
            var tokenLeaves = tokens
                .Where(t => t.Kind is TokenKind.Reference or TokenKind.Name or TokenKind.StructuredReference)
                .Select(t => (t.Start, t.Length)).ToArray();
            var nodeLeaves = parsed.Nodes
                .Where(n => n.Kind is NodeKind.Reference or NodeKind.Name or NodeKind.StructuredReference)
                .OrderBy(n => n.Start).Select(n => (n.Start, n.Length)).ToArray();
            Assert.Equal(tokenLeaves, nodeLeaves);

            // hit map: stable, inside the node span
            var map = HitMap.Build(parsed);
            for (int i = 0; i < text.Length; i++)
            {
                var hit = map.NodeAt(i);
                if (hit is null) continue;
                Assert.True(hit.Start <= i && i < hit.End, $"hit {hit} does not contain index {i}");
                Assert.Contains(hit, parsed.Nodes);
                HitMap.IsLinkCandidate(hit);
            }
            Assert.Null(map.NodeAt(-1));
            Assert.Null(map.NodeAt(text.Length));

            // colors: only on leaves, valid palette indexes
            var colors = RefColorizer.Assign(parsed, "Sheet1", _ => true);
            Assert.True(colors.Nodes.Count == nodeLeaves.Length);
            foreach (var colored in colors.Nodes)
            {
                Assert.InRange(colored.PaletteIndex, 0, RefColorizer.Palette.Count - 1);
                Assert.Equal(RefColorizer.Palette[colored.PaletteIndex], colored.Rgb);
            }
        }
        catch (Exception ex) when (ex is not Xunit.Sdk.XunitException)
        {
            throw new InvalidOperationException($"Exception for formula: {Escape(text)}", ex);
        }
        catch (Xunit.Sdk.XunitException ex)
        {
            throw new InvalidOperationException($"Invariant violated for formula: {Escape(text)}\n{ex.Message}", ex);
        }
    }

    private static string Escape(string text) => text.Replace("\n", "\\n").Replace("\r", "\\r").Replace("\t", "\\t");

    // ---- the fuzzers ------------------------------------------------------------------------------------------

    [Fact]
    public void RandomFragmentSoup()
    {
        var random = new Random(20111016);
        for (int n = 0; n < 4000; n++)
        {
            var sb = new StringBuilder("=");
            int parts = random.Next(1, 25);
            for (int i = 0; i < parts; i++) sb.Append(Fragments[random.Next(Fragments.Length)]);
            CheckInvariants(sb.ToString());
        }
    }

    [Fact]
    public void RandomCharacters()
    {
        var random = new Random(42);
        const string alphabet = "=()[]{}\",;:+-*/^&<>%@#!$?~ '\\_.\t\nAaBbZz09EeXx";
        for (int n = 0; n < 3000; n++)
        {
            var sb = new StringBuilder(random.Next(3) == 0 ? "" : "=");
            int length = random.Next(0, 40);
            for (int i = 0; i < length; i++) sb.Append(alphabet[random.Next(alphabet.Length)]);
            CheckInvariants(sb.ToString());
        }
    }

    [Fact]
    public void MutatedRealisticFormulas()
    {
        var random = new Random(7);
        for (int n = 0; n < 4000; n++)
        {
            var sb = new StringBuilder(Seeds[random.Next(Seeds.Length)]);
            int mutations = random.Next(1, 6);
            for (int m = 0; m < mutations && sb.Length > 0; m++)
            {
                int position = random.Next(sb.Length);
                switch (random.Next(4))
                {
                    case 0: // delete
                        sb.Remove(position, Math.Min(random.Next(1, 4), sb.Length - position));
                        break;
                    case 1: // insert a fragment
                        sb.Insert(position, Fragments[random.Next(Fragments.Length)]);
                        break;
                    case 2: // replace a character
                        sb[position] = Fragments[random.Next(Fragments.Length)][0];
                        break;
                    default: // truncate
                        sb.Length = position;
                        break;
                }
            }
            CheckInvariants(sb.ToString());
        }
    }

    [Fact]
    public void EveryPrefixOfRealisticFormulas()
    {
        foreach (var seed in Seeds)
            for (int length = 0; length <= seed.Length; length++)
                CheckInvariants(seed.Substring(0, length));
    }

    [Fact]
    public void EverySuffixAndDeletionOfRealisticFormulas()
    {
        foreach (var seed in Seeds)
            for (int i = 0; i < seed.Length; i++)
            {
                CheckInvariants("=" + seed.Substring(i));
                CheckInvariants(seed.Remove(i, 1));
            }
    }

    [Fact]
    public void CompleteFormulasFromSeedsAreValid()
    {
        foreach (var seed in new[] { Seeds[0], Seeds[1], Seeds[2], Seeds[3], Seeds[4], Seeds[5] })
            Assert.False(FormulaParser.Parse(seed).HasErrors, seed);
    }

    [Fact]
    public void ExtremeInputsDoNotThrow()
    {
        CheckInvariants("=" + new string('(', 3000));
        CheckInvariants("=" + new string(')', 3000));
        CheckInvariants("=" + new string('{', 3000));
        CheckInvariants("=" + new string('-', 3000) + "1");
        CheckInvariants("=" + new string('"', 3001));
        CheckInvariants("=" + new string('[', 3000));
        CheckInvariants("=" + string.Concat(Enumerable.Repeat("F(", 3000)));
        CheckInvariants("=" + string.Concat(Enumerable.Repeat("A1 ", 3000)));
        CheckInvariants("=" + string.Join(",", Enumerable.Repeat("A1", 3000)));
        CheckInvariants("=SUM(" + string.Join(",", Enumerable.Repeat("A1", 3000)) + ")");
        CheckInvariants("=" + string.Concat(Enumerable.Repeat("'", 3001)));
        CheckInvariants("='" + new string('a', 5000));
        CheckInvariants("=" + new string('9', 5000));
        CheckInvariants("=" + new string('A', 5000));
        CheckInvariants("=A1" + new string('#', 2000));
        CheckInvariants("=" + new string('\uD83D', 100));
    }

    [Fact]
    public void ValueFormatterNeverThrowsOnOddValues()
    {
        var values = new object?[] { null, DBNull.Value, double.NaN, double.PositiveInfinity, double.NegativeInfinity, double.Epsilon, 'c', "", new string('x', 1000), ExcelError.Busy, (ExcelError)(-1), 5, 5L, 5m, new DateTime(1900, 1, 1), new object() };
        foreach (var value in values)
        {
            Assert.NotNull(ValueFormatter.FormatScalar(value, true));
            Assert.NotNull(ValueFormatter.FormatScalar(value, false));
        }
    }
}
