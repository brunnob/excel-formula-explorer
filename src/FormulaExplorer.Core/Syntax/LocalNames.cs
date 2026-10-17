namespace FormulaExplorer.Core.Syntax;

/// <summary>
/// Scope pass over a finished tree: marks the <see cref="NameNode"/>s that are variables bound by LET or LAMBDA
/// (<see cref="NameNode.IsLocal"/>), so a defined name that happens to share the spelling is not applied to them.
/// </summary>
internal static class LocalNames
{
    public static void Mark(FormulaNode root)
    {
        foreach (var node in root.DescendantsAndSelf())
        {
            if (node is not FunctionCallNode { Callee: null } call) continue;
            string function = Normalize(call.Name, "_xlfn.");
            bool isLet = function.Equals("LET", StringComparison.OrdinalIgnoreCase);
            if (!isLet && !function.Equals("LAMBDA", StringComparison.OrdinalIgnoreCase)) continue;

            var args = call.Arguments;
            for (int i = 0; i < args.Count - 1; i++) // the last argument is the result expression / lambda body
            {
                if (isLet && i % 2 != 0) continue;    // LET(name1, value1, name2, value2, ..., result)
                if (args[i] is not NameNode { Workbook: null, Sheet: null } declaration) continue;

                string variable = Normalize(declaration.Name, "_xlpm.");
                declaration.IsLocal = true;

                // LAMBDA: the parameter is visible in every argument. LET: from the next declaration on
                // (its own value expression still sees the outer meaning of the name).
                for (int j = isLet ? i + 2 : i + 1; j < args.Count; j++)
                    foreach (var inner in args[j].DescendantsAndSelf())
                        if (inner is NameNode { Workbook: null, Sheet: null } use
                            && Normalize(use.Name, "_xlpm.").Equals(variable, StringComparison.OrdinalIgnoreCase))
                            use.IsLocal = true;
            }
        }
    }

    private static string Normalize(string name, string prefix) =>
        name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) ? name.Substring(prefix.Length) : name;
}
