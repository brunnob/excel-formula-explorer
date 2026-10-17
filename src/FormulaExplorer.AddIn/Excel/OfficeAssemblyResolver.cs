using System.IO;
using System.Reflection;
using FormulaExplorer.AddIn.Diagnostics;

namespace FormulaExplorer.AddIn.Excel;

/// <summary>
/// For hosts that reference the Excel primary interop assembly WITHOUT embedding its types (tests/integration/GatewayHarness).
/// That assembly references the Office PIA ("office", 16.0.0.0), which the NuGet package does not ship and which .NET
/// (Core) cannot find in the old GAC by itself: loading any Excel interop type then fails with FileNotFoundException("office").
/// This resolver answers that request with the copy installed by Office (legacy GAC, or the add-in folders of a
/// Click-to-Run installation). A .NET Framework build of that assembly loads fine in .NET: it only holds interfaces/enums we
/// never touch. Install it BEFORE the first method that mentions an Excel interop type is JIT-compiled.
/// <para>The add-in itself does NOT use it: its interop types are embedded (see the EmbedExcelInteropTypes target in the csproj),
/// so FormulaExplorer.AddIn.dll references neither the PIA nor "office", and a process-wide AssemblyResolve hook inside
/// Excel would only interfere with other add-ins.</para>
/// </summary>
public static class OfficeAssemblyResolver
{
    private const string OfficeAssemblyName = "office";
    private static int _installed;

    public static void Install()
    {
        if (Interlocked.Exchange(ref _installed, 1) == 1) return;
        AppDomain.CurrentDomain.AssemblyResolve += Resolve;
    }

    private static Assembly? Resolve(object? sender, ResolveEventArgs args)
    {
        try
        {
            if (!string.Equals(new AssemblyName(args.Name).Name, OfficeAssemblyName, StringComparison.OrdinalIgnoreCase))
                return null;

            var path = FindOfficeAssembly();
            if (path is null)
            {
                Log.Info("OfficeAssemblyResolver: no 'office' assembly found on this machine.");
                return null;
            }
            Log.Info($"OfficeAssemblyResolver: resolved 'office' to {path}");
            return Assembly.LoadFrom(path);
        }
        catch (Exception ex)
        {
            Log.Error("OfficeAssemblyResolver.Resolve", ex);
            return null;
        }
    }

    /// <summary>Newest installed copy of office.dll (legacy GAC first, then Click-to-Run add-in folders).</summary>
    private static string? FindOfficeAssembly()
    {
        var candidates = new List<string>();

        var gac = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "assembly", "GAC_MSIL", OfficeAssemblyName);
        if (Directory.Exists(gac))
        {
            foreach (var versionDir in Directory.EnumerateDirectories(gac))
                candidates.AddRange(Directory.EnumerateFiles(versionDir, "office.dll"));
        }

        foreach (var programFiles in new[] { Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86) })
        {
            foreach (var sub in new[] { @"Microsoft Office\root\Office16\ADDINS", @"Microsoft Office\root\vfs\ProgramFilesX86\Microsoft Office\Office16\DCF" })
            {
                var dir = Path.Combine(programFiles, sub);
                if (Directory.Exists(dir))
                    candidates.AddRange(Directory.EnumerateFiles(dir, "office.dll", SearchOption.AllDirectories));
            }
        }

        return candidates
            .Select(p => (Path: p, Version: TryGetVersion(p)))
            .Where(c => c.Version is not null)
            .OrderByDescending(c => c.Version)
            .Select(c => c.Path)
            .FirstOrDefault();
    }

    private static Version? TryGetVersion(string path)
    {
        try { return AssemblyName.GetAssemblyName(path).Version; }
        catch (Exception ex) when (ex is BadImageFormatException or FileLoadException or IOException)
        {
            Log.Info($"OfficeAssemblyResolver: skipped {path}: {ex.Message}");
            return null;
        }
    }
}
