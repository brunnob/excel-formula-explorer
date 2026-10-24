using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace FormulaExplorer.Setup;

/// <summary>
/// Per-user installer for the Formula Explorer add-in (no admin rights): copies the embedded packed .xll to
/// %LOCALAPPDATA%\FormulaExplorer, registers it to open with Excel (HKCU ...\Excel\Options OPEN, OPEN1, ...) and adds an
/// entry to "Installed apps". Usage: FormulaExplorerSetup.exe [/uninstall] [/quiet]
/// </summary>
internal static class Program
{
    private const string Title = "Formula Explorer Setup";
    private const string DisplayName = "Formula Explorer for Excel";
    private const string Version = "1.1.0";
    private const string ExcelOptionsKey = @"Software\Microsoft\Office\16.0\Excel\Options";
    private const string UninstallKey = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\FormulaExplorer";
    private const string RuntimeDownload = "https://aka.ms/dotnet/10.0/windowsdesktop-runtime-win-x64.exe";
    private const string RuntimePage = "https://dotnet.microsoft.com/download/dotnet/10.0";

    private static bool _quiet;

    private static string InstallDir =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "FormulaExplorer");

    private static string XllPath => Path.Combine(InstallDir, "FormulaExplorer64.xll");

    private static string UninstallerPath => Path.Combine(InstallDir, "FormulaExplorerSetup.exe");

    [STAThread]
    private static int Main(string[] args)
    {
        _quiet = args.Any(a => a.Equals("/quiet", StringComparison.OrdinalIgnoreCase));
        bool uninstall = args.Any(a => a.Equals("/uninstall", StringComparison.OrdinalIgnoreCase));
        try
        {
            return uninstall ? Uninstall() : Install();
        }
        catch (Exception ex)
        {
            Show($"Setup failed:\n\n{ex.Message}", MB_ICONERROR);
            return 1;
        }
    }

    // ------------------------------------------------------------------------------------------------------------

    private static int Install()
    {
        if (!Confirm($"Install {DisplayName} {Version}?\n\nIt adds the \"Studio\" tab to Excel's ribbon for the current user. " +
                     "No administrator rights are needed."))
            return 2;

        if (ExcelPlatform() is { } platform && !platform.Equals("x64", StringComparison.OrdinalIgnoreCase))
        {
            Show($"Formula Explorer needs 64-bit Excel, but this Excel is '{platform}'.", MB_ICONERROR);
            return 1;
        }

        if (!WaitForExcelClosed()) return 2;
        if (!EnsureDesktopRuntime()) return 2;

        Directory.CreateDirectory(InstallDir);
        using (var xll = typeof(Program).Assembly.GetManifestResourceStream("FormulaExplorer64.xll")
                         ?? throw new InvalidOperationException("The add-in is missing from this setup file."))
        using (var file = File.Create(XllPath))
            xll.CopyTo(file);

        // Keep a copy of this setup as the uninstaller.
        string self = Environment.ProcessPath ?? throw new InvalidOperationException("Cannot locate the setup file.");
        if (!string.Equals(Path.GetFullPath(self), Path.GetFullPath(UninstallerPath), StringComparison.OrdinalIgnoreCase))
            File.Copy(self, UninstallerPath, overwrite: true);

        AddOpenEntry(XllPath);
        RegisterUninstall();

        Show($"{DisplayName} is installed.\n\nStart Excel: the \"Studio\" tab appears after \"View\". " +
             "Select a cell with a formula and click Formula Explorer.", MB_ICONINFORMATION);
        return 0;
    }

    private static int Uninstall()
    {
        if (!Confirm($"Remove {DisplayName}?")) return 2;
        if (!WaitForExcelClosed()) return 2;

        RemoveOpenEntry(XllPath);
        Registry.CurrentUser.DeleteSubKeyTree(UninstallKey, throwOnMissingSubKey: false);
        TryDelete(XllPath);
        TryDelete(Path.Combine(InstallDir, "settings.json"));

        // This exe may be the installed uninstaller: delete it (and the folder, if empty) after we exit.
        string self = Environment.ProcessPath ?? string.Empty;
        if (string.Equals(Path.GetFullPath(self), Path.GetFullPath(UninstallerPath), StringComparison.OrdinalIgnoreCase))
        {
            Process.Start(new ProcessStartInfo("cmd.exe",
                $"/c ping -n 3 127.0.0.1 >nul & del /f /q \"{UninstallerPath}\" & rmdir \"{InstallDir}\"")
            { CreateNoWindow = true, UseShellExecute = false });
        }
        else
        {
            TryDelete(UninstallerPath);
        }

        Show($"{DisplayName} was removed.", MB_ICONINFORMATION);
        return 0;
    }

    // ------------------------------------------------------------------------------------------------------------
    // Checks

    private static string? ExcelPlatform()
    {
        using var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration");
        return key?.GetValue("Platform") as string;
    }

    private static bool WaitForExcelClosed()
    {
        while (Process.GetProcessesByName("EXCEL").Length > 0)
        {
            if (_quiet) return false;
            if (MessageBox(IntPtr.Zero, "Please close Excel (all windows) and click Retry.", Title,
                    MB_RETRYCANCEL | MB_ICONWARNING) != IDRETRY)
                return false;
        }
        return true;
    }

    private static bool HasDesktopRuntime()
    {
        string root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            "dotnet", "shared", "Microsoft.WindowsDesktop.App");
        return Directory.Exists(root) && Directory.EnumerateDirectories(root, "10.*").Any();
    }

    private static bool EnsureDesktopRuntime()
    {
        if (HasDesktopRuntime()) return true;
        if (_quiet) return false;

        int answer = MessageBox(IntPtr.Zero,
            "Formula Explorer needs the Microsoft .NET 10 Desktop Runtime (x64), which is not installed.\n\n" +
            "Download and install it now? (about 60 MB; Windows will ask for permission)\n\n" +
            "Yes = install it now\nNo = continue without it (the add-in will not load until it is installed)",
            Title, MB_YESNOCANCEL | MB_ICONQUESTION);
        if (answer == IDCANCEL) return false;
        if (answer == IDNO) return true;

        try
        {
            string installer = Path.Combine(Path.GetTempPath(), "windowsdesktop-runtime-10-win-x64.exe");
            using (var http = new HttpClient { Timeout = TimeSpan.FromMinutes(10) })
            using (var response = http.GetAsync(RuntimeDownload, HttpCompletionOption.ResponseHeadersRead).GetAwaiter().GetResult())
            {
                response.EnsureSuccessStatusCode();
                using var stream = response.Content.ReadAsStream();
                using var file = File.Create(installer);
                stream.CopyTo(file);
            }
            using (var p = Process.Start(new ProcessStartInfo(installer, "/install /passive /norestart") { UseShellExecute = true }))
                p?.WaitForExit();
            TryDelete(installer);
        }
        catch (Exception ex)
        {
            Show($"The .NET 10 Desktop Runtime could not be installed automatically:\n{ex.Message}\n\n" +
                 $"Install it from {RuntimePage} and run this setup again.", MB_ICONWARNING);
            Process.Start(new ProcessStartInfo(RuntimePage) { UseShellExecute = true });
            return false;
        }

        if (!HasDesktopRuntime())
        {
            Show($"The .NET 10 Desktop Runtime is still missing. Install it from {RuntimePage} and run this setup again.",
                MB_ICONWARNING);
            return false;
        }
        return true;
    }

    // ------------------------------------------------------------------------------------------------------------
    // Excel "open at startup" entries: OPEN, OPEN1, OPEN2... = /R "path" (kept contiguous)

    private static List<string> OpenValueNames(RegistryKey key) =>
        key.GetValueNames()
            .Where(n => n.Equals("OPEN", StringComparison.OrdinalIgnoreCase) ||
                        (n.StartsWith("OPEN", StringComparison.OrdinalIgnoreCase) && int.TryParse(n.AsSpan(4), out _)))
            .OrderBy(n => n.Length == 4 ? -1 : int.Parse(n.AsSpan(4)))
            .ToList();

    private static bool PointsTo(object? data, string xll)
    {
        string path = (data as string ?? string.Empty).Trim();
        if (path.StartsWith("/R", StringComparison.OrdinalIgnoreCase)) path = path[2..].Trim();
        return string.Equals(path.Trim('"'), xll, StringComparison.OrdinalIgnoreCase);
    }

    private static void AddOpenEntry(string xll)
    {
        using var key = Registry.CurrentUser.CreateSubKey(ExcelOptionsKey);
        var names = OpenValueNames(key);
        if (names.Any(n => PointsTo(key.GetValue(n), xll))) return;
        string name = "OPEN";
        for (int i = 1; names.Contains(name, StringComparer.OrdinalIgnoreCase); i++) name = "OPEN" + i;
        key.SetValue(name, $"/R \"{xll}\"", RegistryValueKind.String);
    }

    private static void RemoveOpenEntry(string xll)
    {
        using var key = Registry.CurrentUser.OpenSubKey(ExcelOptionsKey, writable: true);
        if (key is null) return;
        var names = OpenValueNames(key);
        if (!names.Any(n => PointsTo(key.GetValue(n), xll))) return;
        var keep = names.Select(n => key.GetValue(n)).Where(v => !PointsTo(v, xll)).ToList();
        foreach (var n in names) key.DeleteValue(n, throwOnMissingValue: false);
        for (int i = 0; i < keep.Count; i++)
            key.SetValue(i == 0 ? "OPEN" : "OPEN" + i, keep[i]!, RegistryValueKind.String);
    }

    private static void RegisterUninstall()
    {
        using var key = Registry.CurrentUser.CreateSubKey(UninstallKey);
        key.SetValue("DisplayName", DisplayName);
        key.SetValue("DisplayVersion", Version);
        key.SetValue("Publisher", "Formula Explorer");
        key.SetValue("InstallLocation", InstallDir);
        key.SetValue("DisplayIcon", UninstallerPath);
        key.SetValue("UninstallString", $"\"{UninstallerPath}\" /uninstall");
        key.SetValue("QuietUninstallString", $"\"{UninstallerPath}\" /uninstall /quiet");
        key.SetValue("NoModify", 1, RegistryValueKind.DWord);
        key.SetValue("NoRepair", 1, RegistryValueKind.DWord);
        key.SetValue("EstimatedSize", (int)((new FileInfo(XllPath).Length + new FileInfo(UninstallerPath).Length) / 1024),
            RegistryValueKind.DWord);
        key.SetValue("InstallDate", DateTime.Now.ToString("yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture));
    }

    // ------------------------------------------------------------------------------------------------------------

    private static void TryDelete(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { /* best effort */ }
    }

    private static bool Confirm(string text) =>
        _quiet || MessageBox(IntPtr.Zero, text, Title, MB_OKCANCEL | MB_ICONQUESTION) == IDOK;

    private static void Show(string text, uint icon)
    {
        if (!_quiet) MessageBox(IntPtr.Zero, text, Title, MB_OK | icon);
    }

    private const uint MB_OK = 0x0, MB_OKCANCEL = 0x1, MB_YESNOCANCEL = 0x3, MB_RETRYCANCEL = 0x5;
    private const uint MB_ICONERROR = 0x10, MB_ICONQUESTION = 0x20, MB_ICONWARNING = 0x30, MB_ICONINFORMATION = 0x40;
    private const int IDOK = 1, IDCANCEL = 2, IDRETRY = 4, IDNO = 7;

    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "MessageBoxW")]
    private static extern int MessageBox(IntPtr hWnd, string text, string caption, uint type);
}
