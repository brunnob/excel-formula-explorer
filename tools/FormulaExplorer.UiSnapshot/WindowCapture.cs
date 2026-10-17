using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using FormulaExplorer.AddIn.UI;
using FormulaExplorer.Core.Analysis;
using FormulaExplorer.Core.Syntax;

namespace FormulaExplorer.UiSnapshot;

/// <summary>
/// Shows the real <see cref="ExplorerWindow"/> (not activated, bounds saved to a throw-away settings file) and copies
/// it with PrintWindow, so the title bar and frame can be checked as well. No mouse or keyboard input is generated.
/// </summary>
internal static class WindowCapture
{
    [DllImport("user32.dll")]
    private static extern bool PrintWindow(IntPtr hwnd, IntPtr hdc, uint flags);

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hwnd, out Rect32 rect);

    [DllImport("user32.dll")]
    private static extern IntPtr GetDC(IntPtr hwnd);

    [DllImport("user32.dll")]
    private static extern int ReleaseDC(IntPtr hwnd, IntPtr hdc);

    [DllImport("gdi32.dll")]
    private static extern IntPtr CreateCompatibleDC(IntPtr hdc);

    [DllImport("gdi32.dll")]
    private static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int width, int height);

    [DllImport("gdi32.dll")]
    private static extern IntPtr SelectObject(IntPtr hdc, IntPtr obj);

    [DllImport("gdi32.dll")]
    private static extern bool DeleteObject(IntPtr obj);

    [DllImport("gdi32.dll")]
    private static extern bool DeleteDC(IntPtr hdc);

    private struct Rect32
    {
        public int Left, Top, Right, Bottom;
    }

    private const uint PW_RENDERFULLCONTENT = 2;

    public static string Capture(string formula, string path, Func<ParsedFormula, int?> hoverChar)
    {
        string settings = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "fe-uisnapshot-" + Guid.NewGuid().ToString("N") + ".json");
        var parsed = FormulaParser.Parse(formula);
        var display = new FormulaExplorer.AddIn.Session.FormulaDisplay(parsed, HitMap.Build(parsed),
            RefColorizer.Assign(parsed, "Sheet1", _ => true), parsed.IsFormula);

        var window = new ExplorerWindow(IntPtr.Zero, new CannedProvider(), settings) { ShowActivated = false };
        window.Show();
        window.ShowFormula(display);
        Pump(400);
        if (hoverChar(parsed) is { } index) window.SimulateHover(index);
        Pump(300);

        IntPtr hwnd = window.Handle;
        GetWindowRect(hwnd, out var rect);
        int width = rect.Right - rect.Left, height = rect.Bottom - rect.Top;
        IntPtr screen = GetDC(IntPtr.Zero);
        IntPtr dc = CreateCompatibleDC(screen);
        IntPtr bitmap = CreateCompatibleBitmap(screen, width, height);
        IntPtr old = SelectObject(dc, bitmap);
        PrintWindow(hwnd, dc, PW_RENDERFULLCONTENT);
        SelectObject(dc, old);

        var source = Imaging.CreateBitmapSourceFromHBitmap(bitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(source));
        using (var stream = File.Create(path)) encoder.Save(stream);

        DeleteObject(bitmap);
        DeleteDC(dc);
        ReleaseDC(IntPtr.Zero, screen);
        var snapshot = window.GetSnapshot();
        window.CloseWindow();
        File.Delete(settings);
        return $"{width}x{height} px, ppd={snapshot.PixelsPerDip}, tooltip={snapshot.TooltipVisible}";
    }

    private static void Pump(int milliseconds)
    {
        var frame = new DispatcherFrame();
        var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(milliseconds) };
        timer.Tick += (_, _) =>
        {
            timer.Stop();
            frame.Continue = false;
        };
        timer.Start();
        Dispatcher.PushFrame(frame);
    }

    private sealed class CannedProvider : FormulaExplorer.AddIn.Session.INodeInfoProvider
    {
        public FormulaExplorer.AddIn.Session.NodeInfo GetNodeInfo(SyntaxNode node) =>
            new(node is ReferenceNode or NameNode, "1.062");
    }
}
