using System.IO;
using System.Text.Json;
using System.Text.Json.Nodes;
using FormulaExplorer.AddIn.Diagnostics;

namespace FormulaExplorer.AddIn.UI;

/// <summary>Outer window bounds in DIPs (what WPF's Left/Top/Width/Height use).</summary>
internal readonly record struct WindowBounds(double Left, double Top, double Width, double Height);

/// <summary>
/// Persistence of the explorer window's position and size in <c>%LOCALAPPDATA%\FormulaExplorer\settings.json</c>
/// (<c>{ "window": { "left", "top", "width", "height" } }</c>; other top-level keys are preserved). Reading never throws:
/// a missing, corrupt or implausible file just means "no remembered bounds".
/// </summary>
internal static class WindowSettings
{
    private const string WindowKey = "window";
    private const double MinSize = 50;
    private const double MaxSize = 20000;
    private const double MaxCoordinate = 100000;

    /// <summary>The visible part of a restored window must cover at least this much of its title strip (pixels).</summary>
    private const int MinVisibleWidth = 64;
    private const int MinVisibleHeight = 16;
    private const int TitleStripHeight = 32;

    public static string DefaultPath { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "FormulaExplorer", "settings.json");

    /// <summary>Reads the saved bounds, or null when absent/invalid.</summary>
    public static WindowBounds? Load(string? path = null)
    {
        try
        {
            path ??= DefaultPath;
            if (!File.Exists(path)) return null;
            if (JsonNode.Parse(File.ReadAllText(path)) is not JsonObject root) return null;
            if (root[WindowKey] is not JsonObject window) return null;

            var bounds = new WindowBounds(
                ReadNumber(window, "left"), ReadNumber(window, "top"), ReadNumber(window, "width"), ReadNumber(window, "height"));
            return IsPlausible(bounds) ? bounds : null;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException or InvalidOperationException or FormatException or ArgumentException)
        {
            // ArgumentException: a duplicate property name, which the lazily parsed JsonNode only reports on first access.
            Log.Error("WindowSettings.Load", ex);
            return null;
        }
    }

    /// <summary>Writes the bounds (keeping unrelated keys of the file). Failures are logged, never thrown.</summary>
    public static void Save(WindowBounds bounds, string? path = null)
    {
        try
        {
            if (!IsPlausible(bounds)) return;
            path ??= DefaultPath;
            JsonObject root = TryReadObject(path) ?? new JsonObject();
            root[WindowKey] = new JsonObject
            {
                ["left"] = Math.Round(bounds.Left, 2),
                ["top"] = Math.Round(bounds.Top, 2),
                ["width"] = Math.Round(bounds.Width, 2),
                ["height"] = Math.Round(bounds.Height, 2),
            };
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            string temp = path + ".tmp";
            File.WriteAllText(temp, root.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
            File.Move(temp, path, overwrite: true);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException or InvalidOperationException or ArgumentException)
        {
            Log.Error("WindowSettings.Save", ex);
        }
    }

    /// <summary>True when all four numbers are finite and the size is within sane limits.</summary>
    public static bool IsPlausible(WindowBounds b) =>
        double.IsFinite(b.Left) && double.IsFinite(b.Top) && double.IsFinite(b.Width) && double.IsFinite(b.Height)
        && Math.Abs(b.Left) <= MaxCoordinate && Math.Abs(b.Top) <= MaxCoordinate
        && b.Width >= MinSize && b.Width <= MaxSize && b.Height >= MinSize && b.Height <= MaxSize;

    /// <summary>
    /// True when enough of the window's title strip lies inside some monitor work area for the user to grab it.
    /// </summary>
    /// <param name="scale">DIPs to physical pixels (DPI / 96).</param>
    public static bool IsOnScreen(WindowBounds dipBounds, double scale, IReadOnlyList<PixelRect> workAreas)
    {
        if (!IsPlausible(dipBounds) || scale <= 0) return false;
        double left = dipBounds.Left * scale;
        double top = dipBounds.Top * scale;
        double right = left + dipBounds.Width * scale;
        double strip = top + TitleStripHeight;

        foreach (var area in workAreas)
        {
            double width = Math.Min(right, area.Right) - Math.Max(left, area.Left);
            double height = Math.Min(strip, area.Bottom) - Math.Max(top, area.Top);
            if (width >= MinVisibleWidth && height >= MinVisibleHeight) return true;
        }
        return false;
    }

    private static JsonObject? TryReadObject(string path)
    {
        try
        {
            if (!File.Exists(path)) return null;
            var root = JsonNode.Parse(File.ReadAllText(path)) as JsonObject;
            _ = root?.Count; // parsing is lazy: force it here so a duplicate property name is caught below, not later in Save
            return root;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException or ArgumentException)
        {
            return null; // unreadable file: it is replaced
        }
    }

    private static double ReadNumber(JsonObject obj, string key) =>
        obj[key] is JsonValue value && value.TryGetValue(out double number) ? number : double.NaN;
}
