using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using FormulaExplorer.AddIn.Diagnostics;

namespace FormulaExplorer.AddIn.Ribbon;

/// <summary>
/// Loads the embedded globe PNGs and hands them to the ribbon as <c>IPictureDisp</c> objects (what the ribbon's
/// <c>getImage</c> callback must return). The 32-bit alpha of the PNGs is preserved by the conversion.
/// </summary>
internal static class RibbonIcon
{
    /// <summary>Pixel sizes of the embedded images (see scripts/render-icon.ps1), ascending.</summary>
    private static readonly int[] Sizes = { 16, 32, 64 };

    private static readonly Dictionary<int, object> Cache = new();

    /// <summary>
    /// The smallest embedded image that is at least <paramref name="pixels"/> wide (so Excel only ever scales
    /// down), or the largest one when none is. Null when the resource cannot be loaded.
    /// </summary>
    public static object? GetPicture(int pixels)
    {
        var size = Sizes.FirstOrDefault(s => s >= pixels);
        if (size == 0) size = Sizes[^1];
        lock (Cache)
        {
            if (Cache.TryGetValue(size, out var cached)) return cached;
            try
            {
                var picture = Load(size);
                Cache[size] = picture;
                return picture;
            }
            catch (Exception ex)
            {
                Log.Error($"Loading ribbon icon {size}px", ex);
                return null;
            }
        }
    }

    private static object Load(int size)
    {
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream($"FormulaExplorer.Ribbon.globe{size}.png")
            ?? throw new InvalidOperationException($"Embedded resource globe{size}.png not found.");
        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        copy.Position = 0;
        // The bitmap must keep its stream open for its whole life; GDI+ requires it. Clone into an owned bitmap so
        // the stream can go away: Bitmap(Image) copies the pixels (with alpha) into a fresh 32-bpp ARGB bitmap.
        using var decoded = Image.FromStream(copy);
        using var bitmap = new Bitmap(decoded);
        return PictureConverter.ToPictureDisp(bitmap);
    }

    /// <summary>The OLE picture conversion is a protected static member of AxHost, hence this subclass.</summary>
    private sealed class PictureConverter : AxHost
    {
        // AxHost insists on a CLSID string even though no ActiveX control is ever hosted; the type is never instantiated.
        private PictureConverter() : base("00000000-0000-0000-0000-000000000000") { }

        public static object ToPictureDisp(Image image)
            => GetIPictureDispFromPicture(image) ?? throw new InvalidOperationException("IPictureDisp conversion returned null.");
    }

    /// <summary>Scale factor (1.0 = 96 DPI) of the window with the given handle; 1.0 when unknown.</summary>
    public static double GetDpiScale(IntPtr hwnd)
    {
        try
        {
            var dpi = hwnd != IntPtr.Zero ? GetDpiForWindow(hwnd) : GetDpiForSystem();
            return dpi >= 48 ? dpi / 96.0 : 1.0;
        }
        catch (Exception ex) when (ex is EntryPointNotFoundException or DllNotFoundException)
        {
            return 1.0;
        }
    }

    [DllImport("user32.dll")]
    private static extern uint GetDpiForWindow(IntPtr hwnd);

    [DllImport("user32.dll")]
    private static extern uint GetDpiForSystem();
}
