using FormulaExplorer.AddIn.Diagnostics;

namespace FormulaExplorer.AddIn.Overlay;

/// <summary>
/// Owns the single <see cref="OverlayWindow"/> of the overlay and keeps it usable: creates it on first use, re-owns it when
/// the root's workbook window changed and recreates it when the native window is gone (SDI Excel destroys a workbook
/// window's XLMAIN, and every window it owns, when the workbook is closed).
/// </summary>
internal sealed class OverlayWindowKeeper : IDisposable
{
    private readonly Func<IntPtr, OverlayWindow> _factory;

    public OverlayWindowKeeper()
        : this(owner => new OverlayWindow(owner))
    {
    }

    internal OverlayWindowKeeper(Func<IntPtr, OverlayWindow> factory) => _factory = factory;

    /// <summary>The current window, or null before the first <see cref="Ensure"/>. It may be dead; use <see cref="Ensure"/> to get a live one.</summary>
    public OverlayWindow? Current { get; private set; }

    /// <summary>A live window owned by <paramref name="owner"/>.</summary>
    public OverlayWindow Ensure(IntPtr owner)
    {
        var window = Current;
        if (window is null)
        {
            window = Current = _factory(owner);
            Log.Info($"Overlay window created (owner 0x{owner.ToInt64():X}).");
        }
        else if (!window.IsAlive)
        {
            // Destroyed with its owner: SetOwner/SetWindowPos on the dead handle would fail forever. Release what is left
            // (class registration, brushes) and start over.
            window.Dispose();
            Current = null;
            window = Current = _factory(owner);
            Log.Info($"Overlay window recreated (its owner was destroyed; new owner 0x{owner.ToInt64():X}).");
        }
        else if (window.Owner != owner)
        {
            window.SetOwner(owner);
            Log.Info($"Overlay window re-owned to 0x{owner.ToInt64():X}.");
        }

        return window;
    }

    public void Dispose()
    {
        Current?.Dispose();
        Current = null;
    }
}
