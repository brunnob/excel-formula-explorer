using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using FormulaExplorer.AddIn.Session;
using FormulaExplorer.Core.Syntax;

namespace FormulaExplorer.AddIn.UI;

/// <summary>
/// The client area of the explorer window: formula box (1 px #646464 border, white, scrolling) above right-aligned
/// Back/Close buttons. It is a plain element so the snapshot tool can render it without a window and at any DPI.
/// Metrics (DIPs) reproduce the shipped original: 12 from the left/top/right, box 390×70, buttons 75×23, 7 above and
/// below, 8 apart, Close flush with the box's right edge → a 414×119 client.
/// </summary>
internal sealed class ExplorerPanel : Border
{
    public const double ClientMargin = 12;
    public const double DefaultBoxWidth = 390;
    public const double DefaultBoxHeight = 70;
    public const double ButtonWidth = 75;
    public const double ButtonHeight = 23;
    public const double ButtonGap = 7;
    public const double ButtonSpacing = 8;
    public const double MinBoxWidth = 200;
    public const double MinBoxHeight = 20;

    /// <summary>Client size when the box has its default size.</summary>
    public static Size DefaultClientSize { get; } = new(
        ClientMargin + DefaultBoxWidth + ClientMargin,
        ClientMargin + DefaultBoxHeight + ButtonGap + ButtonHeight + ButtonGap);

    /// <summary>Smallest client size: a 200×20 formula box.</summary>
    public static Size MinClientSize { get; } = new(
        ClientMargin + MinBoxWidth + ClientMargin,
        ClientMargin + MinBoxHeight + ButtonGap + ButtonHeight + ButtonGap);

    private static readonly Brush BoxBorderBrush = Frozen(new SolidColorBrush(Color.FromRgb(0x64, 0x64, 0x64)));

    public ExplorerPanel(INodeInfoProvider provider)
    {
        View = new FormulaView(provider);
        Scroll = new ScrollViewer
        {
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            Content = View,
            Focusable = false,
        };
        View.ScrollHost = Scroll;
        Scroll.ScrollChanged += (_, _) => View.RefreshPointer();

        Box = new Border
        {
            BorderBrush = BoxBorderBrush,
            BorderThickness = new Thickness(1),
            Background = Brushes.White,
            Child = Scroll,
            Width = DefaultBoxWidth,
            Height = DefaultBoxHeight,
            SnapsToDevicePixels = true,
        };

        BackButton = CreateButton("Back", "Retraces your steps");
        BackButton.IsEnabled = false;
        CloseButton = CreateButton("Close", "Exits Formula Explorer");
        CloseButton.IsDefault = true;
        CloseButton.Margin = new Thickness(ButtonSpacing, 0, 0, 0);

        var buttons = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, ButtonGap, 0, ButtonGap),
        };
        buttons.Children.Add(BackButton);
        buttons.Children.Add(CloseButton);

        var grid = new Grid();
        grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        Grid.SetRow(Box, 0);
        Grid.SetRow(buttons, 1);
        grid.Children.Add(Box);
        grid.Children.Add(buttons);

        // Transparent so the whole client area takes part in hit testing (busy cursor over the margins too).
        Background = Brushes.Transparent;
        Padding = new Thickness(ClientMargin, ClientMargin, ClientMargin, 0);
        Child = grid;
        UseLayoutRounding = true;
        SnapsToDevicePixels = true;
        TextOptions.SetTextFormattingMode(this, TextFormattingMode.Display);
    }

    public FormulaView View { get; }

    public ScrollViewer Scroll { get; }

    public Border Box { get; }

    public Button BackButton { get; }

    public Button CloseButton { get; }

    /// <summary>Gives the box its default fixed size (used while the window sizes itself to its content).</summary>
    public void FixBoxSize()
    {
        Box.Width = DefaultBoxWidth;
        Box.Height = DefaultBoxHeight;
    }

    /// <summary>Lets the box stretch with the window.</summary>
    public void ReleaseBoxSize()
    {
        Box.Width = double.NaN;
        Box.Height = double.NaN;
    }

    public void ShowLoading()
    {
        View.ShowLoading();
        Cursor = Cursors.Wait;
        ForceCursor = true;
    }

    public void ShowFormula(FormulaDisplay display)
    {
        View.ShowFormula(display);
        Scroll.ScrollToHome();
        ClearValue(CursorProperty);
        ClearValue(ForceCursorProperty);
    }

    /// <summary>Scrolls so that the char is inside the viewport (test hooks drive the view this way).</summary>
    public void EnsureCharVisible(int charIndex)
    {
        UpdateLayout();
        if (View.GetCharRect(charIndex) is not { } rect) return;
        if (rect.Top < Scroll.VerticalOffset) Scroll.ScrollToVerticalOffset(rect.Top - FormulaView.TextInset.Top);
        else if (rect.Bottom > Scroll.VerticalOffset + Scroll.ViewportHeight)
            Scroll.ScrollToVerticalOffset(rect.Bottom - Scroll.ViewportHeight);
        UpdateLayout();
    }

    private static Button CreateButton(string text, string toolTip)
    {
        var button = new Button { Content = text, Width = ButtonWidth, Height = ButtonHeight, ToolTip = toolTip };
        ToolTipService.SetShowOnDisabled(button, true);
        return button;
    }

    private static T Frozen<T>(T freezable) where T : Freezable
    {
        freezable.Freeze();
        return freezable;
    }
}
