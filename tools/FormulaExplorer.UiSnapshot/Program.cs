using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using FormulaExplorer.AddIn.Session;
using FormulaExplorer.AddIn.UI;
using FormulaExplorer.Core.Analysis;
using FormulaExplorer.Core.Syntax;

namespace FormulaExplorer.UiSnapshot;

/// <summary>
/// Renders the explorer's client area with fake data to PNG files (no Excel, no window on screen) so the look can be
/// compared with docs/referencias/2011-site-*.png. Usage: UiSnapshot [outputDirectory].
/// </summary>
internal static class Program
{
    /// <summary>Client size of the 2011 site screenshots (box 511×30; the Windows 7 frame adds 8 px on each side).</summary>
    private static readonly Size SiteClient = new(535, 79);

    [STAThread]
    private static int Main(string[] args)
    {
        string outDir = args.Length > 0 ? args[0] : System.IO.Path.Combine(AppContext.BaseDirectory, "out");
        Directory.CreateDirectory(outDir);

        foreach (var scenario in Scenarios())
        {
            foreach (double scale in scenario.Scales)
            {
                string suffix = scale == 1.0 ? string.Empty : "-x" + scale.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture);
                string path = System.IO.Path.Combine(outDir, scenario.Name + suffix + ".png");
                var info = Render(scenario, scale, path);
                Console.WriteLine($"{path}  {info}");
            }
        }

        // The real window (title bar, frame, first-open size), copied with PrintWindow; nothing is clicked or typed.
        string windowPath = System.IO.Path.Combine(outDir, "real-window.png");
        Console.WriteLine($"{windowPath}  {WindowCapture.Capture(Site1, windowPath, p => p.Text.IndexOf("Forcast", StringComparison.Ordinal))}");
        return 0;
    }

    private const string Site1 = "=IF(C3,0,C11*'Inventory Forcast'!C$6)";
    private const string Site2 = "=IF(K3=\"On\",1/(((-($I26/Hundred)+1)^(K$31))),0)";
    private const string Site3 = "=IF(J3,(1/((365*(($I27/Hundred)+1)^(K$31)))),0)";
    private const string Offset = "=SUM(I$19:$I$19)+J$19*$I$28/Hundred-SUM(OFFSET($I$19,0,0,1,Aqui_Per))";
    private const string Long = "=IF(K3=\"On\",1/(((-($I26/Hundred)+1)^(K$31))),0)+SUM(I$19:$I$19)+J$19*$I$28/Hundred-SUM(OFFSET($I$19,0,0,1,Aqui_Per))+Sheet8!E7+(E30/C4)^(D4/D5)+D14-E21*E16+Summary!C5";

    private static IEnumerable<Scenario> Scenarios()
    {
        double[] both = { 1.0, 1.5 };

        // The three official site screenshots: pointer positions are taken from where the tooltips sit there.
        yield return new("site-01-hyperlink-hover", Site1, SiteClient, true, both) { PointerX = 155, PointerY = 9, Tip = "25" };
        yield return new("site-02-fragment-solving", Site2, SiteClient, true, both) { PointerX = 235, PointerY = 9, Tip = "0.8210896996" };
        yield return new("site-03-bracket-forest", Site3, SiteClient, true, both) { PointerX = 272, PointerY = 9, Tip = "411.66306" };

        // Default window (414x119 client, 390x70 box).
        yield return new("default-wrapped-long", Long, ExplorerPanelSize(), false, both);
        yield return new("default-hover-offset", Offset, ExplorerPanelSize(), false, both) { HoverText = "OFFSET", Tip = "I19:J19 containing\n[1x2] , 0.252054794520548" };
        yield return new("default-gold-link", Offset, ExplorerPanelSize(), false, both) { GoldText = "$I$28" };
        yield return new("default-gold-hovered", Offset, ExplorerPanelSize(), false, both) { GoldText = "$I$28", HoverText = "$I$28", Tip = "7" };
        yield return new("default-gold-other-hover", Offset, ExplorerPanelSize(), false, both) { GoldText = "$I$28", HoverText = "J$19", Tip = "0.252054794520548" };
        yield return new("default-hover-paren", Site3, ExplorerPanelSize(), false, both) { HoverText = "((365", HoverOffset = 1, Tip = "411.66306" };
        yield return new("default-non-formula", "Quarterly forecast 2012", ExplorerPanelSize(), false, both);
        yield return new("default-empty", string.Empty, ExplorerPanelSize(), false, both);
        yield return new("default-loading", Site1, ExplorerPanelSize(), false, both) { Loading = true };
    }

    private static Size ExplorerPanelSize() => ExplorerPanel.DefaultClientSize;

    private static string Render(Scenario s, double scale, string path)
    {
        var parsed = FormulaParser.Parse(s.Formula);
        var hitMap = HitMap.Build(parsed);
        var colors = RefColorizer.Assign(parsed, "Sheet1", n => n is ReferenceNode { Form: not ReferenceForm.RefError } or NameNode or StructuredReferenceNode);
        var display = new FormulaDisplay(parsed, hitMap, colors, parsed.IsFormula);

        var provider = new FakeProvider(s);
        var panel = new ExplorerPanel(provider);
        if (s.StretchBox) panel.ReleaseBoxSize();

        var overlay = new Canvas { IsHitTestVisible = false };
        var root = new Grid { Width = s.Client.Width, Height = s.Client.Height, Background = new SolidColorBrush(Color.FromRgb(0xF0, 0xF0, 0xF0)) };
        root.Children.Add(panel);
        root.Children.Add(overlay);
        VisualTreeHelper.SetRootDpi(root, new DpiScale(scale, scale));
        Layout(root, s.Client);

        if (s.Loading)
            panel.ShowLoading();
        else
            panel.ShowFormula(display);
        Layout(root, s.Client);

        var view = panel.View;
        if (s.GoldText is { } goldText) view.SetGold(Find(parsed, goldText));

        Point? pointer = null;
        if (s.PointerX is { } px) pointer = new Point(px, s.PointerY ?? FormulaView.TextInset.Top + TextLayout.LinePitch / 2);
        else if (s.HoverText is { } hoverText)
            pointer = view.GetCharCenter(parsed.Text.IndexOf(hoverText, StringComparison.Ordinal) + s.HoverOffset);
        if (pointer is { } p)
        {
            view.SimulateMove(p);
            DrawPointer(overlay, view.TransformToAncestor(root).Transform(p));
            if (view.TooltipVisible && view.TooltipPlacement is { } at)
            {
                var tip = ReturnsToolTipVisual.Create(view.TooltipText!);
                var origin = view.TransformToAncestor(root).Transform(at);
                Canvas.SetLeft(tip, origin.X);
                Canvas.SetTop(tip, origin.Y);
                overlay.Children.Insert(0, tip);
            }
        }
        Layout(root, s.Client);

        var bitmap = new RenderTargetBitmap((int)Math.Round(s.Client.Width * scale), (int)Math.Round(s.Client.Height * scale),
            96 * scale, 96 * scale, PixelFormats.Pbgra32);
        bitmap.Render(root);
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(bitmap));
        using (var stream = File.Create(path)) encoder.Save(stream);

        return $"{bitmap.PixelWidth}x{bitmap.PixelHeight}, hover={(view.HoverNode is null ? "-" : view.HoverNode.GetText(parsed.Text))}, tooltip={(view.TooltipVisible ? "yes" : "no")}";
    }

    private static void Layout(FrameworkElement root, Size size)
    {
        root.Measure(size);
        root.Arrange(new Rect(size));
        root.UpdateLayout();
    }

    /// <summary>The node whose text is <paramref name="text"/> (or the leaf covering the match when it is not a whole node).</summary>
    private static SyntaxNode? Find(ParsedFormula parsed, string text)
    {
        int at = parsed.Text.IndexOf(text, StringComparison.Ordinal);
        if (at < 0) return null;
        var exact = parsed.Nodes.Where(n => n.Start == at && n.Length == text.Length)
            .OrderByDescending(n => n.Kind is NodeKind.Reference or NodeKind.Name ? 1 : 0).FirstOrDefault();
        return exact ?? HitMap.Build(parsed).NodeAt(at);
    }

    private static void DrawPointer(Canvas canvas, Point at)
    {
        var arrow = new Polygon
        {
            Points = new PointCollection { new(0, 0), new(0, 16), new(4, 12), new(7, 18), new(9, 17), new(6, 11), new(11, 11) },
            Fill = Brushes.White,
            Stroke = Brushes.Black,
            StrokeThickness = 1,
            StrokeLineJoin = PenLineJoin.Miter,
        };
        Canvas.SetLeft(arrow, at.X);
        Canvas.SetTop(arrow, at.Y);
        canvas.Children.Add(arrow);
    }

    private sealed record Scenario(string Name, string Formula, Size Client, bool StretchBox, double[] Scales)
    {
        /// <summary>Pointer x in the formula view's coordinates (on the first line); overrides <see cref="HoverText"/>.</summary>
        public double? PointerX { get; init; }

        /// <summary>Pointer y in the view's coordinates (default: the first line's center).</summary>
        public double? PointerY { get; init; }

        public string? HoverText { get; init; }

        public int HoverOffset { get; init; }

        public string? GoldText { get; init; }

        public string? Tip { get; init; }

        public bool Loading { get; init; }
    }

    /// <summary>Links for references, names and OFFSET; every node gets the scenario's canned tooltip.</summary>
    private sealed class FakeProvider : INodeInfoProvider
    {
        private readonly Scenario _scenario;

        public FakeProvider(Scenario scenario) => _scenario = scenario;

        public NodeInfo GetNodeInfo(SyntaxNode node)
        {
            bool link = node is ReferenceNode or NameNode || (node is FunctionCallNode call && call.Name.Equals("OFFSET", StringComparison.OrdinalIgnoreCase));
            return new NodeInfo(link, _scenario.Tip ?? "1.062");
        }
    }
}
