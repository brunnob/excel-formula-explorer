using System.Runtime.ExceptionServices;
using System.Windows;
using System.Windows.Threading;
using FormulaExplorer.AddIn.Session;
using FormulaExplorer.AddIn.UI;
using FormulaExplorer.Core.Analysis;
using FormulaExplorer.Core.Syntax;

namespace FormulaExplorer.AddIn.Tests.UI;

/// <summary>Runs WPF code on a fresh STA thread with its own dispatcher (WPF objects must not be created on xUnit's MTA threads).</summary>
internal static class Sta
{
    public static void Run(Action action)
    {
        Exception? error = null;
        var thread = new Thread(() =>
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                error = ex;
            }
            finally
            {
                Dispatcher.CurrentDispatcher.InvokeShutdown();
            }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
        if (error is not null) ExceptionDispatchInfo.Capture(error).Throw();
    }

    /// <summary>Lets the dispatcher run (timers, posted work) for <paramref name="milliseconds"/>.</summary>
    public static void Pump(int milliseconds)
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
}

/// <summary>Waiting that depends on a condition instead of on the clock.</summary>
internal static class Wait
{
    /// <summary>
    /// Lets the dispatcher run until <paramref name="condition"/> holds; returns whether it did within
    /// <paramref name="timeoutMs"/>. The timeout only bounds a failing test: a loaded machine makes a passing one slower,
    /// never wrong (a fixed <c>Sta.Pump(n)</c> followed by an assertion is a race against the machine).
    /// </summary>
    public static bool Until(Func<bool> condition, int timeoutMs = 10_000)
    {
        var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);
        while (!condition())
        {
            if (DateTime.UtcNow >= deadline) return false;
            Sta.Pump(20);
        }
        return true;
    }
}

/// <summary>Fixed-pitch measurer: every char is <see cref="CharWidth"/> wide except line terminators (zero width).</summary>
internal sealed class FixedMeasurer : ITextMeasurer
{
    private readonly string _text;
    private readonly double _charWidth;

    public FixedMeasurer(string text, double charWidth = 10)
    {
        _text = text;
        _charWidth = charWidth;
    }

    public double Measure(int start, int length) => Offsets(start, length)[length];

    public double[] Offsets(int start, int length)
    {
        var offsets = new double[length + 1];
        for (int i = 0; i < length; i++)
            offsets[i + 1] = offsets[i] + (_text[start + i] is '\r' or '\n' ? 0 : _charWidth);
        return offsets;
    }
}

/// <summary>Node info provider driven by a delegate; counts the queries.</summary>
internal sealed class FakeProvider : INodeInfoProvider
{
    private readonly Func<SyntaxNode, NodeInfo> _info;

    public FakeProvider(Func<SyntaxNode, NodeInfo> info) => _info = info;

    public List<SyntaxNode> Queries { get; } = new();

    /// <summary>Links for references/names, "value" tooltip for everything.</summary>
    public static FakeProvider Standard() => new(node => new NodeInfo(node.Kind is NodeKind.Reference or NodeKind.Name, "value of " + node.Kind));

    public NodeInfo GetNodeInfo(SyntaxNode node)
    {
        Queries.Add(node);
        return _info(node);
    }
}

internal static class Displays
{
    public static FormulaDisplay Of(string text)
    {
        var parsed = FormulaParser.Parse(text);
        var colors = RefColorizer.Assign(parsed, "Sheet1", _ => true);
        return new FormulaDisplay(parsed, HitMap.Build(parsed), colors, parsed.IsFormula);
    }

    public static SyntaxNode Find(FormulaDisplay display, string text, NodeKind kind) =>
        display.Formula.Nodes.First(n => n.Kind == kind && n.GetText(display.Formula.Text) == text);

    /// <summary>A formula view measured at the default text width (390 DIPs wide box → 384 + insets).</summary>
    public static FormulaView NewView(FakeProvider provider, FormulaDisplay display, double width = 388)
    {
        var view = new FormulaView(provider);
        view.ShowFormula(display);
        view.Measure(new Size(width, 400));
        view.Arrange(new Rect(0, 0, width, view.DesiredSize.Height));
        view.UpdateLayout();
        return view;
    }
}
