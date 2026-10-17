using System.Windows.Threading;

namespace FormulaExplorer.AddIn.Session;

/// <summary><see cref="IUiScheduler"/> on the WPF dispatcher of Excel's main thread.</summary>
public sealed class DispatcherUiScheduler : IUiScheduler
{
    private readonly Dispatcher _dispatcher;

    public DispatcherUiScheduler(Dispatcher dispatcher) => _dispatcher = dispatcher;

    public void Post(Action action) => _dispatcher.BeginInvoke(DispatcherPriority.Background, action);

    public void PostDelayed(TimeSpan delay, Action action)
    {
        var timer = new DispatcherTimer(DispatcherPriority.Background, _dispatcher) { Interval = delay };
        timer.Tick += (_, _) =>
        {
            timer.Stop();
            action();
        };
        timer.Start();
    }
}
