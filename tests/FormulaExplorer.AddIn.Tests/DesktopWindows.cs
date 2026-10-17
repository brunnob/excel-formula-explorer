namespace FormulaExplorer.AddIn.Tests;

/// <summary>
/// xUnit collection for every test class that creates or shows real top-level windows. Windows share ONE global z-order and ONE
/// foreground: a test that asserts "this window is directly above that one" or "this window is (still) active" is wrong whenever
/// another test class (running in parallel, which xUnit does between collections) creates, shows or activates a window meanwhile.
/// Classes in the same collection never run in parallel with each other.
/// </summary>
[CollectionDefinition(Name)]
public sealed class DesktopWindows
{
    public const string Name = "Desktop windows";
}
