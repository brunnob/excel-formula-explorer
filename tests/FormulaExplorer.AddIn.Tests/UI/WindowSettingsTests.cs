using System.IO;
using System.Text.Json.Nodes;
using FormulaExplorer.AddIn.UI;

namespace FormulaExplorer.AddIn.Tests.UI;

public sealed class WindowSettingsTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "fe-settings-" + Guid.NewGuid().ToString("N"));

    private string SettingsPath => Path.Combine(_dir, "settings.json");

    public void Dispose()
    {
        if (Directory.Exists(_dir)) Directory.Delete(_dir, recursive: true);
    }

    [Fact]
    public void MissingFileMeansNoBounds() => Assert.Null(WindowSettings.Load(SettingsPath));

    [Fact]
    public void RoundTrip()
    {
        var bounds = new WindowBounds(120.5, 80, 446, 158.25);
        WindowSettings.Save(bounds, SettingsPath);
        Assert.Equal(bounds, WindowSettings.Load(SettingsPath));
    }

    [Fact]
    public void SaveCreatesTheFolderAndLeavesNoTempFile()
    {
        WindowSettings.Save(new WindowBounds(1, 2, 300, 200), SettingsPath);
        Assert.True(File.Exists(SettingsPath));
        Assert.False(File.Exists(SettingsPath + ".tmp"));
    }

    [Fact]
    public void SavePreservesOtherKeys()
    {
        Directory.CreateDirectory(_dir);
        File.WriteAllText(SettingsPath, "{ \"other\": { \"x\": 1 }, \"window\": { \"left\": 0 } }");
        WindowSettings.Save(new WindowBounds(10, 20, 300, 200), SettingsPath);

        var root = JsonNode.Parse(File.ReadAllText(SettingsPath))!.AsObject();
        Assert.Equal(1, (int)root["other"]!["x"]!);
        Assert.Equal(300, (double)root["window"]!["width"]!);
    }

    [Theory]
    [InlineData("")]
    [InlineData("not json")]
    [InlineData("[1,2,3]")]
    [InlineData("{ }")]
    [InlineData("{ \"window\": 5 }")]
    [InlineData("{ \"window\": { \"left\": 1, \"top\": 2, \"width\": 300 } }")]
    [InlineData("{ \"window\": { \"left\": \"a\", \"top\": 2, \"width\": 300, \"height\": 200 } }")]
    [InlineData("{ \"window\": { \"left\": 1, \"top\": 2, \"width\": 3, \"height\": 200 } }")]
    [InlineData("{ \"window\": { \"left\": 1, \"top\": 2, \"width\": 300, \"height\": -4 } }")]
    [InlineData("{ \"window\": { \"left\": 1e300, \"top\": 2, \"width\": 300, \"height\": 200 } }")]
    public void CorruptOrImplausibleFilesMeanNoBounds(string content)
    {
        Directory.CreateDirectory(_dir);
        File.WriteAllText(SettingsPath, content);
        Assert.Null(WindowSettings.Load(SettingsPath));
    }

    [Fact]
    public void ACorruptFileIsReplacedBySave()
    {
        Directory.CreateDirectory(_dir);
        File.WriteAllText(SettingsPath, "{ broken");
        WindowSettings.Save(new WindowBounds(1, 2, 300, 200), SettingsPath);
        Assert.Equal(new WindowBounds(1, 2, 300, 200), WindowSettings.Load(SettingsPath));
    }

    [Fact]
    public void ImplausibleBoundsAreNotSaved()
    {
        WindowSettings.Save(new WindowBounds(double.NaN, 0, 300, 200), SettingsPath);
        Assert.False(File.Exists(SettingsPath));
    }

    [Theory]
    [InlineData(0, 0, 446, 158, true)]
    [InlineData(-100, 10, 446, 158, true)]           // partly off to the left but the title bar is reachable
    [InlineData(-400, 10, 446, 158, false)]          // only 46 px of the title bar left on screen
    [InlineData(1900, 10, 446, 158, false)]          // 20 px visible at the right edge of a 1920 monitor
    [InlineData(1500, 10, 446, 158, true)]
    [InlineData(100, -50, 446, 158, false)]          // title bar above the work area
    [InlineData(100, 1060, 446, 158, false)]         // below the work area
    [InlineData(100, 1000, 446, 158, true)]
    public void IsOnScreenNeedsAGrabbableTitleBar(double left, double top, double width, double height, bool expected)
    {
        var workAreas = new[] { new PixelRect(0, 0, 1920, 1040) };
        Assert.Equal(expected, WindowSettings.IsOnScreen(new WindowBounds(left, top, width, height), 1.0, workAreas));
    }

    [Fact]
    public void IsOnScreenScalesDipsToPixels()
    {
        var workAreas = new[] { new PixelRect(0, 0, 1920, 1040) };
        // 1000 DIPs at 150% are 1500 px: on the monitor. At 2.0 they are 2000 px: off it.
        var bounds = new WindowBounds(1000, 10, 400, 150);
        Assert.True(WindowSettings.IsOnScreen(bounds, 1.5, workAreas));
        Assert.False(WindowSettings.IsOnScreen(bounds, 2.0, workAreas));
    }

    [Fact]
    public void IsOnScreenWorksAcrossMonitors()
    {
        var workAreas = new[] { new PixelRect(0, 0, 1920, 1040), new PixelRect(-1280, 0, 0, 984) };
        Assert.True(WindowSettings.IsOnScreen(new WindowBounds(-1000, 100, 446, 158), 1.0, workAreas));
        Assert.False(WindowSettings.IsOnScreen(new WindowBounds(-1000, 100, 446, 158), 1.0, new[] { new PixelRect(0, 0, 1920, 1040) }));
    }

    [Fact]
    public void NoMonitorsMeansOffScreen() =>
        Assert.False(WindowSettings.IsOnScreen(new WindowBounds(0, 0, 400, 150), 1.0, Array.Empty<PixelRect>()));

    private const string DuplicateKeys =
        "{ \"window\": { \"left\": 1, \"top\": 2, \"width\": 300, \"height\": 200 }, \"window\": { \"left\": 3, \"top\": 4, \"width\": 300, \"height\": 200 } }";

    [Fact]
    public void DuplicateKeysMeanNoBoundsAndDoNotThrow()
    {
        Directory.CreateDirectory(_dir);
        File.WriteAllText(SettingsPath, DuplicateKeys);
        Assert.Null(WindowSettings.Load(SettingsPath));
    }

    [Fact]
    public void ADuplicateKeyFileIsReplacedBySave()
    {
        Directory.CreateDirectory(_dir);
        File.WriteAllText(SettingsPath, DuplicateKeys);
        WindowSettings.Save(new WindowBounds(10, 20, 300, 200), SettingsPath);
        Assert.Equal(new WindowBounds(10, 20, 300, 200), WindowSettings.Load(SettingsPath));
    }
}
