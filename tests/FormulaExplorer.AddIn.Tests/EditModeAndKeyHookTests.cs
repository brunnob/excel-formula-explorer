using System.Runtime.InteropServices;
using FormulaExplorer.AddIn.Excel;
using FormulaExplorer.AddIn.UI;

namespace FormulaExplorer.AddIn.Tests;

public class EditModeAndKeyHookTests
{
    private static readonly int VbaIgnore = unchecked((int)0x800AC472);

    // Stands in for Excel's Application: only CommandBars.GetEnabledMso matters.
    public sealed class FakeApp(Func<string, bool> enabled)
    {
        public FakeBars CommandBars { get; } = new(enabled);
    }

    public sealed class FakeBars(Func<string, bool> enabled)
    {
        public bool GetEnabledMso(string id) => enabled(id);
    }

    private static FakeApp AllEnabled() => new(_ => true);

    private static FakeApp AllDisabled() => new(_ => false);

    // ----- edit-mode detection -----

    [Fact]
    public void TheEditorWindowDecidesWhenItCouldBeObserved()
    {
        Assert.True(ComErrors.IsEditMode(AllEnabled(), () => true));
        Assert.False(ComErrors.IsEditMode(AllDisabled(), () => false));
    }

    [Fact]
    public void ADisabledNewWorkbookCommandAloneIsNotEditMode()
    {
        // FileNewDefault disabled by policy / another add-in / a template: the old single-command test called this edit mode.
        var policy = new FakeApp(id => id != "FileNewDefault");
        Assert.False(ComErrors.IsEditMode(policy, () => null));
        Assert.False(ComErrors.IsEditMode(policy, () => false));
    }

    [Fact]
    public void WithoutAWindowObservationOnlyAClearMajorityOfTheCommandsCountsAsEditMode()
    {
        Assert.True(ComErrors.IsEditMode(AllDisabled(), () => null));
        Assert.False(ComErrors.IsEditMode(AllEnabled(), () => null));

        // two unrelated commands disabled by policy are still not edit mode
        var someDisabled = new FakeApp(id => id is not ("FileNewDefault" or "FileOpen"));
        Assert.False(ComErrors.IsEditMode(someDisabled, () => null));
    }

    [Fact]
    public void OneOrTwoCommandsDisabledForGoodDoNotHideRealEditModeFromTheFallback()
    {
        // While editing Excel disables all nine commands; one or two that answer "enabled" anyway (a command that cannot be
        // asked properly, a customised command bar) must not turn the fallback into "never in edit mode".
        var oneOddEnabled = new FakeApp(id => id == "ZoomDialog");
        Assert.True(ComErrors.IsEditMode(oneOddEnabled, () => null));          // 8 of 9 disabled
        var twoOddEnabled = new FakeApp(id => id is "ZoomDialog" or "MacroPlay");
        Assert.True(ComErrors.IsEditMode(twoOddEnabled, () => null));          // 7 of 9 disabled
        var threeEnabled = new FakeApp(id => id is "ZoomDialog" or "MacroPlay" or "NameManager");
        Assert.False(ComErrors.IsEditMode(threeEnabled, () => null));         // 6 of 9: not a majority this clear
    }

    [Fact]
    public void WhenNothingCanBeObservedEditModeIsNotAssumed()
    {
        Assert.False(ComErrors.IsEditMode(new object(), () => null));
    }

    [Theory]
    [InlineData(null, 9, 9, true)]
    [InlineData(null, 8, 9, true)]
    [InlineData(null, 7, 9, true)]
    [InlineData(null, 6, 9, false)]
    [InlineData(null, 3, 4, true)]     // 4 answered, 3 quarters disabled
    [InlineData(null, 2, 4, false)]
    [InlineData(null, 0, 0, false)]    // nothing answered
    [InlineData(null, 3, 3, false)]    // too few answered to be a vote
    [InlineData(null, 0, 9, false)]    // none disabled: not editing
    [InlineData(true, 0, 9, true)]     // the window wins over the votes
    [InlineData(false, 9, 9, false)]
    public void DecideCombinesTheSignals(bool? window, int disabled, int probed, bool expected) =>
        Assert.Equal(expected, EditModeSignals.Decide(window, disabled, probed));

    [Fact]
    public void ThrowIfEditModeRaisesTheExplicitEditModeExceptionOnlyWhileEditing()
    {
        var ex = Assert.Throws<EditModeException>(() => ComErrors.ThrowIfEditMode(AllDisabled()));
        Assert.True(ComErrors.IsBusy(ex));   // every busy handler still catches it
        Assert.Equal(VbaIgnore, ex.HResult);
        Assert.True(new ExcelBusyException(ex).IsEditMode);
        ComErrors.ThrowIfEditMode(AllEnabled());
    }

    [Fact]
    public void ABareVbaIgnoreFromExcelIsBusyButNotEditMode()
    {
        var bare = new COMException("not ready", VbaIgnore);
        Assert.True(ComErrors.IsBusy(bare));
        Assert.False(new ExcelBusyException(bare).IsEditMode);   // user commands are retried, not dropped
    }

    [Fact]
    public void OtherBusyReasonsAreNotEditMode()
    {
        Assert.False(new ExcelBusyException(new COMException("x", unchecked((int)0x80010001))).IsEditMode);
    }

    // ----- key hook -----

    [Theory]
    [InlineData(0x0100u, 0x1B, true)]    // WM_KEYDOWN Esc
    [InlineData(0x0100u, 0x0D, true)]    // WM_KEYDOWN Enter
    [InlineData(0x0100u, 0x41, false)]   // other keys are left alone
    [InlineData(0x0101u, 0x1B, false)]   // key up is not taken
    [InlineData(0x0200u, 0x1B, false)]   // not a key message
    public void OnlyEscAndEnterKeyDownsAreTakenFromExcelsPump(uint message, int virtualKey, bool expected)
    {
        Assert.Equal(expected, WindowKeyHook.IsInterceptedKey(message, virtualKey));
    }

    [Theory]
    [InlineData(0x00010001L, false)]   // first press: repeat count 1, previous state 0
    [InlineData(0x40010001L, true)]    // bit 30 set: the key was already down (auto-repeat)
    [InlineData(0xC0010001L, true)]    // bits 30 and 31
    [InlineData(0x20010001L, false)]   // bit 29 (Alt) is not the previous state
    public void AutoRepeatIsReadFromTheKeyDownMessageItself(long lParam, bool expected)
    {
        Assert.Equal(expected, WindowKeyHook.IsAutoRepeat(lParam));
    }
}
