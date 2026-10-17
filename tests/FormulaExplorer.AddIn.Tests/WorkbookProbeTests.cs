using System.Runtime.InteropServices;
using FormulaExplorer.AddIn.Excel;

namespace FormulaExplorer.AddIn.Tests;

public class WorkbookProbeTests
{
    private static COMException Com(int hresult) => new("x", hresult);

    private static readonly int BadIndex = unchecked((int)0x8002000B);
    private static readonly int Generic = unchecked((int)0x800A03EC);
    private static readonly int CallRejected = unchecked((int)0x80010001);

    private static void Found() { }

    private static void Throws(Exception ex) => throw ex;

    private static IReadOnlyList<string> Names(params string[] names) => names;

    [Fact]
    public void AWorkbookFoundByNameIsOpenWithoutWalking()
    {
        var presence = WorkbookProbe.Probe("A.xlsx", Found, () => throw new InvalidOperationException("must not walk"));
        Assert.Equal(WorkbookPresence.Open, presence);
    }

    [Fact]
    public void ARejectedCallIsBusyWithoutWalking()
    {
        var presence = WorkbookProbe.Probe("A.xlsx", () => Throws(Com(CallRejected)), () => throw new InvalidOperationException("must not walk"));
        Assert.Equal(WorkbookPresence.Busy, presence);
    }

    [Fact]
    public void ABadIndexIsGoneOnlyWhenTheWalkDoesNotFindTheWorkbook()
    {
        Assert.Equal(WorkbookPresence.Gone, WorkbookProbe.Probe("A.xlsx", () => Throws(Com(BadIndex)), () => Names("B.xlsx", "C.xlsx")));
        Assert.Equal(WorkbookPresence.Gone, WorkbookProbe.Probe("A.xlsx", () => Throws(Com(BadIndex)), () => Names()));
    }

    [Theory]
    [InlineData(0x800A03EC)]   // generic failure while Excel is between dialogs
    [InlineData(0x8002000B)]   // DISP_E_BADINDEX of a lookup that failed spuriously
    public void ATransientLookupFailureWhileTheWorkbookIsStillListedIsOpenNotGone(uint hresult)
    {
        var presence = WorkbookProbe.Probe("A.xlsx", () => Throws(Com(unchecked((int)hresult))), () => Names("b.XLSX", "a.xlsx"));
        Assert.Equal(WorkbookPresence.Open, presence);
    }

    [Fact]
    public void AnUnexpectedLookupExceptionTypeIsAlsoConfirmedByTheWalk()
    {
        Assert.Equal(WorkbookPresence.Open, WorkbookProbe.Probe("A.xlsx", () => Throws(new ArgumentException("range")), () => Names("A.xlsx")));
        Assert.Equal(WorkbookPresence.Open, WorkbookProbe.Probe("A.xlsx", () => Throws(new InvalidCastException()), () => Names("A.xlsx")));
        Assert.Equal(WorkbookPresence.Gone, WorkbookProbe.Probe("A.xlsx", () => Throws(new ArgumentException("range")), () => Names("B.xlsx")));
    }

    [Fact]
    public void AWalkThatCannotCompleteProvesNothingAndIsBusy()
    {
        Assert.Equal(WorkbookPresence.Busy,
            WorkbookProbe.Probe("A.xlsx", () => Throws(Com(Generic)), () => throw Com(CallRejected)));
        Assert.Equal(WorkbookPresence.Busy,
            WorkbookProbe.Probe("A.xlsx", () => Throws(Com(Generic)), () => throw Com(Generic)));
        Assert.Equal(WorkbookPresence.Busy,
            WorkbookProbe.Probe("A.xlsx", () => Throws(Com(BadIndex)), () => throw new ArgumentException("range")));
    }
}
