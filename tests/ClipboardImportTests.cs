using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SnapSlate.Tests;

[TestClass]
public sealed class ClipboardImportTests
{
    [TestMethod]
    public async Task ImportAsync_UnchangedClipboard_DoesNotReadOrDecode()
    {
        var coordinator = new ClipboardImportCoordinator();
        var reads = 0;
        var result = await coordinator.ImportAsync(false, () => { reads++; return Task.FromResult<byte[]?>([1]); }, Prepare, _ => Assert.Fail(), () => true);
        Assert.AreEqual(0, result); Assert.AreEqual(0, reads);
    }

    [TestMethod]
    public async Task ImportAsync_TwoIdenticalCopyEvents_PreservesBothUserCaptures()
    {
        var coordinator = new ClipboardImportCoordinator();
        var documents = new List<ScreenshotDocument>();
        for (var index = 0; index < 2; index++)
        {
            coordinator.NotifyChanged();
            Assert.AreEqual(1, await coordinator.ImportAsync(false, () => Task.FromResult<byte[]?>([1, 2]), Prepare, documents.Add, () => true));
        }
        Assert.AreEqual(2, documents.Count);
        Assert.AreNotEqual(documents[0].Id, documents[1].Id);
    }

    [TestMethod]
    public async Task ImportAsync_ManualRepeat_ImportsUnchangedImageAgain()
    {
        var coordinator = new ClipboardImportCoordinator();
        for (var index = 0; index < 2; index++)
            Assert.AreEqual(1, await coordinator.ImportAsync(true, () => Task.FromResult<byte[]?>([1]), Prepare, _ => { }, () => true));
    }

    [TestMethod]
    public async Task ImportAsync_ChangesDuringRead_DiscardsStaleImageAndReadsLatest()
    {
        var coordinator = new ClipboardImportCoordinator();
        coordinator.NotifyChanged();
        var reads = 0;
        var documents = new List<ScreenshotDocument>();
        await coordinator.ImportAsync(false, () =>
        {
            if (++reads == 1) coordinator.NotifyChanged();
            return Task.FromResult<byte[]?>([(byte)reads]);
        }, Prepare, documents.Add, () => true);
        Assert.AreEqual(2, reads); Assert.AreEqual(1, documents.Count);
        Assert.AreEqual((byte)2, documents[0].ImageBytes![0]);
    }

    [TestMethod]
    public async Task ImportAsync_ProjectReplacedDuringPreparation_DoesNotInsertIntoNewProject()
    {
        var coordinator = new ClipboardImportCoordinator(); coordinator.NotifyChanged();
        var inserted = false;
        Assert.AreEqual(0, await coordinator.ImportAsync(false, () => Task.FromResult<byte[]?>([1]), Prepare, _ => inserted = true, () => false));
        Assert.IsFalse(inserted);
    }

    [TestMethod]
    public async Task ImportAsync_ConcurrentNotification_CoalescesInFlightWork()
    {
        var coordinator = new ClipboardImportCoordinator(); coordinator.NotifyChanged();
        var reading = new TaskCompletionSource<byte[]?>();
        var first = coordinator.ImportAsync(false, () => reading.Task, Prepare, _ => { }, () => true);
        Assert.AreEqual(0, await coordinator.ImportAsync(false, () => throw new System.InvalidOperationException(), Prepare, _ => { }, () => true));
        reading.SetResult([1]); Assert.AreEqual(1, await first);
    }

    private static Task<ScreenshotDocument> Prepare(byte[] bytes) => Task.FromResult(new ScreenshotDocument { ImageBytes = bytes });
}
