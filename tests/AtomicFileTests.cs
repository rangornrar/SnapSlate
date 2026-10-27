using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SnapSlate.Tests;

[TestClass]
public sealed class AtomicFileTests
{
    private string _folder = null!;
    private string _path = null!;

    [TestInitialize]
    public void Initialize()
    {
        _folder = Path.Combine(Path.GetTempPath(), "SnapSlate-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_folder);
        _path = Path.Combine(_folder, "project.snapslate");
    }

    [TestCleanup]
    public void Cleanup() => Directory.Delete(_folder, true);

    [TestMethod]
    public async Task WriteTextAsync_ExistingProject_PreservesPreviousVersionAsBackup()
    {
        File.WriteAllText(_path, "previous");
        await AtomicFile.WriteTextAsync(_path, "next");
        Assert.AreEqual("next", File.ReadAllText(_path));
        Assert.AreEqual("previous", File.ReadAllText(_path + ".bak"));
    }

    [TestMethod]
    public async Task WriteTextAsync_NewProject_CreatesCompleteFile()
    {
        await AtomicFile.WriteTextAsync(_path, "écriture complète");
        Assert.AreEqual("écriture complète", File.ReadAllText(_path));
        Assert.HasCount(1, Directory.GetFiles(_folder));
    }

    [TestMethod]
    public void Write_EncoderFailsAfterPartialOutput_PreservesOriginalAndRemovesTemporaryFile()
    {
        File.WriteAllText(_path, "original");
        Assert.ThrowsExactly<InvalidDataException>(() => AtomicFile.Write(_path, temporary =>
        {
            File.WriteAllText(temporary, "partial");
            throw new InvalidDataException("Encoder failed.");
        }));
        Assert.AreEqual("original", File.ReadAllText(_path));
        Assert.HasCount(1, Directory.GetFiles(_folder));
    }

    [TestMethod]
    public void Write_RedactedImageExport_DoesNotLeaveUnredactedBackup()
    {
        File.WriteAllText(_path, "unredacted image");
        AtomicFile.Write(_path, temporary => File.WriteAllText(temporary, "redacted image"), keepBackup: false);
        Assert.AreEqual("redacted image", File.ReadAllText(_path));
        Assert.IsFalse(File.Exists(_path + ".bak"));
    }

    [TestMethod]
    public async Task WriteTextAsync_DestinationLocked_PreservesOriginal()
    {
        File.WriteAllText(_path, "original");
        using (File.Open(_path, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            await Assert.ThrowsExactlyAsync<IOException>(() => AtomicFile.WriteTextAsync(_path, "next"));
        }
        Assert.AreEqual("original", File.ReadAllText(_path));
        Assert.HasCount(1, Directory.GetFiles(_folder));
    }
}
