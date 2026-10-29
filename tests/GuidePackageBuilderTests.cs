using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SnapSlate.Tests;

[TestClass]
public sealed class GuidePackageBuilderTests
{
    private string _root = null!;
    [TestInitialize]
    public void Initialize() => _root = Path.Combine(Path.GetTempPath(), "SnapSlate-package-tests", Guid.NewGuid().ToString("N"));
    [TestCleanup]
    public void Cleanup() { if (Directory.Exists(_root)) Directory.Delete(_root, true); }

    [TestMethod]
    public async Task BuildAsync_CancelAfterFirstImage_RemovesOnlyPartialPackage()
    {
        Directory.CreateDirectory(_root);
        var existing = Path.Combine(_root, "existing.txt");
        File.WriteAllText(existing, "keep");
        using var cancellation = new CancellationTokenSource();
        await Assert.ThrowsAsync<OperationCanceledException>(async () =>
            await GuidePackageBuilder.BuildAsync(_root, [Source(), Source()], new(), Palettes(), [ExportFormatChoice.Html], "Manual", true,
                (step, _) => { if (step == 2) cancellation.Cancel(); }, cancellation.Token));
        Assert.AreEqual(0, Directory.GetDirectories(_root).Length);
        Assert.AreEqual("keep", File.ReadAllText(existing));
    }

    [TestMethod]
    public async Task BuildAsync_EditWhileExporting_PreservesMetadataOrderAndImageSnapshot()
    {
        var first = Source(); first.StepTitle = "First";
        var second = Source(); second.StepTitle = "Second";
        var documents = new List<ScreenshotDocument> { first, second };
        var metadata = new GuideMetadata { Title = "Guide before", Author = "Author before" };
        var result = await GuidePackageBuilder.BuildAsync(_root, documents, metadata, Palettes(), [ExportFormatChoice.Html], "Manual", true,
            (step, _) =>
            {
                if (step != 1) return;
                metadata.Title = "Changed"; metadata.Author = "Changed";
                first.StepTitle = "Changed"; first.ImageBytes = [1, 2, 3]; documents.Reverse();
            }, CancellationToken.None);
        Assert.AreEqual("Guide before", result.Manifest.Title);
        Assert.AreEqual("Author before", result.Manifest.Author);
        CollectionAssert.AreEqual(new[] { "First", "Second" }, result.Manifest.Steps.Select(step => step.Title).ToArray());
        StringAssert.Contains(File.ReadAllText(Path.Combine(result.PackageFolder, "guide.html")), "Guide before");
    }

    [TestMethod]
    public async Task BuildAsync_AllFormats_ProducesValidDocxPdfAndSharedImages()
    {
        var result = await GuidePackageBuilder.BuildAsync(_root, [Source()], new() { Title = "Guide test" }, Palettes(), [], "Manual", true, null, CancellationToken.None);
        Assert.AreEqual(4, result.GeneratedFormats.Count);
        foreach (var name in new[] { "guide.md", "guide.html", "guide.pdf", "guide.docx", "images/step-01.png" })
            Assert.IsTrue(new FileInfo(Path.Combine(result.PackageFolder, name)).Length > 0, name);
        using var word = WordprocessingDocument.Open(Path.Combine(result.PackageFolder, "guide.docx"), false);
        Assert.AreEqual(0, new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(word).Count());
        var pdf = File.ReadAllBytes(Path.Combine(result.PackageFolder, "guide.pdf"));
        CollectionAssert.AreEqual(new byte[] { 37, 80, 68, 70 }, pdf.Take(4).ToArray());
        var image = word.MainDocumentPart!.ImageParts.Single();
        using var stream = image.GetStream();
        using var embedded = new MemoryStream(); stream.CopyTo(embedded);
        CollectionAssert.AreEqual(File.ReadAllBytes(result.Manifest.Steps[0].ImagePath), embedded.ToArray());
    }

    [TestMethod]
    public async Task BuildAsync_ConcurrentExports_UseDistinctPackages()
    {
        Task<GuidePackageResult> Export() => GuidePackageBuilder.BuildAsync(_root, [Source()], new(), Palettes(), [ExportFormatChoice.Markdown], "Manual", true, null, CancellationToken.None);
        var results = await Task.WhenAll(Export(), Export());
        Assert.AreNotEqual(results[0].PackageFolder, results[1].PackageFolder);
        Assert.IsTrue(results.All(result => File.Exists(Path.Combine(result.PackageFolder, "guide.md"))));
    }

    [TestMethod]
    public async Task BuildAsync_InvalidImage_RemovesFailedPackage()
    {
        var document = Source(); document.ImageBytes = [1, 2, 3];
        await Assert.ThrowsExactlyAsync<ArgumentException>(async () =>
            await GuidePackageBuilder.BuildAsync(_root, [document], new(), Palettes(), [ExportFormatChoice.Html], "Manual", true, null, CancellationToken.None));
        Assert.AreEqual(0, Directory.GetDirectories(_root).Length);
    }

    [TestMethod]
    public async Task BuildAsync_ActiveCaptureScope_ExportsOnlySelectedAnnotatedImage()
    {
        var active = Source();
        active.StepTitle = "Active screenshot";
        active.Annotations.Add(new() { Kind = AnnotationKind.Mask, Bounds = new Windows.Foundation.Rect(10, 10, 20, 20), Opacity = 0.1 });
        var activePackage = await GuidePackageBuilder.BuildAsync(_root, [active], new(), Palettes(), [ExportFormatChoice.Png], "Manual", true, null, CancellationToken.None);
        Assert.AreEqual(1, activePackage.Manifest.Steps.Count);
        Assert.AreEqual("Active screenshot", activePackage.Manifest.Steps[0].Title);
        Assert.AreEqual(1, Directory.GetFiles(Path.Combine(activePackage.PackageFolder, "images"), "*.png").Length);
        Assert.IsFalse(File.Exists(Path.Combine(activePackage.PackageFolder, "guide.html")));
        using var image = new Bitmap(activePackage.Manifest.Steps[0].ImagePath);
        Assert.AreEqual(Color.FromArgb(255, 18, 18, 18).ToArgb(), image.GetPixel(15, 15).ToArgb());
        var completePackage = await GuidePackageBuilder.BuildAsync(_root, [Source(), active, Source()], new(), Palettes(), [ExportFormatChoice.Png], "Manual", true, null, CancellationToken.None);
        Assert.AreEqual(3, completePackage.Manifest.Steps.Count);
        Assert.AreEqual(3, Directory.GetFiles(Path.Combine(completePackage.PackageFolder, "images"), "*.png").Length);
        Assert.AreEqual("Active screenshot", completePackage.Manifest.Steps[1].Title);
    }

    private static ScreenshotDocument Source()
    {
        using var bitmap = new Bitmap(100, 100);
        using (var graphics = Graphics.FromImage(bitmap)) graphics.Clear(Color.White);
        using var stream = new MemoryStream(); bitmap.Save(stream, ImageFormat.Png);
        return new() { SourcePixelWidth = 100, SourcePixelHeight = 100, ImageBytes = stream.ToArray() };
    }

    private static Dictionary<string, GradientPaletteDefinition> Palettes()
    {
        var color = Windows.UI.Color.FromArgb(255, 255, 0, 0);
        return new() { ["Sunset"] = new("Sunset", "Test", color, color) { Shades = [color] } };
    }
}
