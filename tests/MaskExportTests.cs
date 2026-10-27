using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Windows.Foundation;

namespace SnapSlate.Tests;

[TestClass]
public sealed class MaskExportTests
{
    private string _folder = null!;

    [TestInitialize]
    public void Initialize()
    {
        _folder = Path.Combine(Path.GetTempPath(), "SnapSlate-mask-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_folder);
    }

    [TestCleanup]
    public void Cleanup() => Directory.Delete(_folder, true);

    [TestMethod]
    [DataRow(0.0)]
    [DataRow(0.1)]
    [DataRow(0.5)]
    [DataRow(1.0)]
    public void SaveDocumentImage_MaskAtAnyOpacity_CoversAllPixelsIncludingCorners(double opacity)
    {
        using var source = new Bitmap(80, 60);
        using (var graphics = Graphics.FromImage(source)) graphics.Clear(Color.White);
        using var bytes = new MemoryStream();
        source.Save(bytes, ImageFormat.Png);
        var document = new ScreenshotDocument { SourcePixelWidth = 80, SourcePixelHeight = 60, ImageBytes = bytes.ToArray() };
        document.Annotations.Add(new AnnotationModel { Kind = AnnotationKind.Mask, Bounds = new Rect(10.2, 12.1, 30, 20), Opacity = opacity });
        // Even annotations later in the layer list must not reveal or cover a redaction.
        document.Annotations.Add(new AnnotationModel { Kind = AnnotationKind.Rectangle, Bounds = new Rect(12, 14, 25, 16), Opacity = 1 });
        var path = Path.Combine(_folder, "export.png");

        GuideImageRenderer.SaveDocumentImage(document, Palettes(), path);

        using var result = new Bitmap(path);
        for (var y = 12; y < 33; y++)
        for (var x = 10; x < 41; x++)
            Assert.AreEqual(Color.FromArgb(255, 18, 18, 18).ToArgb(), result.GetPixel(x, y).ToArgb(), $"Pixel {x},{y}");
        Assert.AreEqual(Color.White.ToArgb(), result.GetPixel(5, 5).ToArgb());
    }

    [TestMethod]
    public void SaveDocumentImage_CorruptSource_PreservesExistingExport()
    {
        var path = Path.Combine(_folder, "existing.png");
        File.WriteAllText(path, "previous export");
        var document = new ScreenshotDocument { ImageBytes = [1, 2, 3] };
        Assert.ThrowsExactly<ArgumentException>(() => GuideImageRenderer.SaveDocumentImage(document, Palettes(), path));
        Assert.AreEqual("previous export", File.ReadAllText(path));
    }

    private static Dictionary<string, GradientPaletteDefinition> Palettes()
    {
        var color = Windows.UI.Color.FromArgb(255, 255, 0, 0);
        return new Dictionary<string, GradientPaletteDefinition> { ["Sunset"] = new("Sunset", "Test", color, color) { Shades = [color] } };
    }
}
