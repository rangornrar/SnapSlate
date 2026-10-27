using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Windows.Foundation;

namespace SnapSlate.Tests;

[TestClass]
public sealed class EditorConsistencyTests
{
    [TestMethod]
    public void Deserialize_LegacyProject_MigratesFirstStepMetadataIndependentlyOfSelection()
    {
        var first = new ScreenshotDocument { BaseTitle = "Guide initial", Author = "A", Audience = "Équipe", DocumentVersion = "3" };
        var second = new ScreenshotDocument { BaseTitle = "Deuxième étape", Author = "B" };
        var legacy = new SnapSlateProjectState { Documents = [first, second], SelectedDocumentId = second.Id };
        var restored = ProjectPersistence.Deserialize(ProjectPersistence.Serialize(legacy))!;
        Assert.AreEqual(2, restored.Version);
        Assert.AreEqual("Guide initial", restored.Metadata!.Title);
        Assert.AreEqual("A", restored.Metadata.Author);
        Assert.AreEqual("Deuxième étape", restored.Documents[1].StepTitle);
    }

    [TestMethod]
    public void Serialize_GuideMetadata_RoundTripsIndependentlyOfStepTitles()
    {
        var state = new SnapSlateProjectState { Version = 2, Metadata = new() { Title = "Guide", Author = "Autrice", TemplateType = GuideTemplateType.ReleaseNotes }, Documents = [new() { BaseTitle = "Étape" }] };
        var restored = ProjectPersistence.Deserialize(ProjectPersistence.Serialize(state))!;
        Assert.AreEqual("Guide", restored.Metadata!.Title);
        Assert.AreEqual(GuideTemplateType.ReleaseNotes, restored.Metadata.TemplateType);
        Assert.AreEqual("Étape", restored.Documents[0].StepTitle);
    }

    [TestMethod]
    public void HasChanges_GuideMetadataEditAndUndo_TracksGlobalChanges()
    {
        var metadata = new GuideMetadata { Title = "Enregistré" };
        var state = new SnapSlateProjectState { Version = 2, Metadata = metadata, Documents = [new()] };
        var saved = metadata.Clone();
        var tracker = new ProjectSaveTracker();
        tracker.Accept(state);
        metadata.Title = "Modifié";
        Assert.IsTrue(tracker.HasChanges(state));
        state.Metadata = saved;
        Assert.IsFalse(tracker.HasChanges(state));
    }

    [TestMethod]
    public void Clone_MetadataDraft_DoesNotChangeLiveGuide()
    {
        var live = new GuideMetadata { Title = "Guide", Audience = "Public", Author = "A", DocumentVersion = "1" };
        var draft = live.Clone();
        draft.Title = "Brouillon"; draft.Audience = "Autre"; draft.Author = "B"; draft.DocumentVersion = "2";
        Assert.AreEqual("Guide", live.Title);
        Assert.AreEqual("Public", live.Audience);
        Assert.AreEqual("A", live.Author);
        Assert.AreEqual("1", live.DocumentVersion);
    }

    [TestMethod]
    [DataRow("{\"version\":2,\"documents\":[]}")]
    [DataRow("{\"version\":2,\"metadata\":{\"title\":null},\"documents\":[]}")]
    [DataRow("{\"version\":2,\"metadata\":{\"templateType\":999},\"documents\":[]}")]
    public void Deserialize_InvalidVersionTwoMetadata_RejectsProject(string json) =>
        Assert.ThrowsExactly<InvalidDataException>(() => ProjectPersistence.Deserialize(json));

    [TestMethod]
    public void RenderPng_PendingCrop_DoesNotApplyUnconfirmedCrop()
    {
        var document = Source();
        document.PendingCropRect = new Rect(10, 10, 20, 20);
        using var bytes = new MemoryStream(GuideImageRenderer.RenderPng(document, Palettes()));
        using var result = new Bitmap(bytes);
        Assert.AreEqual(400, result.Width);
        Assert.AreEqual(240, result.Height);
    }

    [TestMethod]
    public void RenderPng_NewPendingCrop_PreservesPreviouslyAppliedCrop()
    {
        var document = Source();
        document.AppliedCropRect = new Rect(15, 20, 100, 60);
        document.PendingCropRect = new Rect(0, 0, 20, 20);
        using var bytes = new MemoryStream(GuideImageRenderer.RenderPng(document, Palettes()));
        using var result = new Bitmap(bytes);
        Assert.AreEqual(100, result.Width);
        Assert.AreEqual(60, result.Height);
    }

    [TestMethod]
    public void RenderPng_TransparentAnnotations_LeavesSourcePixelsUnchanged()
    {
        var document = Source();
        foreach (var kind in new[] { AnnotationKind.Text, AnnotationKind.Sticker, AnnotationKind.Rectangle, AnnotationKind.Ellipse, AnnotationKind.Focus, AnnotationKind.ArrowStraight, AnnotationKind.ArrowCurved })
            document.Annotations.Add(new() { Kind = kind, Bounds = new Rect(20, 20, 300, 150), StartPoint = new(20, 180), EndPoint = new(360, 180), Text = "Été\n第二行", Opacity = 0 });
        using var bytes = new MemoryStream(GuideImageRenderer.RenderPng(document, Palettes()));
        using var result = new Bitmap(bytes);
        for (var y = 0; y < result.Height; y++)
        for (var x = 0; x < result.Width; x++) Assert.AreEqual(Color.White.ToArgb(), result.GetPixel(x, y).ToArgb(), $"{x},{y}");
    }

    [TestMethod]
    public void RenderPng_PreviewAndDiskExport_UseIdenticalPixelsForEveryAnnotationKind()
    {
        var document = Source();
        foreach (var kind in Enum.GetValues<AnnotationKind>())
            document.Annotations.Add(new() { Kind = kind, Bounds = new Rect(20, 20, 280, 130), StartPoint = new(20, 180), EndPoint = new(360, 180), Text = "Été\nLong text", Opacity = 0.45, IsBold = true });
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".png");
        try
        {
            var preview = GuideImageRenderer.RenderPng(document, Palettes());
            GuideImageRenderer.SaveDocumentImage(document, Palettes(), path);
            CollectionAssert.AreEqual(preview, File.ReadAllBytes(path));
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [TestMethod]
    [DataRow(false)]
    [DataRow(true)]
    public void RenderPng_AppliedCrop_IsExactRegionOfFullPreview(bool demo)
    {
        var document = demo ? new ScreenshotDocument { BaseTitle = "Demo", StepNote = "Long note" } : Source();
        document.Annotations.Add(new() { Kind = AnnotationKind.Text, Bounds = new Rect(30, 30, 230, 80), Text = "Été\nLong text", Opacity = 0.1, IsBold = true });
        document.Annotations.Add(new() { Kind = AnnotationKind.Rectangle, Bounds = new Rect(0, 0, 200, 150), Opacity = 0.4 });
        using var fullBytes = new MemoryStream(GuideImageRenderer.RenderPng(document, Palettes()));
        using var full = new Bitmap(fullBytes);
        document.AppliedCropRect = new Rect(10, 15, 180, 150);
        using var cropBytes = new MemoryStream(GuideImageRenderer.RenderPng(document, Palettes()));
        using var cropped = new Bitmap(cropBytes);
        for (var y = 0; y < 150; y++)
        for (var x = 0; x < 180; x++) Assert.AreEqual(full.GetPixel(x + 10, y + 15).ToArgb(), cropped.GetPixel(x, y).ToArgb());
    }

    [TestMethod]
    public void Clamp_FractionalCropOutsideSource_AlignsToCoveredPixels()
    {
        Assert.AreEqual(new Rect(0, 10, 100, 70), CropGeometry.Clamp(new Rect(-4, 10.2, 110, 100), 100, 80));
        Assert.AreEqual(new Rect(99, 79, 1, 1), CropGeometry.Clamp(new Rect(200, 200, 10, 10), 100, 80));
    }

    private static ScreenshotDocument Source()
    {
        using var bitmap = new Bitmap(400, 240);
        using (var graphics = Graphics.FromImage(bitmap)) graphics.Clear(Color.White);
        using var stream = new MemoryStream();
        bitmap.Save(stream, ImageFormat.Png);
        return new() { SourcePixelWidth = 400, SourcePixelHeight = 240, ImageBytes = stream.ToArray() };
    }

    private static Dictionary<string, GradientPaletteDefinition> Palettes()
    {
        var start = Windows.UI.Color.FromArgb(255, 230, 20, 40);
        var end = Windows.UI.Color.FromArgb(255, 255, 180, 20);
        return new() { ["Sunset"] = new("Sunset", "Test", start, end) { Shades = [start] } };
    }
}
