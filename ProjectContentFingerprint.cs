using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;

namespace SnapSlate;

// Image arrays are immutable after import and shared by history snapshots.
internal static class ProjectContentFingerprint
{
    private sealed record ImageHash(string Value);
    private static readonly ConditionalWeakTable<byte[], ImageHash> ImageHashes = new();

    public static void WarmImage(byte[] bytes) => _ = ImageHashes.GetValue(bytes, image => new(Convert.ToHexString(SHA256.HashData(image))));

    public static string Project(SnapSlateProjectState state) => Hash(writer =>
    {
        var metadata = state.Metadata;
        writer.Write(metadata is not null);
        if (metadata is not null)
        {
            writer.Write(metadata.Title); writer.Write((int)metadata.TemplateType); writer.Write(metadata.Audience);
            writer.Write(metadata.Author); writer.Write(metadata.DocumentVersion);
        }
        writer.Write(state.Documents.Count);
        foreach (var document in state.Documents) writer.Write(Document(document));
    });

    public static string Document(ScreenshotDocument document) => Hash(writer =>
    {
        writer.Write(document.Id.ToByteArray()); writer.Write((int)document.Origin);
        writer.Write(document.BaseTitle); writer.Write(document.StepNote); writer.Write(document.OriginLabel);
        writer.Write(document.SourceLabel); writer.Write(document.FileNameLabel);
        writer.Write(document.SourcePixelWidth); writer.Write(document.SourcePixelHeight);
        writer.Write(document.ImageBytes is not null);
        if (document.ImageBytes is { } bytes)
            writer.Write(ImageHashes.GetValue(bytes, image => new(Convert.ToHexString(SHA256.HashData(image)))).Value);
        writer.Write(document.SelectedPaletteKey); writer.Write(document.SelectedPaletteShadeIndex);
        writer.Write(document.PaletteDisplayName); writer.Write(document.AnnotationText); writer.Write(document.StepNumber);
        writer.Write((int)document.TemplateType); writer.Write(document.Audience); writer.Write(document.Author); writer.Write(document.DocumentVersion);
        writer.Write(document.StickerModeIndex); writer.Write(document.StrokeThickness); writer.Write(document.DefaultOpacity);
        writer.Write(document.NextStickerIndex); writer.Write(document.ResetStickerNumberOnColorChange);
        WriteCrop(writer, document.PendingCropRect); WriteCrop(writer, document.AppliedCropRect);
        writer.Write(document.Annotations.Count);
        foreach (var annotation in document.Annotations)
        {
            writer.Write(annotation.Id.ToByteArray()); writer.Write((int)annotation.Kind);
            writer.Write(annotation.PaletteKey); writer.Write(annotation.PaletteShadeIndex);
            writer.Write(annotation.StrokeThickness); writer.Write(annotation.Opacity);
            writer.Write(annotation.StartPoint.X); writer.Write(annotation.StartPoint.Y);
            writer.Write(annotation.EndPoint.X); writer.Write(annotation.EndPoint.Y);
            WriteCrop(writer, annotation.Bounds);
            writer.Write(annotation.Text); writer.Write(annotation.StickerModeIndex); writer.Write(annotation.StickerIndex);
            writer.Write(annotation.LegendText); writer.Write(annotation.FontSize); writer.Write(annotation.IsBold);
        }
    });

    private static void WriteCrop(BinaryWriter writer, Windows.Foundation.Rect? rectangle)
    {
        writer.Write(rectangle.HasValue);
        if (rectangle is not { } bounds) return;
        writer.Write(bounds.X); writer.Write(bounds.Y); writer.Write(bounds.Width); writer.Write(bounds.Height);
    }

    private static string Hash(Action<BinaryWriter> write)
    {
        using var stream = new MemoryStream();
        using (var writer = new BinaryWriter(stream, System.Text.Encoding.UTF8, leaveOpen: true)) write(writer);
        return Convert.ToHexString(SHA256.HashData(stream.GetBuffer().AsSpan(0, (int)stream.Length)));
    }
}
