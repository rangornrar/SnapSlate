using System;
using System.Collections.Generic;
using System.IO;
using Windows.Foundation;

namespace SnapSlate;

internal static class ProjectValidator
{
    public const long MaxProjectBytes = 256L * 1024 * 1024;
    private static readonly HashSet<string> PaletteKeys = ["Sunset", "Ember", "Citrus", "Mint", "Lagoon", "Sky", "Ocean", "Rose", "Berry", "Grape"];

    public static void Validate(SnapSlateProjectState state)
    {
        if (state.Version != 1 || state.Documents is null || state.Documents.Count > 200)
        {
            throw new InvalidDataException("Unsupported or invalid project.");
        }

        var ids = new HashSet<Guid>();
        long imageBytes = 0;
        foreach (var document in state.Documents)
        {
            if (document is null || document.Id == Guid.Empty || !ids.Add(document.Id) || document.Annotations is null
                || document.Annotations.Count > 5000 || !Enum.IsDefined(document.Origin) || !Enum.IsDefined(document.TemplateType)
                || document.SelectedPaletteKey is null || !PaletteKeys.Contains(document.SelectedPaletteKey)
                || document.SelectedPaletteShadeIndex is < 0 or > 6 || document.StickerModeIndex is < 0 or > 1
                || document.NextStickerIndex < 1)
            {
                throw new InvalidDataException("Invalid document.");
            }

            ImageSafety.ValidateDimensions(document.SourcePixelWidth, document.SourcePixelHeight);
            ValidateNumber(document.DefaultOpacity, 0, 1);
            ValidateNumber(document.StrokeThickness, 0, 100);
            ValidateStrings(document.BaseTitle, document.StepNote, document.OriginLabel, document.SourceLabel,
                document.FileNameLabel, document.PaletteDisplayName, document.AnnotationText, document.Audience,
                document.Author, document.DocumentVersion);
            ValidateCrop(document.PendingCropRect, document);
            ValidateCrop(document.AppliedCropRect, document);
            if (document.ImageBytes is { } bytes)
            {
                imageBytes += bytes.Length;
                if (imageBytes > 128L * 1024 * 1024)
                {
                    throw new InvalidDataException("Project images exceed the supported limit.");
                }
                ImageSafety.Validate(bytes);
            }

            var annotationIds = new HashSet<Guid>();
            foreach (var annotation in document.Annotations)
            {
                if (annotation is null || annotation.Id == Guid.Empty || !annotationIds.Add(annotation.Id)
                    || !Enum.IsDefined(annotation.Kind) || annotation.PaletteKey is null || !PaletteKeys.Contains(annotation.PaletteKey)
                    || annotation.PaletteShadeIndex is < 0 or > 6)
                {
                    throw new InvalidDataException("Invalid annotation.");
                }
                ValidateNumber(annotation.Opacity, 0, 1);
                ValidateNumber(annotation.StrokeThickness, 0, 100);
                ValidateNumber(annotation.FontSize, 1, 1000);
                ValidatePoint(annotation.StartPoint);
                ValidatePoint(annotation.EndPoint);
                ValidateRect(annotation.Bounds);
                ValidateStrings(annotation.Text, annotation.LegendText);
            }
        }
    }

    private static void ValidateStrings(params string[] values)
    {
        foreach (var value in values)
        {
            if (value is null || value.Length > 100_000)
            {
                throw new InvalidDataException("Invalid project text.");
            }
        }
    }

    private static void ValidateCrop(Rect? crop, ScreenshotDocument document)
    {
        if (crop is not Rect bounds) return;
        ValidateRect(bounds);
        if (bounds.Width < 1 || bounds.Height < 1 || bounds.X < 0 || bounds.Y < 0
            || bounds.Right > document.SourcePixelWidth || bounds.Bottom > document.SourcePixelHeight)
        {
            throw new InvalidDataException("Crop must fit inside the image.");
        }
    }

    private static void ValidateRect(Rect rect)
    {
        ValidatePoint(new Point(rect.X, rect.Y));
        ValidateNumber(rect.Width, 0, 1_000_000);
        ValidateNumber(rect.Height, 0, 1_000_000);
    }

    private static void ValidatePoint(Point point)
    {
        ValidateNumber(point.X, -1_000_000, 1_000_000);
        ValidateNumber(point.Y, -1_000_000, 1_000_000);
    }

    private static void ValidateNumber(double value, double minimum, double maximum)
    {
        if (!double.IsFinite(value) || value < minimum || value > maximum)
        {
            throw new InvalidDataException("Invalid project geometry or style.");
        }
    }
}
