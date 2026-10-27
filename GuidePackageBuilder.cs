using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SnapSlate;

internal static class GuidePackageBuilder
{
    public static async Task<GuidePackageResult> BuildAsync(string rootFolder,
        IEnumerable<ScreenshotDocument> sourceDocuments, GuideMetadata sourceMetadata,
        IReadOnlyDictionary<string, GradientPaletteDefinition> palettes, ExportFormatChoice[] requestedFormats,
        string templateName, bool useEnglish, Action<int, int>? progress, CancellationToken cancellationToken)
    {
        // Clone everything before the first await, including metadata and document order.
        var documents = sourceDocuments.Select(document => document.CloneForHistory()).ToArray();
        var metadata = sourceMetadata.Clone();
        var formats = requestedFormats.Length > 0 ? requestedFormats.Distinct().ToArray()
            : [ExportFormatChoice.Markdown, ExportFormatChoice.Html, ExportFormatChoice.Pdf, ExportFormatChoice.Docx];
        var packageFolder = Path.Combine(rootFolder, $"SnapSlate-{DateTime.Now:yyyyMMdd-HHmmss}-{Guid.NewGuid():N}");
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            var imagesFolder = Path.Combine(packageFolder, "images");
            Directory.CreateDirectory(imagesFolder);
            var steps = new List<GuideExportStep>(documents.Length);
            for (var index = 0; index < documents.Length; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                progress?.Invoke(index + 1, documents.Length);
                var document = documents[index];
                var imagePath = Path.Combine(imagesFolder, $"step-{index + 1:00}.png");
                await Task.Run(() => GuideImageRenderer.SaveDocumentImage(document, palettes, imagePath), cancellationToken);
                cancellationToken.ThrowIfCancellationRequested();
                steps.Add(new(index + 1, document.StepTitle, document.StepNote, document.OriginLabel,
                    document.SourceLabel, imagePath, document.Annotations.Where(annotation => annotation.Kind == AnnotationKind.Sticker)
                    .Select(annotation => new GuideExportLegendItem(annotation.Text,
                        string.IsNullOrWhiteSpace(annotation.LegendText) ? (useEnglish ? "To fill in" : "À compléter") : annotation.LegendText)).ToArray()));
            }
            var manifest = new GuideExportManifest(metadata.Title, templateName, metadata.Audience, metadata.Author,
                metadata.DocumentVersion, DateTimeOffset.Now, packageFolder, steps);
            var generatedFormats = new List<string>();
            foreach (var format in formats)
            {
                cancellationToken.ThrowIfCancellationRequested();
                switch (format)
                {
                    case ExportFormatChoice.Png: generatedFormats.Add("PNG"); break;
                    case ExportFormatChoice.Markdown:
                        await File.WriteAllTextAsync(Path.Combine(packageFolder, "guide.md"), GuideTextExporter.BuildMarkdown(manifest, useEnglish: useEnglish), Encoding.UTF8, cancellationToken);
                        generatedFormats.Add("Markdown"); break;
                    case ExportFormatChoice.Html:
                        await File.WriteAllTextAsync(Path.Combine(packageFolder, "guide.html"), GuideTextExporter.BuildHtml(manifest, useEnglish), Encoding.UTF8, cancellationToken);
                        generatedFormats.Add("HTML"); break;
                    case ExportFormatChoice.Pdf:
                        await GuideExporter.ExportPdfAsync(manifest, Path.Combine(packageFolder, "guide.pdf"));
                        generatedFormats.Add("PDF"); break;
                    case ExportFormatChoice.Docx:
                        await GuideExporter.ExportDocxAsync(manifest, Path.Combine(packageFolder, "guide.docx"));
                        generatedFormats.Add("DOCX"); break;
                }
            }
            cancellationToken.ThrowIfCancellationRequested();
            return new(packageFolder, manifest, generatedFormats);
        }
        catch
        {
            if (Directory.Exists(packageFolder)) Directory.Delete(packageFolder, recursive: true);
            throw;
        }
    }
}
