using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SnapSlate;

internal static class PublicationTextBuilder
{
    public static GuideExportManifest Capture(IEnumerable<ScreenshotDocument> documents, GuideMetadata metadata, string templateName, string language)
    {
        var snapshot = metadata.Clone();
        var steps = documents.Select((document, index) => new GuideExportStep(index + 1, document.StepTitle,
            document.StepNote, document.OriginLabel, ExportSource.ForDocument(document, language), string.Empty,
            document.Annotations.Where(annotation => annotation.Kind == AnnotationKind.Sticker)
                .Select(annotation => new GuideExportLegendItem(annotation.Text, string.IsNullOrWhiteSpace(annotation.LegendText)
                    ? LocalizedText.Get("LegendPlaceholder", language) : annotation.LegendText)).ToArray())).ToArray();
        return new(snapshot.Title, templateName, snapshot.Audience, snapshot.Author, snapshot.DocumentVersion, DateTimeOffset.Now, string.Empty, steps, language);
    }

    public static string Markdown(GuideExportManifest manifest) => LocalizedText.Get("PublicationTextOnlyNotice", manifest.Language)
        + Environment.NewLine + Environment.NewLine + GuideTextExporter.BuildMarkdown(manifest, includeImageReferences: false, useEnglish: manifest.Language == "en-US");

    public static string ConfluenceStorage(GuideExportManifest manifest)
    {
        static string E(string text) => GuideTextExporter.EscapeHtml(text);
        string L(string key) => LocalizedText.Get(key, manifest.Language);
        var storage = new StringBuilder();
        storage.AppendLine($"<p>{E(L("PublicationTextOnlyNotice"))}</p><h1>{E(manifest.Title)}</h1>");
        storage.AppendLine($"<p>{E(manifest.TemplateName)} · {E(manifest.Audience)} · {E(manifest.Version)}</p>");
        storage.AppendLine($"<p>{E(L("ExportAuthor"))}: {E(manifest.Author)}</p>");
        foreach (var step in manifest.Steps)
        {
            storage.AppendLine($"<h2>{step.Index:00}. {E(step.Title)}</h2><p>{E(step.Note).Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", "<br />")}</p>");
            storage.AppendLine($"<p>{E(step.SourceLabel)}</p>");
            if (step.LegendItems.Count == 0) continue;
            storage.AppendLine($"<p>{E(L("ExportLegend"))}</p><ul>");
            foreach (var legend in step.LegendItems) storage.AppendLine($"<li>{E(legend.StickerLabel)}: {E(legend.Description)}</li>");
            storage.AppendLine("</ul>");
        }
        return storage.ToString();
    }
}
