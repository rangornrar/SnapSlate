using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace SnapSlate;

internal static class GuideTextExporter
{
    internal static string EscapeHtml(string value)
    {
        return System.Net.WebUtility.HtmlEncode(value ?? string.Empty);
    }

    private static string NormalizeLineBreaks(string value)
    {
        return (value ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
    }

    private static string FormatMarkdownMultiline(string value)
    {
        return string.Join("  \n", NormalizeLineBreaks(value).Split('\n'));
    }

    private static string FormatHtmlMultiline(string value)
    {
        return EscapeHtml(NormalizeLineBreaks(value)).Replace("\n", "<br />");
    }

    internal static string BuildMarkdown(GuideExportManifest manifest, bool includeImageReferences = true, bool useEnglish = false)
    {
        string T(string french, string english) => useEnglish ? english : french;
        var markdown = new StringBuilder();
        markdown.AppendLine($"# {manifest.Title}");
        markdown.AppendLine();
        markdown.AppendLine($"- {T("Type de document", "Document type")} : {manifest.TemplateName}");
        markdown.AppendLine($"- {T("Public", "Audience")} : {manifest.Audience}");
        markdown.AppendLine($"- {T("Auteur", "Author")} : {manifest.Author}");
        markdown.AppendLine($"- {T("Version", "Version")} : {manifest.Version}");
        markdown.AppendLine();

        foreach (var step in manifest.Steps)
        {
            markdown.AppendLine($"## {step.Index:00}. {step.Title}");
            markdown.AppendLine();

            if (!string.IsNullOrWhiteSpace(step.Note))
            {
                markdown.AppendLine(FormatMarkdownMultiline(step.Note));
                markdown.AppendLine();
            }

            if (includeImageReferences)
            {
                markdown.AppendLine($"![](images/{Path.GetFileName(step.ImagePath)})");
                markdown.AppendLine();
            }

            if (step.LegendItems.Count > 0)
            {
                markdown.AppendLine(T("Légende", "Legend"));
                foreach (var legend in step.LegendItems)
                {
                    markdown.AppendLine($"- {legend.StickerLabel} : {legend.Description}");
                }

                markdown.AppendLine();
            }

            markdown.AppendLine($"_{step.SourceLabel}_");
            markdown.AppendLine();
        }

        return markdown.ToString();
    }

    internal static string BuildHtml(GuideExportManifest manifest, bool useEnglish = false)
    {
        string T(string french, string english) => useEnglish ? english : french;
        var html = new StringBuilder();
        html.AppendLine("<!doctype html>");
        html.AppendLine($"<html lang=\"{(useEnglish ? "en" : "fr")}\">");
        html.AppendLine("<head>");
        html.AppendLine("<meta charset=\"utf-8\">");
        html.AppendLine($"<title>{EscapeHtml(manifest.Title)}</title>");
        html.AppendLine("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">");
        html.AppendLine("<style>body{font-family:Segoe UI,Arial,sans-serif;max-width:1100px;margin:0 auto;padding:32px;color:#1f2937;line-height:1.5}h1,h2{color:#1e3a8a}img{max-width:100%;height:auto;border:1px solid #d1d5db;border-radius:14px}.meta{color:#4b5563}.step{margin:28px 0;padding-bottom:20px;border-bottom:1px solid #e5e7eb}.legend{margin-top:12px;padding:12px 16px;background:#f8fafc;border-radius:12px}.legend ul{margin:8px 0 0 20px}</style>");
        html.AppendLine("</head>");
        html.AppendLine("<body>");
        html.AppendLine($"<h1>{EscapeHtml(manifest.Title)}</h1>");
        html.AppendLine($"<p class=\"meta\">{EscapeHtml(manifest.TemplateName)} · {EscapeHtml(manifest.Audience)} · {EscapeHtml(manifest.Version)}</p>");
        html.AppendLine($"<p class=\"meta\">{EscapeHtml(T("Auteur", "Author"))} : {EscapeHtml(manifest.Author)} · {EscapeHtml(manifest.GeneratedAt.ToString("g", CultureInfo.CurrentCulture))}</p>");

        foreach (var step in manifest.Steps)
        {
            html.AppendLine("<section class=\"step\">");
            html.AppendLine($"<h2>{EscapeHtml($"{step.Index:00}. {step.Title}")}</h2>");
            if (!string.IsNullOrWhiteSpace(step.Note))
            {
                html.AppendLine($"<p>{FormatHtmlMultiline(step.Note)}</p>");
            }

            html.AppendLine($"<img src=\"images/{EscapeHtml(Path.GetFileName(step.ImagePath))}\" alt=\"{EscapeHtml(step.Title)}\">");

            if (step.LegendItems.Count > 0)
            {
                html.AppendLine("<div class=\"legend\">");
                html.AppendLine($"<strong>{EscapeHtml(T("Légende", "Legend"))}</strong>");
                html.AppendLine("<ul>");
                foreach (var legend in step.LegendItems)
                {
                    html.AppendLine($"<li><strong>{EscapeHtml(legend.StickerLabel)}</strong> : {EscapeHtml(legend.Description)}</li>");
                }
                html.AppendLine("</ul>");
                html.AppendLine("</div>");
            }

            html.AppendLine($"<p class=\"meta\">{EscapeHtml(step.SourceLabel)}</p>");
            html.AppendLine("</section>");
        }

        html.AppendLine("</body></html>");
        return html.ToString();
    }

}
