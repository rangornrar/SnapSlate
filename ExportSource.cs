using System;

namespace SnapSlate;

internal static class ExportSource
{
    public static string ForDocument(ScreenshotDocument document, string language)
    {
        if (document.Origin == DocumentOrigin.ClipboardCapture) return LocalizedText.Get("SourceClipboard", language);
        if (document.Origin == DocumentOrigin.FileImport) return FileName(document.FileNameLabel, document.SourceLabel);
        // Old projects may also contain a full source path in a non-file document.
        var source = document.SourceLabel;
        return source.Contains('\\') || source.Contains('/') || source.Contains(':') ? FileName(source, document.FileNameLabel) : source;
    }

    private static string FileName(string value, string fallback)
    {
        var name = string.IsNullOrWhiteSpace(value) ? fallback : value;
        name = name.TrimEnd('/', '\\');
        var slash = name.LastIndexOfAny(['/', '\\']);
        if (slash >= 0) name = name[(slash + 1)..];
        if (name.Contains(':')) name = "capture.png";
        return string.IsNullOrWhiteSpace(name) ? "capture.png" : name;
    }
}
