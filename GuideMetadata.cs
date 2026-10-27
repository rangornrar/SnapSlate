namespace SnapSlate;

internal sealed class GuideMetadata
{
    public string Title { get; set; } = "SnapSlate";
    public GuideTemplateType TemplateType { get; set; }
    public string Audience { get; set; } = string.Empty;
    public string Author { get; set; } = string.Empty;
    public string DocumentVersion { get; set; } = "1.0";

    public GuideMetadata Clone() => (GuideMetadata)MemberwiseClone();

    public static GuideMetadata FromLegacyDocument(ScreenshotDocument? document) => document is null ? new() : new()
    {
        Title = document.BaseTitle,
        TemplateType = document.TemplateType,
        Audience = document.Audience,
        Author = document.Author,
        DocumentVersion = document.DocumentVersion
    };
}
