using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SnapSlate.Tests;

[TestClass]
public sealed class GuideTextExporterTests
{
    private static GuideExportManifest CreateManifest() => new(
        "<script>alert(1)</script>", "Guide", "Users", "A & B", "1.0",
        DateTimeOffset.UnixEpoch, "output",
        [new GuideExportStep(1, "Step \"quoted\"", "First\r\nSecond <img onerror=x>", "Capture", "<source>",
            "capture.png", [new GuideExportLegendItem("1", "<b>Legend</b>")])]);

    [TestMethod]
    public void BuildHtml_WithUserContent_EncodesMarkupAndAttributes()
    {
        var html = GuideTextExporter.BuildHtml(CreateManifest());
        StringAssert.Contains(html, "&lt;script&gt;");
        StringAssert.Contains(html, "A &amp; B");
        StringAssert.Contains(html, "alt=\"Step &quot;quoted&quot;\"");
        StringAssert.Contains(html, "First<br />Second &lt;img onerror=x&gt;");
        StringAssert.Contains(html, "&lt;b&gt;Legend&lt;/b&gt;");
        Assert.IsFalse(html.Contains("<script>"));
    }

    [TestMethod]
    public void BuildMarkdown_WithoutImages_KeepsNotesAndLegend()
    {
        var markdown = GuideTextExporter.BuildMarkdown(CreateManifest(), includeImageReferences: false);
        Assert.IsFalse(markdown.Contains("![]("));
        StringAssert.Contains(markdown, "First  \nSecond");
        StringAssert.Contains(markdown, "- 1 : <b>Legend</b>");
    }

    [TestMethod]
    public void BuildHtml_WithEnglish_UsesEnglishLanguageAndLabels()
    {
        var html = GuideTextExporter.BuildHtml(CreateManifest(), useEnglish: true);
        StringAssert.Contains(html, "<html lang=\"en\">");
        StringAssert.Contains(html, "Author");
        StringAssert.Contains(html, "Legend");
    }

    [TestMethod]
    public void BuildMarkdown_WithImages_UsesRelativeImagePath()
    {
        var markdown = GuideTextExporter.BuildMarkdown(CreateManifest());
        StringAssert.Contains(markdown, "![](images/capture.png)");
    }
}
