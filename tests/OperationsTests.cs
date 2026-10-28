using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Windows.Foundation;

namespace SnapSlate.Tests;

[TestClass]
public sealed class OperationsTests
{
    [TestMethod]
    public void ProjectFingerprint_ImageReplacementAndNestedAnnotationEdit_DetectsChangesWithoutJsonImages()
    {
        var document = Source(); var state = new SnapSlateProjectState { Metadata = new(), Documents = [document] };
        var saved = ProjectContentFingerprint.Project(state);
        document.Annotations.Add(new() { Kind = AnnotationKind.Text, Text = "First", Bounds = new Rect(0, 0, 220, 60) });
        Assert.AreNotEqual(saved, ProjectContentFingerprint.Project(state));
        var annotated = ProjectContentFingerprint.Project(state);
        document.Annotations[0].Text = "Edited";
        Assert.AreNotEqual(annotated, ProjectContentFingerprint.Project(state));
        var edited = ProjectContentFingerprint.Project(state);
        document.ImageBytes = [1, 2, 3];
        Assert.AreNotEqual(edited, ProjectContentFingerprint.Project(state));
    }

    [TestMethod]
    [DataRow(AnnotationKind.Rectangle)]
    [DataRow(AnnotationKind.Text)]
    [DataRow(AnnotationKind.Mask)]
    [DataRow(AnnotationKind.ArrowStraight)]
    [DataRow(AnnotationKind.ArrowCurved)]
    public void Translate_KeyboardMoveAndReverse_RestoresGeometry(AnnotationKind kind)
    {
        var annotation = new AnnotationModel { Kind = kind, Bounds = new Rect(10, 20, 100, 60), StartPoint = new(10, 20), EndPoint = new(110, 80) };
        var before = annotation.Clone(); AnnotationMovement.Translate(annotation, 1, -1); AnnotationMovement.Translate(annotation, -1, 1);
        Assert.AreEqual(before.Bounds, annotation.Bounds); Assert.AreEqual(before.StartPoint, annotation.StartPoint); Assert.AreEqual(before.EndPoint, annotation.EndPoint);
    }

    [TestMethod]
    public void Resources_BothLanguages_HaveMatchingKeysAndFormatArguments()
    {
        Dictionary<string, string> Read(string language)
        {
            using var stream = typeof(LocalizedText).Assembly.GetManifestResourceStream($"SnapSlate.Strings.{language}.Resources.resw")!;
            return XDocument.Load(stream).Root!.Elements("data").ToDictionary(x => x.Attribute("name")!.Value, x => x.Element("value")!.Value);
        }
        var french = Read("fr-FR"); var english = Read("en-US");
        CollectionAssert.AreEquivalent(french.Keys.ToArray(), english.Keys.ToArray());
        foreach (var key in french.Keys)
            CollectionAssert.AreEquivalent(Regex.Matches(french[key], @"\{\d+").Select(match => match.Value).ToArray(), Regex.Matches(english[key], @"\{\d+").Select(match => match.Value).ToArray(), key);
    }

    [TestMethod]
    public void ExportSource_WindowsAndUncPaths_KeepOnlyFileName()
    {
        foreach (var path in new[] { @"C:\Users\Private\Client\shot.png", @"\\internal-server\secret\shot.png", "/home/private/client/shot.png" })
        {
            var document = new ScreenshotDocument { Origin = DocumentOrigin.FileImport, FileNameLabel = path, SourceLabel = path };
            Assert.AreEqual("shot.png", ExportSource.ForDocument(document, "en-US"));
        }
    }

    [TestMethod]
    public void PublicationCapture_TextOnly_RequiresNoImageDecodeAndOmitsPaths()
    {
        var document = new ScreenshotDocument { BaseTitle = "Step <safe>", ImageBytes = [1, 2, 3], Origin = DocumentOrigin.FileImport, FileNameLabel = "shot.png", SourceLabel = @"C:\Private\shot.png" };
        var manifest = PublicationTextBuilder.Capture([document], new() { Title = "Guide" }, "Manual", "en-US");
        var markdown = PublicationTextBuilder.Markdown(manifest); var storage = PublicationTextBuilder.ConfluenceStorage(manifest);
        StringAssert.Contains(markdown, "Text-only"); Assert.IsFalse(markdown.Contains("![]"));
        Assert.IsFalse(storage.Contains("Private")); StringAssert.Contains(storage, "&lt;safe&gt;");
    }

    [TestMethod]
    [DataRow("fr-FR", "Auteur", "Étapes")]
    [DataRow("en-US", "Author", "Steps")]
    public async Task ExportPipeline_ImportAnnotateSaveReopenExport_LocalizesAndPreservesRedaction(string language, string author, string steps)
    {
        var root = Path.Combine(Path.GetTempPath(), "SnapSlate-workflow-" + Guid.NewGuid().ToString("N"));
        var previousCulture = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("ja-JP");
            Directory.CreateDirectory(root);
            var document = Source();
            document.Annotations.Add(new() { Kind = AnnotationKind.Mask, Bounds = new Rect(10, 10, 20, 20), Opacity = 0.1 });
            var state = new SnapSlateProjectState { Version = 2, Metadata = new() { Title = "Guide", Author = "Author A" }, Documents = [document] };
            var path = Path.Combine(root, "project.snapslate");
            await AtomicFile.WriteTextAsync(path, ProjectPersistence.Serialize(state));
            var reopened = ProjectPersistence.Deserialize(await File.ReadAllTextAsync(path))!;
            var package = await GuidePackageBuilder.BuildAsync(root, reopened.Documents, reopened.Metadata!, Palettes(), [], "Manual", language == "en-US", null, CancellationToken.None);
            Assert.AreEqual("shot.png", package.Manifest.Steps[0].SourceLabel);
            foreach (var name in new[] { "guide.md", "guide.html" }) Assert.IsFalse(File.ReadAllText(Path.Combine(package.PackageFolder, name)).Contains("Private"));
            using (var bitmap = new Bitmap(package.Manifest.Steps[0].ImagePath)) Assert.AreEqual(Color.FromArgb(255, 18, 18, 18).ToArgb(), bitmap.GetPixel(10, 10).ToArgb());
            using var word = WordprocessingDocument.Open(Path.Combine(package.PackageFolder, "guide.docx"), false);
            var main = word.MainDocumentPart!;
            var body = main.Document!;
            var text = body.InnerText;
            StringAssert.Contains(text, author); StringAssert.Contains(text, steps); Assert.IsFalse(text.Contains("Private"));
            StringAssert.Contains(body.OuterXml, "Heading1");
            StringAssert.Contains(main.StyleDefinitionsPart!.Styles!.OuterXml, language);
            Assert.AreEqual("shot.png", package.Manifest.Steps[0].SourceLabel);
            Assert.IsTrue(body.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>().All(x => x.Description?.Value == document.StepTitle));
            Assert.AreEqual(0, new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019).Validate(word).Count());
            var pdf = Encoding.Latin1.GetString(File.ReadAllBytes(Path.Combine(package.PackageFolder, "guide.pdf")));
            Assert.IsTrue(Regex.IsMatch(pdf, @"/Lang\s*\(" + language + @"\)"), "PDF language metadata");
        }
        finally { CultureInfo.CurrentCulture = previousCulture; if (Directory.Exists(root)) Directory.Delete(root, true); }
    }

    private static ScreenshotDocument Source()
    {
        using var image = new Bitmap(100, 100); using (var graphics = Graphics.FromImage(image)) graphics.Clear(Color.White);
        using var stream = new MemoryStream(); image.Save(stream, ImageFormat.Png);
        return new() { ImageBytes = stream.ToArray(), SourcePixelWidth = 100, SourcePixelHeight = 100, Origin = DocumentOrigin.FileImport, FileNameLabel = "shot.png", SourceLabel = @"C:\Users\Private\shot.png" };
    }

    private static Dictionary<string, GradientPaletteDefinition> Palettes()
    {
        var color = Windows.UI.Color.FromArgb(255, 255, 0, 0);
        return new() { ["Sunset"] = new("Sunset", "Test", color, color) { Shades = [color] } };
    }
}
