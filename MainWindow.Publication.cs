using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Windows.Foundation;
using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.Storage.Streams;
using WinRT.Interop;

namespace SnapSlate;

public sealed partial class MainWindow
{
    internal static readonly HttpClient PublicationHttpClient = new()
    {
        Timeout = TimeSpan.FromSeconds(90)
    };

    private async void OcrButton_Click(object sender, RoutedEventArgs e)
    {
        await RunOcrAsync();
    }

    private async void PublishNotionButton_Click(object sender, RoutedEventArgs e)
    {
        await PublishToNotionAsync();
    }

    private async void PublishConfluenceButton_Click(object sender, RoutedEventArgs e)
    {
        await PublishToConfluenceAsync();
    }

    private async void PublishSharePointButton_Click(object sender, RoutedEventArgs e)
    {
        await PublishToSharePointAsync();
    }

    private async void BrowseSharePointFolderButton_Click(object sender, RoutedEventArgs e)
    {
        var picker = new FolderPicker();
        picker.FileTypeFilter.Add("*");
        picker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
        InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));

        var folder = await picker.PickSingleFolderAsync();
        if (folder is null)
        {
            StatusText.Text = T("Choix du dossier SharePoint annulé.", "SharePoint folder selection cancelled.");
            return;
        }

        SharePointSyncFolderTextBox.Text = folder.Path;
        App.Preferences.SharePointSyncFolder = folder.Path;
        App.Preferences.Save();
        StatusText.Text = string.Format(CultureInfo.CurrentCulture, T("Dossier SharePoint : {0}", "SharePoint folder: {0}"), folder.Path);
    }

    private void NotionParentPageIdTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializingUi || sender is not TextBox textBox)
        {
            return;
        }

        App.Preferences.NotionParentPageId = textBox.Text.Trim();
        App.Preferences.Save();
    }

    private void ConfluenceBaseUrlTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializingUi || sender is not TextBox textBox)
        {
            return;
        }

        App.Preferences.ConfluenceBaseUrl = textBox.Text.Trim().TrimEnd('/');
        App.Preferences.Save();
    }

    private void ConfluenceSpaceKeyTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializingUi || sender is not TextBox textBox)
        {
            return;
        }

        App.Preferences.ConfluenceSpaceKey = textBox.Text.Trim();
        App.Preferences.Save();
    }

    private void ConfluenceParentPageIdTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializingUi || sender is not TextBox textBox)
        {
            return;
        }

        App.Preferences.ConfluenceParentPageId = textBox.Text.Trim();
        App.Preferences.Save();
    }

    private void ConfluenceEmailTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializingUi || sender is not TextBox textBox)
        {
            return;
        }

        App.Preferences.ConfluenceEmail = textBox.Text.Trim();
        App.Preferences.Save();
    }

    private void SharePointSyncFolderTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializingUi || sender is not TextBox textBox)
        {
            return;
        }

        App.Preferences.SharePointSyncFolder = textBox.Text.Trim();
        App.Preferences.Save();
    }

    private async Task RunOcrAsync()
    {
        CommitPendingTextEdits();

        var document = _currentDocument;
        if (document is null)
        {
            StatusText.Text = T("Aucun document n’est sélectionné.", "No document is selected.");
            return;
        }

        if (document.ImageBytes is not { Length: > 0 })
        {
            StatusText.Text = T("L’OCR nécessite une capture importée.", "OCR requires an imported capture.");
            return;
        }

        try
        {
            var recognizedText = await OcrService.RecognizeTextAsync(document.ImageBytes);
            if (string.IsNullOrWhiteSpace(recognizedText))
            {
                StatusText.Text = T("Aucun texte n’a été détecté.", "No text was detected.");
                return;
            }

            AddTextAnnotation(document, new Point(48, 48), recognizedText);
            var firstLine = recognizedText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? recognizedText;
            StatusText.Text = string.Format(CultureInfo.CurrentCulture, T("OCR ajouté : {0}", "OCR added: {0}"), firstLine);
        }
        catch
        {
            StatusText.Text = T("Impossible d’exécuter l’OCR sur cette image.", "Unable to run OCR on this image.");
        }
    }

    private async Task PublishToNotionAsync()
    {
        CommitPendingTextEdits();

        var token = NotionTokenPasswordBox.Password.Trim();
        var parentPageId = NotionParentPageIdTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(token) || string.IsNullOrWhiteSpace(parentPageId))
        {
            StatusText.Text = T("Renseignez le jeton Notion et l’ID de la page parente.", "Enter the Notion token and the parent page ID.");
            return;
        }

        var tempRoot = CreatePublicationRoot();
        try
        {
            var package = await BuildGuidePackageAsync(tempRoot, []);
            if (package is null)
            {
                return;
            }

            var markdown = BuildPublicationMarkdown(package.Manifest);
            var pageUrl = await NotionPublisher.PublishAsync(token, parentPageId, markdown);
            StatusText.Text = pageUrl is null
                ? T("Publication Notion terminée.", "Notion publication completed.")
                : string.Format(CultureInfo.CurrentCulture, T("Publication Notion terminée : {0}", "Notion publication completed: {0}"), pageUrl);
        }
        catch
        {
            StatusText.Text = T("Impossible de publier vers Notion.", "Unable to publish to Notion.");
        }
        finally
        {
            TryDeleteDirectory(tempRoot);
        }
    }

    private async Task PublishToConfluenceAsync()
    {
        CommitPendingTextEdits();

        var baseUrl = ConfluenceBaseUrlTextBox.Text.Trim().TrimEnd('/');
        var spaceKey = ConfluenceSpaceKeyTextBox.Text.Trim();
        var email = ConfluenceEmailTextBox.Text.Trim();
        var token = ConfluenceTokenPasswordBox.Password.Trim();
        var parentPageText = ConfluenceParentPageIdTextBox.Text.Trim();

        if (string.IsNullOrWhiteSpace(baseUrl) || string.IsNullOrWhiteSpace(spaceKey) || string.IsNullOrWhiteSpace(email) || string.IsNullOrWhiteSpace(token))
        {
            StatusText.Text = T("Renseignez l’URL du site, la clé d’espace, l’e-mail et le jeton Confluence.", "Enter the site URL, space key, email, and Confluence token.");
            return;
        }

        long? parentPageId = null;
        if (!string.IsNullOrWhiteSpace(parentPageText))
        {
            if (!long.TryParse(parentPageText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedParentPageId))
            {
                StatusText.Text = T("L’ID de page parente Confluence doit être numérique.", "The Confluence parent page ID must be numeric.");
                return;
            }

            parentPageId = parsedParentPageId;
        }

        var tempRoot = CreatePublicationRoot();
        try
        {
            var package = await BuildGuidePackageAsync(tempRoot, []);
            if (package is null)
            {
                return;
            }

            var storage = BuildConfluenceStorage(package.Manifest);
            var pageUrl = await ConfluencePublisher.PublishAsync(baseUrl, spaceKey, email, token, package.Manifest.Title, storage, parentPageId);
            StatusText.Text = pageUrl is null
                ? T("Publication Confluence terminée.", "Confluence publication completed.")
                : string.Format(CultureInfo.CurrentCulture, T("Publication Confluence terminée : {0}", "Confluence publication completed: {0}"), pageUrl);
        }
        catch
        {
            StatusText.Text = T("Impossible de publier vers Confluence.", "Unable to publish to Confluence.");
        }
        finally
        {
            TryDeleteDirectory(tempRoot);
        }
    }

    private async Task PublishToSharePointAsync()
    {
        CommitPendingTextEdits();

        var targetFolder = SharePointSyncFolderTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(targetFolder))
        {
            StatusText.Text = T("Choisissez un dossier SharePoint synchronisé.", "Choose a synchronized SharePoint folder.");
            return;
        }

        var tempRoot = CreatePublicationRoot();
        try
        {
            var package = await BuildGuidePackageAsync(tempRoot, []);
            if (package is null)
            {
                return;
            }

            var destinationRoot = Path.GetFullPath(Environment.ExpandEnvironmentVariables(targetFolder));
            Directory.CreateDirectory(destinationRoot);

            var folderName = $"{SanitizePathComponent(package.Manifest.Title)}-{DateTime.Now:yyyyMMdd-HHmmss}";
            var destinationFolder = Path.Combine(destinationRoot, folderName);
            SharePointPublisher.CopyDirectory(package.PackageFolder, destinationFolder);

            StatusText.Text = string.Format(CultureInfo.CurrentCulture, T("Package copié vers {0}", "Package copied to {0}"), destinationFolder);
        }
        catch
        {
            StatusText.Text = T("Impossible de copier le package vers SharePoint.", "Unable to copy the package to SharePoint.");
        }
        finally
        {
            TryDeleteDirectory(tempRoot);
        }
    }

    private async Task<GuidePackageResult?> BuildGuidePackageAsync(string rootFolder, ExportFormatChoice[] requestedFormats)
    {
        if (_documents.Count == 0)
        {
            StatusText.Text = T("Aucune étape à exporter.", "No steps to export.");
            return null;
        }

        var formats = requestedFormats is { Length: > 0 }
            ? requestedFormats.Distinct().ToArray()
            : new[] { ExportFormatChoice.Markdown, ExportFormatChoice.Html, ExportFormatChoice.Pdf, ExportFormatChoice.Docx };

        Directory.CreateDirectory(rootFolder);
        var packageFolder = Path.Combine(rootFolder, $"SnapSlate-{DateTime.Now:yyyyMMdd-HHmmss}");
        var imagesFolder = Path.Combine(packageFolder, "images");
        Directory.CreateDirectory(imagesFolder);

        var exportSteps = new List<GuideExportStep>(_documents.Count);
        var generatedFormats = new List<string>();

        if (formats.Contains(ExportFormatChoice.Png))
        {
            generatedFormats.Add("PNG");
        }

        for (var i = 0; i < _documents.Count; i++)
        {
            var document = _documents[i];
            var imagePath = Path.Combine(imagesFolder, $"step-{i + 1:00}.png");
            await Task.Run(() => GuideImageRenderer.SaveDocumentImage(document, _palettes, imagePath));

            exportSteps.Add(new GuideExportStep(
                i + 1,
                document.BaseTitle,
                document.StepNote,
                document.OriginLabel,
                document.SourceLabel,
                imagePath,
                document.Annotations
                    .Where(annotation => annotation.Kind == AnnotationKind.Sticker)
                    .Select(annotation => new GuideExportLegendItem(
                        annotation.Text,
                        string.IsNullOrWhiteSpace(annotation.LegendText) ? T("À compléter", "To fill in") : annotation.LegendText))
                    .ToArray()));
        }

        var sourceDocument = _currentDocument ?? _documents[0];
        var manifest = new GuideExportManifest(
            sourceDocument.BaseTitle,
            GetTemplateLabel(sourceDocument.TemplateType),
            sourceDocument.Audience,
            sourceDocument.Author,
            sourceDocument.DocumentVersion,
            DateTimeOffset.Now,
            packageFolder,
            exportSteps);

        if (formats.Contains(ExportFormatChoice.Markdown))
        {
            await File.WriteAllTextAsync(Path.Combine(packageFolder, "guide.md"), BuildMarkdown(manifest), Encoding.UTF8);
            generatedFormats.Add("Markdown");
        }

        if (formats.Contains(ExportFormatChoice.Html))
        {
            await File.WriteAllTextAsync(Path.Combine(packageFolder, "guide.html"), BuildHtml(manifest), Encoding.UTF8);
            generatedFormats.Add("HTML");
        }

        if (formats.Contains(ExportFormatChoice.Pdf))
        {
            await GuideExporter.ExportPdfAsync(manifest, Path.Combine(packageFolder, "guide.pdf"));
            generatedFormats.Add("PDF");
        }

        if (formats.Contains(ExportFormatChoice.Docx))
        {
            await GuideExporter.ExportDocxAsync(manifest, Path.Combine(packageFolder, "guide.docx"));
            generatedFormats.Add("DOCX");
        }

        return new GuidePackageResult(packageFolder, manifest, generatedFormats);
    }

    private string BuildPublicationMarkdown(GuideExportManifest manifest)
    {
        return BuildMarkdown(manifest, includeImageReferences: false);
    }

    private string BuildConfluenceStorage(GuideExportManifest manifest)
    {
        var storage = new StringBuilder();
        storage.AppendLine($"<h1>{EscapeHtml(manifest.Title)}</h1>");
        storage.AppendLine($"<p><strong>{EscapeHtml(T("Type de document", "Document type"))}</strong> : {EscapeHtml(manifest.TemplateName)}</p>");
        storage.AppendLine($"<p><strong>{EscapeHtml(T("Public", "Audience"))}</strong> : {EscapeHtml(manifest.Audience)}</p>");
        storage.AppendLine($"<p><strong>{EscapeHtml(T("Auteur", "Author"))}</strong> : {EscapeHtml(manifest.Author)}</p>");
        storage.AppendLine($"<p><strong>{EscapeHtml(T("Version", "Version"))}</strong> : {EscapeHtml(manifest.Version)}</p>");

        foreach (var step in manifest.Steps)
        {
            storage.AppendLine($"<h2>{EscapeHtml($"{step.Index:00}. {step.Title}")}</h2>");

            if (!string.IsNullOrWhiteSpace(step.Note))
            {
                storage.AppendLine($"<p>{EscapeHtml(step.Note).Replace("\r\n", "<br />").Replace("\n", "<br />").Replace("\r", "<br />")}</p>");
            }

            storage.AppendLine($"<p><strong>{EscapeHtml(T("Source", "Source"))}</strong> : {EscapeHtml(step.SourceLabel)}</p>");

            if (step.LegendItems.Count > 0)
            {
                storage.AppendLine($"<p><strong>{EscapeHtml(T("Légende", "Legend"))}</strong></p>");
                storage.AppendLine("<ul>");
                foreach (var legend in step.LegendItems)
                {
                    storage.AppendLine($"<li><strong>{EscapeHtml(legend.StickerLabel)}</strong> : {EscapeHtml(legend.Description)}</li>");
                }
                storage.AppendLine("</ul>");
            }
        }

        return storage.ToString();
    }

    private string CreatePublicationRoot()
    {
        return Path.Combine(Path.GetTempPath(), "SnapSlate", "Publish", Guid.NewGuid().ToString("N"));
    }

    private static void TryDeleteDirectory(string path)
    {
        try
        {
            if (Directory.Exists(path))
            {
                Directory.Delete(path, true);
            }
        }
        catch
        {
        }
    }

    private static string SanitizePathComponent(string value)
    {
        var invalidCharacters = Path.GetInvalidFileNameChars();
        var sanitized = new string(value.Where(character => !invalidCharacters.Contains(character)).ToArray());
        sanitized = sanitized.Replace(' ', '-').Trim('-');
        return string.IsNullOrWhiteSpace(sanitized) ? "SnapSlate" : sanitized;
    }
}

internal sealed record GuidePackageResult(
    string PackageFolder,
    GuideExportManifest Manifest,
    IReadOnlyList<string> GeneratedFormats);

internal static class OcrService
{
    public static async Task<string?> RecognizeTextAsync(byte[] imageBytes)
    {
        using var stream = new InMemoryRandomAccessStream();
        await stream.WriteAsync(imageBytes.AsBuffer());
        stream.Seek(0);

        var decoder = await BitmapDecoder.CreateAsync(stream);
        using var bitmap = await decoder.GetSoftwareBitmapAsync(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Ignore);

        var engine = OcrEngine.TryCreateFromUserProfileLanguages();
        if (engine is null)
        {
            engine = OcrEngine.TryCreateFromLanguage(new Language(CultureInfo.CurrentUICulture.Name))
                ?? OcrEngine.TryCreateFromLanguage(new Language("en-US"))
                ?? OcrEngine.TryCreateFromLanguage(new Language("fr-FR"));
        }

        if (engine is null)
        {
            return null;
        }

        var result = await engine.RecognizeAsync(bitmap);
        var lines = result.Lines
            .Select(line => string.Join(" ", line.Words.Select(word => word.Text).Where(text => !string.IsNullOrWhiteSpace(text))).Trim())
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .ToArray();

        return lines.Length == 0 ? string.Empty : string.Join(Environment.NewLine, lines);
    }
}

internal static class NotionPublisher
{
    private const string NotionVersion = "2026-03-11";

    public static async Task<string?> PublishAsync(string token, string parentPageId, string markdown)
    {
        var request = new HttpRequestMessage(HttpMethod.Post, "https://api.notion.com/v1/pages");
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        request.Headers.TryAddWithoutValidation("Notion-Version", NotionVersion);
        request.Content = JsonContent.Create(new
        {
            parent = new { page_id = parentPageId },
            markdown
        });

        using var response = await MainWindow.PublicationHttpClient.SendAsync(request);
        var payload = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException(string.IsNullOrWhiteSpace(payload) ? $"Notion returned {(int)response.StatusCode}." : payload);
        }

        using var json = JsonDocument.Parse(payload);
        return json.RootElement.TryGetProperty("url", out var urlElement) ? urlElement.GetString() : null;
    }
}

internal static class ConfluencePublisher
{
    public static async Task<string?> PublishAsync(
        string baseUrl,
        string spaceKey,
        string email,
        string token,
        string title,
        string storage,
        long? parentPageId)
    {
        var endpoint = new Uri($"{baseUrl.TrimEnd('/')}/wiki/rest/api/content/");
        var request = new HttpRequestMessage(HttpMethod.Post, endpoint);
        var credentials = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{email}:{token}"));
        request.Headers.Authorization = new AuthenticationHeaderValue("Basic", credentials);
        request.Content = JsonContent.Create(new
        {
            type = "page",
            title,
            space = new { key = spaceKey },
            ancestors = parentPageId.HasValue ? new[] { new { id = parentPageId.Value } } : null,
            body = new
            {
                storage = new
                {
                    value = storage,
                    representation = "storage"
                }
            }
        }, options: new JsonSerializerOptions
        {
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
        });

        using var response = await MainWindow.PublicationHttpClient.SendAsync(request);
        var payload = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException(string.IsNullOrWhiteSpace(payload) ? $"Confluence returned {(int)response.StatusCode}." : payload);
        }

        using var json = JsonDocument.Parse(payload);
        if (json.RootElement.TryGetProperty("_links", out var links) && links.TryGetProperty("webui", out var webUi))
        {
            var relativeUrl = webUi.GetString();
            if (!string.IsNullOrWhiteSpace(relativeUrl))
            {
                return new Uri(new Uri($"{baseUrl.TrimEnd('/')}/"), relativeUrl).ToString();
            }
        }

        return null;
    }
}

internal static class SharePointPublisher
{
    public static void CopyDirectory(string sourceDirectory, string destinationDirectory)
    {
        Directory.CreateDirectory(destinationDirectory);

        foreach (var directory in Directory.GetDirectories(sourceDirectory, "*", SearchOption.AllDirectories))
        {
            Directory.CreateDirectory(directory.Replace(sourceDirectory, destinationDirectory, StringComparison.OrdinalIgnoreCase));
        }

        foreach (var file in Directory.GetFiles(sourceDirectory, "*", SearchOption.AllDirectories))
        {
            var destinationFile = file.Replace(sourceDirectory, destinationDirectory, StringComparison.OrdinalIgnoreCase);
            Directory.CreateDirectory(Path.GetDirectoryName(destinationFile)!);
            File.Copy(file, destinationFile, overwrite: true);
        }
    }
}
