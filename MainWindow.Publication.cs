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
using System.Threading;
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
    internal static readonly HttpClient PublicationHttpClient = new(new HttpClientHandler { AllowAutoRedirect = false })
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
            StatusText.Text = SafetyText("UiSharePointFolderSelectionCancelled");
            return;
        }

        SharePointSyncFolderTextBox.Text = folder.Path;
        EditingPreferences.SharePointSyncFolder = folder.Path;
        SavePreferencesIfNeeded();
        StatusText.Text = string.Format(CultureInfo.CurrentCulture, SafetyText("UiSharePointFolder0"), folder.Path);
    }

    private void NotionParentPageIdTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializingUi || sender is not TextBox textBox)
        {
            return;
        }

        EditingPreferences.NotionParentPageId = textBox.Text.Trim();
        SavePreferencesIfNeeded();
    }

    private void ConfluenceBaseUrlTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializingUi || sender is not TextBox textBox)
        {
            return;
        }

        EditingPreferences.ConfluenceBaseUrl = textBox.Text.Trim().TrimEnd('/');
        SavePreferencesIfNeeded();
    }

    private void ConfluenceSpaceKeyTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializingUi || sender is not TextBox textBox)
        {
            return;
        }

        EditingPreferences.ConfluenceSpaceKey = textBox.Text.Trim();
        SavePreferencesIfNeeded();
    }

    private void ConfluenceParentPageIdTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializingUi || sender is not TextBox textBox)
        {
            return;
        }

        EditingPreferences.ConfluenceParentPageId = textBox.Text.Trim();
        SavePreferencesIfNeeded();
    }

    private void ConfluenceEmailTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializingUi || sender is not TextBox textBox)
        {
            return;
        }

        EditingPreferences.ConfluenceEmail = textBox.Text.Trim();
        SavePreferencesIfNeeded();
    }

    private void SharePointSyncFolderTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializingUi || sender is not TextBox textBox)
        {
            return;
        }

        EditingPreferences.SharePointSyncFolder = textBox.Text.Trim();
        SavePreferencesIfNeeded();
    }

    private async Task RunOcrAsync()
    {
        CommitPendingTextEdits();

        var document = _currentDocument;
        if (document is null)
        {
            StatusText.Text = SafetyText("UiNoDocumentIsSelected");
            return;
        }

        if (document.ImageBytes is not { Length: > 0 })
        {
            StatusText.Text = SafetyText("UiOCRRequiresAnImportedCapture");
            return;
        }

        try
        {
            var recognizedText = await OcrService.RecognizeTextAsync(document.ImageBytes);
            if (string.IsNullOrWhiteSpace(recognizedText))
            {
                StatusText.Text = SafetyText("UiNoTextWasDetected");
                return;
            }

            AddTextAnnotation(document, new Point(48, 48), recognizedText);
            var firstLine = recognizedText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? recognizedText;
            StatusText.Text = string.Format(CultureInfo.CurrentCulture, SafetyText("UiOCRAdded0"), firstLine);
        }
        catch
        {
            StatusText.Text = SafetyText("UiUnableToRunOCROnThisImage");
        }
    }

    private async Task PublishToNotionAsync()
    {
        CommitPendingTextEdits();

        var token = NotionTokenPasswordBox.Password.Trim();
        var parentPageId = NotionParentPageIdTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(token) || string.IsNullOrWhiteSpace(parentPageId))
        {
            StatusText.Text = SafetyText("UiEnterTheNotionTokenAndTheParent");
            return;
        }

        try
        {
            var manifest = PublicationTextBuilder.Capture(_documents, _projectMetadata,
                GetTemplateLabel(_projectMetadata.TemplateType), UseEnglish ? "en-US" : "fr-FR");

            var markdown = PublicationTextBuilder.Markdown(manifest);
            var pageUrl = await TextPublicationService.NotionAsync(PublicationHttpClient, token, parentPageId, manifest.Title, markdown);
            StatusText.Text = pageUrl is null
                ? SafetyText("UiNotionPublicationCompleted")
                : string.Format(CultureInfo.CurrentCulture, SafetyText("UiNotionPublicationCompleted0"), pageUrl);
        }
        catch (Exception error) { ShowPublicationFailure(error); }
    }

    private async Task PublishToConfluenceAsync()
    {
        CommitPendingTextEdits();

        var baseUrl = ConfluenceBaseUrlTextBox.Text.Trim().TrimEnd('/');
        var spaceKey = ConfluenceSpaceKeyTextBox.Text.Trim();
        var email = ConfluenceEmailTextBox.Text.Trim();
        var token = ConfluenceTokenPasswordBox.Password.Trim();
        try { PublicationEndpoint.Confluence(baseUrl); }
        catch (ArgumentException)
        {
            StatusText.Text = SafetyText("HttpsRequired");
            return;
        }
        var parentPageText = ConfluenceParentPageIdTextBox.Text.Trim();

        if (string.IsNullOrWhiteSpace(baseUrl) || string.IsNullOrWhiteSpace(spaceKey) || string.IsNullOrWhiteSpace(email) || string.IsNullOrWhiteSpace(token))
        {
            StatusText.Text = SafetyText("UiEnterTheSiteURLSpaceKeyEmail");
            return;
        }

        long? parentPageId = null;
        if (!string.IsNullOrWhiteSpace(parentPageText))
        {
            if (!long.TryParse(parentPageText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedParentPageId))
            {
                StatusText.Text = SafetyText("UiTheConfluenceParentPageIDMustBe");
                return;
            }

            parentPageId = parsedParentPageId;
        }

        try
        {
            var manifest = PublicationTextBuilder.Capture(_documents, _projectMetadata,
                GetTemplateLabel(_projectMetadata.TemplateType), UseEnglish ? "en-US" : "fr-FR");

            var storage = PublicationTextBuilder.ConfluenceStorage(manifest);
            var pageUrl = await TextPublicationService.ConfluenceAsync(PublicationHttpClient, baseUrl, spaceKey, email, token, manifest.Title, storage, parentPageId);
            StatusText.Text = pageUrl is null
                ? SafetyText("UiConfluencePublicationCompleted")
                : string.Format(CultureInfo.CurrentCulture, SafetyText("UiConfluencePublicationCompleted0"), pageUrl);
        }
        catch (Exception error) { ShowPublicationFailure(error); }
    }

    private async Task PublishToSharePointAsync()
    {
        CommitPendingTextEdits();

        var targetFolder = SharePointSyncFolderTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(targetFolder))
        {
            StatusText.Text = SafetyText("UiChooseASynchronizedSharePointFolder");
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

            StatusText.Text = string.Format(CultureInfo.CurrentCulture, SafetyText("UiPackageCopiedTo0"), destinationFolder);
        }
        catch
        {
            StatusText.Text = SafetyText("UiUnableToCopyThePackageToSharePoint");
        }
        finally
        {
            TryDeleteDirectory(tempRoot);
        }
    }

    private async Task<GuidePackageResult?> BuildGuidePackageAsync(string rootFolder, ExportFormatChoice[] requestedFormats, System.Collections.Generic.IReadOnlyList<ScreenshotDocument>? documents = null)
    {
        if (_showWelcome || _documents.Count == 0)
        {
            StatusText.Text = SafetyText("UiNoStepsToExport");
            return null;
        }

        if (_exportCancellation is not null)
        {
            StatusText.Text = SafetyText("ExportAlreadyRunning");
            return null;
        }
        using var cancellation = new CancellationTokenSource();
        _exportCancellation = cancellation;
        ExportCancelButton.Visibility = ExportProgressRing.Visibility = Visibility.Visible;
        ExportProgressRing.IsActive = true;

        try
        {
            return await GuidePackageBuilder.BuildAsync(rootFolder, documents ?? _documents, _projectMetadata,
                _palettes, requestedFormats, GetTemplateLabel(_projectMetadata.TemplateType), UseEnglish,
                (step, count) => StatusText.Text = string.Format(CultureInfo.CurrentCulture,
                    SafetyText("ExportProgress"), step, count), cancellation.Token);
        }
        catch (OperationCanceledException)
        {
            StatusText.Text = SafetyText("ExportCancelled");
            return null;
        }
        finally
        {
            _exportCancellation = null;
            ExportProgressRing.IsActive = false;
            ExportCancelButton.Visibility = ExportProgressRing.Visibility = Visibility.Collapsed;
        }
    }

    private void ShowPublicationFailure(Exception error)
    {
        AppDiagnostics.Record(error);
        StatusText.Text = error is PublicationException failure ? failure.StatusCode switch
        {
            System.Net.HttpStatusCode.Unauthorized => SafetyText("PublicationUnauthorized"),
            System.Net.HttpStatusCode.Forbidden => SafetyText("PublicationForbidden"),
            System.Net.HttpStatusCode.TooManyRequests => SafetyText("PublicationRateLimited"),
            System.Net.HttpStatusCode.NotFound => SafetyText("PublicationNotFound"),
            _ => string.Format(CultureInfo.CurrentCulture, SafetyText("PublicationHttpFailed"), (int)failure.StatusCode)
        } : SafetyText("PublicationFailed");
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
