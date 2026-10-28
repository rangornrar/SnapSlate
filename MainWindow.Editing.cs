using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;

namespace SnapSlate;

public sealed partial class MainWindow
{
    private GuideMetadata _projectMetadata = new();
    private GuideMetadata? _settingsMetadata;
    private AppPreferences? _settingsDraft;
    private GuideMetadata EditingMetadata => _settingsMetadata ?? _projectMetadata;
    private AppPreferences EditingPreferences => _settingsDraft ?? App.Preferences;
    private int _documentLoadVersion;
    private int _previewVersion;
    private bool _previewRunning;
    private bool _editorClosed;
    private bool _isSavingSettings;
    private CancellationTokenSource? _exportCancellation;

    private void SavePreferencesIfNeeded()
    {
        if (_settingsDraft is null) App.Preferences.Save();
    }

    private void SyncGuideMetadataEditors()
    {
        var wasApplying = _isApplyingDocumentState;
        _isApplyingDocumentState = true;
        try
        {
            var metadata = EditingMetadata;
            DocumentTitleTextBox.Text = metadata.Title;
            TemplateTypeComboBox.SelectedIndex = (int)metadata.TemplateType;
            ExportTemplateComboBox.SelectedIndex = (int)metadata.TemplateType;
            AudienceTextBox.Text = ExportAudienceTextBox.Text = metadata.Audience;
            AuthorTextBox.Text = ExportAuthorTextBox.Text = metadata.Author;
            DocumentVersionTextBox.Text = ExportVersionTextBox.Text = metadata.DocumentVersion;
        }
        finally { _isApplyingDocumentState = wasApplying; }
    }

    private void GuideMetadataChanged()
    {
        SyncGuideMetadataEditors();
        if (_settingsMetadata is null) RefreshProjectDirtyIndicators();
        UpdateExportTemplateSummary();
    }

    private void StepTitleTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isInitializingUi || _isApplyingDocumentState || _currentDocument is null) return;
        if (_currentDocument.StepTitle == StepTitleTextBox.Text) return;
        _currentDocument.StepTitle = StepTitleTextBox.Text;
        if (_currentDocument.ImageBytes is null) RequestScenePreview();
        MarkDocumentDirty(_currentDocument);
    }

    // Coalesce edits into one rendering task and discard results from a previous selection.
    private async void RequestScenePreview()
    {
        ++_previewVersion;
        if (_previewRunning || _editorClosed) return;
        _previewRunning = true;
        try
        {
            while (!_editorClosed && _currentDocument is { } document)
            {
                var version = _previewVersion;
                var snapshot = document.CloneForHistory();
                snapshot.AppliedCropRect = null;
                snapshot.PendingCropRect = null;
                if (_editingTextAnnotation is { } editing)
                    snapshot.Annotations.RemoveAll(annotation => annotation.Id == editing.Id);
                var bytes = await Task.Run(() => GuideImageRenderer.RenderPng(snapshot, _palettes));
                if (_editorClosed) return;
                if (version != _previewVersion || !ReferenceEquals(document, _currentDocument)) continue;
                var image = await CreateBitmapImageAsync(bytes);
                if (_editorClosed) return;
                if (version != _previewVersion || !ReferenceEquals(document, _currentDocument)) continue;
                RenderedSceneImage.Source = image;
                break;
            }
        }
        catch (Exception error)
        {
            AppDiagnostics.Record(error);
            if (!_editorClosed)
            {
                RenderedSceneImage.Source = null;
                StatusText.Text = SafetyText("OperationFailed");
            }
        }
        finally { _previewRunning = false; }
    }

    private void ExportCancelButton_Click(object sender, RoutedEventArgs e) => _exportCancellation?.Cancel();

    private void CanvasWorkspace_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        FloatingToolRail.MaxHeight = Math.Max(48, e.NewSize.Height - 8);
        ApplyFloatingToolRailPosition(EditingPreferences.FloatingToolRailLeft, EditingPreferences.FloatingToolRailTop);
        if (_autoFitPreview) QueueWorkspaceFit();
    }
}
