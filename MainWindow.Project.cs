using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media.Imaging;
using Windows.Foundation;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.System;
using WinRT.Interop;

namespace SnapSlate;

public sealed partial class MainWindow
{
    private const int MaxProjectHistoryEntries = 40;
    private const string ProjectFileExtension = ".snapslate";

    private readonly List<SnapSlateProjectState> _undoHistory = [];
    private readonly List<SnapSlateProjectState> _redoHistory = [];
    private readonly Dictionary<TextBox, SnapSlateProjectState> _pendingTextEditSnapshots = new();

    private readonly ProjectSaveTracker _projectSaveTracker = new();
    private bool _isSavingProject;
    private bool _isOpeningProject;
    private bool _isImportingImage;
    private bool _isProjectDialogOpen;
    private bool _isCloseConfirmationOpen;
    private bool _allowClose;
    private string? _currentProjectPath;
    private SnapSlateProjectState? _pendingDragSnapshot;
    private bool _isApplyingProjectState;

    private void InitializeProjectCommands()
    {
        TrackProjectTextBoxes();
        ConfigureKeyboardAccelerators();
        ApplyAccessibleFieldLabels();
        AddKeyboardAccelerator(VirtualKey.V, VirtualKeyModifiers.Control | VirtualKeyModifiers.Shift, ImportClipboardKeyboardAccelerator_Invoked);
        UpdateUndoRedoButtons();
    }

    private void TrackProjectTextBoxes()
    {
        foreach (var textBox in new[]
        {
            DocumentTitleTextBox,
            StepTitleTextBox,
            StepNoteTextBox,
            AudienceTextBox,
            AuthorTextBox,
            DocumentVersionTextBox,
            AnnotationTextBox,
            StickerLegendTextBox
        })
        {
            textBox.GotFocus += ProjectTextBox_GotFocus;
            textBox.LostFocus += ProjectTextBox_LostFocus;
        }
    }

    private void ConfigureKeyboardAccelerators()
    {
        AddKeyboardAccelerator(VirtualKey.S, VirtualKeyModifiers.Control, SaveProjectKeyboardAccelerator_Invoked);
        AddKeyboardAccelerator(VirtualKey.N, VirtualKeyModifiers.Control, NewProjectKeyboardAccelerator_Invoked);
        AddKeyboardAccelerator(VirtualKey.N, VirtualKeyModifiers.Control | VirtualKeyModifiers.Shift, DemoProjectKeyboardAccelerator_Invoked);
        AddKeyboardAccelerator(VirtualKey.O, VirtualKeyModifiers.Control, OpenProjectKeyboardAccelerator_Invoked);
        AddKeyboardAccelerator(VirtualKey.C, VirtualKeyModifiers.Control | VirtualKeyModifiers.Shift, CopyImageKeyboardAccelerator_Invoked);
        AddKeyboardAccelerator(VirtualKey.Z, VirtualKeyModifiers.Control, UndoProjectKeyboardAccelerator_Invoked);
        AddKeyboardAccelerator(VirtualKey.Y, VirtualKeyModifiers.Control, RedoProjectKeyboardAccelerator_Invoked);
    }

    private async void ImportClipboardKeyboardAccelerator_Invoked(KeyboardAccelerator sender, KeyboardAcceleratorInvokedEventArgs args)
    {
        args.Handled = true;
        await TryImportClipboardCaptureAsync(force: true);
    }

    private void AddKeyboardAccelerator(
        VirtualKey key,
        VirtualKeyModifiers modifiers,
        TypedEventHandler<KeyboardAccelerator, KeyboardAcceleratorInvokedEventArgs> invokedHandler)
    {
        var accelerator = new KeyboardAccelerator
        {
            Key = key,
            Modifiers = modifiers
        };

        accelerator.Invoked += invokedHandler;
        RootGrid.KeyboardAccelerators.Add(accelerator);
    }

    private async void SaveProjectKeyboardAccelerator_Invoked(KeyboardAccelerator sender, KeyboardAcceleratorInvokedEventArgs args)
    {
        args.Handled = true;
        await SaveProjectAsync();
    }

    private async void NewProjectKeyboardAccelerator_Invoked(KeyboardAccelerator sender, KeyboardAcceleratorInvokedEventArgs args)
    {
        args.Handled = true;
        await CreateNewProjectAsync();
    }

    private async void DemoProjectKeyboardAccelerator_Invoked(KeyboardAccelerator sender, KeyboardAcceleratorInvokedEventArgs args)
    {
        args.Handled = true;
        await LoadDemoProjectAsync();
    }

    private async void OpenProjectKeyboardAccelerator_Invoked(KeyboardAccelerator sender, KeyboardAcceleratorInvokedEventArgs args)
    {
        args.Handled = true;
        await OpenProjectAsync();
    }

    private async void UndoProjectKeyboardAccelerator_Invoked(KeyboardAccelerator sender, KeyboardAcceleratorInvokedEventArgs args)
    {
        if (IsTextInputFocused())
        {
            return;
        }

        args.Handled = true;
        await UndoProjectAsync();
    }

    private async void RedoProjectKeyboardAccelerator_Invoked(KeyboardAccelerator sender, KeyboardAcceleratorInvokedEventArgs args)
    {
        if (IsTextInputFocused())
        {
            return;
        }

        args.Handled = true;
        await RedoProjectAsync();
    }

    private bool IsTextInputFocused()
    {
        if (RootGrid.XamlRoot is null)
        {
            return false;
        }

        return FocusManager.GetFocusedElement(RootGrid.XamlRoot) is TextBox or PasswordBox;
    }

    private void ProjectTextBox_GotFocus(object sender, RoutedEventArgs e)
    {
        if (_isApplyingProjectState || sender is not TextBox textBox || _pendingTextEditSnapshots.ContainsKey(textBox))
        {
            return;
        }

        _pendingTextEditSnapshots[textBox] = CaptureProjectState();
    }

    private void ProjectTextBox_LostFocus(object sender, RoutedEventArgs e)
    {
        if (sender is not TextBox textBox)
        {
            return;
        }

        CommitPendingTextEditSnapshot(textBox);
    }

    private void CommitPendingTextEdits()
    {
        foreach (var textBox in _pendingTextEditSnapshots.Keys.ToArray())
        {
            CommitPendingTextEditSnapshot(textBox);
        }
    }

    private void CommitPendingTextEditSnapshot(TextBox textBox)
    {
        if (!_pendingTextEditSnapshots.Remove(textBox, out var snapshot))
        {
            return;
        }

        var currentState = CaptureProjectState();
        if (!AreProjectStatesEquivalent(snapshot, currentState))
        {
            PushUndoState(snapshot);
        }
    }

    private bool AreProjectStatesEquivalent(SnapSlateProjectState left, SnapSlateProjectState right)
    {
        return ProjectContentFingerprint.Project(left) == ProjectContentFingerprint.Project(right);
    }

    private void PushUndoState(SnapSlateProjectState state)
    {
        if (_isApplyingProjectState)
        {
            return;
        }

        _undoHistory.Add(state);
        if (_undoHistory.Count > MaxProjectHistoryEntries)
        {
            _undoHistory.RemoveAt(0);
        }

        _redoHistory.Clear();
        UpdateUndoRedoButtons();
    }

    private SnapSlateProjectState CaptureProjectState()
    {
        return new SnapSlateProjectState
        {
            Version = 2,
            Metadata = _projectMetadata.Clone(),
            SelectedDocumentId = _currentDocument?.Id,
            SelectedAnnotationId = _selectedAnnotation?.Id,
            Documents = _documents.Select(document => document.CloneForHistory()).ToList()
        };
    }

    private void ResetProjectHistory()
    {
        _undoHistory.Clear();
        _redoHistory.Clear();
        _pendingTextEditSnapshots.Clear();
        _pendingDragSnapshot = null;
        UpdateUndoRedoButtons();
    }

    private void UpdateUndoRedoButtons()
    {
        UndoButton.IsEnabled = _undoHistory.Count > 0;
        RedoButton.IsEnabled = _redoHistory.Count > 0;
    }

    private bool HasUnsavedProjectChanges()
    {
        return _projectSaveTracker.HasChanges(CaptureProjectState());
    }

    private async Task<bool> CreateNewProjectAsync()
    {
        if (_isSavingSettings || _isSavingProject || _isOpeningProject || _isApplyingProjectState || _isCloseConfirmationOpen || _isProjectDialogOpen) return false;
        if (!await EnsureProjectCanBeReplacedAsync(SafetyText("UiCreateANewProject")))
        {
            return false;
        }

        await ApplyProjectStateAsync(CreateBlankProjectState(), projectPath: null, markDocumentsClean: true, clearHistory: true);
        StatusText.Text = SafetyText("UiNewProjectCreated");
        return true;
    }

    private async Task<bool> LoadDemoProjectAsync()
    {
        if (_isSavingSettings || _isSavingProject || _isOpeningProject || _isApplyingProjectState || _isCloseConfirmationOpen || _isProjectDialogOpen) return false;
        if (!await EnsureProjectCanBeReplacedAsync(SafetyText("UiLoadTheDemoTemplate")))
        {
            return false;
        }

        await ApplyProjectStateAsync(CreateDemoProjectState(), projectPath: null, markDocumentsClean: true, clearHistory: true);
        StatusText.Text = SafetyText("UiDemoTemplateLoaded");
        return true;
    }

    private async Task<bool> OpenProjectAsync()
    {
        if (_isSavingSettings || _isOpeningProject || _isSavingProject || _isApplyingProjectState || _isCloseConfirmationOpen || _isProjectDialogOpen) return false;
        _isOpeningProject = true;
        try
        {
            if (!await EnsureProjectCanBeReplacedAsync(SafetyText("UiOpenAProject")))
            {
                return false;
            }

            var picker = new FileOpenPicker();
            picker.FileTypeFilter.Add(ProjectFileExtension);
            picker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;
            InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));

            var file = await picker.PickSingleFileAsync();
            if (file is null)
            {
                StatusText.Text = SafetyText("UiProjectOpenCancelled");
                return false;
            }

            try
            {
                if (new FileInfo(file.Path).Length > ProjectValidator.MaxProjectBytes)
                {
                    throw new InvalidDataException("Project file exceeds the supported limit.");
                }
                var json = await File.ReadAllTextAsync(file.Path);
                var state = ProjectPersistence.Deserialize(json);
                if (state is null)
                {
                    StatusText.Text = SafetyText("UiTheProjectFileIsEmptyOrInvalid");
                    return false;
                }

                await ApplyProjectStateAsync(state, file.Path, markDocumentsClean: true, clearHistory: true);
                StatusText.Text = string.Format(CultureInfo.CurrentCulture, SafetyText("UiProjectLoaded0"), file.Path);
                return true;
            }
            catch (Exception error)
            {
                AppDiagnostics.Record(error);
                StatusText.Text = SafetyText("ProjectOpenFailed");
                return false;
            }

        }
        catch (Exception error)
        {
            AppDiagnostics.Record(error);
            StatusText.Text = SafetyText("ProjectOpenFailed");
            return false;
        }
        finally
        {
            _isOpeningProject = false;
        }
    }

    private async Task<bool> SaveProjectAsync(bool saveAs = false)
    {
        if (_isSavingProject || _isApplyingProjectState) return false;
        _isSavingProject = true;
        try
        {
            CommitPendingTextEdits();

            var targetPath = _currentProjectPath;
            if (saveAs || string.IsNullOrWhiteSpace(targetPath))
            {
                var picker = new FileSavePicker();
                picker.FileTypeChoices.Add(SafetyText("UiSnapSlateProject"), [ProjectFileExtension]);
                picker.SuggestedFileName = GetSuggestedProjectFileName();
                InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));

                var file = await picker.PickSaveFileAsync();
                if (file is null)
                {
                    StatusText.Text = SafetyText("UiProjectSaveCancelled");
                    return false;
                }

                targetPath = file.Path;
            }

            if (string.IsNullOrWhiteSpace(targetPath))
            {
                StatusText.Text = SafetyText("UiNoSaveLocationIsAvailable");
                return false;
            }

            var state = CaptureProjectState();
            foreach (var document in state.Documents)
            {
                document.IsDirty = false;
            }

            try
            {
                var json = ProjectPersistence.Serialize(state);
                await AtomicFile.WriteTextAsync(targetPath, json);
                _currentProjectPath = targetPath;
                MarkProjectClean(state);
                StatusText.Text = string.Format(CultureInfo.CurrentCulture, SafetyText("UiProjectSaved0"), targetPath);
                return true;
            }
            catch (Exception error)
            {
                AppDiagnostics.Record(error);
                StatusText.Text = SafetyText("ProjectSaveFailed");
                return false;
            }

        }
        catch (Exception error)
        {
            AppDiagnostics.Record(error);
            StatusText.Text = SafetyText("ProjectSaveFailed");
            return false;
        }
        finally
        {
            _isSavingProject = false;
        }
    }

    private async void SaveProjectButton_Click(object sender, RoutedEventArgs e)
    {
        await SaveProjectAsync();
    }

    private async void NewProjectButton_Click(object sender, RoutedEventArgs e)
    {
        await CreateNewProjectAsync();
    }

    private async void DemoProjectButton_Click(object sender, RoutedEventArgs e)
    {
        await LoadDemoProjectAsync();
    }

    private async void OpenProjectButton_Click(object sender, RoutedEventArgs e)
    {
        await OpenProjectAsync();
    }

    private async Task<bool> UndoProjectAsync()
    {
        if (_isSavingSettings || _isSavingProject || _isOpeningProject || _isApplyingProjectState || _isCloseConfirmationOpen || _isProjectDialogOpen) return false;
        CommitPendingTextEdits();

        if (_undoHistory.Count == 0)
        {
            StatusText.Text = SafetyText("UiNothingToUndo");
            return false;
        }

        var currentState = CaptureProjectState();
        var previousState = _undoHistory[^1];
        _undoHistory.RemoveAt(_undoHistory.Count - 1);
        _redoHistory.Add(currentState);

        await ApplyProjectStateAsync(previousState, _currentProjectPath, markDocumentsClean: false, clearHistory: false);
        StatusText.Text = SafetyText("UiLastActionUndone");
        UpdateUndoRedoButtons();
        return true;
    }

    private async Task<bool> RedoProjectAsync()
    {
        if (_isSavingSettings || _isSavingProject || _isOpeningProject || _isApplyingProjectState || _isCloseConfirmationOpen || _isProjectDialogOpen) return false;
        CommitPendingTextEdits();

        if (_redoHistory.Count == 0)
        {
            StatusText.Text = SafetyText("UiNothingToRedo");
            return false;
        }

        var currentState = CaptureProjectState();
        var nextState = _redoHistory[^1];
        _redoHistory.RemoveAt(_redoHistory.Count - 1);
        _undoHistory.Add(currentState);

        await ApplyProjectStateAsync(nextState, _currentProjectPath, markDocumentsClean: false, clearHistory: false);
        StatusText.Text = SafetyText("UiLastActionRedone");
        UpdateUndoRedoButtons();
        return true;
    }

    private async Task<bool> EnsureProjectCanBeReplacedAsync(string actionDescription)
    {
        CommitPendingTextEdits();

        if (!HasUnsavedProjectChanges())
        {
            return true;
        }

        if (RootGrid.XamlRoot is null)
        {
            return false;
        }

        var dialog = new ContentDialog
        {
            XamlRoot = RootGrid.XamlRoot,
            Title = string.Format(CultureInfo.CurrentCulture, SafetyText("SavePromptTitle"), actionDescription),
            Content = SafetyText("SavePromptBody"),
            PrimaryButtonText = SafetyText("Save"),
            SecondaryButtonText = SafetyText("Discard"),
            CloseButtonText = SafetyText("Cancel"),
            DefaultButton = ContentDialogButton.Primary
        };

        if (_isProjectDialogOpen) return false;
        _isProjectDialogOpen = true;
        ContentDialogResult result;
        try { result = await dialog.ShowAsync(); }
        finally { _isProjectDialogOpen = false; }
        return result switch
        {
            ContentDialogResult.Primary => await SaveProjectAsync() && !HasUnsavedProjectChanges(),
            ContentDialogResult.Secondary => true,
            _ => false
        };
    }


    private async Task ApplyProjectStateAsync(
        SnapSlateProjectState state,
        string? projectPath,
        bool markDocumentsClean,
        bool clearHistory)
    {
        if (_isApplyingProjectState) return;
        _isApplyingProjectState = true;
        try
        {
            ProjectValidator.Validate(state);
            var prepared = state.Documents.Select(document => document.CloneForHistory()).ToList();
            if (prepared.Count == 0) prepared.Add(CreateBlankProjectDocument());
            foreach (var document in prepared)
            {
                await EnsureDocumentThumbnailAsync(document);
            }
            var selected = prepared.FirstOrDefault(document => document.Id == state.SelectedDocumentId) ?? prepared[0];
            var preparedImage = selected.ImageBytes is { Length: > 0 }
                ? await CreateBitmapImageAsync(selected.ImageBytes) : null;

            // All external data and image decoding succeeded before touching the live session.
            if (clearHistory) ResetProjectHistory();
            else
            {
                _pendingTextEditSnapshots.Clear();
                _pendingDragSnapshot = null;
            }
            if (_settingsDraft is not null) ReturnToProcedureSection();
            ++_projectEpoch;
            _projectMetadata = (state.Metadata ?? GuideMetadata.FromLegacyDocument(state.Documents.FirstOrDefault())).Clone();
            _currentProjectPath = projectPath;
            _documents.Clear();
            foreach (var document in prepared) _documents.Add(document);
            RefreshDocumentCounters();
            RefreshDocumentOrderMetadata();
            DocumentTabsListView.SelectedItem = selected;
            await LoadDocumentAsync(selected, preparedImage);
            if (state.SelectedAnnotationId is Guid annotationId)
            {
                SelectAnnotation(selected.Annotations.FirstOrDefault(annotation => annotation.Id == annotationId));
            }
            if (markDocumentsClean) MarkProjectClean();
            else RefreshProjectDirtyIndicators();
            UpdateUndoRedoButtons();
        }
        finally
        {
            _isApplyingProjectState = false;
        }
    }

    private ScreenshotDocument? FindDocumentById(Guid? documentId)
    {
        if (documentId is null)
        {
            return null;
        }

        return _documents.FirstOrDefault(document => document.Id == documentId);
    }

    private void MarkProjectClean(SnapSlateProjectState? savedState = null)
    {
        _projectSaveTracker.Accept(savedState ?? CaptureProjectState());
        RefreshProjectDirtyIndicators();
    }

    private void RefreshProjectDirtyIndicators()
    {
        _projectSaveTracker.RefreshIndicators(new SnapSlateProjectState { Metadata = _projectMetadata, Documents = _documents.ToList() });
    }

    private void RefreshDocumentCounters()
    {
        _nextDemoIndex = Math.Max(GetNextSequentialIndex("Lorem Ipsum "), GetNextSequentialIndex("Planche "));
        _nextCaptureIndex = GetNextSequentialIndex("Capture ");
    }

    private void RefreshDocumentOrderMetadata()
    {
        for (var index = 0; index < _documents.Count; index++)
        {
            _documents[index].StepNumber = index + 1;
        }
        UpdateWelcomeSurface();
    }

    private int GetNextSequentialIndex(string prefix)
    {
        var highest = 0;
        foreach (var document in _documents)
        {
            if (!document.BaseTitle.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var suffix = document.BaseTitle[prefix.Length..];
            if (int.TryParse(suffix, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index) && index > highest)
            {
                highest = index;
            }
        }

        return Math.Max(1, highest + 1);
    }

    private async Task EnsureDocumentThumbnailAsync(ScreenshotDocument document)
    {
        if (document.ThumbnailSource is not null)
        {
            return;
        }

        if (document.ImageBytes is { Length: > 0 })
        {
            document.ThumbnailSource = await CreateBitmapImageAsync(document.ImageBytes, 160);
            return;
        }

        document.ThumbnailSource = new BitmapImage(new Uri("ms-appx:///Assets/Square150x150Logo.scale-200.png"));
    }

    private static SnapSlateProjectState CreateBlankProjectState()
    {
        var document = CreateBlankProjectDocument();

        return new SnapSlateProjectState
        {
            Documents = [document],
            SelectedDocumentId = document.Id
        };
    }

    private static ScreenshotDocument CreateBlankProjectDocument()
    {
        return new ScreenshotDocument
        {
            BaseTitle = LocalizedText.Get("NewProjectTitle"),
            Origin = DocumentOrigin.BlankProject,
            OriginLabel = "Projet vide",
            SourceLabel = LocalizedText.Get("BlankSource"),
            FileNameLabel = "nouveau-projet.snapslate",
            SourcePixelWidth = (int)SceneWidth,
            SourcePixelHeight = (int)SceneHeight,
            ThumbnailSource = new BitmapImage(new Uri("ms-appx:///Assets/Square150x150Logo.scale-200.png")),
            IsDirty = false
        };
    }

    private SnapSlateProjectState CreateDemoProjectState()
    {
        var document = CreateDemoDocument(isInitial: true);
        document.IsDirty = false;

        return new SnapSlateProjectState
        {
            Documents = [document],
            SelectedDocumentId = document.Id
        };
    }

    private string GetSuggestedProjectFileName()
    {
        var title = string.IsNullOrWhiteSpace(_projectMetadata.Title) ? "snapslate-project" : _projectMetadata.Title;
        var invalidCharacters = Path.GetInvalidFileNameChars();
        var sanitized = new string(title.Where(character => !invalidCharacters.Contains(character)).ToArray());
        sanitized = sanitized.Replace(' ', '-').Trim('-').ToLowerInvariant();
        return string.IsNullOrWhiteSpace(sanitized) ? "snapslate-project" : sanitized;
    }
}
