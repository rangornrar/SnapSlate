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

    private string? _currentProjectPath;
    private SnapSlateProjectState? _pendingDragSnapshot;
    private bool _isApplyingProjectState;

    private void InitializeProjectCommands()
    {
        TrackProjectTextBoxes();
        ConfigureKeyboardAccelerators();
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
        AddKeyboardAccelerator(VirtualKey.Z, VirtualKeyModifiers.Control, UndoProjectKeyboardAccelerator_Invoked);
        AddKeyboardAccelerator(VirtualKey.Y, VirtualKeyModifiers.Control, RedoProjectKeyboardAccelerator_Invoked);
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
        return string.Equals(
            ProjectPersistence.Serialize(left),
            ProjectPersistence.Serialize(right),
            StringComparison.Ordinal);
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
        return _documents.Any(document => document.IsDirty);
    }

    private async Task<bool> CreateNewProjectAsync()
    {
        if (!await EnsureProjectCanBeReplacedAsync(T("créer un nouveau projet", "create a new project")))
        {
            return false;
        }

        await ApplyProjectStateAsync(CreateBlankProjectState(), projectPath: null, markDocumentsClean: true, clearHistory: true);
        StatusText.Text = T("Nouveau projet créé.", "New project created.");
        return true;
    }

    private async Task<bool> LoadDemoProjectAsync()
    {
        if (!await EnsureProjectCanBeReplacedAsync(T("charger le modèle de démonstration", "load the demo template")))
        {
            return false;
        }

        await ApplyProjectStateAsync(CreateDemoProjectState(), projectPath: null, markDocumentsClean: true, clearHistory: true);
        StatusText.Text = T("Modèle de démonstration chargé.", "Demo template loaded.");
        return true;
    }

    private async Task<bool> OpenProjectAsync()
    {
        if (!await EnsureProjectCanBeReplacedAsync(T("ouvrir un projet", "open a project")))
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
            StatusText.Text = T("Ouverture de projet annulée.", "Project open cancelled.");
            return false;
        }

        try
        {
            var json = await File.ReadAllTextAsync(file.Path);
            var state = ProjectPersistence.Deserialize(json);
            if (state is null)
            {
                StatusText.Text = T("Le fichier de projet est vide ou invalide.", "The project file is empty or invalid.");
                return false;
            }

            await ApplyProjectStateAsync(state, file.Path, markDocumentsClean: true, clearHistory: true);
            StatusText.Text = string.Format(CultureInfo.CurrentCulture, T("Projet chargé : {0}", "Project loaded: {0}"), file.Path);
            return true;
        }
        catch
        {
            StatusText.Text = T("Impossible de charger ce projet.", "Unable to load this project.");
            return false;
        }
    }

    private async Task<bool> SaveProjectAsync(bool saveAs = false)
    {
        CommitPendingTextEdits();

        var targetPath = _currentProjectPath;
        if (saveAs || string.IsNullOrWhiteSpace(targetPath))
        {
            var picker = new FileSavePicker();
            picker.FileTypeChoices.Add(T("Projet SnapSlate", "SnapSlate project"), [ProjectFileExtension]);
            picker.SuggestedFileName = GetSuggestedProjectFileName();
            InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));

            var file = await picker.PickSaveFileAsync();
            if (file is null)
            {
                StatusText.Text = T("Sauvegarde de projet annulée.", "Project save cancelled.");
                return false;
            }

            targetPath = file.Path;
        }

        if (string.IsNullOrWhiteSpace(targetPath))
        {
            StatusText.Text = T("Aucun emplacement de sauvegarde n'est disponible.", "No save location is available.");
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
            await File.WriteAllTextAsync(targetPath, json);
            _currentProjectPath = targetPath;
            MarkProjectClean();
            StatusText.Text = string.Format(CultureInfo.CurrentCulture, T("Projet sauvegardé : {0}", "Project saved: {0}"), targetPath);
            return true;
        }
        catch
        {
            StatusText.Text = T("Impossible d'écrire le fichier de projet.", "Unable to write the project file.");
            return false;
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
        CommitPendingTextEdits();

        if (_undoHistory.Count == 0)
        {
            StatusText.Text = T("Rien à annuler.", "Nothing to undo.");
            return false;
        }

        var currentState = CaptureProjectState();
        var previousState = _undoHistory[^1];
        _undoHistory.RemoveAt(_undoHistory.Count - 1);
        _redoHistory.Add(currentState);

        await ApplyProjectStateAsync(previousState, _currentProjectPath, markDocumentsClean: false, clearHistory: false);
        StatusText.Text = T("Dernière action annulée.", "Last action undone.");
        UpdateUndoRedoButtons();
        return true;
    }

    private async Task<bool> RedoProjectAsync()
    {
        CommitPendingTextEdits();

        if (_redoHistory.Count == 0)
        {
            StatusText.Text = T("Rien à rétablir.", "Nothing to redo.");
            return false;
        }

        var currentState = CaptureProjectState();
        var nextState = _redoHistory[^1];
        _redoHistory.RemoveAt(_redoHistory.Count - 1);
        _undoHistory.Add(currentState);

        await ApplyProjectStateAsync(nextState, _currentProjectPath, markDocumentsClean: false, clearHistory: false);
        StatusText.Text = T("Dernière action rétablie.", "Last action redone.");
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
            Title = string.Format(CultureInfo.CurrentCulture, T("Enregistrer le projet avant de {0} ?", "Save the project before you {0}?"), actionDescription),
            Content = T("Le projet contient des changements non sauvegardés.", "The project contains unsaved changes."),
            PrimaryButtonText = T("Enregistrer", "Save"),
            SecondaryButtonText = T("Ignorer", "Discard"),
            CloseButtonText = T("Annuler", "Cancel"),
            DefaultButton = ContentDialogButton.Primary
        };

        var result = await dialog.ShowAsync();
        return result switch
        {
            ContentDialogResult.Primary => await SaveProjectAsync(),
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
        _isApplyingProjectState = true;
        try
        {
            if (clearHistory)
            {
                ResetProjectHistory();
            }
            else
            {
                _pendingTextEditSnapshots.Clear();
                _pendingDragSnapshot = null;
            }

            _currentProjectPath = projectPath;
            _documents.Clear();

            var sourceDocuments = state.Documents.Count > 0
                ? state.Documents
                : [CreateBlankProjectDocument()];

            foreach (var sourceDocument in sourceDocuments)
            {
                var document = sourceDocument.CloneForHistory();
                if (markDocumentsClean)
                {
                    document.IsDirty = false;
                }

                await EnsureDocumentThumbnailAsync(document);
                _documents.Add(document);
            }

            RefreshDocumentCounters();
            RefreshDocumentOrderMetadata();

            var selectedDocument = FindDocumentById(state.SelectedDocumentId) ?? _documents.FirstOrDefault();
            if (selectedDocument is null)
            {
                _currentDocument = null;
                _selectedAnnotation = null;
                UpdateUndoRedoButtons();
                UpdateStatus();
                return;
            }

            DocumentTabsListView.SelectedItem = selectedDocument;
            await LoadDocumentAsync(selectedDocument);

            if (state.SelectedAnnotationId is Guid selectedAnnotationId)
            {
                var selectedAnnotation = selectedDocument.Annotations.FirstOrDefault(annotation => annotation.Id == selectedAnnotationId);
                if (selectedAnnotation is not null)
                {
                    SelectAnnotation(selectedAnnotation);
                }
            }

            if (markDocumentsClean)
            {
                MarkProjectClean();
            }

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

    private void MarkProjectClean()
    {
        foreach (var document in _documents)
        {
            document.IsDirty = false;
        }
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
            BaseTitle = "Nouveau projet",
            Origin = DocumentOrigin.BlankProject,
            OriginLabel = "Projet vide",
            SourceLabel = "Démarrer à zéro",
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
        var title = _currentDocument?.BaseTitle ?? "snapslate-project";
        var invalidCharacters = Path.GetInvalidFileNameChars();
        var sanitized = new string(title.Where(character => !invalidCharacters.Contains(character)).ToArray());
        sanitized = sanitized.Replace(' ', '-').Trim('-').ToLowerInvariant();
        return string.IsNullOrWhiteSpace(sanitized) ? "snapslate-project" : sanitized;
    }
}
