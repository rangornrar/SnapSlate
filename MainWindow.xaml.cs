using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Security.Cryptography;
using System.Threading.Tasks;
using Microsoft.UI;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Media.Imaging;
using Microsoft.UI.Xaml.Shapes;
using Windows.ApplicationModel.DataTransfer;
using Windows.Foundation;
using Windows.Graphics;
using Windows.Graphics.Imaging;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.Storage.Streams;
using Windows.UI;
using WinRT.Interop;
using XamlPath = Microsoft.UI.Xaml.Shapes.Path;

namespace SnapSlate;

public sealed partial class MainWindow : Window
{
    private const double SceneWidth = 1360;
    private const double SceneHeight = 820;
    private const double StickerSize = 60;

    private readonly ObservableCollection<ScreenshotDocument> _documents = [];
    private readonly Dictionary<string, GradientPaletteDefinition> _palettes;
    private readonly Dictionary<EditorTool, string> _toolLabels = new()
    {
        [EditorTool.Crop] = "Crop",
        [EditorTool.Text] = "Texte",
        [EditorTool.Sticker] = "Gommette",
        [EditorTool.ArrowStraight] = "Fleche droite",
        [EditorTool.ArrowCurved] = "Fleche courbe",
        [EditorTool.Rectangle] = "Rectangle",
        [EditorTool.Ellipse] = "Ovale"
    };

    private readonly Dictionary<EditorTool, string> _toolHints = new()
    {
        [EditorTool.Crop] = "Glissez une zone puis appliquez le crop.",
        [EditorTool.Text] = "Cliquez dans la scene pour poser un texte.",
        [EditorTool.Sticker] = "Cliquez pour poser une gommette.",
        [EditorTool.ArrowStraight] = "Glissez pour tracer une fleche droite.",
        [EditorTool.ArrowCurved] = "Glissez pour tracer une fleche courbe.",
        [EditorTool.Rectangle] = "Glissez pour entourer une zone rectangulaire.",
        [EditorTool.Ellipse] = "Glissez pour entourer une zone ovale."
    };

    private readonly List<Button> _toolButtons;
    private readonly Dictionary<string, Button> _paletteButtons;
    private EditorTool _currentTool = EditorTool.ArrowStraight;
    private ScreenshotDocument? _currentDocument;
    private FrameworkElement? _activeElement;
    private Point _dragStart;
    private bool _isDragging;
    private bool _isApplyingDocumentState;
    private bool _isProcessingClipboard;
    private string? _lastClipboardHash;
    private DateTimeOffset _lastClipboardImportAt = DateTimeOffset.MinValue;
    private int _nextCaptureIndex = 1;
    private int _nextDemoIndex = 1;

    public MainWindow()
    {
        InitializeComponent();

        ExtendsContentIntoTitleBar = true;
        SetTitleBar(AppTitleBar);

        AppWindow.SetIcon("Assets/AppIcon.ico");
        AppWindow.Resize(new SizeInt32(1760, 1080));
        AppWindow.TitleBar.PreferredHeightOption = TitleBarHeightOption.Tall;

        _palettes = CreatePalettes();
        _toolButtons =
        [
            CropToolButton,
            TextToolButton,
            StickerToolButton,
            ArrowStraightToolButton,
            ArrowCurvedToolButton,
            RectangleToolButton,
            EllipseToolButton
        ];
        _paletteButtons = new Dictionary<string, Button>(StringComparer.Ordinal)
        {
            ["Sunset"] = SunsetPaletteButton,
            ["Ember"] = EmberPaletteButton,
            ["Citrus"] = CitrusPaletteButton,
            ["Mint"] = MintPaletteButton,
            ["Lagoon"] = LagoonPaletteButton,
            ["Sky"] = SkyPaletteButton,
            ["Ocean"] = OceanPaletteButton,
            ["Rose"] = RosePaletteButton,
            ["Berry"] = BerryPaletteButton,
            ["Grape"] = GrapePaletteButton
        };

        PreviewViewport.Clip = new RectangleGeometry();
        DocumentTabsListView.ItemsSource = _documents;
        StrokeThicknessSlider.Minimum = 2;
        StrokeThicknessSlider.Maximum = 14;
        StrokeThicknessSlider.Value = 6;

        SelectTool(EditorTool.ArrowStraight);
        AddDocument(CreateDemoDocument(isInitial: true), select: true);

        Clipboard.ContentChanged += Clipboard_ContentChanged;
    }

    private static Dictionary<string, GradientPaletteDefinition> CreatePalettes()
    {
        return new Dictionary<string, GradientPaletteDefinition>(StringComparer.Ordinal)
        {
            ["Sunset"] = new("Sunset", "sunset note", ColorHelper.FromArgb(255, 240, 90, 65), ColorHelper.FromArgb(255, 255, 198, 110)),
            ["Ember"] = new("Ember", "ember alert", ColorHelper.FromArgb(255, 189, 60, 43), ColorHelper.FromArgb(255, 240, 107, 89)),
            ["Citrus"] = new("Citrus", "citrus focus", ColorHelper.FromArgb(255, 240, 167, 59), ColorHelper.FromArgb(255, 246, 228, 122)),
            ["Mint"] = new("Mint", "mint product", ColorHelper.FromArgb(255, 22, 179, 142), ColorHelper.FromArgb(255, 132, 226, 197)),
            ["Lagoon"] = new("Lagoon", "lagoon flow", ColorHelper.FromArgb(255, 15, 158, 175), ColorHelper.FromArgb(255, 115, 225, 233)),
            ["Sky"] = new("Sky", "sky workflow", ColorHelper.FromArgb(255, 31, 132, 224), ColorHelper.FromArgb(255, 140, 203, 255)),
            ["Ocean"] = new("Ocean", "ocean depth", ColorHelper.FromArgb(255, 23, 58, 115), ColorHelper.FromArgb(255, 70, 135, 224)),
            ["Rose"] = new("Rose", "rose highlight", ColorHelper.FromArgb(255, 216, 90, 136), ColorHelper.FromArgb(255, 245, 160, 189)),
            ["Berry"] = new("Berry", "berry spark", ColorHelper.FromArgb(255, 165, 63, 164), ColorHelper.FromArgb(255, 240, 123, 197)),
            ["Grape"] = new("Grape", "grape accent", ColorHelper.FromArgb(255, 103, 70, 195), ColorHelper.FromArgb(255, 176, 137, 255))
        };
    }

    private ScreenshotDocument CreateDemoDocument(bool isInitial = false)
    {
        var title = isInitial ? "Accueil" : $"Planche {_nextDemoIndex++:00}";
        var palette = _palettes["Sunset"];
        var document = new ScreenshotDocument
        {
            BaseTitle = title,
            Origin = DocumentOrigin.Demo,
            OriginLabel = "Demo",
            SourceLabel = "Planche de demonstration interne",
            FileNameLabel = "demo-reference.png",
            SelectedPaletteKey = palette.Key,
            PaletteDisplayName = palette.DisplayName,
            AnnotationText = "Note rapide",
            StickerModeIndex = 0,
            StrokeThickness = 6,
            ResetStickerNumberOnColorChange = false,
            IsDirty = false
        };

        document.Annotations.Add(new AnnotationModel
        {
            Kind = AnnotationKind.Rectangle,
            Bounds = new Rect(460, 245, 430, 260),
            PaletteKey = "Sunset",
            StrokeThickness = 6
        });
        document.Annotations.Add(new AnnotationModel
        {
            Kind = AnnotationKind.ArrowCurved,
            StartPoint = new Point(980, 205),
            EndPoint = new Point(885, 310),
            PaletteKey = "Sky",
            StrokeThickness = 7
        });
        document.Annotations.Add(new AnnotationModel
        {
            Kind = AnnotationKind.Sticker,
            Bounds = new Rect(970, 170, StickerSize, StickerSize),
            PaletteKey = "Mint",
            Text = "A"
        });

        return document;
    }

    private void AddDocument(ScreenshotDocument document, bool select)
    {
        _documents.Add(document);

        if (select || _currentDocument is null)
        {
            DocumentTabsListView.SelectedItem = document;
        }
    }

    private async Task<ScreenshotDocument?> CreateDocumentFromFileAsync(StorageFile file)
    {
        var buffer = await FileIO.ReadBufferAsync(file);
        var bytes = buffer.ToArray();
        if (bytes.Length == 0)
        {
            return null;
        }

        var palette = _palettes["Sunset"];
        return new ScreenshotDocument
        {
            BaseTitle = System.IO.Path.GetFileNameWithoutExtension(file.Name),
            Origin = DocumentOrigin.FileImport,
            OriginLabel = "Fichier importe",
            SourceLabel = file.Path,
            FileNameLabel = file.Name,
            ImageBytes = bytes,
            SelectedPaletteKey = palette.Key,
            PaletteDisplayName = palette.DisplayName,
            AnnotationText = "Note rapide",
            StickerModeIndex = 0,
            StrokeThickness = 6,
            ResetStickerNumberOnColorChange = false,
            IsDirty = false
        };
    }

    private ScreenshotDocument CreateDocumentFromClipboard(byte[] bytes)
    {
        var palette = _palettes["Sunset"];
        var captureIndex = _nextCaptureIndex++;
        return new ScreenshotDocument
        {
            BaseTitle = $"Capture {captureIndex:00}",
            Origin = DocumentOrigin.ClipboardCapture,
            OriginLabel = "Win + Shift + S",
            SourceLabel = $"Capture Windows importee a {DateTime.Now:HH:mm:ss}",
            FileNameLabel = $"capture-{captureIndex:00}.png",
            ImageBytes = bytes,
            SelectedPaletteKey = palette.Key,
            PaletteDisplayName = palette.DisplayName,
            AnnotationText = "Note rapide",
            StickerModeIndex = 0,
            StrokeThickness = 6,
            ResetStickerNumberOnColorChange = false,
            IsDirty = true
        };
    }

    private async void DocumentTabsListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (DocumentTabsListView.SelectedItem is ScreenshotDocument document)
        {
            await LoadDocumentAsync(document);
        }
    }

    private async Task LoadDocumentAsync(ScreenshotDocument document)
    {
        if (_currentDocument == document && !_isApplyingDocumentState)
        {
            return;
        }

        _currentDocument = document;
        _isApplyingDocumentState = true;

        try
        {
            CanvasSummaryText.Text = document.Origin switch
            {
                DocumentOrigin.ClipboardCapture => "Capture Win + Shift + S importee automatiquement dans un nouvel onglet.",
                DocumentOrigin.FileImport => "Fichier importe dans un nouvel onglet pour annoter sans ecraser l'original.",
                _ => "Planche de demo : annotez la scene ou importez un screenshot."
            };
            CanvasFileNameText.Text = document.FileNameLabel;
            ImageSourceText.Text = $"Source : {document.SourceLabel}";

            AnnotationTextBox.Text = document.AnnotationText;
            StickerModeComboBox.SelectedIndex = document.StickerModeIndex;
            StrokeThicknessSlider.Value = document.StrokeThickness;
            ResetStickerNumberOnColorChangeCheckBox.IsChecked = document.ResetStickerNumberOnColorChange;
            StrokeSummaryText.Text = $"Trait actuel : {document.StrokeThickness:0} px";

            ApplyPaletteSelection(document.SelectedPaletteKey, updateDocument: false, resetStickerCounter: false);

            if (document.ImageBytes is { Length: > 0 })
            {
                ImportedImage.Source = await CreateBitmapImageAsync(document.ImageBytes);
                ImportedImage.Visibility = Visibility.Visible;
                PlaceholderPanel.Visibility = Visibility.Collapsed;
            }
            else
            {
                ImportedImage.Source = null;
                ImportedImage.Visibility = Visibility.Collapsed;
                PlaceholderPanel.Visibility = Visibility.Visible;
            }

            ApplyCropState(document);
            RenderAnnotations(document);
            UpdateStickerSequenceText(document);
            UpdateAnnotationCount(document);
            UpdateStatus();
        }
        finally
        {
            _isApplyingDocumentState = false;
        }
    }

    private static async Task<BitmapImage> CreateBitmapImageAsync(byte[] bytes)
    {
        using var stream = new InMemoryRandomAccessStream();
        await stream.WriteAsync(bytes.AsBuffer());
        stream.Seek(0);

        var bitmap = new BitmapImage();
        await bitmap.SetSourceAsync(stream);
        return bitmap;
    }

    private void SelectTool(EditorTool tool)
    {
        _currentTool = tool;
        foreach (var button in _toolButtons)
        {
            var isSelected = string.Equals(button.Tag?.ToString(), tool.ToString(), StringComparison.Ordinal);
            button.Background = GetBrush(isSelected ? "ToolSelectedBrush" : "ShellPanelAltBrush");
            button.Foreground = GetBrush(isSelected ? "ToolSelectedTextBrush" : "ToolUnselectedTextBrush");
        }

        CurrentToolText.Text = _toolLabels[tool];
        UpdateStatus();
    }

    private void ToolButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button button && Enum.TryParse(button.Tag?.ToString(), out EditorTool tool))
        {
            SelectTool(tool);
        }
    }

    private void PaletteButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentDocument is null || sender is not Button button || button.Tag is not string key)
        {
            return;
        }

        var resetCounter = _currentDocument.ResetStickerNumberOnColorChange &&
            !string.Equals(_currentDocument.SelectedPaletteKey, key, StringComparison.Ordinal);

        ApplyPaletteSelection(key, updateDocument: true, resetStickerCounter: resetCounter);
        UpdateStickerSequenceText(_currentDocument);
        UpdateStatus();
    }

    private void ApplyPaletteSelection(string key, bool updateDocument, bool resetStickerCounter)
    {
        if (!_palettes.TryGetValue(key, out var palette))
        {
            return;
        }

        foreach (var pair in _paletteButtons)
        {
            var isSelected = string.Equals(pair.Key, key, StringComparison.Ordinal);
            pair.Value.BorderThickness = new Thickness(isSelected ? 3 : 1);
            pair.Value.BorderBrush = GetBrush(isSelected ? "ToolSelectedBrush" : "ShellStrokeBrush");
        }

        PaletteSummaryText.Text = $"Palette active : {palette.DisplayName}";

        if (_currentDocument is not null && updateDocument)
        {
            _currentDocument.SelectedPaletteKey = key;
            _currentDocument.PaletteDisplayName = palette.DisplayName;

            if (resetStickerCounter)
            {
                _currentDocument.NextStickerIndex = 1;
            }
        }
    }

    private void AnnotationTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isApplyingDocumentState || _currentDocument is null)
        {
            return;
        }

        _currentDocument.AnnotationText = AnnotationTextBox.Text;
        UpdateStickerSequenceText(_currentDocument);
    }

    private void StickerModeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isApplyingDocumentState || _currentDocument is null)
        {
            return;
        }

        _currentDocument.StickerModeIndex = StickerModeComboBox.SelectedIndex;
        UpdateStickerSequenceText(_currentDocument);
    }

    private void StrokeThicknessSlider_ValueChanged(object sender, RangeBaseValueChangedEventArgs e)
    {
        if (_currentDocument is not null)
        {
            _currentDocument.StrokeThickness = StrokeThicknessSlider.Value;
        }

        StrokeSummaryText.Text = $"Trait actuel : {StrokeThicknessSlider.Value:0} px";
    }

    private void ResetStickerNumberOnColorChangeCheckBox_Changed(object sender, RoutedEventArgs e)
    {
        if (_isApplyingDocumentState || _currentDocument is null)
        {
            return;
        }

        _currentDocument.ResetStickerNumberOnColorChange = ResetStickerNumberOnColorChangeCheckBox.IsChecked == true;
    }

    private void UpdateStickerSequenceText(ScreenshotDocument document)
    {
        var nextLabel = document.StickerModeIndex switch
        {
            1 => ToAlphabetic(document.NextStickerIndex),
            2 when !string.IsNullOrWhiteSpace(document.AnnotationText) => document.AnnotationText.Trim(),
            _ => document.NextStickerIndex.ToString()
        };

        StickerSequenceText.Text = $"Prochaine gommette : {nextLabel}";
    }

    private static string ToAlphabetic(int index)
    {
        var normalized = Math.Max(1, index);
        return ((char)('A' + ((normalized - 1) % 26))).ToString();
    }

    private async void ImportScreenshotButton_Click(object sender, RoutedEventArgs e)
    {
        var picker = new FileOpenPicker();
        picker.FileTypeFilter.Add(".png");
        picker.FileTypeFilter.Add(".jpg");
        picker.FileTypeFilter.Add(".jpeg");
        picker.FileTypeFilter.Add(".bmp");
        picker.FileTypeFilter.Add(".webp");
        InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));

        var file = await picker.PickSingleFileAsync();
        if (file is null)
        {
            StatusText.Text = "Import annule.";
            return;
        }

        var document = await CreateDocumentFromFileAsync(file);
        if (document is null)
        {
            StatusText.Text = "Impossible de lire le fichier selectionne.";
            return;
        }

        AddDocument(document, select: true);
        StatusText.Text = $"Nouveau tab cree pour {file.Name}.";
    }

    private void NewDemoTabButton_Click(object sender, RoutedEventArgs e)
    {
        AddDocument(CreateDemoDocument(), select: true);
        StatusText.Text = "Nouvel onglet de demonstration cree.";
    }

    private async void CloseDocumentButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: ScreenshotDocument document })
        {
            await TryCloseDocumentAsync(document);
        }
    }

    private async Task TryCloseDocumentAsync(ScreenshotDocument document)
    {
        if (document.IsDirty)
        {
            var dialog = new ContentDialog
            {
                XamlRoot = RootGrid.XamlRoot,
                Title = $"Fermer {document.BaseTitle} ?",
                Content = "Cet onglet contient du contenu non sauvegarde. Voulez-vous enregistrer avant de fermer ?",
                PrimaryButtonText = "Enregistrer",
                SecondaryButtonText = "Fermer sans sauvegarder",
                CloseButtonText = "Annuler",
                DefaultButton = ContentDialogButton.Primary
            };

            var result = await dialog.ShowAsync();
            if (result == ContentDialogResult.None)
            {
                return;
            }

            if (result == ContentDialogResult.Primary)
            {
                var saved = await SaveDocumentAsync(document);
                if (!saved)
                {
                    return;
                }
            }
        }

        var previousSelection = _currentDocument;
        var removeIndex = _documents.IndexOf(document);
        _documents.Remove(document);

        if (_documents.Count == 0)
        {
            AddDocument(CreateDemoDocument(isInitial: true), select: true);
            return;
        }

        if (previousSelection is not null && previousSelection != document && _documents.Contains(previousSelection))
        {
            DocumentTabsListView.SelectedItem = previousSelection;
            return;
        }

        var nextIndex = Math.Clamp(removeIndex, 0, _documents.Count - 1);
        DocumentTabsListView.SelectedItem = _documents[nextIndex];
    }

    private async void ExportButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentDocument is null)
        {
            return;
        }

        await SaveDocumentAsync(_currentDocument);
    }

    private async Task<bool> SaveDocumentAsync(ScreenshotDocument document)
    {
        await LoadDocumentAsync(document);

        StorageFile? file = null;
        if (!string.IsNullOrWhiteSpace(document.SavedPath))
        {
            try
            {
                file = await StorageFile.GetFileFromPathAsync(document.SavedPath);
            }
            catch
            {
                file = null;
            }
        }

        if (file is null)
        {
            var picker = new FileSavePicker();
            picker.FileTypeChoices.Add("PNG image", [".png"]);
            picker.SuggestedFileName = document.Origin == DocumentOrigin.FileImport
                ? $"{System.IO.Path.GetFileNameWithoutExtension(document.FileNameLabel)}-annotated"
                : document.BaseTitle.Replace(' ', '-').ToLowerInvariant();

            InitializeWithWindow.Initialize(picker, WindowNative.GetWindowHandle(this));
            file = await picker.PickSaveFileAsync();
            if (file is null)
            {
                StatusText.Text = "Sauvegarde annulee.";
                return false;
            }
        }

        var renderTarget = new RenderTargetBitmap();
        var width = Math.Max(1, (int)Math.Round(PreviewViewport.Width));
        var height = Math.Max(1, (int)Math.Round(PreviewViewport.Height));
        await renderTarget.RenderAsync(PreviewViewport, width, height);
        var pixels = (await renderTarget.GetPixelsAsync()).ToArray();

        using var stream = await file.OpenAsync(FileAccessMode.ReadWrite);
        var encoder = await BitmapEncoder.CreateAsync(BitmapEncoder.PngEncoderId, stream);
        encoder.SetPixelData(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied, (uint)width, (uint)height, 96, 96, pixels);
        await encoder.FlushAsync();

        document.SavedPath = file.Path;
        document.IsDirty = false;
        StatusText.Text = $"Sauvegarde terminee : {file.Path}";
        return true;
    }

    private void ClearAnnotationsButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentDocument is null || _currentDocument.Annotations.Count == 0)
        {
            return;
        }

        _currentDocument.Annotations.Clear();
        MarkDocumentDirty(_currentDocument);
        RenderAnnotations(_currentDocument);
        UpdateAnnotationCount(_currentDocument);
        StatusText.Text = "Les annotations ont ete effacees.";
    }

    private void ApplyCropButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentDocument?.PendingCropRect is not Rect crop || crop.Width < 8 || crop.Height < 8)
        {
            StatusText.Text = "Tracez d'abord une vraie zone de crop.";
            return;
        }

        _currentDocument.AppliedCropRect = crop;
        _currentDocument.PendingCropRect = null;
        MarkDocumentDirty(_currentDocument);
        ApplyCropState(_currentDocument);
        StatusText.Text = "Le crop est applique a cet onglet.";
    }

    private void ResetCropButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentDocument is null)
        {
            return;
        }

        var hadAppliedCrop = _currentDocument.AppliedCropRect is not null;
        _currentDocument.PendingCropRect = null;
        _currentDocument.AppliedCropRect = null;
        if (hadAppliedCrop)
        {
            MarkDocumentDirty(_currentDocument);
        }

        ApplyCropState(_currentDocument);
        StatusText.Text = "Le crop a ete reinitialise.";
    }

    private void ApplyCropState(ScreenshotDocument document)
    {
        if (document.AppliedCropRect is Rect appliedCrop)
        {
            PreviewViewport.Width = appliedCrop.Width;
            PreviewViewport.Height = appliedCrop.Height;
            SceneTransform.X = -appliedCrop.X;
            SceneTransform.Y = -appliedCrop.Y;
            UpdateViewportClip(appliedCrop.Width, appliedCrop.Height);
            CropStateText.Text = $"Crop applique : {appliedCrop.Width:0} x {appliedCrop.Height:0}px.";
            HideCropOverlay();
            ResetCropButton.IsEnabled = true;
            ApplyCropButton.IsEnabled = false;
            return;
        }

        PreviewViewport.Width = SceneWidth;
        PreviewViewport.Height = SceneHeight;
        SceneTransform.X = 0;
        SceneTransform.Y = 0;
        UpdateViewportClip(SceneWidth, SceneHeight);
        CropStateText.Text = document.PendingCropRect is Rect pendingCrop
            ? $"Crop en attente : {pendingCrop.Width:0} x {pendingCrop.Height:0}px."
            : "Aucun crop applique pour le moment.";

        if (document.PendingCropRect is Rect overlayCrop)
        {
            ShowCropOverlay(overlayCrop);
        }
        else
        {
            HideCropOverlay();
        }

        ResetCropButton.IsEnabled = document.PendingCropRect is not null || document.AppliedCropRect is not null;
        ApplyCropButton.IsEnabled = document.PendingCropRect is not null;
    }

    private void UpdateViewportClip(double width, double height)
    {
        if (PreviewViewport.Clip is RectangleGeometry clip)
        {
            clip.Rect = new Rect(0, 0, width, height);
        }
    }

    private void InteractionSurface_PointerPressed(object sender, PointerRoutedEventArgs e)
    {
        if (_currentDocument is null)
        {
            return;
        }

        var point = e.GetCurrentPoint(SceneRoot).Position;
        if (point.X < 0 || point.Y < 0 || point.X > SceneWidth || point.Y > SceneHeight)
        {
            return;
        }

        if (_currentTool == EditorTool.Text)
        {
            AddTextAnnotation(_currentDocument, point);
            return;
        }

        if (_currentTool == EditorTool.Sticker)
        {
            AddStickerAnnotation(_currentDocument, point);
            return;
        }

        _dragStart = point;
        _isDragging = true;
        InteractionSurface.CapturePointer(e.Pointer);

        if (_currentTool == EditorTool.Crop)
        {
            ShowCropOverlay(new Rect(point.X, point.Y, 0, 0));
            return;
        }

        _activeElement = CreateTemporaryElement(_currentDocument);
        if (_activeElement is not null)
        {
            AnnotationCanvas.Children.Add(_activeElement);
        }
    }

    private void InteractionSurface_PointerMoved(object sender, PointerRoutedEventArgs e)
    {
        if (!_isDragging || _currentDocument is null)
        {
            return;
        }

        var point = e.GetCurrentPoint(SceneRoot).Position;
        if (_currentTool == EditorTool.Crop)
        {
            ShowCropOverlay(NormalizeRect(_dragStart, point));
            return;
        }

        UpdateTemporaryElement(point);
    }

    private void InteractionSurface_PointerReleased(object sender, PointerRoutedEventArgs e)
    {
        if (!_isDragging || _currentDocument is null)
        {
            return;
        }

        var point = e.GetCurrentPoint(SceneRoot).Position;

        if (_currentTool == EditorTool.Crop)
        {
            var crop = NormalizeRect(_dragStart, point);
            if (crop.Width >= 8 && crop.Height >= 8)
            {
                _currentDocument.PendingCropRect = crop;
                _currentDocument.AppliedCropRect = null;
                ApplyCropState(_currentDocument);
                StatusText.Text = "Crop capture. Utilisez 'Appliquer crop'.";
            }
            else
            {
                _currentDocument.PendingCropRect = null;
                ApplyCropState(_currentDocument);
                StatusText.Text = "Le crop est trop petit.";
            }

            FinishDrag(e);
            return;
        }

        FinalizeDraggedAnnotation(_currentDocument, point);
        FinishDrag(e);
    }

    private FrameworkElement? CreateTemporaryElement(ScreenshotDocument document)
    {
        var stroke = CreateGradientBrush(document.SelectedPaletteKey);
        return _currentTool switch
        {
            EditorTool.Rectangle => new Rectangle
            {
                Stroke = stroke,
                Fill = CreateHighlightFill(document.SelectedPaletteKey, 36),
                RadiusX = 24,
                RadiusY = 24,
                StrokeThickness = document.StrokeThickness
            },
            EditorTool.Ellipse => new Ellipse
            {
                Stroke = stroke,
                Fill = CreateHighlightFill(document.SelectedPaletteKey, 30),
                StrokeThickness = document.StrokeThickness
            },
            EditorTool.ArrowStraight => CreateArrowPath(stroke, document.StrokeThickness),
            EditorTool.ArrowCurved => CreateArrowPath(stroke, document.StrokeThickness),
            _ => null
        };
    }

    private static XamlPath CreateArrowPath(Brush stroke, double strokeThickness)
    {
        return new XamlPath
        {
            Stroke = stroke,
            StrokeThickness = strokeThickness,
            StrokeLineJoin = PenLineJoin.Round,
            StrokeStartLineCap = PenLineCap.Round,
            StrokeEndLineCap = PenLineCap.Round
        };
    }

    private void UpdateTemporaryElement(Point point)
    {
        if (_activeElement is null)
        {
            return;
        }

        switch (_activeElement)
        {
            case Rectangle rectangle:
                UpdateShapeBounds(rectangle, NormalizeRect(_dragStart, point));
                break;
            case Ellipse ellipse:
                UpdateShapeBounds(ellipse, NormalizeRect(_dragStart, point));
                break;
            case XamlPath path when _currentTool == EditorTool.ArrowStraight:
                path.Data = BuildStraightArrowGeometry(_dragStart, point, path.StrokeThickness);
                break;
            case XamlPath path when _currentTool == EditorTool.ArrowCurved:
                path.Data = BuildCurvedArrowGeometry(_dragStart, point, path.StrokeThickness);
                break;
        }
    }

    private static void UpdateShapeBounds(Shape shape, Rect bounds)
    {
        shape.Width = bounds.Width;
        shape.Height = bounds.Height;
        Canvas.SetLeft(shape, bounds.X);
        Canvas.SetTop(shape, bounds.Y);
    }

    private void FinalizeDraggedAnnotation(ScreenshotDocument document, Point point)
    {
        if (_activeElement is not null)
        {
            AnnotationCanvas.Children.Remove(_activeElement);
            _activeElement = null;
        }

        switch (_currentTool)
        {
            case EditorTool.Rectangle:
            {
                var bounds = NormalizeRect(_dragStart, point);
                if (bounds.Width < 8 || bounds.Height < 8)
                {
                    return;
                }

                document.Annotations.Add(new AnnotationModel
                {
                    Kind = AnnotationKind.Rectangle,
                    Bounds = bounds,
                    PaletteKey = document.SelectedPaletteKey,
                    StrokeThickness = document.StrokeThickness
                });
                break;
            }
            case EditorTool.Ellipse:
            {
                var bounds = NormalizeRect(_dragStart, point);
                if (bounds.Width < 8 || bounds.Height < 8)
                {
                    return;
                }

                document.Annotations.Add(new AnnotationModel
                {
                    Kind = AnnotationKind.Ellipse,
                    Bounds = bounds,
                    PaletteKey = document.SelectedPaletteKey,
                    StrokeThickness = document.StrokeThickness
                });
                break;
            }
            case EditorTool.ArrowStraight:
            {
                if (Distance(_dragStart, point) < 10)
                {
                    return;
                }

                document.Annotations.Add(new AnnotationModel
                {
                    Kind = AnnotationKind.ArrowStraight,
                    StartPoint = _dragStart,
                    EndPoint = point,
                    PaletteKey = document.SelectedPaletteKey,
                    StrokeThickness = document.StrokeThickness
                });
                break;
            }
            case EditorTool.ArrowCurved:
            {
                if (Distance(_dragStart, point) < 10)
                {
                    return;
                }

                document.Annotations.Add(new AnnotationModel
                {
                    Kind = AnnotationKind.ArrowCurved,
                    StartPoint = _dragStart,
                    EndPoint = point,
                    PaletteKey = document.SelectedPaletteKey,
                    StrokeThickness = document.StrokeThickness
                });
                break;
            }
        }

        MarkDocumentDirty(document);
        RenderAnnotations(document);
        UpdateAnnotationCount(document);
        UpdateStatus();
    }

    private void AddTextAnnotation(ScreenshotDocument document, Point point)
    {
        document.Annotations.Add(new AnnotationModel
        {
            Kind = AnnotationKind.Text,
            Bounds = new Rect(point.X, point.Y, 0, 0),
            PaletteKey = document.SelectedPaletteKey,
            Text = string.IsNullOrWhiteSpace(document.AnnotationText) ? "Note rapide" : document.AnnotationText.Trim(),
            StrokeThickness = document.StrokeThickness
        });

        MarkDocumentDirty(document);
        RenderAnnotations(document);
        UpdateAnnotationCount(document);
        UpdateStatus();
    }

    private void AddStickerAnnotation(ScreenshotDocument document, Point point)
    {
        document.Annotations.Add(new AnnotationModel
        {
            Kind = AnnotationKind.Sticker,
            Bounds = new Rect(point.X - (StickerSize / 2), point.Y - (StickerSize / 2), StickerSize, StickerSize),
            PaletteKey = document.SelectedPaletteKey,
            Text = BuildStickerLabel(document)
        });

        MarkDocumentDirty(document);
        RenderAnnotations(document);
        UpdateAnnotationCount(document);
        UpdateStickerSequenceText(document);
        UpdateStatus();
    }

    private static double Distance(Point start, Point end)
    {
        return Math.Sqrt(Math.Pow(end.X - start.X, 2) + Math.Pow(end.Y - start.Y, 2));
    }

    private string BuildStickerLabel(ScreenshotDocument document)
    {
        return document.StickerModeIndex switch
        {
            1 => AdvanceAlphabeticSticker(document),
            2 when !string.IsNullOrWhiteSpace(document.AnnotationText) => document.AnnotationText.Trim(),
            _ => AdvanceNumericSticker(document)
        };
    }

    private static string AdvanceNumericSticker(ScreenshotDocument document)
    {
        var label = document.NextStickerIndex.ToString();
        document.NextStickerIndex++;
        return label;
    }

    private static string AdvanceAlphabeticSticker(ScreenshotDocument document)
    {
        var label = ToAlphabetic(document.NextStickerIndex);
        document.NextStickerIndex++;
        return label;
    }

    private void RenderAnnotations(ScreenshotDocument document)
    {
        AnnotationCanvas.Children.Clear();
        foreach (var annotation in document.Annotations)
        {
            switch (annotation.Kind)
            {
                case AnnotationKind.Text:
                {
                    var border = new Border
                    {
                        Background = new SolidColorBrush(ColorHelper.FromArgb(235, 255, 255, 255)),
                        BorderBrush = CreateGradientBrush(annotation.PaletteKey),
                        BorderThickness = new Thickness(2),
                        CornerRadius = new CornerRadius(14),
                        Padding = new Thickness(14, 10, 14, 10),
                        Child = new TextBlock
                        {
                            Text = annotation.Text,
                            FontFamily = new FontFamily("Bahnschrift SemiBold SemiConden"),
                            FontSize = 28,
                            Foreground = CreateGradientBrush(annotation.PaletteKey)
                        }
                    };

                    AnnotationCanvas.Children.Add(border);
                    Canvas.SetLeft(border, annotation.Bounds.X);
                    Canvas.SetTop(border, annotation.Bounds.Y);
                    break;
                }
                case AnnotationKind.Sticker:
                {
                    var sticker = new Border
                    {
                        Width = annotation.Bounds.Width,
                        Height = annotation.Bounds.Height,
                        Background = CreateGradientBrush(annotation.PaletteKey),
                        BorderBrush = new SolidColorBrush(Colors.White),
                        BorderThickness = new Thickness(3),
                        CornerRadius = new CornerRadius(annotation.Bounds.Width / 2),
                        Child = new TextBlock
                        {
                            Text = annotation.Text,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center,
                            FontFamily = new FontFamily("Bahnschrift SemiBold SemiConden"),
                            FontSize = 28,
                            Foreground = new SolidColorBrush(Colors.White)
                        }
                    };

                    AnnotationCanvas.Children.Add(sticker);
                    Canvas.SetLeft(sticker, annotation.Bounds.X);
                    Canvas.SetTop(sticker, annotation.Bounds.Y);
                    break;
                }
                case AnnotationKind.Rectangle:
                {
                    var rectangle = new Rectangle
                    {
                        Width = annotation.Bounds.Width,
                        Height = annotation.Bounds.Height,
                        RadiusX = 24,
                        RadiusY = 24,
                        Stroke = CreateGradientBrush(annotation.PaletteKey),
                        Fill = CreateHighlightFill(annotation.PaletteKey, 36),
                        StrokeThickness = annotation.StrokeThickness
                    };

                    AnnotationCanvas.Children.Add(rectangle);
                    Canvas.SetLeft(rectangle, annotation.Bounds.X);
                    Canvas.SetTop(rectangle, annotation.Bounds.Y);
                    break;
                }
                case AnnotationKind.Ellipse:
                {
                    var ellipse = new Ellipse
                    {
                        Width = annotation.Bounds.Width,
                        Height = annotation.Bounds.Height,
                        Stroke = CreateGradientBrush(annotation.PaletteKey),
                        Fill = CreateHighlightFill(annotation.PaletteKey, 30),
                        StrokeThickness = annotation.StrokeThickness
                    };

                    AnnotationCanvas.Children.Add(ellipse);
                    Canvas.SetLeft(ellipse, annotation.Bounds.X);
                    Canvas.SetTop(ellipse, annotation.Bounds.Y);
                    break;
                }
                case AnnotationKind.ArrowStraight:
                {
                    AnnotationCanvas.Children.Add(new XamlPath
                    {
                        Stroke = CreateGradientBrush(annotation.PaletteKey),
                        StrokeThickness = annotation.StrokeThickness,
                        StrokeLineJoin = PenLineJoin.Round,
                        StrokeStartLineCap = PenLineCap.Round,
                        StrokeEndLineCap = PenLineCap.Round,
                        Data = BuildStraightArrowGeometry(annotation.StartPoint, annotation.EndPoint, annotation.StrokeThickness)
                    });
                    break;
                }
                case AnnotationKind.ArrowCurved:
                {
                    AnnotationCanvas.Children.Add(new XamlPath
                    {
                        Stroke = CreateGradientBrush(annotation.PaletteKey),
                        StrokeThickness = annotation.StrokeThickness,
                        StrokeLineJoin = PenLineJoin.Round,
                        StrokeStartLineCap = PenLineCap.Round,
                        StrokeEndLineCap = PenLineCap.Round,
                        Data = BuildCurvedArrowGeometry(annotation.StartPoint, annotation.EndPoint, annotation.StrokeThickness)
                    });
                    break;
                }
            }
        }
    }

    private Brush CreateGradientBrush(string paletteKey)
    {
        var palette = _palettes[paletteKey];
        var brush = new LinearGradientBrush
        {
            StartPoint = new Point(0, 0),
            EndPoint = new Point(1, 1)
        };
        brush.GradientStops.Add(new GradientStop { Color = palette.StartColor, Offset = 0 });
        brush.GradientStops.Add(new GradientStop { Color = palette.EndColor, Offset = 1 });
        return brush;
    }

    private Brush CreateHighlightFill(string paletteKey, byte alpha)
    {
        var color = _palettes[paletteKey].StartColor;
        return new SolidColorBrush(ColorHelper.FromArgb(alpha, color.R, color.G, color.B));
    }

    private Geometry BuildStraightArrowGeometry(Point start, Point end, double strokeThickness)
    {
        var group = new GeometryGroup();
        group.Children.Add(BuildLineGeometry(start, end));
        AppendArrowHead(group, end, Math.Atan2(end.Y - start.Y, end.X - start.X), strokeThickness);
        return group;
    }

    private Geometry BuildCurvedArrowGeometry(Point start, Point end, double strokeThickness)
    {
        var group = new GeometryGroup();
        var controlPoint1 = new Point(start.X + ((end.X - start.X) * 0.3), start.Y - 80);
        var controlPoint2 = new Point(start.X + ((end.X - start.X) * 0.82), end.Y - 24);
        var figure = new PathFigure { StartPoint = start, IsClosed = false, IsFilled = false };
        figure.Segments.Add(new BezierSegment
        {
            Point1 = controlPoint1,
            Point2 = controlPoint2,
            Point3 = end
        });
        group.Children.Add(new PathGeometry { Figures = { figure } });
        AppendArrowHead(group, end, Math.Atan2(end.Y - controlPoint2.Y, end.X - controlPoint2.X), strokeThickness);
        return group;
    }

    private void AppendArrowHead(GeometryGroup group, Point end, double angle, double strokeThickness)
    {
        var length = Math.Max(18, strokeThickness * 4);
        var spread = Math.PI / 7;
        var left = new Point(end.X - (length * Math.Cos(angle - spread)), end.Y - (length * Math.Sin(angle - spread)));
        var right = new Point(end.X - (length * Math.Cos(angle + spread)), end.Y - (length * Math.Sin(angle + spread)));
        group.Children.Add(BuildLineGeometry(end, left));
        group.Children.Add(BuildLineGeometry(end, right));
    }

    private static Geometry BuildLineGeometry(Point start, Point end)
    {
        var figure = new PathFigure { StartPoint = start, IsClosed = false, IsFilled = false };
        figure.Segments.Add(new LineSegment { Point = end });
        return new PathGeometry { Figures = { figure } };
    }

    private void ShowCropOverlay(Rect crop)
    {
        CropPreview.Visibility = Visibility.Visible;
        CropBadge.Visibility = Visibility.Visible;
        CropPreview.Width = crop.Width;
        CropPreview.Height = crop.Height;
        Canvas.SetLeft(CropPreview, crop.X);
        Canvas.SetTop(CropPreview, crop.Y);
        CropBadgeText.Text = $"Crop {crop.Width:0} x {crop.Height:0}";
        Canvas.SetLeft(CropBadge, crop.X + 12);
        Canvas.SetTop(CropBadge, Math.Max(12, crop.Y + 12));
    }

    private void HideCropOverlay()
    {
        CropPreview.Visibility = Visibility.Collapsed;
        CropBadge.Visibility = Visibility.Collapsed;
    }

    private static Rect NormalizeRect(Point start, Point end)
    {
        var x = Math.Min(start.X, end.X);
        var y = Math.Min(start.Y, end.Y);
        return new Rect(x, y, Math.Abs(end.X - start.X), Math.Abs(end.Y - start.Y));
    }

    private void FinishDrag(PointerRoutedEventArgs e)
    {
        _isDragging = false;
        _activeElement = null;
        InteractionSurface.ReleasePointerCapture(e.Pointer);
    }

    private void MarkDocumentDirty(ScreenshotDocument document)
    {
        document.IsDirty = true;
    }

    private void UpdateAnnotationCount(ScreenshotDocument document)
    {
        AnnotationCountText.Text = $"Annotations : {document.Annotations.Count}";
    }

    private void UpdateStatus()
    {
        if (_currentDocument is null)
        {
            StatusText.Text = "Aucun document selectionne.";
            return;
        }

        StatusText.Text = $"{_toolHints[_currentTool]} Palette active : {_currentDocument.PaletteDisplayName}. Onglets ouverts : {_documents.Count}.";
        FooterText.Text = "Chaque capture Win + Shift + S cree un nouvel onglet.";
    }

    private SolidColorBrush GetBrush(string key)
    {
        return (SolidColorBrush)Application.Current.Resources[key];
    }

    private async void Clipboard_ContentChanged(object? sender, object e)
    {
        if (_isProcessingClipboard)
        {
            return;
        }

        DispatcherQueue.TryEnqueue(async () =>
        {
            if (_isProcessingClipboard)
            {
                return;
            }

            _isProcessingClipboard = true;
            try
            {
                await TryImportClipboardCaptureAsync();
            }
            finally
            {
                _isProcessingClipboard = false;
            }
        });
    }

    private async Task TryImportClipboardCaptureAsync()
    {
        try
        {
            var dataPackage = Clipboard.GetContent();
            if (!dataPackage.Contains(StandardDataFormats.Bitmap))
            {
                return;
            }

            var streamReference = await dataPackage.GetBitmapAsync();
            using var stream = await streamReference.OpenReadAsync();
            using var memory = new MemoryStream();
            await stream.AsStreamForRead().CopyToAsync(memory);

            var bytes = memory.ToArray();
            if (bytes.Length == 0)
            {
                return;
            }

            var hash = Convert.ToHexString(SHA256.HashData(bytes));
            var now = DateTimeOffset.Now;
            if (string.Equals(hash, _lastClipboardHash, StringComparison.Ordinal) &&
                (now - _lastClipboardImportAt) < TimeSpan.FromSeconds(2))
            {
                return;
            }

            _lastClipboardHash = hash;
            _lastClipboardImportAt = now;

            AddDocument(CreateDocumentFromClipboard(bytes), select: true);
            StatusText.Text = "Nouvelle capture Win + Shift + S importee dans un nouvel onglet.";
        }
        catch
        {
            StatusText.Text = "Impossible de lire la derniere capture depuis le presse-papiers.";
        }
    }
}
