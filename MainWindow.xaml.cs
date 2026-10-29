using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Microsoft.UI;
using Microsoft.UI.Dispatching;
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
    private const double FloatingToolRailGutterWidth = 0;
    private const double StickerSize = 60;
    private const double TextAnnotationMinWidth = 220;
    private const double TextAnnotationMaxWidth = 760;
    private const double TextAnnotationMinHeight = 96;
    private const double TextAnnotationMaxHeight = 520;

    private readonly ObservableCollection<ScreenshotDocument> _documents = [];
    private readonly Dictionary<string, GradientPaletteDefinition> _palettes;
    private readonly ObservableCollection<GradientPaletteDefinition> _paletteOrder = [];
    private readonly Dictionary<EditorTool, string> _toolLabels = new()
    {
        [EditorTool.Select] = "ToolSelect",
        [EditorTool.Crop] = "ToolCrop",
        [EditorTool.Text] = "ToolText",
        [EditorTool.Sticker] = "ToolSticker",
        [EditorTool.ArrowStraight] = "ToolArrowStraight",
        [EditorTool.ArrowCurved] = "ToolArrowCurved",
        [EditorTool.Rectangle] = "ToolRectangle",
        [EditorTool.Ellipse] = "ToolEllipse",
        [EditorTool.Focus] = "ToolFocus",
        [EditorTool.Mask] = "ToolMask"
    };

    private readonly Dictionary<EditorTool, string> _toolHints = new()
    {
        [EditorTool.Select] = "HintSelect",
        [EditorTool.Crop] = "HintCrop",
        [EditorTool.Text] = "HintText",
        [EditorTool.Sticker] = "HintSticker",
        [EditorTool.ArrowStraight] = "HintArrowStraight",
        [EditorTool.ArrowCurved] = "HintArrowCurved",
        [EditorTool.Rectangle] = "HintRectangle",
        [EditorTool.Ellipse] = "HintEllipse",
        [EditorTool.Focus] = "HintFocus",
        [EditorTool.Mask] = "HintMask"
    };

    private readonly List<Button> _toolButtons;
    private readonly Dictionary<string, Button> _paletteButtons;
    private readonly List<Button> _shadeButtons = [];
    private EditorTool _currentTool = EditorTool.ArrowStraight;
    private ScreenshotDocument? _currentDocument;
    private AnnotationModel? _selectedAnnotation;
    private AnnotationModel? _editingTextAnnotation;
    private FrameworkElement? _activeElement;
    private AnnotationModel? _draggingAnnotation;
    private FrameworkElement? _draggingAnnotationElement;
    private Point _dragStart;
    private Point _annotationDragStart;
    private Rect _annotationDragStartBounds;
    private Point _annotationDragStartPoint;
    private Point _annotationDragStartEndPoint;
    private bool _isDragging;
    private bool _isDraggingAnnotation;
    private bool _isApplyingDocumentState;
    private bool _isSyncingAnnotationSelectionUi;
    private bool _isSyncingPaletteUi;
    private bool _isSyncingOpacityUi;
    private bool _isInitializingUi;
    private bool _isInspectorPaneCollapsed = true;
    private bool _isProcessingClipboard;
    private bool _clipboardWatcherPrimed;
    private int _nextCaptureIndex = 1;
    private int _nextDemoIndex = 1;
    private double _currentZoom = 1.0;
    private ShellSection _currentSection = ShellSection.Procedure;

    public MainWindow()
    {
        InitializeComponent();

        ExtendsContentIntoTitleBar = true;
        SetTitleBar(AppTitleBar);

        AppWindow.SetIcon(System.IO.Path.Combine(AppContext.BaseDirectory, "Assets", "AppIcon.ico"));
        var workArea = DisplayArea.GetFromWindowId(AppWindow.Id, DisplayAreaFallback.Primary).WorkArea;
        var width = Math.Min(1440, workArea.Width);
        var height = Math.Min(900, workArea.Height);
        AppWindow.MoveAndResize(new RectInt32(workArea.X + (workArea.Width - width) / 2, workArea.Y + (workArea.Height - height) / 2, width, height));
        AppWindow.TitleBar.PreferredHeightOption = TitleBarHeightOption.Standard;

        _palettes = CreatePalettes();
        _toolButtons =
        [
            ToolbarSelectButton,
            ToolbarCropButton,
            ToolbarTextButton,
            ToolbarStickerButton,
            ToolbarArrowStraightButton,
            ToolbarArrowCurvedButton,
            ToolbarRectangleButton,
            ToolbarEllipseButton,
            ToolbarFocusButton,
            ToolbarMaskButton
        ];
        _paletteButtons = new Dictionary<string, Button>(StringComparer.Ordinal);

        PreviewViewport.Clip = new RectangleGeometry();
        DocumentTabsListView.ItemsSource = _documents;
        ToolbarDocumentComboBox.ItemsSource = _documents;
        ToolbarPaletteComboBox.ItemsSource = _paletteOrder;
        StrokeThicknessSlider.Minimum = 2;
        StrokeThicknessSlider.Maximum = 14;
        StrokeThicknessSlider.Value = 6;

        InitializeShellUi();
        InitializeProjectCommands();
        BuildPaletteStrip();
        BuildShadeStrip();
        SelectTool(EditorTool.Select);
        AddDocument(CreateDemoDocument(isInitial: true), select: true);
        MarkProjectClean();
        AppWindow.Closing += MainWindow_Closing;

        Closed += MainWindow_Closed;

        Clipboard.ContentChanged += Clipboard_ContentChanged;
        _ = PrimeClipboardWatcherAsync();
    }

    private static Dictionary<string, GradientPaletteDefinition> CreatePalettes()
    {
        return new Dictionary<string, GradientPaletteDefinition>(StringComparer.Ordinal)
        {
            ["Sunset"] = CreatePalette("Sunset", "sunset note", ColorHelper.FromArgb(255, 240, 90, 65), ColorHelper.FromArgb(255, 255, 198, 110)),
            ["Ember"] = CreatePalette("Ember", "ember alert", ColorHelper.FromArgb(255, 189, 60, 43), ColorHelper.FromArgb(255, 240, 107, 89)),
            ["Citrus"] = CreatePalette("Citrus", "citrus focus", ColorHelper.FromArgb(255, 240, 167, 59), ColorHelper.FromArgb(255, 246, 228, 122)),
            ["Mint"] = CreatePalette("Mint", "mint product", ColorHelper.FromArgb(255, 22, 179, 142), ColorHelper.FromArgb(255, 132, 226, 197)),
            ["Lagoon"] = CreatePalette("Lagoon", "lagoon flow", ColorHelper.FromArgb(255, 15, 158, 175), ColorHelper.FromArgb(255, 115, 225, 233)),
            ["Sky"] = CreatePalette("Sky", "sky workflow", ColorHelper.FromArgb(255, 31, 132, 224), ColorHelper.FromArgb(255, 140, 203, 255)),
            ["Ocean"] = CreatePalette("Ocean", "ocean depth", ColorHelper.FromArgb(255, 23, 58, 115), ColorHelper.FromArgb(255, 70, 135, 224)),
            ["Rose"] = CreatePalette("Rose", "rose highlight", ColorHelper.FromArgb(255, 216, 90, 136), ColorHelper.FromArgb(255, 245, 160, 189)),
            ["Berry"] = CreatePalette("Berry", "berry spark", ColorHelper.FromArgb(255, 165, 63, 164), ColorHelper.FromArgb(255, 240, 123, 197)),
            ["Grape"] = CreatePalette("Grape", "grape accent", ColorHelper.FromArgb(255, 103, 70, 195), ColorHelper.FromArgb(255, 176, 137, 255))
        };
    }

    private static GradientPaletteDefinition CreatePalette(string key, string displayName, Color startColor, Color endColor)
    {
        return new GradientPaletteDefinition(key, displayName, startColor, endColor)
        {
            Shades = BuildShades(startColor, endColor)
        };
    }

    private static IReadOnlyList<Color> BuildShades(Color startColor, Color endColor)
    {
        var shades = new List<Color>(7);
        for (var i = 0; i < 7; i++)
        {
            var factor = i / 6.0;
            shades.Add(InterpolateColor(startColor, endColor, factor));
        }

        return shades;
    }

    private static Color InterpolateColor(Color start, Color end, double factor)
    {
        byte lerp(byte s, byte e) => (byte)Math.Round(s + ((e - s) * factor));
        return ColorHelper.FromArgb(
            255,
            lerp(start.R, end.R),
            lerp(start.G, end.G),
            lerp(start.B, end.B));
    }

    private ScreenshotDocument CreateDemoDocument(bool isInitial = false)
    {
        var title = isInitial ? "Lorem Ipsum" : $"Lorem Ipsum {_nextDemoIndex++:00}";
        var document = CreateDocumentCore(
            title,
            DocumentOrigin.Demo,
            SafetyText("OriginDemo"),
            SafetyText("DemoSource"),
            "lorem-ipsum.png",
            null,
            (int)SceneWidth,
            (int)SceneHeight,
            new BitmapImage(new Uri("ms-appx:///Assets/Square150x150Logo.scale-200.png")),
            isDirty: false);

        document.StepNote = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Integer mattis, arcu in consequat fermentum, sapien metus consequat justo, et luctus arcu turpis in nisl.";
        document.AnnotationText = "Lorem ipsum dolor sit amet.";

        document.Annotations.Add(new AnnotationModel
        {
            Kind = AnnotationKind.Rectangle,
            Bounds = new Rect(420, 258, 468, 234),
            PaletteKey = "Sunset",
            StrokeThickness = 6
        });
        document.Annotations.Add(new AnnotationModel
        {
            Kind = AnnotationKind.ArrowCurved,
            StartPoint = new Point(980, 214),
            EndPoint = new Point(905, 314),
            PaletteKey = "Sky",
            StrokeThickness = 7
        });
        document.Annotations.Add(new AnnotationModel
        {
            Kind = AnnotationKind.Sticker,
            Bounds = new Rect(948, 164, StickerSize, StickerSize),
            PaletteKey = "Mint",
            Text = "A"
        });

        return document;
    }

    private void AddDocument(ScreenshotDocument document, bool select)
    {
        _documents.Add(document);
        RefreshDocumentOrderMetadata();

        if (select || _currentDocument is null)
        {
            DocumentTabsListView.SelectedItem = document;
        }
    }

    private async Task<ScreenshotDocument?> CreateDocumentFromFileAsync(StorageFile file)
    {
        if ((await file.GetBasicPropertiesAsync()).Size > ImageSafety.MaxImageBytes)
        {
            throw new InvalidDataException("Image file exceeds the supported limit.");
        }
        var buffer = await FileIO.ReadBufferAsync(file);
        var bytes = buffer.ToArray();
        if (bytes.Length == 0)
        {
            return null;
        }

        await Task.Run(() => { ImageSafety.Validate(bytes); ProjectContentFingerprint.WarmImage(bytes); });
        var (sourceWidth, sourceHeight) = await GetImageSizeAsync(bytes);
        var thumbnail = await CreateBitmapImageAsync(bytes, 160);
        return CreateDocumentCore(
            System.IO.Path.GetFileNameWithoutExtension(file.Name),
            DocumentOrigin.FileImport,
            SafetyText("OriginFileImport"),
            file.Path,
            file.Name,
            bytes,
            sourceWidth,
            sourceHeight,
            thumbnail,
            isDirty: false);
    }

    private async Task<ScreenshotDocument> CreateDocumentFromClipboardAsync(byte[] bytes)
    {
        await Task.Run(() => { ImageSafety.Validate(bytes); ProjectContentFingerprint.WarmImage(bytes); });
        var (sourceWidth, sourceHeight) = await GetImageSizeAsync(bytes);
        var captureIndex = _nextCaptureIndex++;
        var thumbnail = await CreateBitmapImageAsync(bytes, 160);
        return CreateDocumentCore(
            $"Capture {captureIndex:00}",
            DocumentOrigin.ClipboardCapture,
            SafetyText("OriginClipboardCapture"),
            SafetyText("SourceClipboard"),
            $"capture-{captureIndex:00}.png",
            bytes,
            sourceWidth,
            sourceHeight,
            thumbnail,
            isDirty: true);
    }

    private ScreenshotDocument CreateDocumentCore(
        string baseTitle,
        DocumentOrigin origin,
        string originLabel,
        string sourceLabel,
        string fileNameLabel,
        byte[]? imageBytes,
        int sourcePixelWidth,
        int sourcePixelHeight,
        BitmapImage? thumbnailSource,
        bool isDirty)
    {
        var palette = _palettes["Sunset"];
        return new ScreenshotDocument
        {
            BaseTitle = baseTitle,
            Origin = origin,
            OriginLabel = originLabel,
            SourceLabel = sourceLabel,
            FileNameLabel = fileNameLabel,
            SourcePixelWidth = sourcePixelWidth,
            SourcePixelHeight = sourcePixelHeight,
            ImageBytes = imageBytes,
            ThumbnailSource = thumbnailSource,
            SelectedPaletteKey = palette.Key,
            SelectedPaletteShadeIndex = 3,
            PaletteDisplayName = palette.DisplayName,
            AnnotationText = SafetyText("QuickNote"),
            StickerModeIndex = 0,
            StrokeThickness = 6,
            ResetStickerNumberOnColorChange = false,
            IsDirty = isDirty
        };
    }

    private async void DocumentTabsListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isApplyingProjectState)
        {
            return;
        }

        if (DocumentTabsListView.SelectedItem is ScreenshotDocument document)
        {
            await LoadDocumentAsync(document);
        }
    }

    private async Task LoadDocumentAsync(ScreenshotDocument document, BitmapImage? preparedImage = null)
    {
        if (_currentDocument == document && !_isApplyingDocumentState)
        {
            return;
        }

        var loadVersion = ++_documentLoadVersion;
        ++_previewVersion;
        RenderedSceneImage.Source = null;
        _currentDocument = document;
        _autoFitPreview = true;
        _isInspectorPaneCollapsed = true;
        _inspectorOpenedManually = false;
        _selectedAnnotation = null;
        _isApplyingDocumentState = true;

        try
        {
            CanvasSummaryText.Text = SafetyText("Origin" + document.Origin);
            CanvasFileNameText.Text = document.FileNameLabel;
            ImageSourceText.Text = string.Format(CultureInfo.CurrentCulture, SafetyText("SourceLabel"), document.SourceLabel);

            AnnotationTextBox.Text = document.AnnotationText;
            StickerModeComboBox.SelectedIndex = document.StickerModeIndex;
            StrokeThicknessSlider.Value = document.StrokeThickness;
            ResetStickerNumberOnColorChangeCheckBox.IsChecked = document.ResetStickerNumberOnColorChange;
            StrokeSummaryText.Text = string.Format(CultureInfo.CurrentCulture, SafetyText("StrokeCurrent"), document.StrokeThickness);

            ApplyPaletteSelection(document.SelectedPaletteKey, document.SelectedPaletteShadeIndex, updateDocument: false, updateSelectedAnnotation: false);
            SyncDocumentEditors(document);

            if (document.ImageBytes is { Length: > 0 })
            {
                var image = preparedImage ?? await CreateBitmapImageAsync(document.ImageBytes);
                if (loadVersion != _documentLoadVersion) return;
                ImportedImage.Source = image;
                ImportedImage.Visibility = Visibility.Visible;
                PlaceholderPanel.Visibility = Visibility.Collapsed;
            }
            else
            {
                ImportedImage.Source = null;
                ImportedImage.Visibility = Visibility.Collapsed;
                PlaceholderPanel.Visibility = Visibility.Visible;
            }

            ApplyDocumentSceneSize(document);
            ApplyCropState(document);
            RenderAnnotations(document);
            SelectAnnotation(null);
            UpdateStickerSequenceText(document);
            UpdateAnnotationCount(document);
            UpdateStatus();
        }
        finally
        {
            if (loadVersion == _documentLoadVersion) _isApplyingDocumentState = false;
        }
    }

    private static async Task<BitmapImage> CreateBitmapImageAsync(byte[] bytes, int? decodePixelWidth = null)
    {
        using var stream = new InMemoryRandomAccessStream();
        await stream.WriteAsync(bytes.AsBuffer());
        stream.Seek(0);

        var bitmap = new BitmapImage();
        if (decodePixelWidth is int width && width > 0)
        {
            bitmap.DecodePixelWidth = width;
        }
        await bitmap.SetSourceAsync(stream);
        return bitmap;
    }

    private static async Task<(int Width, int Height)> GetImageSizeAsync(byte[] bytes)
    {
        using var stream = new InMemoryRandomAccessStream();
        await stream.WriteAsync(bytes.AsBuffer());
        stream.Seek(0);

        try
        {
            var decoder = await BitmapDecoder.CreateAsync(stream);
            return ((int)decoder.PixelWidth, (int)decoder.PixelHeight);
        }
        catch
        {
            return ((int)SceneWidth, (int)SceneHeight);
        }
    }

    private void SelectTool(EditorTool tool)
    {
        if (tool != EditorTool.Select && tool != _currentTool && _selectedAnnotation is not null) SelectAnnotation(null);
        _currentTool = tool;
        foreach (var button in _toolButtons)
        {
            var isSelected = string.Equals(button.Tag?.ToString(), tool.ToString(), StringComparison.Ordinal);
            button.Background = GetBrush(isSelected ? "ToolSelectedBrush" : "ShellPanelBrush");
            button.BorderBrush = GetBrush(isSelected ? "ToolSelectedBrush" : "ShellStrokeBrush");
            button.BorderThickness = new Thickness(isSelected ? 2 : 1);
        }

        _iconBrushOverride = CreateToolAccentBrush(tool);
        ToolbarToolPickerButton.Content = CreateToolIcon(tool);
        _iconBrushOverride = null;
        ToolbarToolPickerButton.Background = GetBrush("ShellPanelBrush");
        ToolbarToolPickerButton.BorderBrush = GetBrush("ShellStrokeBrush");
        ToolbarToolPickerButton.BorderThickness = new Thickness(1);
        CurrentToolText.Text = SafetyText(_toolLabels[tool]);
        ToolbarCurrentToolText.Text = SafetyText(_toolLabels[tool]);
        UpdateWorkspaceContext();
        UpdateStatus();
    }

    private void ToolButton_Click(object sender, RoutedEventArgs e)
    {
        var tag = sender switch
        {
            Button button => button.Tag?.ToString(),
            MenuFlyoutItem item => item.Tag?.ToString(),
            _ => null
        };

        if (tag is not null && Enum.TryParse(tag, out EditorTool tool))
        {
            SelectTool(tool);
        }
    }

    private async void PaletteButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentDocument is null || sender is not Button button || button.Tag is not string key)
        {
            return;
        }

        await ApplyPaletteSelectionAsync(key, _currentDocument.SelectedPaletteShadeIndex);
    }

    private async Task ApplyPaletteSelectionAsync(string key, int shadeIndex)
    {
        if (_currentDocument is null || !_palettes.TryGetValue(key, out _))
        {
            return;
        }

        var paletteChanged = !string.Equals(_currentDocument.SelectedPaletteKey, key, StringComparison.Ordinal);
        var shadeChanged = _currentDocument.SelectedPaletteShadeIndex != shadeIndex;
        bool? stickerRenumberChoice = null;

        if (_selectedAnnotation?.Kind == AnnotationKind.Sticker && (paletteChanged || shadeChanged))
        {
            stickerRenumberChoice = await PromptStickerRenumberAsync();
            if (stickerRenumberChoice is null)
            {
                return;
            }
        }
        else if (_currentDocument.ResetStickerNumberOnColorChange && (paletteChanged || shadeChanged))
        {
            stickerRenumberChoice = true;
        }

        PushUndoState(CaptureProjectState());
        ApplyPaletteSelection(key, shadeIndex, updateDocument: true, updateSelectedAnnotation: true);

        if (stickerRenumberChoice == true)
        {
            RenumberStickers(_currentDocument);
        }

        MarkDocumentDirty(_currentDocument);
        RenderAnnotations(_currentDocument);
        if (_selectedAnnotation is not null)
        {
            SelectAnnotation(_selectedAnnotation);
        }
        UpdateStickerSequenceText(_currentDocument);
        UpdateLegendList();
        UpdateAnnotationCount(_currentDocument);
        UpdateStatus();
    }

    private void ApplyPaletteSelection(string key, int shadeIndex, bool updateDocument, bool updateSelectedAnnotation)
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

        _isSyncingPaletteUi = true;
        try
        {
            ToolbarPaletteComboBox.SelectedItem = _paletteOrder.FirstOrDefault(entry => string.Equals(entry.Key, key, StringComparison.Ordinal));
        }
        finally
        {
            _isSyncingPaletteUi = false;
        }

        var paletteSummary = string.Format(CultureInfo.CurrentCulture, SafetyText("UiActivePalette0"), palette.DisplayName);
        PaletteSummaryText.Text = paletteSummary;

        if (_currentDocument is not null && updateDocument)
        {
            _currentDocument.SelectedPaletteKey = key;
            _currentDocument.SelectedPaletteShadeIndex = Math.Clamp(shadeIndex, 0, 6);
            _currentDocument.PaletteDisplayName = palette.DisplayName;
        }

        if (updateSelectedAnnotation && _selectedAnnotation is not null)
        {
            _selectedAnnotation.PaletteKey = key;
            _selectedAnnotation.PaletteShadeIndex = Math.Clamp(shadeIndex, 0, 6);
        }

        RefreshShadeStrip();
    }

    private async Task<bool?> PromptStickerRenumberAsync()
    {
        if (RootGrid.XamlRoot is null)
        {
            return null;
        }

        var dialog = new ContentDialog
        {
            XamlRoot = RootGrid.XamlRoot,
            Title = SafetyText("UiChangeStickerColor"),
            Content = SafetyText("UiDoYouWantToRenumberTheStickers"),
            PrimaryButtonText = SafetyText("UiRenumber"),
            SecondaryButtonText = SafetyText("UiKeepNumbering"),
            CloseButtonText = SafetyText("UiCancel"),
            DefaultButton = ContentDialogButton.Primary
        };

        var result = await dialog.ShowAsync();
        return result switch
        {
            ContentDialogResult.Primary => true,
            ContentDialogResult.Secondary => false,
            _ => null
        };
    }

    private void AnnotationTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isApplyingDocumentState || _isSyncingAnnotationSelectionUi || _currentDocument is null)
        {
            return;
        }

        if (_selectedAnnotation is not null)
        {
            if (_selectedAnnotation.Kind == AnnotationKind.Text)
            {
                _selectedAnnotation.Text = AnnotationTextBox.Text;
                var size = MeasureTextAnnotationSize(_selectedAnnotation.Text, _selectedAnnotation.FontSize);
                _selectedAnnotation.Bounds = new Rect(_selectedAnnotation.Bounds.X, _selectedAnnotation.Bounds.Y, size.Width, size.Height);
                RenderAnnotations(_currentDocument);
                SelectAnnotation(_selectedAnnotation);
            }
        }
        else
        {
            _currentDocument.AnnotationText = AnnotationTextBox.Text;
        }

        MarkDocumentDirty(_currentDocument);
        UpdateStickerSequenceText(_currentDocument);
        UpdateLegendList();
    }

    private void StickerLegendTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_isApplyingDocumentState || _isSyncingAnnotationSelectionUi || _currentDocument is null)
        {
            return;
        }

        if (_selectedAnnotation?.Kind == AnnotationKind.Sticker)
        {
            _selectedAnnotation.LegendText = StickerLegendTextBox.Text;
            MarkDocumentDirty(_currentDocument);
            UpdateLegendList();
            UpdateStatus();
        }
    }

    private void StickerModeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isApplyingDocumentState || _currentDocument is null)
        {
            return;
        }

        PushUndoState(CaptureProjectState());
        _currentDocument.StickerModeIndex = StickerModeComboBox.SelectedIndex == 1 ? 1 : 0;
        MarkDocumentDirty(_currentDocument);
        UpdateStickerSequenceText(_currentDocument);
    }

    private void StrokeThicknessSlider_ValueChanged(object sender, RangeBaseValueChangedEventArgs e)
    {
        if (_currentDocument is null || _isApplyingDocumentState || _isSyncingAnnotationSelectionUi)
        {
            return;
        }

        PushUndoState(CaptureProjectState());
        _currentDocument.StrokeThickness = StrokeThicknessSlider.Value;
        if (_selectedAnnotation is not null)
        {
            _selectedAnnotation.StrokeThickness = StrokeThicknessSlider.Value;
            RenderAnnotations(_currentDocument);
            SelectAnnotation(_selectedAnnotation);
        }

        MarkDocumentDirty(_currentDocument);
        StrokeSummaryText.Text = string.Format(CultureInfo.CurrentCulture, SafetyText("StrokeCurrent"), StrokeThicknessSlider.Value);
    }

    private void ResetStickerNumberOnColorChangeCheckBox_Changed(object sender, RoutedEventArgs e)
    {
        if (_isApplyingDocumentState || _currentDocument is null)
        {
            return;
        }

        PushUndoState(CaptureProjectState());
        _currentDocument.ResetStickerNumberOnColorChange = ResetStickerNumberOnColorChangeCheckBox.IsChecked == true;
        MarkDocumentDirty(_currentDocument);
    }

    private void UpdateStickerSequenceText(ScreenshotDocument document)
    {
        var nextLabel = GetStickerModeIndex(document) switch
        {
            1 => ToAlphabetic(document.NextStickerIndex),
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
        if (_isImportingImage || _isSavingProject || _isOpeningProject || _isApplyingProjectState || _isCloseConfirmationOpen) return;
        _isImportingImage = true;
        try
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
                StatusText.Text = SafetyText("ImportCancelled");
                return;
            }

            var document = await CreateDocumentFromFileAsync(file);
            if (document is null)
            {
                StatusText.Text = SafetyText("ImportUnreadable");
                return;
            }

            PushUndoState(CaptureProjectState());
            AddDocument(document, select: true);
            document.IsDirty = true;
            StatusText.Text = string.Format(CultureInfo.CurrentCulture, SafetyText("ImportFileSuccess"), file.Name);

        }
        catch (Exception error)
        {
            AppDiagnostics.Record(error);
            StatusText.Text = SafetyText("ImportFailed");
        }
        finally
        {
            _isImportingImage = false;
        }
    }

    private void NewDemoTabButton_Click(object sender, RoutedEventArgs e)
    {
        PushUndoState(CaptureProjectState());
        var document = CreateDemoDocument();
        AddDocument(document, select: true);
        document.IsDirty = true;
        StatusText.Text = SafetyText("DemoCreated");
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
        if (_isSavingProject || _isOpeningProject || _isApplyingProjectState || _isCloseConfirmationOpen || _isProjectDialogOpen) return;
        if (document.IsDirty)
        {
            var dialog = new ContentDialog
            {
                XamlRoot = RootGrid.XamlRoot,
                Title = SafetyText("CloseStepTitle"),
                Content = SafetyText("CloseStepBody"),
                PrimaryButtonText = SafetyText("Save"),
                SecondaryButtonText = SafetyText("Discard"),
                CloseButtonText = SafetyText("Cancel"),
                DefaultButton = ContentDialogButton.Primary
            };

            var result = await dialog.ShowAsync();
            if (result == ContentDialogResult.None)
            {
                return;
            }

            if (result == ContentDialogResult.Primary)
            {
                var saved = await SaveProjectAsync();
                if (HasUnsavedProjectChanges()) return;
                if (!saved)
                {
                    return;
                }
            }
        }

        var previousSelection = _currentDocument;
        var removeIndex = _documents.IndexOf(document);
        PushUndoState(CaptureProjectState());
        _documents.Remove(document);
        RefreshDocumentOrderMetadata();
        var projectRemainsClean = !_documents.Any(item => item.IsDirty);

        if (_documents.Count == 0)
        {
            var replacement = CreateDemoDocument(isInitial: true);
            AddDocument(replacement, select: true);
            RefreshDocumentOrderMetadata();
            if (projectRemainsClean)
            {
                MarkDocumentDirty(replacement);
            }
            return;
        }

        if (previousSelection is not null && previousSelection != document && _documents.Contains(previousSelection))
        {
            DocumentTabsListView.SelectedItem = previousSelection;
            if (projectRemainsClean)
            {
                MarkDocumentDirty(previousSelection);
            }
            return;
        }

        var nextIndex = Math.Clamp(removeIndex, 0, _documents.Count - 1);
        if (projectRemainsClean)
        {
            MarkDocumentDirty(_documents[nextIndex]);
        }

        DocumentTabsListView.SelectedItem = _documents[nextIndex];
    }

    private async void ExportButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentDocument is null)
        {
            return;
        }

        if (ReferenceEquals(sender, ExportCurrentButton))
        {
            var format = App.Preferences.DefaultExportFormat;
            if (format == ExportFormatChoice.Png)
            {
                await SaveDocumentAsync(_currentDocument);
            }
            else
            {
                await ExportGuideAsync(format);
            }

            return;
        }

        await SaveDocumentAsync(_currentDocument);
    }

    private async Task<bool> SaveDocumentAsync(ScreenshotDocument document)
    {
        try
        {
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
                    StatusText.Text = SafetyText("PngCancelled");
                    return false;
                }
            }

            var snapshot = document.CloneForHistory();
            await Task.Run(() => GuideImageRenderer.SaveDocumentImage(snapshot, _palettes, file.Path));

            document.SavedPath = file.Path;
            StatusText.Text = string.Format(CultureInfo.CurrentCulture, SafetyText("ExportSuccess"), file.Path);
            return true;

        }
        catch (Exception error)
        {
            AppDiagnostics.Record(error);
            StatusText.Text = SafetyText("ExportFailed");
            return false;
        }
    }

    private void ClearAnnotationsButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentDocument is null || _currentDocument.Annotations.Count == 0)
        {
            return;
        }

        PushUndoState(CaptureProjectState());
        _currentDocument.Annotations.Clear();
        MarkDocumentDirty(_currentDocument);
        RenderAnnotations(_currentDocument);
        UpdateAnnotationCount(_currentDocument);
        StatusText.Text = SafetyText("AnnotationsCleared");
    }

    private void ApplyCropButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentDocument?.PendingCropRect is not Rect crop || crop.Width < 8 || crop.Height < 8)
        {
            StatusText.Text = SafetyText("CropDrawFirst");
            return;
        }

        PushUndoState(CaptureProjectState());
        _currentDocument.AppliedCropRect = crop;
        _currentDocument.PendingCropRect = null;
        MarkDocumentDirty(_currentDocument);
        ApplyCropState(_currentDocument);
        StatusText.Text = SafetyText("CropApplied");
    }

    private void ResetCropButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentDocument is null)
        {
            return;
        }

        var hadAppliedCrop = _currentDocument.AppliedCropRect is not null;
        PushUndoState(CaptureProjectState());
        _currentDocument.PendingCropRect = null;
        _currentDocument.AppliedCropRect = null;
        if (hadAppliedCrop)
        {
            MarkDocumentDirty(_currentDocument);
        }

        ApplyCropState(_currentDocument);
        StatusText.Text = SafetyText("CropReset");
    }

    private void ApplyCropState(ScreenshotDocument document)
    {
        if (document.PendingCropRect is null && document.AppliedCropRect is Rect appliedCrop)
        {
            ApplySceneViewportLayout(appliedCrop.Width, appliedCrop.Height);
            SceneTransform.X = -appliedCrop.X;
            SceneTransform.Y = -appliedCrop.Y;
            CropStateText.Text = string.Format(CultureInfo.CurrentCulture, SafetyText("CropAppliedSize"), appliedCrop.Width, appliedCrop.Height);
            HideCropOverlay();
            ResetCropButton.IsEnabled = true;
            ApplyCropButton.IsEnabled = false;
            return;
        }

        ApplyDocumentSceneSize(document);
        SceneTransform.X = 0;
        SceneTransform.Y = 0;
        CropStateText.Text = document.PendingCropRect is Rect pendingCrop
            ? string.Format(CultureInfo.CurrentCulture, SafetyText("CropPendingSize"), pendingCrop.Width, pendingCrop.Height)
            : SafetyText("CropNone");

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

    private void ApplyDocumentSceneSize(ScreenshotDocument document)
    {
        ApplySceneViewportLayout(GetSceneWidth(document), GetSceneHeight(document));
    }

    private void ApplySceneViewportLayout(double sceneWidth, double sceneHeight)
    {
        var width = Math.Max(1, sceneWidth);
        var height = Math.Max(1, sceneHeight);

        SceneHost.Width = Math.Max(1, _currentDocument?.SourcePixelWidth ?? width);
        SceneHost.Height = Math.Max(1, _currentDocument?.SourcePixelHeight ?? height);
        SceneHost.Margin = new Thickness(FloatingToolRailGutterWidth, 0, 0, 0);
        PreviewViewport.Width = width * _currentZoom + FloatingToolRailGutterWidth;
        PreviewViewport.Height = height * _currentZoom;
        UpdateViewportClip(PreviewViewport.Width, PreviewViewport.Height);
    }

    private static double GetSceneWidth(ScreenshotDocument document)
    {
        if (document.PendingCropRect is null && document.AppliedCropRect is Rect appliedCrop)
        {
            return Math.Max(1, appliedCrop.Width);
        }

        return document.SourcePixelWidth > 0 ? document.SourcePixelWidth : SceneWidth;
    }

    private static double GetSceneHeight(ScreenshotDocument document)
    {
        if (document.PendingCropRect is null && document.AppliedCropRect is Rect appliedCrop)
        {
            return Math.Max(1, appliedCrop.Height);
        }

        return document.SourcePixelHeight > 0 ? document.SourcePixelHeight : SceneHeight;
    }

    private void UpdateViewportClip(double width, double height)
    {
        if (PreviewViewport.Clip is RectangleGeometry clip)
        {
            clip.Rect = new Rect(0, 0, width, height);
        }
    }

    private void SceneHost_PointerPressed(object sender, PointerRoutedEventArgs e)
    {
        var document = _currentDocument;
        if (document is null)
        {
            return;
        }

        var point = e.GetCurrentPoint(SceneHost).Position;
        var sceneWidth = document.SourcePixelWidth;
        var sceneHeight = document.SourcePixelHeight;
        if (e.GetCurrentPoint(SceneHost).Properties.IsRightButtonPressed)
        {
            ClearSelection();
            e.Handled = true;
            return;
        }

        if (point.X < 0 || point.Y < 0 || point.X > sceneWidth || point.Y > sceneHeight)
        {
            return;
        }

        if (_currentTool == EditorTool.Text)
        {
            AddTextAnnotation(document, point);
            e.Handled = true;
            return;
        }

        if (_currentTool == EditorTool.Sticker)
        {
            AddStickerAnnotation(document, point);
            e.Handled = true;
            return;
        }

        _dragStart = point;
        _isDragging = true;
        SceneHost.CapturePointer(e.Pointer);

        if (_currentTool == EditorTool.Crop)
        {
            ShowCropOverlay(new Rect(point.X, point.Y, 0, 0));
            return;
        }

        _activeElement = CreateTemporaryElement(document);
        if (_activeElement is not null)
        {
            AnnotationCanvas.Children.Add(_activeElement);
        }
    }

    private void SceneHost_PointerMoved(object sender, PointerRoutedEventArgs e)
    {
        if (!_isDragging || _currentDocument is null)
        {
            return;
        }

        var point = e.GetCurrentPoint(SceneHost).Position;
        if (_currentTool == EditorTool.Crop)
        {
            ShowCropOverlay(NormalizeRect(_dragStart, point));
            return;
        }

        UpdateTemporaryElement(point);
    }

    private void SceneHost_PointerReleased(object sender, PointerRoutedEventArgs e)
    {
        if (!_isDragging || _currentDocument is null)
        {
            return;
        }

        var point = e.GetCurrentPoint(SceneHost).Position;

        if (_currentTool == EditorTool.Crop)
        {
            var crop = CropGeometry.Clamp(NormalizeRect(_dragStart, point), _currentDocument.SourcePixelWidth, _currentDocument.SourcePixelHeight);
            if (crop.Width >= 8 && crop.Height >= 8)
            {
                PushUndoState(CaptureProjectState());
                _currentDocument.PendingCropRect = crop;
                MarkDocumentDirty(_currentDocument);
                ApplyCropState(_currentDocument);
                StatusText.Text = SafetyText("CropPendingHint");
            }
            else
            {
                ApplyCropState(_currentDocument);
                StatusText.Text = SafetyText("CropTooSmall");
            }

            FinishDrag(e);
            return;
        }

        if (Distance(_dragStart, point) < 6)
        {
            FinalizeClickedAnnotation(_currentDocument, point);
        }
        else
        {
            FinalizeDraggedAnnotation(_currentDocument, point);
        }
        FinishDrag(e);
    }

    private void ClearSelection()
    {
        SelectAnnotation(null);
        UpdateStatus();
    }

    private void FinalizeClickedAnnotation(ScreenshotDocument document, Point point)
    {
        if (_activeElement is not null)
        {
            AnnotationCanvas.Children.Remove(_activeElement);
            _activeElement = null;
        }

        AnnotationModel? annotation = _currentTool switch
        {
            EditorTool.Rectangle => new AnnotationModel
            {
                Kind = AnnotationKind.Rectangle,
                Bounds = new Rect(point.X - 90, point.Y - 60, 180, 120),
                PaletteKey = document.SelectedPaletteKey,
                PaletteShadeIndex = document.SelectedPaletteShadeIndex,
                StrokeThickness = document.StrokeThickness,
                Opacity = document.DefaultOpacity
            },
            EditorTool.Ellipse => new AnnotationModel
            {
                Kind = AnnotationKind.Ellipse,
                Bounds = new Rect(point.X - 90, point.Y - 60, 180, 120),
                PaletteKey = document.SelectedPaletteKey,
                PaletteShadeIndex = document.SelectedPaletteShadeIndex,
                StrokeThickness = document.StrokeThickness,
                Opacity = document.DefaultOpacity
            },
            EditorTool.Focus => new AnnotationModel
            {
                Kind = AnnotationKind.Focus,
                Bounds = new Rect(point.X - 100, point.Y - 70, 200, 140),
                PaletteKey = document.SelectedPaletteKey,
                PaletteShadeIndex = document.SelectedPaletteShadeIndex,
                StrokeThickness = Math.Max(3, document.StrokeThickness),
                Opacity = document.DefaultOpacity
            },
            EditorTool.Mask => new AnnotationModel
            {
                Kind = AnnotationKind.Mask,
                Bounds = new Rect(point.X - 100, point.Y - 70, 200, 140),
                PaletteKey = document.SelectedPaletteKey,
                PaletteShadeIndex = document.SelectedPaletteShadeIndex,
                StrokeThickness = 2,
                Opacity = document.DefaultOpacity
            },
            EditorTool.ArrowStraight => new AnnotationModel
            {
                Kind = AnnotationKind.ArrowStraight,
                StartPoint = point,
                EndPoint = new Point(point.X + 140, point.Y - 40),
                PaletteKey = document.SelectedPaletteKey,
                PaletteShadeIndex = document.SelectedPaletteShadeIndex,
                StrokeThickness = document.StrokeThickness,
                Opacity = document.DefaultOpacity
            },
            EditorTool.ArrowCurved => new AnnotationModel
            {
                Kind = AnnotationKind.ArrowCurved,
                StartPoint = point,
                EndPoint = new Point(point.X + 140, point.Y - 40),
                PaletteKey = document.SelectedPaletteKey,
                PaletteShadeIndex = document.SelectedPaletteShadeIndex,
                StrokeThickness = document.StrokeThickness,
                Opacity = document.DefaultOpacity
            },
            _ => null
        };

        if (annotation is null)
        {
            return;
        }

        PushUndoState(CaptureProjectState());
        document.Annotations.Add(annotation);
        MarkDocumentDirty(document);
        RenderAnnotations(document);
        SelectAnnotation(annotation);
        UpdateAnnotationCount(document);
        UpdateStatus();
    }

    private FrameworkElement? CreateTemporaryElement(ScreenshotDocument document)
    {
        var stroke = CreateGradientBrush(document.SelectedPaletteKey, document.SelectedPaletteShadeIndex);
        return _currentTool switch
        {
            EditorTool.Rectangle => new Rectangle
            {
                Stroke = stroke,
                Fill = CreateHighlightFill(document.SelectedPaletteKey, 36, document.SelectedPaletteShadeIndex),
                RadiusX = 24,
                RadiusY = 24,
                StrokeThickness = document.StrokeThickness,
                Opacity = document.DefaultOpacity
            },
            EditorTool.Ellipse => new Ellipse
            {
                Stroke = stroke,
                Fill = CreateHighlightFill(document.SelectedPaletteKey, 30, document.SelectedPaletteShadeIndex),
                StrokeThickness = document.StrokeThickness,
                Opacity = document.DefaultOpacity
            },
            EditorTool.ArrowStraight => CreateArrowPath(stroke, document.StrokeThickness),
            EditorTool.ArrowCurved => CreateArrowPath(stroke, document.StrokeThickness),
            EditorTool.Focus => new Rectangle
            {
                Stroke = stroke,
                Fill = CreateHighlightFill(document.SelectedPaletteKey, 72, document.SelectedPaletteShadeIndex),
                RadiusX = 24,
                RadiusY = 24,
                StrokeThickness = Math.Max(3, document.StrokeThickness),
                Opacity = document.DefaultOpacity
            },
            EditorTool.Mask => new Rectangle
            {
                Stroke = new SolidColorBrush(ColorHelper.FromArgb(255, 18, 18, 18)),
                Fill = new SolidColorBrush(ColorHelper.FromArgb(215, 18, 18, 18)),
                RadiusX = 20,
                RadiusY = 20,
                StrokeThickness = 2,
                Opacity = document.DefaultOpacity
            },
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

        var before = CaptureProjectState();
        var annotationAdded = false;
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
                    PaletteShadeIndex = document.SelectedPaletteShadeIndex,
                    StrokeThickness = document.StrokeThickness,
                    Opacity = document.DefaultOpacity
                });
                annotationAdded = true;
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
                    PaletteShadeIndex = document.SelectedPaletteShadeIndex,
                    StrokeThickness = document.StrokeThickness,
                    Opacity = document.DefaultOpacity
                });
                annotationAdded = true;
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
                    PaletteShadeIndex = document.SelectedPaletteShadeIndex,
                    StrokeThickness = document.StrokeThickness,
                    Opacity = document.DefaultOpacity
                });
                annotationAdded = true;
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
                    PaletteShadeIndex = document.SelectedPaletteShadeIndex,
                    StrokeThickness = document.StrokeThickness,
                    Opacity = document.DefaultOpacity
                });
                annotationAdded = true;
                break;
            }
            case EditorTool.Focus:
            {
                var bounds = NormalizeRect(_dragStart, point);
                if (bounds.Width < 8 || bounds.Height < 8)
                {
                    return;
                }

                document.Annotations.Add(new AnnotationModel
                {
                    Kind = AnnotationKind.Focus,
                    Bounds = bounds,
                    PaletteKey = document.SelectedPaletteKey,
                    PaletteShadeIndex = document.SelectedPaletteShadeIndex,
                    StrokeThickness = Math.Max(3, document.StrokeThickness),
                    Opacity = document.DefaultOpacity
                });
                annotationAdded = true;
                break;
            }
            case EditorTool.Mask:
            {
                var bounds = NormalizeRect(_dragStart, point);
                if (bounds.Width < 8 || bounds.Height < 8)
                {
                    return;
                }

                document.Annotations.Add(new AnnotationModel
                {
                    Kind = AnnotationKind.Mask,
                    Bounds = bounds,
                    PaletteKey = document.SelectedPaletteKey,
                    PaletteShadeIndex = document.SelectedPaletteShadeIndex,
                    StrokeThickness = 2,
                    Opacity = document.DefaultOpacity
                });
                annotationAdded = true;
                break;
            }
        }

        if (!annotationAdded)
        {
            return;
        }

        PushUndoState(before);
        MarkDocumentDirty(document);
        RenderAnnotations(document);
        SelectAnnotation(document.Annotations[^1]);
        UpdateAnnotationCount(document);
        UpdateStatus();
    }

    private void AddTextAnnotation(ScreenshotDocument document, Point point, string? textOverride = null)
    {
        PushUndoState(CaptureProjectState());
        var text = string.IsNullOrWhiteSpace(textOverride)
            ? (string.IsNullOrWhiteSpace(document.AnnotationText) ? SafetyText("QuickNote") : document.AnnotationText.Trim())
            : textOverride.Trim();
        var size = MeasureTextAnnotationSize(text, textOverride is null ? 28 : 24);
        document.Annotations.Add(new AnnotationModel
        {
            Kind = AnnotationKind.Text,
            Bounds = new Rect(point.X, point.Y, size.Width, size.Height),
            PaletteKey = document.SelectedPaletteKey,
            Text = text,
            PaletteShadeIndex = document.SelectedPaletteShadeIndex,
            StrokeThickness = document.StrokeThickness,
            Opacity = document.DefaultOpacity,
            FontSize = textOverride is null ? 28 : 24
        });

        MarkDocumentDirty(document);
        RenderAnnotations(document);
        var annotation = document.Annotations[^1];
        SelectAnnotation(annotation);
        if (textOverride is null)
        {
            // Focus after layout and pointer handling have opened the inspector.
            DispatcherQueue.TryEnqueue(Microsoft.UI.Dispatching.DispatcherQueuePriority.Low, () =>
            {
                if (_editorClosed || !ReferenceEquals(_selectedAnnotation, annotation)) return;
                AnnotationTextBox.Focus(FocusState.Programmatic);
                AnnotationTextBox.SelectAll();
            });
        }
        UpdateAnnotationCount(document);
        UpdateStatus();
    }

    private static Size MeasureTextAnnotationSize(string text, double fontSize)
    {
        var probe = new TextBlock
        {
            Text = string.IsNullOrWhiteSpace(text) ? " " : text,
            FontFamily = new FontFamily("Segoe UI"),
            FontSize = fontSize,
            TextWrapping = TextWrapping.Wrap,
            TextTrimming = TextTrimming.None
        };

        probe.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
        var width = Math.Clamp(Math.Ceiling(probe.DesiredSize.Width + 28), TextAnnotationMinWidth, TextAnnotationMaxWidth);
        probe.Measure(new Size(Math.Max(1, width - 28), double.PositiveInfinity));
        var height = Math.Clamp(Math.Ceiling(probe.DesiredSize.Height + 20), TextAnnotationMinHeight, TextAnnotationMaxHeight);
        return new Size(width, height);
    }

    private void AddStickerAnnotation(ScreenshotDocument document, Point point)
    {
        PushUndoState(CaptureProjectState());
        document.Annotations.Add(new AnnotationModel
        {
            Kind = AnnotationKind.Sticker,
            Bounds = new Rect(point.X - (StickerSize / 2), point.Y - (StickerSize / 2), StickerSize, StickerSize),
            PaletteKey = document.SelectedPaletteKey,
            PaletteShadeIndex = document.SelectedPaletteShadeIndex,
            Text = BuildStickerLabel(document),
            StickerIndex = document.NextStickerIndex - 1,
            LegendText = string.Empty,
            Opacity = document.DefaultOpacity
        });

        MarkDocumentDirty(document);
        RenderAnnotations(document);
        SelectAnnotation(document.Annotations[^1]);
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
        return GetStickerModeIndex(document) switch
        {
            1 => AdvanceAlphabeticSticker(document),
            _ => AdvanceNumericSticker(document)
        };
    }

    private static int GetStickerModeIndex(ScreenshotDocument document)
    {
        return document.StickerModeIndex == 1 ? 1 : 0;
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
        RefreshAnnotationObjects();
        AnnotationCanvas.Children.Clear();
        foreach (var annotation in document.Annotations.Where(item => item.Kind != AnnotationKind.Mask)
            .Concat(document.Annotations.Where(item => item.Kind == AnnotationKind.Mask)))
        {
            switch (annotation.Kind)
            {
                case AnnotationKind.Text:
                {
                    var isEditing = ReferenceEquals(_editingTextAnnotation, annotation);
                    var border = new Border
                    {
                        Background = new SolidColorBrush(ColorHelper.FromArgb(235, 255, 255, 255)),
                        BorderBrush = CreateGradientBrush(annotation.PaletteKey, annotation.PaletteShadeIndex),
                        BorderThickness = new Thickness(2),
                        CornerRadius = new CornerRadius(14),
                        Padding = new Thickness(12, 8, 12, 8),
                        Width = Math.Max(220, annotation.Bounds.Width),
                        Height = Math.Max(60, annotation.Bounds.Height),
                        Opacity = annotation.Opacity
                    };

                    if (isEditing)
                    {
                        var editor = new TextBox
                        {
                            Tag = annotation,
                            Text = annotation.Text,
                            FontFamily = new FontFamily("Segoe UI"),
                            FontSize = annotation.FontSize,
                            TextWrapping = TextWrapping.Wrap,
                            AcceptsReturn = true,
                            BorderThickness = new Thickness(0),
                            Background = new SolidColorBrush(Colors.Transparent),
                            Foreground = CreateGradientBrush(annotation.PaletteKey, annotation.PaletteShadeIndex),
                            HorizontalAlignment = HorizontalAlignment.Stretch,
                            VerticalAlignment = VerticalAlignment.Stretch
                        };
                        ScrollViewer.SetVerticalScrollBarVisibility(editor, ScrollBarVisibility.Auto);
                        editor.TextChanged += InlineTextEditor_TextChanged;
                        editor.LostFocus += InlineTextEditor_LostFocus;
                        editor.Loaded += (_, _) =>
                        {
                            editor.Focus(FocusState.Programmatic);
                            editor.SelectAll();
                        };
                        border.Child = editor;
                    }
                    else
                    {
                        border.Child = new TextBlock
                        {
                            Text = annotation.Text,
                            FontFamily = new FontFamily("Segoe UI"),
                            FontSize = annotation.FontSize,
                            TextWrapping = TextWrapping.Wrap,
                            Foreground = CreateGradientBrush(annotation.PaletteKey, annotation.PaletteShadeIndex)
                        };
                    }

                    AttachAnnotationInteraction(border, annotation);
                    AnnotationCanvas.Children.Add(border);
                    SetCanvasPosition(border, annotation.Bounds);
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
                        Opacity = annotation.Opacity,
                        Child = new TextBlock
                        {
                            Text = annotation.Text,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center,
                            FontFamily = new FontFamily("Segoe UI"),
                            FontSize = 28,
                            Foreground = new SolidColorBrush(Colors.White)
                        }
                    };

                    AttachAnnotationInteraction(sticker, annotation);
                    AnnotationCanvas.Children.Add(sticker);
                    SetCanvasPosition(sticker, annotation.Bounds);
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
                        Stroke = CreateGradientBrush(annotation.PaletteKey, annotation.PaletteShadeIndex),
                        Fill = CreateHighlightFill(annotation.PaletteKey, 36, annotation.PaletteShadeIndex),
                        StrokeThickness = annotation.StrokeThickness,
                        Opacity = annotation.Opacity
                    };

                    AttachAnnotationInteraction(rectangle, annotation);
                    AnnotationCanvas.Children.Add(rectangle);
                    SetCanvasPosition(rectangle, annotation.Bounds);
                    break;
                }
                case AnnotationKind.Ellipse:
                {
                    var ellipse = new Ellipse
                    {
                        Width = annotation.Bounds.Width,
                        Height = annotation.Bounds.Height,
                        Stroke = CreateGradientBrush(annotation.PaletteKey, annotation.PaletteShadeIndex),
                        Fill = CreateHighlightFill(annotation.PaletteKey, 30, annotation.PaletteShadeIndex),
                        StrokeThickness = annotation.StrokeThickness,
                        Opacity = annotation.Opacity
                    };

                    AttachAnnotationInteraction(ellipse, annotation);
                    AnnotationCanvas.Children.Add(ellipse);
                    SetCanvasPosition(ellipse, annotation.Bounds);
                    break;
                }
                case AnnotationKind.Focus:
                {
                    var rectangle = new Rectangle
                    {
                        Width = annotation.Bounds.Width,
                        Height = annotation.Bounds.Height,
                        RadiusX = 24,
                        RadiusY = 24,
                        Stroke = CreateGradientBrush(annotation.PaletteKey, annotation.PaletteShadeIndex),
                        Fill = CreateHighlightFill(annotation.PaletteKey, 72, annotation.PaletteShadeIndex),
                        StrokeThickness = Math.Max(3, annotation.StrokeThickness),
                        Opacity = annotation.Opacity
                    };

                    AttachAnnotationInteraction(rectangle, annotation);
                    AnnotationCanvas.Children.Add(rectangle);
                    SetCanvasPosition(rectangle, annotation.Bounds);
                    break;
                }
                case AnnotationKind.Mask:
                {
                    var rectangle = new Rectangle
                    {
                        Width = annotation.Bounds.Width,
                        Height = annotation.Bounds.Height,
                        RadiusX = 0,
                        RadiusY = 0,
                        Stroke = new SolidColorBrush(ColorHelper.FromArgb(255, 18, 18, 18)),
                        Fill = new SolidColorBrush(ColorHelper.FromArgb(255, 18, 18, 18)),
                        StrokeThickness = 2,
                        Opacity = 1
                    };

                    AttachAnnotationInteraction(rectangle, annotation);
                    AnnotationCanvas.Children.Add(rectangle);
                    SetCanvasPosition(rectangle, annotation.Bounds);
                    break;
                }
                case AnnotationKind.ArrowStraight:
                {
                    var path = new XamlPath
                    {
                        Stroke = CreateGradientBrush(annotation.PaletteKey, annotation.PaletteShadeIndex),
                        StrokeThickness = annotation.StrokeThickness,
                        StrokeLineJoin = PenLineJoin.Round,
                        StrokeStartLineCap = PenLineCap.Round,
                        StrokeEndLineCap = PenLineCap.Round,
                        Data = BuildStraightArrowGeometry(annotation.StartPoint, annotation.EndPoint, annotation.StrokeThickness),
                        Opacity = annotation.Opacity,
                    };

                    AttachAnnotationInteraction(path, annotation);
                    AnnotationCanvas.Children.Add(path);
                    break;
                }
                case AnnotationKind.ArrowCurved:
                {
                    var path = new XamlPath
                    {
                        Stroke = CreateGradientBrush(annotation.PaletteKey, annotation.PaletteShadeIndex),
                        StrokeThickness = annotation.StrokeThickness,
                        StrokeLineJoin = PenLineJoin.Round,
                        StrokeStartLineCap = PenLineCap.Round,
                        StrokeEndLineCap = PenLineCap.Round,
                        Data = BuildCurvedArrowGeometry(annotation.StartPoint, annotation.EndPoint, annotation.StrokeThickness),
                        Opacity = annotation.Opacity,
                    };

                    AttachAnnotationInteraction(path, annotation);
                    AnnotationCanvas.Children.Add(path);
                    break;
                }
            }
        }

        foreach (var visual in AnnotationCanvas.Children.OfType<FrameworkElement>())
        {
            if (visual.Tag is AnnotationModel model && !ReferenceEquals(model, _editingTextAnnotation))
                visual.Opacity = 0;
        }
        ImportedImage.Visibility = Visibility.Collapsed;
        PlaceholderPanel.Visibility = Visibility.Collapsed;
        RequestScenePreview();
        AnnotationCanvas.Children.Add(SelectionOutlineHalo);
        AnnotationCanvas.Children.Add(SelectionOutline);
        AnnotationCanvas.Children.Add(CropPreview);
        AnnotationCanvas.Children.Add(CropBadge);
    }

    private static void SetCanvasPosition(FrameworkElement element, Rect bounds)
    {
        Canvas.SetLeft(element, bounds.X);
        Canvas.SetTop(element, bounds.Y);
    }

    private void AttachAnnotationInteraction(FrameworkElement element, AnnotationModel annotation)
    {
        element.Tag = annotation;
        element.PointerPressed += AnnotationElement_PointerPressed;
        element.PointerMoved += AnnotationElement_PointerMoved;
        element.PointerReleased += AnnotationElement_PointerReleased;
        element.PointerCaptureLost += AnnotationElement_PointerCaptureLost;
        element.DoubleTapped += AnnotationElement_DoubleTapped;
    }

    private Brush CreateGradientBrush(string paletteKey, int shadeIndex = 3)
    {
        var palette = _palettes[paletteKey];
        var index = Math.Clamp(shadeIndex, 0, Math.Max(0, palette.Shades.Count - 1));
        var shade = palette.Shades.Count > 0 ? palette.Shades[index] : palette.StartColor;
        var highlight = InterpolateColor(shade, palette.EndColor, 0.25);
        var brush = new LinearGradientBrush
        {
            StartPoint = new Point(0, 0),
            EndPoint = new Point(1, 1)
        };
        brush.GradientStops.Add(new GradientStop { Color = shade, Offset = 0 });
        brush.GradientStops.Add(new GradientStop { Color = highlight, Offset = 1 });
        return brush;
    }

    private Brush CreateHighlightFill(string paletteKey, byte alpha, int shadeIndex = 3)
    {
        var palette = _palettes[paletteKey];
        var index = Math.Clamp(shadeIndex, 0, Math.Max(0, palette.Shades.Count - 1));
        var color = palette.Shades.Count > 0 ? palette.Shades[index] : palette.StartColor;
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
        CropBadgeText.Text = string.Format(CultureInfo.CurrentCulture, SafetyText("CropBadge"), crop.Width, crop.Height);
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
        SceneHost.ReleasePointerCapture(e.Pointer);
    }

    private void MarkDocumentDirty(ScreenshotDocument document)
    {
        document.IsDirty = true;
    }

    private void UpdateAnnotationCount(ScreenshotDocument document)
    {
        UpdateStatus();
    }

    private void UpdateStatus()
    {
        var hint = _editingTextAnnotation is not null
            ? SafetyText("UiDoubleClickToFinishEditingTheText")
            : SafetyText(_toolHints[_currentTool]);

        StatusText.Text = hint;

        var annotationCount = _currentDocument?.Annotations.Count ?? 0;
        AnnotationCountText.Text = string.Format(
            CultureInfo.CurrentCulture,
            SafetyText(annotationCount == 1 ? "WorkspaceAnnotationCountSingle" : "WorkspaceAnnotationCount"),
            annotationCount);
    }

    private SolidColorBrush GetBrush(string key)
    {
        return (SolidColorBrush)Application.Current.Resources[key];
    }

    private readonly ClipboardImportCoordinator _clipboardImports = new();
    private long _projectEpoch;
    private bool _clipboardDrainQueued;

    private void Clipboard_ContentChanged(object? sender, object e)
    {
        _clipboardImports.NotifyChanged();
        if (_clipboardDrainQueued) return;
        _clipboardDrainQueued = true;
        if (!DispatcherQueue.TryEnqueue(async () =>
        {
            try
            {
                // Wait only for a pending change; no periodic clipboard read or image decode.
                while (!_editorClosed && App.Preferences.AutoImportClipboard)
                {
                    if (_isApplyingProjectState || _isCloseConfirmationOpen || _isProjectDialogOpen || _isSavingProject || _isOpeningProject || _isProcessingClipboard)
                    {
                        await Task.Delay(100);
                        continue;
                    }
                    await TryImportClipboardCaptureAsync();
                    break;
                }
            }
            finally { _clipboardDrainQueued = false; }
        })) _clipboardDrainQueued = false;
    }

    private void MainWindow_Closed(object sender, WindowEventArgs args)
    {
        _editorClosed = true;
        ++_previewVersion;
        Clipboard.ContentChanged -= Clipboard_ContentChanged;
    }

    private Task PrimeClipboardWatcherAsync()
    {
        _clipboardWatcherPrimed = true;
        return Task.CompletedTask;
    }

    private async Task<byte[]?> ReadClipboardBitmapBytesAsync(bool force)
    {
        for (var attempt = 0; attempt < 5; attempt++)
        {
            if (attempt > 0)
            {
                await Task.Delay(150);
            }

            var dataPackage = Clipboard.GetContent();
            if (!dataPackage.Contains(StandardDataFormats.Bitmap))
            {
                return null;
            }

            // Ignore our own output for automatic import, but allow an explicit paste.
            if (!force && dataPackage.Contains(ClipboardImageSourceFormat)
                && await dataPackage.GetDataAsync(ClipboardImageSourceFormat) is string source
                && source == _clipboardImageSource)
                return null;

            var streamReference = await dataPackage.GetBitmapAsync();
            using var stream = await streamReference.OpenReadAsync();
            using var memory = new MemoryStream();
            await stream.AsStreamForRead().CopyToAsync(memory);

            var bytes = memory.ToArray();
            if (bytes.Length > 0)
            {
                return bytes;
            }
        }

        return null;
    }

    private async Task TryImportClipboardCaptureAsync(bool force = false)
    {
        if ((!force && !App.Preferences.AutoImportClipboard) || !_clipboardWatcherPrimed || _editorClosed
            || _isApplyingProjectState || _isCloseConfirmationOpen || _isProjectDialogOpen || _isSavingProject || _isOpeningProject || _isProcessingClipboard) return;
        _isProcessingClipboard = true;
        var epoch = _projectEpoch;
        try
        {
            var imported = await _clipboardImports.ImportAsync(force, () => ReadClipboardBitmapBytesAsync(force),
                CreateDocumentFromClipboardAsync,
                document => { PushUndoState(CaptureProjectState()); AddDocument(document, select: true); },
                () => epoch == _projectEpoch && !_editorClosed && !_isApplyingProjectState);
            if (imported > 0) StatusText.Text = SafetyText("ClipboardImported");
            else if (force) StatusText.Text = SafetyText("ClipboardNoImage");
        }
        catch (Exception error)
        {
            AppDiagnostics.Record(error);
            StatusText.Text = SafetyText("ClipboardFailed");
        }
        finally { _isProcessingClipboard = false; }
    }

    private async void ImportClipboardMenuItem_Click(object sender, RoutedEventArgs e) => await TryImportClipboardCaptureAsync(force: true);
}
