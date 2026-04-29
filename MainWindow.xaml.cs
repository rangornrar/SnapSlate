using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Security.Cryptography;
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
    private const double FloatingToolRailGutterWidth = 96;
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
        [EditorTool.Select] = "Sélection",
        [EditorTool.Crop] = "Recadrer",
        [EditorTool.Text] = "Texte",
        [EditorTool.Sticker] = "Gommette",
        [EditorTool.ArrowStraight] = "Flèche droite",
        [EditorTool.ArrowCurved] = "Flèche courbe",
        [EditorTool.Rectangle] = "Rectangle",
        [EditorTool.Ellipse] = "Ovale",
        [EditorTool.Focus] = "Focus",
        [EditorTool.Mask] = "Masquage"
    };

    private readonly Dictionary<EditorTool, string> _toolHints = new()
    {
        [EditorTool.Select] = "Cliquez un objet pour le sélectionner et le déplacer.",
        [EditorTool.Crop] = "Glissez une zone puis appliquez le recadrage.",
        [EditorTool.Text] = "Cliquez dans la scène pour poser un texte.",
        [EditorTool.Sticker] = "Cliquez pour poser une gommette.",
        [EditorTool.ArrowStraight] = "Glissez pour tracer une flèche droite.",
        [EditorTool.ArrowCurved] = "Glissez pour tracer une flèche courbe.",
        [EditorTool.Rectangle] = "Glissez pour entourer une zone rectangulaire.",
        [EditorTool.Ellipse] = "Glissez pour entourer une zone ovale.",
        [EditorTool.Focus] = "Glissez pour mettre une zone en avant.",
        [EditorTool.Mask] = "Glissez pour masquer une zone."
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
    private bool _isInspectorPaneCollapsed;
    private bool _isProcessingClipboard;
    private bool _clipboardWatcherPrimed;
    private readonly DispatcherQueueTimer _clipboardPollTimer;
    private string? _lastClipboardFingerprint;
    private int _nextCaptureIndex = 1;
    private int _nextDemoIndex = 1;
    private double _currentZoom = 1.0;
    private ShellSection _currentSection = ShellSection.Procedure;
    private AppPreferences? _settingsSnapshot;

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

        _clipboardPollTimer = DispatcherQueue.CreateTimer();
        _clipboardPollTimer.Interval = TimeSpan.FromMilliseconds(850);
        _clipboardPollTimer.IsRepeating = true;
        _clipboardPollTimer.Tick += ClipboardPollTimer_Tick;
        _clipboardPollTimer.Start();
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
            "Démo Lorem Ipsum",
            "Maquette éditoriale Lorem Ipsum",
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
        var buffer = await FileIO.ReadBufferAsync(file);
        var bytes = buffer.ToArray();
        if (bytes.Length == 0)
        {
            return null;
        }

        var (sourceWidth, sourceHeight) = await GetImageSizeAsync(bytes);
        var thumbnail = await CreateBitmapImageAsync(bytes, 160);
        return CreateDocumentCore(
            System.IO.Path.GetFileNameWithoutExtension(file.Name),
            DocumentOrigin.FileImport,
            "Fichier importe",
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
        var (sourceWidth, sourceHeight) = await GetImageSizeAsync(bytes);
        var captureIndex = _nextCaptureIndex++;
        var thumbnail = await CreateBitmapImageAsync(bytes, 160);
        return CreateDocumentCore(
            $"Capture {captureIndex:00}",
            DocumentOrigin.ClipboardCapture,
            "Win + Shift + S",
            $"Capture Windows importee a {DateTime.Now:HH:mm:ss}",
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
            AnnotationText = "Note rapide",
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

    private async Task LoadDocumentAsync(ScreenshotDocument document)
    {
        if (_currentDocument == document && !_isApplyingDocumentState)
        {
            return;
        }

        _currentDocument = document;
        _selectedAnnotation = null;
        _isApplyingDocumentState = true;

        try
        {
            CanvasSummaryText.Text = document.Origin switch
            {
                DocumentOrigin.ClipboardCapture => "Nouvelle capture importee.",
                DocumentOrigin.FileImport => "Image importee.",
                DocumentOrigin.BlankProject => "Projet vide.",
                DocumentOrigin.Demo => "Page Lorem Ipsum.",
                _ => "Planche de demonstration."
            };
            CanvasFileNameText.Text = document.FileNameLabel;
            ImageSourceText.Text = $"Source : {document.SourceLabel}";

            AnnotationTextBox.Text = document.AnnotationText;
            StickerModeComboBox.SelectedIndex = document.StickerModeIndex;
            StrokeThicknessSlider.Value = document.StrokeThickness;
            ResetStickerNumberOnColorChangeCheckBox.IsChecked = document.ResetStickerNumberOnColorChange;
            StrokeSummaryText.Text = $"Trait actuel : {document.StrokeThickness:0} px";

            ApplyPaletteSelection(document.SelectedPaletteKey, document.SelectedPaletteShadeIndex, updateDocument: false, updateSelectedAnnotation: false);
            SyncDocumentEditors(document);

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
            _isApplyingDocumentState = false;
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
        CurrentToolText.Text = _toolLabels[tool];
        ToolbarCurrentToolText.Text = _toolLabels[tool];
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

        var paletteSummary = string.Format(CultureInfo.CurrentCulture, T("Palette active : {0}", "Active palette: {0}"), palette.DisplayName);
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
            Title = T("Changer la couleur de la gommette ?", "Change sticker color?"),
            Content = T("Voulez-vous renuméroter les gommettes après ce changement de couleur ?", "Do you want to renumber the stickers after this color change?"),
            PrimaryButtonText = T("Renuméroter", "Renumber"),
            SecondaryButtonText = T("Continuer", "Keep numbering"),
            CloseButtonText = T("Annuler", "Cancel"),
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
        StrokeSummaryText.Text = $"Trait actuel : {StrokeThicknessSlider.Value:0} px";
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

        PushUndoState(CaptureProjectState());
        AddDocument(document, select: true);
        document.IsDirty = true;
        StatusText.Text = $"Nouveau tab cree pour {file.Name}.";
    }

    private void NewDemoTabButton_Click(object sender, RoutedEventArgs e)
    {
        PushUndoState(CaptureProjectState());
        var document = CreateDemoDocument();
        AddDocument(document, select: true);
        document.IsDirty = true;
        StatusText.Text = "Nouvelle page Lorem Ipsum creee.";
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

        GuideImageRenderer.SaveDocumentImage(document, _palettes, file.Path);

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

        PushUndoState(CaptureProjectState());
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

        PushUndoState(CaptureProjectState());
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
        PushUndoState(CaptureProjectState());
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
            ApplySceneViewportLayout(appliedCrop.Width, appliedCrop.Height);
            SceneTransform.X = -appliedCrop.X;
            SceneTransform.Y = -appliedCrop.Y;
            CropStateText.Text = $"Crop applique : {appliedCrop.Width:0} x {appliedCrop.Height:0}px.";
            HideCropOverlay();
            ResetCropButton.IsEnabled = true;
            ApplyCropButton.IsEnabled = false;
            return;
        }

        ApplyDocumentSceneSize(document);
        SceneTransform.X = 0;
        SceneTransform.Y = 0;
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

    private void ApplyDocumentSceneSize(ScreenshotDocument document)
    {
        ApplySceneViewportLayout(GetSceneWidth(document), GetSceneHeight(document));
    }

    private void ApplySceneViewportLayout(double sceneWidth, double sceneHeight)
    {
        var width = Math.Max(1, sceneWidth);
        var height = Math.Max(1, sceneHeight);

        SceneHost.Width = width;
        SceneHost.Height = height;
        SceneHost.Margin = new Thickness(FloatingToolRailGutterWidth, 0, 0, 0);
        PreviewViewport.Width = width + FloatingToolRailGutterWidth;
        PreviewViewport.Height = height;
        UpdateViewportClip(PreviewViewport.Width, PreviewViewport.Height);
    }

    private static double GetSceneWidth(ScreenshotDocument document)
    {
        if (document.AppliedCropRect is Rect appliedCrop)
        {
            return Math.Max(1, appliedCrop.Width);
        }

        return document.SourcePixelWidth > 0 ? document.SourcePixelWidth : SceneWidth;
    }

    private static double GetSceneHeight(ScreenshotDocument document)
    {
        if (document.AppliedCropRect is Rect appliedCrop)
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
        var sceneWidth = GetSceneWidth(document);
        var sceneHeight = GetSceneHeight(document);
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
            return;
        }

        if (_currentTool == EditorTool.Sticker)
        {
            AddStickerAnnotation(document, point);
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

        PushUndoState(CaptureProjectState());
        MarkDocumentDirty(document);
        RenderAnnotations(document);
        UpdateAnnotationCount(document);
        UpdateStatus();
    }

    private void AddTextAnnotation(ScreenshotDocument document, Point point, string? textOverride = null)
    {
        PushUndoState(CaptureProjectState());
        var text = string.IsNullOrWhiteSpace(textOverride)
            ? (string.IsNullOrWhiteSpace(document.AnnotationText) ? "Note rapide" : document.AnnotationText.Trim())
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
        UpdateAnnotationCount(document);
        UpdateStatus();
    }

    private static Size MeasureTextAnnotationSize(string text, double fontSize)
    {
        var probe = new TextBlock
        {
            Text = string.IsNullOrWhiteSpace(text) ? " " : text,
            FontFamily = new FontFamily("Bahnschrift SemiBold SemiConden"),
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
        AnnotationCanvas.Children.Clear();
        foreach (var annotation in document.Annotations)
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
                        Padding = new Thickness(14, 10, 14, 10),
                        Width = Math.Max(220, annotation.Bounds.Width),
                        Height = Math.Max(96, annotation.Bounds.Height),
                        Opacity = annotation.Opacity
                    };

                    if (isEditing)
                    {
                        var editor = new TextBox
                        {
                            Tag = annotation,
                            Text = annotation.Text,
                            FontFamily = new FontFamily("Bahnschrift SemiBold SemiConden"),
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
                            FontFamily = new FontFamily("Bahnschrift SemiBold SemiConden"),
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
                            FontFamily = new FontFamily("Bahnschrift SemiBold SemiConden"),
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
                        RadiusX = 20,
                        RadiusY = 20,
                        Stroke = new SolidColorBrush(ColorHelper.FromArgb(255, 18, 18, 18)),
                        Fill = new SolidColorBrush(ColorHelper.FromArgb(215, 18, 18, 18)),
                        StrokeThickness = 2,
                        Opacity = annotation.Opacity
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
            ? T("Double-cliquez pour terminer la modification du texte.", "Double-click to finish editing the text.")
            : _toolHints[_currentTool];

        StatusText.Text = hint;

        var annotationCount = _currentDocument?.Annotations.Count ?? 0;
        AnnotationCountText.Text = string.Format(
            CultureInfo.CurrentCulture,
            T("{0} annotations · Zoom {1:0} %", "{0} annotations · Zoom {1:0} %"),
            annotationCount,
            _currentZoom * 100);
    }

    private SolidColorBrush GetBrush(string key)
    {
        return (SolidColorBrush)Application.Current.Resources[key];
    }

    private async void Clipboard_ContentChanged(object? sender, object e)
    {
        DispatcherQueue.TryEnqueue(async () =>
        {
            await TryImportClipboardCaptureAsync();
        });
    }

    private async void ClipboardPollTimer_Tick(DispatcherQueueTimer sender, object args)
    {
        await TryImportClipboardCaptureAsync();
    }

    private void MainWindow_Closed(object sender, WindowEventArgs args)
    {
        _clipboardPollTimer.Stop();
        Clipboard.ContentChanged -= Clipboard_ContentChanged;
    }

    private async Task PrimeClipboardWatcherAsync()
    {
        try
        {
            var bytes = await ReadClipboardBitmapBytesAsync();
            if (bytes is { Length: > 0 })
            {
                _lastClipboardFingerprint = await CreateClipboardFingerprintAsync(bytes);
            }
        }
        catch
        {
            // Best effort only.
        }
        finally
        {
            _clipboardWatcherPrimed = true;
        }
    }

    private async Task<byte[]?> ReadClipboardBitmapBytesAsync()
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
                continue;
            }

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

    private static async Task<string> CreateClipboardFingerprintAsync(byte[] bytes)
    {
        using var stream = new InMemoryRandomAccessStream();
        await stream.WriteAsync(bytes.AsBuffer());
        stream.Seek(0);

        try
        {
            var decoder = await BitmapDecoder.CreateAsync(stream);
            var pixelData = await decoder.GetPixelDataAsync(
                BitmapPixelFormat.Bgra8,
                BitmapAlphaMode.Premultiplied,
                new BitmapTransform(),
                ExifOrientationMode.IgnoreExifOrientation,
                ColorManagementMode.ColorManageToSRgb);

            var pixels = pixelData.DetachPixelData();
            var fingerprintBytes = new byte[8 + pixels.Length];
            global::System.Buffer.BlockCopy(BitConverter.GetBytes((int)decoder.PixelWidth), 0, fingerprintBytes, 0, 4);
            global::System.Buffer.BlockCopy(BitConverter.GetBytes((int)decoder.PixelHeight), 0, fingerprintBytes, 4, 4);
            global::System.Buffer.BlockCopy(pixels, 0, fingerprintBytes, 8, pixels.Length);
            return Convert.ToHexString(SHA256.HashData(fingerprintBytes));
        }
        catch
        {
            return Convert.ToHexString(SHA256.HashData(bytes));
        }
    }

    private async Task TryImportClipboardCaptureAsync()
    {
        try
        {
            if (!App.Preferences.AutoImportClipboard || !_clipboardWatcherPrimed)
            {
                return;
            }

            if (_isProcessingClipboard)
            {
                return;
            }

            _isProcessingClipboard = true;
            var bytes = await ReadClipboardBitmapBytesAsync();

            if (bytes is null || bytes.Length == 0)
            {
                return;
            }

            var fingerprint = await CreateClipboardFingerprintAsync(bytes);
            if (string.Equals(fingerprint, _lastClipboardFingerprint, StringComparison.Ordinal))
            {
                return;
            }

            _lastClipboardFingerprint = fingerprint;

            AddDocument(await CreateDocumentFromClipboardAsync(bytes), select: true);
            StatusText.Text = "Nouvelle capture importee.";
        }
        catch
        {
            StatusText.Text = "Impossible de lire la derniere capture depuis le presse-papiers.";
        }
        finally
        {
            _isProcessingClipboard = false;
        }
    }
}
