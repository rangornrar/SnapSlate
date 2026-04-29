using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.Json.Serialization;
using Microsoft.UI.Xaml.Media.Imaging;
using Windows.Foundation;

namespace SnapSlate;

public sealed class ScreenshotDocument : INotifyPropertyChanged
{
    private string _baseTitle = "Capture";
    private string _stepNote = string.Empty;
    private string _originLabel = "Demo";
    private string _sourceLabel = "Source inconnue";
    private string _fileNameLabel = "capture.png";
    private int _sourcePixelWidth = 1360;
    private int _sourcePixelHeight = 820;
    private byte[]? _imageBytes;
    private BitmapImage? _thumbnailSource;
    private string _selectedPaletteKey = "Sunset";
    private int _selectedPaletteShadeIndex = 3;
    private string _paletteDisplayName = "sunset note";
    private string _annotationText = "Note rapide";
    private int _stepNumber = 1;
    private GuideTemplateType _templateType = GuideTemplateType.UserManual;
    private string _audience = "Utilisateurs finaux";
    private string _author = "SnapSlate";
    private string _documentVersion = "1.0";
    private double _defaultOpacity = 1;
    private int _stickerModeIndex;
    private double _strokeThickness = 6;
    private int _nextStickerIndex = 1;
    private bool _resetStickerNumberOnColorChange;
    private Rect? _pendingCropRect;
    private Rect? _appliedCropRect;
    private string? _savedPath;
    private bool _isDirty;

    public event PropertyChangedEventHandler? PropertyChanged;

    public Guid Id { get; set; } = Guid.NewGuid();

    public DocumentOrigin Origin { get; set; } = DocumentOrigin.Demo;

    public List<AnnotationModel> Annotations { get; set; } = [];

    public string BaseTitle
    {
        get => _baseTitle;
        set
        {
            if (SetField(ref _baseTitle, value))
            {
                OnPropertyChanged(nameof(TabTitle));
                OnPropertyChanged(nameof(StepTitle));
            }
        }
    }

    public string StepTitle
    {
        get => BaseTitle;
        set => BaseTitle = value;
    }

    public string StepNote
    {
        get => _stepNote;
        set => SetField(ref _stepNote, value);
    }

    public string OriginLabel
    {
        get => _originLabel;
        set
        {
            if (SetField(ref _originLabel, value))
            {
                OnPropertyChanged(nameof(TabSubtitle));
            }
        }
    }

    public string SourceLabel
    {
        get => _sourceLabel;
        set => SetField(ref _sourceLabel, value);
    }

    public string FileNameLabel
    {
        get => _fileNameLabel;
        set => SetField(ref _fileNameLabel, value);
    }

    public int SourcePixelWidth
    {
        get => _sourcePixelWidth;
        set => SetField(ref _sourcePixelWidth, value);
    }

    public int SourcePixelHeight
    {
        get => _sourcePixelHeight;
        set => SetField(ref _sourcePixelHeight, value);
    }

    public byte[]? ImageBytes
    {
        get => _imageBytes;
        set => SetField(ref _imageBytes, value);
    }

    [JsonIgnore]
    public BitmapImage? ThumbnailSource
    {
        get => _thumbnailSource;
        set => SetField(ref _thumbnailSource, value);
    }

    public string SelectedPaletteKey
    {
        get => _selectedPaletteKey;
        set => SetField(ref _selectedPaletteKey, value);
    }

    public int SelectedPaletteShadeIndex
    {
        get => _selectedPaletteShadeIndex;
        set => SetField(ref _selectedPaletteShadeIndex, value);
    }

    public string PaletteDisplayName
    {
        get => _paletteDisplayName;
        set => SetField(ref _paletteDisplayName, value);
    }

    public string AnnotationText
    {
        get => _annotationText;
        set => SetField(ref _annotationText, value);
    }

    public int StepNumber
    {
        get => _stepNumber;
        set => SetField(ref _stepNumber, value);
    }

    public GuideTemplateType TemplateType
    {
        get => _templateType;
        set => SetField(ref _templateType, value);
    }

    public string Audience
    {
        get => _audience;
        set => SetField(ref _audience, value);
    }

    public string Author
    {
        get => _author;
        set => SetField(ref _author, value);
    }

    public string DocumentVersion
    {
        get => _documentVersion;
        set => SetField(ref _documentVersion, value);
    }

    public int StickerModeIndex
    {
        get => _stickerModeIndex;
        set => SetField(ref _stickerModeIndex, value);
    }

    public double StrokeThickness
    {
        get => _strokeThickness;
        set => SetField(ref _strokeThickness, value);
    }

    public double DefaultOpacity
    {
        get => _defaultOpacity;
        set => SetField(ref _defaultOpacity, value);
    }

    public int NextStickerIndex
    {
        get => _nextStickerIndex;
        set => SetField(ref _nextStickerIndex, value);
    }

    public bool ResetStickerNumberOnColorChange
    {
        get => _resetStickerNumberOnColorChange;
        set => SetField(ref _resetStickerNumberOnColorChange, value);
    }

    public Rect? PendingCropRect
    {
        get => _pendingCropRect;
        set => SetField(ref _pendingCropRect, value);
    }

    public Rect? AppliedCropRect
    {
        get => _appliedCropRect;
        set => SetField(ref _appliedCropRect, value);
    }

    [JsonIgnore]
    public string? SavedPath
    {
        get => _savedPath;
        set => SetField(ref _savedPath, value);
    }

    public bool IsDirty
    {
        get => _isDirty;
        set
        {
            if (SetField(ref _isDirty, value))
            {
                OnPropertyChanged(nameof(TabTitle));
                OnPropertyChanged(nameof(TabSubtitle));
            }
        }
    }

    public string TabTitle => IsDirty ? $"{BaseTitle} *" : BaseTitle;

    public string TabSubtitle => IsDirty ? "A sauvegarder" : OriginLabel;

    public ScreenshotDocument CloneForHistory()
    {
        var clone = new ScreenshotDocument
        {
            Id = Id,
            BaseTitle = BaseTitle,
            StepNote = StepNote,
            Origin = Origin,
            OriginLabel = OriginLabel,
            SourceLabel = SourceLabel,
            FileNameLabel = FileNameLabel,
            SourcePixelWidth = SourcePixelWidth,
            SourcePixelHeight = SourcePixelHeight,
            ImageBytes = ImageBytes,
            ThumbnailSource = ThumbnailSource,
            SelectedPaletteKey = SelectedPaletteKey,
            SelectedPaletteShadeIndex = SelectedPaletteShadeIndex,
            PaletteDisplayName = PaletteDisplayName,
            AnnotationText = AnnotationText,
            StepNumber = StepNumber,
            TemplateType = TemplateType,
            Audience = Audience,
            Author = Author,
            DocumentVersion = DocumentVersion,
            StickerModeIndex = StickerModeIndex,
            StrokeThickness = StrokeThickness,
            DefaultOpacity = DefaultOpacity,
            NextStickerIndex = NextStickerIndex,
            ResetStickerNumberOnColorChange = ResetStickerNumberOnColorChange,
            PendingCropRect = PendingCropRect,
            AppliedCropRect = AppliedCropRect,
            SavedPath = SavedPath,
            IsDirty = IsDirty
        };

        foreach (var annotation in Annotations)
        {
            clone.Annotations.Add(annotation.Clone());
        }

        return clone;
    }

    private bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return false;
        }

        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
