using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Windows.Foundation;

namespace SnapSlate;

public sealed class ScreenshotDocument : INotifyPropertyChanged
{
    private string _baseTitle = "Capture";
    private string _originLabel = "Demo";
    private string _sourceLabel = "Source inconnue";
    private string _fileNameLabel = "capture.png";
    private string _selectedPaletteKey = "Sunset";
    private string _paletteDisplayName = "sunset note";
    private string _annotationText = "Note rapide";
    private int _stickerModeIndex;
    private double _strokeThickness = 6;
    private int _nextStickerIndex = 1;
    private bool _resetStickerNumberOnColorChange;
    private Rect? _pendingCropRect;
    private Rect? _appliedCropRect;
    private string? _savedPath;
    private bool _isDirty;
    private byte[]? _imageBytes;

    public event PropertyChangedEventHandler? PropertyChanged;

    public Guid Id { get; } = Guid.NewGuid();

    public DocumentOrigin Origin { get; init; } = DocumentOrigin.Demo;

    public List<AnnotationModel> Annotations { get; } = [];

    public string BaseTitle
    {
        get => _baseTitle;
        set
        {
            if (SetField(ref _baseTitle, value))
            {
                OnPropertyChanged(nameof(TabTitle));
            }
        }
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

    public byte[]? ImageBytes
    {
        get => _imageBytes;
        set => SetField(ref _imageBytes, value);
    }

    public string SelectedPaletteKey
    {
        get => _selectedPaletteKey;
        set => SetField(ref _selectedPaletteKey, value);
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
