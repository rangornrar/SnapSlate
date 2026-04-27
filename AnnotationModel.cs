using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Windows.Foundation;

namespace SnapSlate;

public enum AnnotationKind
{
    Text,
    Sticker,
    Rectangle,
    Ellipse,
    ArrowStraight,
    ArrowCurved,
    Focus,
    Mask
}

public sealed class AnnotationModel : INotifyPropertyChanged
{
    private string _paletteKey = "Sunset";
    private int _paletteShadeIndex = 3;
    private double _strokeThickness = 6;
    private double _opacity = 1;
    private Point _startPoint;
    private Point _endPoint;
    private Rect _bounds;
    private string _text = string.Empty;
    private int _stickerModeIndex;
    private int _stickerIndex;
    private string _legendText = string.Empty;
    private double _fontSize = 28;
    private bool _isBold;

    public event PropertyChangedEventHandler? PropertyChanged;

    public Guid Id { get; set; } = Guid.NewGuid();

    public AnnotationKind Kind { get; init; }

    public string PaletteKey
    {
        get => _paletteKey;
        set => SetField(ref _paletteKey, value);
    }

    public int PaletteShadeIndex
    {
        get => _paletteShadeIndex;
        set => SetField(ref _paletteShadeIndex, value);
    }

    public double StrokeThickness
    {
        get => _strokeThickness;
        set => SetField(ref _strokeThickness, value);
    }

    public double Opacity
    {
        get => _opacity;
        set => SetField(ref _opacity, value);
    }

    public Point StartPoint
    {
        get => _startPoint;
        set => SetField(ref _startPoint, value);
    }

    public Point EndPoint
    {
        get => _endPoint;
        set => SetField(ref _endPoint, value);
    }

    public Rect Bounds
    {
        get => _bounds;
        set => SetField(ref _bounds, value);
    }

    public string Text
    {
        get => _text;
        set => SetField(ref _text, value);
    }

    public int StickerModeIndex
    {
        get => _stickerModeIndex;
        set => SetField(ref _stickerModeIndex, value);
    }

    public int StickerIndex
    {
        get => _stickerIndex;
        set => SetField(ref _stickerIndex, value);
    }

    public string LegendText
    {
        get => _legendText;
        set => SetField(ref _legendText, value);
    }

    public double FontSize
    {
        get => _fontSize;
        set => SetField(ref _fontSize, value);
    }

    public bool IsBold
    {
        get => _isBold;
        set => SetField(ref _isBold, value);
    }

    public AnnotationModel Clone()
    {
        return new AnnotationModel
        {
            Id = Id,
            Kind = Kind,
            PaletteKey = PaletteKey,
            PaletteShadeIndex = PaletteShadeIndex,
            StrokeThickness = StrokeThickness,
            Opacity = Opacity,
            StartPoint = StartPoint,
            EndPoint = EndPoint,
            Bounds = Bounds,
            Text = Text,
            StickerModeIndex = StickerModeIndex,
            StickerIndex = StickerIndex,
            LegendText = LegendText,
            FontSize = FontSize,
            IsBold = IsBold
        };
    }

    private bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return false;
        }

        field = value;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        return true;
    }
}
