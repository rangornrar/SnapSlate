using Windows.Foundation;

namespace SnapSlate;

public enum AnnotationKind
{
    Text,
    Sticker,
    Rectangle,
    Ellipse,
    ArrowStraight,
    ArrowCurved
}

public sealed class AnnotationModel
{
    public AnnotationKind Kind { get; init; }

    public string PaletteKey { get; init; } = "Sunset";

    public double StrokeThickness { get; init; } = 6;

    public Point StartPoint { get; init; }

    public Point EndPoint { get; init; }

    public Rect Bounds { get; init; }

    public string Text { get; init; } = string.Empty;
}
