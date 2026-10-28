using System;
using Windows.Foundation;

namespace SnapSlate;

internal static class AnnotationMovement
{
    public static void Translate(AnnotationModel annotation, double x, double y)
    {
        if (!double.IsFinite(x) || !double.IsFinite(y)) throw new ArgumentOutOfRangeException(nameof(x));
        if (annotation.Kind is AnnotationKind.ArrowStraight or AnnotationKind.ArrowCurved)
        {
            annotation.StartPoint = new Point(annotation.StartPoint.X + x, annotation.StartPoint.Y + y);
            annotation.EndPoint = new Point(annotation.EndPoint.X + x, annotation.EndPoint.Y + y);
        }
        else annotation.Bounds = new Rect(annotation.Bounds.X + x, annotation.Bounds.Y + y, annotation.Bounds.Width, annotation.Bounds.Height);
    }
}
