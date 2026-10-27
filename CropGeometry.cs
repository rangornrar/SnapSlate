using System;
using Windows.Foundation;

namespace SnapSlate;

internal static class CropGeometry
{
    public static Rect Clamp(Rect crop, double width, double height)
    {
        var left = Math.Clamp(Math.Floor(crop.X), 0, Math.Max(0, width - 1));
        var top = Math.Clamp(Math.Floor(crop.Y), 0, Math.Max(0, height - 1));
        var right = Math.Clamp(Math.Ceiling(crop.Right), left + 1, width);
        var bottom = Math.Clamp(Math.Ceiling(crop.Bottom), top + 1, height);
        return new Rect(left, top, right - left, bottom - top);
    }
}
