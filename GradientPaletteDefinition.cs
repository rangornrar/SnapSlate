using System.Collections.Generic;
using Microsoft.UI.Xaml.Media;
using Windows.Foundation;
using Windows.UI;

namespace SnapSlate;

public sealed record GradientPaletteDefinition(
    string Key,
    string DisplayName,
    Color StartColor,
    Color EndColor)
{
    public IReadOnlyList<Color> Shades { get; init; } = [];

    public Brush DisplayBrush
    {
        get
        {
            var brush = new LinearGradientBrush
            {
                StartPoint = new Point(0, 0),
                EndPoint = new Point(1, 1)
            };

            brush.GradientStops.Add(new GradientStop { Color = StartColor, Offset = 0 });
            brush.GradientStops.Add(new GradientStop { Color = EndColor, Offset = 1 });
            return brush;
        }
    }
}
