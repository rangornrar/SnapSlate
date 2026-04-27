using System.Collections.Generic;
using Windows.UI;

namespace SnapSlate;

public sealed record GradientPaletteDefinition(
    string Key,
    string DisplayName,
    Color StartColor,
    Color EndColor)
{
    public IReadOnlyList<Color> Shades { get; init; } = [];
}
