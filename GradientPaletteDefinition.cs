using Windows.UI;

namespace SnapSlate;

public sealed record GradientPaletteDefinition(
    string Key,
    string DisplayName,
    Color StartColor,
    Color EndColor);
