using System;
using System.Collections.Generic;

namespace SnapSlate;

public sealed record GuideExportManifest(
    string Title,
    string TemplateName,
    string Audience,
    string Author,
    string Version,
    DateTimeOffset GeneratedAt,
    string ExportFolder,
    IReadOnlyList<GuideExportStep> Steps);

public sealed record GuideExportStep(
    int Index,
    string Title,
    string Note,
    string OriginLabel,
    string SourceLabel,
    string ImagePath,
    IReadOnlyList<GuideExportLegendItem> LegendItems);

public sealed record GuideExportLegendItem(
    string StickerLabel,
    string Description);
