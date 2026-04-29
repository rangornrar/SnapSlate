using System;
using System.Collections.Generic;

namespace SnapSlate;

internal sealed class SnapSlateProjectState
{
    public int Version { get; set; } = 1;

    public Guid? SelectedDocumentId { get; set; }

    public Guid? SelectedAnnotationId { get; set; }

    public List<ScreenshotDocument> Documents { get; set; } = [];
}
