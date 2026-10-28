using System;
using System.Collections.Generic;
using System.Linq;

namespace SnapSlate;

internal sealed class ProjectSaveTracker
{
    private string? _savedFingerprint;
    private Dictionary<Guid, string> _savedDocuments = [];

    public bool HasChanges(SnapSlateProjectState state) => _savedFingerprint != ProjectContentFingerprint.Project(state);

    public void Accept(SnapSlateProjectState savedState)
    {
        _savedFingerprint = ProjectContentFingerprint.Project(savedState);
        _savedDocuments = savedState.Documents.ToDictionary(document => document.Id, ProjectContentFingerprint.Document);
    }

    public void RefreshIndicators(SnapSlateProjectState currentState)
    {
        foreach (var document in currentState.Documents)
        {
            document.IsDirty = !_savedDocuments.TryGetValue(document.Id, out var fingerprint)
                || fingerprint != ProjectContentFingerprint.Document(document);
        }
    }

}
