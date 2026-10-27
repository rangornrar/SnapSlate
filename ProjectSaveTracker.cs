using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace SnapSlate;

internal sealed class ProjectSaveTracker
{
    private string? _savedFingerprint;
    private Dictionary<Guid, string> _savedDocuments = [];

    public bool HasChanges(SnapSlateProjectState state) => _savedFingerprint != Fingerprint(state);

    public void Accept(SnapSlateProjectState savedState)
    {
        _savedFingerprint = Fingerprint(savedState);
        _savedDocuments = savedState.Documents.ToDictionary(document => document.Id, DocumentFingerprint);
    }

    public void RefreshIndicators(SnapSlateProjectState currentState)
    {
        foreach (var document in currentState.Documents)
        {
            document.IsDirty = !_savedDocuments.TryGetValue(document.Id, out var fingerprint)
                || fingerprint != DocumentFingerprint(document);
        }
    }

    private static string Fingerprint(SnapSlateProjectState state) => Hash(string.Join("|", state.Documents.Select(DocumentFingerprint)));

    private static string DocumentFingerprint(ScreenshotDocument document)
    {
        var content = document.CloneForHistory();
        content.IsDirty = false;
        return Hash(ProjectJson.Serialize(content));
    }

    private static string Hash(string content) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(content)));
}
