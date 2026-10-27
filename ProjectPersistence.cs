namespace SnapSlate;

internal static class ProjectPersistence
{
    public static string Serialize(SnapSlateProjectState state)
    {
        return ProjectJson.Serialize(state);
    }

    public static SnapSlateProjectState? Deserialize(string json)
    {
        var state = ProjectJson.Deserialize<SnapSlateProjectState>(json);
        if (state is not null)
        {
            ProjectValidator.Validate(state);
            if (state.Version == 1)
            {
                state.Metadata = GuideMetadata.FromLegacyDocument(state.Documents.Count > 0 ? state.Documents[0] : null);
                state.Version = 2;
            }
        }
        return state;
    }
}
