namespace SnapSlate;

internal static class ProjectPersistence
{
    public static string Serialize(SnapSlateProjectState state)
    {
        return ProjectJson.Serialize(state);
    }

    public static SnapSlateProjectState? Deserialize(string json)
    {
        return ProjectJson.Deserialize<SnapSlateProjectState>(json);
    }
}
