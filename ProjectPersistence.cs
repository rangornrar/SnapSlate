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
        if (state is not null) ProjectValidator.Validate(state);
        return state;
    }
}
