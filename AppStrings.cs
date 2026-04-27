using Microsoft.Windows.ApplicationModel.Resources;

namespace SnapSlate;

public static class AppStrings
{
    private static readonly ResourceLoader Loader = new();

    public static string Get(string key)
    {
        return Loader.GetString(key);
    }
}
