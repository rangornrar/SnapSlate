using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.UI.Xaml;

namespace SnapSlate;

public enum ThemeChoice
{
    System,
    Light,
    Dark
}

public enum LanguageChoice
{
    System,
    French,
    English
}

public enum ExportFormatChoice
{
    Png,
    Html,
    Markdown,
    Pdf,
    Docx
}

public sealed class AppPreferences
{
    private static string DefaultExportFolder =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");

    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        WriteIndented = true,
        Converters =
        {
            new JsonStringEnumConverter(JsonNamingPolicy.CamelCase)
        }
    };

    public ThemeChoice Theme { get; set; } = ThemeChoice.System;

    public LanguageChoice Language { get; set; } = LanguageChoice.System;

    public string ExportFolder { get; set; } = string.Empty;

    public ExportFormatChoice DefaultExportFormat { get; set; } = ExportFormatChoice.Png;

    public bool AutoImportClipboard { get; set; } = true;

    public bool CollapseStepsPane { get; set; } = false;

    public static AppPreferences Load()
    {
        var path = GetPreferencesPath();
        if (!File.Exists(path))
        {
            return new AppPreferences { ExportFolder = DefaultExportFolder };
        }

        try
        {
            var json = File.ReadAllText(path);
            var preferences = JsonSerializer.Deserialize<AppPreferences>(json, SerializerOptions) ?? new AppPreferences();
            if (string.IsNullOrWhiteSpace(preferences.ExportFolder))
            {
                preferences.ExportFolder = DefaultExportFolder;
            }

            return preferences;
        }
        catch
        {
            return new AppPreferences { ExportFolder = DefaultExportFolder };
        }
    }

    public void Save()
    {
        var path = GetPreferencesPath();
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var json = JsonSerializer.Serialize(this, SerializerOptions);
        File.WriteAllText(path, json);
    }

    public string? GetLanguageOverride()
    {
        return Language switch
        {
            LanguageChoice.French => "fr-FR",
            LanguageChoice.English => "en-US",
            _ => null
        };
    }

    public ElementTheme GetRequestedTheme()
    {
        return Theme switch
        {
            ThemeChoice.Light => ElementTheme.Light,
            ThemeChoice.Dark => ElementTheme.Dark,
            _ => ElementTheme.Default
        };
    }

    private static string GetPreferencesPath()
    {
        return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SnapSlate", "preferences.json");
    }
}
