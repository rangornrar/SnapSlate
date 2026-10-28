using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace SnapSlate;

internal static class LocalizedText
{
    private static readonly Lazy<Dictionary<string, string>> French = new(() => Load("fr-FR"));
    private static readonly Lazy<Dictionary<string, string>> English = new(() => Load("en-US"));
    public static string CurrentLanguage { get; set; } = "fr-FR";

    public static CultureInfo Culture(string language) => CultureInfo.GetCultureInfo(language.StartsWith("en", StringComparison.OrdinalIgnoreCase) ? "en-US" : "fr-FR");

    public static string Get(string key, string? language = null)
    {
        var values = Culture(language ?? CurrentLanguage).Name == "en-US" ? English.Value : French.Value;
        return values.TryGetValue(key, out var value) ? value : throw new KeyNotFoundException("Missing resource: " + key);
    }

    private static Dictionary<string, string> Load(string language)
    {
        using var stream = typeof(LocalizedText).Assembly.GetManifestResourceStream($"SnapSlate.Strings.{language}.Resources.resw")
            ?? throw new InvalidOperationException("Missing language resource: " + language);
        return XDocument.Load(stream).Root!.Elements("data").ToDictionary(data => data.Attribute("name")!.Value, data => data.Element("value")!.Value);
    }
}
