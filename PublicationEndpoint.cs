using System;

namespace SnapSlate;

internal static class PublicationEndpoint
{
    public static Uri ConfluenceV2(string baseUrl)
    {
        var legacy = Confluence(baseUrl).AbsoluteUri;
        return new Uri(legacy[..^"rest/api/content/".Length] + "api/v2/");
    }

    public static Uri Confluence(string baseUrl)
    {
        if (!Uri.TryCreate(baseUrl, UriKind.Absolute, out var site) || site.Scheme != Uri.UriSchemeHttps
            || string.IsNullOrWhiteSpace(site.Host) || !string.IsNullOrEmpty(site.UserInfo)
            || !string.IsNullOrEmpty(site.Query) || !string.IsNullOrEmpty(site.Fragment))
        {
            throw new ArgumentException("Confluence requires an absolute HTTPS site URL.", nameof(baseUrl));
        }

        var prefix = site.AbsoluteUri.TrimEnd('/');
        if (prefix.EndsWith("/wiki", StringComparison.OrdinalIgnoreCase)) prefix = prefix[..^5];
        return new Uri(prefix + "/wiki/rest/api/content/");
    }
}
