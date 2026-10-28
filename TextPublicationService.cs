using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace SnapSlate;

internal sealed class PublicationException(HttpStatusCode statusCode, TimeSpan? retryAfter = null)
    : Exception("Publication HTTP " + (int)statusCode)
{
    public HttpStatusCode StatusCode { get; } = statusCode;
    public TimeSpan? RetryAfter { get; } = retryAfter;
}

internal static class TextPublicationService
{
    public static async Task<string?> NotionAsync(HttpClient client, string token, string parentPageId, string title, string markdown)
    {
        using var request = new HttpRequestMessage(HttpMethod.Post, "https://api.notion.com/v1/pages");
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        request.Headers.TryAddWithoutValidation("Notion-Version", "2026-03-11");
        request.Content = JsonContent.Create(new
        {
            parent = new { page_id = parentPageId },
            properties = new { title = new { title = new[] { new { text = new { content = title } } } } },
            markdown
        });
        using var response = await client.SendAsync(request);
        EnsureSuccess(response);
        using var json = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
        return json.RootElement.TryGetProperty("url", out var url) ? url.GetString() : null;
    }

    public static async Task<string?> ConfluenceAsync(HttpClient client, string baseUrl, string spaceKey,
        string email, string token, string title, string storage, long? parentPageId)
    {
        var endpoint = PublicationEndpoint.ConfluenceV2(baseUrl);
        var authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(Encoding.UTF8.GetBytes($"{email}:{token}")));
        using var findSpace = new HttpRequestMessage(HttpMethod.Get, new Uri(endpoint, "spaces?keys=" + Uri.EscapeDataString(spaceKey) + "&limit=2"));
        findSpace.Headers.Authorization = authorization;
        using var spaceResponse = await client.SendAsync(findSpace);
        EnsureSuccess(spaceResponse);
        using var spaces = JsonDocument.Parse(await spaceResponse.Content.ReadAsStringAsync());
        var results = spaces.RootElement.GetProperty("results");
        if (results.GetArrayLength() != 1) throw new PublicationException(HttpStatusCode.NotFound);
        var spaceId = results[0].GetProperty("id").GetString();
        using var request = new HttpRequestMessage(HttpMethod.Post, new Uri(endpoint, "pages"));
        request.Headers.Authorization = authorization;
        request.Content = JsonContent.Create(new
        {
            spaceId, status = "current", title,
            parentId = parentPageId?.ToString(System.Globalization.CultureInfo.InvariantCulture),
            body = new { representation = "storage", value = storage }
        }, options: new JsonSerializerOptions { DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull });
        using var response = await client.SendAsync(request);
        EnsureSuccess(response);
        using var page = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
        if (!page.RootElement.TryGetProperty("_links", out var links) || !links.TryGetProperty("webui", out var webUi)) return null;
        var wikiRoot = endpoint.AbsoluteUri[..^"api/v2/".Length];
        return new Uri(new Uri(wikiRoot), (webUi.GetString() ?? string.Empty).TrimStart('/')).AbsoluteUri;
    }

    private static void EnsureSuccess(HttpResponseMessage response)
    {
        if (response.IsSuccessStatusCode) return;
        var retry = response.Headers.RetryAfter;
        throw new PublicationException(response.StatusCode, retry?.Delta ?? (retry?.Date is { } date ? date - DateTimeOffset.UtcNow : null));
    }
}
