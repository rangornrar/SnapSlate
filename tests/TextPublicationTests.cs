using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SnapSlate.Tests;

[TestClass]
public sealed class TextPublicationTests
{
    [TestMethod]
    public async Task Notion_TextPublication_SendsTitleMarkdownAndScopedAuthorization()
    {
        using var client = new HttpClient(new Handler(async request =>
        {
            Assert.AreEqual("https://api.notion.com/v1/pages", request.RequestUri!.AbsoluteUri);
            Assert.AreEqual("Bearer", request.Headers.Authorization!.Scheme);
            Assert.AreEqual("test-token", request.Headers.Authorization.Parameter);
            Assert.IsTrue(request.Headers.Contains("Notion-Version"));
            using var body = JsonDocument.Parse(await request.Content!.ReadAsStringAsync());
            Assert.AreEqual("parent", body.RootElement.GetProperty("parent").GetProperty("page_id").GetString());
            Assert.AreEqual("Guide", body.RootElement.GetProperty("properties").GetProperty("title").GetProperty("title")[0].GetProperty("text").GetProperty("content").GetString());
            Assert.AreEqual("Text-only", body.RootElement.GetProperty("markdown").GetString());
            return Json("{\"url\":\"https://notion.so/test\"}");
        }));
        Assert.AreEqual("https://notion.so/test", await TextPublicationService.NotionAsync(client, "test-token", "parent", "Guide", "Text-only"));
        Assert.IsNull(client.DefaultRequestHeaders.Authorization);
    }

    [TestMethod]
    [DataRow(401)]
    [DataRow(403)]
    [DataRow(429)]
    public async Task Publication_HttpFailure_PreservesStatusAndRetryWithoutResponseSecrets(int status)
    {
        using var client = new HttpClient(new Handler(_ =>
        {
            var response = new HttpResponseMessage((HttpStatusCode)status) { Content = new StringContent("private-token-and-path") };
            response.Headers.RetryAfter = new RetryConditionHeaderValue(TimeSpan.FromSeconds(30));
            return Task.FromResult(response);
        }));
        var failure = await Assert.ThrowsAsync<PublicationException>(() => TextPublicationService.NotionAsync(client, "test-token", "parent", "Guide", "Text"));
        Assert.AreEqual((HttpStatusCode)status, failure.StatusCode);
        Assert.AreEqual(TimeSpan.FromSeconds(30), failure.RetryAfter);
        Assert.IsFalse(failure.ToString().Contains("private-token-and-path"));
    }

    [TestMethod]
    public async Task Confluence_CloudV2_ResolvesSpaceAndPostsStorageText()
    {
        var methods = new List<HttpMethod>();
        using var client = new HttpClient(new Handler(async request =>
        {
            methods.Add(request.Method);
            Assert.AreEqual("Basic", request.Headers.Authorization!.Scheme);
            if (request.Method == HttpMethod.Get)
            {
                Assert.AreEqual("https://example.atlassian.net/wiki/api/v2/spaces?keys=MY%20SPACE&limit=2", request.RequestUri!.AbsoluteUri);
                return Json("{\"results\":[{\"id\":\"42\"}]}");
            }
            Assert.AreEqual("https://example.atlassian.net/wiki/api/v2/pages", request.RequestUri!.AbsoluteUri);
            using var body = JsonDocument.Parse(await request.Content!.ReadAsStringAsync());
            Assert.AreEqual("42", body.RootElement.GetProperty("spaceId").GetString());
            Assert.AreEqual("7", body.RootElement.GetProperty("parentId").GetString());
            Assert.AreEqual("storage", body.RootElement.GetProperty("body").GetProperty("representation").GetString());
            Assert.AreEqual("<p>Text-only</p>", body.RootElement.GetProperty("body").GetProperty("value").GetString());
            return Json("{\"_links\":{\"webui\":\"/spaces/MY/pages/99\"}}");
        }));
        Assert.AreEqual("https://example.atlassian.net/wiki/spaces/MY/pages/99", await TextPublicationService.ConfluenceAsync(client,
            "https://example.atlassian.net/wiki", "MY SPACE", "test@example.com", "token", "Guide", "<p>Text-only</p>", 7));
        CollectionAssert.AreEqual(new[] { HttpMethod.Get, HttpMethod.Post }, methods.ToArray());
        Assert.IsNull(client.DefaultRequestHeaders.Authorization);
    }

    private static HttpResponseMessage Json(string content) => new(HttpStatusCode.OK) { Content = new StringContent(content) };
    private sealed class Handler(Func<HttpRequestMessage, Task<HttpResponseMessage>> send) : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) => send(request);
    }
}
