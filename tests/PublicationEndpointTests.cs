using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SnapSlate.Tests;

[TestClass]
public sealed class PublicationEndpointTests
{
    [TestMethod]
    [DataRow("http://example.atlassian.net")]
    [DataRow("file:///C:/site")]
    [DataRow("example.atlassian.net")]
    [DataRow("https://user:secret@example.atlassian.net")]
    [DataRow("https://example.atlassian.net?token=secret")]
    [DataRow("https://example.atlassian.net/#fragment")]
    public void Confluence_UnsafeUrl_RejectsBeforeSendingRequest(string url)
    {
        Assert.ThrowsExactly<ArgumentException>(() => PublicationEndpoint.Confluence(url));
    }

    [TestMethod]
    [DataRow("https://example.atlassian.net")]
    [DataRow("https://example.atlassian.net/")]
    [DataRow("https://example.atlassian.net/wiki")]
    public void Confluence_HttpsSite_BuildsSingleWikiPath(string url)
    {
        Assert.AreEqual("https://example.atlassian.net/wiki/rest/api/content/", PublicationEndpoint.Confluence(url).AbsoluteUri);
    }
}
