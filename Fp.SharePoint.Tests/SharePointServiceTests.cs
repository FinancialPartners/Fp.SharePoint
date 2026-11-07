using System.Net;
using System.Text;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using Microsoft.Graph;

namespace Fp.SharePoint.Tests;

public sealed class SharePointServiceTests
{
    private const string DriveId = "drive-1";

    private sealed class FakeHandler(Func<HttpRequestMessage, HttpResponseMessage> respond) : HttpMessageHandler
    {
        public List<(HttpMethod Method, Uri Uri, string Body)> Requests { get; } = [];

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            var body = request.Content is null ? string.Empty : await request.Content.ReadAsStringAsync(cancellationToken);
            Requests.Add((request.Method, request.RequestUri!, body));
            return respond(request);
        }
    }

    private static HttpResponseMessage Json(string json, HttpStatusCode status = HttpStatusCode.OK) =>
        new(status) { Content = new StringContent(json, Encoding.UTF8, "application/json") };

    private static HttpResponseMessage NotFound() =>
        Json("""{"error":{"code":"itemNotFound","message":"not found"}}""", HttpStatusCode.NotFound);

    private static SharePointService CreateService(FakeHandler handler, long largeThreshold = 4 * 1024 * 1024)
    {
        var client = new GraphServiceClient(new HttpClient(handler));
        var options = Options.Create(new SharePointOptions { DriveId = DriveId, LargeFileThresholdBytes = largeThreshold });
        return new SharePointService(client, options, NullLogger<SharePointService>.Instance);
    }

    [Fact]
    public async Task UploadFile_PutsContentOnNormalizedPath_AndReturnsWebUrl()
    {
        var handler = new FakeHandler(_ => Json("""{"id":"1","webUrl":"https://contoso/file.xml"}"""));
        var service = CreateService(handler);
        using var stream = new MemoryStream("hello"u8.ToArray());

        var url = await service.UploadFile(stream, "\\Invoices\\2026/file.xml");

        Assert.Equal("https://contoso/file.xml", url);
        var request = Assert.Single(handler.Requests);
        Assert.Equal(HttpMethod.Put, request.Method);
        Assert.Contains($"/drives/{DriveId}/root:/Invoices/2026/file.xml:/content", request.Uri.AbsoluteUri);
        Assert.Equal("hello", request.Body);
    }

    [Fact]
    public async Task UploadFile_RewindsStreamBeforeUpload()
    {
        var handler = new FakeHandler(_ => Json("""{"id":"1"}"""));
        var service = CreateService(handler);
        using var stream = new MemoryStream("data"u8.ToArray());
        stream.Position = stream.Length;

        await service.UploadFile(stream, "a.txt");

        Assert.Equal("data", handler.Requests.Single().Body);
    }

    [Fact]
    public async Task UploadFile_LargeFile_CreatesUploadSession()
    {
        var handler = new FakeHandler(_ => Json("""{"error":{"code":"stop","message":"stop"}}""", HttpStatusCode.BadRequest));
        var service = CreateService(handler, largeThreshold: 2);
        using var stream = new MemoryStream(new byte[10]);

        await Assert.ThrowsAnyAsync<Exception>(() => service.UploadFile(stream, "big.zip"));

        var request = Assert.Single(handler.Requests);
        Assert.Equal(HttpMethod.Post, request.Method);
        Assert.Contains("createUploadSession", request.Uri.AbsoluteUri);
        Assert.Contains("replace", request.Body);
    }

    [Theory]
    [InlineData("")]
    [InlineData("  ")]
    public async Task UploadFile_RejectsEmptyPath(string path)
    {
        var service = CreateService(new FakeHandler(_ => Json("{}")));

        await Assert.ThrowsAnyAsync<ArgumentException>(() => service.UploadFile(new MemoryStream(), path));
    }

    [Fact]
    public async Task UploadFile_RejectsNullStream()
    {
        var service = CreateService(new FakeHandler(_ => Json("{}")));

        await Assert.ThrowsAsync<ArgumentNullException>(() => service.UploadFile(null!, "a.txt"));
    }

    [Fact]
    public async Task DownloadFile_ReturnsContent()
    {
        var handler = new FakeHandler(_ => new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent("content") });
        var service = CreateService(handler);

        await using var stream = await service.DownloadFile("a.txt");

        Assert.NotNull(stream);
        using var reader = new StreamReader(stream);
        Assert.Equal("content", await reader.ReadToEndAsync());
    }

    [Fact]
    public async Task DownloadFile_ReturnsNull_WhenNotFound()
    {
        var service = CreateService(new FakeHandler(_ => NotFound()));

        Assert.Null(await service.DownloadFile("missing.txt"));
    }

    [Fact]
    public async Task DeleteFile_SendsDelete()
    {
        var handler = new FakeHandler(_ => new HttpResponseMessage(HttpStatusCode.NoContent));
        var service = CreateService(handler);

        await service.DeleteFile("folder/a.txt");

        var request = Assert.Single(handler.Requests);
        Assert.Equal(HttpMethod.Delete, request.Method);
        Assert.Contains("root:/folder/a.txt", request.Uri.AbsoluteUri);
    }

    [Fact]
    public async Task DeleteFile_IgnoresNotFound()
    {
        var service = CreateService(new FakeHandler(_ => NotFound()));

        await service.DeleteFile("missing.txt");
    }

    [Fact]
    public async Task ListItems_Root_MapsItemsAndFollowsPaging()
    {
        var handler = new FakeHandler(request => request.RequestUri!.AbsoluteUri.Contains("page2")
            ? Json("""{"value":[{"name":"b.txt","size":5}]}""")
            : Json("""{"value":[{"name":"docs","folder":{"childCount":1}}],"@odata.nextLink":"https://graph.microsoft.com/v1.0/page2"}"""));
        var service = CreateService(handler);

        var items = await service.ListItems();

        Assert.Equal(2, items.Count);
        Assert.Contains(items, i => i.Name == "docs" && i.IsFolder);
        Assert.Contains(items, i => i.Name == "b.txt" && !i.IsFolder && i.Size == 5);
        Assert.Contains("/items/root/children", handler.Requests[0].Uri.AbsoluteUri);
    }

    [Fact]
    public async Task ListItems_Folder_UsesPath()
    {
        var handler = new FakeHandler(_ => Json("""{"value":[]}"""));
        var service = CreateService(handler);

        var items = await service.ListItems("\\Invoices/2026/");

        Assert.Empty(items);
        Assert.Contains("root:/Invoices/2026:/children", handler.Requests.Single().Uri.AbsoluteUri);
    }
}
