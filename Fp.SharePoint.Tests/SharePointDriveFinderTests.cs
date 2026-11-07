using System.Net;
using System.Text;
using Microsoft.Graph;

namespace Fp.SharePoint.Tests;

public sealed class SharePointDriveFinderTests
{
    [Fact]
    public void ToSiteIdentifiers_SiteUrl_HasTrailingColon()
    {
        Assert.Equal(
            ["contoso.sharepoint.com:/sites/Finance:"],
            SharePointDriveFinder.ToSiteIdentifiers("  https://contoso.sharepoint.com/sites/Finance/  "));
    }

    [Fact]
    public void ToSiteIdentifiers_DeepUrl_FallsBackToParents()
    {
        Assert.Equal(
            [
                "contoso.sharepoint.com:/sites/Finance/Shared%20Documents/Forms:",
                "contoso.sharepoint.com:/sites/Finance/Shared%20Documents:",
                "contoso.sharepoint.com:/sites/Finance:",
            ],
            SharePointDriveFinder.ToSiteIdentifiers("https://contoso.sharepoint.com/sites/Finance/Shared%20Documents/Forms"));
    }

    [Fact]
    public void ToSiteIdentifiers_RootSite_IsHostOnly() =>
        Assert.Equal(["contoso.sharepoint.com"], SharePointDriveFinder.ToSiteIdentifiers("https://contoso.sharepoint.com"));

    [Theory]
    [InlineData("")]
    [InlineData("not a url")]
    [InlineData("http://contoso.sharepoint.com/sites/Finance")]
    public void ToSiteIdentifiers_RejectsInvalidUrl(string url) =>
        Assert.Throws<ArgumentException>(() => SharePointDriveFinder.ToSiteIdentifiers(url));

    [Fact]
    public async Task GetDrives_ReturnsDrivesOfSite()
    {
        Uri? requested = null;
        var handler = new StubHandler(request =>
        {
            requested = request.RequestUri;
            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(
                    """{"value":[{"id":"b!1","name":"Documents","webUrl":"https://c/Shared%20Documents","driveType":"documentLibrary"},{"id":"b!2","name":"Invoices"}]}""",
                    Encoding.UTF8,
                    "application/json")
            };
        });
        var finder = new SharePointDriveFinder(new GraphServiceClient(new HttpClient(handler)));

        var drives = await finder.GetDrives("https://contoso.sharepoint.com/sites/Finance");

        Assert.Equal(2, drives.Count);
        Assert.Equal(new SharePointDrive("b!1", "Documents", "https://c/Shared%20Documents", "documentLibrary"), drives[0]);
        Assert.Equal("b!2", drives[1].Id);
        Assert.Contains("/v1.0/sites/contoso.sharepoint.com:/sites/Finance:/drives", requested!.AbsoluteUri);
    }

    [Fact]
    public async Task GetDrives_PageUrl_FallsBackToSitePath()
    {
        var requests = new List<string>();
        var handler = new StubHandler(request =>
        {
            requests.Add(request.RequestUri!.AbsoluteUri);
            return request.RequestUri.AbsoluteUri.Contains("Finance/SitePages")
                ? new HttpResponseMessage(HttpStatusCode.NotFound)
                {
                    Content = new StringContent("""{"error":{"code":"itemNotFound","message":"not a site"}}""", Encoding.UTF8, "application/json")
                }
                : new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent("""{"value":[{"id":"b!1","name":"Documents"}]}""", Encoding.UTF8, "application/json")
                };
        });
        var finder = new SharePointDriveFinder(new GraphServiceClient(new HttpClient(handler)));

        var drives = await finder.GetDrives("https://contoso.sharepoint.com/sites/Finance/SitePages");

        Assert.Single(drives);
        Assert.Equal(2, requests.Count);
        Assert.Contains("sites/Finance:/drives", requests[1]);
    }

    [Fact]
    public async Task GetDrives_SiteNotFound_Throws()
    {
        var handler = new StubHandler(_ => new HttpResponseMessage(HttpStatusCode.NotFound)
        {
            Content = new StringContent("""{"error":{"code":"itemNotFound","message":"nope"}}""", Encoding.UTF8, "application/json")
        });
        var finder = new SharePointDriveFinder(new GraphServiceClient(new HttpClient(handler)));

        await Assert.ThrowsAsync<Microsoft.Graph.Models.ODataErrors.ODataError>(
            () => finder.GetDrives("https://contoso.sharepoint.com/sites/Missing"));
    }

    private sealed class StubHandler(Func<HttpRequestMessage, HttpResponseMessage> respond) : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) =>
            Task.FromResult(respond(request));
    }
}
