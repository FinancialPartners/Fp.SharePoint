using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models.ODataErrors;

namespace Fp.SharePoint;

public record SharePointDrive(string Id, string Name, string? WebUrl, string? DriveType);

/// <summary>
/// Finds the document libraries (drives) of a SharePoint site, so you can look up the <c>DriveId</c> for <see cref="SharePointOptions"/>.
/// </summary>
public class SharePointDriveFinder(GraphServiceClient graphClient)
{
    /// <summary>
    /// Creates a credential: a client secret credential when a secret is given, otherwise <c>DefaultAzureCredential</c>
    /// (Azure CLI, Visual Studio, managed identity, ...).
    /// </summary>
    public static TokenCredential CreateCredential(string tenantId, string? clientId, string? secret) =>
        string.IsNullOrWhiteSpace(secret)
            ? new DefaultAzureCredential(new DefaultAzureCredentialOptions { TenantId = tenantId })
            : new ClientSecretCredential(tenantId, clientId, secret);

    public static SharePointDriveFinder Create(TokenCredential credential) =>
        new(new GraphServiceClient(credential, [".default"]));

    /// <summary>
    /// Lists the drives of a site, e.g. <c>https://contoso.sharepoint.com/sites/Finance</c>.
    /// </summary>
    public async Task<IReadOnlyList<SharePointDrive>> GetDrives(string siteUrl, CancellationToken cancellationToken = default)
    {
        ODataError? lastError = null;

        // The url may point to a page or library inside the site, so fall back to parent paths until a site is found.
        foreach (var siteIdentifier in ToSiteIdentifiers(siteUrl))
        {
            try
            {
                // A raw url keeps ':' and '/' literal; the SDK would percent-encode them in the path segment.
                var url = $"{graphClient.RequestAdapter.BaseUrl!.TrimEnd('/')}/sites/{siteIdentifier}/drives";
                var response = await graphClient.Sites[siteIdentifier].Drives.WithUrl(url).GetAsync(cancellationToken: cancellationToken);

                return (response?.Value ?? [])
                    .Where(d => d.Id is not null)
                    .Select(d => new SharePointDrive(d.Id!, d.Name ?? string.Empty, d.WebUrl, d.DriveType))
                    .ToList();
            }
            catch (ODataError ex) when (ex.ResponseStatusCode is 404 or 400)
            {
                lastError = ex;
            }
        }

        throw lastError!;
    }

    /// <summary>
    /// Converts a site url to the Graph site identifiers to try, most specific first: <c>host:/sites/name:</c>.
    /// The trailing colon is required by Graph when a segment such as <c>/drives</c> follows. The root site is just the host.
    /// </summary>
    public static IReadOnlyList<string> ToSiteIdentifiers(string siteUrl)
    {
        if (!Uri.TryCreate(siteUrl?.Trim(), UriKind.Absolute, out var uri)
            || uri.Scheme != Uri.UriSchemeHttps
            || string.IsNullOrEmpty(uri.Host))
        {
            throw new ArgumentException("Enter a valid https site url, e.g. https://contoso.sharepoint.com/sites/Finance", nameof(siteUrl));
        }

        var segments = uri.AbsolutePath.Split('/', StringSplitOptions.RemoveEmptyEntries);
        if (segments.Length == 0)
        {
            return [uri.Host];
        }

        // Down to two segments (sites/name, teams/name); anything shorter is not a site path.
        var identifiers = new List<string>();
        for (var count = segments.Length; count >= Math.Min(2, segments.Length); count--)
        {
            identifiers.Add($"{uri.Host}:/{string.Join('/', segments.Take(count))}:");
        }

        return identifiers;
    }
}
