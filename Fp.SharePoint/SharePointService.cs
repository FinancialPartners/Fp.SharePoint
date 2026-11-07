using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

namespace Fp.SharePoint;

public class SharePointService(
    GraphServiceClient graphClient,
    IOptions<SharePointOptions> options,
    ILogger<SharePointService> logger) : ISharePointService
{
    private const int ChunkSize = 320 * 1024 * 10;
    private readonly SharePointOptions sharePointOptions = options.Value;

    public async Task<string?> UploadFile(Stream stream, string fileNameWithPath, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(stream);
        var path = NormalizePath(fileNameWithPath);

        if (stream.CanSeek)
        {
            stream.Seek(0, SeekOrigin.Begin);
        }

        try
        {
            var driveItem = stream.CanSeek && stream.Length > sharePointOptions.LargeFileThresholdBytes
                ? await UploadLargeFile(stream, path, cancellationToken)
                : await graphClient.Drives[sharePointOptions.DriveId].Root.ItemWithPath(path).Content
                    .PutAsync(stream, cancellationToken: cancellationToken);

            return driveItem?.WebUrl;
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            logger.LogError(ex, "Uploading {Path} to SharePoint failed", path);
            throw;
        }
    }

    public async Task<Stream?> DownloadFile(string fileNameWithPath, CancellationToken cancellationToken = default)
    {
        var path = NormalizePath(fileNameWithPath);

        try
        {
            return await graphClient.Drives[sharePointOptions.DriveId].Root.ItemWithPath(path).Content
                .GetAsync(cancellationToken: cancellationToken);
        }
        catch (ODataError ex) when (ex.ResponseStatusCode == 404)
        {
            return null;
        }
    }

    public async Task DeleteFile(string fileNameWithPath, CancellationToken cancellationToken = default)
    {
        var path = NormalizePath(fileNameWithPath);

        try
        {
            await graphClient.Drives[sharePointOptions.DriveId].Root.ItemWithPath(path)
                .DeleteAsync(cancellationToken: cancellationToken);
        }
        catch (ODataError ex) when (ex.ResponseStatusCode == 404)
        {
            logger.LogDebug("File {Path} does not exist, nothing to delete", path);
        }
    }

    public async Task<IReadOnlyList<SharePointItem>> ListItems(string folderPath = "", CancellationToken cancellationToken = default)
    {
        var path = folderPath.Replace('\\', '/').Trim('/');
        var drive = graphClient.Drives[sharePointOptions.DriveId];

        var page = path.Length == 0
            ? await drive.Items["root"].Children.GetAsync(cancellationToken: cancellationToken)
            : await drive.Root.ItemWithPath(path).Children.GetAsync(cancellationToken: cancellationToken);

        var items = new List<SharePointItem>();
        if (page is null)
        {
            return items;
        }

        var iterator = PageIterator<DriveItem, DriveItemCollectionResponse>.CreatePageIterator(
            graphClient,
            page,
            item =>
            {
                items.Add(new SharePointItem(item.Name ?? string.Empty, item.WebUrl, item.Size, item.Folder is not null, item.LastModifiedDateTime));
                return true;
            });
        await iterator.IterateAsync(cancellationToken);

        return items;
    }

    private async Task<DriveItem?> UploadLargeFile(Stream stream, string path, CancellationToken cancellationToken)
    {
        var body = new CreateUploadSessionPostRequestBody
        {
            Item = new DriveItemUploadableProperties
            {
                AdditionalData = new Dictionary<string, object> { ["@microsoft.graph.conflictBehavior"] = "replace" }
            }
        };

        var session = await graphClient.Drives[sharePointOptions.DriveId].Root.ItemWithPath(path).CreateUploadSession
            .PostAsync(body, cancellationToken: cancellationToken);

        var task = new LargeFileUploadTask<DriveItem>(session, stream, ChunkSize, graphClient.RequestAdapter);
        var result = await task.UploadAsync(cancellationToken: cancellationToken);

        return result.UploadSucceeded
            ? result.ItemResponse
            : throw new InvalidOperationException($"Chunked upload of '{path}' to SharePoint did not succeed.");
    }

    private static string NormalizePath(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        return path.Replace('\\', '/').Trim('/');
    }
}
