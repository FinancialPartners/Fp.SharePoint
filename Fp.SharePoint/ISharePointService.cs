namespace Fp.SharePoint;

public interface ISharePointService
{
    /// <summary>
    /// Uploads a stream to the configured drive, replacing an existing file. Returns the web URL of the file.
    /// </summary>
    Task<string?> UploadFile(Stream stream, string fileNameWithPath, CancellationToken cancellationToken = default);

    /// <summary>
    /// Downloads a file, returns null when it does not exist.
    /// </summary>
    Task<Stream?> DownloadFile(string fileNameWithPath, CancellationToken cancellationToken = default);

    Task DeleteFile(string fileNameWithPath, CancellationToken cancellationToken = default);

    /// <summary>
    /// Lists the items in a folder. An empty path lists the root of the drive.
    /// </summary>
    Task<IReadOnlyList<SharePointItem>> ListItems(string folderPath = "", CancellationToken cancellationToken = default);
}

public record SharePointItem(string Name, string? WebUrl, long? Size, bool IsFolder, DateTimeOffset? LastModified);
