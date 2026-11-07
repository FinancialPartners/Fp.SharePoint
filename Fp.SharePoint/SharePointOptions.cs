using System.ComponentModel.DataAnnotations;

namespace Fp.SharePoint;

public class SharePointOptions
{
    public const string DefaultSectionName = "SharePointOptions";

    [Required]
    public string TenantId { get; set; } = string.Empty;

    [Required]
    public string ClientId { get; set; } = string.Empty;

    /// <summary>
    /// Client secret of the app registration. When empty, <c>DefaultAzureCredential</c> is used (e.g. managed identity).
    /// </summary>
    public string Secret { get; set; } = string.Empty;

    /// <summary>
    /// Not used anymore, kept so existing configuration keeps binding.
    /// </summary>
    public string SiteId { get; set; } = string.Empty;

    [Required]
    public string DriveId { get; set; } = string.Empty;

    /// <summary>
    /// Files larger than this (in bytes) are uploaded in chunks. Graph limits a simple upload to 4 MB.
    /// </summary>
    [Range(1, 250L * 1024 * 1024)]
    public long LargeFileThresholdBytes { get; set; } = 4 * 1024 * 1024;
}
