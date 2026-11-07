using Azure.Core;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Microsoft.Extensions.Options;
using Microsoft.Graph;

namespace Fp.SharePoint;

public static class ServiceExtensions
{
    /// <summary>
    /// Registers the SharePoint services, binding <see cref="SharePointOptions"/> from the given configuration section.
    /// </summary>
    public static IServiceCollection AddSharePointServices(
        this IServiceCollection services,
        string sectionName = SharePointOptions.DefaultSectionName)
    {
        services.AddOptions<SharePointOptions>()
            .BindConfiguration(sectionName)
            .ValidateDataAnnotations()
            .ValidateOnStart();

        return services.AddSharePointCore();
    }

    /// <summary>
    /// Registers the SharePoint services, configuring <see cref="SharePointOptions"/> in code.
    /// </summary>
    public static IServiceCollection AddSharePointServices(
        this IServiceCollection services,
        Action<SharePointOptions> configure)
    {
        services.AddOptions<SharePointOptions>()
            .Configure(configure)
            .ValidateDataAnnotations()
            .ValidateOnStart();

        return services.AddSharePointCore();
    }

    private static IServiceCollection AddSharePointCore(this IServiceCollection services)
    {
        // TryAdd lets the host application supply its own credential or Graph client.
        services.TryAddSingleton<TokenCredential>(sp => CreateCredential(sp.GetRequiredService<IOptions<SharePointOptions>>().Value));
        services.TryAddSingleton(sp => new GraphServiceClient(sp.GetRequiredService<TokenCredential>(), [".default"]));
        services.TryAddSingleton<ISharePointService, SharePointService>();

        return services;
    }

    private static TokenCredential CreateCredential(SharePointOptions options) =>
        SharePointDriveFinder.CreateCredential(options.TenantId, options.ClientId, options.Secret);
}
