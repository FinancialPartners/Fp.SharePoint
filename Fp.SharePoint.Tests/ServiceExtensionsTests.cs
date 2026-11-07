using Azure.Core;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using Microsoft.Graph;

namespace Fp.SharePoint.Tests;

public sealed class ServiceExtensionsTests
{
    private sealed class FakeCredential : TokenCredential
    {
        public override AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken) =>
            new("token", DateTimeOffset.UtcNow.AddHours(1));

        public override ValueTask<AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken) =>
            new(GetToken(requestContext, cancellationToken));
    }

    private static ServiceCollection CreateServices(Dictionary<string, string?> settings)
    {
        var services = new ServiceCollection();
        services.AddLogging();
        services.AddSingleton<IConfiguration>(new ConfigurationBuilder().AddInMemoryCollection(settings).Build());
        return services;
    }

    [Fact]
    public void AddSharePointServices_BindsDefaultSection()
    {
        var services = CreateServices(new()
        {
            ["SharePointOptions:TenantId"] = "tenant",
            ["SharePointOptions:ClientId"] = "client",
            ["SharePointOptions:Secret"] = "secret",
            ["SharePointOptions:DriveId"] = "drive",
        });
        services.AddSharePointServices();
        using var provider = services.BuildServiceProvider();

        var options = provider.GetRequiredService<IOptions<SharePointOptions>>().Value;

        Assert.Equal("tenant", options.TenantId);
        Assert.Equal("client", options.ClientId);
        Assert.Equal("drive", options.DriveId);
        Assert.IsType<SharePointService>(provider.GetRequiredService<ISharePointService>());
    }

    [Fact]
    public void AddSharePointServices_BindsCustomSection()
    {
        var services = CreateServices(new()
        {
            ["Custom:TenantId"] = "t",
            ["Custom:ClientId"] = "c",
            ["Custom:DriveId"] = "d",
        });
        services.AddSharePointServices("Custom");
        using var provider = services.BuildServiceProvider();

        Assert.Equal("d", provider.GetRequiredService<IOptions<SharePointOptions>>().Value.DriveId);
    }

    [Fact]
    public void AddSharePointServices_WithDelegate_ConfiguresOptions()
    {
        var services = CreateServices([]);
        services.AddSharePointServices(o =>
        {
            o.TenantId = "t";
            o.ClientId = "c";
            o.DriveId = "d";
        });
        using var provider = services.BuildServiceProvider();

        Assert.Equal("c", provider.GetRequiredService<IOptions<SharePointOptions>>().Value.ClientId);
    }

    [Fact]
    public void Options_MissingRequiredValues_FailValidation()
    {
        var services = CreateServices(new() { ["SharePointOptions:TenantId"] = "t" });
        services.AddSharePointServices();
        using var provider = services.BuildServiceProvider();

        Assert.Throws<OptionsValidationException>(() => provider.GetRequiredService<IOptions<SharePointOptions>>().Value);
    }

    [Fact]
    public void Options_LargeFileThresholdOutOfRange_FailsValidation()
    {
        var services = CreateServices([]);
        services.AddSharePointServices(o =>
        {
            o.TenantId = "t";
            o.ClientId = "c";
            o.DriveId = "d";
            o.LargeFileThresholdBytes = 0;
        });
        using var provider = services.BuildServiceProvider();

        Assert.Throws<OptionsValidationException>(() => provider.GetRequiredService<IOptions<SharePointOptions>>().Value);
    }

    [Fact]
    public void AddSharePointServices_KeepsHostSuppliedCredentialAndGraphClient()
    {
        var credential = new FakeCredential();
        var graphClient = new GraphServiceClient(credential);
        var services = CreateServices([]);
        services.AddSingleton<TokenCredential>(credential);
        services.AddSingleton(graphClient);
        services.AddSharePointServices(o =>
        {
            o.TenantId = "t";
            o.ClientId = "c";
            o.DriveId = "d";
        });
        using var provider = services.BuildServiceProvider();

        Assert.Same(credential, provider.GetRequiredService<TokenCredential>());
        Assert.Same(graphClient, provider.GetRequiredService<GraphServiceClient>());
    }
}
