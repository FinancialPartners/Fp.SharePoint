# Fp.SharePoint

[![NuGet](https://img.shields.io/nuget/v/Fp.SharePoint.svg)](https://www.nuget.org/packages/Fp.SharePoint)
[![Build](https://github.com/FinancialPartners/Fp.SharePoint/actions/workflows/build.yml/badge.svg)](https://github.com/FinancialPartners/Fp.SharePoint/actions/workflows/build.yml)

Upload, download, delete and list files in a SharePoint document library through Microsoft Graph.

```
dotnet add package Fp.SharePoint
```

See the [package readme](Fp.SharePoint/README.md) for setup and usage.

## Finding your DriveId

`Fp.SharePoint.DriveFinder` is a small console tool that lists the document libraries of a SharePoint site with their `DriveId`:

```
dotnet run --project Fp.SharePoint.DriveFinder
```

It asks for your tenant id, optionally a client id and secret (leave the client id empty to sign in with your own account via Azure CLI / Visual Studio) and a site url such as `https://contoso.sharepoint.com/sites/Finance`.

## Releasing

Versions come from git tags ([MinVer](https://github.com/adamralph/minver)). Push a tag like `v1.0.0` and the release workflow tests, packs and publishes the package to nuget.org.

## License

[MIT](LICENSE)
