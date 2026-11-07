# Fp.SharePoint

Upload, download, delete and list files in a SharePoint document library through Microsoft Graph.

## How to use
1. Add the NuGet package to your project.
2. In your `ConfigureServices` add:
```csharp
services.AddSharePointServices();                 // binds section "SharePointOptions"
services.AddSharePointServices("MySection");      // custom section name
services.AddSharePointServices(o => { o.TenantId = "..."; o.ClientId = "..."; o.DriveId = "..."; });
```
3. Add the configuration section to your appsettings:
```json
"SharePointOptions": {
  "TenantId": "{Your TenantId}",
  "ClientId": "{Your ClientId}",
  "Secret": "{Secret}",
  "DriveId": "ID of the document library where you want to store files"
}
```
   `SiteId` is no longer needed. Leave `Secret` empty to use `DefaultAzureCredential` (for example managed identity). Keep the secret out of source control (user secrets, Key Vault).
4. In Azure create an app registration:
   - Microsoft Graph -> Application permissions -> `Sites.ReadWrite.All` (or `Sites.Selected` for a single site)
   - Create a secret for the app registration
5. Inject `ISharePointService`:
```csharp
var url = await sharePoint.UploadFile(stream, "Invoices/2026/invoice.xml");
await using var file = await sharePoint.DownloadFile("Invoices/2026/invoice.xml");
var items = await sharePoint.ListItems("Invoices/2026");
await sharePoint.DeleteFile("Invoices/2026/invoice.xml");
```

Files above `LargeFileThresholdBytes` (default 4 MB) are uploaded in chunks. Register your own `TokenCredential` or `GraphServiceClient` before calling `AddSharePointServices` to override the defaults.
