using Fp.SharePoint;

Console.WriteLine("SharePoint drive finder");
Console.WriteLine("Lists the document libraries of a site with their DriveId.");
Console.WriteLine();

var tenantId = Prompt("Tenant id");
var clientId = Prompt("Client id (empty = use your own login: Azure CLI / Visual Studio)", required: false);
var secret = string.IsNullOrWhiteSpace(clientId) ? null : PromptSecret("Client secret");

SharePointDriveFinder finder;
try
{
    finder = SharePointDriveFinder.Create(SharePointDriveFinder.CreateCredential(tenantId, clientId, secret));
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Could not create credentials: {ex.Message}");
    return 1;
}

while (true)
{
    Console.WriteLine();
    var siteUrl = Prompt("SharePoint site url (e.g. https://contoso.sharepoint.com/sites/Finance, empty = quit)", required: false);
    if (string.IsNullOrWhiteSpace(siteUrl))
    {
        return 0;
    }

    try
    {
        var drives = await finder.GetDrives(siteUrl);
        if (drives.Count == 0)
        {
            Console.WriteLine("No drives found.");
            continue;
        }

        Console.WriteLine();
        foreach (var drive in drives)
        {
            Console.WriteLine($"{drive.Name}");
            Console.WriteLine($"  DriveId: {drive.Id}");
            Console.WriteLine($"  Url:     {drive.WebUrl}");
        }
    }
    catch (ArgumentException ex)
    {
        Console.Error.WriteLine(ex.Message);
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"Failed to get drives: {ex.Message}");
    }
}

static string Prompt(string label, bool required = true)
{
    while (true)
    {
        Console.Write($"{label}: ");
        var value = Console.ReadLine()?.Trim() ?? string.Empty;
        if (value.Length > 0 || !required)
        {
            return value;
        }
    }
}

static string PromptSecret(string label)
{
    Console.Write($"{label}: ");
    var secret = new System.Text.StringBuilder();
    while (true)
    {
        var key = Console.ReadKey(intercept: true);
        if (key.Key == ConsoleKey.Enter)
        {
            Console.WriteLine();
            return secret.ToString();
        }

        if (key.Key == ConsoleKey.Backspace)
        {
            if (secret.Length > 0)
            {
                secret.Length--;
            }
        }
        else if (!char.IsControl(key.KeyChar))
        {
            secret.Append(key.KeyChar);
        }
    }
}
