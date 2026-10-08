namespace ItemIdGenerator.Settings;

public sealed record AppSettings(
    string StorageConnectionString,
    string ConfigSiteUrl,
    string ConfigListTitle,
    string TenantId,
    string ClientId,
    string ClientSecret,
    string WebhookClientState,
    string NotificationUrl)
{
    public const string DefaultListTitle = "RequestNumberConfig";

    public static AppSettings FromEnvironment()
    {
        var title = Environment.GetEnvironmentVariable("NUMBERING_CONFIG_LIST_TITLE");
        return new AppSettings(
            Environment.GetEnvironmentVariable("AzureWebJobsStorage") ?? string.Empty,
            string.Empty,
            string.IsNullOrWhiteSpace(title) ? DefaultListTitle : title.Trim(),
            Environment.GetEnvironmentVariable("SHAREPOINT_TENANT_ID") ?? string.Empty,
            Environment.GetEnvironmentVariable("SHAREPOINT_CLIENT_ID") ?? string.Empty,
            Environment.GetEnvironmentVariable("SHAREPOINT_CLIENT_SECRET") ?? string.Empty,
            Environment.GetEnvironmentVariable("SHAREPOINT_WEBHOOK_CLIENT_STATE") ?? string.Empty,
            Environment.GetEnvironmentVariable("SPO_WEBHOOK_NOTIFICATION_URL") ?? string.Empty);
    }
}
