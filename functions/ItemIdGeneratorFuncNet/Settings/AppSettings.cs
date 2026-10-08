using System.Security.Cryptography;
using System.Text;

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
    public const string DefaultListTitle = "AutoGenFeatureConfiguration";

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

    public string ResolveClientState()
    {
        if (IsUsable(WebhookClientState))
        {
            return WebhookClientState.Trim();
        }

        if (!IsUsable(ClientSecret))
        {
            return string.Empty;
        }

        var hash = SHA256.HashData(Encoding.UTF8.GetBytes("request-number|" + ClientSecret));
        return Convert.ToHexString(hash)[..32];
    }

    public string ResolveNotificationUrl(Uri? requestUrl)
    {
        if (IsUsable(NotificationUrl))
        {
            return NotificationUrl.Trim();
        }

        if (requestUrl is null || !string.Equals(requestUrl.Scheme, "https", StringComparison.OrdinalIgnoreCase))
        {
            return string.Empty;
        }

        return requestUrl.GetLeftPart(UriPartial.Authority).TrimEnd('/') + "/api/spoWebhook";
    }

    private static bool IsUsable(string? value)
    {
        var text = (value ?? string.Empty).Trim();
        return text.Length > 0 && !text.Contains('<', StringComparison.Ordinal);
    }
}
