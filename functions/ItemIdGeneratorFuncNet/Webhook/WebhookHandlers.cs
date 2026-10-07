using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using ItemIdGenerator.Settings;
using ItemIdGenerator.SharePoint;

namespace ItemIdGenerator.Webhook;

public sealed class WebhookResult
{
    public int Status { get; init; }
    public string Body { get; init; } = string.Empty;
    public bool PlainText { get; init; }
    public IReadOnlyList<string> QueueMessages { get; init; } = Array.Empty<string>();
}

public sealed class SharePointWebhookHandler
{
    private static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web);

    public WebhookResult Handle(string? validationToken, string? body, string? configuredState)
    {
        if (!string.IsNullOrEmpty(validationToken))
        {
            return new WebhookResult
            {
                Status = 200,
                PlainText = true,
                Body = validationToken
            };
        }

        if (!IsUsableSecret(configuredState))
        {
            return JsonResult(500, new { error = "SHAREPOINT_WEBHOOK_CLIENT_STATE is not configured." });
        }

        if (!TryReadNotifications(body, out var notifications, out var error))
        {
            return JsonResult(400, new { error });
        }

        var messages = new List<string>();
        foreach (var entry in notifications)
        {
            if (!ClientStateMatches(configuredState!, entry.ClientState))
            {
                return JsonResult(403, new { error = "Webhook clientState does not match." });
            }

            var listId = SharePointUrls.ListIdFromResource(entry.Resource);
            if (listId.Length == 0)
            {
                return JsonResult(400, new { error = "Webhook resource must be a list GUID." });
            }

            var siteUrl = (entry.SiteUrl ?? string.Empty).Trim();
            if (siteUrl.Length == 0)
            {
                return JsonResult(400, new { error = "Webhook notification is missing siteUrl." });
            }

            messages.Add(JsonSerializer.Serialize(new
            {
                siteUrl,
                listId,
                subscriptionId = entry.SubscriptionId ?? string.Empty
            }, Json));
        }

        return new WebhookResult
        {
            Status = 200,
            Body = JsonSerializer.Serialize(new { accepted = messages.Count }, Json),
            QueueMessages = messages
        };
    }

    private static bool TryReadNotifications(string? body, out List<NotificationItem> notifications, out string error)
    {
        notifications = new List<NotificationItem>();
        error = string.Empty;
        if (string.IsNullOrWhiteSpace(body))
        {
            error = "Webhook body must be JSON.";
            return false;
        }

        try
        {
            using var document = JsonDocument.Parse(body);
            if (document.RootElement.ValueKind != JsonValueKind.Object
                || !document.RootElement.TryGetProperty("value", out var value)
                || value.ValueKind != JsonValueKind.Array)
            {
                error = "Webhook body must be a JSON object with a value array.";
                return false;
            }

            foreach (var item in value.EnumerateArray())
            {
                notifications.Add(new NotificationItem
                {
                    ClientState = StringOrNull(item, "clientState"),
                    Resource = StringOrNull(item, "resource"),
                    SiteUrl = StringOrNull(item, "siteUrl"),
                    SubscriptionId = StringOrNull(item, "subscriptionId")
                });
            }

            return true;
        }
        catch (JsonException)
        {
            error = "Webhook body must be JSON.";
            return false;
        }
    }

    private static string? StringOrNull(JsonElement item, string name)
    {
        if (!item.TryGetProperty(name, out var value) || value.ValueKind != JsonValueKind.String)
        {
            return null;
        }

        return value.GetString();
    }

    private static bool IsUsableSecret(string? value)
    {
        var text = (value ?? string.Empty).Trim();
        return text.Length > 0 && !text.Contains('<', StringComparison.Ordinal);
    }

    private static bool ClientStateMatches(string expected, string? received)
    {
        var left = Encoding.UTF8.GetBytes(expected);
        var right = Encoding.UTF8.GetBytes(received ?? string.Empty);
        return left.Length == right.Length && CryptographicOperations.FixedTimeEquals(left, right);
    }

    private static WebhookResult JsonResult(int status, object body)
    {
        return new WebhookResult
        {
            Status = status,
            Body = JsonSerializer.Serialize(body, Json)
        };
    }

    private sealed class NotificationItem
    {
        public string? ClientState { get; init; }
        public string? Resource { get; init; }
        public string? SiteUrl { get; init; }
        public string? SubscriptionId { get; init; }
    }
}

public sealed class RegisterWebhookHandler
{
    private static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web);

    public async Task<WebhookResult> HandleAsync(string? body, AppSettings settings, SharePointClient sharePoint)
    {
        if (!TryReadRequest(body, out var itemId, out var configSiteUrl, out var error))
        {
            return JsonResult(400, new { error });
        }

        if (!string.IsNullOrWhiteSpace(configSiteUrl))
        {
            settings = settings with { ConfigSiteUrl = configSiteUrl.Trim() };
        }

        try
        {
            var config = await sharePoint.GetConfigItemByIdAsync(settings, itemId);
            if (config is null)
            {
                return JsonResult(404, new { error = "Config row was not found." });
            }

            var result = await sharePoint.RegisterListWebhookAsync(settings, config);
            return JsonResult(200, new { subscriptionId = result.SubscriptionId, alreadyRegistered = result.AlreadyRegistered });
        }
        catch (Exception exception)
        {
            return JsonResult(502, new { error = exception.Message });
        }
    }

    private static bool TryReadRequest(string? body, out int itemId, out string? configSiteUrl, out string error)
    {
        itemId = 0;
        configSiteUrl = null;
        error = string.Empty;
        if (string.IsNullOrWhiteSpace(body))
        {
            error = "Request body must be JSON.";
            return false;
        }

        try
        {
            using var document = JsonDocument.Parse(body);
            if (document.RootElement.ValueKind != JsonValueKind.Object)
            {
                error = "Request body must be a JSON object.";
                return false;
            }

            if (!document.RootElement.TryGetProperty("configItemId", out var idProperty)
                || idProperty.ValueKind != JsonValueKind.Number
                || !idProperty.TryGetInt32(out itemId)
                || itemId <= 0)
            {
                error = "configItemId must be a positive integer.";
                return false;
            }

            if (document.RootElement.TryGetProperty("configSiteUrl", out var site) && site.ValueKind == JsonValueKind.String)
            {
                configSiteUrl = site.GetString();
            }

            return true;
        }
        catch (JsonException)
        {
            error = "Request body must be JSON.";
            return false;
        }
    }

    private static WebhookResult JsonResult(int status, object body)
    {
        return new WebhookResult
        {
            Status = status,
            Body = JsonSerializer.Serialize(body, Json)
        };
    }
}
