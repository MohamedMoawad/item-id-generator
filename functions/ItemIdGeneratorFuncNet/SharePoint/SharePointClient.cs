using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using ItemIdGenerator.Numbering;
using ItemIdGenerator.Settings;

namespace ItemIdGenerator.SharePoint;

public sealed class RegisteredWebhook
{
    public string SubscriptionId { get; init; } = string.Empty;
    public bool AlreadyRegistered { get; init; }
}

public sealed class SharePointClient
{
    private readonly HttpClient _http;
    private readonly object _cacheGate = new();
    private string _token = string.Empty;
    private string _tokenHost = string.Empty;
    private DateTimeOffset _tokenExpiresAt;

    public SharePointClient(HttpClient http)
    {
        _http = http;
    }

    public async Task<NumberingConfig?> GetEnabledConfigAsync(AppSettings settings, string webOrSiteUrl, string listId)
    {
        var guid = SharePointUrls.NormalizeGuid(listId);
        if (guid.Length == 0)
        {
            throw new InvalidOperationException("Target list id must be a GUID.");
        }

        var siteUrl = await ResolveConfigSiteAsync(settings, webOrSiteUrl);
        var url = SharePointUrls.BuildConfigItemsUrl(siteUrl, ConfigListTitle(settings));
        using var response = await SendAsync(settings, HttpMethod.Get, url, siteUrl, null, null, false);
        var text = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
        {
            throw HttpError(response, text, "config list read");
        }

        using var document = JsonDocument.Parse(text);
        var matches = SharePointUrls.CollectionItems(document.RootElement)
            .Select(SharePointUrls.NormalizeConfigItem)
            .Where(item => item is not null && item.IsActive && item.TargetListGuid == guid)
            .Cast<NumberingConfig>()
            .ToList();
        if (matches.Count > 1)
        {
            throw new InvalidOperationException($"More than one active numbering config targets list {guid}.");
        }

        return matches.FirstOrDefault()?.WithConfigSite(siteUrl);
    }

    public async Task<string> ResolveConfigSiteAsync(AppSettings settings, string webOrSiteUrl)
    {
        var web = SharePointUrls.SiteRoot(webOrSiteUrl);
        var url = $"{web}/_api/site?$select=Url";
        using var response = await SendAsync(settings, HttpMethod.Get, url, web, null, null, false);
        var text = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
        {
            throw HttpError(response, text, "site collection lookup");
        }

        return SharePointUrls.ReadSiteCollectionUrl(text, web);
    }

    public async Task<NumberingConfig?> GetConfigItemByIdAsync(AppSettings settings, int itemId)
    {
        var siteUrl = ConfigSiteUrl(settings);
        var url = SharePointUrls.BuildConfigItemUrl(siteUrl, ConfigListTitle(settings), itemId);
        using var response = await SendAsync(settings, HttpMethod.Get, url, siteUrl, null, null, false);
        var text = await response.Content.ReadAsStringAsync();
        if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
        {
            return null;
        }

        if (!response.IsSuccessStatusCode)
        {
            throw HttpError(response, text, "config item read");
        }

        using var document = JsonDocument.Parse(text);
        return SharePointUrls.NormalizeConfigItem(document.RootElement)?.WithConfigSite(siteUrl);
    }

    public async Task<IReadOnlyList<int>> ListUnnumberedItemsAsync(AppSettings settings, NumberingConfig config, int? explicitItemId)
    {
        SharePointUrls.AssertInternalName(config.NumberColumnInternalName);
        if (string.IsNullOrWhiteSpace(config.TargetSiteUrl))
        {
            throw new InvalidOperationException("Target list URL is missing or is not an https list URL.");
        }

        if (explicitItemId is int itemId)
        {
            var item = await GetListItemAsync(settings, config.TargetSiteUrl, config.TargetListGuid, itemId, config.NumberColumnInternalName);
            if (item is null || !item.Value.Blank)
            {
                return Array.Empty<int>();
            }

            return new[] { item.Value.Id };
        }

        var url = SharePointUrls.BuildRecentItemsUrl(config.TargetSiteUrl, config.TargetListGuid, config.NumberColumnInternalName);
        using var response = await SendAsync(settings, HttpMethod.Get, url, config.TargetSiteUrl, null, null, false);
        var text = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
        {
            throw HttpError(response, text, "recent items read");
        }

        using var document = JsonDocument.Parse(text);
        return SharePointUrls.CollectionItems(document.RootElement)
            .Where(item => SharePointUrls.IsBlankNumber(item, config.NumberColumnInternalName) && TryItemId(item, out _))
            .Select(item =>
            {
                TryItemId(item, out var id);
                return id;
            })
            .OrderBy(id => id)
            .ToArray();
    }

    public async Task<bool> WriteNumberIfBlankAsync(AppSettings settings, NumberingConfig config, int itemId, string code)
    {
        SharePointUrls.AssertInternalName(config.NumberColumnInternalName);
        for (var attempt = 1; attempt <= 3; attempt++)
        {
            var item = await GetListItemAsync(settings, config.TargetSiteUrl, config.TargetListGuid, itemId, config.NumberColumnInternalName);
            if (item is null)
            {
                throw new InvalidOperationException($"List item {itemId} was not found.");
            }

            if (!item.Value.Blank)
            {
                return false;
            }

            var body = JsonSerializer.Serialize(new Dictionary<string, string> { [config.NumberColumnInternalName] = code });
            using var response = await SendAsync(
                settings,
                HttpMethod.Post,
                SharePointUrls.BuildTargetItemUrl(config.TargetSiteUrl, config.TargetListGuid, itemId),
                config.TargetSiteUrl,
                body,
                item.Value.Etag,
                true);
            if ((int)response.StatusCode == 412)
            {
                continue;
            }

            if (!response.IsSuccessStatusCode)
            {
                throw HttpError(response, await response.Content.ReadAsStringAsync(), "list item update");
            }

            return true;
        }

        throw new InvalidOperationException($"Could not write the request number on item {itemId}.");
    }

    public IConfigStore CreateConfigStore(AppSettings settings, NumberingConfig config)
    {
        var scoped = string.IsNullOrWhiteSpace(config.ConfigSiteUrl)
            ? settings
            : settings with { ConfigSiteUrl = config.ConfigSiteUrl };
        return new ConfigStore(this, scoped, config.Id);
    }

    public async Task<RegisteredWebhook> RegisterListWebhookAsync(AppSettings settings, NumberingConfig config)
    {
        if (!config.IsActive)
        {
            throw new InvalidOperationException("Activate the config row before registering a webhook.");
        }

        if (!string.IsNullOrWhiteSpace(config.WebhookSubscriptionId))
        {
            return new RegisteredWebhook { SubscriptionId = config.WebhookSubscriptionId, AlreadyRegistered = true };
        }

        var notificationUrl = settings.NotificationUrl.Trim();
        var clientState = settings.WebhookClientState.Trim();
        if (notificationUrl.Length == 0 || notificationUrl.Contains('<', StringComparison.Ordinal))
        {
            throw new InvalidOperationException("Set SPO_WEBHOOK_NOTIFICATION_URL to the spoWebhook URL, including the function key.");
        }

        if (clientState.Length == 0 || clientState.Contains('<', StringComparison.Ordinal))
        {
            throw new InvalidOperationException("Set SHAREPOINT_WEBHOOK_CLIENT_STATE before registering a webhook.");
        }

        if (string.IsNullOrWhiteSpace(config.TargetListUrl) || string.IsNullOrWhiteSpace(config.TargetListGuid))
        {
            throw new InvalidOperationException("The config row needs TargetListUrl and TargetListGuid.");
        }

        var location = SharePointUrls.DeriveWebFromListUrl(config.TargetListUrl);
        var expiration = DateTimeOffset.UtcNow.AddDays(170).ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
        var body = JsonSerializer.Serialize(new Dictionary<string, string>
        {
            ["resource"] = $"{location.WebAbsoluteUrl}/_api/web/lists('{config.TargetListGuid}')",
            ["notificationUrl"] = notificationUrl,
            ["expirationDateTime"] = expiration,
            ["clientState"] = clientState
        });
        using var response = await SendAsync(
            settings,
            HttpMethod.Post,
            $"{location.WebAbsoluteUrl}/_api/web/lists(guid'{config.TargetListGuid}')/subscriptions",
            location.WebAbsoluteUrl,
            body,
            null,
            false);
        var text = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
        {
            throw HttpError(response, text, "webhook registration");
        }

        var subscriptionId = ReadSubscriptionId(text);
        if (subscriptionId.Length == 0)
        {
            throw new InvalidOperationException("SharePoint did not return a webhook subscription id.");
        }

        var saved = await MergeFieldsAsync(settings, config.Id, config.Etag, new Dictionary<string, string>
        {
            ["WebhookSubscriptionId"] = subscriptionId
        });
        if (!saved)
        {
            throw new InvalidOperationException("The webhook was created but the config row changed. Register it again.");
        }

        return new RegisteredWebhook { SubscriptionId = subscriptionId, AlreadyRegistered = false };
    }

    private async Task<bool> MergeCounterAsync(AppSettings settings, int itemId, string etag, long currentCount, string lastResetDate)
    {
        var siteUrl = ConfigSiteUrl(settings);
        var body = JsonSerializer.Serialize(new Dictionary<string, object>
        {
            ["CurrentCount"] = currentCount,
            ["LastResetDate"] = lastResetDate
        });
        using var response = await SendAsync(settings, HttpMethod.Post, SharePointUrls.BuildConfigItemUrl(siteUrl, ConfigListTitle(settings), itemId), siteUrl, body, etag, true);
        if ((int)response.StatusCode == 412)
        {
            return false;
        }

        if (!response.IsSuccessStatusCode)
        {
            throw HttpError(response, await response.Content.ReadAsStringAsync(), "config item update");
        }

        return true;
    }

    private async Task<bool> MergeFieldsAsync(AppSettings settings, int itemId, string etag, Dictionary<string, string> fields)
    {
        var siteUrl = ConfigSiteUrl(settings);
        var body = JsonSerializer.Serialize(fields);
        using var response = await SendAsync(
            settings,
            HttpMethod.Post,
            SharePointUrls.BuildConfigItemUrl(siteUrl, ConfigListTitle(settings), itemId),
            siteUrl,
            body,
            string.IsNullOrEmpty(etag) ? "*" : etag,
            true);
        if ((int)response.StatusCode == 412)
        {
            return false;
        }

        if (!response.IsSuccessStatusCode)
        {
            throw HttpError(response, await response.Content.ReadAsStringAsync(), "config field update");
        }

        return true;
    }

    private async Task<(int Id, string Etag, bool Blank)?> GetListItemAsync(AppSettings settings, string siteUrl, string listId, int itemId, string fieldName)
    {
        var url = SharePointUrls.BuildTargetItemUrl(siteUrl, listId, itemId) + "?$select=Id," + fieldName;
        using var response = await SendAsync(settings, HttpMethod.Get, url, siteUrl, null, null, false);
        var text = await response.Content.ReadAsStringAsync();
        if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
        {
            return null;
        }

        if (!response.IsSuccessStatusCode)
        {
            throw HttpError(response, text, "list item read");
        }

        using var document = JsonDocument.Parse(text);
        if (!TryItemId(document.RootElement, out var id))
        {
            return null;
        }

        var etag = SharePointUrls.ReadEtag(document.RootElement);
        return (id, etag.Length == 0 ? "*" : etag, SharePointUrls.IsBlankNumber(document.RootElement, fieldName));
    }

    private async Task<HttpResponseMessage> SendAsync(AppSettings settings, HttpMethod method, string url, string siteUrl, string? jsonBody, string? etag, bool merge)
    {
        var token = await GetAccessTokenAsync(settings, siteUrl);
        using var request = new HttpRequestMessage(method, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        request.Headers.TryAddWithoutValidation("Accept", "application/json;odata=minimalmetadata");
        if (jsonBody is not null)
        {
            request.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");
            request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json;odata=nometadata");
        }

        if (merge)
        {
            request.Headers.TryAddWithoutValidation("IF-MATCH", etag ?? "*");
            request.Headers.TryAddWithoutValidation("X-HTTP-Method", "MERGE");
        }

        return await _http.SendAsync(request);
    }

    private async Task<string> GetAccessTokenAsync(AppSettings settings, string siteUrl)
    {
        var identity = AssertAppIdentity(settings);
        var host = SharePointUrls.AssertHttpsUrl(siteUrl, "SharePoint site URL").Host;
        lock (_cacheGate)
        {
            if (_token.Length > 0 && _tokenHost == host && _tokenExpiresAt > DateTimeOffset.UtcNow.AddMinutes(1))
            {
                return _token;
            }
        }

        var form = new Dictionary<string, string>
        {
            ["client_id"] = identity.ClientId,
            ["client_secret"] = identity.ClientSecret,
            ["scope"] = $"https://{host}/.default",
            ["grant_type"] = "client_credentials"
        };
        using var request = new HttpRequestMessage(HttpMethod.Post, $"https://login.microsoftonline.com/{identity.TenantId}/oauth2/v2.0/token")
        {
            Content = new FormUrlEncodedContent(form)
        };
        using var response = await _http.SendAsync(request);
        var text = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
        {
            throw HttpError(response, text, "token request");
        }

        using var document = JsonDocument.Parse(text);
        var token = document.RootElement.TryGetProperty("access_token", out var access) ? access.GetString() : null;
        if (string.IsNullOrEmpty(token))
        {
            throw new InvalidOperationException("SharePoint token response did not include an access token.");
        }

        var expiresIn = 3600;
        if (document.RootElement.TryGetProperty("expires_in", out var expires) && expires.TryGetInt32(out var parsed))
        {
            expiresIn = parsed;
        }

        lock (_cacheGate)
        {
            _token = token;
            _tokenHost = host;
            _tokenExpiresAt = DateTimeOffset.UtcNow.AddSeconds(expiresIn);
        }

        return token;
    }

    private static (string TenantId, string ClientId, string ClientSecret) AssertAppIdentity(AppSettings settings)
    {
        if (SharePointUrls.NormalizeGuid(settings.TenantId).Length == 0
            || SharePointUrls.NormalizeGuid(settings.ClientId).Length == 0
            || string.IsNullOrEmpty(settings.ClientSecret)
            || settings.ClientSecret.Contains('<', StringComparison.Ordinal))
        {
            throw new InvalidOperationException("Set SHAREPOINT_TENANT_ID, SHAREPOINT_CLIENT_ID, and SHAREPOINT_CLIENT_SECRET. Do not commit the secret.");
        }

        return (settings.TenantId.Trim(), settings.ClientId.Trim(), settings.ClientSecret);
    }

    private static string ConfigSiteUrl(AppSettings settings)
    {
        var siteUrl = settings.ConfigSiteUrl.Trim();
        if (siteUrl.Length == 0 || siteUrl.Contains('<', StringComparison.Ordinal))
        {
            throw new InvalidOperationException("The site collection URL is missing. RequestNumberConfig is read on the site collection that owns the list.");
        }

        return SharePointUrls.SiteRoot(siteUrl);
    }

    private static string ConfigListTitle(AppSettings settings)
    {
        var title = settings.ConfigListTitle.Trim();
        if (title.Length == 0 || title.Contains('\r') || title.Contains('\n'))
        {
            throw new InvalidOperationException("NUMBERING_CONFIG_LIST_TITLE is invalid.");
        }

        return title;
    }

    private static bool TryItemId(JsonElement item, out int id)
    {
        id = 0;
        return item.ValueKind == JsonValueKind.Object
            && item.TryGetProperty("Id", out var property)
            && property.ValueKind == JsonValueKind.Number
            && property.TryGetInt32(out id);
    }

    private static string ReadSubscriptionId(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        using var document = JsonDocument.Parse(text);
        var root = document.RootElement;
        var direct = IdString(root);
        if (direct.Length > 0)
        {
            return direct;
        }

        if (root.TryGetProperty("d", out var nested))
        {
            return IdString(nested);
        }

        return string.Empty;
    }

    private static string IdString(JsonElement item)
    {
        foreach (var name in new[] { "Id", "id" })
        {
            if (!item.TryGetProperty(name, out var value))
            {
                continue;
            }

            if (value.ValueKind == JsonValueKind.String)
            {
                return value.GetString() ?? string.Empty;
            }

            if (value.ValueKind == JsonValueKind.Number)
            {
                return value.ToString();
            }
        }

        return string.Empty;
    }

    private static InvalidOperationException HttpError(HttpResponseMessage response, string body, string action)
    {
        var detail = (body ?? string.Empty).Trim();
        if (detail.Length > 300)
        {
            detail = detail[..300];
        }

        return new InvalidOperationException($"SharePoint {action} failed (HTTP {(int)response.StatusCode}). {detail}");
    }

    private sealed class ConfigStore : IConfigStore
    {
        private readonly SharePointClient _client;
        private readonly AppSettings _settings;
        private readonly int _itemId;

        public ConfigStore(SharePointClient client, AppSettings settings, int itemId)
        {
            _client = client;
            _settings = settings;
            _itemId = itemId;
        }

        public Task<NumberingConfig?> ReadAsync()
        {
            return _client.GetConfigItemByIdAsync(_settings, _itemId);
        }

        public Task<bool> CompareAndSwapAsync(string etag, long currentCount, string lastResetDate)
        {
            return _client.MergeCounterAsync(_settings, _itemId, etag, currentCount, lastResetDate);
        }
    }
}
