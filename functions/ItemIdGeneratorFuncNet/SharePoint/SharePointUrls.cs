using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using ItemIdGenerator.Numbering;

namespace ItemIdGenerator.SharePoint;

public static class SharePointUrls
{
    public const string ConfigListTitle = "AutoGenFeatureConfiguration";

    private static readonly Regex GuidPattern = new(@"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex GuidSearch = new(@"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex InternalName = new(@"^[A-Za-z_][A-Za-z0-9_]*$", RegexOptions.Compiled);
    private static readonly Regex ListsSegment = new(@"^(.*)/lists/[^/]+$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static string NormalizeGuid(string? value)
    {
        var text = (value ?? string.Empty).Trim().Trim('{', '}').ToLowerInvariant();
        return GuidPattern.IsMatch(text) ? text : string.Empty;
    }

    public static string ListIdFromResource(string? resource)
    {
        var direct = NormalizeGuid(resource);
        if (direct.Length > 0)
        {
            return direct;
        }

        var match = GuidSearch.Match(resource ?? string.Empty);
        return match.Success ? NormalizeGuid(match.Value) : string.Empty;
    }

    public static void AssertInternalName(string? name)
    {
        if (name is null || !InternalName.IsMatch(name))
        {
            throw new InvalidOperationException($"Invalid column internal name \"{name}\".");
        }
    }

    public static Uri AssertHttpsUrl(string? value, string label)
    {
        if (!Uri.TryCreate(value, UriKind.Absolute, out var url))
        {
            throw new InvalidOperationException($"{label} is not a URL.");
        }

        if (!string.Equals(url.Scheme, "https", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException($"{label} must use https.");
        }

        return url;
    }

    public static string SiteRoot(string siteUrl)
    {
        return AssertHttpsUrl(siteUrl, "SharePoint site URL").GetLeftPart(UriPartial.Path).TrimEnd('/');
    }

    public static string ReadSiteCollectionUrl(string? json, string fallbackWebUrl)
    {
        var fallback = SiteRoot(fallbackWebUrl);
        if (string.IsNullOrWhiteSpace(json))
        {
            return fallback;
        }

        try
        {
            using var document = JsonDocument.Parse(json);
            var url = SiteUrlValue(document.RootElement);
            return url.Length == 0 ? fallback : SiteRoot(url);
        }
        catch (JsonException)
        {
            return fallback;
        }
    }

    private static string SiteUrlValue(JsonElement item)
    {
        var direct = FirstString(item, "Url", "url");
        if (direct.Length > 0)
        {
            return direct;
        }

        if (item.ValueKind == JsonValueKind.Object && item.TryGetProperty("d", out var nested))
        {
            return FirstString(nested, "Url", "url");
        }

        return string.Empty;
    }

    public static string EscapeODataString(string value)
    {
        return value.Replace("'", "''", StringComparison.Ordinal);
    }

    public static string BuildConfigItemsUrl(string siteUrl, string listTitle)
    {
        var select = "Id,Title,TargetListUrl,TargetListGuid,NumberColumnInternalName,Formula,CurrentCount,ResetPeriod,LastResetDate,IsActive,PadLength,WebhookSubscriptionId";
        return $"{SiteRoot(siteUrl)}/_api/web/lists/getbytitle('{EscapeODataString(listTitle)}')/items?$select={select}&$top=500";
    }

    public static string BuildConfigItemUrl(string siteUrl, string listTitle, int itemId)
    {
        return $"{SiteRoot(siteUrl)}/_api/web/lists/getbytitle('{EscapeODataString(listTitle)}')/items({itemId.ToString(CultureInfo.InvariantCulture)})";
    }

    public static string BuildTargetItemUrl(string siteUrl, string listId, int itemId)
    {
        var guid = NormalizeGuid(listId);
        if (guid.Length == 0)
        {
            throw new InvalidOperationException("Target list id must be a GUID.");
        }

        return $"{SiteRoot(siteUrl)}/_api/web/lists(guid'{guid}')/items({itemId.ToString(CultureInfo.InvariantCulture)})";
    }

    public static string BuildRecentItemsUrl(string siteUrl, string listId, string fieldInternalName)
    {
        AssertInternalName(fieldInternalName);
        var guid = NormalizeGuid(listId);
        if (guid.Length == 0)
        {
            throw new InvalidOperationException("Target list id must be a GUID.");
        }

        return $"{SiteRoot(siteUrl)}/_api/web/lists(guid'{guid}')/items?$select=Id,{fieldInternalName}&$orderby=Id desc&$top=200";
    }

    public static bool IsBlankNumber(JsonElement item, string fieldName)
    {
        if (!item.TryGetProperty(fieldName, out var value) || value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
        {
            return true;
        }

        return string.IsNullOrWhiteSpace(value.ToString());
    }

    public static bool IsActiveValue(JsonElement item)
    {
        if (!item.TryGetProperty("IsActive", out var value))
        {
            return false;
        }

        return value.ValueKind switch
        {
            JsonValueKind.True => true,
            JsonValueKind.Number => value.TryGetInt32(out var number) && number == 1,
            JsonValueKind.String => value.GetString() is "1" or "true" or "Yes",
            _ => false
        };
    }

    public static (string WebAbsoluteUrl, string ServerRelativeUrl) DeriveWebFromListUrl(string listUrl)
    {
        var url = AssertHttpsUrl(listUrl, "Target list URL");
        var path = Uri.UnescapeDataString(url.AbsolutePath);
        path = Regex.Replace(path, @"/Forms/AllItems\.aspx$", string.Empty, RegexOptions.IgnoreCase);
        path = Regex.Replace(path, @"/AllItems\.aspx$", string.Empty, RegexOptions.IgnoreCase);
        path = path.TrimEnd('/');
        var listsMatch = ListsSegment.Match(path);
        var webPath = listsMatch.Success ? listsMatch.Groups[1].Value : path[..path.LastIndexOf('/')];
        return ($"{url.GetLeftPart(UriPartial.Authority)}{webPath}", path);
    }

    public static NumberingConfig? NormalizeConfigItem(JsonElement raw)
    {
        if (raw.ValueKind != JsonValueKind.Object || !TryInt(raw, "Id", out var id))
        {
            return null;
        }

        var targetListUrl = StringValue(raw, "TargetListUrl").Trim();
        var targetSiteUrl = string.Empty;
        if (targetListUrl.Length > 0)
        {
            try
            {
                targetSiteUrl = DeriveWebFromListUrl(targetListUrl).WebAbsoluteUrl;
            }
            catch (InvalidOperationException)
            {
                targetSiteUrl = string.Empty;
            }
        }

        var padLength = 0;
        if (raw.TryGetProperty("PadLength", out var pad) && pad.ValueKind == JsonValueKind.Number && pad.TryGetInt32(out var parsedPad) && parsedPad > 0)
        {
            padLength = parsedPad;
        }

        long currentCount = 0;
        if (raw.TryGetProperty("CurrentCount", out var count) && count.ValueKind == JsonValueKind.Number && count.TryGetInt64(out var parsedCount))
        {
            currentCount = parsedCount;
        }

        return new NumberingConfig
        {
            Id = id,
            Etag = ReadEtag(raw),
            Title = StringValue(raw, "Title"),
            TargetListUrl = targetListUrl,
            TargetSiteUrl = targetSiteUrl,
            TargetListGuid = NormalizeGuid(StringValue(raw, "TargetListGuid")),
            NumberColumnInternalName = StringValue(raw, "NumberColumnInternalName"),
            CurrentCount = currentCount,
            Formula = StringValue(raw, "Formula"),
            ResetPeriod = string.IsNullOrEmpty(StringValue(raw, "ResetPeriod")) ? "None" : StringValue(raw, "ResetPeriod"),
            LastResetDate = StringValue(raw, "LastResetDate"),
            IsActive = IsActiveValue(raw),
            PadLength = padLength,
            WebhookSubscriptionId = StringValue(raw, "WebhookSubscriptionId")
        };
    }

    public static string ReadEtag(JsonElement item)
    {
        if (item.ValueKind != JsonValueKind.Object)
        {
            return string.Empty;
        }

        return FirstString(item, "odata.etag", "@odata.etag");
    }

    public static IEnumerable<JsonElement> CollectionItems(JsonElement payload)
    {
        if (payload.ValueKind == JsonValueKind.Array)
        {
            return payload.EnumerateArray();
        }

        if (payload.ValueKind == JsonValueKind.Object && payload.TryGetProperty("value", out var value) && value.ValueKind == JsonValueKind.Array)
        {
            return value.EnumerateArray();
        }

        if (payload.ValueKind == JsonValueKind.Object
            && payload.TryGetProperty("d", out var nested)
            && nested.ValueKind == JsonValueKind.Object
            && nested.TryGetProperty("results", out var results)
            && results.ValueKind == JsonValueKind.Array)
        {
            return results.EnumerateArray();
        }

        return Array.Empty<JsonElement>();
    }

    private static bool TryInt(JsonElement item, string name, out int value)
    {
        value = 0;
        return item.TryGetProperty(name, out var property) && property.ValueKind == JsonValueKind.Number && property.TryGetInt32(out value);
    }

    private static string StringValue(JsonElement item, string name)
    {
        if (!item.TryGetProperty(name, out var property) || property.ValueKind != JsonValueKind.String)
        {
            return string.Empty;
        }

        return property.GetString() ?? string.Empty;
    }

    private static string FirstString(JsonElement item, params string[] names)
    {
        foreach (var name in names)
        {
            var value = StringValue(item, name);
            if (value.Length > 0)
            {
                return value;
            }
        }

        return string.Empty;
    }
}
