using System.Globalization;
using System.Text.Json;
using ItemIdGenerator.Lease;
using ItemIdGenerator.Numbering;
using ItemIdGenerator.Settings;
using ItemIdGenerator.SharePoint;
using Microsoft.Extensions.Logging;

namespace ItemIdGenerator.Queue;

public interface INumberingGateway
{
    Task<NumberingConfig?> GetEnabledConfigAsync(string listId);
    Task<IReadOnlyList<int>> ListUnnumberedAsync(NumberingConfig config, int? itemId);
    Task<IssuedCode> ReserveNextCodeAsync(NumberingConfig config, DateTimeOffset now);
    Task<bool> WriteNumberIfBlankAsync(NumberingConfig config, int itemId, string code);
}

public sealed class SharePointNumberingGateway : INumberingGateway
{
    private readonly SharePointClient _sharePoint;
    private readonly AppSettings _settings;

    public SharePointNumberingGateway(SharePointClient sharePoint, AppSettings settings)
    {
        _sharePoint = sharePoint;
        _settings = settings;
    }

    public Task<NumberingConfig?> GetEnabledConfigAsync(string listId)
    {
        return _sharePoint.GetEnabledConfigAsync(_settings, listId);
    }

    public Task<IReadOnlyList<int>> ListUnnumberedAsync(NumberingConfig config, int? itemId)
    {
        return _sharePoint.ListUnnumberedItemsAsync(_settings, config, itemId);
    }

    public Task<IssuedCode> ReserveNextCodeAsync(NumberingConfig config, DateTimeOffset now)
    {
        return IssueNextCode.ReserveAsync(_sharePoint.CreateConfigStore(_settings, config), now);
    }

    public Task<bool> WriteNumberIfBlankAsync(NumberingConfig config, int itemId, string code)
    {
        return _sharePoint.WriteNumberIfBlankAsync(_settings, config, itemId, code);
    }
}

public sealed class NumberingProcessor
{
    public async Task<IReadOnlyList<IssuedItem>> ProcessAsync(
        string message,
        INumberingGateway gateway,
        ILeaseStore leaseStore,
        Func<DateTimeOffset> now,
        ILogger? logger)
    {
        var work = ParseWorkMessage(message);
        var lease = await leaseStore.AcquireAsync(work.ListId);
        if (!lease.Acquired)
        {
            throw new InvalidOperationException($"Numbering lease for list {work.ListId} is held. The queue message will retry.");
        }

        try
        {
            var config = await gateway.GetEnabledConfigAsync(work.ListId);
            if (config is null)
            {
                logger?.LogInformation("No active numbering config for list {ListId}.", work.ListId);
                return Array.Empty<IssuedItem>();
            }

            var siteUrl = string.IsNullOrWhiteSpace(config.TargetSiteUrl) ? work.SiteUrl : config.TargetSiteUrl;
            if (string.IsNullOrWhiteSpace(siteUrl))
            {
                throw new InvalidOperationException($"Numbering config {config.Id} has no target site URL.");
            }

            var ready = config.WithTargetSite(siteUrl);
            var items = (await gateway.ListUnnumberedAsync(ready, work.ItemId)).OrderBy(id => id).ToArray();
            var issued = new List<IssuedItem>();
            foreach (var itemId in items)
            {
                var next = await gateway.ReserveNextCodeAsync(ready, now());
                var written = await gateway.WriteNumberIfBlankAsync(ready, itemId, next.Code);
                if (written)
                {
                    issued.Add(new IssuedItem { ItemId = itemId, Code = next.Code, Sequence = next.Sequence });
                    logger?.LogInformation("Issued {Code} to item {ItemId} on list {ListId}.", next.Code, itemId, work.ListId);
                }
            }

            return issued;
        }
        finally
        {
            await lease.Release();
        }
    }

    public static NumberingWork ParseWorkMessage(string message)
    {
        JsonElement body;
        try
        {
            using var document = JsonDocument.Parse(message);
            body = document.RootElement.Clone();
        }
        catch (JsonException)
        {
            throw new InvalidOperationException("Queue message must be a JSON object.");
        }

        if (body.ValueKind != JsonValueKind.Object)
        {
            throw new InvalidOperationException("Queue message must be a JSON object.");
        }

        var listId = SharePointUrls.NormalizeGuid(StringValue(body, "listId"));
        if (listId.Length == 0)
        {
            throw new InvalidOperationException("Queue message listId must be a GUID.");
        }

        int? itemId = null;
        if (body.TryGetProperty("itemId", out var item) && item.ValueKind is not JsonValueKind.Null and not JsonValueKind.Undefined)
        {
            if (item.ValueKind == JsonValueKind.String && string.IsNullOrWhiteSpace(item.GetString()))
            {
                itemId = null;
            }
            else if (!TryItemId(item, out var parsed))
            {
                throw new InvalidOperationException("Queue message itemId must be a positive integer when it is present.");
            }
            else
            {
                itemId = parsed;
            }
        }

        return new NumberingWork
        {
            SiteUrl = StringValue(body, "siteUrl").Trim(),
            ListId = listId,
            ItemId = itemId
        };
    }

    private static bool TryItemId(JsonElement value, out int itemId)
    {
        itemId = 0;
        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var number) && number > 0)
        {
            itemId = number;
            return true;
        }

        if (value.ValueKind == JsonValueKind.String)
        {
            var text = (value.GetString() ?? string.Empty).Trim();
            if (text.Length > 0 && text[0] != '0' && int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out var parsed) && parsed > 0)
            {
                itemId = parsed;
                return true;
            }
        }

        return false;
    }

    private static string StringValue(JsonElement item, string name)
    {
        if (!item.TryGetProperty(name, out var value) || value.ValueKind != JsonValueKind.String)
        {
            return string.Empty;
        }

        return value.GetString() ?? string.Empty;
    }
}

public sealed class NumberingWork
{
    public string SiteUrl { get; init; } = string.Empty;
    public string ListId { get; init; } = string.Empty;
    public int? ItemId { get; init; }
}

public sealed class IssuedItem
{
    public int ItemId { get; init; }
    public string Code { get; init; } = string.Empty;
    public long Sequence { get; init; }
}
