namespace ItemIdGenerator.Numbering;

public sealed class NumberingConfig
{
    public int Id { get; init; }
    public string Etag { get; init; } = string.Empty;
    public string Title { get; init; } = string.Empty;
    public string TargetListUrl { get; init; } = string.Empty;
    public string TargetSiteUrl { get; init; } = string.Empty;
    public string ConfigSiteUrl { get; init; } = string.Empty;
    public string TargetListGuid { get; init; } = string.Empty;
    public string NumberColumnInternalName { get; init; } = string.Empty;
    public long CurrentCount { get; init; }
    public string Formula { get; init; } = string.Empty;
    public string ResetPeriod { get; init; } = "None";
    public string LastResetDate { get; init; } = string.Empty;
    public bool IsActive { get; init; }
    public int PadLength { get; init; }
    public string WebhookSubscriptionId { get; init; } = string.Empty;

    public NumberingConfig WithTargetSite(string siteUrl)
    {
        return new NumberingConfig
        {
            Id = Id,
            Etag = Etag,
            Title = Title,
            TargetListUrl = TargetListUrl,
            TargetSiteUrl = siteUrl,
            ConfigSiteUrl = ConfigSiteUrl,
            TargetListGuid = TargetListGuid,
            NumberColumnInternalName = NumberColumnInternalName,
            CurrentCount = CurrentCount,
            Formula = Formula,
            ResetPeriod = ResetPeriod,
            LastResetDate = LastResetDate,
            IsActive = IsActive,
            PadLength = PadLength,
            WebhookSubscriptionId = WebhookSubscriptionId
        };
    }

    public NumberingConfig WithConfigSite(string siteUrl)
    {
        return new NumberingConfig
        {
            Id = Id,
            Etag = Etag,
            Title = Title,
            TargetListUrl = TargetListUrl,
            TargetSiteUrl = TargetSiteUrl,
            ConfigSiteUrl = siteUrl,
            TargetListGuid = TargetListGuid,
            NumberColumnInternalName = NumberColumnInternalName,
            CurrentCount = CurrentCount,
            Formula = Formula,
            ResetPeriod = ResetPeriod,
            LastResetDate = LastResetDate,
            IsActive = IsActive,
            PadLength = PadLength,
            WebhookSubscriptionId = WebhookSubscriptionId
        };
    }
}

public sealed class IssuedCode
{
    public string Code { get; init; } = string.Empty;
    public long Sequence { get; init; }
    public string LastResetDate { get; init; } = string.Empty;
}

public interface IConfigStore
{
    Task<NumberingConfig?> ReadAsync();
    Task<bool> CompareAndSwapAsync(string etag, long currentCount, string lastResetDate);
}
