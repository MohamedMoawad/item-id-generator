using ItemIdGenerator.Lease;
using Xunit;
using ItemIdGenerator.Numbering;
using ItemIdGenerator.Queue;
using ItemIdGenerator.SharePoint;
using ItemIdGenerator.Webhook;

namespace ItemIdGeneratorFuncNet.Tests;

public class NumberingTests
{
    private static readonly DateTimeOffset October = DateTimeOffset.Parse("2026-10-06T15:04:00Z");
    private const string ListId = "aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee";

    [Fact]
    public void Formula_pads_sequence_and_formats_utc_tokens()
    {
        Assert.Equal("REQ-202610-0012", Formula.Apply("REQ-{yyyy}{MM}-{seq:4}", 12, October, 0));
        Assert.Equal("REQ-0012", Formula.Apply("REQ-{seq}", 12, October, 4));
        Assert.Equal("REQ-12", Formula.Apply("REQ-{seq}", 12, October, 0));
        Assert.Equal("REQ-12", Formula.Apply("REQ-{seq:2}", 12, October, 6));
        Assert.Equal("2026-{department}-03", Formula.Apply("{yyyy}-{department}-{seq:2}", 3, October, 0));
    }

    [Fact]
    public void Formula_rejects_a_code_that_would_repeat()
    {
        var error = Assert.Throws<InvalidOperationException>(() => Formula.Apply("REQ-{yyyy}", 1, October, 0));
        Assert.Contains("{seq}", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Reset_uses_utc_period_keys()
    {
        Assert.Equal("2026-10-06", Formula.PeriodKeyFor("Day", October));
        Assert.Equal("2026-10", Formula.PeriodKeyFor("month", October));
        Assert.Equal("2026", Formula.PeriodKeyFor("YEAR", October));
        Assert.Equal(string.Empty, Formula.PeriodKeyFor("None", October));
        Assert.False(Formula.PeriodRolledOver("None", "2020-01-01T00:00:00.000Z", October));
        Assert.False(Formula.PeriodRolledOver("Month", "", October));
        Assert.False(Formula.PeriodRolledOver("Month", "not-a-date", October));
        Assert.True(Formula.PeriodRolledOver("Month", "2026-09-30T23:00:00.000Z", October));
        Assert.False(Formula.PeriodRolledOver("Month", "2026-10-01T00:00:00.000Z", October));
    }

    [Fact]
    public async Task Reserve_retries_a_lost_etag_and_refuses_an_inactive_row()
    {
        var store = new MemoryStore(Active(etag: "\"1\"", count: 4));
        store.FailNextSwap = true;
        var issued = await IssueNextCode.ReserveAsync(store, October, sleep: _ => Task.CompletedTask);
        Assert.Equal("REQ-0005", issued.Code);
        Assert.Equal(5, issued.Sequence);

        var inactive = new MemoryStore(Active(etag: "\"1\"", count: 1, isActive: false));
        await Assert.ThrowsAsync<InvalidOperationException>(() => IssueNextCode.ReserveAsync(inactive, October, sleep: _ => Task.CompletedTask));
    }

    [Fact]
    public async Task Parallel_reservations_get_distinct_sequences()
    {
        var store = new MemoryStore(Active(etag: "\"0\"", count: 0));
        var tasks = Enumerable.Range(0, 100)
            .Select(_ => IssueNextCode.ReserveAsync(store, October, sleep: _ => Task.CompletedTask))
            .ToArray();
        var issued = await Task.WhenAll(tasks);
        var sequences = issued.Select(item => item.Sequence).OrderBy(value => value).ToArray();
        Assert.Equal(Enumerable.Range(1, 100).Select(value => (long)value), sequences);
    }

    [Fact]
    public void List_urls_drop_the_form_page()
    {
        var list = SharePointUrls.DeriveWebFromListUrl("https://contoso.sharepoint.com/sites/ops/Lists/Requests/AllItems.aspx");
        Assert.Equal("https://contoso.sharepoint.com/sites/ops", list.WebAbsoluteUrl);
        Assert.Equal("/sites/ops/Lists/Requests", list.ServerRelativeUrl);

        var library = SharePointUrls.DeriveWebFromListUrl("https://contoso.sharepoint.com/sites/ops/Shared Documents/Forms/AllItems.aspx");
        Assert.Equal("https://contoso.sharepoint.com/sites/ops", library.WebAbsoluteUrl);
        Assert.Equal("/sites/ops/Shared Documents", library.ServerRelativeUrl);
        Assert.Equal(ListId, SharePointUrls.ListIdFromResource($"https://contoso.sharepoint.com/sites/ops/_api/web/lists('{ListId.ToUpperInvariant()}')"));
        Assert.Equal("list-" + ListId, BlobLeaseStore.BlobNameForList(ListId));
    }

    [Fact]
    public void Webhook_echoes_the_validation_token_and_enqueues_a_matching_notification()
    {
        var handler = new SharePointWebhookHandler();
        var handshake = handler.Handle("abc 123", null, null);
        Assert.Equal(200, handshake.Status);
        Assert.True(handshake.PlainText);
        Assert.Equal("abc 123", handshake.Body);

        var missing = handler.Handle(null, """{"value":[]}""", "<webhook-client-state>");
        Assert.Equal(500, missing.Status);

        var rejected = handler.Handle(null, Notification("wrong"), "expected-state");
        Assert.Equal(403, rejected.Status);
        Assert.Empty(rejected.QueueMessages);

        var accepted = handler.Handle(null, Notification("expected-state"), "expected-state");
        Assert.Equal(200, accepted.Status);
        Assert.Contains("\"accepted\":1", accepted.Body, StringComparison.Ordinal);
        Assert.Contains(ListId, accepted.QueueMessages[0], StringComparison.Ordinal);
        Assert.Contains("https://contoso.sharepoint.com/sites/ops", accepted.QueueMessages[0], StringComparison.Ordinal);
    }

    [Fact]
    public async Task Processor_numbers_blank_items_and_retries_when_the_lease_is_held()
    {
        var gateway = new FakeGateway();
        var processor = new NumberingProcessor();
        var issued = await processor.ProcessAsync(
            $$"""{"siteUrl":"https://contoso.sharepoint.com/sites/ops","listId":"{{ListId}}"}""",
            gateway,
            new FakeLease(true),
            () => October,
            null);
        Assert.Equal(new[] { 3, 8 }, issued.Select(item => item.ItemId).ToArray());
        Assert.Equal(new long[] { 1, 2 }, issued.Select(item => item.Sequence).ToArray());
        Assert.Equal("https://contoso.sharepoint.com/sites/ops", gateway.RequestedSite);

        await Assert.ThrowsAsync<InvalidOperationException>(() => processor.ProcessAsync(
            $$"""{"listId":"{{ListId}}"}""",
            gateway,
            new FakeLease(false),
            () => October,
            null));

        var idle = new FakeGateway { Config = null };
        var none = await processor.ProcessAsync(
            $$"""{"siteUrl":"https://contoso.sharepoint.com/sites/hr","listId":"{{ListId}}"}""",
            idle,
            new FakeLease(true),
            () => October,
            null);
        Assert.Empty(none);
    }

    private static string Notification(string clientState)
    {
        return $$"""
        {"value":[{"clientState":"{{clientState}}","resource":"{AAAAAAAA-BBBB-4CCC-8DDD-EEEEEEEEEEEE}","siteUrl":"https://contoso.sharepoint.com/sites/ops","subscriptionId":"sub-1"}]}
        """;
    }

    private static NumberingConfig Active(string etag, long count, bool isActive = true)
    {
        return new NumberingConfig
        {
            Id = 7,
            Etag = etag,
            Formula = "REQ-{seq:4}",
            ResetPeriod = "None",
            CurrentCount = count,
            IsActive = isActive,
            PadLength = 4,
            TargetListGuid = ListId,
            TargetSiteUrl = "https://contoso.sharepoint.com/sites/ops"
        };
    }

    private sealed class MemoryStore : IConfigStore
    {
        private readonly object _gate = new();
        private NumberingConfig _current;
        public bool FailNextSwap { get; set; }

        public MemoryStore(NumberingConfig current)
        {
            _current = current;
        }

        public Task<NumberingConfig?> ReadAsync()
        {
            lock (_gate)
            {
                return Task.FromResult<NumberingConfig?>(_current);
            }
        }

        public Task<bool> CompareAndSwapAsync(string etag, long currentCount, string lastResetDate)
        {
            lock (_gate)
            {
                if (FailNextSwap)
                {
                    FailNextSwap = false;
                    _current = Copy(_current, "\"lost\"", _current.CurrentCount);
                    return Task.FromResult(false);
                }

                if (_current.Etag != etag)
                {
                    return Task.FromResult(false);
                }

                _current = Copy(_current, "\"" + currentCount + "\"", currentCount);
                return Task.FromResult(true);
            }
        }

        private static NumberingConfig Copy(NumberingConfig current, string etag, long count)
        {
            return new NumberingConfig
            {
                Id = current.Id,
                Etag = etag,
                Formula = current.Formula,
                ResetPeriod = current.ResetPeriod,
                CurrentCount = count,
                IsActive = current.IsActive,
                PadLength = current.PadLength,
                LastResetDate = current.LastResetDate,
                TargetListGuid = current.TargetListGuid,
                TargetSiteUrl = current.TargetSiteUrl
            };
        }
    }

    private sealed class FakeLease : ILeaseStore
    {
        private readonly bool _acquired;
        public FakeLease(bool acquired) => _acquired = acquired;
        public Task<LeaseHold> AcquireAsync(string listId) => Task.FromResult(new LeaseHold(_acquired, () => Task.CompletedTask));
    }

    private sealed class FakeGateway : INumberingGateway
    {
        public string RequestedSite { get; private set; } = string.Empty;
        public NumberingConfig? Config { get; set; } = new NumberingConfig
        {
            Id = 7,
            IsActive = true,
            Formula = "REQ-{seq:4}",
            ResetPeriod = "None",
            CurrentCount = 0,
            Etag = "\"0\"",
            PadLength = 4,
            TargetListGuid = ListId,
            TargetSiteUrl = "https://contoso.sharepoint.com/sites/ops",
            NumberColumnInternalName = "RequestNumber"
        };

        private readonly MemoryStore _store;

        public FakeGateway()
        {
            _store = new MemoryStore(Config!);
        }

        public Task<NumberingConfig?> GetEnabledConfigAsync(string siteUrl, string listId)
        {
            RequestedSite = siteUrl;
            return Task.FromResult(Config);
        }

        public Task<IReadOnlyList<int>> ListUnnumberedAsync(NumberingConfig config, int? itemId)
        {
            return Task.FromResult<IReadOnlyList<int>>(new[] { 8, 3 });
        }

        public Task<IssuedCode> ReserveNextCodeAsync(NumberingConfig config, DateTimeOffset now)
        {
            return IssueNextCode.ReserveAsync(_store, now, sleep: _ => Task.CompletedTask);
        }

        public Task<bool> WriteNumberIfBlankAsync(NumberingConfig config, int itemId, string code)
        {
            return Task.FromResult(true);
        }
    }
}
