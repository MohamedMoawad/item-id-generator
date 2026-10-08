using System.Net;
using Xunit;
using ItemIdGenerator.Settings;
using ItemIdGenerator.SharePoint;
using ItemIdGenerator.Webhook;

namespace ItemIdGeneratorFuncNet.Tests;

public class SharePointClientTests
{
    private const string ListId = "aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee";

    [Fact]
    public async Task Register_rejects_a_missing_row_and_stores_a_new_subscription()
    {
        var handler = new RegisterWebhookHandler();
        var settings = SampleSettings();
        var missing = new SharePointClient(new HttpClient(new ScriptedHandler(request =>
            request.RequestUri!.ToString().Contains("oauth2", StringComparison.Ordinal)
                ? Response(HttpStatusCode.OK, """{"access_token":"token","expires_in":3600}""")
                : Response(HttpStatusCode.NotFound, ""))));
        var notFound = await handler.HandleAsync("""{"configItemId":7,"configSiteUrl":"https://contoso.sharepoint.com/sites/config"}""", settings, missing);
        Assert.Equal(404, notFound.Status);

        var script = new ScriptedHandler(request =>
        {
            var url = request.RequestUri!.ToString();
            if (url.Contains("oauth2", StringComparison.Ordinal))
            {
                return Response(HttpStatusCode.OK, """{"access_token":"token","expires_in":3600}""");
            }

            if (url.EndsWith("/subscriptions", StringComparison.Ordinal))
            {
                Assert.False(request.Headers.Contains("X-HTTP-Method"));
                return Response(HttpStatusCode.Created, """{"Id":"sub-from-handler"}""");
            }

            if (request.Headers.TryGetValues("X-HTTP-Method", out var method) && method.Contains("MERGE"))
            {
                return Response(HttpStatusCode.NoContent, "");
            }

            return Response(HttpStatusCode.OK, $$"""
            {"Id":7,"TargetListUrl":"https://contoso.sharepoint.com/sites/ops/Lists/Requests","TargetListGuid":"{{ListId}}","IsActive":true,"WebhookSubscriptionId":"","odata.etag":"\"1\""}
            """);
        });
        var client = new SharePointClient(new HttpClient(script));
        var created = await handler.HandleAsync("""{"configItemId":7,"configSiteUrl":"https://contoso.sharepoint.com/sites/ops"}""", settings, client);
        Assert.Equal(200, created.Status);
        Assert.Contains("sub-from-handler", created.Body, StringComparison.Ordinal);
        Assert.Contains("\"alreadyRegistered\":false", created.Body, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Enabled_config_is_read_from_the_site_collection_of_the_webhook()
    {
        var requested = new List<string>();
        var client = new SharePointClient(new HttpClient(new ScriptedHandler(request =>
        {
            var url = request.RequestUri!.ToString();
            requested.Add(url);
            if (url.Contains("oauth2", StringComparison.Ordinal))
            {
                return Response(HttpStatusCode.OK, """{"access_token":"token","expires_in":3600}""");
            }

            if (url.Contains("/_api/site?", StringComparison.Ordinal))
            {
                return Response(HttpStatusCode.OK, """{"Url":"https://contoso.sharepoint.com/sites/ops"}""");
            }

            if (url.Contains("https://contoso.sharepoint.com/sites/ops/_api/web/lists/getbytitle", StringComparison.Ordinal))
            {
                return Response(HttpStatusCode.OK, $$"""
                {"value":[{"Id":7,"TargetListUrl":"https://contoso.sharepoint.com/sites/ops/team/Lists/Requests","TargetListGuid":"{{ListId}}","IsActive":true,"Formula":"REQ-{seq}","NumberColumnInternalName":"RequestNumber","odata.etag":"\"1\""}]}
                """);
            }

            return Response(HttpStatusCode.NotFound, "");
        })));

        var config = await client.GetEnabledConfigAsync(
            SampleSettings() with { ConfigSiteUrl = string.Empty },
            "https://contoso.sharepoint.com/sites/ops/team",
            ListId);
        Assert.NotNull(config);
        Assert.Equal("https://contoso.sharepoint.com/sites/ops", config!.ConfigSiteUrl);
        Assert.Contains(requested, url => url.Contains("/sites/ops/team/_api/site?", StringComparison.Ordinal));
        Assert.Contains(requested, url => url.Contains("https://contoso.sharepoint.com/sites/ops/_api/web/lists/getbytitle", StringComparison.Ordinal));
        Assert.Equal("https://contoso.sharepoint.com/sites/hr", SharePointUrls.ReadSiteCollectionUrl("""{"d":{"Url":"https://contoso.sharepoint.com/sites/hr/"}}""", "https://contoso.sharepoint.com/sites/hr/team"));
    }

    [Fact]
    public void Register_call_supplies_the_spoWebhook_address_and_client_state()
    {
        var settings = SampleSettings() with { NotificationUrl = string.Empty, WebhookClientState = string.Empty };
        Assert.Equal(
            "https://numbers.azurewebsites.net/api/spoWebhook",
            settings.ResolveNotificationUrl(new Uri("https://numbers.azurewebsites.net/api/RegisterWebhook?code=abc")));
        Assert.Equal(string.Empty, settings.ResolveNotificationUrl(new Uri("http://localhost:7071/api/RegisterWebhook")));
        Assert.Equal(32, settings.ResolveClientState().Length);
        Assert.Equal(settings.ResolveClientState(), settings.ResolveClientState());
    }

    [Fact]
    public async Task Register_rejects_a_body_without_a_config_item_id()
    {
        var handler = new RegisterWebhookHandler();
        var result = await handler.HandleAsync("""{"configItemId":0}""", SampleSettings(), new SharePointClient(new HttpClient()));
        Assert.Equal(400, result.Status);

        var missingSite = await handler.HandleAsync("""{"configItemId":7}""", SampleSettings(), new SharePointClient(new HttpClient()));
        Assert.Equal(400, missingSite.Status);
        Assert.Contains("configSiteUrl", missingSite.Body, StringComparison.Ordinal);
    }

    private static AppSettings SampleSettings()
    {
        return new AppSettings(
            "UseDevelopmentStorage=true",
            "https://contoso.sharepoint.com/sites/config",
            "RequestNumberConfig",
            "11111111-2222-4333-8444-555555555555",
            "66666666-7777-4888-8999-aaaaaaaaaaaa",
            "not-a-real-secret",
            "not-a-real-client-state",
            "https://contoso.azurewebsites.net/api/spoWebhook?code=not-a-real-key");
    }

    private static HttpResponseMessage Response(HttpStatusCode status, string body)
    {
        return new HttpResponseMessage(status)
        {
            Content = new StringContent(body)
        };
    }

    private sealed class ScriptedHandler : HttpMessageHandler
    {
        private readonly Func<HttpRequestMessage, HttpResponseMessage> _handle;
        public ScriptedHandler(Func<HttpRequestMessage, HttpResponseMessage> handle) => _handle = handle;
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            return Task.FromResult(_handle(request));
        }
    }
}
