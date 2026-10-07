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
        var created = await handler.HandleAsync("""{"configItemId":7}""", settings, client);
        Assert.Equal(200, created.Status);
        Assert.Contains("sub-from-handler", created.Body, StringComparison.Ordinal);
        Assert.Contains("\"alreadyRegistered\":false", created.Body, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Register_rejects_a_body_without_a_config_item_id()
    {
        var handler = new RegisterWebhookHandler();
        var result = await handler.HandleAsync("""{"configItemId":0}""", SampleSettings(), new SharePointClient(new HttpClient()));
        Assert.Equal(400, result.Status);
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
