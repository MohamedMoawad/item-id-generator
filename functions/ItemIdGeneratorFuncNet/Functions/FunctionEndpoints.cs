using System.Net;
using ItemIdGenerator.Lease;
using ItemIdGenerator.Queue;
using ItemIdGenerator.Settings;
using ItemIdGenerator.SharePoint;
using ItemIdGenerator.Webhook;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;

namespace ItemIdGenerator.Functions;

public sealed class SpoWebhookFunction
{
    private readonly SharePointWebhookHandler _handler;

    public SpoWebhookFunction(SharePointWebhookHandler handler)
    {
        _handler = handler;
    }

    [Function("spoWebhook")]
    public async Task<SpoWebhookResponse> Run(
        [HttpTrigger(AuthorizationLevel.Anonymous, "post")] HttpRequestData request,
        FunctionContext context)
    {
        var token = QueryValue(request.Url, "validationtoken");
        var body = string.IsNullOrEmpty(token) ? await request.ReadAsStringAsync() : null;
        var result = _handler.Handle(token, body, AppSettings.FromEnvironment().ResolveClientState());
        if (result.Status == 200 && result.QueueMessages.Count > 0)
        {
            context.GetLogger("spoWebhook").LogInformation("Accepted {Count} SharePoint webhook notification(s).", result.QueueMessages.Count);
        }

        return new SpoWebhookResponse
        {
            Messages = result.QueueMessages.Count == 0 ? null : result.QueueMessages.ToArray(),
            HttpResponse = await WriteAsync(request, result)
        };
    }

    internal static string? QueryValue(Uri url, string name)
    {
        var query = url.Query;
        if (query.StartsWith('?'))
        {
            query = query[1..];
        }

        foreach (var part in query.Split('&', StringSplitOptions.RemoveEmptyEntries))
        {
            var split = part.IndexOf('=');
            var key = Uri.UnescapeDataString(split < 0 ? part : part[..split]);
            if (!key.Equals(name, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            return split < 0 ? string.Empty : Uri.UnescapeDataString(part[(split + 1)..]);
        }

        return null;
    }

    private static async Task<HttpResponseData> WriteAsync(HttpRequestData request, WebhookResult result)
    {
        var response = request.CreateResponse((HttpStatusCode)result.Status);
        response.Headers.Add("Content-Type", result.PlainText ? "text/plain; charset=utf-8" : "application/json; charset=utf-8");
        await response.WriteStringAsync(result.Body);
        return response;
    }
}

public sealed class SpoWebhookResponse
{
    [QueueOutput("request-numbers", Connection = "AzureWebJobsStorage")]
    public string[]? Messages { get; set; }

    public HttpResponseData? HttpResponse { get; set; }
}

public sealed class RegisterWebhookFunction
{
    private readonly RegisterWebhookHandler _handler;
    private readonly SharePointClient _sharePoint;

    public RegisterWebhookFunction(RegisterWebhookHandler handler, SharePointClient sharePoint)
    {
        _handler = handler;
        _sharePoint = sharePoint;
    }

    [Function("RegisterWebhook")]
    public async Task<HttpResponseData> Run(
        [HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequestData request,
        FunctionContext context)
    {
        var body = await request.ReadAsStringAsync();
        var result = await _handler.HandleAsync(body, AppSettings.FromEnvironment(), _sharePoint, request.Url);
        if (result.Status == 200)
        {
            context.GetLogger("RegisterWebhook").LogInformation("Registered a SharePoint list webhook.");
        }

        var response = request.CreateResponse((HttpStatusCode)result.Status);
        response.Headers.Add("Content-Type", "application/json; charset=utf-8");
        await response.WriteStringAsync(result.Body);
        return response;
    }
}

public sealed class ProcessRequestNumberFunction
{
    private readonly SharePointClient _sharePoint;
    private readonly NumberingProcessor _processor;

    public ProcessRequestNumberFunction(SharePointClient sharePoint, NumberingProcessor processor)
    {
        _sharePoint = sharePoint;
        _processor = processor;
    }

    [Function("processRequestNumber")]
    public async Task Run(
        [QueueTrigger("request-numbers", Connection = "AzureWebJobsStorage")] string message,
        FunctionContext context)
    {
        var settings = AppSettings.FromEnvironment();
        var gateway = new SharePointNumberingGateway(_sharePoint, settings);
        var leases = new BlobLeaseStore(settings.StorageConnectionString);
        await _processor.ProcessAsync(message, gateway, leases, () => DateTimeOffset.UtcNow, context.GetLogger("processRequestNumber"));
    }
}
