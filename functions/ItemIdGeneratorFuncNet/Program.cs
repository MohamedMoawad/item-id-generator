using ItemIdGenerator.Queue;
using ItemIdGenerator.SharePoint;
using ItemIdGenerator.Webhook;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

var host = new HostBuilder()
    .ConfigureFunctionsWorkerDefaults()
    .ConfigureServices(services =>
    {
        services.AddHttpClient<SharePointClient>(client =>
        {
            client.Timeout = TimeSpan.FromSeconds(100);
        });
        services.AddSingleton<SharePointWebhookHandler>();
        services.AddSingleton<RegisterWebhookHandler>();
        services.AddSingleton<NumberingProcessor>();
    })
    .Build();

host.Run();
