# Request numbers, .NET 8 isolated

`spoWebhook`, `processRequestNumber`, and `RegisterWebhook`. Same behavior as the Node worker: SharePoint calls `spoWebhook`, the queue worker numbers blank items, and the SharePoint panel calls `RegisterWebhook`.

The upload file is [deploy/item-id-generator-func-net.zip](../../deploy/item-id-generator-func-net.zip). It is a `dotnet publish` output, so `functions.metadata` is already in the zip. Create the Function App as **.NET 8 Isolated**. Do not upload this zip to a Node Function App.

```bash
dotnet test
dotnet publish -c Release -o publish
```

Copy `local.settings.json.example` to `local.settings.json` for local runs. Leave the secret empty in source control.
