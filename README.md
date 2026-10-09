# Request numbers for SharePoint Online

When someone creates a list item, this repo assigns the next request number. Upload the SharePoint package, open the list, choose **Autogen Setting**, and save the formula. The list writes the number while it is open. No Azure Function is required for that.

Do not also run a Power Automate flow that writes the same number column. Two writers will race. An optional Azure Function can number items that are created while nobody has the list open. Those steps are in [deploy/README.md](deploy/README.md).

## Upload files

Built packages are in [deploy/](deploy/README.md).

| File | Use |
| --- | --- |
| [deploy/item-id-generator.sppkg](deploy/item-id-generator.sppkg) | The only file you upload to SharePoint. Add the app to each site collection. List owners set the formula. They do not paste a service address. |
| [deploy/item-id-generator-func-net.zip](deploy/item-id-generator-func-net.zip) | .NET 8 isolated webhook. Create a new Function App with stack **.NET 8 Isolated**, then Zip Push Deploy this file. Steps are in [deploy/README.md](deploy/README.md). |

Saving an active rule is enough. The list command numbers blank items while the list is open. The function package is optional.

## How a number is chosen

The list command reads `AutoGenFeatureConfiguration`, reserves the next count with the row etag, and writes the number column. Two people editing at the same time get different numbers because a lost etag is retried. An item created while the list is closed stays blank until the next person opens the list, unless the optional Azure Function is running.

## Flow

```text
SPFx configuration web part
  ensure AutoGenFeatureConfiguration on the site collection
  paste list URL -> GetList -> save row
  POST RegisterWebhook
        |
        v
SharePoint list subscription
  notificationUrl = spoWebhook
        |
        v
any create (form, grid, Power Automate)
        |
        v
spoWebhook
  echo validationtoken as text/plain, or check clientState and enqueue
        |
        v
queue request-numbers
        |
        v
processRequestNumber
  blob lease for that list
  ETag compare-and-swap on AutoGenFeatureConfiguration
  PATCH the item when the number column is still blank
```

| Path | Role |
| --- | --- |
| `schema/` | `AutoGenFeatureConfiguration` columns, a sample row, and a PnP script. |
| `functions/ItemIdGeneratorFunc` | `spoWebhook`, `processRequestNumber`, and `RegisterWebhook`. |
| `spfx/ItemIdGenerator` | Configuration web part. It provisions the list, resolves a list URL, and calls `RegisterWebhook`. |

## Concurrency

1. **Lease.** `processRequestNumber` takes a 60-second blob lease named `list-{guid}` in container `numbering-locks`. `host.json` sets queue `batchSize` to 1, which only limits one instance. A second instance that dequeues the same list fails the lease and the message retries after the 60-second visibility timeout (`maxDequeueCount` 5). Different lists run in parallel.
2. **ETag.** The winner reads `CurrentCount` and MERGEs `CurrentCount` plus `LastResetDate` with `If-Match`. HTTP 412 reads the row again. One hundred parallel reservations in the unit test produce sequences 1 through 100.
3. **Blank column.** The worker skips an item that already has a number. SharePoint notifications do not include the item id, so the worker reads the latest 200 items and numbers the blank ones, oldest id first. A queue message may also carry `itemId` for a manual replay.

A crash after the counter moves and before the item PATCH leaves a gap. It does not reuse the number.

Reset uses UTC. `None` never resets. An empty `LastResetDate` does not wipe `CurrentCount`. `Day`, `Month`, and `Year` reset when the UTC period key changes.

## Versions

| Piece | Version |
| --- | --- |
| Node.js | `>=22.14.0 <23` (`.nvmrc` is 22) |
| SPFx | 1.23.2, Heft, React 17. SPFx 1.24 is still a release candidate. |
| Azure Functions | programming model v4, `@azure/functions` 4.16.5 |
| Storage | `@azure/storage-blob` 12.34.0 |

This is not an npm workspace. SPFx Heft breaks when dependencies are hoisted.

## Local development

1. Create `AutoGenFeatureConfiguration` with the web part or `schema/provision-numbering-config.ps1`. Add a single-line text column such as `RequestNumber` on each target list.
2. Register an Entra application with SharePoint application permission **Sites.Manage.All** and admin-consent it once. That principal can number lists on every site collection. Copy `functions/ItemIdGeneratorFunc/local.settings.json.example` to `local.settings.json` and fill the placeholders locally. Do not commit that file. Do not set a site collection URL.
3. Start Azurite so the queue and the lease container have a storage account (`UseDevelopmentStorage=true`).
4. From `functions/ItemIdGeneratorFunc`: `npm install`, `npm test`, then `npm start`.
5. Put the `spoWebhook` URL, including `?code=<function-key>`, in `SPO_WEBHOOK_NOTIFICATION_URL`. Put the same style of URL for `RegisterWebhook` in the web part property pane. Allow the SharePoint origin in Function CORS. The example `local.settings.json` sets `Host.CORS` for local runs.
6. From `spfx/ItemIdGenerator`: `npm install`, `npm test`, `npm start`. The debug server is `https://localhost:4321`. Open the hosted workbench on the site collection. Hosted workbench retires on 2026-12-01. Package with `npm run package-solution`.

`gulp serve` in older docs is `npm start`. `gulp package-solution --ship` is `npm run package-solution`.

## App registration

`RegisterWebhook` and the queue worker use client credentials (`client_secret`) against `https://{sharepoint-host}/.default`. Certificate client assertion is not implemented. If the tenant rejects secrets, that is a gap.

The function key on `RegisterWebhook` is the gate. Anyone with the key can pass `configSiteUrl` and point the app at a site collection the app principal can read. The property pane shows that URL to page editors. Treat the key as a secret and prefer a short-lived key for a prototype.

## Gaps

- Webhook renewal is not implemented. SharePoint expires a subscription after at most 180 days. Registration asks for about 170 days.
- The worker scans the latest 200 items. It does not store a change token. A burst larger than that can leave older blank items unnumbered until another notification or a replay message with `itemId`.
- Deactivating a row does not delete the SharePoint subscription.
- Changing the target list clears `WebhookSubscriptionId` and registers a new subscription. The old subscription remains until it expires.
- A failed item PATCH after a successful increment burns a sequence.
- Certificate authentication is not implemented.
- The SPFx UI was not exercised in a SharePoint tenant. Unit tests cover the list contract, field XML, URL parsing, the counter, and webhook registration with a fake `fetch`.
