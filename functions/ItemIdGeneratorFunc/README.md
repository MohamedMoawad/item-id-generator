# ItemIdGeneratorFunc

Azure Functions v4 (Node.js) worker for request numbers.

| Function | Trigger | Job |
| --- | --- | --- |
| `spoWebhook` | HTTP POST, auth level `function` | Echo `validationtoken` as `text/plain`, or check `clientState` and enqueue the list. |
| `processRequestNumber` | Queue `request-numbers` | Lease the list, increment `RequestNumberConfig`, PATCH a blank number column. |
| `RegisterWebhook` | HTTP POST, auth level `function` | App-only. Subscribe the target list to `spoWebhook` and store the subscription id. |

`src/index.js` loads those three modules. Do not add a second numbering flow.

## Settings

Copy `local.settings.json.example` to `local.settings.json`. Placeholders only. Leave the secret empty in source control.

| Name | Purpose |
| --- | --- |
| `AzureWebJobsStorage` | Queue and blob leases. Local value `UseDevelopmentStorage=true` needs Azurite. |
| `configSiteUrl` | Sent by the SharePoint panel. The worker reads `RequestNumberConfig` on that site collection. There is no Function App setting for a site URL. |
| `NUMBERING_CONFIG_LIST_TITLE` | Default `RequestNumberConfig`. |
| `SHAREPOINT_TENANT_ID` | Directory (tenant) id. |
| `SHAREPOINT_CLIENT_ID` | Application (client) id. |
| `SHAREPOINT_CLIENT_SECRET` | Client secret. Not committed. Certificate auth is not implemented. |
| `SHAREPOINT_WEBHOOK_CLIENT_STATE` | Shared secret SharePoint echoes on each notification. Rejected when empty or still a `<placeholder>`. |
| `SPO_WEBHOOK_NOTIFICATION_URL` | Full `spoWebhook` URL including `?code=<function-key>`. SharePoint calls this during subscribe and on each change. |
| `APPLICATIONINSIGHTS_CONNECTION_STRING` | Optional. |

`Host.CORS` in the example allows the SharePoint origin so the browser can call `RegisterWebhook`. `spoWebhook` is called by SharePoint, not by the browser, so it does not need CORS. Set the same origin on the Function App in Azure (`az functionapp cors add`).

The app needs SharePoint application permission **Sites.Manage.All** and admin consent once. That covers every site collection. Do not set a site collection URL on the Function App.

## spoWebhook

SharePoint proves the notification URL with:

```text
POST /api/spoWebhook?code=<function-key>&validationtoken=<token>
```

The response is the raw token, `Content-Type: text/plain`, not JSON.

A change notification has `value[]` with `resource`, `siteUrl`, `clientState`, and `subscriptionId`. It does not include the new item id. A mismatched `clientState` is HTTP 403. An unusable server secret is HTTP 500 so SharePoint retries instead of accepting a forged call. The queue message is `{ siteUrl, listId, subscriptionId }`.

## processRequestNumber

1. Acquire blob lease `list-{guid}` in `numbering-locks` for 60 seconds. A 409 throws so the message retries.
2. Resolve the site collection from the webhook `siteUrl`, then load the one active `RequestNumberConfig` row whose `TargetListGuid` matches. Zero rows is an ack. Two rows is an error.
3. Read the latest 200 items. Number those whose number column is blank, lowest id first. A message that includes `itemId` numbers only that item.
4. MERGE `CurrentCount` and `LastResetDate` with `If-Match`. HTTP 412 retries, up to 200 attempts.
5. MERGE the number onto the item with its own ETag. If the column is already filled, skip it. A crash between steps 4 and 5 leaves a gap.

`host.json` sets `batchSize` 1, `newBatchThreshold` 0, `visibilityTimeout` 1 minute, and `maxDequeueCount` 5. `batchSize` does not cover a second Function instance. The lease does.

Formula tokens are UTC: `{yyyy}` `{yy}` `{MM}` `{dd}` `{HH}` `{mm}` `{seq}` `{seq:n}`. `{seq}` uses `PadLength`. An empty `LastResetDate` does not reset the count.

## RegisterWebhook

```json
{ "configSiteUrl": "https://<tenant>.sharepoint.com/sites/<site>", "configItemId": 7 }
```

The function reads that row with the app principal, POSTs to `{web}/_api/web/lists('{guid}')/subscriptions`, and MERGEs `WebhookSubscriptionId`. The subscription resource is the list API URL. Expiration is about 170 days (SharePoint allows 180). If the id is already set, the response is `{ "alreadyRegistered": true }` and SharePoint is not called. An inactive row is not registered.

Renewal is not implemented. Deactivating a row does not delete the subscription.

## Commands

Functions are already in `src/functions`. These templates match them if you recreate the project:

```bash
npm install
npm test
npm install -g azure-functions-core-tools@4 --unsafe-perm true
func start
func azure functionapp publish <function-app>
```

`func start` is `npm start`. Run Azurite first. `npm test` does not load `src/index.js`, so it does not need the Functions host.

There is no separate `GenerateItemId` HTTP API. The web part must not mint numbers.
