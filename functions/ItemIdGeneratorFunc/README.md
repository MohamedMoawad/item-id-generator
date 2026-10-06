# ItemIdGeneratorFunc

Azure Functions (Node.js programming model v4) HTTP trigger that generates a conditional SharePoint item ID.

- Runtime: Azure Functions host v4, `@azure/functions` 4.16.5
- Node.js: 22.x (`>=22.14.0 <23`)
- Trigger: `GenerateItemId`, HTTP POST, auth level `function`

## Request and response

```http
POST /api/GenerateItemId
Content-Type: application/json

{ "listName": "Requests", "condition": "VIP", "itemId": 15 }
```

Extra JSON fields are kept as metadata for rules you add (department, role, list column values). They are not required.

Success:

```json
{ "id": "VIP-Requests-1700000000000" }
```

That is HTTP status 200 and a JSON body `{ id }`. The v4 programming model returns it as `{ status: 200, jsonBody: { id } }`, and the host serializes `jsonBody`.

| Condition | ID |
| --- | --- |
| `VIP` (trim, case-insensitive) | `VIP-{listName}-{timestamp}` |
| anything else, including `default` or empty | `{listName}-{timestamp}` |

`listName` is trimmed and internal whitespace becomes a hyphen. `timestamp` is Unix epoch milliseconds. `itemId` must be a positive integer (the SharePoint list item ID).

Validation failures return HTTP 400 and `{ "error": "..." }`.

## Add a rule

Edit `src/rules/idRules.js`. `registerRule` inserts at the front, so the new rule runs before VIP. First match wins.

```javascript
const { registerRule } = require('./rules/idRules');

registerRule({
  id: 'department',
  matches: (ctx) => typeof ctx.metadata.department === 'string' && ctx.metadata.department.trim() !== '',
  format: (ctx) => `${ctx.metadata.department.trim()}-${ctx.listName}-${ctx.timestamp}`
});
```

Call `registerRule` from `src/index.js` after the function module loads, or append to the `rules` array in `idRules.js` if this rule should lose to VIP.

## Optional webhook

If `WEBHOOK_URL` is unset or blank, nothing is called. If it is set, the Function POSTs:

```json
{ "id": "VIP-Requests-1700000000000", "itemId": 15, "condition": "VIP", "listName": "Requests" }
```

A webhook failure is logged and the generated ID is still returned to the caller.

## Prerequisites

- Node.js 22.14 or newer, below 23
- [Azure Functions Core Tools](https://learn.microsoft.com/azure/azure-functions/functions-run-local) v4 (`func`)
- Azurite, only if you keep `AzureWebJobsStorage` as `UseDevelopmentStorage=true`

## Create, run, and test

This folder is already a v4 JavaScript function app (`package.json` `main` loads `src/index.js`, which registers `GenerateItemId`). You do not need to run `func init` again unless you want an empty app beside it.

Equivalent commands the Core Tools would use to create this layout:

```bash
func init ItemIdGeneratorFunc --worker-runtime node --model v4 --language javascript
cd ItemIdGeneratorFunc
func new --name GenerateItemId --template "HTTP trigger" --authlevel function
```

Run what is already here:

```bash
cd functions/ItemIdGeneratorFunc
cp local.settings.json.example local.settings.json
npm install
npm test
func start
```

`func start` prints a function key and a local URL:

```text
http://localhost:7071/api/GenerateItemId
```

Try it:

```bash
curl -X POST "http://localhost:7071/api/GenerateItemId?code=<function-key>" \
  -H "Content-Type: application/json" \
  -d '{"listName":"Requests","condition":"VIP","itemId":15}'
```

`npm test` runs the Node.js test runner (`node --test`) against the rules, the webhook, and the HTTP handler. It does not start the Functions host.

In `local.settings.json`, replace `https://<tenant>.sharepoint.com` with the SharePoint origin that will call the Function. `local.settings.json` is gitignored. The example file has empty secrets only.

For an HTTP-only local run without Azurite, set `AzureWebJobsStorage` to `""`.

## Deploy

Create a Function App on the Azure Functions v4 runtime and Node 22. Then:

```bash
func azure functionapp publish <function-app-name>
```

Set application settings in the portal or with the Azure CLI. Do not put real values in source:

| Setting | Value |
| --- | --- |
| `WEBHOOK_URL` | `https://<webhook-host>/path` or leave empty |
| `APPLICATIONINSIGHTS_CONNECTION_STRING` | From the Application Insights resource |

`host.json` already turns on Application Insights sampling. The connection string is what attaches this app to a resource.

## CORS

Local CORS is the `Host.CORS` value in `local.settings.json`. That file is not deployed.

On the Function App CORS blade, allow the SharePoint origin and the local debug origin you actually use:

- `https://<tenant>.sharepoint.com`
- `https://localhost:4321` while you are debugging SPFx

Do not use `*` once the app is called from SharePoint with anything other than a public anonymous API. This starter does not send credentialed CORS requests (`CORSCredentials` is false). The function key travels in the query string.

## Azure AD authentication

Placeholder names only. This repo does not contain a client ID, tenant ID, or client secret.

The trigger uses `authLevel: 'function'` so local and prototype calls pass `?code=<function-key>`. Treat that key like a password. Do not commit it, and do not leave it in a web part property on a production page: page editors can read property pane values.

Production path:

1. Create an app registration. Client ID `<application-client-id>`, tenant `<directory-tenant-id>`, application ID URI `api://<application-id-uri>`.
2. On the Function App, enable authentication (Easy Auth) with the Microsoft identity provider. Redirect URI `https://<function-app>.azurewebsites.net/.auth/login/aad/callback`.
3. Change `authLevel` to `anonymous` so Easy Auth is the only gate. A function key and Easy Auth at the same time is easy to misconfigure.
4. Call the Function from SPFx with `AadHttpClient` and a `webApiPermissionRequests` entry. The shape is in the repository README.

No sample client ID is checked in on purpose.
