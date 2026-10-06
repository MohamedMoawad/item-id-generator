# Item ID generator

Starter monorepo for SharePoint Online conditional item IDs, in the spirit of an Infowise-style "generate an ID when a condition matches" rule. A SharePoint Framework web part asks an Azure Function for an ID, then writes that ID back to a list column. The Function can optionally notify a webhook.

```
SharePoint page
  ItemIdGenerator web part
    POST { listName, condition, itemId }
      -> Azure Function GenerateItemId
           -> rules (default, VIP, or rules you add)
           -> optional POST { id, itemId, condition, listName } to WEBHOOK_URL
    <- HTTP 200 { id }
    MERGE the ID into the list item field (SPHttpClient)
```

The two projects install separately. This is not an npm workspace: SharePoint Framework's Heft rig breaks when its packages are hoisted.

| Path | What it is |
| --- | --- |
| `spfx/ItemIdGenerator` | SPFx 1.23.2 React web part. npm package name `item-id-generator`. |
| `functions/ItemIdGeneratorFunc` | Azure Functions v4 Node.js HTTP trigger `GenerateItemId`. |

## Versions

Chosen from the current stable releases on 6 October 2026, not the SPFx 1.24 preview.

| Piece | Version | Why |
| --- | --- | --- |
| Node.js | 22.14.0 up to, but not including, 23 | SPFx 1.23 engine range. Azure Functions Node.js v4 also supports Node 22. |
| `@microsoft/generator-sharepoint` | 1.23.2 (`latest`) | Current SharePoint Online stable generator. 1.24 is still `next` (1.24.0-rc.0). |
| Toolchain | Heft (generator default) | Gulp is opt-in with `--use-gulp` and is legacy from SPFx 1.22 onward. |
| React | 17.0.1 | The version the 1.23 generator scaffolds. React 18 lands with SPFx 1.24. |
| `@azure/functions` | 4.16.5 | Current v4 programming model. |

## Local development order

1. **Function.** From `functions/ItemIdGeneratorFunc`, copy `local.settings.json.example` to `local.settings.json`, replace the `<tenant>` CORS placeholder, run `npm install`, `npm test`, then `func start`. Copy the function key from the console.
2. **Web part.** From `spfx/ItemIdGenerator`, run `npm install`, then `npm start`. In the property pane, set the Function URL to `http://localhost:7071/api/GenerateItemId?code=<function-key>`.
3. **List.** Create a single-line text column (example internal name `GeneratedItemId`) on the list named in the property pane (example title `Requests`).
4. Open the web part, enter a list item ID, and generate. The part shows the ID and writes it to that column.

Details, including App Catalog deployment, are in each project's README.

## Environment variables

Set these in `functions/ItemIdGeneratorFunc/local.settings.json` locally, and in the Function App configuration when you deploy. Do not commit `local.settings.json` or real keys.

| Name | Required | Purpose |
| --- | --- | --- |
| `FUNCTIONS_WORKER_RUNTIME` | yes | `node` |
| `AzureWebJobsStorage` | local tools | `UseDevelopmentStorage=true` when Azurite is running. HTTP-only local runs can use an empty string if storage is not needed. |
| `WEBHOOK_URL` | no | When non-empty, the Function POSTs `{ id, itemId, condition, listName }` after it generates an ID. When empty, the webhook is skipped. |
| `APPLICATIONINSIGHTS_CONNECTION_STRING` | no | Application Insights. Leave empty until the resource exists. |

The web part does not use environment variables. Its Function URL, list name, target field internal name, and condition are property pane settings.

## ID rules

| Condition | ID |
| --- | --- |
| anything except VIP, including `default` | `{listName}-{timestamp}` |
| `VIP` (case-insensitive) | `VIP-{listName}-{timestamp}` |

`timestamp` is Unix epoch milliseconds. Whitespace in the list title collapses to a hyphen. Add department, role, or column-value rules in `functions/ItemIdGeneratorFunc/src/rules/idRules.js` with `registerRule`. First match wins.

## Next steps

**App registration (Azure AD).** The Function is registered with `authLevel: 'function'`, so a function key in the web part URL works for a prototype. Anyone who can edit the page can read that key. For a real tenant, turn on App Service authentication (Easy Auth) with a Microsoft Entra app registration and stop embedding the key. Placeholder values only:

- Application (client) ID: `<application-client-id>`
- Directory (tenant) ID: `<directory-tenant-id>`
- Application ID URI: `api://<application-id-uri>`
- Redirect URI: `https://<function-app>.azurewebsites.net/.auth/login/aad/callback`

No client ID is included in this repo. After Easy Auth is on, set the Function `authLevel` to `anonymous` so the platform, not a key, is the gate, and call the API with `AadHttpClient` instead of `fetch`. A permission request you would add later in `spfx/ItemIdGenerator/config/package-solution.json` looks like this, still with placeholders:

```json
"webApiPermissionRequests": [
  {
    "resource": "api://<application-id-uri>",
    "scope": "access_as_user"
  }
]
```

**CORS.** Local `local.settings.json` is not deployed. On the Function App, allow `https://<tenant>.sharepoint.com`. Do not leave `*` in production. The hosted SharePoint workbench retires on 1 December 2026; debug against a real page.

**Application Insights.** Create an Application Insights resource, then set `APPLICATIONINSIGHTS_CONNECTION_STRING` on the Function App. `host.json` already enables request sampling. Do not commit the connection string.

**SharePoint package.** Build and upload the `.sppkg` from the SPFx README. Create the target text column before authors use the web part.

**Uniqueness.** Epoch milliseconds are enough for a starter and can collide if two items are generated in the same millisecond. Add a short random suffix in the rules module if that matters.
