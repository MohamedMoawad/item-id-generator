# ItemIdGenerator

SharePoint Framework 1.23.2 React web part. It configures request numbers. It does not generate them.

On the site collection (`pageContext.site.absoluteUrl`):

1. **Ensure AutoGenFeatureConfiguration** creates the generic list and any missing columns (`TargetListUrl`, `TargetListGuid`, `NumberColumnInternalName`, `Formula`, `CurrentCount`, `ResetPeriod`, `LastResetDate`, `IsActive`, `PadLength`, `WebhookSubscriptionId`). Field XML is added with `CreateFieldAsXml` and Options 25. Opening the page does not create the list. Saving a rule ensures it first.
2. Paste a list or library URL and choose **Resolve list**. The web part calls `GetList` on the web derived from that URL and stores the GUID.
3. **Save rule** writes the row. The body for an edit does not include `CurrentCount` or `LastResetDate`. If the target GUID changed, `WebhookSubscriptionId` is cleared.
4. When the row is active and the property pane has a real `RegisterWebhook` URL, the web part POSTs `{ configSiteUrl, configItemId }`. A placeholder URL (anything containing `<`) saves the row and warns that the webhook was not registered. A registration failure does not roll back the row.

Create the number column, such as `RequestNumber`, on the target list yourself.

The property pane holds the config list title (default `AutoGenFeatureConfiguration`) and the full `RegisterWebhook` URL, including the function key. Page editors can read that URL. Function CORS must allow `https://<tenant>.sharepoint.com`.

## Commands

```bash
npm install
npm test
npm start
npm run package-solution
```

The generator uses Heft. `npm start` is the old `gulp serve`. `npm run package-solution` is `gulp package-solution --ship`. The debug server listens on `https://localhost:4321`. `config/serve.json` loads the hosted workbench. That workbench retires on 2026-12-01.

`npm test` covers the pure contract: create versus update bodies, URL parsing, field internal names, and the RegisterWebhook payload. It does not import `@microsoft/sp-http`.

## Deploy

Upload `sharepoint/solution/item-id-generator.sppkg` to the App Catalog. `skipFeatureDeployment` is true, so make the solution available to all sites when the catalog asks. Add **AutoGen Feature** to a page on the site collection.

The web part uses the current user's `SPHttpClient`. That user needs permission to create a list on the site collection and to read the target list. Webhook registration itself is app-only inside the Function, because SharePoint calls `spoWebhook` back during subscribe and the browser cannot finish that handshake.

This UI was not opened against a live tenant in this environment.
