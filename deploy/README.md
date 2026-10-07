# Deploy

Two files in this folder:

| File | Where it goes |
| --- | --- |
| `item-id-generator.sppkg` | SharePoint app catalog. This is the app you add to the site. |
| `item-id-generator-func.zip` | Upload this zip to the Function App. It already contains the Node packages. |

The list subscription is a SharePoint webhook, which is the current replacement for a remote event receiver. SharePoint calls `spoWebhook` when an item is added, changed, or deleted. The worker numbers items whose number column is still blank, so an already numbered item is left alone.

## 1. Upload the SharePoint app

Package version **1.2.1.0** is named **Request number config**. Upload this file over the previous package in the app catalog and choose to replace it. Version 1.2.1 fixes the HTTP 400 on **Request number settings** (`__metadata` is not valid on `SP.XmlSchemaFieldCreationInformation`). If you still have the original tenant-wide **ItemIdGenerator** solution, remove it from the app catalog and from the catalog recycle bin first. Do not check “Enable this app and add it to all sites.”

1. Download `item-id-generator.sppkg` from this folder.
2. Open the app catalog (`https://<tenant>.sharepoint.com/sites/appcatalog`) and go to **Apps for SharePoint**.
3. Upload `item-id-generator.sppkg` and deploy it. Do not choose "add it to all sites".
4. Open the site collection and go to `https://<tenant>.sharepoint.com/sites/<site>/_layouts/15/addanapp.aspx`. Add **Request number config**. The modern **New > App** page hides some catalog apps; this classic page lists them.
5. Adding the app creates the list **RequestNumberConfig** on that site collection and adds missing columns the first time an owner opens a page.
6. Open any list. In the command bar, choose **Request number settings**. That panel is the configuration for that list. Save stores the row on `RequestNumberConfig`. People who can edit the config list can change it. Other people can open the same link and see the settings.

Paste the **RegisterWebhook** URL into that panel once. It is stored for the whole site collection. Each list keeps its own formula, column, and reset.

## 2. Upload the Azure Function zip

Download [item-id-generator-func.zip](item-id-generator-func.zip) and keep it as a zip. Do not unzip it and zip the folder again. `host.json` has to sit at the root of the archive, and the Node packages are already inside.

The Function App has to be **Code**, **Node.js 22**, **Functions 4.x**, **Linux**. Consumption or App Service is fine. A container app cannot take this zip.

In Microsoft Entra, before the webhook can write numbers:

1. App registration with application permission **Sites.Selected**, and admin consent.
2. Grant that app access to the site collection that holds `RequestNumberConfig` and to each target web.
3. Create a client secret. Keep it out of git.
4. Pick a long random string. That string is `SHAREPOINT_WEBHOOK_CLIENT_STATE`.

### App settings

On the Function App, open **Settings** → **Environment variables** → **App settings**. Add these, then **Apply** and restart.

| Name | Value |
| --- | --- |
| `NUMBERING_CONFIG_SITE_URL` | `https://<tenant>.sharepoint.com/sites/<site>` |
| `NUMBERING_CONFIG_LIST_TITLE` | `RequestNumberConfig` |
| `SHAREPOINT_TENANT_ID` | Directory (tenant) id |
| `SHAREPOINT_CLIENT_ID` | Application (client) id |
| `SHAREPOINT_CLIENT_SECRET` | The client secret |
| `SHAREPOINT_WEBHOOK_CLIENT_STATE` | The random string from above |
| `WEBSITE_NODE_DEFAULT_VERSION` | `~22` |
| `AzureWebJobsFeatureFlags` | `EnableWorkerIndexing` |
| `FUNCTIONS_NODE_BLOCK_ON_ENTRY_POINT_ERROR` | `true` |
| `FUNCTIONS_WORKER_RUNTIME` | `node` |

`AzureWebJobsFeatureFlags=EnableWorkerIndexing` is what makes `spoWebhook`, `RegisterWebhook`, and `processRequestNumber` show on the **Functions** page. Without it the host finds no `function.json` files and the page stays empty after a successful zip deploy. If that setting already has another value, add `,EnableWorkerIndexing` on the end. Do not replace the other flags.

Leave `WEBSITE_RUN_FROM_PACKAGE` unset. Leave `SCM_DO_BUILD_DURING_DEPLOYMENT` unset. This zip is already built.

`AzureWebJobsStorage` is created with the Function App. That connection is the queue and the numbering lock. Do not delete it.

### Zip upload

**Deployment Center** connects GitHub. It does not take a zip file. The upload page is one menu away:

1. Open the Function App.
2. **Development Tools** → **Advanced Tools** → **Go**. A new tab opens.
3. In that tab, **Tools** → **Zip Push Deploy**.
4. Drag `item-id-generator-func.zip` onto the page.
5. Wait until the log says the deployment succeeded.
6. Back on the Function App, confirm the app settings above, especially `AzureWebJobsFeatureFlags`, then **Restart**.
7. Open **Functions** and choose **Refresh**. You should see `spoWebhook`, `RegisterWebhook`, and `processRequestNumber`.

### URLs to copy

1. Open `spoWebhook` → **Function keys** → copy the **default** key.
2. Add this app setting, then apply and restart again:

   `SPO_WEBHOOK_NOTIFICATION_URL` = `https://<function-app>.azurewebsites.net/api/spoWebhook?code=<that-key>`

3. Open `RegisterWebhook` → **Get function URL** and copy it. That is the URL you paste into **Request number settings** in SharePoint.
4. On the Function App, open **CORS** (under **API**) and add `https://<tenant>.sharepoint.com`, then save.

`publish-function.ps1` is an optional Azure CLI path that creates the app and uploads the same zip. The portal steps above are enough.

## 3. Register the webhook on the list

1. Add a single-line text column on the target list, for example internal name `RequestNumber`.
2. Open that list and choose **Request number settings**.
3. Paste the **RegisterWebhook** URL, including `?code=`. People who can edit the page can see that URL.
4. Set the formula, the number column, and the reset, leave **Active** checked, and save.

Save calls `RegisterWebhook`. That function subscribes the target list to `spoWebhook` and stores the subscription id on the row. From then on, a new item from the form, the grid, or Power Automate is queued and numbered.

There is no separate classic event receiver to attach. Do not also add a Power Automate flow that writes the same number column.
