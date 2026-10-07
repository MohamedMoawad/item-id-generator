# Deploy

Files in this folder:

| File | Where it goes |
| --- | --- |
| `item-id-generator.sppkg` | SharePoint app catalog. This is the app you add to the site. |
| `item-id-generator-func-net.zip` | .NET 8 isolated Function App. Upload this zip. It is the webhook package to use. |
| `item-id-generator-func.zip` | Older Node package. Do not upload this onto a .NET Function App. |

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

## 2. Create a .NET Function App and upload the zip

Download [item-id-generator-func-net.zip](item-id-generator-func-net.zip) and keep it zipped. `host.json` and `functions.metadata` are at the root. The metadata file is what makes **spoWebhook**, **RegisterWebhook**, and **processRequestNumber** show in the portal. This is a published .NET 8 isolated app, not the Node zip.

`spo-autogen-function` is a Node app. Do not upload this zip there. Create a new Function App:

1. Azure portal → **Create a resource** → **Function App**.
2. Publish **Code**. Runtime stack **.NET**. Version **8 (Isolated)**. Operating system **Windows** if you use the existing Windows plan `ASP-rgshared-a442`, otherwise Linux.
3. Region **Qatar Central** if you want it next to the other app. Give it a new name. Do not reuse `spo-autogen-function`.
4. Create it. **Settings → Configuration → General settings** should show **.NET 8 Isolated**. `FUNCTIONS_WORKER_RUNTIME` is `dotnet-isolated`. Leave `AzureWebJobsFeatureFlags` unset. That flag is only for the Node package.

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
| `FUNCTIONS_WORKER_RUNTIME` | `dotnet-isolated` |

The portal sets `FUNCTIONS_WORKER_RUNTIME` when you create a .NET 8 Isolated app. On **Windows**, set `WEBSITE_RUN_FROM_PACKAGE` to `1` and save **before** the zip upload. Leave `SCM_DO_BUILD_DURING_DEPLOYMENT` unset. The zip is already built. Do not set `WEBSITE_RUN_FROM_PACKAGE` on a Linux Consumption app.

`AzureWebJobsStorage` is created with the Function App. That connection is the queue and the numbering lock. Do not delete it.

### Zip upload

**Deployment Center** connects GitHub. It does not take a zip file. The upload page is one menu away:

1. Open the Function App.
2. **Development Tools** → **Advanced Tools** → **Go**. A new tab opens.
3. In that tab, **Tools** → **Zip Push Deploy**.
4. Drag `item-id-generator-func-net.zip` onto the page.
5. Wait until the log says the deployment succeeded.
6. Restart the Function App.
7. Open **Functions** and choose **Refresh**. You should see `spoWebhook`, `RegisterWebhook`, and `processRequestNumber`.

### URLs to copy

1. Open `spoWebhook` → **Function keys** → copy the **default** key.
2. Add this app setting, then apply and restart again:

   `SPO_WEBHOOK_NOTIFICATION_URL` = `https://<function-app>.azurewebsites.net/api/spoWebhook?code=<that-key>`

3. Open `RegisterWebhook` → **Get function URL** and copy it. That is the URL you paste into **Request number settings** in SharePoint.
4. On the Function App, open **CORS** (under **API**) and add `https://<tenant>.sharepoint.com`, then save.

`publish-function.ps1` creates a Node Function App and uploads the older Node zip. Use the portal steps above for the .NET package.

## 3. Register the webhook on the list

1. Add a single-line text column on the target list, for example internal name `RequestNumber`.
2. Open that list and choose **Request number settings**.
3. Paste the **RegisterWebhook** URL, including `?code=`. People who can edit the page can see that URL.
4. Set the formula, the number column, and the reset, leave **Active** checked, and save.

Save calls `RegisterWebhook`. That function subscribes the target list to `spoWebhook` and stores the subscription id on the row. From then on, a new item from the form, the grid, or Power Automate is queued and numbered.

There is no separate classic event receiver to attach. Do not also add a Power Automate flow that writes the same number column.
