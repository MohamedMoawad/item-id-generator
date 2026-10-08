# Deploy

Files in this folder:

| File | Where it goes |
| --- | --- |
| `item-id-generator.sppkg` | SharePoint app catalog. This is the app you add to the site. |
| `item-id-generator-func-net.zip` | .NET 8 isolated Function App. Upload this zip. It is the webhook package to use. |
| `item-id-generator-func.zip` | Older Node package. Do not upload this onto a .NET Function App. |

The list subscription is a SharePoint webhook, which is the current replacement for a remote event receiver. SharePoint calls `spoWebhook` when an item is added, changed, or deleted. The worker numbers items whose number column is still blank, so an already numbered item is left alone.

## 1. Upload the SharePoint app

Package version **1.6.0.0** is named **AutoGen Feature**. Upload `item-id-generator.sppkg` only. You do not create an Azure Function, and you do not paste a webhook address.

1. Upload the package to the app catalog and add **AutoGen Feature** to the site.
2. Open the list and choose **Request number settings**.
3. Set the formula, leave **Active** checked, and save.
4. Add a new item, then refresh the list. The number shows in the number column.

The number is written while someone has that list open. Items created while the list is closed receive a number the next time the list is opened. The Azure Function section below is optional and is only for numbering when nobody has the list open.

If you still have the original tenant-wide **ItemIdGenerator** solution, remove it from the app catalog and from the catalog recycle bin first. Do not check “Enable this app and add it to all sites.”

1. Download `item-id-generator.sppkg` from this folder.
2. Open the app catalog (`https://<tenant>.sharepoint.com/sites/appcatalog`) and go to **Apps for SharePoint**.
3. Upload `item-id-generator.sppkg` and deploy it. Do not choose "add it to all sites".
4. Open the site collection and go to `https://<tenant>.sharepoint.com/sites/<site>/_layouts/15/addanapp.aspx`. Add **AutoGen Feature**. The modern **New > App** page hides some catalog apps; this classic page lists them.
5. Adding the app creates the list **AutoGenFeatureConfiguration** on that site collection and adds missing columns the first time an owner opens a page.
6. Open any list. In the command bar, choose **Request number settings**. That panel is the configuration for that list. Save stores the row on `AutoGenFeatureConfiguration`. People who can edit the config list can change it. Other people can open the same link and see the settings.

List owners set the formula, the number column, and the reset. They do not see a webhook address. Saving an active rule is enough. Opening the list fills blank number columns.

## 2. Create a .NET Function App and upload the zip

Download [item-id-generator-func-net.zip](item-id-generator-func-net.zip) and keep it zipped. `host.json` and `functions.metadata` are at the root. The metadata file is what makes **spoWebhook**, **RegisterWebhook**, and **processRequestNumber** show in the portal. This is a published .NET 8 isolated app, not the Node zip.

`spo-autogen-function` is a Node app. Do not upload this zip there. Create a new Function App:

1. Azure portal → **Create a resource** → **Function App**.
2. Publish **Code**. Runtime stack **.NET**. Version **8 (Isolated)**. Operating system **Windows** if you use the existing Windows plan `ASP-rgshared-a442`, otherwise Linux.
3. Region **Qatar Central** if you want it next to the other app. Give it a new name. Do not reuse `spo-autogen-function`.
4. Create it. **Settings → Configuration → General settings** should show **.NET 8 Isolated**. `FUNCTIONS_WORKER_RUNTIME` is `dotnet-isolated`. Leave `AzureWebJobsFeatureFlags` unset. That flag is only for the Node package.

In Microsoft Entra, before the webhook can write numbers:

1. App registration with SharePoint application permission **Sites.Manage.All**, and admin consent. That one consent covers every site collection. Do not use **Sites.Selected** when the app must number lists across the organization.
2. Create a client secret. Keep it out of git. Put that secret in `SHAREPOINT_CLIENT_SECRET`.

### App settings

On the Function App, open **Settings** → **Environment variables** → **App settings**. Add these, then **Apply** and restart.

| Name | Value |
| --- | --- |
| `NUMBERING_CONFIG_LIST_TITLE` | `AutoGenFeatureConfiguration` |

If this setting is still `RequestNumberConfig`, change it. The function reads **AutoGenFeatureConfiguration**.
| `SHAREPOINT_TENANT_ID` | Directory (tenant) id |
| `SHAREPOINT_CLIENT_ID` | Application (client) id |
| `SHAREPOINT_CLIENT_SECRET` | The client secret |
| `FUNCTIONS_WORKER_RUNTIME` | `dotnet-isolated` |

`SHAREPOINT_WEBHOOK_CLIENT_STATE` and `SPO_WEBHOOK_NOTIFICATION_URL` are optional. Leave them empty. The function builds the spoWebhook address from the RegisterWebhook call and creates its own client state from the client secret. Do not paste the RegisterWebhook URL into `SPO_WEBHOOK_NOTIFICATION_URL`.

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

### The one link to copy

1. On the Function App, open **CORS** (under **API**) and add `https://<tenant>.sharepoint.com`, then save. Without this, the list panel cannot call the function.
2. Open `RegisterWebhook` → **Get function URL** and copy the whole URL, including `?code=`.
3. In SharePoint, open the list → **Request number settings**. Paste that URL into **One-time service address**. Leave **Active** checked. Save.

Save creates the number column if it is missing, stores the address, and subscribes the list. The success line is **Saved. New items on this list will get the next request number.** Add a new item after that. The number column fills in a few seconds later.

`publish-function.ps1` creates a Node Function App and uploads the older Node zip. Use the portal steps above for the .NET package.

## 3. Turn numbering on for a list

1. Add a single-line text column on the target list, for example internal name `RequestNumber`.
2. Open that list and choose **Request number settings**.
3. Set the formula, the number column, and the reset. Leave **Active** checked and save.

You do not paste an address on this list. Save uses the address stored for the organization and subscribes the list. The same thing happens when someone opens a list that already has an active rule. From then on, a new item from the form, the grid, or Power Automate is queued and numbered.

There is no separate classic event receiver to attach. Do not also add a Power Automate flow that writes the same number column.
