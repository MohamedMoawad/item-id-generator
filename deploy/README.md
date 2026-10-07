# Deploy

Two files in this folder:

| File | Where it goes |
| --- | --- |
| `item-id-generator.sppkg` | SharePoint app catalog. This is the app you add to the site. |
| `item-id-generator-func.zip` | Azure Function App. This is the webhook that numbers new items. |

The list subscription is a SharePoint webhook, which is the current replacement for a remote event receiver. SharePoint calls `spoWebhook` when an item is added, changed, or deleted. The worker numbers items whose number column is still blank, so an already numbered item is left alone.

## 1. Upload the SharePoint app

Package version **1.1.0.0** is a site app, not a tenant-wide package. If 1.0.0.0 is already in the catalog, remove that solution first, then upload this file. SharePoint cannot switch an existing solution from tenant-wide deployment to a site app.

1. Download `item-id-generator.sppkg` from this folder.
2. Open the app catalog (`https://<tenant>.sharepoint.com/sites/appcatalog`) and go to **Apps for SharePoint**.
3. Upload `item-id-generator.sppkg` and deploy it. Do not choose "add it to all sites".
4. Open the site collection. **Site contents** > **New** > **App** > add **ItemIdGenerator**.
5. Adding the app creates the list **RequestNumberConfig** on that site collection and adds missing columns the first time an owner opens a page.
6. Open any list. In the command bar, choose **Request number settings**. That panel is the configuration for that list. Save stores the row on `RequestNumberConfig`. People who can edit the config list can change it. Other people can open the same link and see the settings.

Paste the **RegisterWebhook** URL into that panel once. It is stored for the whole site collection. Each list keeps its own formula, column, and reset.

## 2. Publish the Azure Function

This environment cannot sign in to your Azure subscription. Run the script from your machine in PowerShell.

Before that, in Microsoft Entra:

1. App registration, application permission **Sites.Selected**, admin consent.
2. Grant that app access to the site collection that will hold `RequestNumberConfig` and to each target web. SharePoint admin center or `Grant-PnPAzureADAppSitePermission`.
3. Create a client secret. Keep it out of git.
4. Pick a long random string for the webhook client state. That same string is `SHAREPOINT_WEBHOOK_CLIENT_STATE`.

```powershell
az login
cd deploy
.\publish-function.ps1 `
  -ResourceGroup 'item-id-rg' `
  -FunctionAppName 'item-id-func' `
  -StorageAccount 'itemidfuncstore' `
  -Location 'eastus' `
  -SharePointSiteUrl 'https://<tenant>.sharepoint.com/sites/<site>' `
  -TenantId '<directory-tenant-id>' `
  -ClientId '<application-client-id>' `
  -ClientSecret '<client-secret>' `
  -WebhookClientState '<webhook-client-state>' `
  -SharePointOrigin 'https://<tenant>.sharepoint.com'
```

The script creates the resource group, storage account, and Node 22 Function App, deploys the zip, turns on remote `npm install`, saves `SPO_WEBHOOK_NOTIFICATION_URL`, and allows the SharePoint origin through CORS. It prints two URLs.

Storage account names are 3–24 lowercase letters and numbers, and must be globally unique. Function App names must be globally unique too.

## 3. Register the webhook on the list

1. Edit the **Request number config** web part.
2. Property pane: list title `RequestNumberConfig`.
3. Paste the **RegisterWebhook** URL from the script, including `?code=`. Page editors can see that URL.
4. Add a single-line text column on the target list, for example internal name `RequestNumber`.
5. On the page, choose **Ensure RequestNumberConfig**, paste the target list URL, **Resolve list**, then **Save rule** with **Active** checked.

Save calls `RegisterWebhook`. That function subscribes the target list to `spoWebhook` and stores the subscription id on the row. From then on, a new item from the form, the grid, or Power Automate is queued and numbered.

There is no separate classic event receiver to attach. Do not also add a Power Automate flow that writes the same number column.
