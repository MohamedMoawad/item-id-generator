# ItemIdGenerator

SharePoint Framework 1.23.2 React web part for SharePoint Online. It POSTs `{ listName, condition, itemId }` to the `GenerateItemId` Azure Function, shows the returned ID, and writes that ID to a list column with `SPHttpClient`.

Scaffolded with `@microsoft/generator-sharepoint` 1.23.2 (`yo @microsoft/sharepoint`), React, SharePoint Online only. The generator's default toolchain is Heft. The npm package name is `item-id-generator` (the generator kebab-cases the solution name). The web part and the App Catalog solution name are `ItemIdGenerator`.

Node.js must be at least 22.14.0 and below 23.

## Property pane

| Setting | Example | Sent to |
| --- | --- | --- |
| Function URL | `https://<function-app>.azurewebsites.net/api/GenerateItemId` | `fetch` POST |
| List name | `Requests` | JSON `listName`, and the list updated afterwards |
| Target field internal name | `GeneratedItemId` | SharePoint field on the list item |
| Condition | `Default` or `VIP` | JSON `condition` |

The list item ID is typed in the web part, because it changes per item. It is posted as `itemId` and used in the REST URL `items(<id>)`.

Create the target column first: single line of text. The internal name is the `Field=` value in the column settings URL, not the display name. `src/webparts/itemIdGenerator/services/updateListItemField.ts` documents the list title and the field internal name next to the MERGE request.

While the Function uses a function key, the property pane URL is:

```text
http://localhost:7071/api/GenerateItemId?code=<function-key>
```

Do not store a real key in the manifest or in git. The manifest ships the placeholder `https://<function-app>.azurewebsites.net/api/GenerateItemId`. Page editors can read property values, so a function key in the URL is only for local debugging.

## Local serve

`npm start` is this project's debug command. It runs `heft start --clean`, which serves `https://localhost:4321`. That replaces `gulp serve`.

Gulp is not installed here. SPFx 1.22 and later default to Heft. `gulp serve` and `gulp package-solution --ship` apply only if you scaffold again with `yo @microsoft/sharepoint --use-gulp`.

```bash
cd spfx/ItemIdGenerator
npm install
npm start
```

Trust the developer certificate when the toolchain asks. `config/serve.json` opens `https://{tenantDomain}/_layouts/workbench.aspx`. Replace `{tenantDomain}` with your tenant host if serve does not prompt for it.

The hosted workbench retires on 1 December 2026. For a normal page, load the debug manifests on that page:

```text
?debug=true&noredir=true&debugManifestsFile=https://localhost:4321/temp/manifests.js
```

`npm test` runs the Heft Jest suite. It covers the JSON body and the list MERGE URL. It does not call Azure or SharePoint.

## Package and App Catalog

`npm run build` runs the production tests and then packages the solution. `npm run package-solution` only packages. That command replaces `gulp package-solution --ship`.

```bash
npm run build
```

Upload `sharepoint/solution/item-id-generator.sppkg` to the tenant App Catalog (or a site collection catalog). The package was built with `skipFeatureDeployment: true`, so when the catalog asks, make the solution available to all sites. Add the **ItemIdGenerator** web part to a modern page, edit the web part, and fill in the property pane.

Deploy the Function before authors use the page. The browser calls the Function directly, so the Function App has to allow this site's origin. See CORS below.

`config/deploy-azure-storage.json` still has the generator placeholders `<!-- STORAGE ACCOUNT NAME -->` and `<!-- ACCESS KEY -->`. Leave them, or replace them locally and do not commit a real access key. Client-side assets are included in the `.sppkg` (`includeClientSideAssets` is true), so Azure CDN deployment is optional.

## CORS

The web part uses `fetch` from the SharePoint page, not `SPHttpClient`, because the Function is not on the SharePoint domain. The Function App must allow:

- `https://<tenant>.sharepoint.com`
- `https://localhost:4321` during local debug

Local Function CORS is `Host.CORS` in `functions/ItemIdGeneratorFunc/local.settings.json`. Production CORS is configured on the Function App. Details and the Azure AD placeholder are in that project's README and the repository README.

If the browser reports a CORS or network failure, the web part keeps the message on the page and does not write the list item.

## List update

After a successful Function response, `updateListItemField` sends `SPHttpClient.post` with `X-HTTP-Method: MERGE` to:

```text
{web}/_api/web/lists/getbytitle('<List name>')/items(<item id>)
```

The body is `{ "<target field internal name>": "<generated id>" }`. If SharePoint asks for the list item entity type, the comments in `updateListItemField.ts` show the verbose `__metadata` form. The generated ID stays visible even when the MERGE fails, so it can be copied.
