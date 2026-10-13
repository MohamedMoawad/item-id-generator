# ItemIdGenerator

SharePoint Framework 1.23.2 app. The list page assigns the request number. The web part stores the rule on `AutoGenFeatureConfiguration`.

On the site collection (`pageContext.site.absoluteUrl`):

1. **Ensure AutoGenFeatureConfiguration** creates the list and any missing columns (`TargetListUrl`, `TargetListGuid`, `NumberColumnInternalName`, `Formula`, `CurrentCount`, `ResetPeriod`, `LastResetDate`, `IsActive`, `PadLength`). Field XML is added with `CreateFieldAsXml` and Options 25.
2. Paste a list or library URL and choose **Resolve list**. The web part calls `GetList` and stores the GUID.
3. **Save rule** writes the row. An edit does not send `CurrentCount` or `LastResetDate`.

Create the number column, such as `RequestNumber`, on the target list yourself. The list command **Autogen Feature** is the settings dialog people use day to day.

## Commands

```bash
npm install
npm test
npm run build
```

`npm run build` runs the tests and writes `sharepoint/solution/item-id-generator.sppkg`.

## Deploy

Upload that package to the app catalog and deploy it. On the site collection, open **Site contents → Add an app → From your organization** and add **AutoGen Feature**.
