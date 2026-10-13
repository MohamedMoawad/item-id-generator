# Deploy

Upload `item-id-generator.sppkg` to the SharePoint app catalog. One upload updates every site that already has the app.

Package version **1.31.0.0** is named **AutoGen Feature**. A new item saved from the open list stores its number in that same save. The counter moves once for that item. Add it from the site: **Site contents → Add an app → From your organization → AutoGen Feature**. **AutoGenFeatureConfiguration** stays visible in site contents. The dialog is 40% of the screen width. It shows **Add AutoGen Configuration** until the list is saved, then one row with Modified, Modified by, Edit, and View. A new rule starts with Reset set to **None**.

1. Upload `item-id-generator.sppkg` in the app catalog, replace the existing app, and deploy it. If SharePoint refuses the upgrade, delete **AutoGen Feature** from the app catalog and upload this package again.
2. On the site collection, open **Site contents**, choose **Add an app**, open **From your organization**, and add **AutoGen Feature**.
3. Open a list and choose **Autogen Feature** on the command bar. The dialog opens in the middle of the page.
4. Select the single-line text column, set the formula, leave **Active** checked, and save.
5. Add a new item and leave the number column empty. The number is part of the save. When the form closes, the list page reloads and the number appears. You do not type it or click Refresh.

The number is written while someone has that list open. Items created while the list is closed, including items created by Power Automate, receive a number the next time the list is opened.

Later versions are the same upload. Replace the package in the app catalog and deploy it again. On each site, open **Site contents** and update **AutoGen Feature** if SharePoint shows an update. Refresh the list.

List owners set the formula, the number column, and the reset. Saving an active rule is enough. Opening the list fills blank number columns.
