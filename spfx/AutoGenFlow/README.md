# AutoGen Flow

AutoGen Flow gives each list its own running number. Open a list, choose **AutoGen Flow**, and use the same popup as a list setting: add a rule, edit it, or delete it. One numbering flow for the site collection reads `AutoGenFlowConfiguration` and writes the number on each related list. You choose the Power Automate environment. The flow uses this account’s SharePoint connection.

The flow runs about once a minute. It loads the active rows and, on each related list, fills a blank number column. Runs wait in line, so items created together keep a continuous sequence.

## Upload

The package is [deploy/autogen-flow.sppkg](../../deploy/autogen-flow.sppkg), solution **AutoGen Flow**, version **1.33.0.0**.

1. Upload it in the SharePoint app catalog and deploy it.
2. The package uses **Microsoft Flow Service** / **Flows.Read.All** and **Flows.Manage.All**. Those two permissions are already enough. Update the app on the site after the catalog deploy.
3. On the site collection, open **Site contents → Add an app → From your organization** and add **AutoGen Flow**.
4. Open a list or library. On the command bar choose **AutoGen Flow**.
5. Choose **Add AutoGen Configuration**.
6. Choose **Environment**. Select the single-line text column and the formula. Leave **Active** checked. **Save settings**. Power Automate opens in a new tab so this account can sign in. The app then creates or updates one numbering flow for the site collection. Saving another list updates that same flow.
7. On the saved rule, open **Settings** and choose **Manage flow**. That opens the same Power Automate tab. When the flow is on, the tab shows the flow and its run history.

`AutoGenFlowConfiguration` appears in site contents. Each list has one row. The popup shows that row with **Edit**, **View**, **Manage flow**, and **Delete**. The SharePoint connection needs Contribute on each numbered list and on the site collection.

An older flow that numbered only one list is replaced the next time you save a rule on that list. Delete removes the list’s row and leaves the site numbering flow running for the other lists.

## A list rule

| Setting | Meaning |
| --- | --- |
| Number column | The single line of text that receives the code. |
| Formula | `{counter}` or `{counter:4}`, plus `{yyyy}` `{MM}` `{dd}` and columns such as `{Title}`. |
| Pad length | Used when the formula says `{counter}` without a width. |
| Reset period | Never, every day, every month, or every year. |
| Last number used | The next item adds one. `0` starts at `1`. |
| Active | An inactive rule keeps its row and receives no new numbers. |

`REQ-{counter}` with pad length 5 and a last number of 0 shows **REQ-00001**. A formula typed as `{counter:4}{counter}` is saved as `{counter:4}`.

Date tokens use the item’s Created time in UTC. `{Department.title}` uses the lookup or person display text. Changing the formula, the column, or the count is picked up on the next flow run. Save once after an app upgrade so the flow definition is replaced.

## What happens on save

1. The dialog writes one row in `AutoGenFlowConfiguration` on the site collection.
2. The dialog lists Power Automate environments. Choose one. The default environment is selected first.
3. It uses this account’s SharePoint connection in that environment.
4. It creates or updates one flow for the site collection, bound to that connection, and turns the flow on.
5. The same flow id and connection are stored on the list row and on the site row.

Delete removes the list row. The site numbering flow stays on.

## What the flow does

About once a minute, the flow loads the active rows from `AutoGenFlowConfiguration`. For each row it opens that list and fills a blank number column. It reserves the next integer, writes the code, and stores the new count. A failed write puts the count back. A blank number is filled within about a minute after the item is created or saved. Status **Active** means the site flow is running.

## Local development

From `spfx/AutoGenFlow`, with Node `>=22.14.0 <23`:

```bash
npm install
npm test
npm run build
```

`npm run build` runs the tests and writes `sharepoint/solution/autogen-flow.sppkg`. Copy that file to `deploy/` after a release build.
