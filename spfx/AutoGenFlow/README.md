# AutoGen Flow

AutoGen Flow gives each list its own running number. Open a list, choose **AutoGen Flow**, and use the same popup as a list setting: add a rule, edit it, or delete it. Saving an active rule creates that list’s numbering flow. You choose the Power Automate environment. The flow uses this account’s SharePoint connection.

The flow reads that list’s row in `AutoGenFlowConfiguration` on the site collection and writes the next number when an item is created. Runs wait in line, so items created together keep a continuous sequence.

## Upload

The package is [deploy/autogen-flow.sppkg](../../deploy/autogen-flow.sppkg), solution **AutoGen Flow**, version **1.9.0.0**.

1. Upload it in the SharePoint app catalog and deploy it.
2. The app calls Power Automate with **Microsoft Flow Service** / **Flows.Read.All** and **Flows.Manage.All**. Those are the same two permissions an alert app already asks a Global Administrator to approve under **Advanced → API access**. If both rows are already under **Approved requests**, upload this package and leave that page alone.
3. On the site collection, open **Site contents → Add an app → From your organization** and add **AutoGen Flow**.
4. Open a list or library. On the command bar choose **AutoGen Flow**.
5. Choose **Add AutoGen Configuration**.
6. Choose **Environment**. Select the single-line text column and the formula. Leave **Active** checked. **Save settings**. The app creates the flow with this account’s SharePoint connection. If this list already has a rule, open **AutoGen Flow** and save again so the flow is replaced.
7. On the saved rule, choose **Open flow in Power Automate**. That opens the flow named **AutoGen ·** and the list name.

`AutoGenFlowConfiguration` appears in site contents. Each list has one row. The popup shows that row with **Edit**, **View**, **Open flow**, and **Delete**. The SharePoint connection needs Contribute on this list and on the site collection.

If an older site-wide flow named **AutoGen ·** and the site name is still turned on, turn it off in Power Automate. Saving an active list rule also turns that older flow off when this account can reach it.

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

Date tokens use the item’s Created time in UTC. `{Department.title}` uses the lookup or person display text. Changing the formula, the column, or the count is picked up on the next item. Save again after an app upgrade so the flow definition is replaced.

## What happens on save

1. The dialog writes one row in `AutoGenFlowConfiguration` on the site collection.
2. The dialog lists Power Automate environments. Choose one. The default environment is selected first.
3. It uses this account’s SharePoint connection in that environment.
4. It creates a flow whose trigger is **When an item is created or modified** on this list, bound to that connection, and turns the flow on.
5. The flow id and connection are stored on the same row.

Delete stops and removes that flow, then deletes the row.

## What the flow does

When an item is created or saved, the flow loads that list’s row from `AutoGenFlowConfiguration`. If the number column is still empty, it reserves the next integer, writes the code, and stores the new count. A failed write puts the count back. An item that was created before the flow was saved stays blank until you open that item and save it.

## Local development

From `spfx/AutoGenFlow`, with Node `>=22.14.0 <23`:

```bash
npm install
npm test
npm run build
```

`npm run build` runs the tests and writes `sharepoint/solution/autogen-flow.sppkg`. Copy that file to `deploy/` after a release build.
