# AutoGen Flow

AutoGen Flow gives each list its own running number. Open a list, choose **AutoGen Flow**, and set the number column and the formula. Saving the rule creates an automated cloud flow for that list. The trigger is When an item is created. Leave the number column empty. The flow writes the number.

The number is written while someone has that list open. An item created while nobody has the list open receives its number the next time the list is opened.

## Upload

The package is [deploy/autogen-flow.sppkg](../../deploy/autogen-flow.sppkg), solution **AutoGen Flow**, version **1.39.0.0**.

1. Upload it in the SharePoint app catalog and deploy it.
2. The package uses **Microsoft Flow Service** / **Flows.Read.All** and **Flows.Manage.All**. Update the app on the site after the catalog deploy.
3. On the site collection, open **Site contents → Add an app → From your organization** and add **AutoGen Flow**.
4. Open a list or library. On the command bar choose **AutoGen Flow**.
5. Choose **Add AutoGen Configuration**.
6. Select the single-line text column and the formula. **Columns on this list** shows every column on that list. A lookup or person column is inserted as `{Department.title}`. Leave **Active** checked. **Save settings**.
7. Add an item and leave the number column empty. The number appears on the list.

`AutoGenFlowConfiguration` appears in site contents. Each list has one row. The popup shows that row with **Edit**, **View**, and **Delete**.

`AutoGenFlowLog` appears in site contents. Sort it by **Created**, newest first. **Create flow** records the flow name, environment id, flow id, state, and the error text when the create call fails. Id none and State none mean Power Automate did not return a flow. The open list also records **Load rule**, **Load items**, **Plan number**, **Write number**, **Save count**, and **Refresh list**.

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

Date tokens use the item’s Created time in UTC. `{Department.title}` uses the lookup or person display text. Changing the formula, the column, or the count applies to the next blank item on the open list.

## What happens on save

1. The dialog writes one row in `AutoGenFlowConfiguration` on the site collection.
2. Blank numbers already on this list are filled from that row.
3. While the list stays open, a blank number column is filled with the next code.

Delete removes the list row. New items on that list stay blank.

## Local development

From `spfx/AutoGenFlow`, with Node `>=22.14.0 <23`:

```bash
npm install
npm test
npm run build
```

`npm run build` runs the tests and writes `sharepoint/solution/autogen-flow.sppkg`. Copy that file to `deploy/` after a release build.
