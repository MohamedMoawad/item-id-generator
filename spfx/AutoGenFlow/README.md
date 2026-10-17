# AutoGen

AutoGen gives each list its own running number. Open a list, choose **AutoGen Flow**, set the column and the formula, then choose the Power Automate environment and the SharePoint connection. Turning numbering on creates a flow for that list only.

The flow reads that list’s row in `AutoGenFlowConfiguration` on the site collection and writes the next number when an item is created. A person, a grid edit, and another flow all receive the next number. The counter moves only after the item is updated, and the flow runs one item at a time, so the sequence stays continuous.

## Upload

The package is [deploy/autogen-flow.sppkg](../../deploy/autogen-flow.sppkg), solution **AutoGen Flow**, version **1.2.0.0**.

1. Upload it in the SharePoint app catalog and deploy it.
2. In the SharePoint admin center, open **Advanced → API access** and approve **AutoGen Flow**. This approval is once for the tenant. It lets the app create the flow.
3. On the site collection, open **Site contents → Add an app → From your organization** and add **AutoGen Flow**.
4. Open a list or library. On the command bar choose **AutoGen Flow**.
5. Choose the single-line text column and the formula. **Save**.
6. Choose the environment and the SharePoint connection, then **Turn on numbering**. The connection needs Contribute on this list and on the site collection.

`AutoGenFlowConfiguration` appears in site contents. Each list has one row. Each list also has its own flow, named **AutoGen ·** and the list name, in the environment you chose.

If an older site-wide flow named **AutoGen ·** and the site name is still turned on, turn it off in Power Automate. Otherwise that flow and the list flow both write numbers. Turning numbering on for a list also turns that older site flow off when this account can reach it.

## A list rule

| Setting | Meaning |
| --- | --- |
| Environment | The Power Automate environment that hosts this list’s flow. |
| SharePoint connection | The connection the flow uses. Change it, then choose **Redeploy**. |
| Number column | The single line of text that receives the code. |
| Formula | `{counter}` or `{counter:4}`, plus `{yyyy}` `{MM}` `{dd}` and columns such as `{Title}`. |
| Digits | Used when the formula says `{counter}` without a width. |
| Start again | Never, every day, every month, or every year. |
| Last number used | The next item adds one. `0` starts at `1`. |
| This list | Paused lists keep their rule and receive no new numbers. |

`REQ-{counter}` with 5 digits and a last number of 0 shows **REQ-00001**. A formula typed as `{counter:4}{counter}` is saved as `{counter:4}`, so the digit is written once.

Date tokens use the item’s Created time in UTC. `{Department.title}` uses the lookup or person display text. Changing the formula, the column, or the count is picked up on the next item. Choose **Redeploy** after an app upgrade, or after you change the environment or the connection.

## What the list flow does

When an item is created, the flow loads that list’s active row from `AutoGenFlowConfiguration` on the site collection. If the number column is still empty, it reserves the next integer, writes the code, and stores the new count. A failed write puts the count back. Runs wait in line, so two items created together keep a continuous sequence.

Turn off stops this list’s flow and leaves the saved rule in place. Another list is another flow. Its admin chooses the environment and connection for that list.

## Local development

From `spfx/AutoGenFlow`, with Node `>=22.14.0 <23`:

```bash
npm install
npm test
npm run build
```

`npm run build` runs the tests and writes `sharepoint/solution/autogen-flow.sppkg`. Copy that file to `deploy/` after a release build.
