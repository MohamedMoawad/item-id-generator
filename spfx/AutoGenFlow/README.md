# AutoGen

AutoGen gives every list on a site collection its own running number. Open a list, choose **AutoGen Flow**, set the column and the formula, and turn numbering on once. The same numbering service covers every other list on that site collection.

Numbers are written within a minute, in the order the items were created. A person, a grid edit, and another flow all receive the next number. The counter moves only after the item is updated, and the service runs one pass at a time, so the sequence stays continuous.

## Upload

The package is [deploy/autogen-flow.sppkg](../../deploy/autogen-flow.sppkg), solution **AutoGen Flow**, version **1.1.0.0**.

1. Upload it in the SharePoint app catalog and deploy it.
2. In the SharePoint admin center, open **Advanced → API access** and approve **AutoGen Flow**. This approval is once for the tenant. It lets the app create the numbering flow.
3. On the site collection, open **Site contents → Add an app → From your organization** and add **AutoGen Flow**.
4. Open any list or library. On the command bar choose **AutoGen Flow**.
5. Choose the single-line text column and the formula. **Save**.
6. Choose **Turn on numbering**. Pick the SharePoint connection when the account has more than one. That connection needs Contribute on the site collection.

`AutoGenFlowConfiguration` appears in site contents. Each list has one row. The site has one numbering flow, named **AutoGen ·** and the site name, in the default Power Automate environment.

## A list rule

| Setting | Meaning |
| --- | --- |
| Number column | The single line of text that receives the code. |
| Formula | `{counter}` or `{counter:4}`, plus `{yyyy}` `{MM}` `{dd}` and columns such as `{Title}`. |
| Digits | Used when the formula says `{counter}` without a width. |
| Start again | Never, every day, every month, or every year. |
| Last number used | The next item adds one. `0` starts at `1`. |
| This list | Paused lists keep their rule and receive no new numbers. |

`REQ-{counter}` with 5 digits and a last number of 0 shows **REQ-00001**. A formula typed as `{counter:4}{counter}` is saved as `{counter:4}`, so the digit is written once.

Date tokens use the item’s Created time in UTC. `{Department.title}` uses the lookup or person display text.

## What the numbering service does

Once a minute it reads the active rows in `AutoGenFlowConfiguration`, skips the site row, and for each list loads blank number columns in Id order. It assigns the next integer, writes the code, then stores the new count. Up to 100 blank items on a list are numbered in one pass. The rest continue on the next minute.

Turn off stops the service and leaves the saved rules in place. Redeploy replaces the service after an app upgrade. Adding another list is only a new rule. The same service picks it up.

## Local development

From `spfx/AutoGenFlow`, with Node `>=22.14.0 <23`:

```bash
npm install
npm test
npm run build
```

`npm run build` runs the tests and writes `sharepoint/solution/autogen-flow.sppkg`. Copy that file to `deploy/` after a release build.
