# AutoGen Flow

AutoGen Flow stores the numbering rule in the master list `AutoGenFlowConfiguration`. Power Automate writes the number when an item is created.

The existing **AutoGen Feature** app is unchanged. It still numbers items from the open list page. Install this package beside it. On a list that Power Automate creates, save the rule here and leave the AutoGen Feature rule for that list inactive.

This package does not hook the browser, and it does not number items by itself.

## Upload

The package is [deploy/autogen-flow.sppkg](../../deploy/autogen-flow.sppkg), solution **AutoGen Flow**, version **1.0.0.0**.

1. Upload it in the SharePoint app catalog and deploy it.
2. On the site collection, open **Site contents → Add an app → From your organization** and add **AutoGen Flow**.
3. Open the list. On the command bar choose **AutoGen Flow**.
4. Choose the single-line text column, set the formula, and save. The first save creates `AutoGenFlowConfiguration` on the site collection if it is not there yet.
5. Build the flow below. One flow handles one list. Clone it for the next list.

`AutoGenFlowConfiguration` stays visible in site contents. Each list has one row. The row holds the formula, the number column, the reset period, the current count, and the column tokens the flow replaces.

## What the master list stores

| Column | How the flow uses it |
| --- | --- |
| TargetListGuid | Filter query finds the row for this list. |
| NumberColumnInternalName | The text column the flow updates. |
| Formula | The pattern, with `{counter}` written once. |
| CounterToken | The exact token to replace, such as `{counter:4}`. |
| CounterWidth | Digits to pad. `{counter:4}` stores 4. A plain `{counter}` uses Pad length. |
| FormulaColumns | Comma-separated tokens such as `Title,Department.title`. |
| CurrentCount | The last number used. The next item adds one. |
| ResetPeriod | `None`, `Day`, `Month`, or `Year`. |
| LastResetDate | The flow sets this to the item Created time after a successful number. |
| IsActive | The filter keeps inactive rows out of the flow. |

Tokens:

| Token | Value |
| --- | --- |
| `{yyyy}` `{yy}` `{MM}` `{dd}` `{HH}` `{mm}` | Item Created time, UTC |
| `{counter}` | Next integer, padded with Pad length |
| `{counter:4}` | Next integer, padded to that width |
| `{Title}` | Text column from the created item |
| `{Department.title}` | Lookup or person display text |

`{counter:4}{counter}` is saved as `{counter:4}` so the digit is written once. A current count of 0 produces 1.

## Build the flow

Create an automated cloud flow.

1. Trigger: **When an item is created** (SharePoint). Site and list are the list you just configured.
2. Open the trigger settings.
   - Concurrency control: **On**
   - Degree of parallelism: **1**
   - Trigger condition, pasted from the AutoGen Flow dialog:
     `@empty(triggerOutputs()?['body/RequestNumber'])`
     Replace `RequestNumber` with the number column internal name.
3. **Get item**. Same site and list. Id is the trigger **ID**.
4. **Get items**.
   - Site: the master list site shown in the dialog (the site collection).
   - List: `AutoGenFlowConfiguration`
   - Filter query from the dialog, for example `TargetListGuid eq 'guid' and IsActive eq 1`
   - Top Count: `1`
5. **Initialize variable**. Name `Code`, type String, value empty. This action stays outside the condition.
6. **Condition**. Paste the condition expression from the dialog. It continues only when the fresh item has an empty number column and a master row exists. Put the actions below in **Yes**.
7. Add a **Compose** for each name. Rename the action to that name, then paste the expression from the dialog into the expression editor.

| Compose name | Role |
| --- | --- |
| Config | First master row |
| ResetPeriod | None, Day, Month, or Year |
| LastResetDate | Previous numbered item |
| CurrentCount | Integer last value, or 0 |
| CounterWidth | Pad width |
| CounterToken | `{counter}` or `{counter:4}` |
| FormulaColumns | Column tokens, or empty |
| Rolled | True when the reset period changed |
| Next | 1 after a reset, otherwise current count plus 1 |
| Padded | Next, padded with twelve zeros to CounterWidth |
| CodeStart | Formula with the counter and the UTC date tokens replaced |

8. **Set variable** `Code` to the output of `CodeStart`.
9. When FormulaColumns is not empty, add **Apply to each**.
   - Select: `split(outputs('FormulaColumns'), ',')`
   - Inside the loop, **Set variable** `Code` to the column expression from the dialog.
   - The expression reads `Value` or `Title` for a lookup or person, and the plain value for text.
10. **Update item** on the created item. Set the number column to `variables('Code')`.
11. **Update item** on `AutoGenFlowConfiguration`.
    - Id: `outputs('Config')?['ID']` (if that is empty, use `Id`)
    - Current count: output of `Next`
    - Last reset date: **Created** from the trigger

The counter moves only after the item update succeeds. Concurrency 1 keeps one hundred creates in one line: 0001, 0002, 0003, with no skipped values.

If the item update succeeds and the master-list update fails, open that master row and set **Current count** to the number just used. The next run then continues from there.

## One flow for each list

The SharePoint trigger is bound to one list. For a second list, save its AutoGen Flow rule, clone the flow, and change three things: the trigger list, the trigger condition column, and the filter query. The master list can live on the site collection while the target list is in a subsite.

## Local development

From `spfx/AutoGenFlow`, with Node `>=22.14.0 <23`:

```bash
npm install
npm test
npm run build
```

`npm run build` runs the tests and writes `sharepoint/solution/autogen-flow.sppkg`. Copy that file to `deploy/` after a release build.
