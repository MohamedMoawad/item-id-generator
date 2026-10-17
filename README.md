# Request numbers for SharePoint Online

When someone creates a list item from an open list, this SharePoint app assigns the next request number. Upload the package, open the list, choose **Autogen Feature**, and save the formula.

The number is written by the list page. There is no Azure Function.

## Upload

The package is [deploy/item-id-generator.sppkg](deploy/item-id-generator.sppkg). Steps are in [deploy/README.md](deploy/README.md).

## How a number is chosen

The open list reads `AutoGenFeatureConfiguration`, reserves the next count with the row etag, and puts that code into the new-item save. Two people saving at the same time get different numbers because a lost etag is retried. An item that already has a number is left alone.

An item created while nobody has the list open stays blank until the next person opens that list. The page then fills the blank number column.

## Project

| Path | Role |
| --- | --- |
| `spfx/ItemIdGenerator` | The SharePoint Framework app: settings dialog, config list, and numbering on the open list. |
| `spfx/AutoGenFlow` | Sequential numbers for every list. One site numbering flow writes the code. See [spfx/AutoGenFlow/README.md](spfx/AutoGenFlow/README.md). |
| `spfx/RepeatingRows` | Repeating line items stored in a child list. See [spfx/RepeatingRows/README.md](spfx/RepeatingRows/README.md). |

## Versions

| Piece | Version |
| --- | --- |
| Node.js | `>=22.14.0 <23` |
| SPFx | 1.23.2, Heft, React 17 |

## Local development

From `spfx/ItemIdGenerator`:

```bash
npm install
npm test
npm start
```

`npm start` is the old `gulp serve`. `npm run build` runs the tests and packages `sharepoint/solution/item-id-generator.sppkg`. Copy that file to `deploy/` after a release build.

## AutoGen Flow

[spfx/AutoGenFlow](spfx/AutoGenFlow/README.md) is a second numbering app. Open a list, choose **AutoGen Flow**, and add, edit, or delete that list’s rule. Choose the environment and connection user, then save. The app creates the numbering flow. Upload [deploy/autogen-flow.sppkg](deploy/autogen-flow.sppkg) separately. On a list numbered by AutoGen Flow, clear **Active** on that list’s Autogen Feature rule.
