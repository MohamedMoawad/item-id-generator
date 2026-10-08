# AutoGenFeatureConfiguration

One row per list or library that should receive an automatic request number. Create the list on the **site collection root**. The configuration web part button does this, and `provision-numbering-config.ps1` is the PnP alternative.

The Azure Function owns `CurrentCount`, `LastResetDate`, and `WebhookSubscriptionId`. An admin edit must not write the first two.

| Column | Type | Who writes it |
| --- | --- | --- |
| `Title` | Text | Admin. Friendly name. |
| `TargetListUrl` | Text, required | Admin. https URL of the list or library. `/AllItems.aspx` and `/Forms/AllItems.aspx` are stripped. A path containing `/lists/{name}` uses the site before that segment. Anything else uses the parent of the last segment. |
| `TargetListGuid` | Text, required | Admin, after `GetList`. One active row per GUID. |
| `NumberColumnInternalName` | Text, required | Admin. Single-line text column on the target list, such as `RequestNumber`. This repo does not create that column. |
| `Formula` | Text, required | Admin. Must contain `{seq}` or `{seq:n}`. |
| `CurrentCount` | Number | Function. New rows start at 0. |
| `ResetPeriod` | Choice `None`, `Day`, `Month`, `Year` | Admin. `None` never resets. |
| `LastResetDate` | DateTime | Function. Empty does not reset an existing count. The first issued number stamps the date. Later issues reset only when the UTC day, month, or year key changed. |
| `IsActive` | Boolean | Admin. Missing is inactive. |
| `PadLength` | Number 0–12, default 4 | Admin. Width for `{seq}`. `{seq:n}` still wins. |
| `WebhookSubscriptionId` | Text | `RegisterWebhook`. Cleared when the target GUID changes so the next save can subscribe again. The previous subscription is left in place until it expires. |

UTC tokens: `{yyyy}` `{yy}` `{MM}` `{dd}` `{HH}` `{mm}` `{seq}` `{seq:n}` with `n` from 1 to 12.

`REQ-{yyyy}{MM}-{seq}` with `PadLength` 4 at 2026-10-06 UTC and sequence 12 becomes `REQ-202610-0012`.

`CreateFieldAsXml` uses Options **25** (add to the default content type, honor the internal name, add to the default view).
