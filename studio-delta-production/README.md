# Studio Delta Production

Google Apps Script shop-floor app for Studio Delta (South Africa). Paste `Code.gs` and `index.html` into the Apps Script project bound to the production spreadsheet. The HTML file **must** be named `index` in the script editor.

Timezone: **Africa/Johannesburg**. A standard work day is **8 hours (480 minutes)**. Anything above that on a calendar day is overtime.

## Floor rules

- One person has **one running clock**. Extra open jobs are paused automatically.
- Same product, same process (e.g. assembly slats for two matching gates) can be **batched**: one clock, time split across the orders. Use **Work this order only** to leave the batch.
- When plate cutting finishes on an order, welders waiting on that plate are **auto-switched** back. They can tap **Still on other job** if they are not ready.
- If nobody has a running job for more than ~10 minutes during shift hours (07:00–17:00 weekdays), admin gets an email. Admin can assign an **indirect task**. Starting a real order closes that task.

## New spreadsheet pieces

| Sheet / column | Purpose |
| --- | --- |
| `Production_Log` column M | JSON meta: pause intervals, batch id, share, entry type |
| `Idle_Alerts` | One row per idle email so the same person is not emailed twice in a day |

Legacy pause columns J/K/L are still written for older reports.

## First deploy

1. Replace the existing `Code.gs` and `index` HTML with these files.
2. Open the web app once as admin. `doGet` tries to create a 5-minute idle check trigger.
3. Confirm **Project Settings → Time zone** is `Africa/Johannesburg`.

## Activity report

Admin sidebar → **Activity**. Daily / weekly / monthly view of what each person did, with regular hours capped at 8h per day and overtime shown separately.
