# Studio Delta Production

Google Apps Script shop-floor app for Studio Delta (South Africa). Paste `Code.gs` and `index.html` into the Apps Script project bound to the production spreadsheet. The HTML file **must** be named `index` in the script editor.

Timezone: **Africa/Johannesburg**. A standard work day is **8 hours (480 minutes)**. Anything above that on a calendar day is overtime.

## Floor rules

- One person has **one running clock**. Extra open jobs are paused automatically.
- Same product, same process (e.g. assembly slats for two matching gates) can be **batched**: one clock, time split across the orders. Use **Work this order only** to leave the batch.
- When plate cutting finishes on an order, welders waiting on that plate are **auto-switched** back. They can tap **Still on other job** if they are not ready.
- If nobody has a running job for more than ~10 minutes during shift hours (07:00–17:00 weekdays), **Admin and QC who currently have the app open** get a popup modal listing those idle people. Assign an **indirect task** from the popup (or from Admin → Workers). Starting a real order closes that task. Idle alerts are **not** emailed.

## New spreadsheet pieces

| Sheet / column | Purpose |
| --- | --- |
| `Production_Log` column M | JSON meta: pause intervals, batch id, share, entry type |
| `Idle_Alerts` | Open idle people for the day. Admin/QC popups read this list; assigning a task sets Status to Assigned |
| `Backboards` | Category + profile name list for assembly / final QC (same idea as `Steel_Profiles`) |
| `Backboard_Usage` | Timestamp, order, worker, process, type, size logged when assembly or final QC is finished |

Legacy pause columns J/K/L are still written for older reports.

## Users sheet (login)

Workers log in **once per visit** with name + access code. Every page load shows the login screen first (the last login is not remembered).

Add a **Tasks** column (column D). The first time the app opens it will create the header if it is missing.

| Name | Role | Password | Tasks |
| --- | --- | --- | --- |
| Sipho | Welder Tagger | 1234 | Welding, Tagging |
| Thabo | Quality Control | 1234 | Quality Control |
| Admin | Admin | **** | |

- **Admin** sees Production, Workers, Metrics, QC Reports, Activity, and can open any floor task.
- **Quality Control / QC** only sees QC work.
- Anyone else only sees the tasks listed. `Welding, Tagging` (or a role like `Welder Tagger`) means they pick Welding or Tagging after login, and only those boards appear.
- If Tasks is blank, the app reads the Role cell (`Welder Tagger` → Welding + Tagging).

The tablet does **not** remember the last login. After Log Out (or a refresh) everyone sees the login screen again.

Floor screens cache order lists for about 20 seconds and only reread the spreadsheet after a start, pause, finish, or similar change. Auto-refresh is once a minute while the tab is visible.

Assembly and Final QC must log **which backboard was used**, the same way plate/profile cutting logs steel. Add backboard names on the `Backboards` sheet (Category, Profile Name), or pick Custom during QC.

## First deploy

1. Replace the existing `Code.gs` and `index` HTML with these files.
2. Open the web app once as admin. `doGet` tries to create a 5-minute idle check trigger.
3. Confirm **Project Settings → Time zone** is `Africa/Johannesburg`.

## Activity report

Admin sidebar → **Activity**. Daily / weekly / monthly view of what each person did, with regular hours capped at 8h per day and overtime shown separately.
