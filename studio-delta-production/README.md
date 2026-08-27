# Studio Delta Production

Google Apps Script shop-floor app for Studio Delta (South Africa). Paste `Code.gs` and `index.html` into the Apps Script project bound to the production spreadsheet. The HTML file **must** be named `index` in the script editor.

Timezone: **Africa/Johannesburg**. The paid shift is **07:45–15:45** with a **30-minute break 12:00–12:30** (7.5 hours / 450 minutes). Minutes after that on a calendar day are overtime. Overnight jobs are split by calendar day on Activity and on Admin → Workers (e.g. 14:00 yesterday–10:00 today → yesterday 14:00–15:45 and today 07:45–10:00).

## Floor rules

- One person has **one running clock**. Starting or resuming another order asks **Switch from A to B** (pause A, start B — still needs a reason) or **Work on A and B** (one clock, time split; Together badge). **Work this order only** leaves a batch. Pausing or switching requires **No materials**, **Touch up** (plus the order number), or **Other**. Activity and the Workers hour log split those hours by calendar day and show each bout with the pause reason — not a single start-minus-end total.
- Same product, same process (e.g. assembly slats for two matching gates) can still be started together from Available. Use **Work this order only** to leave a batch.
- When plate cutting finishes on an order, welders waiting on that plate are **auto-switched** back. They can tap **Still on other job** if they are not ready.
- **Ready for Assembly** can go to **Assembly** or **Paint Preparation**. Assemblers see **Assemble** and **Paint prep**. Painters (Users task `Painting`) also see Ready for Assembly and can **Start paint prep**, then paint. Paint prep finish → **Ready for Painting**. Painting finish → **Ready for Assembly** so assembly can happen.
- If nobody has a running job for more than ~10 minutes during shift hours (07:45–15:45 weekdays), **Admin and QC who currently have the app open** get a popup modal listing those idle people. Assign an **indirect task** from the popup (or from Admin → Workers). Starting a real order closes that task. Idle alerts are **not** emailed.

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
- Painters need `Painting` on the Tasks column. That also lets them prep items for painting (no extra `Paint Preparation` task required). Assemblers with `Assembly` also get **Paint prep** on Ready for Assembly cards.

## Worker schedule

Admin sidebar → **Schedule**. Click Not Yet Started (and Ready for Steelwork) orders in the sequence they should be done, pick a process and a worker, then **Build calendar**. Block lengths come from past jobs of the same product and process (or a typical process time if that product is new). **Insert other task** (cleaning, meeting, etc.) pins a block on that worker’s calendar and pushes later work through 07:45–15:45, skipping lunch and weekends. This is a plan, not a live clock — the floor still starts/pauses/finishes as usual.
- If Tasks is blank, the app reads the Role cell (`Welder Tagger` → Welding + Tagging).

The tablet does **not** remember the last login. After Log Out (or a refresh) everyone sees the login screen again.

Floor boards read only the latest production-log rows (not the whole history) and cache the order list for about 90 seconds. Start/pause/finish reuse that same log slice instead of rereading the sheet several times. Auto-refresh is once a minute while the tab is visible.

Assembly and Final QC must log **which backboard was used**, the same way plate/profile cutting logs steel. Add backboard names on the `Backboards` sheet (Category, Profile Name), or pick Custom during QC.

## First deploy

1. Replace the existing `Code.gs` and `index` HTML with these files.
2. Open the web app once as admin. `doGet` tries to create a 5-minute idle check trigger.
3. Confirm **Project Settings → Time zone** is `Africa/Johannesburg`.

## Activity report

Admin sidebar → **Activity**. Daily / weekly / monthly view of what each person did, with regular hours capped at 7.5h per day and overtime shown separately. Overnight jobs and pause/resume bouts are split by calendar day, and each pause shows **when** and **why**. Admin → **Workers** uses the same per-day split for the hour log (one row per calendar day) and weekly totals.
