# Studio Delta Production

Shop-floor app for Studio Delta (South Africa). The UI is the same `index.html` as before. **Google Sheets is still the database.** The live host is meant to be **Railway** (Node/Express). Apps Script paste deploy still works if you need a fallback: paste `Code.gs` and `index.html` into the bound script project (HTML file **must** be named `index` there).

Timezone: **Africa/Johannesburg**. The paid shift is **07:45–15:45** with a **30-minute break 12:00–12:30** (7.5 hours / 450 minutes). Minutes after that on a calendar day are overtime. Overnight jobs are split by calendar day on Activity and on Admin → Workers (e.g. 14:00 yesterday–10:00 today → yesterday 14:00–15:45 and today 07:45–10:00).

## Host on Railway (Google Sheets stays the DB)

The GitHub repo is a website plus this app. Railway should build from the **repository root** `Dockerfile`, which copies `studio-delta-production/` only.

1. In [Google Cloud](https://console.cloud.google.com/), create a project (or pick one) → **APIs & Services** → enable **Google Sheets API**, **Google Drive API**, **Google Docs API**. Enable **Gmail API** only if you want QC / powder-list / glass-alert emails.
2. **IAM** → **Service accounts** → create one (e.g. `studio-delta-floor`) → **Keys** → JSON. Copy the JSON.
3. Open the production spreadsheet → **Share** → add the service account email (`...@....iam.gserviceaccount.com`) as **Editor**.
4. Share these Drive folders with the same email as **Editor** (same IDs already in `Code.gs`):
   - QC reports folder
   - PDF job-queue folder
   - QC Google Doc templates (the service account must be able to copy them)
   - Powder-coating list folder (or let the first generate create one in the service account’s Drive)
5. In [Railway](https://railway.com/) → **New project** → **Deploy from GitHub** → this repo.
6. Variables (service → **Variables**):

   | Variable | Value |
   | --- | --- |
   | `SHEET_ID` | `1pdvAFTIyd5sf8Wbf38MSd4cfk3mb3McPqJrYeM8SOYk` (or your copy) |
   | `TZ` | `Africa/Johannesburg` |
   | `GOOGLE_SERVICE_ACCOUNT_JSON` | the **full JSON key** as one line |
   | `GMAIL_SENDER` | optional. A Workspace mailbox the service account can send as (domain-wide delegation). If unset, the floor still works; QC PDFs / powder emails / glass alerts are skipped or fail softly. |

7. **Settings → Networking → Generate domain**. Tablets open that HTTPS URL. Login is still name + access code from the `Users` sheet.
8. Keep **one replica**. The in-memory lock/cache is per process.

`GET /health` should return `"ok": true` and `"sheetsConfigured": true` after the variables are set. The first request loads the whole spreadsheet into memory, runs the same shop functions as Apps Script, then writes changes back.

QC PDFs still fill worker / order / Yes-No answers. Inserting photos into the Google Doc template is not wired on Railway (text tags for missing photos stay as-is). Floor start / pause / finish does not need Drive.

Local run (from `studio-delta-production/`):

```bash
cp .env.example .env   # then fill credentials
npm install
npm test
npm start              # http://localhost:8080
```

Do **not** commit the JSON key or a `.env` file.

## Office pages (not Google Sheets)

The live floor still uses the spreadsheet until we finish moving it. Office **Orders** and **Schedule** already save in the app database (a file on the server).

- `/orders` — all ORDERS columns (quote, client, prices, etc.)
- `/schedule` — order list + week grid (Mon–Fri)

On Orders, **Import from Sheets** copies the ORDERS tab once into the app. After that, edit in the app.

On Railway, add a **Volume** mounted at `/app/data` so orders survive deploys. Or we can switch this file to Postgres next.

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

## Apps Script fallback (optional)

Paste `Code.gs` and `index.html` into the bound Apps Script project. The HTML file **must** be named `index` in the script editor.

1. Replace the existing `Code.gs` and `index` HTML with these files.
2. Open the web app once as admin. `doGet` tries to create a 5-minute idle check trigger.
3. Confirm **Project Settings → Time zone** is `Africa/Johannesburg`.

## Activity report

Admin sidebar → **Activity**. Daily / weekly / monthly view of what each person did, with regular hours capped at 7.5h per day and overtime shown separately. Overnight jobs and pause/resume bouts are split by calendar day, and each pause shows **when** and **why**. Admin → **Workers** uses the same per-day split for the hour log (one row per calendar day) and weekly totals.
