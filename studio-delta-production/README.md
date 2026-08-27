# Studio Delta Production

Shop-floor + office app for Studio Delta (South Africa). The UI is the same `index.html` as before. **The live database is on Railway** (a workbook file on the server volume). Google Sheets is only used to **copy data in once**, and for Drive QC PDFs / emails.

Timezone: **Africa/Johannesburg**. The paid shift is **07:45–15:45** with a **30-minute break 12:00–12:30** (7.5 hours / 450 minutes). Minutes after that on a calendar day are overtime. Overnight jobs are split by calendar day on Activity and on Admin → Workers (e.g. 14:00 yesterday–10:00 today → yesterday 14:00–15:45 and today 07:45–10:00).

## Host on Railway (Railway is the database)

The GitHub repo is a website plus this app. Railway should build from the **repository root** `Dockerfile`, which copies `studio-delta-production/` only.

Floor start / pause / resume / finish, durations, steel usage, backboard usage, idle alerts, and order **STATUS** / **ASSIGNED OPERATOR** all save in Railway. The office **Orders** page reads the **same ORDERS table**, so a tablet clock updates the status the office sees.

1. In [Google Cloud](https://console.cloud.google.com/), create a project (or pick one) → **APIs & Services** → enable **Google Sheets API**, **Google Drive API**, **Google Docs API**. Enable **Gmail API** only if you want QC / powder-list / glass-alert emails.
2. **IAM** → **Service accounts** → create one (e.g. `studio-delta-floor`) → **Keys** → JSON. Copy the JSON.
3. Open the production spreadsheet → **Share** → add the service account email (`...@....iam.gserviceaccount.com`) as **Viewer** (import is read-only) or **Editor** if you still use Drive QC.
4. Share these Drive folders with the same email as **Editor** (same IDs already in `Code.gs`):
   - QC reports folder
   - PDF job-queue folder
   - QC Google Doc templates (the service account must be able to copy them)
   - Powder-coating list folder (or let the first generate create one in the service account’s Drive)
5. In [Railway](https://railway.com/) → **New project** → **Deploy from GitHub** → this repo.
6. Variables (service → **Variables**):

   | Variable | Value |
   | --- | --- |
   | `SHEET_ID` | `1pdvAFTIyd5sf8Wbf38MSd4cfk3mb3McPqJrYeM8SOYk` (or your copy) — used for one-time import |
   | `TZ` | `Africa/Johannesburg` |
   | `GOOGLE_SERVICE_ACCOUNT_JSON` | the **full JSON key** as one line (import + Drive QC) |
   | `GMAIL_SENDER` | optional. A Workspace mailbox the service account can send as (domain-wide delegation). If unset, the floor still works; QC PDFs / powder emails / glass alerts are skipped or fail softly. |
   | `DATA_DIR` | `/app/data` (already set in Docker) |

7. **Settings → Networking → Generate domain**. Tablets open that HTTPS URL. Login is still name + access code from the `Users` table (copied from Google on first boot if the Railway volume is empty).
8. Keep **one replica**. The in-memory lock/cache is per process.
9. Add a **Volume** mounted at **`/app/data`**. Without it, shop data is lost on every deploy. That folder holds `floor-workbook.json` (Users, ORDERS, Production_Log, steel, backboards, idle, schedule) plus office extras (dropdowns, payments).

`GET /health` should return `"ok": true` and `"db": "railway"`. The first boot copies Google into Railway if Users and ORDERS are still empty. After that, Google is not the live store.

On **Orders**, **Import from Sheets** copies Users, orders, production logs, steel and backboards from Google **into Railway** (this replaces current Railway shop data). Use it once when you first move, or if you need a fresh copy.

QC PDFs still fill worker / order / Yes-No answers. Inserting photos into the Google Doc template is not wired on Railway (text tags for missing photos stay as-is). Floor start / pause / finish does not need Drive.

Local run (from `studio-delta-production/`):

```bash
cp .env.example .env   # then fill credentials if you want Google import
npm install
npm test
npm start              # http://localhost:8080
```

Do **not** commit the JSON key or a `.env` file.

## Office pages (same Railway database as the floor)

- `/orders` — all ORDERS columns (quote, client, prices, etc.). Status and assigned operator are the same fields the floor writes.
- `/schedule` — order list + week grid (Mon–Fri)
- `/dropdowns` — lists for Type, Category, Product, and the other order dropdowns
- `/debtors` — orders not fully paid

## Floor rules

- One person has **one running clock**. Starting or resuming another order asks **Switch from A to B** (pause A, start B — still needs a reason) or **Work on A and B** (one clock, time split; Together badge). **Work this order only** leaves a batch. Pausing or switching requires **No materials**, **Touch up** (plus the order number), or **Other**. Activity and the Workers hour log split those hours by calendar day and show each bout with the pause reason — not a single start-minus-end total.
- Same product, same process (e.g. assembly slats for two matching gates) can still be started together from Available. Use **Work this order only** to leave a batch.
- When plate cutting finishes on an order, welders waiting on that plate are **auto-switched** back. They can tap **Still on other job** if they are not ready.
- **Ready for Assembly** can go to **Assembly** or **Paint Preparation**. Assemblers see **Assemble** and **Paint prep**. Painters (Users task `Painting`) also see Ready for Assembly and can **Start paint prep**, then paint. Paint prep finish → **Ready for Painting**. Painting finish → **Ready for Assembly** so assembly can happen.
- If nobody has a running job for more than ~10 minutes during shift hours (07:45–15:45 weekdays), **Admin and QC who currently have the app open** get a popup modal listing those idle people. Assign an **indirect task** from the popup (or from Admin → Workers). Starting a real order closes that task. Idle alerts are **not** emailed.

## Railway workbook tabs

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
