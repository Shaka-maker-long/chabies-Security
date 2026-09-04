# Studio Delta Production

Shop-floor + office app for Studio Delta (South Africa). One login: **Admin** sees office and shop pages; **Production** only sees the floor. The shell is a compact SAP-style ERP in Studio Delta steel, linen, and brass. After login, **office users always open Home** — clicking Home in the same tab reuses that office session and must not ask for the access code again. Home is a **live 3D digital twin** of the shop. Drag people and piles to place them. People **animate** the job they are on (welding and tagging share a weld arc; idle people stand still). Count badges open the order list. **Ready for final QC** sits in finished goods. **Production** in the menu is the list-report. After you enter name and access code, a 5-second **Welcome to Studio Delta** screen draws the S mark, then the app opens. Quote values are typed **including VAT (15%)**; exclusive VAT is saved too, and a running total updates as you type so you can check it against the quote PDF.

Orders (~400+) load as a text grid first; click a row to edit. Floor reads no longer rewrite the whole workbook to disk.

## Database: SQLite on the Railway volume (not Google Sheets, not Postgres)

The live store is **SQLite** at `DATA_DIR/studio-delta.db` on the Railway volume. Every page’s data is in SQL tables. Google Sheets is **not** read or written while the app runs. The old JSON files stay as a backup copy. First boot copies the existing files into SQLite so the shop is not empty.

| Page | SQLite tables |
| --- | --- |
| Users / login | `users`, `sessions` |
| Orders / debtors | `orders`, `payments`, `sheet_rows` (ORDERS tab) |
| Enquiries / Tasks | `enquiries` |
| Dropdowns | `dropdowns` (order lists and `enquiry:` lists) |
| Office schedule | `office_schedule_rows`, `office_schedule_cells` |
| Floor / Production / Workers / Metrics / QC / Activity | `sheet_rows` for `Production_Log`, `Overview`, `Idle_Alerts`, `Rates` |
| Steel / backboards | `sheet_rows` for `Steel_Profiles`, `Steel_Usage`, `Backboards`, `Backboard_Usage` |
| Worker schedule / task times | `sheet_rows` for `Schedule`, `Task_Durations` |

| File | What it holds |
| --- | --- |
| `DATA_DIR/studio-delta.db` | All page tables above, plus JSON blobs as a second copy |
| `DATA_DIR/floor-workbook.json` | Backup of every workbook tab |
| `DATA_DIR/studio-delta.json` | Backup of enquiries, dropdowns, payments, office schedule |
| `DATA_DIR/office-sessions.json` | Backup of office login sessions |
| `DATA_DIR/enquiry-quotes/` | Quote PDFs |
| `DATA_DIR/enquiry-files/` | Cost sheets, follow-up screenshots, POP, drawings, Outlook emails |

On Railway, `DATA_DIR` is `/app/data`. **A Volume must be mounted at `/app/data`**. Without it, every deploy wipes the database. Users → **Download backup** saves a JSON copy of shop + office data.

### Automatic backups (volume + Google Drive)

Every night at **02:00 Africa/Johannesburg** (and two minutes after boot if today’s copy is missing) the app writes a dated snapshot into `/app/data/backups/` and keeps **14 days**. That protects against a bad save or a deleted enquiry. Users → **Backup now** runs the same job immediately.

Those local copies are still on the same Railway disk. **Off-site** is Google Drive (a different company than Railway):

1. Keep `GOOGLE_SERVICE_ACCOUNT_JSON` on the Railway service (the same key used for QC PDFs).
2. In Google Drive, create a folder e.g. **Studio Delta ERP backups**.
3. Share that folder with the service account email (`client_email` in the JSON key) as **Editor**.
4. Set `BACKUP_DRIVE_FOLDER_ID` to the folder ID from the Drive URL.
5. Optional: `BACKUP_EMAIL` (or `GMAIL_SENDER`) so a “backup OK” mail goes out, with the `.db` attached when it is small.

If Drive is not set, nightly copies still run on the volume only. `/health` shows `backupOk`, `backupOffsite`, and `backupAt`. Restore: download a dated `.db` from Users or from Drive, then replace `/app/data/studio-delta.db` on the volume (stop the service, replace the file, start it).

`GET /health` must show `"sheetsLive": false` and `"usingEphemeralDisk": false`. Office pages show a red banner if the volume is missing.

### One-time copy from the old spreadsheet

Only if you still need the Google workbook on Railway:

1. Set `GOOGLE_MIGRATE=1`, `SHEET_ID`, and `GOOGLE_SERVICE_ACCOUNT_JSON`.
2. Open **Users** → **Copy old Google spreadsheet once** (Users, ORDERS, logs, and an Enquiries tab if it exists).
3. **Remove `GOOGLE_MIGRATE`** so Google cannot overwrite Railway again.

If Railway Users is empty on boot (for example after a deploy with no volume), the app copies Users from the old spreadsheet once so people can log in. That is not the live database. **Attach a volume at `/app/data`** or the next deploy will wipe logins again.

Timezone: **Africa/Johannesburg**. The paid shift is **07:45–15:45** with a **30-minute break 12:00–12:30** (7.5 hours / 450 minutes). Minutes after that on a calendar day are overtime. Overnight jobs are split by calendar day on Activity and on Admin → Workers (e.g. 14:00 yesterday–10:00 today → yesterday 14:00–15:45 and today 07:45–10:00).

## Host on Railway

The GitHub repo is a website plus this app. Railway should build from the **repository root** `Dockerfile`, which copies `studio-delta-production/` only.

1. In [Railway](https://railway.com/) → **New project** → **Deploy from GitHub** → this repo.
2. Variables:

   | Variable | Value |
   | --- | --- |
   | `TZ` | `Africa/Johannesburg` |
   | `DATA_DIR` | `/app/data` (already set in Docker) |
   | `GOOGLE_SERVICE_ACCOUNT_JSON` | optional for QC PDFs. **Required for off-site Drive backups** |
   | `GMAIL_SENDER` | optional. Workspace mailbox for QC / powder / glass emails and backup notices |
   | `BACKUP_DRIVE_FOLDER_ID` | optional. Google Drive folder ID for nightly off-site copies |
   | `BACKUP_EMAIL` | optional. Address that receives “backup OK / failed” mail |

3. **Volume (required):** service → **Volumes** → mount path **`/app/data`**. Confirm `GET /health` shows `"usingEphemeralDisk": false`.
4. **Settings → Networking → Generate domain**. Login is name + access code from the Railway `Users` table. If Users is empty, the app seeds **Admin** / **admin** — change that code on Users.
5. Keep **one replica**.

Google Drive QC is optional. If you still generate QC PDFs: enable Sheets/Drive/Docs APIs, share the Drive folders with the service account (same IDs as `Code.gs`). Floor start / pause / finish does not need Google.

QC PDFs still fill worker / order / Yes-No answers. Inserting photos into the Google Doc template is not wired on Railway (text tags for missing photos stay as-is). Floor start / pause / finish does not need Drive.

Local run (from `studio-delta-production/`):

```bash
cp .env.example .env   # Google credentials are optional (QC PDFs only)
npm install
npm test
npm start              # http://localhost:8080
```

Do **not** commit the JSON key or a `.env` file.

## Office pages (same Railway database as the floor)

- `/orders` — all ORDERS columns (quote, client, prices, etc.). Status and assigned operator are the same fields the floor writes.
- `/enquiries` — Enquiry log (view-only sheet) plus **Dashboard** at `/enquiries/dashboard`. The dashboard is a weekly or monthly line of enquiry and quote **counts**, switchable to **revenue**. Tick two or more months on **Week of month** to overlay week 1–5 (first week vs later weeks). A **Month** filter applies to every chart. Click Custom or New Design for a split pie (Dimensions / Colour / Other, or each description). Quote value by type replaces the old Outlook pie. Product shows the top items plus a search list. Province can switch to quote value. Click a line-graph dot (or a bar) to list the enquiries, then a row for the full enquiry popup. On the sheet, numbers start at `#1996` and count up. Click a row to **view** the enquiry in a read-only popup. **New enquiry** and each row’s pencil icon open the edit popup — you do not type in the cells. The list icon next to the pencil opens status, costing, and files. **STATUS** is text only. **MONTH ENQUIRED** fills from **DATE ENQUIRED**. **QUOTE NO.** shows the latest quotation number; **DATE QUOTED** is the day the quote PDF was uploaded. Columns through **CLIENT NAME** stay frozen while you scroll. Capture product names here. **Custom** asks whether the change is **Dimensions**, **Colour**, or **Other** (Other names are saved onto the dropdown). **New Design** needs a full description. **STATUS is not a free dropdown** — it moves when the assigned office person saves a deliverable (cost sheet, approval, quote PDF, follow-up screenshot, POP, drawing). If a quoted client wants changes, stay on that enquiry: **Issue another quote** (new SOQ, previous PDFs stay in Files) or **Client wants changes — recost**. Lifespan stays in the status popup, not on the sheet. Costing uses the **Costing / Quoting / Approval** people ticked on Users, and you can still pick someone else on that enquiry. Open the list icon on any enquiry to see **Files**: Outlook emails, cost sheets, quote PDF, follow-up screenshots, proof of payment, and drawing stay listed and openable at any time. **CORRESPONDANCE**: paste any **file link or path** (including those with `//`). It is saved on the server as **Correspondance link** (do not attach a file). Files shows that name plus **Copy link** / **Open**. Cost sheets upload **per product**, and you can add more than one sheet per item; earlier sheets stay in Files on recost. Filter the sheet by **day, week, or month** (from DATE ENQUIRED) and tick **Quoted only**. The **quote total** at the top (incl and excl VAT) follows those filters. Product **values, delivery excl VAT, and quotation number** are entered when the quote PDF is uploaded. The quotation number defaults to the next `SOQ` after the last used number (enquiries and orders) and cannot be a duplicate.
- `/tasks` — **My tasks**. Office Admins see open work assigned to them and update the system from that queue (preview the file, confirm it, save). Finished work moves to **Completed** (`/tasks/completed`). Follow-ups due 7 days after the quote or last follow-up show as overdue on the to-do list.
- `/schedule` — order list + week grid (Mon–Fri)
- `/dropdowns` — lists for Type, Category, Product, and the other order dropdowns
- `/users` — office and floor people, Railway backup download. Google Sheets is not used. One person with job title **Manager** can add, delete, and set Access (Admin vs Production). The Manager sets someone else’s access code on Users (**Access code**, then Save) or in **Change access code** by typing that person’s name — the old code must stop working. Until someone is Manager, the first office Admin can still open Users. Everyone else only changes their own access code (**Change access code** in the menu, or on Home after login). After login, the sidebar shows **Logged in as** with the person’s name and job title. On Admin people, tick **Costing**, **Quoting**, and **Approval** so those names fill in automatically on new enquiries. You can still pick someone else on a single enquiry. Leave Approval unticked to skip cost-sheet approval.
- `/outlook-addin` — optional install page for Outlook with Get Add-ins. Classic Outlook without add-ins should paste the email link on the enquiry instead of attaching the `.msg`.

## Floor rules

- Home is a **live 3D digital twin** of the shop. Drag people and piles to place them; **Reset layout** puts them back. Tap a count badge for the order list. Welding and tagging use the same weld animation. The roller-door QC pins are hidden so you can place those people yourself. **Ready for final QC** still sits in finished goods. The plan refreshes while you stay on Home.
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

Workers log in when they **open this tab**. Moving between Home, Enquiries, Orders, and the other pages stays logged in. A **new tab or window** must log in even if another copy of the app is already open — the login is not shared across tabs. Log Out (or closing the tab) ends that login. Each person uses **only their own access code**. After they change it, the old code must not work — including the Manager’s code.

Add a **Tasks** column (column D). The first time the app opens it will create the header if it is missing.

| Name | Role | Password | Tasks |
| --- | --- | --- | --- |
| Sipho | Welder Tagger | 1234 | Welding, Tagging |
| Thabo | Quality Control | 1234 | Quality Control |
| Admin | Manager | **** | |

- **Admin** Home is the first page after login: enquiry/quote scorecard plus the live shop twin. The office job cart is new orders. Each station and pile shows a count button; click it for the order list. Logging in from an office page (Enquiries, Orders, …) also returns to Home. Production, Workers, Metrics, QC Reports, and Activity stay in the menu. Job title **Manager** is the person who adds and deletes people on Users.
- **Quality Control / QC** only sees QC work.
- Anyone else only sees the tasks listed. `Welding, Tagging` (or a role like `Welder Tagger`) means they pick Welding or Tagging after login, and only those boards appear.
- Painters need `Painting` on the Tasks column. That also lets them prep items for painting (no extra `Paint Preparation` task required). Assemblers with `Assembly` also get **Paint prep** on Ready for Assembly cards.

## Worker schedule

Admin sidebar → **Schedule**. Click Not Yet Started (and Ready for Steelwork) orders in the sequence they should be done, pick a process and a worker, then **Build calendar**. Block lengths come from past jobs of the same product and process (or a typical process time if that product is new). **Insert other task** (cleaning, meeting, etc.) pins a block on that worker’s calendar and pushes later work through 07:45–15:45, skipping lunch and weekends. This is a plan, not a live clock — the floor still starts/pauses/finishes as usual.
- If Tasks is blank, the app reads the Role cell (`Welder Tagger` → Welding + Tagging).

The tablet does **not** share a login across tabs. Opening the link in a new tab always asks for name and access code. Log Out ends that tab’s login.

Floor boards read only the latest production-log rows (not the whole history) and cache the order list for about 90 seconds. Start/pause/finish reuse that same log slice instead of rereading the sheet several times. Auto-refresh is once a minute while the tab is visible.

Assembly and Final QC must log **which backboard was used**, the same way plate/profile cutting logs steel. Add backboard names on the `Backboards` sheet (Category, Profile Name), or pick Custom during QC.

## Apps Script fallback (optional)

Paste `Code.gs` and `index.html` into the bound Apps Script project. The HTML file **must** be named `index` in the script editor.

1. Replace the existing `Code.gs` and `index` HTML with these files.
2. Open the web app once as admin. `doGet` tries to create a 5-minute idle check trigger.
3. Confirm **Project Settings → Time zone** is `Africa/Johannesburg`.

## Activity report

Admin sidebar → **Activity**. Daily / weekly / monthly view of what each person did, with regular hours capped at 7.5h per day and overtime shown separately. Overnight jobs and pause/resume bouts are split by calendar day, and each pause shows **when** and **why**. Admin → **Workers** uses the same per-day split for the hour log (one row per calendar day) and weekly totals.
