# Demo Request — Form & Queue

An intake workflow for demo change requests:

**Reps submit a request → it lands in a central Queue → you triage it (priority, sprint, status).**

## ⭐ Recommended: the web app (`webapp/`) — one link for reps, one for you

A self-hosted form + live dashboard built on Google Apps Script. Reps get a
**single link** to a polished, mobile-friendly form; every submission lands
instantly in a spreadsheet (downloadable as `.xlsx` any time); and you get a
**live dashboard** with KPI cards, search, status/priority filters, grouping,
and one-click status triage — no copying rows, no manual collection.

### Deploy it (about 3 minutes, one time)

1. Go to **[script.google.com](https://script.google.com)** → **New project**.
2. Replace the default `Code.gs` with **`webapp/Code.gs`**, then use the
   **+** next to *Files* → *HTML* to add **`Form`** and **`Dashboard`**, pasting
   in `webapp/Form.html` and `webapp/Dashboard.html`.
3. **Deploy → New deployment → Web app**:
   - *Execute as:* **Me**
   - *Who has access:* **Anyone with the link** (or restrict to your org)
4. Copy the web app URL. That's it:
   - **Send the URL to your reps** → that's the request form.
   - **Add `?page=dashboard` to the URL** → that's your live admin dashboard.

The data spreadsheet (*Demo Request Queue — Live Data*) is created in your
Drive automatically on the first submission; the dashboard's **Open Sheet**
button takes you straight to it, and **File → Download → .xlsx** gives you
the Excel copy whenever you need one.

### What each page does

- **The form** — three guided sections (The Basics / Priority & Tracking /
  Tell Us More), tap-to-pick priority cards, inline validation, a success
  screen with a reference number (e.g. `REQ-0007`), and a "submit another"
  button. Works great on phones.
- **The dashboard** — five live KPI cards (Total / Needs Review / In
  Progress / Completed / Critical Open), full-text search, status filter
  chips, a priority filter, **Group by** Status / Priority / Application /
  Requester, color-coded priority bars and status pills, and a status
  dropdown on every request that writes back to the sheet instantly.

---

## Alternative: the Excel/Sheets workbooks

If you'd rather stay file-based (no web app), the original workbooks are here:

| File | What it does |
|------|--------------|
| `webapp/` | **⭐ The web form + live dashboard** (see above) |
| `Admin-Queue.xlsx` | **The admin dashboard** in Excel — KPI cards + filterable request table |
| `build_admin_queue.py` | Generator for `Admin-Queue.xlsx` (re-run to tweak data/colors) |
| `build_admin_queue.gs` | **The same dashboard for Google Sheets** — paste into script.google.com and run |
| `Demo-Request-Queue.xlsx` | Intake workbook: a **New Request** form sheet + a **Queue** triage table |
| `build_queue_workbook.py` | Generator for `Demo-Request-Queue.xlsx` |

## Admin Queue dashboard (Excel + Google Sheets)

A viewing queue for all requests, modeled on the admin-dashboard mockup:

- **5 KPI cards** — Pending Review / In Progress / Scheduled / Completed / Critical.
  These are **live `COUNTIF` formulas**: change a Status and the counts update.
- **Request table** — Request, Requester, Application(s), Priority, Status, Jira #,
  Sprint, Submitted, Notes. Priority is color-coded; Status shows as colored pills;
  both are dropdowns. Header is frozen with a filter (this replaces the mockup's
  search box + filter chips — click a column's filter to search/narrow).
- The mockup's **Approve** buttons map to the **Status dropdown** — set a row to
  *Approved* / *Scheduled* / *In Progress* to move it through the pipeline.

**Excel:** open `Admin-Queue.xlsx`. To regenerate: `python3 build_admin_queue.py`.

**Google Sheets:** go to [script.google.com](https://script.google.com) → New project →
paste in `build_admin_queue.gs` → Run `buildAdminQueue` → authorize → open the link
printed in **View → Logs**. It builds the same dashboard as a live Google Sheet.

---

## Intake workbook (Demo-Request-Queue.xlsx)

## The workbook

- **Read Me** — how the form and queue fit together, plus Microsoft Form setup steps
- **New Request** — a clean, styled intake form mirroring the original mockup
  (The Basics / Priority & Tracking / Tell Us More), with a Priority dropdown.
  A rep can fill this in and send the file back if you don't want the web form.
- **Queue** — the master triage table. Each request is one row. Columns are split:
  - **From the request** — Request ID, Date Received, Requester, Email, Request Name,
    Application(s), Sprint # (requested), Jira #, Rep Priority, Full Description
  - **Your triage** — Status, Assigned Sprint, Final Priority, Owner, Notes
  - Priority and Status cells are dropdowns; priority columns are color-coded
    (Critical = red → Low = green); the header row is frozen with filters on.

## Microsoft Forms front-end (if you can't use the web app)

This session can't create a live Microsoft Form for you (no access to your M365
account, and MS Forms has no script generator), so here's the 2-minute setup:

1. Go to **[forms.office.com](https://forms.office.com)** → **New Form**.
2. Add these questions (they match the original form):
   - Request Name — *Text, Required*
   - Your Name — *Text, Required*
   - Date Requested — *Date*
   - Application(s) — *Text, Required*
   - Jira # — *Text*
   - Sprint # — *Text*
   - Priority — *Choice, Required* — Critical / High / Medium / Low
   - Full Description — *Long answer text, Required*
3. **Share** the form link with the sales team.
4. On the **Responses** tab, click **Open in Excel** — Microsoft auto-collects every
   submission into a workbook in your OneDrive.

### Feeding responses into the Queue

- **Manual (simplest):** copy new rows from the Forms response workbook into the
  **Queue** sheet, then triage.
- **Automated:** use **Power Automate** → trigger *"When a new response is submitted"*
  → action *"Add a row into a table"* pointing at the Queue table (store this file on
  OneDrive/SharePoint). New requests then drop into the Queue automatically.

## Regenerating the workbook

```bash
pip install openpyxl
python3 build_queue_workbook.py   # writes Demo-Request-Queue.xlsx
```

Edit field labels, dropdown options, or colors at the top of `build_queue_workbook.py`.
