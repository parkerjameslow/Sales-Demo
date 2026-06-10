# Demo Request — Form & Queue

An intake workflow for demo change requests:

**Reps submit a request → it lands in a central Queue → you triage it (priority, sprint, status).**

## ⭐ The web app (`docs/`) — a real site, two links

A clean web form for reps + an **Admin Queue** dashboard for you, hosted free
on GitHub Pages with the data in a free Supabase (Postgres) database:

- **The form** (`docs/index.html`) — three guided sections (The Basics /
  Priority & Tracking / Tell Us More), tap-to-pick priority cards with
  plain-language hints, inline validation, and a success screen with a
  reference number (e.g. `REQ-0007`). Works great on phones, no login.
- **The Admin Queue** (`docs/dashboard.html`) — five KPI cards (Pending
  Review / In Progress / Scheduled / Completed / Critical), full-text search,
  status + priority filter chips, sortable columns, color-coded priority and
  status pills, **Approve →** buttons, a **Details** modal where you edit
  status, priority, Jira # and Sprint, and **Export CSV**.

**Setup is one-time, ~5 minutes — follow [`SETUP.md`](SETUP.md).**
Database schema lives in [`supabase/schema.sql`](supabase/schema.sql);
the site's connection values go in [`docs/config.js`](docs/config.js).

---

## Alternatives also in this repo

| File | What it does |
|------|--------------|
| `docs/` + `supabase/` | **⭐ The web form + Admin Queue dashboard** (see above) |
| `webapp/` | Google Apps Script version (form + dashboard backed by a Google Sheet) |
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
