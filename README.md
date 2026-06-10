# Demo Request — Form & Queue

A Microsoft-based intake workflow for demo change requests:

**Reps submit a request → it lands in a central Queue → you triage it (priority, sprint, status).**

## What's in here

| File | What it does |
|------|--------------|
| `Demo-Request-Queue.xlsx` | The deliverable: an Excel workbook with a **New Request** intake sheet, a **Queue** triage table (the heart of it), and a **Read Me** tab |
| `build_queue_workbook.py` | The generator that builds the `.xlsx` (re-run to tweak fields/colors) |

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

## Recommended front-end: Microsoft Forms

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
