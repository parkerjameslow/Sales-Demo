"""
Builds the Demo Request "Queue" Excel workbook (.xlsx).

Run:  python3 build_queue_workbook.py
Output: Demo-Request-Queue.xlsx

Sheets:
  - Read Me      : how the form + queue fit together
  - New Request  : a clean intake form a rep can fill in and send back
  - Queue        : the master triage table you prioritize / assign to sprints
  - Lists        : dropdown source values (hidden)
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

# ---- palette (mirrors the original form mockup) ---------------------------
TEAL      = "3BA9C4"   # section headers
DARK      = "2C3A4B"   # body text
LAVENDER  = "EEF0FA"   # page background
WHITE     = "FFFFFF"
GREY_BORD = "C9D2E0"
RED       = "F4CCCC"   # Critical
ORANGE    = "FCE5CD"   # High
AMBER     = "FFF2CC"   # Medium
GREEN     = "D9EAD3"   # Low

thin = Side(style="thin", color=GREY_BORD)
box = Border(left=thin, right=thin, top=thin, bottom=thin)

PRIORITIES = ["Critical", "High", "Medium", "Low"]
STATUSES = ["New", "Triaged", "Scheduled", "In Progress", "Blocked", "Done", "Rejected"]

wb = Workbook()

# ===========================================================================
# Lists (hidden) — dropdown sources
# ===========================================================================
lists = wb.active
lists.title = "Lists"
lists["A1"] = "Priority"
lists["B1"] = "Status"
for i, p in enumerate(PRIORITIES, start=2):
    lists.cell(row=i, column=1, value=p)
for i, s in enumerate(STATUSES, start=2):
    lists.cell(row=i, column=2, value=s)
lists.sheet_state = "hidden"

# ===========================================================================
# Read Me
# ===========================================================================
rm = wb.create_sheet("Read Me")
rm.sheet_view.showGridLines = False
rm.column_dimensions["A"].width = 100

def rm_line(row, text, size=11, bold=False, color=DARK):
    c = rm.cell(row=row, column=1, value=text)
    c.font = Font(size=size, bold=bold, color=color)
    c.alignment = Alignment(wrap_text=True, vertical="top")

rm_line(1, "Demo Request — Form & Queue", size=18, bold=True, color=TEAL)
rm_line(3, "HOW THIS WORKS", bold=True, color=TEAL)
rm_line(4, "1. Reps submit a request (via the Microsoft Form you set up, or by filling the 'New Request' sheet and sending it back).")
rm_line(5, "2. Each request becomes one row on the 'Queue' sheet.")
rm_line(6, "3. You triage on the Queue: set Status, Final Priority, and the Assigned Sprint. Use the column filters to organize.")
rm_line(8, "SET UP THE MICROSOFT FORM (recommended front-end)", bold=True, color=TEAL)
rm_line(9, "Go to forms.office.com → New Form. Add these questions (matching the original form):")
rm_line(10, "   • Request Name — Text — Required")
rm_line(11, "   • Your Name — Text — Required   (or turn on 'record name' if internal only)")
rm_line(12, "   • Date Requested — Date")
rm_line(13, "   • Application(s) — Text — Required")
rm_line(14, "   • Jira # — Text")
rm_line(15, "   • Sprint # — Text")
rm_line(16, "   • Priority — Choice — Required — options: Critical, High, Medium, Low")
rm_line(17, "   • Full Description — Text (long answer) — Required")
rm_line(18, "Then: Responses tab → 'Open in Excel' to get the auto-collected workbook in OneDrive.")
rm_line(20, "FEED RESPONSES INTO THIS QUEUE", bold=True, color=TEAL)
rm_line(21, "Easiest: paste new response rows from the Forms Excel into the Queue sheet here, then triage.")
rm_line(22, "Automated: use Power Automate → trigger 'When a new response is submitted' → 'Add a row into a table' pointing at the Queue table in this file (stored on OneDrive/SharePoint).")
rm_line(24, "Prefer reps to skip the web form? They can fill the 'New Request' sheet and email this file back; copy the row into the Queue.")

# ===========================================================================
# New Request (intake form)
# ===========================================================================
nr = wb.create_sheet("New Request")
nr.sheet_view.showGridLines = False
nr.column_dimensions["A"].width = 3
nr.column_dimensions["B"].width = 22
nr.column_dimensions["C"].width = 60
nr.column_dimensions["D"].width = 3

def section(row, text):
    c = nr.cell(row=row, column=2, value=text.upper())
    c.font = Font(bold=True, size=12, color=TEAL)
    nr.cell(row=row, column=3)

def field(row, label, hint="", required=False):
    lab = label + (" *" if required else "")
    lc = nr.cell(row=row, column=2, value=lab)
    lc.font = Font(bold=True, color=DARK)
    lc.alignment = Alignment(vertical="center")
    ic = nr.cell(row=row, column=3)
    ic.border = box
    ic.fill = PatternFill("solid", fgColor=WHITE)
    if hint:
        hc = nr.cell(row=row + 1, column=3, value=hint)
        hc.font = Font(italic=True, size=9, color="8A97A8")
    return ic

title = nr.cell(row=1, column=2, value="Demo Request Form")
title.font = Font(bold=True, size=18, color=DARK)
nr.cell(row=2, column=2, value="Fields marked * are required.").font = Font(italic=True, size=9, color="8A97A8")

section(4, "The Basics")
field(5, "Request Name", "App name + what needs to change. Keep it skimmable.", required=True)
field(7, "Your Name", "First Last", required=True)
field(9, "Date Requested", "")

section(11, "Priority & Tracking")
field(12, "Application(s)", "Which demo app(s) does this change apply to?", required=True)
field(14, "Jira #", "Leave blank — Eternals will link after approval.")
field(16, "Sprint #", "e.g. Sprint 44")
prio_cell = field(18, "Priority", "Critical / High / Medium / Low", required=True)

section(20, "Tell Us More")
desc_lab = nr.cell(row=21, column=2, value="Full Description *")
desc_lab.font = Font(bold=True, color=DARK)
desc_lab.alignment = Alignment(vertical="top")
desc_cell = nr.cell(row=21, column=3)
desc_cell.border = box
desc_cell.fill = PatternFill("solid", fgColor=WHITE)
desc_cell.alignment = Alignment(wrap_text=True, vertical="top")
nr.merge_cells(start_row=21, start_column=3, end_row=26, end_column=3)
nr.cell(row=27, column=3,
        value="What needs to change and why? Current behavior, expected behavior, scenarios to cover.").font = Font(
        italic=True, size=9, color="8A97A8")

# priority dropdown on the intake form
dv_prio = DataValidation(type="list", formula1="=Lists!$A$2:$A$5", allow_blank=True)
nr.add_data_validation(dv_prio)
dv_prio.add(prio_cell)

# ===========================================================================
# Queue (triage table)
# ===========================================================================
q = wb.create_sheet("Queue")
q.sheet_view.showGridLines = False

headers = [
    "Request ID", "Date Received", "Requester", "Email", "Request Name",
    "Application(s)", "Sprint # (requested)", "Jira #", "Rep Priority",
    "Full Description",
    "Status", "Assigned Sprint", "Final Priority", "Owner", "Notes",
]
widths = [12, 14, 18, 24, 30, 18, 18, 12, 13, 45, 14, 16, 13, 16, 30]

# group banner: intake vs triage
q.merge_cells(start_row=1, start_column=1, end_row=1, end_column=10)
q.merge_cells(start_row=1, start_column=11, end_row=1, end_column=15)
b1 = q.cell(row=1, column=1, value="FROM THE REQUEST")
b2 = q.cell(row=1, column=11, value="YOUR TRIAGE")
for b in (b1, b2):
    b.font = Font(bold=True, color=WHITE, size=11)
    b.alignment = Alignment(horizontal="center")
b1.fill = PatternFill("solid", fgColor=TEAL)
b2.fill = PatternFill("solid", fgColor=DARK)

for col, (h, w) in enumerate(zip(headers, widths), start=1):
    c = q.cell(row=2, column=col, value=h)
    c.font = Font(bold=True, color=WHITE)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.fill = PatternFill("solid", fgColor=(TEAL if col <= 10 else DARK))
    c.border = box
    q.column_dimensions[get_column_letter(col)].width = w

# light borders + sample empty rows
LAST = 200
for r in range(3, LAST + 1):
    for col in range(1, len(headers) + 1):
        q.cell(row=r, column=col).border = box

q.freeze_panes = "A3"
q.auto_filter.ref = f"A2:{get_column_letter(len(headers))}{LAST}"

# dropdowns on Queue
dv_rep = DataValidation(type="list", formula1="=Lists!$A$2:$A$5", allow_blank=True)
dv_final = DataValidation(type="list", formula1="=Lists!$A$2:$A$5", allow_blank=True)
dv_status = DataValidation(type="list", formula1="=Lists!$B$2:$B$8", allow_blank=True)
for dv in (dv_rep, dv_final, dv_status):
    q.add_data_validation(dv)
dv_rep.add(f"I3:I{LAST}")     # Rep Priority
dv_status.add(f"K3:K{LAST}")  # Status
dv_final.add(f"M3:M{LAST}")   # Final Priority

# conditional color on the two priority columns
for col_letter in ("I", "M"):
    rng = f"{col_letter}3:{col_letter}{LAST}"
    q.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Critical"'], fill=PatternFill("solid", fgColor=RED)))
    q.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"High"'], fill=PatternFill("solid", fgColor=ORANGE)))
    q.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Medium"'], fill=PatternFill("solid", fgColor=AMBER)))
    q.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Low"'], fill=PatternFill("solid", fgColor=GREEN)))

# order: Read Me, New Request, Queue, (Lists hidden)
wb.move_sheet("Read Me", -(wb.sheetnames.index("Read Me")))
wb._sheets.sort(key=lambda s: {"Read Me": 0, "New Request": 1, "Queue": 2, "Lists": 3}[s.title])
wb.active = wb.sheetnames.index("Read Me")

wb.save("Demo-Request-Queue.xlsx")
print("Wrote Demo-Request-Queue.xlsx")
