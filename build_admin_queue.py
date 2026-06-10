"""
Builds the Admin Queue dashboard as an Excel workbook (.xlsx).

Run:    python3 build_admin_queue.py
Output: Admin-Queue.xlsx

A single "Admin Queue" sheet with:
  - 5 KPI cards (Pending Review / In Progress / Scheduled / Completed / Critical)
    driven by live COUNTIF formulas off the table below
  - A request table with Priority/Status dropdowns, color-coded priority,
    colored status pills, frozen headers and AutoFilter (the search/filter)
Sample rows mirror the dashboard mockup.
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

# ---- palette --------------------------------------------------------------
INK      = "1F2A44"   # headings
SLATE    = "6B7A90"   # muted text
HAIR     = "D8DEE9"   # hairlines
WHITE    = "FFFFFF"
HEADERBG = "F2F5FA"

BLUE   = "2F5BEA"   # pending
ORANGE = "E0892B"   # in progress / high
TEAL   = "3BA9C4"   # scheduled
GREEN  = "3FA34D"   # completed / low
RED    = "D64545"   # critical
PURPLE = "7C4DD6"   # approved
AMBER  = "C9A227"   # medium

# status pill fills (light) + text
STATUS_STYLE = {
    "New":         ("EDF0F5", SLATE),
    "In Review":   ("E3ECFF", BLUE),
    "Approved":    ("EEE6FB", PURPLE),
    "Scheduled":   ("DFF3F7", TEAL),
    "In Progress": ("FCEBD6", ORANGE),
    "Completed":   ("DCEFDD", GREEN),
    "Blocked":     ("F7DDDD", RED),
    "Rejected":    ("E6E9EE", INK),
}
PRIORITY_FILL = {
    "Critical": "F7D7D7",
    "High":     "FCE6CE",
    "Medium":   "FFF1C9",
    "Low":      "DCEFDD",
}
PRIORITIES = ["Critical", "High", "Medium", "Low"]
STATUSES = list(STATUS_STYLE.keys())

hair = Side(style="thin", color=HAIR)
box = Border(left=hair, right=hair, top=hair, bottom=hair)

# ---- sample data (from the mockup) ----------------------------------------
ROWS = [
    # Request, Requester, Application(s), Priority, Status, Jira#, Sprint, Submitted
    ["Agent — Screen Pop Customization",        "Sarah M.",  "Agent",              "High",     "Scheduled",   "DEMO-482", "Sprint 44", "May 2"],
    ["WFM — Real-Time Adherence Dashboard",      "James T.",  "WFM, Supervisor",    "Critical", "In Progress", "DEMO-479", "Sprint 44", "Apr 30"],
    ["QM — AI Evaluation Scorecard",             "Priya K.",  "Quality Management", "Medium",   "Approved",    "DEMO-491", "Sprint 45", "May 4"],
    ["Supervisor — Live Monitoring Panel",       "Marcus L.", "Supervisor",         "High",     "In Review",   "",         "",          "May 5"],
    ["Interactions Hub — Recording Playback",    "Dana W.",   "Interactions Hub",   "Low",      "New",         "",         "",          "May 6"],
    ["Smart Reach — Campaign Builder Flow",      "Jordan B.", "Smart Reach",        "High",     "Completed",   "DEMO-465", "Sprint 43", "Apr 22"],
    ["My Zone — Schedule Visibility",            "Sarah M.",  "My Zone",            "Medium",   "New",         "",         "",          "May 6"],
    ["Agent — After Call Work Timer",            "Priya K.",  "Agent",              "Medium",   "New",         "",         "",          "May 6"],
]

wb = Workbook()
ws = wb.active
ws.title = "Admin Queue"
ws.sheet_view.showGridLines = False

# columns sized so each KPI card (a column pair) is an even ~30 wide
COL_W = [20, 10, 18, 12, 16, 14, 16, 14, 30]  # A..I
for i, w in enumerate(COL_W, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w

# ---- title ----------------------------------------------------------------
ws.merge_cells("A1:I1")
t = ws["A1"]; t.value = "Admin Queue"
t.font = Font(size=20, bold=True, color=INK)
ws.merge_cells("A2:I2")
s = ws["A2"]; s.value = "Requests waiting for your review. Set a Status to move a request through the pipeline."
s.font = Font(size=10, color=SLATE)
ws.row_dimensions[1].height = 28

# ---- KPI cards ------------------------------------------------------------
# card n occupies a column pair across rows 4..7
DATA_RANGE_STATUS = "E:E"
DATA_RANGE_PRIO = "D:D"
cards = [
    ("PENDING REVIEW", "Need your decision",     BLUE,   f'=COUNTIF({DATA_RANGE_STATUS},"New")+COUNTIF({DATA_RANGE_STATUS},"In Review")'),
    ("IN PROGRESS",    "Being prepared",         ORANGE, f'=COUNTIF({DATA_RANGE_STATUS},"In Progress")'),
    ("SCHEDULED",      "On the calendar",        TEAL,   f'=COUNTIF({DATA_RANGE_STATUS},"Scheduled")'),
    ("COMPLETED",      "Delivered",              GREEN,  f'=COUNTIF({DATA_RANGE_STATUS},"Completed")'),
    ("CRITICAL",       "Needs immediate action", RED,    f'=COUNTIF({DATA_RANGE_PRIO},"Critical")'),
]
card_cols = [(1, 2), (3, 4), (5, 6), (7, 8), (9, 9)]  # I is double-width, so single col

ws.row_dimensions[4].height = 5    # accent bar
ws.row_dimensions[5].height = 34   # big number
ws.row_dimensions[6].height = 15   # label
ws.row_dimensions[7].height = 14   # subtitle

def style_card(c0, c1, label, subtitle, accent, formula):
    L, R = get_column_letter(c0), get_column_letter(c1)
    # accent bar
    ws.merge_cells(f"{L}4:{R}4")
    ws[f"{L}4"].fill = PatternFill("solid", fgColor=accent)
    # number
    ws.merge_cells(f"{L}5:{R}5")
    n = ws[f"{L}5"]; n.value = formula
    n.font = Font(size=26, bold=True, color=accent)
    n.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    # label
    ws.merge_cells(f"{L}6:{R}6")
    lb = ws[f"{L}6"]; lb.value = label
    lb.font = Font(size=9, bold=True, color=SLATE)
    lb.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    # subtitle
    ws.merge_cells(f"{L}7:{R}7")
    sb = ws[f"{L}7"]; sb.value = subtitle
    sb.font = Font(size=8, color=SLATE)
    sb.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    # white body + outline border
    for r in range(5, 8):
        for c in range(c0, c1 + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = PatternFill("solid", fgColor=WHITE)
    for c in range(c0, c1 + 1):
        for r in range(4, 8):
            cur = ws.cell(row=r, column=c).border
            top = Side(style="thin", color=HAIR) if r == 4 else cur.top
            bot = Side(style="thin", color=HAIR) if r == 7 else cur.bottom
            left = Side(style="thin", color=HAIR) if c == c0 else cur.left
            right = Side(style="thin", color=HAIR) if c == c1 else cur.right
            ws.cell(row=r, column=c).border = Border(left=left, right=right, top=top, bottom=bot)

for (c0, c1), (label, sub, accent, formula) in zip(card_cols, cards):
    style_card(c0, c1, label, sub, accent, formula)

# ---- table ----------------------------------------------------------------
HEADER_ROW = 10
headers = ["Request", "Requester", "Application(s)", "Priority",
           "Status", "Jira #", "Sprint", "Submitted", "Notes"]
for c, h in enumerate(headers, start=1):
    cell = ws.cell(row=HEADER_ROW, column=c, value=h)
    cell.font = Font(bold=True, color=INK, size=10)
    cell.fill = PatternFill("solid", fgColor=HEADERBG)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    cell.border = box
ws.row_dimensions[HEADER_ROW].height = 22

first = HEADER_ROW + 1
for i, row in enumerate(ROWS):
    r = first + i
    for c, val in enumerate(row, start=1):
        cell = ws.cell(row=r, column=c, value=(val if val else "—"))
        cell.border = box
        cell.alignment = Alignment(vertical="center", wrap_text=(c == 1))
        if c == 1:
            cell.font = Font(bold=True, color=INK, size=10)
        elif c in (4, 5):
            cell.font = Font(bold=True, size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        else:
            cell.font = Font(color=INK, size=10)
    ws.row_dimensions[r].height = 22

# extra blank rows so the queue can grow (formatted + dropdowns)
LAST = first + 60
for r in range(first + len(ROWS), LAST + 1):
    for c in range(1, len(headers) + 1):
        ws.cell(row=r, column=c).border = box
    ws.row_dimensions[r].height = 20

ws.freeze_panes = f"A{first}"
ws.auto_filter.ref = f"A{HEADER_ROW}:I{LAST}"

# dropdowns
dv_prio = DataValidation(type="list", formula1='"%s"' % ",".join(PRIORITIES), allow_blank=True)
dv_stat = DataValidation(type="list", formula1='"%s"' % ",".join(STATUSES), allow_blank=True)
ws.add_data_validation(dv_prio); ws.add_data_validation(dv_stat)
dv_prio.add(f"D{first}:D{LAST}")
dv_stat.add(f"E{first}:E{LAST}")

# conditional formatting — priority fills
for val, fill in PRIORITY_FILL.items():
    ws.conditional_formatting.add(
        f"D{first}:D{LAST}",
        CellIsRule(operator="equal", formula=[f'"{val}"'],
                   fill=PatternFill("solid", fgColor=fill)))
# conditional formatting — status pills
for val, (fill, txt) in STATUS_STYLE.items():
    ws.conditional_formatting.add(
        f"E{first}:E{LAST}",
        CellIsRule(operator="equal", formula=[f'"{val}"'],
                   fill=PatternFill("solid", fgColor=fill),
                   font=Font(bold=True, color=txt)))

wb.save("Admin-Queue.xlsx")
print("Wrote Admin-Queue.xlsx")
