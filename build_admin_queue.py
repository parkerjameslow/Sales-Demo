"""
Builds the Admin Queue dashboard as an Excel workbook (.xlsx).

Run:    python3 build_admin_queue.py
Output: Admin-Queue.xlsx

Design goals (bulletproof + polished):
  - The request list is a real Excel TABLE ("Requests"). KPI cards use
    STRUCTURED references (Requests[Status]) so they stay correct even if you
    move or insert columns.
  - The workbook is saved in AUTOMATIC calc mode with full-calc-on-load, so the
    cards compute the instant you open the file (no stuck zeros).
  - 5 KPI "status boxes" with colored top bars, big centered numbers and soft
    tinted backgrounds, plus a Top-Priority banner.
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter

# ---- palette --------------------------------------------------------------
INK   = "1F2A44"   # headings
SLATE = "6B7A90"   # muted text
HAIR  = "D8DEE9"   # hairlines
WHITE = "FFFFFF"

BLUE   = "2F5BEA"   # pending
ORANGE = "E0892B"   # in progress / high
TEAL   = "2BA7C4"   # scheduled
GREEN  = "3FA34D"   # completed / low
RED    = "D64545"   # critical
PURPLE = "7C4DD6"   # approved
AMBER  = "C9A227"   # medium

# soft tints for KPI card bodies
TINT = {BLUE: "EEF2FE", ORANGE: "FDF1E6", TEAL: "E8F5F8", GREEN: "EAF4EB", RED: "FBECEC"}

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
PRIORITY_STYLE = {
    "Critical": ("F7D7D7", RED),
    "High":     ("FCE6CE", ORANGE),
    "Medium":   ("FFF1C9", "9A7B0A"),
    "Low":      ("DCEFDD", GREEN),
}
PRIORITIES = list(PRIORITY_STYLE.keys())
STATUSES = list(STATUS_STYLE.keys())

hair = Side(style="thin", color=HAIR)
box = Border(left=hair, right=hair, top=hair, bottom=hair)

# ---- sample data (from the mockup) ----------------------------------------
ROWS = [
    ["Agent — Screen Pop Customization",     "Sarah M.",  "Agent",              "High",     "Scheduled",   "DEMO-482", "Sprint 44", "May 2"],
    ["WFM — Real-Time Adherence Dashboard",   "James T.",  "WFM, Supervisor",    "Critical", "In Progress", "DEMO-479", "Sprint 44", "Apr 30"],
    ["QM — AI Evaluation Scorecard",          "Priya K.",  "Quality Management", "Medium",   "Approved",    "DEMO-491", "Sprint 45", "May 4"],
    ["Supervisor — Live Monitoring Panel",    "Marcus L.", "Supervisor",         "High",     "In Review",   "",         "",          "May 5"],
    ["Interactions Hub — Recording Playback", "Dana W.",   "Interactions Hub",   "Low",      "New",         "",         "",          "May 6"],
    ["Smart Reach — Campaign Builder Flow",   "Jordan B.", "Smart Reach",        "High",     "Completed",   "DEMO-465", "Sprint 43", "Apr 22"],
    ["My Zone — Schedule Visibility",         "Sarah M.",  "My Zone",            "Medium",   "New",         "",         "",          "May 6"],
    ["Agent — After Call Work Timer",         "Priya K.",  "Agent",              "Medium",   "New",         "",         "",          "May 6"],
]
HEADERS = ["Request", "Requester", "Application(s)", "Priority",
           "Status", "Jira #", "Sprint", "Submitted", "Notes"]

wb = Workbook()
ws = wb.active
ws.title = "Admin Queue"
ws.sheet_view.showGridLines = False

# make the file always recompute on open, in automatic mode (bulletproofing)
wb.calculation.calcMode = "auto"
wb.calculation.fullCalcOnLoad = True

# columns sized so each KPI card (a column pair) is an even ~30 wide
COL_W = [20, 10, 18, 12, 16, 14, 16, 14, 30]  # A..I
for i, w in enumerate(COL_W, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w

# ---- title ----------------------------------------------------------------
ws.merge_cells("A1:I1")
ws["A1"].value = "Admin Queue"
ws["A1"].font = Font(size=22, bold=True, color=INK)
ws.merge_cells("A2:I2")
ws["A2"].value = "Requests waiting for your review. Set a Status to move a request through the pipeline."
ws["A2"].font = Font(size=10, color=SLATE)
ws.row_dimensions[1].height = 30
ws.row_dimensions[3].height = 6

# ---- KPI cards (structured-reference COUNTIFs) ----------------------------
T = "Requests"
cards = [
    ("PENDING REVIEW", "Need your decision",     BLUE,
     f'=COUNTIF({T}[Status],"New")+COUNTIF({T}[Status],"In Review")'),
    ("IN PROGRESS",    "Being prepared",         ORANGE,
     f'=COUNTIF({T}[Status],"In Progress")'),
    ("SCHEDULED",      "On the calendar",        TEAL,
     f'=COUNTIF({T}[Status],"Scheduled")'),
    ("COMPLETED",      "Delivered",              GREEN,
     f'=COUNTIF({T}[Status],"Completed")'),
    ("CRITICAL",       "Needs immediate action", RED,
     f'=COUNTIF({T}[Priority],"Critical")'),
]
card_cols = [(1, 2), (3, 4), (5, 6), (7, 8), (9, 9)]

R0 = 4
ws.row_dimensions[R0].height = 7       # accent bar
ws.row_dimensions[R0 + 1].height = 46  # big number
ws.row_dimensions[R0 + 2].height = 17  # label
ws.row_dimensions[R0 + 3].height = 16  # subtitle

def kpi_card(c0, c1, label, subtitle, accent, formula):
    L, R = get_column_letter(c0), get_column_letter(c1)
    body = TINT.get(accent, WHITE)
    # accent bar
    ws.merge_cells(f"{L}{R0}:{R}{R0}")
    ws[f"{L}{R0}"].fill = PatternFill("solid", fgColor=accent)
    # number
    ws.merge_cells(f"{L}{R0+1}:{R}{R0+1}")
    n = ws[f"{L}{R0+1}"]; n.value = formula
    n.font = Font(size=34, bold=True, color=accent)
    n.alignment = Alignment(horizontal="center", vertical="center")
    # label
    ws.merge_cells(f"{L}{R0+2}:{R}{R0+2}")
    lb = ws[f"{L}{R0+2}"]; lb.value = label
    lb.font = Font(size=10, bold=True, color=INK)
    lb.alignment = Alignment(horizontal="center", vertical="center")
    # subtitle
    ws.merge_cells(f"{L}{R0+3}:{R}{R0+3}")
    sb = ws[f"{L}{R0+3}"]; sb.value = subtitle
    sb.font = Font(size=8, color=SLATE)
    sb.alignment = Alignment(horizontal="center", vertical="center")
    # tinted body + outline
    for r in range(R0 + 1, R0 + 4):
        for c in range(c0, c1 + 1):
            ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor=body)
    for c in range(c0, c1 + 1):
        for r in range(R0, R0 + 4):
            cur = ws.cell(row=r, column=c).border
            top = hair if r == R0 else cur.top
            bot = hair if r == R0 + 3 else cur.bottom
            left = hair if c == c0 else cur.left
            right = hair if c == c1 else cur.right
            ws.cell(row=r, column=c).border = Border(left=left, right=right, top=top, bottom=bot)

for (c0, c1), (label, sub, accent, formula) in zip(card_cols, cards):
    kpi_card(c0, c1, label, sub, accent, formula)

# ---- Top Priority banner (full width, color-tinted by the priority) -------
ws.row_dimensions[8].height = 6
BANNER = 9
ws.merge_cells(f"A{BANNER}:I{BANNER}")
bn = ws[f"A{BANNER}"]
bn.value = (f'="⚑   Top priority in the queue:   "&'
            f'IFS('
            f'COUNTIF({T}[Priority],"Critical")>0,"Critical",'
            f'COUNTIF({T}[Priority],"High")>0,"High",'
            f'COUNTIF({T}[Priority],"Medium")>0,"Medium",'
            f'COUNTIF({T}[Priority],"Low")>0,"Low",'
            f'TRUE,"None set")')
bn.font = Font(size=12, bold=True, color=INK)
bn.alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws.row_dimensions[BANNER].height = 26
for c in range(1, 10):
    ws.cell(row=BANNER, column=c).border = box
# tint the banner by which priority word it contains
for word, (fill, txt) in PRIORITY_STYLE.items():
    ws.conditional_formatting.add(
        f"A{BANNER}:I{BANNER}",
        FormulaRule(formula=[f'ISNUMBER(SEARCH("{word}",$A${BANNER}))'],
                    fill=PatternFill("solid", fgColor=fill),
                    font=Font(size=12, bold=True, color=txt)))
ws.row_dimensions[10].height = 8

# ---- request table --------------------------------------------------------
HEADER_ROW = 11
for c, h in enumerate(HEADERS, start=1):
    cell = ws.cell(row=HEADER_ROW, column=c, value=h)
    cell.font = Font(bold=True, color=INK, size=10)
    cell.alignment = Alignment(horizontal="left", vertical="center")
ws.row_dimensions[HEADER_ROW].height = 22

first = HEADER_ROW + 1
for i, row in enumerate(ROWS):
    r = first + i
    for c, val in enumerate(row, start=1):
        cell = ws.cell(row=r, column=c, value=(val if val else ""))
        cell.alignment = Alignment(vertical="center", wrap_text=(c in (1, 9)))
        if c == 1:
            cell.font = Font(bold=True, color=INK, size=10)
        elif c in (4, 5):
            cell.font = Font(bold=True, size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center")
        else:
            cell.font = Font(color=INK, size=10)
    ws.row_dimensions[r].height = 24

# blank rows so the queue can grow
GROW = 50
LAST = first + len(ROWS) + GROW - 1
for r in range(first + len(ROWS), LAST + 1):
    ws.row_dimensions[r].height = 22
    ws.cell(row=r, column=4).alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=5).alignment = Alignment(horizontal="center")

# real Excel Table -> structured refs + banded styling + filter buttons
tbl = Table(displayName=T, ref=f"A{HEADER_ROW}:I{LAST}")
tbl.tableStyleInfo = TableStyleInfo(
    name="TableStyleLight15", showFirstColumn=False, showLastColumn=False,
    showRowStripes=True, showColumnStripes=False)
ws.add_table(tbl)

ws.freeze_panes = f"A{first}"

# dropdowns
dv_prio = DataValidation(type="list", formula1='"%s"' % ",".join(PRIORITIES), allow_blank=True)
dv_stat = DataValidation(type="list", formula1='"%s"' % ",".join(STATUSES), allow_blank=True)
ws.add_data_validation(dv_prio); ws.add_data_validation(dv_stat)
dv_prio.add(f"D{first}:D{LAST}")
dv_stat.add(f"E{first}:E{LAST}")

# conditional formatting — priority + status colors
for val, (fill, txt) in PRIORITY_STYLE.items():
    ws.conditional_formatting.add(
        f"D{first}:D{LAST}",
        CellIsRule(operator="equal", formula=[f'"{val}"'],
                   fill=PatternFill("solid", fgColor=fill),
                   font=Font(bold=True, color=txt)))
for val, (fill, txt) in STATUS_STYLE.items():
    ws.conditional_formatting.add(
        f"E{first}:E{LAST}",
        CellIsRule(operator="equal", formula=[f'"{val}"'],
                   fill=PatternFill("solid", fgColor=fill),
                   font=Font(bold=True, color=txt)))

wb.save("Admin-Queue.xlsx")
print("Wrote Admin-Queue.xlsx")
