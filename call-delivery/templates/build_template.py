import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter

OUT = "/home/user/mygamemystory/call-delivery/templates/Call_Delivery_Intake_Template.xlsx"

# ---- palette ----
NAVY   = "1F2A44"
BLUE   = "2E5AAC"
LTBLUE = "DCE6F7"
GREY   = "F2F4F8"
GOLD   = "C7A008"
WHITE  = "FFFFFF"
BORDER = "B8C0D0"

thin = Side(style="thin", color=BORDER)
box  = Border(left=thin, right=thin, top=thin, bottom=thin)

def fill(hexc): return PatternFill("solid", fgColor=hexc)

wb = openpyxl.Workbook()

# =========================================================
# LISTS (dropdown sources) -- put first so we can reference it
# =========================================================
lists = wb.active
lists.title = "Lists"
lists.sheet_properties.tabColor = GREY

groups = ["OM_All","FE","BE","PCO","MOD","AUTO Direct Collections","Repo","Auto",
          "ARC","CPOD / NCC","CARE (Outbound)","Branch Central","Branch Vendor",
          "Optional Products","West Coast Pilot","Spanish","Card","Call Escalation Team",
          "Sales","Recoveries"]
waves  = ["E","C","M","P","E-C","C-M","M-P","E-C-M","C-M-P","E-C-M-P","All"]
tags   = ["AMs","Branch","Blitz","TSC (All)","WFO","Vendor","NCs","1s & Slows",
          "Recent Pay","Broken Prom","All","Other"]
lists_ids = ["ADAN","ADAO","ADAR","FEAO","FEAR","BEAO","BEAR","PCAO","PCAR",
             "REAO","REAH","BWAO","BWAR","NCAO","NCAH","LDA3","LDA4","SAMO","SAMR",
             "BR_CAP_CONS_HB","BR_CAP_CONS_LB","BR_CAP_CONS_SLOW_HB","BR_CAP_CONS_SLOW_LB",
             "BR_1PAY_HB","BR_1PAY_LB","BR_SLOW_HB","BR_SLOW_LB","BR_ADHOC_HB","BR_ADHOC_LB",
             "N/A"]

cols = {"A":("Business Groups",groups),"B":("Timezone Waves",waves),
        "C":("Segment / Tag",tags),"D":("List ID / Campaign",lists_ids)}
for col,(hdr,vals) in cols.items():
    c = lists[f"{col}1"]; c.value = hdr; c.font = Font(bold=True,color=WHITE); c.fill = fill(BLUE)
    for i,v in enumerate(vals, start=2):
        lists[f"{col}{i}"] = v
    lists.column_dimensions[col].width = 24
lists["F1"] = "DO NOT EDIT — this sheet powers the dropdowns on the Schedule tab."
lists["F1"].font = Font(italic=True, color="B00020")

def rng(col, vals): return f"Lists!${col}$2:${col}${len(vals)+1}"
GROUPS_R = rng("A",groups); WAVES_R = rng("B",waves)
TAGS_R = rng("C",tags); IDS_R = rng("D",lists_ids)

# =========================================================
# SCHEDULE (the entry sheet)
# =========================================================
sh = wb.create_sheet("Schedule", 0)
sh.sheet_properties.tabColor = BLUE
sh.sheet_view.showGridLines = False

# Title banner
sh.merge_cells("B2:I2")
t = sh["B2"]; t.value = "CALL DELIVERY  —  Outbound Dialing Schedule (Intake)"
t.font = Font(bold=True, size=16, color=WHITE); t.fill = fill(NAVY)
t.alignment = Alignment(vertical="center", horizontal="left", indent=1)
sh.row_dimensions[2].height = 30
sh.merge_cells("B3:I3")
s = sh["B3"]; s.value = ("Fill one row per scheduled pass.  Use the dropdowns — do not free-type.  "
                         "All times are EASTERN.  Save this file to the shared Intake folder (see README).")
s.font = Font(italic=True, size=10, color=NAVY); s.fill = fill(LTBLUE)
s.alignment = Alignment(vertical="center", horizontal="left", indent=1)
sh.row_dimensions[3].height = 20

# Metadata header cells
def meta(cell, label, lcell, hint=None, width=None):
    lc = sh[cell]; lc.value = label; lc.font = Font(bold=True, color=NAVY)
    lc.alignment = Alignment(horizontal="right", vertical="center")
    v = sh[lcell]; v.fill = fill(WHITE); v.border = box
    v.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    if hint: v.value = hint; v.font = Font(color="9AA3B2", italic=True)

meta("B5","Business Group:","C5")
meta("E5","Period (Month):","F5")
meta("B6","Submitted By:","C6")
meta("E6","Submitted Date:","F6")
sh["C5"].font = Font(bold=True, color=BLUE, size=12)
for r in (5,6): sh.row_dimensions[r].height = 20

# Business group dropdown on C5 (single-select)
dv_grp = DataValidation(type="list", formula1=GROUPS_R, allow_blank=False, showDropDown=False)
dv_grp.error = "Pick your business group from the list."; dv_grp.errorTitle = "Business Group"
sh.add_data_validation(dv_grp); dv_grp.add(sh["C5"])

# ---- Table ----
HDR_ROW = 9
headers = ["Business Group","Date","Start Time (ET)","Timezone Wave",
           "List ID / Campaign","Priority","Segment / Tag","Special Instructions"]
first_col = 2  # B
N_ROWS = 400
last_col = first_col + len(headers) - 1
for i,h in enumerate(headers):
    c = sh.cell(row=HDR_ROW, column=first_col+i, value=h)
    c.font = Font(bold=True, color=WHITE); c.fill = fill(BLUE)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = box
sh.row_dimensions[HDR_ROW].height = 30

data_first = HDR_ROW + 1
data_last  = HDR_ROW + N_ROWS
Lc = lambda n: get_column_letter(n)
BG,DT,TM,TZ,LI,PR,SG,SI = [Lc(first_col+i) for i in range(8)]

# Auto-stamp Business Group from C5
for r in range(data_first, data_last+1):
    cell = sh[f"{BG}{r}"]
    cell.value = f'=IF($C$5="","",$C$5)'
    cell.font = Font(color="6B7280")
    cell.alignment = Alignment(horizontal="center", vertical="center")

# number/date formats
for r in range(data_first, data_last+1):
    sh[f"{DT}{r}"].number_format = "mm/dd/yyyy"
    sh[f"{TM}{r}"].number_format = "h:mm AM/PM"
    for cc in (DT,TM,TZ,LI,PR,SG,SI):
        sh[f"{cc}{r}"].border = box
        sh[f"{cc}{r}"].alignment = Alignment(vertical="center",
            horizontal="left" if cc==SI else "center", indent=1 if cc==SI else 0)
    sh[f"{BG}{r}"].border = box

# widths
widths = {BG:20, DT:13, TM:14, TZ:14, LI:20, PR:9, SG:15, SI:40, "A":2}
for c,w in widths.items(): sh.column_dimensions[c].width = w

# Excel Table
ref = f"{BG}{HDR_ROW}:{SI}{data_last}"
tbl = Table(displayName="tblSchedule", ref=ref)
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True,
                                    showColumnStripes=False, showFirstColumn=False,
                                    showLastColumn=False)
sh.add_table(tbl)

# Validations
def add_dv(**kw):
    dv = DataValidation(**kw); sh.add_data_validation(dv); return dv
dv_tz = add_dv(type="list", formula1=WAVES_R, allow_blank=True, showDropDown=False)
dv_li = add_dv(type="list", formula1=IDS_R,   allow_blank=True, showDropDown=False)
dv_sg = add_dv(type="list", formula1=TAGS_R,  allow_blank=True, showDropDown=False)
dv_pr = add_dv(type="whole", operator="between", formula1="1", formula2="20", allow_blank=True)
dv_pr.error = "Priority must be a whole number 1–20."; dv_pr.errorTitle="Priority"
dv_dt = add_dv(type="date", operator="greaterThanOrEqual", formula1="DATE(2020,1,1)", allow_blank=True)
dv_tm = add_dv(type="time", operator="between", formula1="0", formula2="1", allow_blank=True)
dr = f"{data_first}:{data_last}"
dv_tz.add(f"{TZ}{data_first}:{TZ}{data_last}")
dv_li.add(f"{LI}{data_first}:{LI}{data_last}")
dv_sg.add(f"{SG}{data_first}:{SG}{data_last}")
dv_pr.add(f"{PR}{data_first}:{PR}{data_last}")
dv_dt.add(f"{DT}{data_first}:{DT}{data_last}")
dv_tm.add(f"{TM}{data_first}:{TM}{data_last}")

# Seed a couple of example rows (light grey, clearly examples)
examples = [
    ("=IF($C$5=\"\",\"\",$C$5)", "2026-07-20", "08:15", "E", "FEAO", 1, "AMs", "Eastern AM launch"),
    ("=IF($C$5=\"\",\"\",$C$5)", "2026-07-20", "11:30", "C-M-P", "FEAR", 2, "Branch", "Mid-day branch pass"),
    ("=IF($C$5=\"\",\"\",$C$5)", "2026-07-20", "17:00", "All", "N/A", 3, "Blitz", "5pm blitz"),
]
import datetime as _dt
for i,(bg,d,tm,tz,li,pr,sg,si) in enumerate(examples):
    r = data_first + i
    sh[f"{DT}{r}"] = _dt.datetime.strptime(d,"%Y-%m-%d")
    hh,mm = map(int,tm.split(":")); sh[f"{TM}{r}"] = _dt.time(hh,mm)
    sh[f"{TZ}{r}"]=tz; sh[f"{LI}{r}"]=li; sh[f"{PR}{r}"]=pr; sh[f"{SG}{r}"]=sg; sh[f"{SI}{r}"]=si

sh.freeze_panes = f"{DT}{data_first}"

# =========================================================
# README sheet
# =========================================================
rd = wb.create_sheet("README")
rd.sheet_properties.tabColor = GOLD
rd.sheet_view.showGridLines = False
rd.column_dimensions["B"].width = 100
lines = [
    ("CALL DELIVERY — INTAKE TEMPLATE  ·  HOW TO USE", "h1"),
    ("", ""),
    ("1.  Set the header once.", "h2"),
    ("     • Pick your Business Group in cell C5.  It auto-stamps every row.", "p"),
    ("     • Fill Period (month), Submitted By, and Submitted Date.", "p"),
    ("", ""),
    ("2.  Enter one row per scheduled pass on the Schedule tab.", "h2"),
    ("     • Date .............. the calendar day of the pass (mm/dd/yyyy).", "p"),
    ("     • Start Time (ET) ... EASTERN time the pass starts (e.g. 8:15 AM).", "p"),
    ("     • Timezone Wave ..... which zones dial in this pass: E, C, M, P or a combo (E-C-M-P).", "p"),
    ("     • List ID / Campaign  the campaign/list code (dropdown).", "p"),
    ("     • Priority .......... 1 = highest.  Whole number 1–20.", "p"),
    ("     • Segment / Tag ..... AMs, Branch, Blitz, TSC, Vendor, etc.", "p"),
    ("     • Special Instructions  anything the Call Delivery team should know.", "p"),
    ("", ""),
    ("3.  Use the dropdowns — do not free-type.", "h2"),
    ("     Consistent values are what let the master workbook import every group automatically.", "p"),
    ("     Need a value that isn't listed? Put it in Special Instructions and tell Call Delivery.", "p"),
    ("", ""),
    ("4.  Save to the shared Intake folder using the exact name:", "h2"),
    ("     \\CallDelivery\\Intake\\<YYYY-MM>\\<GROUP>.xlsx     e.g.  \\CallDelivery\\Intake\\2026-07\\FE.xlsx", "code"),
    ("     One file per group per month.  Overwrite the same file when your schedule changes.", "p"),
    ("", ""),
    ("What happens next (automatic):", "h2"),
    ("     Call Delivery's master workbook refreshes from this folder with Power Query.", "p"),
    ("     Your rows flow straight into the consolidated Daily Rollup — no copy/paste, no retyping.", "p"),
    ("", ""),
    ("Notes", "h2"),
    ("     • All times are Eastern.  • Delete the three grey example rows before submitting.", "p"),
    ("     • Do not rename the tabs or the 'tblSchedule' table — the importer looks for them by name.", "p"),
]
r = 2
for text, style in lines:
    c = rd.cell(row=r, column=2, value=text)
    if style=="h1":
        c.font = Font(bold=True, size=15, color=WHITE); c.fill = fill(NAVY)
        rd.row_dimensions[r].height = 26
        c.alignment = Alignment(vertical="center", indent=1)
    elif style=="h2":
        c.font = Font(bold=True, size=11, color=BLUE)
    elif style=="code":
        c.font = Font(name="Consolas", size=10, color="1B4D3E"); c.fill = fill(GREY)
    else:
        c.font = Font(size=10, color="222222")
    if style != "h1":
        c.alignment = Alignment(vertical="center", indent=1)
    r += 1

# order tabs: Schedule, README, Lists
wb.move_sheet("README", offset=-1)
wb.active = wb["Schedule"]
wb.save(OUT)
print("Saved:", OUT)
