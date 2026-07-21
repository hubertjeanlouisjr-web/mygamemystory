"""Step 4: redesigned Daily Rollup tab inside the master workbook.
Reads tblConsolidated via a hidden helper sheet; ranges are sized to the
actual Consolidated table so late-month dates are never dropped."""
import openpyxl, datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import Rule
from openpyxl.utils import get_column_letter

P = "/home/user/mygamemystory/call-delivery/templates/Call_Delivery_MASTER.xlsx"
NAVY="1F2A44"; BLUE="2E5AAC"; LTBLUE="DCE6F7"; GREY="F2F4F8"; WHITE="FFFFFF"
GOLD="C7A008"; GOLDL="FDF3D0"; GREEN="1B4D3E"; GREENL="DFF0E4"; BORDER="B8C0D0"
INPUT_BLUE="0000FF"; YELLOW="FFF2CC"; PURPLEL="EBE3F5"; TEALL="D8EEF0"; REDL="FBDAD5"
thin=Side(style="thin",color=BORDER); box=Border(left=thin,right=thin,top=thin,bottom=thin)
def fill(h): return PatternFill("solid",fgColor=h)
F = lambda **kw: Font(name="Arial", **kw)

wb = openpyxl.load_workbook(P)
for name in ("Daily Rollup","RollupData"):
    if name in wb.sheetnames: del wb[name]

# ---------- hidden helper (sized to the real Consolidated table) ----------
co = wb["Consolidated"]
cref = co.tables["tblConsolidated"].ref                      # e.g. "B4:J821"
N_FIRST = 5                                                  # data starts row 5 (header row 4)
N_LAST  = int(''.join(ch for ch in cref.split(":")[1] if ch.isdigit()))
CB=f"Consolidated!$B$5:$B${N_LAST}"; CC=f"Consolidated!$C$5:$C${N_LAST}"

hd = wb.create_sheet("RollupData")
hd["A1"]="key grp|date|hour"; hd["B1"]="display"; hd["C1"]="key grp|date"; hd["D1"]="note"
for i in range(N_FIRST, N_LAST+1):
    r = i - N_FIRST + 2
    hd[f"A{r}"]=(f'=IF(Consolidated!B{i}="","",Consolidated!B{i}&"|"'
                 f'&TEXT(Consolidated!C{i},"yyyymmdd")&"|"&HOUR(Consolidated!D{i}))')
    hd[f"B{r}"]=(f'=IF(Consolidated!B{i}="","",SUBSTITUTE(SUBSTITUTE(TEXT(Consolidated!D{i},"h:mm AM/PM")," AM","a")," PM","p")'
                 f'&IF(Consolidated!E{i}="",""," "&Consolidated!E{i})'
                 f'&IF(Consolidated!H{i}="",""," "&Consolidated!H{i}))')
    hd[f"C{r}"]=(f'=IF(Consolidated!B{i}="","",Consolidated!B{i}&"|"'
                 f'&TEXT(Consolidated!C{i},"yyyymmdd"))')
    hd[f"D{r}"]=f'=IF(Consolidated!I{i}="","",Consolidated!I{i})'
hd.sheet_state = "hidden"
HELP_LAST = N_LAST - N_FIRST + 2
KEY=f"RollupData!$A$2:$A${HELP_LAST}"
DISP=f"RollupData!$B$2:$B${HELP_LAST}"
KEY2=f"RollupData!$C$2:$C${HELP_LAST}"
NOTE=f"RollupData!$D$2:$D${HELP_LAST}"
GNUM_LK="'Ref Group Numbers'!$A$3:$A$40"
GNUM_VAL="'Ref Group Numbers'!$B$3:$B$40"

# ---------- Daily Rollup ----------
dr = wb.create_sheet("Daily Rollup", 0)
dr.sheet_view.showGridLines=False
dr.sheet_properties.tabColor=GOLD

dr.merge_cells("B2:S2")
t=dr["B2"]; t.value="GOOD MORNING!!!   —   Today's Call Delivery Strategy"
t.font=F(bold=True,size=18,color=WHITE); t.fill=fill(NAVY)
t.alignment=Alignment(vertical="center",indent=1); dr.row_dimensions[2].height=32

dr.merge_cells("B3:S3")
s=dr["B3"]
s.value=("ALL TIMES EASTERN.   Changes are communicated through the day — reach your Point of "
         "Contact via the Teams channel or email Call Delivery Systems.")
s.font=F(italic=True,size=10,color=NAVY); s.fill=fill(LTBLUE)
s.alignment=Alignment(vertical="center",indent=1); dr.row_dimensions[3].height=18

# controls row
dr["B5"]="Strategy Date:"; dr["B5"].font=F(bold=True,color=NAVY)
dr["B5"].alignment=Alignment(horizontal="right")
dr["C5"]=datetime.date(2026,7,17)
dr["C5"].number_format="ddd mm/dd/yyyy"
dr["C5"].font=F(bold=True,color=INPUT_BLUE,size=12); dr["C5"].fill=fill(YELLOW); dr["C5"].border=box
dr["C5"].alignment=Alignment(horizontal="center")
dr["E5"]="Auto-Messaging Voice:"; dr["E5"].font=F(bold=True,color=NAVY)
dr["E5"].alignment=Alignment(horizontal="right")
dr.merge_cells("E5:F5")
dr["G5"]='=IFERROR(INDEX(\'Ref Voice\'!$B$2:$B$40,MATCH($C$5,\'Ref Voice\'!$A$2:$A$40,0)),"")'; dr["G5"].font=F(bold=True,color=INPUT_BLUE); dr["G5"].fill=fill(YELLOW)
dr["G5"].border=box; dr["G5"].alignment=Alignment(horizontal="center")
dr.merge_cells("I5:S5")
n=dr["I5"]; n.value="Yellow cells are yours: set the date (or type =TODAY() for live use) and the voice for the day."
n.font=F(italic=True,size=9,color="6B7280"); n.alignment=Alignment(vertical="center")

# ---- schedule grid ----
GRID_HDR = 7
GROUPS = ["FE","BE","PCO","Auto","Repo","MOD","Branch Central","Branch Vendor",
          "Card","ARC","Optional Products","Call Escalation Team"]
HOURS = list(range(8,22))   # 8:00 .. 21:00
COL_GROUP, COL_GNUM, COL_PASSES = 2, 3, 4
COL_H0 = 5                            # first hour column (E)
COL_INSTR = COL_H0 + len(HOURS)       # Special Instructions (S)

c=dr.cell(row=GRID_HDR,column=COL_GROUP,value="Business Group")
c.font=F(bold=True,color=WHITE); c.fill=fill(NAVY); c.border=box
c.alignment=Alignment(horizontal="left",vertical="center",indent=1)
c=dr.cell(row=GRID_HDR,column=COL_GNUM,value="Group #'s")
c.font=F(bold=True,color=WHITE,size=9); c.fill=fill(NAVY); c.border=box
c.alignment=Alignment(horizontal="center",vertical="center")
c=dr.cell(row=GRID_HDR,column=COL_PASSES,value="Passes")
c.font=F(bold=True,color=WHITE,size=9); c.fill=fill(NAVY); c.border=box
c.alignment=Alignment(horizontal="center",vertical="center")
for j,h in enumerate(HOURS):
    c=dr.cell(row=GRID_HDR,column=COL_H0+j,value=datetime.time(h,0))
    c.number_format="h AM/PM"
    c.font=F(bold=True,color=WHITE,size=9); c.fill=fill(BLUE); c.border=box
    c.alignment=Alignment(horizontal="center",vertical="center")
c=dr.cell(row=GRID_HDR,column=COL_INSTR,value="Special Instructions")
c.font=F(bold=True,color=WHITE); c.fill=fill(NAVY); c.border=box
c.alignment=Alignment(horizontal="center",vertical="center")
dr.row_dimensions[GRID_HDR].height=20

for i,g in enumerate(GROUPS):
    r=GRID_HDR+1+i
    rowfill = GREY if i%2 else WHITE
    c=dr.cell(row=r,column=COL_GROUP,value=g)
    c.font=F(bold=True,color=NAVY); c.border=box
    c.alignment=Alignment(vertical="center",indent=1); c.fill=fill(rowfill)
    c=dr.cell(row=r,column=COL_GNUM,
        value=f'=IFERROR(INDEX({GNUM_VAL},MATCH($B{r},{GNUM_LK},0)),"")')
    c.font=F(size=9,bold=True,color=NAVY); c.border=box
    c.alignment=Alignment(horizontal="center",vertical="center"); c.fill=fill(rowfill)
    c=dr.cell(row=r,column=COL_PASSES,
        value=f'=COUNTIFS({CB},$B{r},{CC},$C$5)')
    c.font=F(size=9,color="6B7280"); c.border=box; c.alignment=Alignment(horizontal="center",vertical="center")
    for j,h in enumerate(HOURS):
        cell=dr.cell(row=r,column=COL_H0+j)
        cell.value=(f'=IFERROR(INDEX({DISP},MATCH($B{r}&"|"&TEXT($C$5,"yyyymmdd")&"|"&{h},{KEY},0)),"")')
        cell.border=box; cell.font=F(size=8)
        cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    c=dr.cell(row=r,column=COL_INSTR,
        value=f'=IFERROR(INDEX({NOTE},MATCH($B{r}&"|"&TEXT($C$5,"yyyymmdd"),{KEY2},0)),"")')
    c.border=box; c.font=F(size=9,italic=True,color=GREEN)
    c.alignment=Alignment(vertical="center",indent=1,wrap_text=True)
    dr.row_dimensions[r].height=30
LAST_GRID_ROW = GRID_HDR+len(GROUPS)

# tag colors via conditional formatting (containsText) over the hour cells
grid_rng=f"{get_column_letter(COL_H0)}{GRID_HDR+1}:{get_column_letter(COL_H0+len(HOURS)-1)}{LAST_GRID_ROW}"
def contains_rule(text, color):
    first = f"{get_column_letter(COL_H0)}{GRID_HDR+1}"
    return Rule(type="containsText", operator="containsText", text=text,
        formula=[f'NOT(ISERROR(SEARCH("{text}",{first})))'],
        dxf=openpyxl.styles.differential.DifferentialStyle(fill=fill(color)))
for text,color in [("Blitz",GOLDL),("TSC",PURPLEL),("Branch",GREENL),
                   ("Vendor",GREY),("WFO",TEALL),("AMs",LTBLUE)]:
    dr.conditional_formatting.add(grid_rng, contains_rule(text,color))

# legend
lr = LAST_GRID_ROW+2
dr.cell(row=lr,column=2,value="Legend:").font=F(bold=True,size=9,color=NAVY)
for k,(text,color) in enumerate([("AMs (auto-msg)",LTBLUE),("Branch",GREENL),("WFO",TEALL),
                                  ("Blitz",GOLDL),("TSC",PURPLEL),("Vendor",GREY)]):
    c=dr.cell(row=lr,column=4+k*2)
    dr.merge_cells(start_row=lr,start_column=4+k*2,end_row=lr,end_column=5+k*2)
    c.value=text; c.fill=fill(color); c.border=box
    c.font=F(size=8); c.alignment=Alignment(horizontal="center")

# ---- coverage + hours section ----
cr = lr+2
dr.merge_cells(start_row=cr,start_column=2,end_row=cr,end_column=8)
c=dr.cell(row=cr,column=2,value="CD TEAM COVERAGE  (fill from the Assignments workbook)")
c.font=F(bold=True,size=11,color=WHITE); c.fill=fill(NAVY)
c.alignment=Alignment(vertical="center",indent=1); dr.row_dimensions[cr].height=20
hdr_r=cr+1
for k,h in enumerate(["Area","AM","Mid-Day","PM"]):
    c=dr.cell(row=hdr_r,column=2+k,value=h)
    c.font=F(bold=True,color=WHITE,size=9); c.fill=fill(BLUE); c.border=box
    c.alignment=Alignment(horizontal="center")
for k,area in enumerate(["Central Collections","Inbound (Sales / Care)","Card","Escalations / NICE"]):
    r=hdr_r+1+k
    c=dr.cell(row=r,column=2,value=area); c.font=F(bold=True,size=9,color=NAVY); c.border=box
    c.alignment=Alignment(indent=1)
    for kk in range(3):
        c=dr.cell(row=r,column=3+kk); c.border=box; c.fill=fill(YELLOW)
        c.font=F(color=INPUT_BLUE,size=9); c.alignment=Alignment(horizontal="center")

hr_c = 10
dr.merge_cells(start_row=cr,start_column=hr_c,end_row=cr,end_column=hr_c+3)
c=dr.cell(row=cr,column=hr_c,value="HOURS OF OPERATION")
c.font=F(bold=True,size=11,color=WHITE); c.fill=fill(NAVY)
c.alignment=Alignment(vertical="center",indent=1)
for k,h in enumerate(["Area","Groups","Open","Close"]):
    c=dr.cell(row=hdr_r,column=hr_c+k,value=h)
    c.font=F(bold=True,color=WHITE,size=9); c.fill=fill(BLUE); c.border=box
    c.alignment=Alignment(horizontal="center")
for k,(area,grps,o,cl) in enumerate([
        ("Card Collections","40 / 80",'=IFERROR(INDEX(\'Ref Card Hours\'!$B$2:$B$40,MATCH(\'Daily Rollup\'!$C$5,\'Ref Card Hours\'!$A$2:$A$40,0)),"")','=IFERROR(INDEX(\'Ref Card Hours\'!$C$2:$C$40,MATCH(\'Daily Rollup\'!$C$5,\'Ref Card Hours\'!$A$2:$A$40,0)),"")'),
        ("Call Escalation Team","37 / 77","8:00 AM","5:00 PM"),
        ("Optional Products","130 / 131","8:00 AM","6:00 PM")]):
    r=hdr_r+1+k
    vals=[area,grps,o,cl]
    for kk,v in enumerate(vals):
        c=dr.cell(row=r,column=hr_c+kk,value=v); c.border=box
        c.font=F(size=9, color=INPUT_BLUE if kk>=2 else "000000")
        if kk>=2: c.fill=fill(YELLOW)
        c.alignment=Alignment(horizontal="center" if kk else "left",indent=0 if kk else 1)

fr = hdr_r+6
dr.merge_cells(start_row=fr,start_column=2,end_row=fr,end_column=19)
c=dr.cell(row=fr,column=2,
    value="This page rebuilds itself from the Consolidated tab — refresh the import, pick a date, and it's ready to send. No cross-sheet formula web, nothing to break.")
c.font=F(italic=True,size=9,color="6B7280")

dr.column_dimensions["B"].width=17
dr.column_dimensions[get_column_letter(COL_GNUM)].width=12
dr.column_dimensions[get_column_letter(COL_PASSES)].width=7
for j in range(len(HOURS)):
    dr.column_dimensions[get_column_letter(COL_H0+j)].width=11
dr.column_dimensions[get_column_letter(COL_INSTR)].width=30
dr.freeze_panes=f"{get_column_letter(COL_H0)}{GRID_HDR+1}"

wb.save(P)
print(f"Daily Rollup added (Group # column + dynamic ranges to row {N_LAST})")
