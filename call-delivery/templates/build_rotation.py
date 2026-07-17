"""Step 3: fair-rotation assignment randomizer workbook."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

OUT = "/home/user/mygamemystory/call-delivery/templates/Call_Delivery_Assignments.xlsx"

NAVY="1F2A44"; BLUE="2E5AAC"; LTBLUE="DCE6F7"; GREY="F2F4F8"; WHITE="FFFFFF"
GOLD="C7A008"; GREEN="1B4D3E"; BORDER="B8C0D0"; INPUT_BLUE="0000FF"; YELLOW="FFF2CC"
thin=Side(style="thin",color=BORDER); box=Border(left=thin,right=thin,top=thin,bottom=thin)
def fill(h): return PatternFill("solid",fgColor=h)
F = lambda **kw: Font(name="Arial", **kw)

ROSTER = ["Nick","Ron","Jamela","Magaly","Chris","Cynthia","David","Lee","Brad"]
PERSON_COLORS = {"Nick":"DCE6F7","Ron":"FDE2CE","Jamela":"E2F0DA","Magaly":"F3DCE9",
                 "Chris":"FFF3C4","Cynthia":"D8F0F0","David":"E8E0F5","Lee":"FBD9D3","Brad":"E5E7EB"}
# (group, default staffed) — from the July Assignments sheet: N = was N/A all month
GROUPS = [("OM_All","N"),("FE","Y"),("BE","Y"),("PCO","Y"),("MOD","N"),
          ("AUTO Direct Collections","Y"),("Repo","Y"),("ARC","Y"),("CPOD / NCC","N"),
          ("Call Escalation Team","Y"),("Branch Vendor","Y"),("Sales","Y"),("Care","Y"),
          ("Card","Y"),("Branch Central","Y"),("Optional Products","Y"),
          ("CDEM-EML-STRAT","Y"),("CARE (Outbound)","N"),("West Coast Pilot","N"),
          ("Spanish","N"),("BA Monitoring","Y")]
N_DAYS = 31

wb = openpyxl.Workbook()

# ============ TEAM ============
tm = wb.active; tm.title = "Team"; tm.sheet_view.showGridLines=False
tm.sheet_properties.tabColor = BLUE
hdrs = [("A1","Team Member"),("B1","Active? (Y/N)"),("C1","Active Rank"),("E1","Active Roster (auto)")]
for cell,v in hdrs:
    tm[cell]=v; tm[cell].font=F(bold=True,color=WHITE); tm[cell].fill=fill(BLUE); tm[cell].border=box
for i,name in enumerate(ROSTER, start=2):
    tm[f"A{i}"]=name; tm[f"A{i}"].border=box; tm[f"A{i}"].font=F()
    tm[f"B{i}"]="Y"; tm[f"B{i}"].border=box
    tm[f"B{i}"].font=F(color=INPUT_BLUE,bold=True); tm[f"B{i}"].fill=fill(YELLOW)
    tm[f"B{i}"].alignment=Alignment(horizontal="center")
    tm[f"C{i}"]=f'=IF(B{i}="Y",COUNTIF($B$2:B{i},"Y"),"")'
    tm[f"C{i}"].font=F(color="6B7280"); tm[f"C{i}"].alignment=Alignment(horizontal="center")
    tm[f"E{i}"]=f'=IFERROR(INDEX($A$2:$A$10,MATCH(ROW()-1,$C$2:$C$10,0)),"")'
    tm[f"E{i}"].font=F(color="6B7280")
tm["G1"]="Active Count"; tm["G1"].font=F(bold=True,color=WHITE); tm["G1"].fill=fill(BLUE); tm["G1"].border=box
tm["G2"]='=COUNTIF(B2:B10,"Y")'; tm["G2"].font=F(bold=True); tm["G2"].border=box
tm["G2"].alignment=Alignment(horizontal="center")
tm["A12"]="Edit only the yellow Active? column. Set N to bench someone (PTO, new duties) — the whole rotation re-balances instantly."
tm["A12"].font=F(italic=True,size=9,color="6B7280")
for c,w in {"A":16,"B":14,"C":12,"E":20,"G":13}.items(): tm.column_dimensions[c].width=w

# ============ SETTINGS ============
se = wb.create_sheet("Settings"); se.sheet_view.showGridLines=False
se.sheet_properties.tabColor = GOLD
se["B2"]="Setting"; se["C2"]="Value"; se["D2"]="What it does"
for cell in ("B2","C2","D2"):
    se[cell].font=F(bold=True,color=WHITE); se[cell].fill=fill(BLUE); se[cell].border=box
rows=[("Month start date","2026-08-01","First day shown on the Rotation tab (use the 1st of the month)."),
      ("Shuffle seed",7,"THE RANDOMIZER. Any whole number. Change it (7→8→23...) to deal a completely new fair rotation."),
      ("PM offset",3,"How far the PM person is shifted from the AM person. Keep between 1 and active-count minus 1.")]
import datetime
for i,(k,v,d) in enumerate(rows, start=3):
    se[f"B{i}"]=k; se[f"B{i}"].font=F(bold=True,color=NAVY); se[f"B{i}"].border=box
    if k=="Month start date":
        se[f"C{i}"]=datetime.date(2026,8,1); se[f"C{i}"].number_format="mm/dd/yyyy"
    else:
        se[f"C{i}"]=v
    se[f"C{i}"].font=F(color=INPUT_BLUE,bold=True); se[f"C{i}"].fill=fill(YELLOW); se[f"C{i}"].border=box
    se[f"C{i}"].alignment=Alignment(horizontal="center")
    se[f"D{i}"]=d; se[f"D{i}"].font=F(size=9,color="6B7280")
for c,w in {"B":20,"C":14,"D":90}.items(): se.column_dimensions[c].width=w

# ============ ROTATION ============
ro = wb.create_sheet("Rotation"); ro.sheet_view.showGridLines=False
ro.sheet_properties.tabColor = GREEN
FIRST_COL = 3  # C; A=weekday, B=date
GROUP_ROW, STAFF_ROW, SHIFT_ROW, DATA_ROW = 2, 3, 4, 5

ro["A4"]="Day"; ro["B4"]="Date"
for cell in ("A4","B4"):
    ro[cell].font=F(bold=True,color=WHITE); ro[cell].fill=fill(NAVY); ro[cell].border=box
    ro[cell].alignment=Alignment(horizontal="center",vertical="center")
ro["B3"]="Staffed? →"; ro["B3"].font=F(bold=True,size=9,color=NAVY)
ro["B3"].alignment=Alignment(horizontal="right",vertical="center")

for g,(gname,staffed) in enumerate(GROUPS):
    c0 = FIRST_COL + g*3
    L0,L1,L2 = get_column_letter(c0), get_column_letter(c0+1), get_column_letter(c0+2)
    ro.merge_cells(start_row=GROUP_ROW,start_column=c0,end_row=GROUP_ROW,end_column=c0+2)
    h=ro.cell(row=GROUP_ROW,column=c0,value=gname)
    h.font=F(bold=True,color=WHITE,size=10); h.fill=fill(NAVY); h.border=box
    h.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    # staffed flag (editable)
    ro.merge_cells(start_row=STAFF_ROW,start_column=c0,end_row=STAFF_ROW,end_column=c0+2)
    s=ro.cell(row=STAFF_ROW,column=c0,value=staffed)
    s.font=F(color=INPUT_BLUE,bold=True); s.fill=fill(YELLOW); s.border=box
    s.alignment=Alignment(horizontal="center")
    for j,sh in enumerate(("AM","Mid-Day","PM")):
        c=ro.cell(row=SHIFT_ROW,column=c0+j,value=sh)
        c.font=F(bold=True,color=WHITE,size=9); c.fill=fill(BLUE); c.border=box
        c.alignment=Alignment(horizontal="center")
    for d in range(N_DAYS):
        r = DATA_ROW + d
        date_ref = f"$B{r}"
        staff_ref = f"{L0}${STAFF_ROW}"
        base = (f'MOD(Settings!$C$4+(DAY({date_ref})-1)+{g}+{{off}},Team!$G$2)+1')
        for j,off in enumerate(("0","0","Settings!$C$5")):
            cell = ro.cell(row=r, column=c0+j)
            cell.value=(f'=IF({staff_ref}<>"Y","N/A",'
                        f'IF(WEEKDAY({date_ref},2)>5,"",'
                        f'INDEX(Team!$E$2:$E$10,{base.format(off=off)})))')
            cell.border=box; cell.font=F(size=10)
            cell.alignment=Alignment(horizontal="center",vertical="center")
for d in range(N_DAYS):
    r=DATA_ROW+d
    ro[f"B{r}"] = "=Settings!$C$3" if d==0 else f"=B{r-1}+1"
    ro[f"B{r}"].number_format="mm/dd"; ro[f"B{r}"].font=F(bold=True,size=10)
    ro[f"B{r}"].border=box; ro[f"B{r}"].alignment=Alignment(horizontal="center")
    ro[f"A{r}"]=f'=TEXT(B{r},"ddd")'; ro[f"A{r}"].font=F(size=9,color="6B7280")
    ro[f"A{r}"].border=box; ro[f"A{r}"].alignment=Alignment(horizontal="center")
    if d%2==1:
        for c in (1,2): ro.cell(row=r,column=c).fill=fill(GREY)

last_col_letter = get_column_letter(FIRST_COL + len(GROUPS)*3 - 1)
data_range = f"C{DATA_ROW}:{last_col_letter}{DATA_ROW+N_DAYS-1}"
for name,color in PERSON_COLORS.items():
    ro.conditional_formatting.add(data_range,
        CellIsRule(operator="equal", formula=[f'"{name}"'], fill=fill(color)))
ro.column_dimensions["A"].width=6; ro.column_dimensions["B"].width=8
for g in range(len(GROUPS)):
    for j in range(3):
        ro.column_dimensions[get_column_letter(FIRST_COL+g*3+j)].width=10
ro.row_dimensions[GROUP_ROW].height=30
ro.freeze_panes = f"C{DATA_ROW}"
t=ro.cell(row=1,column=1,value="ASSIGNMENT ROTATION — generated from the Settings seed. Type over any cell to override. Weekends left blank for manual fill.")
t.font=F(bold=True,size=11,color=NAVY)

# ============ FAIRNESS CHECK ============
fc = wb.create_sheet("Fairness Check"); fc.sheet_view.showGridLines=False
fc.sheet_properties.tabColor = LTBLUE
t=fc["B2"]; t.value="FAIRNESS CHECK — assignments per person (all shifts, weekdays of the month)"
t.font=F(bold=True,size=12,color=WHITE); t.fill=fill(NAVY)
fc.merge_cells("B2:H2"); t.alignment=Alignment(vertical="center",indent=1)
fc.row_dimensions[2].height=22
HR=4
fc.cell(row=HR,column=2,value="Team Member").font=F(bold=True,color=WHITE)
fc.cell(row=HR,column=2).fill=fill(BLUE); fc.cell(row=HR,column=2).border=box
for g,(gname,_) in enumerate(GROUPS):
    c=fc.cell(row=HR,column=3+g,value=gname)
    c.font=F(bold=True,color=WHITE,size=8); c.fill=fill(BLUE); c.border=box
    c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
tc=fc.cell(row=HR,column=3+len(GROUPS),value="TOTAL")
tc.font=F(bold=True,color=WHITE); tc.fill=fill(NAVY); tc.border=box
tc.alignment=Alignment(horizontal="center",vertical="center")
fc.row_dimensions[HR].height=34
for i,name in enumerate(ROSTER):
    r=HR+1+i
    fc.cell(row=r,column=2,value=name).font=F(bold=True); fc.cell(row=r,column=2).border=box
    for g in range(len(GROUPS)):
        c0=FIRST_COL+g*3
        rng=f"Rotation!{get_column_letter(c0)}${DATA_ROW}:{get_column_letter(c0+2)}${DATA_ROW+N_DAYS-1}"
        c=fc.cell(row=r,column=3+g,value=f'=COUNTIF({rng},$B{r})')
        c.border=box; c.alignment=Alignment(horizontal="center"); c.font=F(size=10)
    c=fc.cell(row=r,column=3+len(GROUPS),
        value=f"=SUM({get_column_letter(3)}{r}:{get_column_letter(2+len(GROUPS))}{r})")
    c.font=F(bold=True); c.border=box; c.alignment=Alignment(horizontal="center"); c.fill=fill(GREY)
sp_r=HR+len(ROSTER)+2
fc.cell(row=sp_r,column=2,value="Spread (max − min, active only):").font=F(bold=True,color=NAVY)
tot_col=get_column_letter(3+len(GROUPS))
tot_rng=f"{tot_col}{HR+1}:{tot_col}{HR+len(ROSTER)}"
act_rng=f"Team!$B$2:$B$10"
c=fc.cell(row=sp_r,column=3+len(GROUPS),
    value=f'=SUMPRODUCT(MAX(({act_rng}="Y")*{tot_rng}))-SUMPRODUCT(MIN(IF({act_rng}="Y",{tot_rng})))')
c.font=F(bold=True); c.border=box; c.alignment=Alignment(horizontal="center"); c.fill=fill(YELLOW)
fc.cell(row=sp_r+1,column=2,
    value="A small spread (0–3) means the load is balanced. Re-check after changing the seed or benching someone.").font=F(italic=True,size=9,color="6B7280")
fc.column_dimensions["B"].width=14
for g in range(len(GROUPS)+1): fc.column_dimensions[get_column_letter(3+g)].width=9

# ============ README ============
rd = wb.create_sheet("README",0); rd.sheet_view.showGridLines=False
rd.sheet_properties.tabColor=GOLD
rd.column_dimensions["B"].width=105
lines=[
 ("CALL DELIVERY — ASSIGNMENT ROTATION GENERATOR","h1"),
 ("",""),
 ("Replaces hand-typing the monthly Assignments matrix. A seeded rotation deals every","p"),
 ("business group's AM / Mid-Day / PM coverage across the active team, fairly and instantly.","p"),
 ("",""),
 ("HOW TO USE (30 seconds each month)","h2"),
 ("   1.  Settings tab → set the Month start date.","p"),
 ("   2.  Settings tab → change the Shuffle seed to any new number.  That's the randomizer —","p"),
 ("        every seed deals a different, equally-fair rotation.","p"),
 ("   3.  Team tab → mark anyone out (PTO, new duties) with N.  The rotation re-balances.","p"),
 ("   4.  Rotation tab → review.  Type over any cell to make a manual swap — overrides win.","p"),
 ("   5.  Check the Fairness Check tab: the Spread number should be small (0–3).","p"),
 ("",""),
 ("RULES BUILT IN","h2"),
 ("   •  Fair rotation: everyone active gets a near-equal share, and the same person never","p"),
 ("      camps on the same group day after day — the wheel turns every day.","p"),
 ("   •  AM and Mid-Day stay with the same person (matches how the team actually runs);","p"),
 ("      PM goes to a different person (the PM offset in Settings).","p"),
 ("   •  Groups marked N in the Staffed? row show N/A (e.g., OM_All, MOD) — flip to Y anytime.","p"),
 ("   •  Weekends are left blank for manual staffing decisions.","p"),
 ("",""),
 ("EDITING LEGEND","h2"),
 ("   •  Yellow cells with blue text = yours to edit (seed, dates, Active flags, Staffed flags).","p"),
 ("   •  Everything else is formulas.  To freeze a finished month: copy the Rotation grid and","p"),
 ("      Paste Special → Values, then hand-tune freely.","p"),
 ("   •  Each person has a consistent color across the whole grid for quick visual scanning.","p"),
]
r=2
for text,style in lines:
    c=rd.cell(row=r,column=2,value=text)
    if style=="h1":
        c.font=F(bold=True,size=15,color=WHITE); c.fill=fill(NAVY)
        rd.row_dimensions[r].height=26; c.alignment=Alignment(vertical="center",indent=1)
    elif style=="h2": c.font=F(bold=True,size=11,color=BLUE)
    else: c.font=F(size=10)
    r+=1

wb.save(OUT)
print("saved", OUT)
