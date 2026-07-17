"""Build the Call Delivery MASTER workbook: Start Here, Settings, Consolidated sample."""
import openpyxl, datetime, os, sys
sys.path.insert(0, os.path.dirname(__file__))
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE="/home/user/mygamemystory/call-delivery"
INTAKE=f"{BASE}/intake/2026-07"
OUT=f"{BASE}/templates/Call_Delivery_MASTER.xlsx"

NAVY="1F2A44"; BLUE="2E5AAC"; LTBLUE="DCE6F7"; GREY="F2F4F8"; WHITE="FFFFFF"
GOLD="C7A008"; GREEN="1B4D3E"; BORDER="B8C0D0"
thin=Side(style="thin",color=BORDER); box=Border(left=thin,right=thin,top=thin,bottom=thin)
def fill(h): return PatternFill("solid",fgColor=h)

wb = openpyxl.Workbook()

# ================= START HERE =================
st = wb.active; st.title="START HERE"; st.sheet_view.showGridLines=False
st.sheet_properties.tabColor=GOLD
st.column_dimensions["B"].width=110
st.merge_cells("B2:B2")
rows = [
 ("CALL DELIVERY — MASTER WORKBOOK","h1"),
 ("One refresh pulls every business group's schedule out of the shared Intake folder.",""),
 ("",""),
 ("DAILY USE (after one-time setup)","h2"),
 ("    1.  Open this workbook.","p"),
 ("    2.  Data → Refresh All.","p"),
 ("    3.  Done — the Consolidated tab now holds every group's current schedule.","p"),
 ("",""),
 ("ONE-TIME SETUP (about 5 minutes)","h2"),
 ("    1.  Put the intake files in your shared folder (one per group per month):","p"),
 ("            \\\\CallDelivery\\Intake\\2026-07\\FE.xlsx,  BE.xlsx,  PCO.xlsx, ...","code"),
 ("    2.  Set the folder path on the Settings tab (cell C4).","p"),
 ("    3.  Data → Get Data → Launch Power Query Editor → New Source → Blank Query.","p"),
 ("    4.  Open Advanced Editor, paste the query from the process guide","p"),
 ("        (docs/Call_Delivery_Intake_Process_Guide.docx), set FolderPath, Close & Load.","p"),
 ("    5.  Load the query to the Consolidated tab, replacing the sample table there.","p"),
 ("",""),
 ("WHAT'S IN THIS WORKBOOK","h2"),
 ("    •  Settings ......... the intake folder path and current month.","p"),
 ("    •  Consolidated ..... every group's schedule in one table.  Currently pre-loaded with a","p"),
 ("                          REAL sample: all 655 July passes converted from the old workbook,","p"),
 ("                          so you can see exactly what a refresh produces.","p"),
 ("",""),
 ("WHY THIS REPLACES THE OLD TAB-PER-GROUP DESIGN","h2"),
 ("    The old master pulled each group's 'today' row through cross-sheet formulas, which is why","p"),
 ("    it accumulated 219 broken-reference errors, one dead group (Auto), and an external-workbook link.","p"),
 ("    Here, data arrives by import, not by formula — deleting or renaming a group's file can","p"),
 ("    never break another group or the rollup.","p"),
]
r=2
for text,style in rows:
    c=st.cell(row=r,column=2,value=text)
    if style=="h1":
        c.font=Font(bold=True,size=16,color=WHITE); c.fill=fill(NAVY)
        st.row_dimensions[r].height=28; c.alignment=Alignment(vertical="center",indent=1)
    elif style=="h2": c.font=Font(bold=True,size=12,color=BLUE)
    elif style=="code": c.font=Font(name="Consolas",size=10,color=GREEN); c.fill=fill(GREY)
    else: c.font=Font(size=10.5)
    r+=1

# ================= SETTINGS =================
se = wb.create_sheet("Settings"); se.sheet_view.showGridLines=False
se.sheet_properties.tabColor=BLUE
se.column_dimensions["B"].width=26; se.column_dimensions["C"].width=60
se["B2"]="Setting"; se["C2"]="Value"
for cell in ("B2","C2"):
    se[cell].font=Font(bold=True,color=WHITE); se[cell].fill=fill(BLUE); se[cell].border=box
vals=[("Intake folder root","\\\\CallDelivery\\Intake"),
      ("Current month (folder name)","2026-07"),
      ("Full folder path (used by the query)",'=C3&"\\"&C4')]
for i,(k,v) in enumerate(vals, start=3):
    se[f"B{i}"]=k; se[f"B{i}"].font=Font(bold=True,color=NAVY); se[f"B{i}"].border=box
    se[f"C{i}"]=v; se[f"C{i}"].border=box; se[f"C{i}"].font=Font(name="Consolas",size=10)
se["E3"]="Tip: in Power Query, you can bind FolderPath to this cell (right-click the Consolidated query → Edit) so a new month is just an edit here + Refresh."
se["E3"].font=Font(italic=True,size=9,color="6B7280")

# ================= CONSOLIDATED (sample = real July data) =================
co = wb.create_sheet("Consolidated"); co.sheet_view.showGridLines=False
co.sheet_properties.tabColor=GREEN
co.merge_cells("B2:J2")
t=co["B2"]; t.value="CONSOLIDATED SCHEDULE  —  sample pre-load (this becomes the Power Query output)"
t.font=Font(bold=True,size=13,color=WHITE); t.fill=fill(NAVY)
t.alignment=Alignment(vertical="center",indent=1); co.row_dimensions[2].height=24

headers=["Business Group","Date","Start Time (ET)","Timezone Wave","List ID / Campaign",
         "Priority","Segment / Tag","Special Instructions","Source File"]
HDR=4
for i,h in enumerate(headers):
    c=co.cell(row=HDR,column=2+i,value=h); c.font=Font(bold=True,color=WHITE)
    c.fill=fill(BLUE); c.border=box
    c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
co.row_dimensions[HDR].height=28

# read back the 8 intake files -> stacked rows (proves the ingest contract works)
import glob
allrows=[]
for f in sorted(glob.glob(f"{INTAKE}/*.xlsx")):
    iwb=openpyxl.load_workbook(f, data_only=False)
    sh=iwb["Schedule"]
    ref=sh.tables["tblSchedule"].ref
    first=int(ref.split(":")[0][2:] if ref.split(":")[0][1].isalpha() else 10)+1  # header row +1
    # simpler: iterate rows 10.. until no date
    grp=sh["C5"].value
    for rr in range(10, sh.max_row+1):
        d=sh.cell(row=rr,column=3).value
        if d is None: continue
        allrows.append([grp, d, sh.cell(row=rr,column=4).value, sh.cell(row=rr,column=5).value,
                        sh.cell(row=rr,column=6).value, sh.cell(row=rr,column=7).value,
                        sh.cell(row=rr,column=8).value, sh.cell(row=rr,column=9).value,
                        os.path.basename(f)])
allrows.sort(key=lambda x:(x[1], str(x[0]), x[5] if isinstance(x[5],int) else 99))

r=HDR+1
for row in allrows:
    for i,v in enumerate(row):
        c=co.cell(row=r,column=2+i,value=v); c.border=box
        if i==1: c.number_format="mm/dd/yyyy"
        if i==2: c.number_format="h:mm AM/PM"
        c.alignment=Alignment(vertical="center",
            horizontal="left" if i in (7,8) else "center", indent=1 if i in (7,8) else 0)
    r+=1
last=r-1
for col,w in zip("BCDEFGHIJ",[18,12,14,13,18,9,14,34,20]):
    co.column_dimensions[col].width=w
tbl=Table(displayName="tblConsolidated",ref=f"B{HDR}:J{last}")
tbl.tableStyleInfo=TableStyleInfo(name="TableStyleLight9",showRowStripes=True,
    showColumnStripes=False,showFirstColumn=False,showLastColumn=False)
co.add_table(tbl)
co.freeze_panes=f"B{HDR+1}"
co.auto_filter.ref=None

wb.save(OUT)
print(f"Master saved: {OUT}  |  consolidated rows: {len(allrows)}")
