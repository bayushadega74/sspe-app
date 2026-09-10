"""
Build "Level & Jadwal Coach" workbook for Twenty Swim (Bandung).

Input : data.json  (students pulled from Database_Siswa_Twenty_SSI.xlsx, see extract step)
Output: Level_Jadwal_Coach_Twenty_Swim.xlsx

Tab order: DASHBOARD, Member Aktif, Panduan Level SSI, [30 coach tabs A-Z],
           DataMember (hidden), Data Dashboard (hidden), Daftar (hidden)
"""
import json, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import DoughnutChart, BarChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, Font as DFont, RichTextProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.workbook.properties import CalcProperties

SRC = sys.argv[1] if len(sys.argv) > 1 else "data.json"
OUT = sys.argv[2] if len(sys.argv) > 2 else "Level_Jadwal_Coach_Twenty_Swim.xlsx"
J = json.load(open(SRC, encoding="utf-8"))
COACHES = sorted(J["coaches"], key=str.lower)
DATA = J["data"]

# ---------------------------------------------------------------- palette
NAVY, BLUE = "1F2A44", "159BD3"
GRAY, LINE = "6B7280", "D9DDE6"
LBLUE, LORANGE = "EAF3FA", "FFF4E0"
GREEN_BG, GREEN_TX = "E8F3EC", "1E6B3C"
RED_BG, RED_TX = "FBEAE8", "A03B3B"
YEL_BG, YEL_TX = "FDF3D8", "8A6D1A"
PINK_BG = "FCE4EC"
BLUE_BG, BLUE_TX = "E9F0FA", "1D4F91"
GREEN_STRONG, RED_STRONG = "C8E6C9", "FFCDD2"

FONT = "Arial"
def font(sz=10, b=False, color="000000", i=False):
    return Font(name=FONT, size=sz, bold=b, color=color, italic=i)
def fill(rgb):
    return PatternFill("solid", fgColor=rgb)
thin = Side(style="thin", color=LINE)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")
LEFT_WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)

# ---------------------------------------------------------------- lists
LEVELS = ["Baby Swim I", "Baby Swim II", "Aquatike I", "Aquatike II", "Aquatike III",
          "Preschool I", "Preschool II", "Beginner I", "Beginner II", "Beginner III",
          "Intermediate I", "Intermediate II", "Advanced I", "Advanced II", "Swim Team Prep"]
PROGRAMS = ["Private Class", "Couple Class", "Group Class", "Reguler Class", "Baby Swim", "Hydrotherapy Class"]
LOKASI = ["Bandung Kota", "Bandung Utara", "Bandung Selatan", "Bandung Timur", "Bandung Barat", "Cimahi"]
KOLAM = ["BSD Bandung (Riung Mekar)", "Calma Kitchen & Pool", "Cherryfield Club House",
         "Hotel Cozy Catellya Boutique", "Hotel Grand Arjuna Bandung", "Hotel Grand Pacific",
         "Hotel Kytos Setiabudi", "Hotel Neo Dipatiukur", "Hotel Oakwood Merdeka Bandung",
         "Hotel Point C Hotel Bandung", "Hotel Savoy Homan", "Kolam Renang Buah Batu Regency",
         "Kolam Renang Bukit Cipaku", "Kolam Renang Rumah Oma Opa", "Kolam Renang Saraga ITB",
         "Kolam Renang Tirta Mulya Cigending", "Kolam Renang Tongkeng", "Kolam Renang UPI (Gelanggang UPI)",
         "Kolam Renang Wika Bandung", "Pandiga Sport Club", "Sampoerna Sport Club",
         "Sari Ater Kamboti Bandung", "Villa Salse (Dago Giri)", "Lainnya"]
RISIKO = ["Rendah", "Sedang", "Tinggi"]
STATUS = ["PA", "Standalone"]
SOSIAL = ["Sudah", "Belum"]

# coach-tab geometry
R0, R1 = 5, 54            # student rows (50)
SCH_HDR = 5               # schedule header row
DAYS = [("Senin", 6, 10), ("Selasa", 11, 15), ("Rabu", 16, 20), ("Kamis", 21, 25),
        ("Jumat", 26, 30), ("Sabtu", 31, 40), ("Minggu", 41, 50)]
SCH_R0, SCH_R1 = 6, 50

def q(name):  # quoted sheet ref
    return "'" + name.replace("'", "''") + "'"

def xsum(fn, col, crit=None):
    """Sum of COUNTA/COUNTIF over all 30 coach tabs."""
    parts = []
    for c in COACHES:
        rng = f"{q(c)}!${col}${R0}:${col}${R1}"
        parts.append(f"COUNTA({rng})" if fn == "COUNTA" else f"COUNTIF({rng},{crit})")
    return "=" + "+".join(parts)

def xsum_sched(col, crit):
    return "=" + "+".join(f"COUNTIF({q(c)}!${col}${SCH_R0}:${col}${SCH_R1},{crit})" for c in COACHES)

wb = Workbook()
wb.remove(wb.active)

# ======================================================================
# Daftar (hidden) - dropdown sources
# ======================================================================
wsD = wb.create_sheet("Daftar")
def put_list(ws, col, header, items):
    ws[f"{col}1"] = header; ws[f"{col}1"].font = font(10, True)
    for i, v in enumerate(items, start=2):
        ws[f"{col}{i}"] = v; ws[f"{col}{i}"].font = font(10)
put_list(wsD, "A", "Kolam", KOLAM)
put_list(wsD, "C", "Lokasi", LOKASI)
put_list(wsD, "E", "Level", LEVELS)
put_list(wsD, "G", "Program", PROGRAMS)
put_list(wsD, "I", "Coach", COACHES)
put_list(wsD, "K", "Status", STATUS)
put_list(wsD, "M", "Risiko", RISIKO)
put_list(wsD, "O", "Sosialisasi", SOSIAL)
for col, w in {"A": 34, "C": 18, "E": 18, "G": 20, "I": 30, "K": 12, "M": 10, "O": 12}.items():
    wsD.column_dimensions[col].width = w
REF_KOLAM = f"Daftar!$A$2:$A${1+len(KOLAM)}"
REF_LOK = f"Daftar!$C$2:$C${1+len(LOKASI)}"
REF_LVL = f"Daftar!$E$2:$E${1+len(LEVELS)}"
REF_PROG = f"Daftar!$G$2:$G${1+len(PROGRAMS)}"
REF_COACH = f"Daftar!$I$2:$I${1+len(COACHES)}"
REF_STAT = f"Daftar!$K$2:$K${1+len(STATUS)}"
REF_RISK = f"Daftar!$M$2:$M${1+len(RISIKO)}"
REF_SOS = f"Daftar!$O$2:$O${1+len(SOSIAL)}"

# ======================================================================
# Coach tabs
# ======================================================================
def build_coach(name):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 100
    widths = {"A": 5, "B": 30, "C": 19, "D": 17, "E": 12, "F": 11, "G": 27, "H": 3, "I": 3,
              "J": 10, "K": 13, "L": 28, "M": 18, "N": 32}
    for c, w in widths.items():
        ws.column_dimensions[c].width = w
    ws.column_dimensions["H"].hidden = True

    ws["A1"] = "TWENTY SWIM  -  LEVEL SISWA & JADWAL LATIHAN (BANDUNG)"
    ws["A1"].font = font(14, True, NAVY)
    ws["A2"] = f"Pelatih: {name}"
    ws["A2"].font = font(11, True, BLUE)
    ws["D2"] = "Terisi:"; ws["D2"].font = font(9, False, GRAY); ws["D2"].alignment = Alignment(horizontal="right")
    ws["E2"] = f"=COUNTA($B${R0}:$B${R1})"; ws["E2"].font = font(10, True, NAVY)
    ws["F2"] = "siswa"; ws["F2"].font = font(9, False, GRAY)
    ws["A3"] = "Isi kolom putih. Kolom berwarna = dropdown, tinggal pilih. No & Coach terisi otomatis."
    ws["A3"].font = font(8, False, GRAY, i=True)
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[3].height = 13
    ws.row_dimensions[4].height = 30

    heads = ["No", "Nama Siswa", "Program/Kelas", "Level Saat Ini", "Status", "Risiko",
             "Sosialisasi Keselamatan ke Ortu", "Coach"]
    for i, h in enumerate(heads, start=1):
        c = ws.cell(4, i, h)
        c.font = font(10, True, "FFFFFF"); c.fill = fill(NAVY); c.alignment = CENTER; c.border = BORDER

    rows = DATA.get(name, [])
    for r in range(R0, R1 + 1):
        ws.cell(r, 1, f'=IF($B{r}<>"",ROW()-4,"")')
        ws.cell(r, 8, f'=IF($B{r}<>"","{name}","")')
        idx = r - R0
        if idx < len(rows):
            nm, prog, lvl, st = rows[idx]
            ws.cell(r, 2, nm)
            if prog: ws.cell(r, 3, prog)
            if lvl: ws.cell(r, 4, lvl)
            if st: ws.cell(r, 5, st)
        for cidx in range(1, 9):
            c = ws.cell(r, cidx)
            c.font = font(10); c.border = BORDER
            c.alignment = CENTER if cidx in (1, 5, 6, 7) else LEFT
        ws.row_dimensions[r].height = 17

    # dropdowns (long lists reference Daftar)
    dvs = [
        (f"C{R0}:C{R1}", f"={REF_PROG}"), (f"D{R0}:D{R1}", f"={REF_LVL}"),
        (f"E{R0}:E{R1}", f"={REF_STAT}"), (f"F{R0}:F{R1}", f"={REF_RISK}"),
        (f"G{R0}:G{R1}", f"={REF_SOS}"),
        (f"M{SCH_R0}:M{SCH_R1}", f"={REF_LOK}"), (f"N{SCH_R0}:N{SCH_R1}", f"={REF_KOLAM}"),
    ]
    for rng, f in dvs:
        dv = DataValidation(type="list", formula1=f, allow_blank=True, showDropDown=False,
                            showErrorMessage=True, errorTitle="Pilihan tidak valid",
                            error="Silakan pilih dari daftar dropdown.")
        ws.add_data_validation(dv); dv.add(rng)
    # schedule student name: suggest from this tab's list, but allow free typing
    dv = DataValidation(type="list", formula1=f"=$B${R0}:$B${R1}", allow_blank=True, showErrorMessage=False)
    ws.add_data_validation(dv); dv.add(f"L{SCH_R0}:L{SCH_R1}")

    # conditional colours
    rngC, rngF, rngG = f"C{R0}:C{R1}", f"F{R0}:F{R1}", f"G{R0}:G{R1}"
    for p in ["Private Class", "Couple Class", "Group Class", "Reguler Class"]:
        ws.conditional_formatting.add(rngC, FormulaRule(formula=[f'$C{R0}="{p}"'], fill=fill(LBLUE)))
    for p in ["Baby Swim", "Hydrotherapy Class"]:
        ws.conditional_formatting.add(rngC, FormulaRule(formula=[f'$C{R0}="{p}"'], fill=fill(PINK_BG)))
    ws.conditional_formatting.add(rngF, FormulaRule(formula=[f'$F{R0}="Rendah"'], fill=fill(GREEN_BG), font=Font(color=GREEN_TX)))
    ws.conditional_formatting.add(rngF, FormulaRule(formula=[f'$F{R0}="Sedang"'], fill=fill(YEL_BG), font=Font(color=YEL_TX)))
    ws.conditional_formatting.add(rngF, FormulaRule(formula=[f'$F{R0}="Tinggi"'], fill=fill(RED_BG), font=Font(color=RED_TX)))
    ws.conditional_formatting.add(rngG, FormulaRule(formula=[f'$G{R0}="Sudah"'], fill=fill(GREEN_STRONG), font=Font(color=GREEN_TX, bold=True)))
    ws.conditional_formatting.add(rngG, FormulaRule(formula=[f'$G{R0}="Belum"'], fill=fill(RED_STRONG), font=Font(color=RED_TX, bold=True)))

    # ---- schedule table (right)
    ws["J4"] = "JADWAL LATIHAN  (Senin-Jumat 5 baris, Sabtu-Minggu 10 baris)"
    ws["J4"].font = font(10, True, NAVY); ws["J4"].alignment = Alignment(vertical="center")
    for i, h in enumerate(["Hari", "Jam", "Nama Siswa", "Lokasi (wilayah)", "Kolam"], start=10):
        c = ws.cell(SCH_HDR, i, h)
        c.font = font(10, True, "FFFFFF"); c.fill = fill(BLUE); c.alignment = CENTER; c.border = BORDER
    for day, a, b in DAYS:
        weekend = day in ("Sabtu", "Minggu")
        ws.merge_cells(start_row=a, start_column=10, end_row=b, end_column=10)
        c = ws.cell(a, 10, day)
        c.font = font(10, True, NAVY); c.alignment = CENTER
        for r in range(a, b + 1):
            for cidx in range(10, 15):
                cc = ws.cell(r, cidx)
                cc.border = BORDER; cc.font = font(10)
                if cidx == 10:
                    cc.fill = fill(LORANGE if weekend else LBLUE)
                elif cidx == 11:
                    cc.alignment = CENTER; cc.number_format = "@"
                else:
                    cc.alignment = LEFT
    ws.freeze_panes = "B5"
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws

for c in COACHES:
    build_coach(c)

# ======================================================================
# DataMember (hidden) - consolidated list via SORT+FILTER (Google Sheets)
# ======================================================================
wsM = wb.create_sheet("DataMember")
for i, h in enumerate(["Nama", "Coach", "Program", "Level"], start=1):
    c = wsM.cell(1, i, h); c.font = font(10, True, "FFFFFF"); c.fill = fill(NAVY)
blocks = ";".join(
    ",".join(f"{q(c)}!${col}${R0}:${col}${R1}" for col in "BHCD") for c in COACHES)
names = ";".join(f"{q(c)}!$B${R0}:$B${R1}" for c in COACHES)
if "--no-array" not in sys.argv:
    wsM["A2"] = f'=IFERROR(SORT(FILTER({{{blocks}}},{{{names}}}<>""),1,TRUE),"")'
wsM["F1"] = "Rumus di A2 menggabungkan 30 tab coach (SORT+FILTER, Google Sheets). Jangan diubah."
wsM["F1"].font = font(9, False, GRAY, i=True)
for col, w in {"A": 34, "B": 30, "C": 20, "D": 18}.items():
    wsM.column_dimensions[col].width = w

# ======================================================================
# Data Dashboard (hidden) - chart sources
# ======================================================================
wsX = wb.create_sheet("Data Dashboard")
wsX["A1"] = "SUMBER DATA DASHBOARD - otomatis dari 30 tab coach, jangan diubah manual"
wsX["A1"].font = font(10, True, NAVY)

def block(ws, col, row, title, items, fmls, val_hdr="Siswa"):
    ci = ord(col) - 64
    h1 = ws.cell(row, ci, title); h2 = ws.cell(row, ci + 1, val_hdr)
    for h in (h1, h2): h.font = font(10, True, "FFFFFF"); h.fill = fill(NAVY)
    for i, (it, f) in enumerate(zip(items, fmls), start=1):
        ws.cell(row + i, ci, it).font = font(10)
        ws.cell(row + i, ci + 1, f).font = font(10)
    return row + 1, row + len(items)

# Level (A:B)  rows 4..18 + (Belum diisi) row 19
lv0, lv1 = block(wsX, "A", 3, "Level SSI", LEVELS, [xsum("COUNTIF", "D", f"$A{4+i}") for i in range(len(LEVELS))])
wsX.cell(lv1 + 1, 1, "(Belum diisi)").font = font(10, False, GRAY)
wsX.cell(lv1 + 1, 2, f"=B{lv1+3}-SUM(B{lv0}:B{lv1})").font = font(10)
wsX.cell(lv1 + 3, 1, "Total siswa").font = font(10, True)
wsX.cell(lv1 + 3, 2, xsum("COUNTA", "B")).font = font(10, True)
TOTAL_CELL = f"'Data Dashboard'!$B${lv1+3}"

# Program (D:E)
pg0, pg1 = block(wsX, "D", 3, "Program/Kelas", PROGRAMS, [xsum("COUNTIF", "C", f"$D{4+i}") for i in range(len(PROGRAMS))])
wsX.cell(pg1 + 1, 4, "(Belum diisi)").font = font(10, False, GRAY)
wsX.cell(pg1 + 1, 5, f"=B{lv1+3}-SUM(E{pg0}:E{pg1})").font = font(10)

# Risiko (G:H)
rk0, rk1 = block(wsX, "G", 3, "Risiko", RISIKO, [xsum("COUNTIF", "F", f"$G{4+i}") for i in range(len(RISIKO))])
wsX.cell(rk1 + 1, 7, "(Belum diisi)").font = font(10, False, GRAY)
wsX.cell(rk1 + 1, 8, f"=B{lv1+3}-SUM(H{rk0}:H{rk1})").font = font(10)

# Sosialisasi (G:H lower)
so_row = rk1 + 4
block(wsX, "G", so_row, "Sosialisasi", SOSIAL, [xsum("COUNTIF", "G", f"$G{so_row+1+i}") for i in range(len(SOSIAL))])
wsX.cell(so_row + 3, 7, "(Belum diisi)").font = font(10, False, GRAY)
wsX.cell(so_row + 3, 8, f"=B{lv1+3}-SUM(H{so_row+1}:H{so_row+2})").font = font(10)
SUDAH_CELL = f"'Data Dashboard'!$H${so_row+1}"; BELUM_CELL = f"'Data Dashboard'!$H${so_row+2}"

# Status PA/Standalone (G:H lower 2)
st_row = so_row + 6
block(wsX, "G", st_row, "Status", STATUS, [xsum("COUNTIF", "E", f"$G{st_row+1+i}") for i in range(len(STATUS))])

# Lokasi (J:K) from schedule
lo0, lo1 = block(wsX, "J", 3, "Lokasi (jadwal)", LOKASI, [xsum_sched("M", f"$J{4+i}") for i in range(len(LOKASI))], "Sesi")
# Kolam (M:N) from schedule
ko0, ko1 = block(wsX, "M", 3, "Kolam (jadwal)", KOLAM, [xsum_sched("N", f"$M{4+i}") for i in range(len(KOLAM))], "Sesi")

# Coach (P:R) + Top 8 (T:V)
co0, co1 = block(wsX, "P", 3, "Coach (semua)", COACHES, [f"=COUNTA({q(c)}!$B${R0}:$B${R1})" for c in COACHES])
wsX["R3"] = "skor"; wsX["R3"].font = font(10, True, "FFFFFF"); wsX["R3"].fill = fill(NAVY)
for r in range(co0, co1 + 1):
    wsX.cell(r, 18, f"=Q{r}+ROW()/100000").font = font(9, False, GRAY)
for i, h in enumerate(["Top 8 Coach", "Siswa", "skor"], start=20):
    c = wsX.cell(3, i, h); c.font = font(10, True, "FFFFFF"); c.fill = fill(NAVY)
for k in range(1, 9):
    r = 3 + k
    wsX.cell(r, 22, f"=LARGE($R${co0}:$R${co1},{k})").font = font(9, False, GRAY)
    wsX.cell(r, 20, f"=INDEX($P${co0}:$P${co1},MATCH(V{r},$R${co0}:$R${co1},0))").font = font(10)
    wsX.cell(r, 21, f"=INDEX($Q${co0}:$Q${co1},MATCH(V{r},$R${co0}:$R${co1},0))").font = font(10)
COACH_COUNT_CELL = f"COUNTA({REF_COACH})"
for col, w in {"A": 18, "B": 8, "D": 20, "E": 8, "G": 14, "H": 8, "J": 18, "K": 8, "M": 34, "N": 8,
               "P": 30, "Q": 8, "R": 10, "T": 30, "U": 8, "V": 10}.items():
    wsX.column_dimensions[col].width = w

# ======================================================================
# DASHBOARD
# ======================================================================
ws = wb.create_sheet("DASHBOARD", 0)
ws.sheet_view.showGridLines = False
ws.sheet_view.zoomScale = 100
CARD_COLS = [("B", "C", "D"), ("F", "G", "H"), ("J", "K", "L"), ("N", "O", "P")]
ws.column_dimensions["A"].width = 2.2
for grp in CARD_COLS:
    for c in grp: ws.column_dimensions[c].width = 9.6
for c in ("E", "I", "M"): ws.column_dimensions[c].width = 1.8
ws.column_dimensions["Q"].width = 2.2
ws.row_dimensions[1].height = 7.5
ws.row_dimensions[2].height = 36
ws.row_dimensions[3].height = 15
ws.row_dimensions[4].height = 9.75
ws.row_dimensions[5].height = 18
ws.row_dimensions[6].height = 21
ws.row_dimensions[7].height = 21
ws.row_dimensions[8].height = 15
ws.row_dimensions[9].height = 9.75

ws.merge_cells("B2:P2")
ws["B2"] = "TWENTY SWIM  -  DASHBOARD LEVEL & JADWAL COACH (BANDUNG)"
ws["B2"].font = font(17, True, "FFFFFF"); ws["B2"].fill = fill(NAVY)
ws["B2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws.merge_cells("B3:P3")
ws["B3"] = "Semua angka dihitung otomatis dari 30 tab coach - tambah/ubah siswa di tab coach, dashboard ikut berubah."
ws["B3"].font = font(9, False, GRAY); ws["B3"].alignment = LEFT

cards = [
    ("TOTAL SISWA", xsum("COUNTA", "B"), "nama terisi di 30 tab coach", BLUE_BG, BLUE_TX),
    ("JUMLAH COACH", f"={COACH_COUNT_CELL}", "tab coach aktif", "F0F1F4", "3D4A66"),
    ("SOSIALISASI SUDAH", xsum("COUNTIF", "G", '"Sudah"'), "keselamatan ke ortu: Sudah", GREEN_BG, GREEN_TX),
    ("SOSIALISASI BELUM", xsum("COUNTIF", "G", '"Belum"'), "keselamatan ke ortu: Belum", RED_BG, RED_TX),
]
for (c1, c2, c3), (title, fml, sub, bg, tx) in zip(CARD_COLS, cards):
    ws.merge_cells(f"{c1}5:{c3}5"); ws.merge_cells(f"{c1}6:{c3}7"); ws.merge_cells(f"{c1}8:{c3}8")
    ws[f"{c1}5"] = title; ws[f"{c1}5"].font = font(9, True, tx)
    ws[f"{c1}6"] = fml; ws[f"{c1}6"].font = font(24, True, NAVY); ws[f"{c1}6"].number_format = "#,##0"
    ws[f"{c1}8"] = sub; ws[f"{c1}8"].font = font(8, False, GRAY)
    for r in range(5, 9):
        for cc in (c1, c2, c3):
            cell = ws[f"{cc}{r}"]
            cell.fill = fill(bg); cell.alignment = CENTER; cell.border = BORDER

# ---- chart helpers
def rich(sz, b=False, color="000000"):
    cp = CharacterProperties(sz=sz, b=b, solidFill=color, latin=DFont(typeface=FONT))
    return RichText(bodyPr=RichTextProperties(), p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])

def style_chart(ch, title, w, h):
    ch.title = title
    ch.title.tx.rich.p[0].pPr = ParagraphProperties(defRPr=CharacterProperties(sz=1100, b=True, solidFill=NAVY, latin=DFont(typeface=FONT)))
    for r in ch.title.tx.rich.p[0].r:
        r.rPr = CharacterProperties(sz=1100, b=True, solidFill=NAVY, latin=DFont(typeface=FONT))
    ch.title.overlay = False
    ch.width, ch.height = w, h
    ch.graphical_properties = GraphicalProperties(solidFill="FFFFFF")
    ch.graphical_properties.line = LineProperties(solidFill="D9D9D9", w=9525)
    ch.legend = None

def labels(numfmt="#,##0", pos=None):
    d = DataLabelList()
    d.showVal = True; d.showSerName = False; d.showCatName = False; d.showLegendKey = False; d.showPercent = False
    d.numFmt = numfmt
    d.txPr = rich(900, False, "1F2A44")
    if pos: d.position = pos
    return d

def axes(ch, horizontal=False):
    ch.x_axis.delete = False; ch.y_axis.delete = False
    ch.x_axis.txPr = rich(900); ch.y_axis.txPr = rich(900)
    ch.x_axis.majorTickMark = "none"; ch.y_axis.majorTickMark = "none"
    ch.x_axis.graphicalProperties = GraphicalProperties(); ch.x_axis.graphicalProperties.line = LineProperties(solidFill="B3B3B3")
    ch.y_axis.graphicalProperties = GraphicalProperties(); ch.y_axis.graphicalProperties.line = LineProperties(solidFill="B3B3B3")
    ch.y_axis.majorGridlines.spPr = GraphicalProperties(); ch.y_axis.majorGridlines.spPr.line = LineProperties(solidFill="E5E7EB")
    ch.y_axis.number_format = "#,##0"
    ch.y_axis.scaling.min = 0
    if horizontal:
        ch.x_axis.scaling.orientation = "maxMin"   # first category on top
        ch.y_axis.crosses = "max"                   # keep value axis at the bottom
        ch.x_axis.axPos = "l"; ch.y_axis.axPos = "b"

def bar(title, cats, vals, color, horizontal=False, w=11.6, h=7.6, gap=60, colors=None):
    ch = BarChart()
    ch.type = "bar" if horizontal else "col"
    ch.grouping = "clustered"; ch.gapWidth = gap; ch.overlap = 0
    ch.add_data(vals, titles_from_data=True); ch.set_categories(cats)
    s = ch.series[0]
    s.graphicalProperties = GraphicalProperties(solidFill=color)
    s.graphicalProperties.line = LineProperties(noFill=True)
    if colors:
        for i, col in enumerate(colors):
            dp = DataPoint(idx=i); dp.graphicalProperties = GraphicalProperties(solidFill=col)
            dp.graphicalProperties.line = LineProperties(noFill=True); s.dPt.append(dp)
    ch.dataLabels = labels(numfmt="#,##0;;", pos="outEnd")   # hide zero labels
    style_chart(ch, title, w, h); axes(ch, horizontal)
    return ch

X = wsX.title
def ref(col, r0, r1, hdr=True):
    ci = ord(col) - 64
    return Reference(wsX, min_col=ci, min_row=(r0 - 1) if hdr else r0, max_row=r1)

# Donut: level distribution (15 levels + belum diisi)
dn = DoughnutChart(holeSize=55)
dn.add_data(ref("B", lv0, lv1 + 1), titles_from_data=True)
dn.set_categories(ref("A", lv0, lv1 + 1, hdr=False))
PALETTE = ["1F2A44", "2E3D66", "3D5A99", "159BD3", "4FB8E3", "8AD0EC", "1E6B3C", "3E9D63", "7DC38F",
           "C8A24B", "E0BE6E", "F0D9A0", "C25450", "D98380", "EAB1AE", "9CA3AF"]
s = dn.series[0]
for i, col in enumerate(PALETTE):
    dp = DataPoint(idx=i); dp.graphicalProperties = GraphicalProperties(solidFill=col)
    dp.graphicalProperties.line = LineProperties(solidFill="FFFFFF", w=6350); s.dPt.append(dp)
dn.dataLabels = labels(numfmt="#,##0;;")   # hide zero slices
style_chart(dn, "Sebaran Level SSI", 11.6, 7.6)
from openpyxl.chart.legend import Legend
dn.legend = Legend(); dn.legend.position = "r"; dn.legend.overlay = False; dn.legend.txPr = rich(800)
ws.add_chart(dn, "B11")

top8 = bar("8 Coach dengan Siswa Terbanyak", ref("T", 4, 11, hdr=False), ref("U", 4, 11), NAVY, horizontal=True)
ws.add_chart(top8, "J11")

prog = bar("Siswa per Program/Kelas", ref("D", pg0, pg1, hdr=False), ref("E", pg0, pg1), BLUE)
ws.add_chart(prog, "B28")

risk = bar("Siswa per Risiko", ref("G", rk0, rk1, hdr=False), ref("H", rk0, rk1), NAVY,
           gap=110, colors=["3E9D63", "C8A24B", "C25450"])
ws.add_chart(risk, "J28")

lok = bar("Sesi Latihan per Lokasi (dari jadwal)", ref("J", lo0, lo1, hdr=False), ref("K", lo0, lo1), NAVY, h=11.0)
ws.add_chart(lok, "B45")

kol = bar("Sesi Latihan per Kolam (dari jadwal)", ref("M", ko0, ko1, hdr=False), ref("N", ko0, ko1), BLUE, horizontal=True, h=11.0, gap=40)
ws.add_chart(kol, "J45")

def section(row, text):
    ws.merge_cells(f"B{row}:P{row}")
    ws[f"B{row}"] = text; ws[f"B{row}"].font = font(10, True, NAVY)
    ws[f"B{row}"].alignment = Alignment(vertical="center")
    ws[f"B{row}"].border = Border(bottom=Side(style="thin", color=BLUE))
    ws.row_dimensions[row].height = 18
section(10, "SEBARAN SISWA")
section(27, "PROGRAM & RISIKO")
section(44, "LOKASI & KOLAM (DARI TABEL JADWAL)")

# small numeric summary under the charts (aggregate only)
r = 68
ws[f"B{r}"] = "RINGKASAN ANGKA"; ws[f"B{r}"].font = font(10, True, NAVY)
ws.merge_cells(f"B{r}:P{r}"); ws[f"B{r}"].border = Border(bottom=Side(style="thin", color=BLUE)); ws.row_dimensions[r].height = 18
summ = [("Status PA", f"='{X}'!H{st_row+1}"), ("Status Standalone", f"='{X}'!H{st_row+2}"),
        ("Level belum diisi", f"='{X}'!B{lv1+1}"), ("Program belum diisi", f"='{X}'!E{pg1+1}"),
        ("Risiko belum diisi", f"='{X}'!H{rk1+1}"), ("Sosialisasi belum diisi", f"='{X}'!H{so_row+3}")]
for i, (lab, f) in enumerate(summ):
    rr = r + 1 + i % 3
    c_lab, c_val = ("B", "D") if i < 3 else ("J", "L")
    ws[f"{c_lab}{rr}"] = lab; ws[f"{c_lab}{rr}"].font = font(9, False, GRAY)
    ws[f"{c_val}{rr}"] = f; ws[f"{c_val}{rr}"].font = font(10, True, NAVY); ws[f"{c_val}{rr}"].alignment = Alignment(horizontal="right")
    ws.row_dimensions[rr].height = 15

ws["B74"] = "Privasi: dashboard hanya menampilkan angka agregat - tanpa nama siswa."
ws["B74"].font = font(8, False, GRAY, i=True)
ws["B75"] = "Catatan: grafik Lokasi & Kolam menghitung baris jadwal yang diisi di tabel kanan tiap tab coach; Top 8 coach dihitung dari jumlah nama di tiap tab."
ws["B75"].font = font(8, False, GRAY, i=True)
ws.page_setup.orientation = "portrait"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True

# ======================================================================
# Member Aktif (interactive filter)
# ======================================================================
wsA = wb.create_sheet("Member Aktif", 1)
wsA.sheet_view.showGridLines = False
for col, w in {"A": 34, "B": 30, "C": 20, "D": 8, "E": 18, "F": 4}.items():
    wsA.column_dimensions[col].width = w
wsA["A1"] = "TWENTY SWIM  -  DAFTAR MEMBER AKTIF"; wsA["A1"].font = font(14, True, NAVY)
wsA.row_dimensions[1].height = 22
wsA["A2"] = "Gabungan semua siswa dari 30 tab coach, urut abjad. Pilih filter di bawah (kosongkan = tampilkan semua)."
wsA["A2"].font = font(9, False, GRAY)
wsA["A3"] = "Coach:"; wsA["C3"] = "Program:"; wsA["E3"] = "Level:"
for c in ("A3", "C3", "E3"):
    wsA[c].font = font(10, True, NAVY); wsA[c].alignment = Alignment(horizontal="right", vertical="center")
# filter cells sit right next to labels: B3 (coach), D3 (program), F3 (level)
wsA.column_dimensions["B"].width = 30; wsA.column_dimensions["D"].width = 20; wsA.column_dimensions["F"].width = 18
for c in ("B3", "D3", "F3"):
    wsA[c].fill = fill(YEL_BG); wsA[c].border = BORDER; wsA[c].font = font(10); wsA[c].alignment = LEFT
wsA.row_dimensions[3].height = 20
for c, f in (("B3", REF_COACH), ("D3", REF_PROG), ("F3", REF_LVL)):
    dv = DataValidation(type="list", formula1=f"={f}", allow_blank=True, showErrorMessage=True,
                        errorTitle="Pilihan tidak valid", error="Pilih dari daftar, atau kosongkan untuk semua.")
    wsA.add_data_validation(dv); dv.add(c)
wsA["A4"] = "Ditampilkan:"; wsA["A4"].font = font(9, False, GRAY); wsA["A4"].alignment = Alignment(horizontal="right")
wsA["B4"] = "=COUNTA($B$6:$B$2000)"; wsA["B4"].font = font(10, True, NAVY); wsA["B4"].alignment = LEFT
wsA["C4"] = "siswa"; wsA["C4"].font = font(9, False, GRAY)
for i, h in enumerate(["Nama Siswa", "Coach", "Program/Kelas", "Level"], start=1):
    c = wsA.cell(5, i, h); c.font = font(10, True, "FFFFFF"); c.fill = fill(NAVY); c.alignment = CENTER; c.border = BORDER
wsA.row_dimensions[5].height = 22
# D5 header sits in column D which is also the Program filter column - realign widths for table
wsA.column_dimensions["A"].width = 34; wsA.column_dimensions["B"].width = 30
wsA.column_dimensions["C"].width = 20; wsA.column_dimensions["D"].width = 18
wsA.column_dimensions["E"].width = 8; wsA.column_dimensions["F"].width = 18
if "--no-array" not in sys.argv:
  wsA["A6"] = ('=IFERROR(FILTER(DataMember!$A$2:$D$2000,DataMember!$A$2:$A$2000<>"",'
             '(((DataMember!$B$2:$B$2000=$B$3)+($B$3=""))>0),'
             '(((DataMember!$C$2:$C$2000=$D$3)+($D$3=""))>0),'
             '(((DataMember!$D$2:$D$2000=$F$3)+($F$3=""))>0)),"Tidak ada siswa untuk filter ini.")')
for r in range(6, 500):
    for cidx in range(1, 5):
        c = wsA.cell(r, cidx); c.font = font(10); c.border = BORDER
        c.alignment = LEFT
wsA.conditional_formatting.add("A6:D500", FormulaRule(formula=['AND($A6<>"",MOD(ROW(),2)=0)'], fill=fill("F5F7FA")))
wsA.freeze_panes = "A6"
wsA.page_setup.fitToWidth = 1; wsA.page_setup.fitToHeight = 0; wsA.sheet_properties.pageSetUpPr.fitToPage = True

# ======================================================================
# Panduan Level SSI
# ======================================================================
wsP = wb.create_sheet("Panduan Level SSI", 2)
wsP.sheet_view.showGridLines = False
for col, w in {"A": 24, "B": 26, "C": 34, "D": 3, "E": 26, "F": 40}.items():
    wsP.column_dimensions[col].width = w
wsP["A1"] = "PANDUAN LEVEL SSI (SWIM SCHOOL INTERNATIONAL)"; wsP["A1"].font = font(14, True, NAVY)
wsP.row_dimensions[1].height = 22
def hdr(ws, row, cols, texts):
    for c, t in zip(cols, texts):
        cell = ws[f"{c}{row}"]; cell.value = t
        cell.font = font(10, True, "FFFFFF"); cell.fill = fill(NAVY); cell.alignment = CENTER; cell.border = BORDER
hdr(wsP, 3, "ABC", ["Kategori", "Level", "Rentang Usia"])
guide = [("Baby Swim", "Baby Swim I", "3 bulan - 2 thn 11 bln"), ("Baby Swim", "Baby Swim II", "3 bulan - 2 thn 11 bln"),
         ("Aquatike", "Aquatike I", "3 thn - 4 thn 11 bln"), ("Aquatike", "Aquatike II", "3 thn - 4 thn 11 bln"),
         ("Aquatike", "Aquatike III", "3 thn - 4 thn 11 bln"), ("Preschool", "Preschool I", "5 thn - 6 thn"),
         ("Preschool", "Preschool II", "5 thn - 6 thn"), ("Beginner", "Beginner I", "6 thn - dewasa"),
         ("Beginner", "Beginner II", "6 thn - dewasa"), ("Beginner", "Beginner III", "6 thn - dewasa"),
         ("Intermediate", "Intermediate I", "6 thn - dewasa"), ("Intermediate", "Intermediate II", "6 thn - dewasa"),
         ("Advanced", "Advanced I", "6 thn - dewasa"), ("Advanced", "Advanced II", "6 thn - dewasa"),
         ("Swim Team Prep", "Swim Team Prep", "Transisi atlet kompetitif")]
for i, row in enumerate(guide, start=4):
    for c, v in zip("ABC", row):
        cell = wsP[f"{c}{i}"]; cell.value = v; cell.font = font(10); cell.border = BORDER; cell.alignment = LEFT
        if i % 2 == 1: cell.fill = fill("F5F7FA")
r = 21
wsP[f"A{r}"] = "KETERANGAN STATUS"; wsP[f"A{r}"].font = font(11, True, NAVY)
hdr(wsP, r + 1, "AB", ["Status", "Keterangan"])
for i, (a, b) in enumerate([("PA (Parent Assist)", "Ada pendamping di air"), ("Standalone", "Mandiri tanpa pendamping")]):
    for c, v in zip("AB", (a, b)):
        cell = wsP[f"{c}{r+2+i}"]; cell.value = v; cell.font = font(10); cell.border = BORDER; cell.alignment = LEFT

# right side: program & colour legend, how-to
wsP["E3"] = "PROGRAM / KELAS"; wsP["E3"].font = font(11, True, NAVY)
hdr(wsP, 4, "EF", ["Program", "Warna di tab coach"])
for i, p in enumerate(PROGRAMS):
    a = wsP[f"E{5+i}"]; b = wsP[f"F{5+i}"]
    a.value = p; b.value = "Biru muda" if i < 4 else "Merah muda"
    for c in (a, b): c.font = font(10); c.border = BORDER; c.alignment = LEFT
    b.fill = fill(LBLUE if i < 4 else PINK_BG)
wsP["E12"] = "RISIKO & SOSIALISASI"; wsP["E12"].font = font(11, True, NAVY)
hdr(wsP, 13, "EF", ["Nilai", "Warna"])
leg = [("Risiko: Rendah", GREEN_BG), ("Risiko: Sedang", YEL_BG), ("Risiko: Tinggi", RED_BG),
       ("Sosialisasi: Sudah", GREEN_STRONG), ("Sosialisasi: Belum", RED_STRONG)]
for i, (t, col) in enumerate(leg):
    a = wsP[f"E{14+i}"]; b = wsP[f"F{14+i}"]
    a.value = t; b.value = "otomatis saat dipilih"; b.fill = fill(col)
    for c in (a, b): c.font = font(10); c.border = BORDER; c.alignment = LEFT
wsP["E21"] = "CARA PAKAI FILE INI"; wsP["E21"].font = font(11, True, NAVY)
howto = [
    "1. Isi data hanya di tab coach (kolom putih). Kolom dropdown tinggal pilih.",
    "2. Kolom No dan Coach terisi otomatis - jangan diketik manual.",
    "3. Tabel kanan di tiap tab coach adalah jadwal latihan: isi Jam, Nama Siswa, Lokasi, Kolam.",
    "4. Tab DASHBOARD dan Member Aktif menghitung sendiri dari 30 tab coach.",
    "5. Tab Member Aktif: pilih Coach / Program / Level untuk menyaring; kosongkan untuk semua.",
    "6. Tab DataMember, Data Dashboard, dan Daftar disembunyikan (sumber rumus & dropdown). Jangan diubah.",
    "7. Rumus SORT + FILTER di Member Aktif dibuat untuk Google Sheets (Excel 365 juga mendukung).",
    "Sumber data awal: Database_Siswa_Twenty_SSI.xlsx (Nama, Level Saat Ini, Status) per tab coach.",
    "Program/Kelas diisi otomatis hanya jika ada petunjuk di data asal, mis. '(Private)', '(Grup)',",
    "'(Couple)', 'Swim/Hydro Therapy' -> Hydrotherapy Class, level Baby Swim -> Baby Swim. Sisanya kosong.",
]
for i, t in enumerate(howto):
    c = wsP[f"E{22+i}"]; c.value = t; c.font = font(9, False, "374151"); c.alignment = Alignment(vertical="center")
    wsP.merge_cells(f"E{22+i}:F{22+i}")

# ======================================================================
# order, hidden tabs, workbook settings
# ======================================================================
for nm in ("DataMember", "Data Dashboard", "Daftar"):
    wb[nm].sheet_state = "hidden"
order = ["DASHBOARD", "Member Aktif", "Panduan Level SSI"] + COACHES + ["DataMember", "Data Dashboard", "Daftar"]
wb._sheets = [wb[n] for n in order]
wb.active = 0
wb.calculation = CalcProperties(fullCalcOnLoad=True)
wb.save(OUT)
print("saved", OUT, "sheets:", len(wb.sheetnames))
