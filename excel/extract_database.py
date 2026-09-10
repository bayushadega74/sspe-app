"""
Pull Nama / Level Saat Ini / Status from every coach tab of Database_Siswa_Twenty_SSI.xlsx
into data.json (input for build_level_jadwal_coach.py).

Columns are located by header text on row 4 (two tabs have Status shifted to column F).
Program/Kelas is inferred only from explicit hints in the source:
  "(Private)" -> Private Class, "(Grup)/(Group)" -> Group Class, "(Couple ...)" -> Couple Class,
  "Swim/Hydro Therapy" (level or name) -> Hydrotherapy Class, level Baby Swim -> Baby Swim.
Levels outside the 15 SSI levels (Hydro/Swim Therapy) are left blank.
"""
import sys, re, json, collections, openpyxl

SRC = sys.argv[1]
OUT = sys.argv[2] if len(sys.argv) > 2 else "data.json"
LEVELS = ["Baby Swim I", "Baby Swim II", "Aquatike I", "Aquatike II", "Aquatike III", "Preschool I", "Preschool II",
          "Beginner I", "Beginner II", "Beginner III", "Intermediate I", "Intermediate II", "Advanced I", "Advanced II",
          "Swim Team Prep"]
lv_norm = {l.lower(): l for l in LEVELS}
SKIP = {"Daftar Isi", "Panduan Level SSI", "Rekap data", "Sheet5"}

wb = openpyxl.load_workbook(SRC, data_only=True)
coaches = [ws.title for ws in wb.worksheets if ws.title not in SKIP]
data, unmatched = {}, collections.Counter()
for name in coaches:
    ws = wb[name]
    hdr = {str(c.value).strip(): c.column for c in ws[4] if c.value}
    cn, cl, cs = hdr["Nama Siswa"], hdr["Level Saat Ini"], hdr["Status"]
    rows = []
    for r in range(5, 80):
        nm = ws.cell(r, cn).value
        if nm is None or str(nm).strip() == "":
            continue
        nm = re.sub(r"\s+", " ", str(nm)).strip()
        lv = ws.cell(r, cl).value; lv = re.sub(r"\s+", " ", str(lv)).strip() if lv else ""
        st = ws.cell(r, cs).value; st = str(st).strip() if st else ""
        low = nm.lower(); prog = ""
        if "therapy" in lv.lower() or "terapi" in lv.lower() or "therapy" in low or "terapi" in low:
            prog = "Hydrotherapy Class"
        elif "couple" in low: prog = "Couple Class"
        elif "private" in low: prog = "Private Class"
        elif "grup" in low or "group" in low: prog = "Group Class"
        level = lv_norm.get(lv.lower(), "")
        if lv and not level: unmatched[lv] += 1
        if level.startswith("Baby Swim") and not prog: prog = "Baby Swim"
        if st not in ("PA", "Standalone"): st = ""
        rows.append([nm, prog, level, st])
    data[name] = rows
json.dump({"coaches": coaches, "data": data}, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(len(coaches), "coaches,", sum(len(v) for v in data.values()), "students; unmatched levels:", dict(unmatched))
