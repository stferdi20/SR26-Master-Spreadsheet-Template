"""Cross-tab consistency check for the generated workbooks. Run after build.py:  python3 check.py

Checks that a change in one place (units, indicators, quant routing, versions) has flowed to every dependent tab
and request workbook. Exits non-zero if anything is inconsistent.
"""
import glob
import re
import sys
import warnings

from openpyxl import load_workbook

import build as B
import quant
import rasci

warnings.filterwarnings("ignore")
fails, notes = [], []


def check(cond, msg):
    (notes if cond else fails).append(("OK  " if cond else "FAIL") + "  " + msg)


V = B.VERSION
master_path = f"output/SR26 end-year reporting master spreadsheet - {V}.xlsx"
M = load_workbook(master_path)
req_files = sorted(glob.glob(f"output/*_requests/* - {V}.xlsx"))
stale = [f for f in glob.glob("output/**/*.xlsx", recursive=True) if not f.endswith(f"{V}.xlsx")]
check(bool(glob.glob(f"output/SR26 project management - {V}.xlsx")), "separate project management workbook exists")
check(not any(n.startswith("P") and n[1].isdigit() for n in M.sheetnames), "no project (P0–P5) tabs left in the master")
check(not stale, f"all output files carry the current version {V} ({len(req_files) + 1} files)")

qual = B.read_qual()
B.DATABOOK = quant.read_databook(qual)
purged = quant.quant_only_rows(qual)
ref_ok = {q["ref"] for q in qual if q["row"] not in purged}
units = list(B.PILOT_FACULTIES) + [c[0] for c in B.CENTRAL]
UNCODE = {B.file_code(u): u for u in units}

# ---- request workbooks: rows, tabs, links
req_rows = {}
for f in req_files:
    code = UNCODE.get(f.split("/")[-1].split(" - ")[0], f.split("/")[-1].split(" - ")[0])
    wb = load_workbook(f)
    t = wb["1. Reporting template"]
    rows = [t.cell(r, 2).value for r in range(1, t.max_row + 1) if re.match(r"^[A-Z]{2}\d\(", str(t.cell(r, 2).value or ""))]
    req_rows[code] = rows
    check(set(rows) <= ref_ok, f"{code}: all {len(rows)} requested refs exist in the requirements matrix")
    check(wb.sheetnames[:3] == ["1. Reporting template", "2. OPTIONAL highlighted stories", "3. SR25 responses"],
          f"{code}: tab order 1-2-3")
    owned = quant.owned_points(B.DATABOOK, code) if code not in B.PILOT_FACULTIES else []
    check(bool(owned) == ("4. Quantitative data" in wb.sheetnames), f"{code}: Tab 4 present only if it owns Databook metrics ({len(owned)})")
    s3 = wb["3. SR25 responses"]
    bad = [c.coordinate for row in t.iter_rows(min_col=8, max_col=8) for c in row
           if c.hyperlink and s3.cell(int(c.hyperlink.location.split("!B")[1]), 2).value != str(c.value).split("\n")[0]]
    check(not bad, f"{code}: SR25 links land on the right row")
    for ws in wb:
        check(len(ws.title) <= 31, f"{code}/{ws.title}: sheet name length")
        for dv in ws.data_validations.dataValidation:
            check(not ((dv.formula1 or "").startswith('"') and len(dv.formula1) > 257), f"{code}/{ws.title}: dropdown list length")
check(set(req_rows) == set(units), f"one request workbook per unit ({len(units)})")

# ---- tracker matches request files
tr = M["4. Request tracker"]
trk = {}
for r in range(10, tr.max_row + 1):
    f = tr.cell(r, 7).value
    if f:
        code = UNCODE.get(f.split(" - ")[0], f.split(" - ")[0])
        trk[code] = trk.get(code, 0) + (tr.cell(r, 6).value or 0)
        check(f.endswith(f"{V}.xlsx"), f"tracker file name for {code} is current version")
for code, rows in req_rows.items():
    check(trk.get(code) == len(rows), f"tracker indicator count for {code} = {trk.get(code)} vs request rows {len(rows)}")

# ---- requirements matrix 'issued to' matches requests
rm = M["1. Requirements matrix"]
issued = {rm.cell(r, 2).value: str(rm.cell(r, 9).value) for r in range(6, rm.max_row + 1) if rm.cell(r, 2).value}
check(set(issued) == ref_ok, f"requirements matrix holds exactly the {len(ref_ok)} non-quant-only indicators")
for code, rows in req_rows.items():
    tag = "Faculties (pilot)" if code in B.PILOT_FACULTIES else code
    miss = [r for r in rows if tag not in issued.get(r, "")]
    check(not miss, f"requirements matrix 'issued to' names {tag} for all its requests" + (f" – missing {miss}" if miss else ""))

# ---- scoring tabs use the same unit list
a1 = M["A1. Data scoring"]
a1_units = [a1.cell(10, c).value for c in range(8, a1.max_column + 1)]
a1_units = a1_units[:a1_units.index(None)] if None in a1_units else a1_units
tc = M["0.5 Target Check"]
tc_units = [tc.cell(6, c).value for c in range(3, 3 + len(a1_units))]
t2 = M["T2. Single target review"]
t2_units = [t2.cell(14, c).value for c in range(3, 3 + len(a1_units))]
check(a1_units == tc_units == t2_units, f"A1, 0.5 and T2 list the same {len(a1_units)} unit sections")
n_sections = len(B.PILOT_FACULTIES) + sum(1 for c in B.CENTRAL for _, m in c[3] if any(m(q) for q in qual))
check(len(a1_units) == n_sections, f"scoring covers every unit section with requests ({n_sections})")
a1_refs = {a1.cell(r, 2).value for r in range(11, a1.max_row + 1) if a1.cell(r, 2).value}
check(a1_refs == ref_ok, "A1 scores exactly the requirements-matrix indicators")

# ---- RASCI: every request row is an R for that area; no R without a request (except quant/compile)
rs = M["3b. RASCI matrix"]
hdr = {c: rs.cell(6, c).value for c in range(9, rs.max_column + 1) if rs.cell(6, c).value}
codes = {}
for r in range(7, rs.max_row + 1):
    ref = rs.cell(r, 1).value
    if rs.cell(r, 2).value == "Indicator":
        codes[ref] = {hdr[c]: rs.cell(r, c).value for c in hdr if rs.cell(r, c).value}
check(set(codes) == {q["ref"] for q in qual}, "RASCI lists every Sustainability 2030 indicator")
for code, rows in req_rows.items():
    if code in B.PILOT_FACULTIES:
        cols = [rasci.FACULTIES_LABEL(B.PILOT_FACULTIES)]
    else:
        cols = [rasci.area_label(l) for c in B.CENTRAL if c[0] == code for l, _ in c[3]]
        check(all(c in hdr.values() for c in cols), f"RASCI has a column for every {code} section")
    have = {ref for ref, d in codes.items() if any(d.get(c) in ("R", "A and R") for c in cols)}
    check(set(rows) == have, f"RASCI 'R' for {code} matches its request rows" +
          ("" if set(rows) == have else f" – request only {sorted(set(rows) - have)}, RASCI only {sorted(have - set(rows))}"))

# ---- stakeholder map lists every central section
sm = M["3. Stakeholder map"]
sm_text = " ".join(str(c.value) for row in sm.iter_rows() for c in row if c.value)
for code, _, _, secs in B.CENTRAL:
    for label, _ in secs:
        check(label in sm_text, f"stakeholder map lists {label}")

# ---- quant coverage owners have the Tab 4
qc = M["2. Quant coverage"]
for r in range(6, qc.max_row + 1):
    own, n, where = qc.cell(r, 10).value, qc.cell(r, 9).value, qc.cell(r, 12).value
    if own and own != "SST" and n:
        check(own in req_rows and where, f"quant coverage {qc.cell(r, 2).value}: owner {own} has a request with Tab 4")

# ---- triangulation uses SR26 area names that exist in the RASCI
tri = M["3c. RASCI triangulation"]
areas = set(hdr.values())
for r in range(6, tri.max_row + 1):
    for col in (7, 8, 9):
        for a in str(tri.cell(r, col).value or "").split("\n"):
            if a and not a.startswith("Not in SR26") and tri.cell(r, 2).value and re.match(r"[A-Z]{2}\d", str(tri.cell(r, 2).value)):
                check(a in areas, f"triangulation {tri.cell(r, 2).value}: area '{a}' exists in RASCI")

# ---- no stale hard-coded text
alltext = " ".join(str(c.value) for ws in M for row in ws.iter_rows() for c in row if isinstance(c.value, str))
for phrase in ["no SR26 request issued", "6 central units", "V0.1", "Quant – Databook mapping", "Stakeholder map & RASCI",
               "P0 Project overview"]:
    check(phrase not in alltext, f"no stale text '{phrase}'")
check(f"DRAFT {V}" in alltext, f"Read me shows {V}")
# old timeline must not survive anywhere (all workbooks)
for f in [master_path] + req_files + glob.glob(f"output/SR26 project management - {V}.xlsx"):
    txt = " ".join(str(c.value) for ws in load_workbook(f) for row in ws.iter_rows() for c in row if isinstance(c.value, str))
    for phrase in ["15 January 2027", "4 December 2026", "4 Dec (TBC)", "15 Jan (TBC)", "Support meetings: January"]:
        check(phrase not in txt, f"{f.split('/')[-1][:30]}: no old date '{phrase}'")

# ---- report (de-duplicated)
seen = set()
for line in notes + fails:
    if line not in seen:
        seen.add(line)
ok = len([n for n in notes])
print(f"{len(set(notes))} checks passed, {len(set(fails))} failed")
for f in dict.fromkeys(fails):
    print(f)
sys.exit(1 if fails else 0)
