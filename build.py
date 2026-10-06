"""Build the SR26 data collection master spreadsheet and per-unit reporting request workbooks.

Inputs (not committed, place in ./source/):
  S2030_targets.xlsx   - Sustainability 2030 updated targets (Qual + Quan worksheets)
  SR25_master.xlsx     - 2025 end-year reporting master spreadsheet (for End-2025 reference responses)

Run:  python3 build.py
"""
import re
import warnings
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink

from faculty_requests import USE, faculty_request
import quant
import rasci

warnings.filterwarnings("ignore")
ROOT = Path(__file__).parent
SRC = ROOT / "source"
OUT = ROOT / "output"

# ---------------------------------------------------------------- settings
VERSION = "v0.7"  # bump on every revision; appears in file names and Read me
VERSION_HISTORY = [
    ("v0.1", "First draft: requirements matrix, pilot faculty and central-unit requests, tracker, evidence register."),
    ("v0.2", "SR25-style faculty wording and 'How SST will use'; scoring/T1/T2 formulas; Legal & Risk; project tabs; linked SR25 responses tab."),
    ("v0.3", "Quantitative indicators separated: removed from requirements matrix; '2. Quant coverage' and '2b. Databook register' against the Databook draft; '4. Quantitative data' tab in central-unit requests."),
    ("v0.4", "RASCI rebuilt to the SR25 standard ('3b. RASCI matrix'): indicator hierarchy, consolidated portfolios, SST contacts, A/R/S/C/I codes, summary counts."),
    ("v0.5", "'3c. RASCI triangulation': each SR26 indicator compared with the SR25 RASCI (equivalent targets, translated areas), gaps and suggested actions."),
    ("v0.6", "Chancellery units added following SR25 (ACM, CGCE, Chancellery Education, SASS, Chancellery Indigenous) with linked SR25 responses; cross-tab consistency check (check.py); Impact column added to Issues log (as SR25)."),
    ("v0.7", "Team review: Chancellery units as separate tabs (AC&M, MRE, GCE, Education, SaSS, Indigenous); faculty requests cut down (ABP walkthrough) – EE1(a)+(c) merged, 'any examples' wording, alumni to SaSS, CL2(a)/(c) to AC&M, TR3(a)/(c)/(d) to MRE; simpler evidence register; highlighted stories compiled per unit with SST priority-area columns; project tabs moved to a separate project workbook."),
]
DUE_DATE = "Friday 15 January 2027 (TBC)"
STORIES_DUE = "Friday 4 December 2026 (TBC)"
CONTACT = "Sustainability Strategy team (sustainability-strategy@unimelb.edu.au – TBC)"
PILOT_FACULTIES = {  # SR25 code -> full name
    "ABP": "Faculty of Architecture, Building and Planning",
    "ARTS": "Faculty of Arts",
    "FBE": "Faculty of Business and Economics",
    "SCI": "Faculty of Science",
}
# Central units: SR26 name, SR25 equivalent, sections (name, matcher on Qual row)
ESTATE_PA = {"Estate and infrastructure"}


def _is_estate(r):
    return r["pa"] in ESTATE_PA or r["row"] == 30  # row 30 = climate resilience maturity (Gerard)


CENTRAL = [
    ("CI&S", "Commercial, Infrastructure & Sustainability (CI&S)",
     "Corporate Development (CD Development & Estate Planning teams; CD Sustainability Strategy)",
     [("CI&S – Estate Planning & Development (TBC: confirm with Gerard)", lambda r: "CI&S" in r["stake"] and _is_estate(r)),
      ("CI&S – Sustainability Strategy (incl. CLLAP)", lambda r: "CI&S" in r["stake"] and not _is_estate(r))]),
    ("CFOG", "Chief Finance Officer Group (CFOG)",
     "CFOG (Procurement; Accounting; Treasury & Investments; Enterprise Portfolio & Investments)",
     [("CFOG – Procurement", lambda r: r["row"] == 46 or "CFOG" in r["stake"] and r["pa"] in {"Responsible procurement", "Modern slavery and human rights", "Nature and biodiversity", "Transformational research", "Enablers"}),
      ("CFOG – Treasury & Investments", lambda r: "CFOG" in r["stake"] and r["pa"] == "Responsible investments")]),
    ("ESG", "Enterprise Services Group (ESG)",
     "Business Services – Sustainability Delivery (nature & biodiversity, waste & circular economy, part of quant. climate leadership)",
     [("ESG – Campus Operations & Sustainability Delivery", lambda r: "ESG" in r["stake"] or "Rachel" in r["stake"])]),
    ("CIOG", "Chief Information Officer Group (CIOG)",
     "Business Services (AI – not requested in SR25)",
     [("CIOG – Responsible AI & digital", lambda r: "CIOG" in r["stake"])]),
    ("MRE", "Melbourne Research and Enterprise (MRE)",
     "Chancellery Research and Enterprise",
     [("MRE – Research (incl. MBI – TBC)", lambda r: "MRE" in r["stake"] or "MBI" in r["stake"])]),
    # Chancellery units – each reported separately (SR25 Chancellery tab sections), named as the team uses them
    ("AC&M", "Advancement, Communication and Marketing (AC&M)", "Chancellery – Advancement, Communications & Marketing",
     [("AC&M – Advancement, Communication and Marketing", lambda r: r["row"] in (16, 54))]),
    ("GCE", "Global, Culture and Engagement (GCE)", "Chancellery – Global, Culture and Engagement",
     [("GCE – Global, Culture and Engagement", lambda r: r["row"] in (52, 54))]),
    ("Education", "Chancellery Education", "Chancellery – Chancellery Education",
     [("Education – Chancellery Education", lambda r: r["row"] in (8, 9, 54))]),
    ("SaSS", "Student and Scholarly Services (SaSS)", "Chancellery – Student and scholarly services",
     [("SaSS – Student and Scholarly Services", lambda r: r["row"] == 10)]),
    ("Indigenous", "Chancellery Indigenous", "Chancellery – Chancellery Indigenous",
     [("Indigenous – Chancellery Indigenous", lambda r: r["row"] == 52)]),
    # SR25 Legal and Risk reported 8a(i)/(ii) climate change preparedness (University Risk 16, flood emergency plans)
    ("L&R", "Legal and Risk", "Legal and Risk",
     [("Legal & Risk – Risk & resilience", lambda r: r["row"] == 30)]),
]

# New S2030 Qual row -> related SR25 faculty target ref (for End-2025 reference columns)
SR25_MAP = {8: "4a(ii)", 11: "4b(i)", 12: "2a", 16: "5b(i)", 19: "5b(i)", 20: "5c(ii)",
            21: "5c(ii)", 52: "6a", 53: "2a", 54: "3a(ii)"}
PA_CODE = {"Exceptional education": "EE", "Transformational research": "TR", "Climate leadership": "CL",
           "Nature and biodiversity": "NB", "Responsible AI": "AI", "Estate and infrastructure": "EI",
           "Circular economy": "CE", "Responsible procurement": "RP", "Modern slavery and human rights": "MS",
           "Responsible investments": "RI", "Enablers": "EN"}
STATUS = ["Met or exceeded", "Partially met", "Not met", "Not yet started", "Not applicable"]
STATUS_DEF = [
    ("Met or exceeded", "The data and information provided demonstrates that the Faculty/Portfolio has met or exceeded the target"),
    ("Partially met", "The data and information provided indicates that the target has not been met, but there has been a significant improvement and performance is trending towards the target"),
    ("Not met", "Some action has been taken to address the target, but the data and information provided demonstrates that this has been insufficient to meet the target"),
    ("Not yet started", "No action has been taken."),
]
SCORES = [("Met or exceeded", 3), ("Partially met", 2), ("Not met", 1), ("Not yet started", 0)]  # as SR25 'Lists (Hide)'
TRACK_STATUS = ["Not sent", "Sent", "Chased", "Support meeting", "Received", "Under review", "Complete", "Not required"]

# ---------------------------------------------------------------- styles (SR25 look)
NAVY, DARK = "002060", "0E2841"
HDR = PatternFill("solid", fgColor="0B3040")      # accent1 -50%
REFF = PatternFill("solid", fgColor="C1E4F5")     # accent1 +80%
GOLD = PatternFill("solid", fgColor="FFC000")
GREY = PatternFill("solid", fgColor="F2F2F2")
GREY2 = PatternFill("solid", fgColor="D9D9D9")
WHITE_B = Font(name="Aptos", size=12, bold=True, color="FFFFFF")
thin = Side(style="thin", color="808080")
med = Side(style="medium")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
MBOX = Border(left=med, right=med, top=med, bottom=med)
WRAP = Alignment(wrap_text=True, vertical="top")
CWRAP = Alignment(wrap_text=True, vertical="center", horizontal="center")


def title(ws, text, sub):
    ws["A1"] = "↑"; ws["A1"].font = Font(name="Aptos", color=NAVY)
    ws["B1"] = text; ws["B1"].font = Font(name="Arial", size=11)
    ws["B2"] = sub; ws["B2"].font = Font(name="Aptos", size=22, bold=True, color=NAVY)
    ws.sheet_view.showGridLines = False


def hdr_row(ws, row, headers, col=2, fill=HDR, font=WHITE_B):
    for i, h in enumerate(headers):
        c = ws.cell(row, col + i, h)
        c.fill, c.font, c.alignment, c.border = fill, font, CWRAP, BOX


def body(c, fill=None, bold=False):
    c.alignment, c.border = WRAP, BOX
    c.font = Font(name="Aptos Narrow", size=11, bold=bold)
    if fill:
        c.fill = fill


def widths(ws, w):
    for k, v in w.items():
        ws.column_dimensions[k].width = v


def dv_list(ws, src, rng):
    dv = DataValidation(type="list", formula1=src, allow_blank=True)
    dv.error, dv.errorTitle = "Please select from the dropdown", "Invalid entry"
    ws.add_data_validation(dv)
    dv.add(rng)


def lists_sheet(wb):
    ws = wb.create_sheet("Lists")
    ws["B1"] = "End 2026 target status"
    ws["B2"] = "Please select from dropdown"
    for i, s in enumerate(STATUS):
        ws.cell(3 + i, 2, s)
    ws["C1"] = "Data owner confirms"
    for i, s in enumerate(["Yes", "No", "Partially – see comments"]):
        ws.cell(2 + i, 3, s)
    ws["D1"] = "Request status"
    for i, s in enumerate(TRACK_STATUS):
        ws.cell(2 + i, 4, s)
    ws["E1"] = "Y/N"; ws["E2"] = "Yes"; ws["E3"] = "No"; ws["E4"] = "TBC"
    ws["G1"] = "Target status"; ws["H1"] = "Score"; ws["I1"] = "Response count"  # used by A1. Data scoring
    for i, (st, sc) in enumerate(SCORES):
        ws.cell(2 + i, 7, st); ws.cell(2 + i, 8, sc); ws.cell(2 + i, 9, 1)
    ws.sheet_state = "hidden"
    return {"status": "=Lists!$B$3:$B$7", "confirm": "=Lists!$C$2:$C$4",
            "track": f"=Lists!$D$2:$D${1 + len(TRACK_STATUS)}", "yn": "=Lists!$E$2:$E$4"}


# ---------------------------------------------------------------- read inputs
def clean(v):
    return re.sub(r"\s+$", "", str(v)).replace("\xa0", " ") if v not in (None, "") else ""


def read_qual():
    ws = load_workbook(SRC / "S2030_targets.xlsx", data_only=True)["Qual"]
    rows, dom, pa, tgt, counters = [], "", "", "", {}
    for r in range(8, ws.max_row + 1):
        g = lambda col: clean(ws[f"{col}{r}"].value)
        if not g("F"):
            continue
        dom = g("C") or dom
        if g("D") and g("D") != pa:
            pa = g("D"); counters[pa] = [0, 0]
        if g("E") and g("E") != tgt:
            tgt = g("E"); counters[pa][0] += 1; counters[pa][1] = 0
        counters[pa][1] += 1
        code = PA_CODE.get(pa, "XX")
        rows.append(dict(row=r, ref=f"{code}{counters[pa][0]}({chr(96 + counters[pa][1])})",
                         domain=dom.replace("Educationand", "Education and"), pa=pa, target=tgt,
                         indicator=g("F"), qual=g("G"), request=g("H"), stake=g("I"), projects=g("J"),
                         intext=g("K"), ci=g("L"), notes=g("M")))
    return rows


def read_quan(qual_by_row):
    ws = load_workbook(SRC / "S2030_targets.xlsx", data_only=True)["Quan"]
    out = []
    for r in range(8, ws.max_row + 1):
        g = lambda col: clean(ws[f"{col}{r}"].value)
        if g("G") in ("Y", "M") and g("F"):
            q = qual_by_row.get(r, {})
            out.append(dict(row=r, ref=q.get("ref", ""), pa=q.get("pa", ""), indicator=g("F"), flag=g("G"),
                            metric=g("H"), person=g("I"), intext=g("J"), notes=g("L")))
    return out


# Indicators asked of one unit only (team review, ABP walkthrough): Qual row -> unit code
EXCLUSIVE = {11: "SaSS",                      # EE1(d) alumni
             27: "AC&M", 29: "AC&M",          # CL2(a), CL2(c) climate communications / mapping
             20: "MRE", 22: "MRE", 23: "MRE"}  # TR3(a), (c), (d) research systems, culture, assessment
FACULTY_MERGED = {10: 8}  # EE1(c) is covered by the faculty EE1(a) question


def asks_faculties(r):
    """True if the pilot faculties receive a request for this Qual row."""
    return "Faculties" in r["stake"] and r["row"] not in EXCLUSIVE and r["row"] not in FACULTY_MERGED


def _exclusive(m, code, is_first):
    """Exclusive rows go to their owner's first section only; other rows use the section's own matcher."""
    return lambda r: (EXCLUSIVE[r["row"]] == code and is_first) if r["row"] in EXCLUSIVE else m(r)


for _code, _name, _sr25, _secs in CENTRAL:
    _secs[:] = [(lab, _exclusive(m, _code, k == 0)) for k, (lab, m) in enumerate(_secs)]


def file_code(code):
    """File-name-safe unit code."""
    return {"CI&S": "CIandS", "L&R": "LandR", "AC&M": "ACM"}.get(code, code)


# SR25 master tab/section holding each unit's end-2025 responses (None = no SR25 equivalent)
SR25_SOURCES = {
    "CI&S": ("CD", None),
    "CFOG": ("CFOG", None),
    "ESG": ("Business Services", None),
    "CIOG": None,  # AI not requested in SR25
    "MRE": ("Chancellery", "Research and Enterprise"),
    "AC&M": ("Chancellery", "Advancement"),
    "GCE": ("Chancellery", "Global, Culture"),
    "Education": ("Chancellery", "Chancellery Education"),
    "SaSS": ("Chancellery", "Student and scholarly"),
    "Indigenous": ("Chancellery", "Chancellery Indigenous"),
    "L&R": ("Legal and Risk", None),
}
# Central units: S2030 Qual row -> closest SR25 (SP2030) ref answered by that unit
CENTRAL_SR25_MAP = {
    "CI&S": {30: "8a(i)", 53: "2a", 12: "5a(ii)", 51: "12b(i)", 49: "11c(ii)", 40: "7b"},
    "CFOG": {49: "11c(i)", 46: "11c(iii)", 51: "12b(i)", 47: "11d", 48: "11d"},
    "ESG": {32: "9a(ii)", 33: "9a(i)", 34: "9b(i)", 54: "7a", 53: "2a"},
    "CIOG": {},
    "MRE": {12: "5a(i)", 16: "5b(ii)", 18: "5b(iii)", 20: "5c(i)", 21: "5c(ii)", 53: "2a"},
    "L&R": {30: "8a(i)"},
    "AC&M": {16: "5b(ii)", 29: "5b(ii)", 54: "3e"},
    "GCE": {54: "7c(ii)"},
    "Education": {8: "4a(ii)", 54: "3a(ii)"},
    "SaSS": {10: "3b(vii)", 11: "3b(ii)"},
    "Indigenous": {52: "6a"},
}
SR25_COLS = {"status": "Confirm end-20", "opt1": "OPTION 1", "opt2": "OPTION 2", "support": "Supporting inf"}


def read_sr25_tab(tab, section=None):
    """[(section name, [row dict])] for one SR25 master tab, located by header text (layouts differ per tab)."""
    ws = load_workbook(SRC / "SR25_master.xlsx", data_only=True)[tab]
    out, cur, cmap = [], None, {}
    for r in range(1, ws.max_row + 1):
        b, c = clean(ws.cell(r, 2).value).strip(), clean(ws.cell(r, 3).value).strip()
        if re.fullmatch(r"\d\.\d(\.\d)?", b) and c:
            cur = (c, []); out.append(cur); continue
        if b == "Ref":
            hdr = {clean(ws.cell(r, k).value): k for k in range(2, ws.max_column + 1) if ws.cell(r, k).value}
            cmap = {key: next((k for h, k in hdr.items() if h.startswith(pre)), None) for key, pre in SR25_COLS.items()}
            cmap["request"] = 4
            if cur is None:
                cur = (tab, []); out.append(cur)
            continue
        if cur is not None and cmap and re.match(r"^\d+[a-z]", b):
            g = lambda k: clean(ws.cell(r, cmap[k]).value) if cmap.get(k) else ""
            cur[1].append(dict(ref=b, target=c, request=g("request"), status=g("status"), opt1=g("opt1"),
                               opt2=g("opt2"), support=g("support")))
    if section:
        out = [x for x in out if section.lower() in x[0].lower()]
    return [x for x in out if x[1]]


def sr25_for(code):
    """SR25 sections for a faculty code or central unit code."""
    if code in PILOT_FACULTIES:
        return [x for x in read_sr25_tab("Faculties") if x[0].upper() == code]
    src = SR25_SOURCES.get(code)
    return read_sr25_tab(*src) if src else []


def prev_lookup(sections):
    """{sr25 ref: (status, response, links)} – first occurrence wins."""
    d = {}
    for _, rows in sections:
        for x in rows:
            resp = "\n\n".join(v for v in (x["opt1"], x["opt2"]) if v)
            if x["ref"] not in d or not (d[x["ref"]][0] or d[x["ref"]][1]):  # prefer an answered occurrence
                d[x["ref"]] = (x["status"], resp, x["support"])
    return d


def read_sr25_faculty():
    """{faculty code: {sr25 ref: (status, response, links)}} from SR25 master 'Faculties' tab."""
    ws = load_workbook(SRC / "SR25_master.xlsx", data_only=True)["Faculties"]
    res, cur = {}, None
    for r in range(1, ws.max_row + 1):
        b, c = clean(ws[f"B{r}"].value), clean(ws[f"C{r}"].value)
        if re.fullmatch(r"1\.\d+", b):
            cur = c.strip(); res[cur] = {}
        elif cur and re.match(r"\d[a-z]", b):
            resp = "\n\n".join(x for x in (clean(ws[f"M{r}"].value), clean(ws[f"N{r}"].value)) if x)
            res[cur][b] = (clean(ws[f"L{r}"].value), resp, clean(ws[f"O{r}"].value))
    return res


def use_text(r):
    if r["row"] in USE:
        return USE[r["row"]]
    if r["intext"] == "Internal":
        return "For internal management reporting only – this information will not be published."
    return (f"We will use this to inform the {r['pa']} section of the 2026 Sustainability Report (transition year), "
            "and to establish a baseline for reporting against Sustainability 2030 from 2027. Strong examples may be "
            "followed up as case studies.")


# ---------------------------------------------------------------- reporting block (template + master tabs)
COLS = ["Ref", "Priority area", "Sustainability 2030 target", "Indicator", "2026 Reporting request",
        "How SST will use this data",
        "Related SR25 ref", "End-2025 confirmed status", "End-2025 reporting response", "Links and supporting information",
        "Confirm end-2026 target status", "OPTION 1:  Target-level reporting (Column D)",
        "OPTION 2:  Indicator level reporting (Column E/F)", "Supporting information/comments",
        "Data source and limitations", "Data owner confirms information is accurate", "Confirmed by (name, role, date)"]
SUBS = {12: "(Met or exceeded; partially met; not met; not yet started)",
        13: "Provide further evidence of target status assessment in the form of:\nA general update on activities in relation to the overarching target (Column D)",
        14: "Provide further evidence of target status assessment in the form of:\nData and commentary in direct response to the reporting request (Column F)",
        15: "Provide any additional information to support the target status assessment (e.g. links, images, rationale for selection of target status)",
        16: "Where did this information come from (system, report, survey, person)? Note any gaps or limitations.",
        17: "Select 'Yes' to confirm the information provided is accurate and can be used for reporting",
        18: "Name, role and date of the person confirming"}
WIDTHS = {"A": 4, "B": 10, "C": 16, "D": 38, "E": 38, "F": 50, "G": 34, "H": 11, "I": 16, "J": 50, "K": 24,
          "L": 22, "M": 48, "N": 48, "O": 30, "P": 30, "Q": 20, "R": 24}


def block(ws, top, unit_label, rows, sr25, lists, reg=None, key=None, ref_map=None, tailor=None, anchors=None, quant_refs=()):
    """Write a reporting block at row `top`. Returns next free row.

    sr25: {SR25 ref: (status, response, links)} or None (no SR25 equivalent). ref_map: Qual row -> SR25 ref
    (defaults to the faculty map). tailor: use the SR25-style faculty wording (default: faculty blocks).
    anchors: {SR25 ref: row on '3. SR25 responses'} – makes Column H a link to the full SR25 response.
    If `reg` is given, (key, sheet, first_row, last_row) is recorded for the scoring formulas."""
    c = ws.cell(top, 2, unit_label)
    c.font = Font(name="Aptos Narrow", size=22, bold=True); c.fill = GREY
    h1 = top + 2
    for rng, text, fill, fnt in [
        ((3, 7), "Sustainability 2030 target and indicator", HDR, Font(name="Aptos Narrow", size=14, bold=True, color="FFFFFF")),
        ((8, 11), "End-2025 reporting (click the ref to see your full SR25 response; [+] above expands a summary)", HDR, Font(name="Aptos Narrow", size=14, bold=True, color="FFFFFF")),
        ((12, 18), "End-2026 reporting (PLEASE COMPLETE THIS SECTION)", GOLD, Font(name="Aptos Narrow", size=14, bold=True, color=DARK))]:
        ws.merge_cells(start_row=h1, start_column=rng[0], end_row=h1, end_column=rng[1])
        cell = ws.cell(h1, rng[0], text); cell.fill, cell.font, cell.alignment = fill, fnt, CWRAP
    for i, h in enumerate(COLS):
        col = 2 + i
        cell = ws.cell(h1 + 1, col, h)
        gold = col >= 12
        cell.fill = GOLD if gold else HDR
        cell.font = Font(name="Aptos", size=12, bold=True, color=DARK if gold else "FFFFFF")
        cell.alignment, cell.border = CWRAP, BOX
        sub = ws.cell(h1 + 2, col, SUBS.get(col, ""))
        sub.fill, sub.font, sub.alignment, sub.border = GREY, Font(name="Aptos Narrow", size=10, bold=True, color=DARK), CWRAP, BOX
    r0 = h1 + 3
    if ref_map is None:
        ref_map = SR25_MAP
    if tailor is None:
        tailor = sr25 is not None and ref_map is SR25_MAP
    for k, r in enumerate(rows):
        rr = r0 + k
        sref = ref_map.get(r["row"]) if sr25 is not None else None
        prev = sr25.get(sref, ("", "", "")) if sref else ("", "", "")
        request = r["request"] or "(Reporting request TBC)"
        if sr25 is not None:
            request = faculty_request(r["row"], sref, prev[0], prev[1], request, tailor=tailor,
                                      linked=bool(anchors and sref in anchors))
        if r["ref"] in quant_refs:
            request += "\n\nQuantitative figures for this indicator are requested separately in Tab 4 (Quantitative data)."
        href = sref if sref else ("N/A – new indicator" if sr25 is not None else "")
        vals = [r["ref"], r["pa"], r["target"], r["indicator"], request, use_text(r), href, prev[0], prev[1], prev[2]]
        for i, v in enumerate(vals):
            cell = ws.cell(rr, 2 + i, v)
            body(cell, REFF if i < 4 else None, bold=(i == 0))
        if anchors and sref in anchors:
            h = ws.cell(rr, 8)
            h.hyperlink = Hyperlink(ref=h.coordinate, location=f"'{SR25_SHEET}'!B{anchors[sref]}")
            h.value = f"{sref}\n→ view SR25 response"
            h.font = Font(name="Aptos Narrow", size=11, bold=True, color="0563C1", underline="single")
        for col in range(12, 19):
            body(ws.cell(rr, col))
        ws.row_dimensions[rr].height = 150
    last = r0 + len(rows) - 1
    if reg is not None:
        reg.append((key or unit_label, ws.title, r0, last + 1))  # +1 blank row: keeps single-row blocks a true range
    if rows:
        dv_list(ws, lists["status"], f"L{r0}:L{last}")
        dv_list(ws, lists["confirm"], f"Q{r0}:Q{last}")
    return last + 3


def setup_block_sheet(ws):
    widths(ws, WIDTHS)
    for col in "IJK":  # H (related SR25 ref + link) stays visible
        ws.column_dimensions[col].outlineLevel = 1
        ws.column_dimensions[col].hidden = True
    ws.sheet_properties.outlinePr.summaryRight = False


def instructions(ws, unit_name):
    ws.merge_cells("B4:C8"); ws["B4"] = "How to use this spreadsheet:"
    ws["B4"].font = Font(name="Aptos", bold=True, color="FFFFFF"); ws["B4"].fill = HDR; ws["B4"].alignment = WRAP
    ws.merge_cells("D4:G8")
    ws["D4"] = ("• Confirm target status for each row below (Column L). Refer to target status definitions to the right for guidance.\n"
                "• Provide commentary on progress towards each target (Column M or N)\n"
                "• Provide supporting documentation if required (Column O), including links to images\n"
                "• Note the source of your information and any limitations (Column P), then confirm the information is accurate (Columns Q–R)\n"
                "• Where a related end-2025 target existed, click the ref in Column H to see your full SR25 response (Tab 3)\n"
                "• Use Tab 2 of this spreadsheet to highlight key sustainability stories (optional)\n"
                f"• 2026 is a transition year: Sustainability 2030 launched ~20 October 2026, so report on activity across all of 2026 "
                f"and any early actions since launch.\n• Due: {DUE_DATE}. Questions: {CONTACT}")
    ws["D4"].alignment = WRAP; ws["D4"].font = Font(name="Aptos", size=11)
    ws["D4"].border = MBOX
    ws.merge_cells("L3:N3"); ws["L3"] = "Target status assessment and definitions"
    ws["L3"].font = Font(name="Aptos", bold=True, color="FFFFFF"); ws["L3"].fill = HDR
    for i, (s, d) in enumerate(STATUS_DEF):
        ws.cell(4 + i, 12, s).font = Font(name="Aptos", bold=True)
        ws.merge_cells(start_row=4 + i, start_column=13, end_row=4 + i, end_column=14)
        cell = ws.cell(4 + i, 13, d); cell.alignment = WRAP; cell.font = Font(name="Aptos")
        ws.row_dimensions[4 + i].height = 32
    ws["L8"].fill = GREY2
    ws["L9"] = "Note: target status definitions are under review for 2027 reporting. 2026 ratings are for internal management reporting only."
    ws["L9"].font = Font(name="Aptos", italic=True, size=9)
    ws.row_dimensions[8].height = 60


STORY_COLS = ["Title", "Description", "Related Sustainability 2030 targets (optional)", "Images and other media", "Links and supporting documentation"]


def stories_sheet(ws, unit=None):
    title(ws, "End-2026 target status assessment and reporting", "Highlighted stories (optional)")
    ws["B5"] = "How to use this spreadsheet"; ws["B5"].font = Font(name="Aptos", bold=True)
    ws.merge_cells("C5:F6")
    ws["C5"] = ("Please share up to 3 stories from 2026 which highlight your faculty's/portfolio's sustainability progress and/or impact. "
                "These stories can directly relate to Sustainability 2030 targets and indicators, or sustainability more generally. "
                f"Due: {STORIES_DUE} (early submission helps us develop case studies).")
    ws["C5"].alignment = WRAP
    cols = (["Faculty/Portfolio"] if unit is None else []) + STORY_COLS + \
           (["SST: shortlisted for case study?", "SST: permission to publish confirmed?", "SST notes"] if unit is None else [])
    hdr_row(ws, 8, cols, fill=HDR, font=Font(name="Aptos", bold=True, color="FFFFFF"))
    off = 1 if unit is None else 0
    guide = {3 + off: "• Please include a brief description of a project or activity that has made a significant contribution to sustainability at the University (up to 250 words)\n• Summarise the benefits. Benefits may relate to operational outcomes, environmental/social outcomes, and sector benefits.",
             5 + off: "(link to OneDrive/SharePoint files or published online content preferred)"}
    for i in range(len(cols)):
        c = ws.cell(9, 2 + i, guide.get(2 + i, "")); body(c, GREY)
    n = 3 if unit else 12
    for r in range(10, 10 + n):
        if unit:
            ws.cell(r, 1, r - 9)
        for i in range(len(cols)):
            body(ws.cell(r, 2 + i))
        ws.row_dimensions[r].height = 120
    widths(ws, {"A": 4, "B": 28, "C": 60, "D": 28, "E": 28, "F": 30, "G": 30, "H": 22, "I": 22, "J": 30})
    ws.row_dimensions[9].height = 90
    return 10, 10 + n - 1


PRIORITY_AREAS = ["Exceptional education", "Transformational research", "Climate leadership", "Nature and biodiversity",
                  "Responsible AI", "Estate and infrastructure", "Circular economy", "Responsible procurement",
                  "Modern slavery and human rights", "Responsible investments", "Enablers", "Not priority-area specific"]


def stories_master(wb):
    """One block of 3 story slots per unit, in the same order as each request's Tab 2 rows 10–12, so slots can be
    pasted now and linked to the request files in the Teams version. SST columns judge the priority area."""
    ws = wb.create_sheet("Highlighted stories")
    title(ws, "End-2026 target status assessment and reporting", "Highlighted stories – compiled")
    ws["B3"] = ("Each faculty/unit has 3 story slots, matching rows 10–12 of Tab 2 in its request workbook. Copy (or, in the Teams "
                "version, link) the respondent columns; SST completes the gold columns to assign priority areas and shortlist.")
    ws["B3"].alignment = WRAP; ws.merge_cells("B3:N3"); ws.row_dimensions[3].height = 32
    lst = wb["Lists"]; lst["L1"] = "Priority area"
    for i, pa in enumerate(PRIORITY_AREAS):
        lst.cell(2 + i, 12, pa)
    pa_rng = f"=Lists!$L$2:$L${1 + len(PRIORITY_AREAS)}"
    resp = ["Faculty / unit", "Story #"] + STORY_COLS
    sst = ["SST: primary priority area", "SST: secondary priority area", "SST: related S2030 indicator ref",
           "SST: use in report?", "SST: permission to publish?", "SST notes"]
    hdr_row(ws, 5, resp + sst)
    for i in range(len(resp), len(resp) + len(sst)):
        ws.cell(5, 2 + i).fill = GOLD; ws.cell(5, 2 + i).font = Font(name="Aptos", size=12, bold=True, color=DARK)
    units = [(c, f"{c} – {n}") for c, n in PILOT_FACULTIES.items()] + [(c[0], c[1]) for c in CENTRAL]
    r = 6
    for code, name in units:
        for k in range(3):
            vals = [name if k == 0 else "", k + 1] + [""] * (len(resp) - 2 + len(sst))
            for i, v in enumerate(vals):
                body(ws.cell(r, 2 + i, v), REFF if i < 2 else None, bold=i == 0)
            ws.row_dimensions[r].height = 60
            r += 1
    last = r - 1
    c0 = 2 + len(resp)
    L = lambda k: get_column_letter(c0 + k)
    dv_list(ws, pa_rng, f"{L(0)}6:{L(1)}{last}")
    dv_list(ws, '"Shortlist,Maybe,No"', f"{L(3)}6:{L(3)}{last}")
    dv_list(ws, '"Yes,No,TBC"', f"{L(4)}6:{L(4)}{last}")
    # summary: stories per priority area
    r = last + 3
    ws.cell(r, 2, "Stories by priority area (SST primary)").font = Font(name="Aptos", bold=True, size=14, color=NAVY)
    hdr_row(ws, r + 1, ["Priority area", "Stories", "Shortlisted"])
    for i, pa in enumerate(PRIORITY_AREAS):
        rr = r + 2 + i
        body(ws.cell(rr, 2, pa))
        body(ws.cell(rr, 3, f'=COUNTIF(${L(0)}$6:${L(0)}${last},B{rr})'))
        body(ws.cell(rr, 4, f'=COUNTIFS(${L(0)}$6:${L(0)}${last},B{rr},${L(3)}$6:${L(3)}${last},"Shortlist")'))
    widths(ws, {"A": 4, "B": 30, "C": 8, "D": 26, "E": 50, "F": 22, "G": 22, "H": 26, L(0): 22, L(1): 22, L(2): 14, L(3): 13, L(4): 13, L(5): 30})
    ws.freeze_panes = "D6"


# ---------------------------------------------------------------- per-unit request workbook
SR25_SHEET = "3. SR25 responses"


def sr25_sheet(wb, name, sr25_sections, ref_map, qual_by_row):
    """Read-only copy of the unit's end-2025 responses. Returns {SR25 ref: row} for hyperlinks."""
    ws = wb.create_sheet(SR25_SHEET)
    title(ws, "End-2025 target status assessment and reporting (for reference)", f"Your SR25 responses – {name}")
    ws["B3"] = ("Your responses to the end-2025 reporting request, against the former Sustainability Plan 2030 (SP2030) targets. "
                "For reference only – please do not edit. Use the 'Back to request' links to return to the reporting template.")
    ws["B3"].alignment = WRAP; ws.merge_cells("B3:I3"); ws.row_dimensions[3].height = 32
    rev = {}
    for row, ref in ref_map.items():
        if row in qual_by_row:
            rev.setdefault(ref, []).append(qual_by_row[row]["ref"])
    cols = ["SR25 ref", "SP2030 target", "2025 reporting request", "Confirmed end-2025 status", "OPTION 1: Target-level reporting",
            "OPTION 2: Indicator-level reporting", "Supporting information/comments", "Related SR26 indicator(s)"]
    anchors, answered, r = {}, {}, 5
    if not sr25_sections:
        ws.cell(r, 2, "No end-2025 request was issued to this area (new to Sustainability 2030 reporting).").font = Font(name="Aptos", italic=True)
    for sec, rows in sr25_sections:
        ws.cell(r, 2, sec).font = Font(name="Aptos Narrow", size=16, bold=True); ws.cell(r, 2).fill = GREY
        r += 1
        hdr_row(ws, r, cols); r += 1
        for x in rows:
            vals = [x["ref"], x["target"], x["request"], x["status"] or "Not provided", x["opt1"], x["opt2"], x["support"],
                    ", ".join(rev.get(x["ref"], []))]
            for i, v in enumerate(vals):
                body(ws.cell(r, 2 + i, v), REFF if i == 0 else None, bold=i == 0)
            if x["ref"] not in anchors or (x["status"] or x["opt1"] or x["opt2"]) and not answered.get(x["ref"]):
                anchors[x["ref"]] = r  # link to the answered occurrence where a ref appears twice
                answered[x["ref"]] = bool(x["status"] or x["opt1"] or x["opt2"])
            ws.row_dimensions[r].height = 160
            r += 1
        r += 1
    widths(ws, {"A": 4, "B": 10, "C": 34, "D": 36, "E": 16, "F": 60, "G": 60, "H": 34, "I": 16})
    back = ws.cell(4, 2, "← Back to request"); back.hyperlink = Hyperlink(ref="B4", location="'1. Reporting template'!A1")
    back.font = Font(name="Aptos", color="0563C1", underline="single", bold=True)
    ws.freeze_panes = "C5"
    return anchors


def request_workbook(code, name, sections, sr25, ref_map=None, sr25_sections=None, qual_by_row=None, folder="faculty_requests"):
    wb = Workbook(); ws = wb.active; ws.title = "1. Reporting template"
    lists = lists_sheet(wb)
    quant_refs = quant.unit_quant_tab(wb, code, name, quant.owned_points(DATABOOK, code), DUE_DATE) if folder != "faculty_requests" else set()
    anchors = sr25_sheet(wb, name, sr25_sections or [], ref_map if ref_map is not None else SR25_MAP, qual_by_row or {})
    title(ws, "End-2026 target status assessment and reporting", "End-2026 target status confirmation and reporting")
    setup_block_sheet(ws); instructions(ws, name)
    top = 11
    for label, rows in sections:
        top = block(ws, top, label, rows, sr25, lists, ref_map=ref_map, anchors=anchors, quant_refs=quant_refs)
    ws.freeze_panes = "C14"
    st = wb.create_sheet("2. OPTIONAL highlighted stories", 1)
    stories_sheet(st, unit=code)
    order = ["1. Reporting template", "2. OPTIONAL highlighted stories", SR25_SHEET, "4. Quantitative data", "Lists"]
    wb._sheets.sort(key=lambda w: order.index(w.title))
    wb.active = 0
    safe = file_code(code)
    path = OUT / folder / f"{safe} - end-2026 sustainability reporting request - {VERSION}.xlsx"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path.name


# ---------------------------------------------------------------- master workbook
def master(qual, quan, fac_rows, central_sections, sr25_all):
    wb = Workbook()
    lists = None

    # 0. Read me
    ws = wb.active; ws.title = "0. Read me"
    title(ws, "SR26 data collection master spreadsheet", "2026 Sustainability Report – data collection master")
    readme = [
        ("Purpose", "Central record of SR26 data requirements, owners, reporting requests, responses and evidence. Adapted from the 2025 end-year reporting master spreadsheet (V0.2) and the Sustainability 2030 updated targets (Qual and Quan worksheets)."),
        ("Status", f"DRAFT {VERSION} – pilot with {len(PILOT_FACULTIES)} faculties ({', '.join(PILOT_FACULTIES)}) and {len(CENTRAL)} central units ({', '.join(c[0] for c in CENTRAL)})."),
        ("Reporting approach", "Transition year: cover all of CY2026, distinguishing former Sustainability Plan 2030 activity from foundations and early actions after the Sustainability 2030 launch (~20 Oct 2026). Databook is the authoritative quantitative source. Target status (Met/Partially met/Not met) retained for now for internal management reporting – rating framework under review."),
        ("Key dates (TBC)", f"Requests issued: Nov–Dec 2026 | Highlighted stories due: {STORIES_DUE} | Reporting template due: {DUE_DATE} | Support meetings: January 2027 | Sustainability Reporting Review Group: ~15 Feb 2027 | VCAG: 16 Feb & ~2 Mar 2027"),
        ("Tabs", "1. Requirements matrix – every S2030 qualitative indicator, owner and request\n"
                 "2. Quant coverage – every quantitative indicator checked against the Databook draft, owner and action\n"
                 "2b. Databook register – every Databook data point with columns to confirm definition, source, owner\n"
                 "3. Stakeholder map – SR25 → SR26 unit mapping\n"
                 "3b. RASCI matrix – A/R/S/C/I roles per indicator and area (SR25 RASCI format), with counts\n"
                 "3c. RASCI triangulation – SR26 roles cross-checked against the SR25 RASCI, gaps flagged\n"
                 "4. Request tracker – sent/chased/received/confirmed status and response rate\n"
                 "Faculties / CI&S / CFOG / ESG / CIOG / MRE / L&R – consolidated responses (copy in from returned request workbooks; same layout)\n"
                 "Highlighted stories – consolidated optional stories and case-study shortlist\n"
                 "A1. Data scoring – auto-pulls every unit's responses per indicator and scores target status (SR25 method)\n"
                 "0.5 Target Check – which units were asked about each indicator\n"
                 "T1. Manual target review – every indicator with responses from the units asked, side by side, plus reviewer notes\n"
                 "T2. Single target review – pick an indicator and see all units' responses and the scored status\n"
                 "Request wording log – original S2030 wording vs tailored faculty request wording\n"
                 "Evidence register – source, evidence, limitations and owner confirmation for every material claim\n"
                 "Issues log – open questions, TBCs and decisions\n"
                 "Project management tabs (overview, plan, timeline, comms/web to-do, content checklist, page index) are in the separate 'SR26 project management' workbook"),
        ("Refs", "Refs are new SR26 codes: priority-area code + target number + indicator letter (e.g. EE1(a) = Exceptional education, target 1, indicator a). 'Related SR25 ref' links to the closest SP2030 (2025) target where one exists."),
        ("Regenerating", "Workbooks are generated by build.py in the repository. Edit the settings at the top (due dates, pilot faculties, unit mapping) and re-run to produce request workbooks for additional faculties/units."),
        ("Version history", "\n".join(f"{v}: {t}" for v, t in VERSION_HISTORY)),
        ("Request wording", "Faculty requests are rewritten in the SR25 style and point to each faculty's end-2025 response (see Request wording log). Central-unit requests use the S2030 'Qual' worksheet wording as-is."),
        ("Formulas", "Responses must be pasted into the Faculties/central-unit tabs keeping the same rows – A1, 0.5 and T2 look them up by Ref. Do not insert rows inside a block; add new units by re-running build.py."),
    ]
    for i, (k, v) in enumerate(readme):
        a = ws.cell(4 + i, 2, k); a.font = Font(name="Aptos", bold=True, color="FFFFFF"); a.fill = HDR; a.alignment = WRAP; a.border = BOX
        b = ws.cell(4 + i, 3, v); b.alignment = WRAP; b.border = BOX; b.font = Font(name="Aptos")
        ws.row_dimensions[4 + i].height = max(30, 15 * (v.count("\n") + 1 + len(v) // 110))
    widths(ws, {"A": 4, "B": 22, "C": 120})

    # 1. Requirements matrix
    ws = wb.create_sheet("1. Requirements matrix")
    title(ws, "End-2026 target status assessment and reporting", "1. Requirements matrix (qualitative)")
    cols = ["Ref", "Domain", "Priority area", "Target", "Indicator", "Reporting requirement (S2030 Qual sheet)",
            "Key stakeholders (S2030 sheet)", "SR26 request issued to", "Specific projects we are aware of (area & person)",
            "Internal vs External", "Flag for continuous improvement", "Notes", "Related SR25 ref", "Channel",
            "Quantitative component", "Owner confirmed?"]
    hdr_row(ws, 5, cols)
    qcomp = {c["ref"]: c for c in COVERAGE}
    for k, r in enumerate(qual):
        rr = 6 + k
        to = [u[0] for u in CENTRAL for s in u[3] if s[1](r)]
        if asks_faculties(r):
            to = ["Faculties (pilot)"] + to
        channel = "Tailored request" if to else "TBC"
        qc = qcomp.get(r["ref"])
        qtxt = f"Yes – {qc['indb']}; see 2. Quant coverage" if qc else ""
        vals = [r["ref"], r["domain"], r["pa"], r["target"], r["indicator"], r["request"], r["stake"],
                ", ".join(dict.fromkeys(to)) or "TBC – no owner identified", r["projects"], r["intext"], r["ci"], r["notes"],
                SR25_MAP.get(r["row"], ""), channel, qtxt, ""]
        for i, v in enumerate(vals):
            body(ws.cell(rr, 2 + i, v), REFF if i == 0 else None, bold=i == 0)
        ws.row_dimensions[rr].height = 90
    ws.auto_filter.ref = f"B5:{get_column_letter(1 + len(cols))}{5 + len(qual)}"
    ws.freeze_panes = "D6"
    widths(ws, dict(zip("BCDEFGHIJKLMNOPQ", [9, 16, 18, 40, 40, 50, 18, 22, 40, 11, 12, 30, 10, 18, 24, 12])))

    # 2. Quant coverage + 2b. Databook register
    quant.coverage_tab(wb, COVERAGE)
    quant.register_tab(wb, DATABOOK)

    # 3. Stakeholder map & RASCI
    ws = wb.create_sheet("3. Stakeholder map")
    title(ws, "End-2026 target status assessment and reporting", "3. Stakeholder map (SR25 → SR26)")
    hdr_row(ws, 4, ["SR26 unit", "SR26 full name", "SR25 equivalent (2025 master)", "Section / team", "Key contact", "Status / notes"])
    units = [("Faculty", name, f"Faculties tab – {code}", "Associate Dean Sustainability / faculty sustainability lead", "TBC", "Pilot faculty") for code, name in PILOT_FACULTIES.items()]
    for code, name, sr25name, secs in CENTRAL:
        for label, _ in secs:
            units.append((code, name, sr25name, label, "TBC", ""))
    units += [
        ("TBC", "CGOP / CDEP / CDSS", "Not in SR25 master", "—", "TBC", "Structure changes TBC – confirm before requests are issued."),
    ]
    for k, u in enumerate(units):
        for i, v in enumerate(u):
            body(ws.cell(5 + k, 2 + i, v))
    widths(ws, {"A": 4, "B": 14, "C": 40, "D": 44, "E": 40, "F": 18, "G": 50})
    ws.cell(7 + len(units), 2, "Who is responsible for each indicator: see '3b. RASCI matrix'.").font = Font(name="Aptos", italic=True)
    rasci.rasci_tab(wb, QUAL_ALL, COVERAGE, PILOT_FACULTIES, CENTRAL)
    print("  SR25-only gaps by area:", rasci.triangulation_tab(wb, QUAL_ALL))

    # 4. Request tracker
    ws = wb.create_sheet("4. Request tracker")
    title(ws, "End-2026 target status assessment and reporting", "4. Request tracker")
    cols = ["#", "Group", "Unit / section", "Key contact", "No. of indicators requested", "Request workbook", "Date sent",
            "Due date", "Reminder date", "Support meeting offered/held", "Response received", "Owner confirmed",
            "Status", "Days overdue", "Notes / follow-up"]
    hdr_row(ws, 9, cols)
    track = [(f"F{i+1}", "Faculty", f"{code} – {name}", len(fac_rows), f"{code} - end-2026 sustainability reporting request - {VERSION}.xlsx")
             for i, (code, name) in enumerate(PILOT_FACULTIES.items())]
    n = 0
    for code, name, _, secs in CENTRAL:
        for label, rows in central_sections[code]:
            n += 1
            track.append((f"C{n}", "Central", label, len(rows), f"{file_code(code)} - end-2026 sustainability reporting request - {VERSION}.xlsx"))
    for k, t in enumerate(track):
        rr = 10 + k
        vals = [t[0], t[1], t[2], "TBC", t[3], t[4], None, "15/01/2027", f"=IF(I{rr}=\"\",\"\",I{rr}-7)", "", None, "", "Not sent",
                f"=IF(OR(I{rr}=\"\",L{rr}<>\"\"),\"\",MAX(0,TODAY()-I{rr}))", ""]
        for i, v in enumerate(vals):
            body(ws.cell(rr, 2 + i, v))
        from datetime import date
        ws.cell(rr, 9).value = date(2027, 1, 15)
        for col in (8, 9, 10, 12):
            ws.cell(rr, col).number_format = "dd/mm/yyyy"
    last = 9 + len(track)
    dv_list(ws, "=Lists!$D$2:$D$9", f"N10:N{last}")
    dv_list(ws, "=Lists!$E$2:$E$4", f"M10:M{last}")
    # summary
    ws["B4"] = "Summary"; ws["B4"].font = Font(name="Aptos", bold=True, size=14, color=NAVY)
    for i, s in enumerate(TRACK_STATUS):
        ws.cell(5 + i // 4 * 2, 3 + (i % 4) * 2, s).font = Font(name="Aptos", bold=True)
        ws.cell(6 + i // 4 * 2, 3 + (i % 4) * 2, f'=COUNTIF($N$10:$N${last},"{s}")')
    ws["L5"] = "Response rate"; ws["L5"].font = Font(name="Aptos", bold=True)
    ws["L6"] = f'=IFERROR((COUNTIF($N$10:$N${last},"Received")+COUNTIF($N$10:$N${last},"Under review")+COUNTIF($N$10:$N${last},"Complete"))/(COUNTA($N$10:$N${last})-COUNTIF($N$10:$N${last},"Not required")),0)'
    ws["L6"].number_format = "0%"
    ws["N5"] = "Owner confirmed"; ws["N5"].font = Font(name="Aptos", bold=True)
    ws["N6"] = f'=COUNTIF($M$10:$M${last},"Yes")&" / "&COUNTA($C$10:$C${last})'
    widths(ws, dict(zip("ABCDEFGHIJKLMNOP", [4, 6, 10, 46, 16, 12, 44, 12, 12, 12, 16, 14, 12, 16, 10, 40])))
    ws.freeze_panes = "E10"

    # Unit tabs (consolidated responses)
    reg = []

    def unit_tab(tab, subtitle, sections, sr25_for):
        ws = wb.create_sheet(tab)
        title(ws, "End-2026 target status assessment and reporting", subtitle)
        setup_block_sheet(ws)
        ws["B3"] = "Consolidated responses – paste from returned request workbooks (identical column layout). Click [+] above Columns H–K for end-2025 responses."
        ws["B3"].font = Font(name="Aptos", italic=True)
        top = 5
        for idx, (key, label, rows, sr, rmap) in enumerate(sections):
            top = block(ws, top, f"1.{idx + 1}  {label}", rows, sr, LISTS, reg, key, ref_map=rmap)
        return ws

    global LISTS
    lists = lists_sheet(wb); LISTS = lists
    unit_tab("Faculties", "Faculties", [(c, f"{c} – {n}", fac_rows, prev_lookup(sr25_for(c)), None) for c, n in PILOT_FACULTIES.items()], True)
    for code, name, _, secs in CENTRAL:
        prev = prev_lookup(sr25_for(code)) if SR25_SOURCES.get(code) else None
        unit_tab(code, name, [(label.split(" (")[0], label, rows, prev, CENTRAL_SR25_MAP[code]) for label, rows in central_sections[code] if rows], False)

    scoring_tabs(wb, qual, reg)
    wording_log(wb, fac_rows)

    # Highlighted stories – compiled from every request's Tab 2, with SST assessment columns
    stories_master(wb)

    # Evidence register – kept simple: one line per material claim or figure used in the report
    ws = wb.create_sheet("Evidence register")
    title(ws, "End-2026 target status assessment and reporting", "Evidence register")
    ws["B3"] = "One line per material claim or figure used in the report: what it supports, where the evidence is, and whether the owner has confirmed it."
    ws["B3"].font = Font(name="Aptos", italic=True)
    cols = ["Evidence ID", "Ref", "Faculty / unit", "Claim or figure in the report", "Evidence (link or file)", "Source",
            "Owner confirmed?", "Notes"]
    hdr_row(ws, 5, cols)
    for r in range(6, 206):
        ws.cell(r, 2, f'=IF(E{r}="","","EV-"&TEXT(ROW()-5,"000"))')
        for c in range(2, 2 + len(cols)):
            body(ws.cell(r, c))
    dv_list(ws, "=Lists!$E$2:$E$4", "H6:H205")
    widths(ws, dict(zip("BCDEFGHI", [10, 9, 18, 50, 40, 20, 12, 30])))
    ws.freeze_panes = "D6"

    # Issues log
    ws = wb.create_sheet("Issues log")
    title(ws, "End-2026 target status assessment and reporting", "Issues log")
    hdr_row(ws, 4, ["#", "Date raised", "Issue / question", "Impact", "Ref / unit", "Owner", "Proposed action", "Status"])
    issues = [
        ("Climate resilience maturity (CL3(a)) and Estate & infrastructure indicators – confirm owner and request with Gerard.", "CL3(a); EI1(a)–(d)", "Stefanus", "Meet Gerard", "Open"),
        ("Confirm whether CGOP, CDEP and CDSS structures have changed since the 2025 master spreadsheet stakeholder mapping.", "Stakeholder map", "Stefanus", "Confirm with team", "Open"),
        ("Business Services split into CIOG (AI) and ESG (nature & biodiversity, waste & circular economy, part of quant. climate leadership) – confirm contacts.", "CIOG; ESG", "TBC", "Confirm contacts", "Open"),
        ("Procurement sits with CFOG (RP1(a) procurement case studies routed to CFOG – Procurement although the Qual sheet names no stakeholder); estate planning centralised in CI&S – requests issued centrally, not to faculties.", "CFOG; CI&S", "—", "Noted", "Closed"),
        ("Chancellery units reported separately, as SR25: AC&M, MRE, GCE, Education, SaSS and Indigenous each have their own request and master tab. Alumni (EE1(d)) asked of SaSS only.", "Stakeholder map", "Stefanus", "Confirm contacts", "Closed"),
        ("Faculty requests cut down after ABP walkthrough: EE1(a)+(c) merged; CL2(a)/(c) to AC&M only; TR3(a)/(c)/(d) to MRE only; EE1(d) to SaSS only. Wattle Fellowship to be asked directly (not tracked in master).", "Faculties", "Stefanus", "Confirm with team; apply to other faculties", "Open"),
        ("Legal & Risk asked for climate resilience maturity CL3(a), following SR25 (8a(i)/(ii): University Risk 16 Climate Change; flood emergency response plans). CL3(a) is also with CI&S (Gerard) – agree who leads.", "CL3(a); L&R; CI&S", "Stefanus", "Confirm with Gerard", "Open"),
        ("TR2(b) 'Documented progress of strategic initiatives, incl. Impact Accelerators' reuses the TR2(d) case-study wording and has no stakeholder in the Qual sheet – left as-is.", "TR2(b)", "TBC", "Review wording/owner", "Open"),
        ("Quantitative-only rows (no stakeholder in the Qual sheet) removed from the requirements matrix and moved to '2. Quant coverage'; Databook gaps requested in Tab 4 of the owning central unit's request.", "TR2(b); CL1(a)-(b); CL2(b); NB1(a); CE1-2; RP2; RI1(a)", "Stefanus", "Confirm owners and Databook coverage with Chris", "Open"),
        ("Target status rating (Met/Partially met/Not met) retained pending revamp of the traffic-light framework.", "All", "Director, Sustainability", "Update templates once agreed", "Open"),
        ("Due dates are placeholders based on the SR26 approach timeline (stories early Dec; requests due January).", "All", "TBC", "Confirm dates", "Open"),
        ("20 indicators are requested from each faculty (SR25 sent 10). Monitor burden in pilot.", "Faculties", "TBC", "Review after pilot", "Open"),
    ]
    from datetime import date
    for k, it in enumerate(issues):
        vals = [k + 1, date(2026, 10, 5), it[0], ""] + list(it[1:])  # Impact column as SR25 issues log
        for i, v in enumerate(vals):
            body(ws.cell(5 + k, 2 + i, v))
        ws.cell(5 + k, 3).number_format = "dd/mm/yyyy"
    for r in range(5 + len(issues), 40):  # blank rows for new issues
        for i in range(8):
            body(ws.cell(r, 2 + i))
    dv_list(ws, '"Open,In progress,Closed"', "I5:I39")
    widths(ws, {"A": 4, "B": 5, "C": 12, "D": 60, "E": 30, "F": 24, "G": 18, "H": 26, "I": 10})

    wb.move_sheet("Lists", offset=len(wb.sheetnames))
    OUT.mkdir(exist_ok=True)
    p = OUT / f"SR26 end-year reporting master spreadsheet - {VERSION}.xlsx"
    wb.save(p)
    return p.name


# ---------------------------------------------------------------- scoring engine (adapted from SR25 A1 / 0.5 / T2 tabs)
FIELDS = [("Confirm end-2026 target status", "L"), ("OPTION 1:  Target-level reporting", "M"),
          ("OPTION 2:  Indicator level reporting", "N"), ("Supporting information/comments", "O"),
          ("Data source and limitations", "P"), ("Data owner confirms", "Q")]
AUTOFILL = "Cell will autofill based on response"


def scoring_tabs(wb, qual, reg):
    units = []  # (key, [(sheet, r0, last)]) – a unit may appear once per sheet
    for key, sheet, r0, last in reg:
        units.append((key, sheet, r0, last))
    n_u = len(units)
    uc = lambda i: get_column_letter(8 + i)            # unit columns start at H
    lastu = uc(n_u - 1)

    # ---- A1. Data scoring
    ws = wb.create_sheet("A1. Data scoring")
    title(ws, "End-2026 target status assessment and reporting", "A1. Data scoring summary tab")
    ws["C4"] = "Description"; ws["C4"].font = Font(name="Aptos", bold=True)
    ws["D4"] = ("Pulls every unit's end-2026 response for each indicator from the Faculties and central-unit tabs (do not type in "
                "the grey cells). 'N/A' = not requested from that unit. Status is scored Met or exceeded = 3, Partially met = 2, "
                "Not met = 1, Not yet started = 0 (as SR25). Threshold status applies the SR25 final assessment rule: >75% 'met' = "
                "Met or exceeded; 50–75% 'met' or 'partially met' = Partially met; otherwise Not met.")
    ws["D4"].alignment = WRAP; ws.merge_cells("D4:N6")
    groups = ["Faculty" if k in PILOT_FACULTIES else "Central" for k, *_ in units]
    head = ["Ref.", "Priority area", "Indicator", "Reporting requirement", "Target code", "Target progress"]
    for i, h in enumerate(head):
        ws.cell(10, 2 + i, h)
    for i, (k, *_r) in enumerate(units):
        ws.cell(9, 8 + i, groups[i]); ws.cell(10, 8 + i, k)
    sc0 = 8 + n_u + 1
    score_cols = ["Requested, no response received", "Not yet started", "Not met", "Partially met", "Met or exceeded",
                  "Total score", "Total valid responses", "Target score (average)", "Target score (text)",
                  "% met", "% met or partially met", "Threshold status (SR25 rule)"]
    for i, h in enumerate(score_cols):
        ws.cell(10, sc0 + i, h)
    hdr_row(ws, 10, [ws.cell(10, c).value for c in range(2, sc0 + len(score_cols))])
    for c in range(8, 8 + n_u):
        ws.cell(9, c).font = Font(name="Aptos", italic=True, size=9)
    C = {h: get_column_letter(sc0 + i) for i, h in enumerate(score_cols)}
    r = 11
    status_rows = {}
    for q in qual:
        for fi, (fname, col) in enumerate(FIELDS):
            ws.cell(r, 1, f'=F{r}&"|"&G{r}')
            vals = [q["ref"] if fi == 0 else "", q["pa"] if fi == 0 else "", q["indicator"] if fi == 0 else "",
                    q["request"] if fi == 0 else "", q["ref"], fname]
            for i, v in enumerate(vals):
                body(ws.cell(r, 2 + i, v), REFF if fi == 0 else None, bold=(i == 0))
            for i, (k, sheet, r0, last) in enumerate(units):
                rng = f"'{sheet}'!${col}${r0}:${col}${last}"
                key = f"'{sheet}'!$B${r0}:$B${last}"
                f = (f'=IFERROR(IF(INDEX({rng},MATCH($F{r},{key},0))="","{AUTOFILL}",'
                     f'INDEX({rng},MATCH($F{r},{key},0))),"N/A")')
                c = ws.cell(r, 8 + i, f); body(c, GREY)
                c.alignment = Alignment(wrap_text=False, vertical="top")
            if fi == 0:
                status_rows[q["ref"]] = r
                rowrng = f"$H{r}:${lastu}{r}"
                fs = {
                    "Requested, no response received": f'=COUNTIF({rowrng},"{AUTOFILL}")',
                    "Not yet started": f'=COUNTIF({rowrng},"Not yet started")',
                    "Not met": f'=COUNTIF({rowrng},"Not met")',
                    "Partially met": f'=COUNTIF({rowrng},"Partially met")',
                    "Met or exceeded": f'=COUNTIF({rowrng},"Met or exceeded")',
                }
                fs["Total score"] = f'=3*{C["Met or exceeded"]}{r}+2*{C["Partially met"]}{r}+{C["Not met"]}{r}'
                fs["Total valid responses"] = f'={C["Met or exceeded"]}{r}+{C["Partially met"]}{r}+{C["Not met"]}{r}+{C["Not yet started"]}{r}'
                v = f'{C["Total valid responses"]}{r}'
                fs["Target score (average)"] = f'=IF({v}>0,{C["Total score"]}{r}/{v},"No valid responses to date")'
                a = f'{C["Target score (average)"]}{r}'
                fs["Target score (text)"] = f'=IF(ISNUMBER({a}),INDEX(Lists!$G$2:$G$5,MATCH(ROUND({a},0),Lists!$H$2:$H$5,0)),"")'
                fs["% met"] = f'=IF({v}>0,{C["Met or exceeded"]}{r}/{v},"")'
                fs["% met or partially met"] = f'=IF({v}>0,({C["Met or exceeded"]}{r}+{C["Partially met"]}{r})/{v},"")'
                pm, pmp = f'{C["% met"]}{r}', f'{C["% met or partially met"]}{r}'
                fs["Threshold status (SR25 rule)"] = (f'=IF({v}=0,"",IF({pm}>0.75,"Met or exceeded",'
                                                      f'IF({pmp}>=0.5,"Partially met","Not met")))')
                for h, f in fs.items():
                    c = ws.cell(r, sc0 + score_cols.index(h), f); body(c, PatternFill("solid", fgColor="E2EFDA"))
                    if h.startswith("%"):
                        c.number_format = "0%"
                    if h == "Target score (average)":
                        c.number_format = "0.00"
            r += 1
    last_a1 = r - 1
    ws.column_dimensions["A"].hidden = True
    widths(ws, {"B": 9, "C": 16, "D": 36, "E": 36, "F": 9, "G": 30})
    for i in range(n_u):
        ws.column_dimensions[uc(i)].width = 18
    for i in range(len(score_cols)):
        ws.column_dimensions[get_column_letter(sc0 + i)].width = 13
    ws.freeze_panes = "H11"
    ws.auto_filter.ref = f"B10:{get_column_letter(sc0 + len(score_cols) - 1)}{last_a1}"

    # ---- 0.5 Target Check
    tc = wb.create_sheet("0.5 Target Check")
    title(tc, "End-2026 target status assessment and reporting", "0.5 Target Check")
    tc["B3"] = "Which units were asked to report on each indicator (Yes = indicator appears in that unit's tab). Feeds T2."
    hdr_row(tc, 6, ["Target list"] + [k for k, *_ in units] + ["No. of units requested"])
    for k, q in enumerate(qual):
        rr = 7 + k
        sr = status_rows[q["ref"]]
        body(tc.cell(rr, 2, q["ref"]), REFF, bold=True)
        for i in range(n_u):
            c = tc.cell(rr, 3 + i, f"=IF('A1. Data scoring'!{uc(i)}{sr}<>\"N/A\",\"Yes\",\"No\")"); body(c)
            c.alignment = CWRAP
        body(tc.cell(rr, 3 + n_u, f'=COUNTIF(C{rr}:{get_column_letter(2 + n_u)}{rr},"Yes")'))
    widths(tc, {"B": 10})
    for i in range(n_u + 1):
        tc.column_dimensions[get_column_letter(3 + i)].width = 14
    tc.freeze_panes = "C7"
    last_tc = 6 + len(qual)

    # ---- T1. Manual target review: every indicator, only the units asked, side by side (base review sheet)
    t1 = wb.create_sheet("T1. Manual target review")
    title(t1, "End-2026 target status assessment and reporting", "T1. Manual target review")
    t1["C4"] = "Description"; t1["C4"].font = Font(name="Aptos", bold=True)
    t1["D4"] = ("Review of individual unit responses for every indicator – only the units asked about each indicator are shown. "
                "Values come from A1 (do not type in grey cells); record your review in the 'Reviewer notes' column. "
                "Use T2 to focus on a single indicator.")
    t1["D4"].alignment = WRAP; t1.merge_cells("D4:J5"); t1.row_dimensions[4].height = 30
    asked = {}
    for i, (k, sheet, r0, last) in enumerate(units):
        for rr in range(r0, last + 1):
            v = wb[sheet].cell(rr, 2).value
            if v:
                asked.setdefault(v, []).append(i)
    maxu = max((len(v) for v in asked.values()), default=1)
    ucol = lambda j: get_column_letter(5 + j)
    sc = [get_column_letter(5 + maxu + i) for i in range(3)]
    hdr_row(t1, 7, ["Ref.", "Indicator", "Target progress"] + [f"Unit {j + 1}" for j in range(maxu)] +
            ["Target score (text)", "Threshold status (SR25 rule)", "Reviewer notes"])
    r = 8
    for q in qual:
        idx = asked.get(q["ref"], [])
        if not idx:
            continue
        a1 = status_rows[q["ref"]]
        body(t1.cell(r, 2, q["ref"]), REFF, bold=True)
        body(t1.cell(r, 3, q["indicator"]), REFF)
        body(t1.cell(r, 4, "Unit"), GREY2, bold=True)
        for j, ui in enumerate(idx):
            body(t1.cell(r, 5 + j, units[ui][0]), GREY2, bold=True)
        for col, h in zip(sc[:2], ["Target score (text)", "Threshold status (SR25 rule)"]):
            c = t1[f"{col}{r}"]; c.value = f"='A1. Data scoring'!{C[h]}{a1}"; body(c, PatternFill("solid", fgColor="E2EFDA"), bold=True)
        body(t1[f"{sc[2]}{r}"])
        for fi, (fname, _) in enumerate(FIELDS):
            rr = r + 1 + fi
            body(t1.cell(rr, 2, q["ref"])); t1.cell(rr, 2).font = Font(name="Aptos Narrow", size=9, color="808080")
            body(t1.cell(rr, 4, fname), bold=True)
            for j, ui in enumerate(idx):
                body(t1.cell(rr, 5 + j, f"='A1. Data scoring'!{uc(ui)}{a1 + fi}"), GREY)
            body(t1[f"{sc[2]}{rr}"])
            t1.row_dimensions[rr].height = 60 if fi in (1, 2) else 18
        t1.merge_cells(start_row=r + 1, start_column=3, end_row=r + len(FIELDS), end_column=3)
        t1.merge_cells(f"{sc[2]}{r}:{sc[2]}{r + len(FIELDS)}")
        r += len(FIELDS) + 2
    widths(t1, {"A": 4, "B": 9, "C": 34, "D": 26, sc[0]: 16, sc[1]: 16, sc[2]: 40})
    for j in range(maxu):
        t1.column_dimensions[ucol(j)].width = 30
    t1.freeze_panes = "E8"
    t1.auto_filter.ref = f"B7:{sc[2]}{r}"

    # ---- T2. Single target review
    t2 = wb.create_sheet("T2. Single target review")
    title(t2, "End-2026 target status assessment and reporting", "T2. Single target review")
    t2["B4"] = "Target selected"; t2["B4"].font = Font(name="Aptos", bold=True)
    t2["C4"] = qual[0]["ref"]; t2["C4"].fill = GOLD; t2["C4"].font = Font(name="Aptos", bold=True, size=14); t2["C4"].border = MBOX
    t2["D4"] = "← select an indicator ref from the dropdown"
    rm = "'1. Requirements matrix'"
    dv_list(t2, f"={rm}!$B$6:$B${5 + len(qual)}", "C4")
    info = [("Priority area", "D"), ("Target", "E"), ("Indicator", "F"), ("Reporting requirement", "G"), ("Request issued to", "I")]
    for i, (lbl, col) in enumerate(info):
        a = t2.cell(5 + i, 2, lbl); a.font = Font(name="Aptos", bold=True, color="FFFFFF"); a.fill = HDR; a.border = BOX
        t2.merge_cells(start_row=5 + i, start_column=3, end_row=5 + i, end_column=8)
        c = t2.cell(5 + i, 3, f"=IFERROR(INDEX({rm}!${col}$6:${col}${5 + len(qual)},MATCH($C$4,{rm}!$B$6:$B${5 + len(qual)},0)),\"\")")
        c.alignment = WRAP; c.border = BOX
        t2.row_dimensions[5 + i].height = 45
    srow = 11
    for i, h in enumerate(["Target score (text)", "Threshold status (SR25 rule)", "% met", "Total valid responses", "Requested, no response received"]):
        a = t2.cell(srow, 2 + i * 2, h); a.font = Font(name="Aptos", bold=True); a.alignment = WRAP
        c = t2.cell(srow + 1, 2 + i * 2, f"=IFERROR(INDEX('A1. Data scoring'!${C[h]}$11:${C[h]}${last_a1},"
                                         f"MATCH($C$4&\"|\"&\"{FIELDS[0][0]}\",'A1. Data scoring'!$A$11:$A${last_a1},0)),\"\")")
        c.font = Font(name="Aptos", bold=True, size=13, color=NAVY)
        if h == "% met":
            c.number_format = "0%"
    hr = 14
    hdr_row(t2, hr, ["Target progress"] + [k for k, *_ in units])
    body(t2.cell(hr + 1, 2, "Requested?"), GREY2, bold=True)
    for i in range(n_u):
        c = t2.cell(hr + 1, 3 + i, f"=IFERROR(INDEX('0.5 Target Check'!{get_column_letter(3 + i)}$7:{get_column_letter(3 + i)}${last_tc},"
                                    f"MATCH($C$4,'0.5 Target Check'!$B$7:$B${last_tc},0)),\"\")")
        body(c, GREY2); c.alignment = CWRAP
    for fi, (fname, _) in enumerate(FIELDS):
        rr = hr + 2 + fi
        body(t2.cell(rr, 2, fname), REFF, bold=True)
        for i in range(n_u):
            c = t2.cell(rr, 3 + i, f"=IFERROR(INDEX('A1. Data scoring'!{uc(i)}$11:{uc(i)}${last_a1},"
                                     f"MATCH($C$4&\"|\"&$B{rr},'A1. Data scoring'!$A$11:$A${last_a1},0)),\"\")")
            body(c)
        t2.row_dimensions[rr].height = 120 if fi in (1, 2) else 45
    widths(t2, {"B": 26})
    for i in range(n_u):
        t2.column_dimensions[get_column_letter(3 + i)].width = 30
    t2.column_dimensions["C"].width = 30
    t2.freeze_panes = "C15"


def wording_log(wb, fac_rows):
    ws = wb.create_sheet("Request wording log")
    title(ws, "End-2026 target status assessment and reporting", "Faculty request wording log")
    ws["B3"] = ("Faculty requests were rewritten in the SR25 style (specific ask, scope, what SST already holds or will collect elsewhere, "
                "and a pointer to the faculty's own end-2025 response). The original S2030 Qual-sheet wording is kept here; central-unit "
                "requests still use the original wording.")
    ws["B3"].alignment = WRAP; ws.merge_cells("B3:G3"); ws.row_dimensions[3].height = 45
    hdr_row(ws, 5, ["Ref", "Indicator", "Original wording (S2030 Qual sheet)", "SR26 faculty request (before faculty-specific opening line)",
                    "Related SR25 ref", "Faculty-specific opening line"])
    from faculty_requests import TAILORED
    for k, r in enumerate(fac_rows):
        req, note = TAILORED.get(r["row"], (r["request"], ""))
        vals = [r["ref"], r["indicator"], r["request"], req + ("\n\n" + note if note else ""), SR25_MAP.get(r["row"], "N/A – new indicator"),
                "Points to the faculty's end-2025 response and rating" if r["row"] in SR25_MAP else "None (new indicator)"]
        for i, v in enumerate(vals):
            body(ws.cell(6 + k, 2 + i, v), REFF if i == 0 else None, bold=i == 0)
        ws.row_dimensions[6 + k].height = 130
    widths(ws, {"B": 9, "C": 36, "D": 50, "E": 60, "F": 12, "G": 28})


from project_tabs import project_tabs  # noqa: E402


def main():
    import shutil
    shutil.rmtree(OUT, ignore_errors=True)  # old versions removed; previous versions stay in git history
    qual = read_qual()
    by_row = {r["row"]: r for r in qual}
    quan = read_quan(by_row)
    global DATABOOK, COVERAGE, QUAL_ALL
    QUAL_ALL = qual
    DATABOOK = quant.read_databook(qual)
    COVERAGE = quant.coverage(qual, quan, DATABOOK)
    sr25 = read_sr25_faculty()
    fac_rows = [r for r in qual if asks_faculties(r)]
    central_sections = {code: [(label, [r for r in qual if m(r)]) for label, m in secs] for code, _, _, secs in CENTRAL}
    print("Qual indicators:", len(qual), "| faculty rows:", len(fac_rows), "| quant rows:", len(quan))
    for code, name in PILOT_FACULTIES.items():
        s25 = sr25_for(code)
        print(" ", request_workbook(code, name, [(f"{code} – {name}", fac_rows)], prev_lookup(s25),
                                    sr25_sections=s25, qual_by_row=by_row), sum(len(x[1]) for x in s25), "SR25 rows")
    for code, name, _, _ in CENTRAL:
        secs = [(l, rows) for l, rows in central_sections[code] if rows]
        s25 = sr25_for(code)
        print(" ", request_workbook(code, name, secs, prev_lookup(s25) if s25 else None, ref_map=CENTRAL_SR25_MAP[code],
                                    sr25_sections=s25, qual_by_row=by_row, folder="central_unit_requests"),
              [len(r) for _, r in secs], sum(len(x[1]) for x in s25), "SR25 rows")
    purged = quant.quant_only_rows(qual)
    print("  quant-only rows removed from requirements matrix:", [r["ref"] for r in qual if r["row"] in purged])
    print(" ", master([r for r in qual if r["row"] not in purged], quan, fac_rows, central_sections, sr25))
    pwb = Workbook()
    lists_sheet(pwb)
    project_tabs(pwb)
    del pwb[pwb.sheetnames[0]]
    pwb.move_sheet("Lists", offset=len(pwb.sheetnames))
    pwb.active = 0
    pwb.save(OUT / f"SR26 project management - {VERSION}.xlsx")
    print("  SR26 project management -", VERSION)


if __name__ == "__main__":
    main()
