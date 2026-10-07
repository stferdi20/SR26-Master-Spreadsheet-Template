"""Build the SR26 data collection master spreadsheet and per-unit reporting request workbooks.

Inputs (not committed, place in ./source/):
  S2030_targets.xlsx   - Sustainability 2030 updated targets (Qual + Quan worksheets)
  SR25_master.xlsx     - 2025 end-year reporting master spreadsheet (for End-2025 reference responses)

Run:  python3 build.py
"""
import re
import warnings
from datetime import date
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
VERSION = "v0.17"  # bump on every revision; appears in file names and Read me
VERSION_HISTORY = [
    ("v0.1", "First draft: requirements matrix, pilot faculty and central-unit requests, tracker, evidence register."),
    ("v0.2", "SR25-style faculty wording and 'How SST will use'; scoring/T1/T2 formulas; Legal & Risk; project tabs; linked SR25 responses tab."),
    ("v0.3", "Quantitative indicators separated: removed from requirements matrix; '2. Quant coverage' and '2b. Databook register' against the Databook draft; '4. Quantitative data' tab in central-unit requests."),
    ("v0.4", "RASCI rebuilt to the SR25 standard ('3b. RASCI matrix'): indicator hierarchy, consolidated portfolios, SST contacts, A/R/S/C/I codes, summary counts."),
    ("v0.5", "'3c. RASCI triangulation': each SR26 indicator compared with the SR25 RASCI (equivalent targets, translated areas), gaps and suggested actions."),
    ("v0.6", "Chancellery units added following SR25 (ACM, CGCE, Chancellery Education, SASS, Chancellery Indigenous) with linked SR25 responses; cross-tab consistency check (check.py); Impact column added to Issues log (as SR25)."),
    ("v0.7", "Team review: Chancellery units as separate tabs (AC&M, MRE, GCE, Education, SaSS, Indigenous); faculty requests cut down (ABP walkthrough) – EE1(a)+(c) merged, 'any examples' wording, alumni to SaSS, CL2(a)/(c) to AC&M, TR3(a)/(c)/(d) to MRE; simpler evidence register; highlighted stories compiled per unit with SST priority-area columns; project tabs moved to a separate project workbook."),
    ("v0.8", "TR1(b)–(d) to MRE only; EN1(d) removed from faculties (9 faculty questions); note on EE1(c) in T1/T2 that faculty answers sit under EE1(a)."),
    ("v0.9", "Team timeline applied: early engagement w/c 26 Oct; requests W1 Nov; due 15 Dec 2026 (responses and stories); first review W3 Dec; follow-up W3–W4 Dec; consolidation W4 Dec; CDSS as contact; optional 1:1 meetings; tracker statuses and project timeline updated."),
    ("v0.10", "RASCI: the four pilot faculties merged into one 'Faculties' column (identical requests; SR25 contacts kept in the contact row)."),
    ("v0.11", "Removed '3c. RASCI triangulation' (agreed and applied) and '0.5 Target Check' (T2 now reads 'Requested?' from A1)."),
    ("v0.12", "Teams linking (pilot: ABP): request files get stable names and protected layouts; master Faculties ABP rows and ABP story slots linked to the ABP file in the same folder; tracker counts answered and owner-confirmed rows automatically."),
    ("v0.13", "Formatting fixes in request files: no frozen panes (titles were clipped), 'SR25 response' header no longer merged across hidden columns, wider column H."),
    ("v0.14", "No target status rating in 2026 (Rose): rating column and definitions removed from all requests; A1 renamed 'A1. Response summary' (units asked, responses received, awaiting, owner confirmed); T1/T2 and tracker updated; columns shift left by one."),
    ("v0.15", "Fix workbook links: external reference written in Excel's syntax ('[1]Sheet'!A1, not [1]'Sheet'!A1); link lists the request file's actual sheet names."),
    ("v0.16", "Master unit tabs: SR25 columns I–K group now shows its [+] button (outline level declared as in SR25; collapsed flag on column H)."),
    ("v0.17", "Excel-online formatting pass on every tab: row 1 header and version removed; header rows frozen only (no frozen columns clipping titles); row heights sized to wrapped text; taller title row."),
]
# Team timeline (Oct 2026): early engagement w/c 26 Oct; requests issued W1 Nov; 6-week collection to 15 Dec;
# first review W3 Dec; targeted follow-up W3–W4 Dec; consolidated master W4 Dec.
DUE_DT = date(2026, 12, 15)
DUE_DATE = "Tuesday 15 December 2026"
STORIES_DUE = DUE_DATE
CONTACT = "CDSS – questions and clarifications are recorded centrally (contact email TBC); optional 1:1 meetings available on request"
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
TRACK_ROWS = {}  # tracker key (faculty code / section label) -> tracker row
TRACK_STATUS = ["Not sent", "Sent", "1:1 meeting", "Received", "First review", "Follow-up", "Complete", "Not required"]

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
    """Sheet title in B2 (row 1 left empty – no running header or version)."""
    ws["B2"] = sub; ws["B2"].font = Font(name="Aptos", size=22, bold=True, color=NAVY)
    ws.row_dimensions[2].height = 34
    ws.sheet_view.showGridLines = False


def finalize(wb):
    """Last pass before saving, for Excel online (which does not auto-fit): freeze header rows only (a frozen column
    clips titles in column B), and size rows to their wrapped text."""
    import math
    from openpyxl.utils import column_index_from_string
    for ws in wb.worksheets:
        fp = ws.freeze_panes
        if fp:
            row = int("".join(ch for ch in fp if ch.isdigit()))
            ws.freeze_panes = f"A{row}" if row <= 12 else None
        width = {}
        for c in range(1, ws.max_column + 1):
            d = ws.column_dimensions[get_column_letter(c)]
            width[c] = 0 if d.hidden else (d.width or 8.43)
        span, skip = {}, set()
        for m in ws.merged_cells.ranges:
            cells = {(r, c) for r in range(m.min_row, m.max_row + 1) for c in range(m.min_col, m.max_col + 1)}
            if m.max_row > m.min_row:
                skip |= cells           # multi-row blocks keep their own heights
            else:
                span[(m.min_row, m.min_col)] = sum(width[c] for c in range(m.min_col, m.max_col + 1))
                skip |= cells - {(m.min_row, m.min_col)}
        for row in ws.iter_rows():
            need = 0
            for cell in row:
                v = cell.value
                if not isinstance(v, str) or v.startswith("=") or (cell.row, cell.column) in skip:
                    continue
                size = cell.font.sz or 11
                w = span.get((cell.row, cell.column), width.get(cell.column, 8.43))
                if not w:
                    continue
                if cell.alignment.wrap_text:
                    per_line = max(1.0, w * 1.05 * 11 / size * (0.9 if cell.font.b else 1))
                    lines = sum(max(1, math.ceil(len(p) / per_line)) for p in v.split("\n"))
                else:
                    lines = 1
                need = max(need, lines * size * 1.5 + 6)  # margin: Excel online renders slightly wider
            if need:
                cur = ws.row_dimensions[row[0].row].height or 15
                if need > cur:
                    ws.row_dimensions[row[0].row].height = min(409, round(need))


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
    ws["B1"] = "(not used – no target status rating in 2026)"
    ws["B2"] = "Please select from dropdown"
    for i, s in enumerate([]):
        ws.cell(3 + i, 2, s)
    ws["C1"] = "Data owner confirms"
    for i, s in enumerate(["Yes", "No", "Partially – see comments"]):
        ws.cell(2 + i, 3, s)
    ws["D1"] = "Request status"
    for i, s in enumerate(TRACK_STATUS):
        ws.cell(2 + i, 4, s)
    ws["E1"] = "Y/N"; ws["E2"] = "Yes"; ws["E3"] = "No"; ws["E4"] = "TBC"
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
             13: "MRE", 14: "MRE", 15: "MRE",  # TR1(b), (c), (d) academic–professional collaboration mechanisms
             20: "MRE", 22: "MRE", 23: "MRE"}  # TR3(a), (c), (d) research systems, culture, assessment
FACULTY_REMOVED = {55}  # EN1(d) governance – not asked of faculties (central units keep it)
FACULTY_MERGED = {10: 8}  # EE1(c) is covered by the faculty EE1(a) question


def asks_faculties(r):
    """True if the pilot faculties receive a request for this Qual row."""
    return ("Faculties" in r["stake"] and r["row"] not in EXCLUSIVE and r["row"] not in FACULTY_MERGED
            and r["row"] not in FACULTY_REMOVED)


def merged_note(ref_by_row):
    """{ref: note} for indicators whose faculty answers are filed under another indicator."""
    return {ref_by_row[a]: f"Faculties answer this indicator in their {ref_by_row[b]} response (merged question) – review both."
            for a, b in FACULTY_MERGED.items() if a in ref_by_row and b in ref_by_row}


def _exclusive(m, code, is_first):
    """Exclusive rows go to their owner's first section only; other rows use the section's own matcher."""
    return lambda r: (EXCLUSIVE[r["row"]] == code and is_first) if r["row"] in EXCLUSIVE else m(r)


for _code, _name, _sr25, _secs in CENTRAL:
    _secs[:] = [(lab, _exclusive(m, _code, k == 0)) for k, (lab, m) in enumerate(_secs)]


# Request workbooks linked into the master (same Teams folder; relative workbook links). Add codes as files are uploaded.
LINKED_UNITS = ["ABP"]
REQ_SHEETS = {}  # code -> sheet names of the request workbook (for the external link part)
REQ_ROWS = {}  # code -> {"refs": {ref: row on tab 1}, "stories": [rows on tab 2]} – filled by request_workbook


def request_name(code):
    """Stable request file name (no version) so workbook links survive revisions."""
    return f"{file_code(code)} - SR26 sustainability reporting request.xlsx"


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
        "OPTION 1:  Target-level reporting (Column D)",
        "OPTION 2:  Indicator level reporting (Column E/F)", "Supporting information/comments",
        "Data source and limitations", "Data owner confirms information is accurate", "Confirmed by (name, role, date)"]
# No target status rating is requested in 2026 (team decision, Rose) – responses are commentary and evidence only
SUBS = {12: "A general update on activities in relation to the overarching target (Column D)",
        13: "Data and commentary in direct response to the reporting request (Column F)",
        14: "Any additional information (e.g. links, images, documents)",
        15: "Where did this information come from (system, report, survey, person)? Note any gaps or limitations.",
        16: "Select 'Yes' to confirm the information provided is accurate and can be used for reporting",
        17: "Name, role and date of the person confirming"}
WIDTHS = {"A": 4, "B": 10, "C": 16, "D": 38, "E": 38, "F": 50, "G": 34, "H": 18, "I": 16, "J": 50, "K": 24,
          "L": 48, "M": 48, "N": 30, "O": 30, "P": 20, "Q": 24}


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
        ((8, 8), "SR25 response", HDR, Font(name="Aptos Narrow", size=14, bold=True, color="FFFFFF")),
        ((12, 17), "End-2026 reporting (PLEASE COMPLETE THIS SECTION)", GOLD, Font(name="Aptos Narrow", size=14, bold=True, color=DARK))]:
        if rng[1] > rng[0]:  # no merge across hidden columns (I–K) – the text would be squashed into H
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
        for col in range(12, 18):
            body(ws.cell(rr, col))
        ws.row_dimensions[rr].height = 150
    last = r0 + len(rows) - 1
    if reg is not None:
        reg.append((key or unit_label, ws.title, r0, last + 1))  # +1 blank row: keeps single-row blocks a true range
    if rows:
        dv_list(ws, lists["confirm"], f"P{r0}:P{last}")
    return last + 3


def protect(ws, editable):
    """Lock the layout (no inserting/deleting rows, which would break workbook links); only answer cells are editable."""
    from openpyxl.styles import Protection
    for r, c in editable:
        ws.cell(r, c).protection = Protection(locked=False)
    ws.protection.sheet = True
    ws.protection.formatRows = False      # allow resizing rows
    ws.protection.formatColumns = False   # allow resizing columns
    ws.protection.formatCells = True


def setup_block_sheet(ws, outline=True):
    widths(ws, WIDTHS)
    for col in "IJK":  # H (related SR25 ref + link) stays visible
        if outline:  # grouping can't be toggled on a protected sheet, so request files just hide the summary
            ws.column_dimensions[col].outlineLevel = 1
        ws.column_dimensions[col].hidden = True
    ws.sheet_properties.outlinePr.summaryRight = False
    if outline:  # Excel only draws the [+] when the summary column next to the group (H) is flagged collapsed
        ws.column_dimensions["H"].collapsed = True
        ws.sheet_view.showOutlineSymbols = True
        ws.sheet_format.outlineLevelCol = 1  # as SR25: reserves the outline bar where the [+] is drawn


def instructions(ws, unit_name):
    ws.merge_cells("B4:C8"); ws["B4"] = "How to use this spreadsheet:"
    ws["B4"].font = Font(name="Aptos", bold=True, color="FFFFFF"); ws["B4"].fill = HDR; ws["B4"].alignment = WRAP
    ws.merge_cells("D4:G8")
    ws["D4"] = ("• For each row below, describe progress towards the target (Column L) and/or respond to the 2026 reporting request (Column M)\n"
                "• Provide supporting information if required (Column N), including links to documents or images\n"
                "• Note the source of your information and any limitations (Column O), then confirm the information is accurate (Columns P–Q)\n"
                "• Where a related end-2025 target existed, click the ref in Column H to see your full SR25 response (Tab 3)\n"
                "• Use Tab 2 of this spreadsheet to highlight key sustainability stories (optional)\n"
                f"• 2026 is a transition year: Sustainability 2030 launched ~20 October 2026, so report on activity across all of 2026 "
                f"and any early actions since launch. No target status rating is requested this year.\n• Due: {DUE_DATE}. Questions: {CONTACT}")
    ws["D4"].alignment = WRAP; ws["D4"].font = Font(name="Aptos", size=11)
    ws["D4"].border = MBOX
    ws.row_dimensions[8].height = 60


STORY_COLS = ["Title", "Description", "Related Sustainability 2030 targets (optional)", "Images and other media", "Links and supporting documentation"]


def stories_sheet(ws, unit=None):
    title(ws, "SR26 sustainability reporting", "Highlighted stories (optional)")
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
    title(ws, "SR26 sustainability reporting", "Highlighted stories – compiled")
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


def link_cells(wb, reg):
    """Point the master's answer cells (and story slots) at each linked request workbook: [n] = n-th external link."""
    tmpl, stor = "1. Reporting template", "2. OPTIONAL highlighted stories"
    for n, code in enumerate(LINKED_UNITS, start=1):
        rows = REQ_ROWS[code]
        for key, sheet, r0, r1 in reg:
            if key != code:
                continue
            ws = wb[sheet]
            for r in range(r0, r1 + 1):
                src = rows["refs"].get(ws.cell(r, 2).value)
                if not src:
                    continue
                for c in range(12, 18):
                    a = f"'[{n}]{tmpl}'!{get_column_letter(c)}{src}"  # Excel syntax: '[n]Sheet name'!A1
                    ws.cell(r, c).value = f'=IF({a}="","",{a})'
                    ws.cell(r, c).fill = PatternFill("solid", fgColor="E2EFDA")
        st = wb["Highlighted stories"]
        first = next(r for r in range(6, st.max_row + 1) if str(st.cell(r, 2).value or "").startswith(code + " "))
        for k, src in enumerate(rows["stories"]):
            for i in range(len(STORY_COLS)):
                a = f"'[{n}]{stor}'!{get_column_letter(2 + i)}{src}"
                st.cell(first + k, 4 + i).value = f'=IF({a}="","",{a})'
                st.cell(first + k, 4 + i).fill = PatternFill("solid", fgColor="E2EFDA")


def patch_outline(path):
    """openpyxl drops sheetFormatPr/@outlineLevelCol on save; without it Excel draws no [+] for grouped columns."""
    import zipfile
    zf = zipfile.ZipFile(path)
    files = {n: zf.read(n) for n in zf.namelist()}
    zf.close()
    for n, b in files.items():
        if n.startswith("xl/worksheets/sheet"):
            x = b.decode()
            if 'outlineLevel="1"' in x and "outlineLevelCol" not in x:
                files[n] = x.replace("<sheetFormatPr ", '<sheetFormatPr outlineLevelCol="1" ', 1).encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n in ["[Content_Types].xml"] + [k for k in files if k != "[Content_Types].xml"]:
            z.writestr(n, files[n])


def add_external_links(path, codes):
    """Inject Excel external-link parts (relative to the master's folder) for each linked request workbook."""
    import zipfile
    from urllib.parse import quote
    src = zipfile.ZipFile(path)
    files = {n: src.read(n) for n in src.namelist()}
    src.close()
    refs, rels, ctypes = [], [], []
    for n, code in enumerate(codes, start=1):
        sheets = REQ_SHEETS[code]  # the linked file's actual sheet names, in order
        names = "".join(f'<sheetName val="{s}"/>' for s in sheets)
        data = "".join(f'<sheetData sheetId="{i}"/>' for i in range(len(sheets)))
        files[f"xl/externalLinks/externalLink{n}.xml"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<externalBook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1">'
            f'<sheetNames>{names}</sheetNames><sheetDataSet>{data}</sheetDataSet></externalBook></externalLink>').encode()
        files[f"xl/externalLinks/_rels/externalLink{n}.xml.rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" '
            f'Target="{quote(request_name(code))}" TargetMode="External"/></Relationships>').encode()
        refs.append(f'<externalReference r:id="rIdExt{n}"/>')
        rels.append(f'<Relationship Id="rIdExt{n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" '
                    f'Target="externalLinks/externalLink{n}.xml"/>')
        ctypes.append(f'<Override PartName="/xl/externalLinks/externalLink{n}.xml" '
                      'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>')
    wbx = files["xl/workbook.xml"].decode()
    if 'xmlns:r=' not in wbx.split(">", 2)[1]:
        wbx = wbx.replace("<workbook ", '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ', 1)
    files["xl/workbook.xml"] = wbx.replace("</sheets>", "</sheets><externalReferences>" + "".join(refs) + "</externalReferences>", 1).encode()
    files["xl/_rels/workbook.xml.rels"] = files["xl/_rels/workbook.xml.rels"].decode().replace("</Relationships>", "".join(rels) + "</Relationships>").encode()
    files["[Content_Types].xml"] = files["[Content_Types].xml"].decode().replace("</Types>", "".join(ctypes) + "</Types>").encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n in ["[Content_Types].xml"] + [k for k in files if k != "[Content_Types].xml"]:
            z.writestr(n, files[n])


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
    title(ws, "SR26 sustainability reporting", "SR26 sustainability reporting request")
    setup_block_sheet(ws, outline=False); instructions(ws, name)
    top, rreg = 11, []
    for label, rows in sections:
        top = block(ws, top, label, rows, sr25, lists, ref_map=ref_map, anchors=anchors, quant_refs=quant_refs, reg=rreg, key=label)
    refs = {}
    for _, _, r0, r1 in rreg:
        for r in range(r0, r1 + 1):
            if ws.cell(r, 2).value:
                refs[ws.cell(r, 2).value] = r
    REQ_ROWS[code] = {"refs": refs, "stories": [10, 11, 12]}
    REQ_SHEETS[code] = None  # set after tab order is final
    protect(ws, [(r, c) for r in refs.values() for c in range(12, 18)])
    # no frozen panes: a frozen column clips the title and section names in column B
    st = wb.create_sheet("2. OPTIONAL highlighted stories", 1)
    stories_sheet(st, unit=code)
    protect(st, [(r, c) for r in (10, 11, 12) for c in range(2, 2 + len(STORY_COLS))])
    if "4. Quantitative data" in wb.sheetnames:
        q4 = wb["4. Quantitative data"]
        protect(q4, [(r, c) for r in range(8, q4.max_row + 1) for c in range(11, 16) if q4.cell(r, 3).value
                     and q4.cell(r, 2).fill.fgColor.rgb not in ("00F2F2F2", "FFF2F2F2")])
    order = ["1. Reporting template", "2. OPTIONAL highlighted stories", SR25_SHEET, "4. Quantitative data", "Lists"]
    wb._sheets.sort(key=lambda w: order.index(w.title))
    REQ_SHEETS[code] = list(wb.sheetnames)
    wb.active = 0
    safe = file_code(code)
    path = OUT / folder / request_name(code)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.properties.version = VERSION
    finalize(wb)
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
        ("Reporting approach", "Transition year: cover all of CY2026, distinguishing former Sustainability Plan 2030 activity from foundations and early actions after the Sustainability 2030 launch (~20 Oct 2026). Databook is the authoritative quantitative source. No target status rating (Met/Partially met/Not met) is requested in 2026 (team decision) – requests collect commentary and evidence only."),
        ("Key dates (TBC)", f"Early engagement (existing meetings): w/c 26 Oct 2026 | Requests issued: W1 Nov (from 2 Nov) | Collection period incl. optional 1:1 meetings: 2 Nov – 15 Dec (6 weeks) | Responses and highlighted stories due: {DUE_DATE} | First review: W3 Dec | Targeted follow-up: W3–W4 Dec | Consolidated master: W4 Dec | Sustainability Reporting Review Group: ~15 Feb 2027 | VCAG: 16 Feb & ~2 Mar 2027"),
        ("Tabs", "1. Requirements matrix – every S2030 qualitative indicator, owner and request\n"
                 "2. Quant coverage – every quantitative indicator checked against the Databook draft, owner and action\n"
                 "2b. Databook register – every Databook data point with columns to confirm definition, source, owner\n"
                 "3. Stakeholder map – SR25 → SR26 unit mapping\n"
                 "3b. RASCI matrix – A/R/S/C/I roles per indicator and area (SR25 RASCI format), with counts\n"
                 "4. Request tracker – sent/chased/received/confirmed status and response rate\n"
                 "Faculties / CI&S / CFOG / ESG / CIOG / MRE / L&R – consolidated responses (copy in from returned request workbooks; same layout)\n"
                 "Highlighted stories – consolidated optional stories and case-study shortlist\n"
                 "A1. Response summary – auto-pulls every unit's responses per indicator; counts units asked, responses received and owner-confirmed\n"
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
    title(ws, "SR26 sustainability reporting", "1. Requirements matrix (qualitative)")
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
    title(ws, "SR26 sustainability reporting", "3. Stakeholder map (SR25 → SR26)")
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

    # 4. Request tracker
    ws = wb.create_sheet("4. Request tracker")
    title(ws, "SR26 sustainability reporting", "4. Request tracker")
    cols = ["#", "Group", "Unit / section", "Key contact", "No. of indicators requested", "Request workbook", "Date sent",
            "Due date", "Reminder date", "Optional 1:1 meeting (date)", "Rows answered (auto)", "Owner confirmed rows (auto)",
            "Status", "Days overdue", "Notes / follow-up"]
    hdr_row(ws, 9, cols)
    track = [(f"F{i+1}", "Faculty", f"{code} – {name}", len(fac_rows), request_name(code), code)
             for i, (code, name) in enumerate(PILOT_FACULTIES.items())]
    n = 0
    for code, name, _, secs in CENTRAL:
        for label, rows in central_sections[code]:
            n += 1
            track.append((f"C{n}", "Central", label, len(rows), request_name(code), label.split(" (")[0]))
    for k, t in enumerate(track):
        rr = 10 + k
        vals = [t[0], t[1], t[2], "TBC", t[3], t[4], None, DUE_DT, f"=IF(I{rr}=\"\",\"\",I{rr}-7)", "", None, "", "Not sent",
                f"=IF(OR(I{rr}=\"\",L{rr}>=F{rr}),\"\",MAX(0,TODAY()-I{rr}))", ""]
        for i, v in enumerate(vals):
            body(ws.cell(rr, 2 + i, v))
        from datetime import date
        for col in (8, 9, 10, 11):
            ws.cell(rr, col).number_format = "dd/mm/yyyy"
    last = 9 + len(track)
    dv_list(ws, f"=Lists!$D$2:$D${1 + len(TRACK_STATUS)}", f"N10:N{last}")
    TRACK_ROWS.update({t[5]: 10 + k for k, t in enumerate(track)})
    tracker_ws = ws
    # summary
    ws["B4"] = "Summary"; ws["B4"].font = Font(name="Aptos", bold=True, size=14, color=NAVY)
    for i, s in enumerate(TRACK_STATUS):
        ws.cell(5 + i // 4 * 2, 3 + (i % 4) * 2, s).font = Font(name="Aptos", bold=True)
        ws.cell(6 + i // 4 * 2, 3 + (i % 4) * 2, f'=COUNTIF($N$10:$N${last},"{s}")')
    ws["L5"] = "Response rate"; ws["L5"].font = Font(name="Aptos", bold=True)
    ws["L6"] = f'=IFERROR((COUNTIF($N$10:$N${last},"Received")+COUNTIF($N$10:$N${last},"First review")+COUNTIF($N$10:$N${last},"Follow-up")+COUNTIF($N$10:$N${last},"Complete"))/(COUNTA($N$10:$N${last})-COUNTIF($N$10:$N${last},"Not required")),0)'
    ws["L6"].number_format = "0%"
    ws["N5"] = "Owner confirmed"; ws["N5"].font = Font(name="Aptos", bold=True)
    ws["N5"].value = "Rows owner-confirmed"
    ws["N6"] = f'=SUM($M$10:$M${last})&" / "&SUM($F$10:$F${last})'
    widths(ws, dict(zip("ABCDEFGHIJKLMNOP", [4, 6, 10, 46, 16, 12, 44, 12, 12, 12, 16, 14, 12, 16, 10, 40])))
    ws.freeze_panes = "E10"

    # Unit tabs (consolidated responses)
    reg = []

    def unit_tab(tab, subtitle, sections, sr25_for):
        ws = wb.create_sheet(tab)
        title(ws, "SR26 sustainability reporting", subtitle)
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
    for key, sheet, r0, r1 in reg:  # tracker progress, read from the unit tabs (which are linked to the request files)
        tr = TRACK_ROWS.get(key)
        if tr:
            q = f"'{sheet}'!"
            # a row counts as answered when Option 1 or Option 2 has text
            tracker_ws.cell(tr, 12, f'=SUMPRODUCT(--((LEN({q}$L${r0}:$L${r1})+LEN({q}$M${r0}:$M${r1}))>0))')
            tracker_ws.cell(tr, 13, f'=COUNTIF({q}$P${r0}:$P${r1},"Yes")')
    wording_log(wb, fac_rows)

    # Highlighted stories – compiled from every request's Tab 2, with SST assessment columns
    stories_master(wb)
    link_cells(wb, reg)  # after the stories tab exists

    # Evidence register – kept simple: one line per material claim or figure used in the report
    ws = wb.create_sheet("Evidence register")
    title(ws, "SR26 sustainability reporting", "Evidence register")
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
    title(ws, "SR26 sustainability reporting", "Issues log")
    hdr_row(ws, 4, ["#", "Date raised", "Issue / question", "Impact", "Ref / unit", "Owner", "Proposed action", "Status"])
    issues = [
        ("Climate resilience maturity (CL3(a)) and Estate & infrastructure indicators – confirm owner and request with Gerard.", "CL3(a); EI1(a)–(d)", "Stefanus", "Meet Gerard", "Open"),
        ("Confirm whether CGOP, CDEP and CDSS structures have changed since the 2025 master spreadsheet stakeholder mapping.", "Stakeholder map", "Stefanus", "Confirm with team", "Open"),
        ("Business Services split into CIOG (AI) and ESG (nature & biodiversity, waste & circular economy, part of quant. climate leadership) – confirm contacts.", "CIOG; ESG", "TBC", "Confirm contacts", "Open"),
        ("Procurement sits with CFOG (RP1(a) procurement case studies routed to CFOG – Procurement although the Qual sheet names no stakeholder); estate planning centralised in CI&S – requests issued centrally, not to faculties.", "CFOG; CI&S", "—", "Noted", "Closed"),
        ("Chancellery units reported separately, as SR25: AC&M, MRE, GCE, Education, SaSS and Indigenous each have their own request and master tab. Alumni (EE1(d)) asked of SaSS only.", "Stakeholder map", "Stefanus", "Confirm contacts", "Closed"),
        ("Faculty requests cut down after ABP walkthrough: EE1(a)+(c) merged; CL2(a)/(c) to AC&M only; TR1(b)–(d) and TR3(a)/(c)/(d) to MRE only; EN1(d) not asked of faculties; EE1(d) to SaSS only. Wattle Fellowship to be asked directly (not tracked in master).", "Faculties", "Stefanus", "Confirm with team; apply to other faculties", "Open"),
        ("Legal & Risk asked for climate resilience maturity CL3(a), following SR25 (8a(i)/(ii): University Risk 16 Climate Change; flood emergency response plans). CL3(a) is also with CI&S (Gerard) – agree who leads.", "CL3(a); L&R; CI&S", "Stefanus", "Confirm with Gerard", "Open"),
        ("TR2(b) 'Documented progress of strategic initiatives, incl. Impact Accelerators' reuses the TR2(d) case-study wording and has no stakeholder in the Qual sheet – left as-is.", "TR2(b)", "TBC", "Review wording/owner", "Open"),
        ("Quantitative-only rows (no stakeholder in the Qual sheet) removed from the requirements matrix and moved to '2. Quant coverage'; Databook gaps requested in Tab 4 of the owning central unit's request.", "TR2(b); CL1(a)-(b); CL2(b); NB1(a); CE1-2; RP2; RI1(a)", "Stefanus", "Confirm owners and Databook coverage with Chris", "Open"),
        ("No target status rating requested in 2026 (Rose): rating column removed from all requests; A1 now summarises responses instead of scoring.", "All", "Director, Sustainability", "Revisit for 2027 with the new traffic-light framework", "Closed"),
        ("Timeline updated to the team plan: requests issued W1 Nov, responses and stories due 15 Dec 2026, first review W3 Dec, follow-up W3–W4 Dec, consolidation W4 Dec. Databook figures given by 15 Dec are provisional; final figures confirmed Jan–Feb.", "All", "Stefanus", "Confirm exact issue date and CDSS contact", "Open"),
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
    wb.properties.version = VERSION
    finalize(wb)
    wb.save(p)
    if LINKED_UNITS:
        add_external_links(p, LINKED_UNITS)
    patch_outline(p)
    return p.name


# ---------------------------------------------------------------- scoring engine (adapted from SR25 A1 / 0.5 / T2 tabs)
FIELDS = [("OPTION 1:  Target-level reporting", "L"), ("OPTION 2:  Indicator level reporting", "M"),
          ("Supporting information/comments", "N"), ("Data source and limitations", "O"), ("Data owner confirms", "P")]
AUTOFILL = "Cell will autofill based on response"


def scoring_tabs(wb, qual, reg):
    units = []  # (key, [(sheet, r0, last)]) – a unit may appear once per sheet
    for key, sheet, r0, last in reg:
        units.append((key, sheet, r0, last))
    n_u = len(units)
    uc = lambda i: get_column_letter(8 + i)            # unit columns start at H
    lastu = uc(n_u - 1)

    # ---- A1. Data scoring
    ws = wb.create_sheet("A1. Response summary")
    title(ws, "SR26 sustainability reporting", "A1. Response summary")
    ws["C4"] = "Description"; ws["C4"].font = Font(name="Aptos", bold=True)
    ws["D4"] = ("Pulls every unit's end-2026 response for each indicator from the Faculties and central-unit tabs (do not type in "
                "the grey cells). 'N/A' = not requested from that unit. No target status rating is collected in 2026, so the summary "
                "columns count responses: units asked, responses received (Option 1 or 2 answered), awaiting response and owner-confirmed.")
    ws["D4"].alignment = WRAP; ws.merge_cells("D4:N6")
    groups = ["Faculty" if k in PILOT_FACULTIES else "Central" for k, *_ in units]
    head = ["Ref.", "Priority area", "Indicator", "Reporting requirement", "Target code", "Target progress"]
    for i, h in enumerate(head):
        ws.cell(10, 2 + i, h)
    for i, (k, *_r) in enumerate(units):
        ws.cell(9, 8 + i, groups[i]); ws.cell(10, 8 + i, k)
    sc0 = 8 + n_u + 1
    score_cols = ["Units asked", "Responses received", "Awaiting response", "Owner confirmed"]
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
                o1, o2, ok = (f"$H{r + k}:${lastu}{r + k}" for k in (0, 1, 4))
                got = lambda rg: f'(({rg}<>"N/A")*({rg}<>"{AUTOFILL}"))'
                fs = {"Units asked": f'=SUMPRODUCT(--({o1}<>"N/A"))',
                      "Responses received": f"=SUMPRODUCT(--(({got(o1)}+{got(o2)})>0))",
                      "Awaiting response": f'={C["Units asked"]}{r}-{C["Responses received"]}{r}',
                      "Owner confirmed": f'=COUNTIF({ok},"Yes")'}
                for h, f in fs.items():
                    body(ws.cell(r, sc0 + score_cols.index(h), f), PatternFill("solid", fgColor="E2EFDA"))
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

    # ---- T1. Manual target review: every indicator, only the units asked, side by side (base review sheet)
    MERGED = merged_note({q["row"]: q["ref"] for q in qual})
    t1 = wb.create_sheet("T1. Manual target review")
    title(t1, "SR26 sustainability reporting", "T1. Manual target review")
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
            ["Responses received", "Owner confirmed", "Reviewer notes"])
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
        for col, h in zip(sc[:2], ["Responses received", "Owner confirmed"]):
            c = t1[f"{col}{r}"]; c.value = f"='A1. Response summary'!{C[h]}{a1}"; body(c, PatternFill("solid", fgColor="E2EFDA"), bold=True)
        body(t1[f"{sc[2]}{r}"])
        t1[f"{sc[2]}{r}"].value = MERGED.get(q["ref"])
        for fi, (fname, _) in enumerate(FIELDS):
            rr = r + 1 + fi
            body(t1.cell(rr, 2, q["ref"])); t1.cell(rr, 2).font = Font(name="Aptos Narrow", size=9, color="808080")
            body(t1.cell(rr, 4, fname), bold=True)
            for j, ui in enumerate(idx):
                body(t1.cell(rr, 5 + j, f"='A1. Response summary'!{uc(ui)}{a1 + fi}"), GREY)
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
    title(t2, "SR26 sustainability reporting", "T2. Single target review")
    t2["B4"] = "Target selected"; t2["B4"].font = Font(name="Aptos", bold=True)
    t2["C4"] = qual[0]["ref"]; t2["C4"].fill = GOLD; t2["C4"].font = Font(name="Aptos", bold=True, size=14); t2["C4"].border = MBOX
    note_f = "".join(f'IF($C$4="{k}","{v}",' for k, v in MERGED.items())
    t2["D4"] = f'={note_f}"← select an indicator ref from the dropdown"' + ")" * len(MERGED)
    t2["D4"].font = Font(name="Aptos", bold=True, color="C00000")
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
    for i, h in enumerate(["Units asked", "Responses received", "Awaiting response", "Owner confirmed"]):
        a = t2.cell(srow, 2 + i * 2, h); a.font = Font(name="Aptos", bold=True); a.alignment = WRAP
        c = t2.cell(srow + 1, 2 + i * 2, f"=IFERROR(INDEX('A1. Response summary'!${C[h]}$11:${C[h]}${last_a1},"
                                         f"MATCH($C$4&\"|\"&\"{FIELDS[0][0]}\",'A1. Response summary'!$A$11:$A${last_a1},0)),\"\")")
        c.font = Font(name="Aptos", bold=True, size=13, color=NAVY)
    hr = 14
    hdr_row(t2, hr, ["Target progress"] + [k for k, *_ in units])
    body(t2.cell(hr + 1, 2, "Requested?"), GREY2, bold=True)
    for i in range(n_u):
        # Requested? read straight from A1: 'N/A' there means the unit was not asked
        c = t2.cell(hr + 1, 3 + i, f"=IFERROR(IF(INDEX('A1. Response summary'!{uc(i)}$11:{uc(i)}${last_a1},"
                                    f"MATCH($C$4&\"|\"&\"{FIELDS[0][0]}\",'A1. Response summary'!$A$11:$A${last_a1},0))=\"N/A\",\"No\",\"Yes\"),\"\")")
        body(c, GREY2); c.alignment = CWRAP
    for fi, (fname, _) in enumerate(FIELDS):
        rr = hr + 2 + fi
        body(t2.cell(rr, 2, fname), REFF, bold=True)
        for i in range(n_u):
            c = t2.cell(rr, 3 + i, f"=IFERROR(INDEX('A1. Response summary'!{uc(i)}$11:{uc(i)}${last_a1},"
                                     f"MATCH($C$4&\"|\"&$B{rr},'A1. Response summary'!$A$11:$A${last_a1},0)),\"\")")
            body(c)
        t2.row_dimensions[rr].height = 120 if fi in (1, 2) else 45
    widths(t2, {"B": 26})
    for i in range(n_u):
        t2.column_dimensions[get_column_letter(3 + i)].width = 30
    t2.column_dimensions["C"].width = 30
    t2.freeze_panes = "C15"


def wording_log(wb, fac_rows):
    ws = wb.create_sheet("Request wording log")
    title(ws, "SR26 sustainability reporting", "Faculty request wording log")
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
    finalize(pwb)
    pwb.save(OUT / f"SR26 project management - {VERSION}.xlsx")
    print("  SR26 project management -", VERSION)


if __name__ == "__main__":
    main()
