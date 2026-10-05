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

warnings.filterwarnings("ignore")
ROOT = Path(__file__).parent
SRC = ROOT / "source"
OUT = ROOT / "output"

# ---------------------------------------------------------------- settings
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


def block(ws, top, unit_label, rows, sr25, lists):
    """Write a reporting block at row `top`. Returns next free row."""
    c = ws.cell(top, 2, unit_label)
    c.font = Font(name="Aptos Narrow", size=22, bold=True); c.fill = GREY
    h1 = top + 2
    for rng, text, fill, fnt in [
        ((3, 7), "Sustainability 2030 target and indicator", HDR, Font(name="Aptos Narrow", size=14, bold=True, color="FFFFFF")),
        ((8, 11), "End-2025 reporting (for reference – click [+] above to expand)", HDR, Font(name="Aptos Narrow", size=14, bold=True, color="FFFFFF")),
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
    for k, r in enumerate(rows):
        rr = r0 + k
        prev = sr25.get(SR25_MAP.get(r["row"]), ("", "", "")) if sr25 is not None else ("", "", "")
        vals = [r["ref"], r["pa"], r["target"], r["indicator"], r["request"] or "(Reporting request TBC)", use_text(r),
                SR25_MAP.get(r["row"], "N/A – new indicator") if sr25 is not None else "", prev[0], prev[1], prev[2]]
        for i, v in enumerate(vals):
            cell = ws.cell(rr, 2 + i, v)
            body(cell, REFF if i < 4 else None, bold=(i == 0))
        for col in range(12, 19):
            body(ws.cell(rr, col))
        ws.row_dimensions[rr].height = 150
    last = r0 + len(rows) - 1
    if rows:
        dv_list(ws, lists["status"], f"L{r0}:L{last}")
        dv_list(ws, lists["confirm"], f"Q{r0}:Q{last}")
    return last + 3


def setup_block_sheet(ws):
    widths(ws, WIDTHS)
    for col in "HIJK":
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
                "• Your end-2025 responses (where a related target existed) can be viewed by clicking the [+] icon above Columns H–K\n"
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


# ---------------------------------------------------------------- per-unit request workbook
def request_workbook(code, name, sections, sr25):
    wb = Workbook(); ws = wb.active; ws.title = "1. Reporting template"
    lists = lists_sheet(wb)
    title(ws, "End-2026 target status assessment and reporting", "End-2026 target status confirmation and reporting")
    setup_block_sheet(ws); instructions(ws, name)
    top = 11
    for label, rows in sections:
        top = block(ws, top, label, rows, sr25, lists)
    ws.freeze_panes = "C14"
    st = wb.create_sheet("2. OPTIONAL highlighted stories", 1)
    stories_sheet(st, unit=code)
    wb.move_sheet("Lists", offset=0)
    safe = code.replace("&", "and")
    path = OUT / ("faculty_requests" if sr25 is not None else "central_unit_requests") / f"{safe} - end-2026 sustainability reporting request.xlsx"
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
        ("Status", "DRAFT v0.1 – pilot with 4 faculties (ABP, Arts, FBE, Science) and 5 central units (CI&S, CFOG, ESG, CIOG, MRE)."),
        ("Reporting approach", "Transition year: cover all of CY2026, distinguishing former Sustainability Plan 2030 activity from foundations and early actions after the Sustainability 2030 launch (~20 Oct 2026). Databook is the authoritative quantitative source. Target status (Met/Partially met/Not met) retained for now for internal management reporting – rating framework under review."),
        ("Key dates (TBC)", f"Requests issued: Nov–Dec 2026 | Highlighted stories due: {STORIES_DUE} | Reporting template due: {DUE_DATE} | Support meetings: January 2027 | Sustainability Reporting Review Group: ~15 Feb 2027 | VCAG: 16 Feb & ~2 Mar 2027"),
        ("Tabs", "1. Requirements matrix – every S2030 qualitative indicator, owner and request\n"
                 "2. Quant – Databook mapping – quantitative indicators to confirm with the Databook/central data owners\n"
                 "3. Stakeholder map & RASCI – SR25 → SR26 unit mapping and who is asked for what\n"
                 "4. Request tracker – sent/chased/received/confirmed status and response rate\n"
                 "Faculties / CI&S / CFOG / ESG / CIOG / MRE – consolidated responses (copy in from returned request workbooks; same layout)\n"
                 "Highlighted stories – consolidated optional stories and case-study shortlist\n"
                 "Evidence register – source, evidence, limitations and owner confirmation for every material claim\n"
                 "Issues log – open questions, TBCs and decisions"),
        ("Refs", "Refs are new SR26 codes: priority-area code + target number + indicator letter (e.g. EE1(a) = Exceptional education, target 1, indicator a). 'Related SR25 ref' links to the closest SP2030 (2025) target where one exists."),
        ("Regenerating", "Workbooks are generated by build.py in the repository. Edit the settings at the top (due dates, pilot faculties, unit mapping) and re-run to produce request workbooks for additional faculties/units."),
        ("Not changed", "Request wording is taken as-is from the S2030 'Qual' worksheet (Reporting Requirements column); known wording issues are recorded in the Issues log rather than edited."),
    ]
    for i, (k, v) in enumerate(readme):
        a = ws.cell(4 + i, 2, k); a.font = Font(name="Aptos", bold=True, color="FFFFFF"); a.fill = HDR; a.alignment = WRAP; a.border = BOX
        b = ws.cell(4 + i, 3, v); b.alignment = WRAP; b.border = BOX; b.font = Font(name="Aptos")
        ws.row_dimensions[4 + i].height = max(30, 15 * (v.count("\n") + 1 + len(v) // 110))
    widths(ws, {"A": 4, "B": 22, "C": 120})

    # 1. Requirements matrix
    ws = wb.create_sheet("1. Requirements matrix")
    title(ws, "End-2026 target status assessment and reporting", "1. Requirements matrix (qualitative)")
    cols = ["Ref", "Domain", "Priority area", "Target", "Indicator", "Qualitative target? (Y/M)", "Reporting requirement",
            "Key stakeholders (S2030 sheet)", "SR26 request issued to", "Specific projects we are aware of (area & person)",
            "Internal vs External", "Flag for continuous improvement", "Notes", "Related SR25 ref", "Channel", "Owner confirmed?"]
    hdr_row(ws, 5, cols)
    for k, r in enumerate(qual):
        rr = 6 + k
        to = [u[0] for u in CENTRAL for s in u[3] if s[1](r)]
        if "Faculties" in r["stake"]:
            to = ["Faculties (pilot)"] + to
        channel = "Tailored request" if to else ("Databook / central data owner" if "Quantitative" in r["notes"] or not r["stake"] else "TBC")
        vals = [r["ref"], r["domain"], r["pa"], r["target"], r["indicator"], r["qual"], r["request"], r["stake"],
                ", ".join(dict.fromkeys(to)) or "TBC – no owner identified", r["projects"], r["intext"], r["ci"], r["notes"],
                SR25_MAP.get(r["row"], ""), channel, ""]
        for i, v in enumerate(vals):
            body(ws.cell(rr, 2 + i, v), REFF if i == 0 else None, bold=i == 0)
        ws.row_dimensions[rr].height = 90
    ws.auto_filter.ref = f"B5:{get_column_letter(1 + len(cols))}{5 + len(qual)}"
    ws.freeze_panes = "D6"
    widths(ws, dict(zip("BCDEFGHIJKLMNOPQ", [9, 16, 18, 40, 40, 11, 50, 18, 22, 40, 11, 12, 30, 10, 18, 12])))

    # 2. Quant
    ws = wb.create_sheet("2. Quant – Databook mapping")
    title(ws, "End-2026 target status assessment and reporting", "2. Quantitative indicators – Databook mapping")
    ws["B3"] = ("Databook is the authoritative quantitative source – confirm coverage first and only request data separately where an indicator is not covered. "
                "Per the SR26 approach: map Sep–Oct; confirm provisional data in January; finalise by February.")
    ws["B3"].font = Font(name="Aptos", italic=True)
    cols = ["Ref", "Priority area", "Indicator", "Quant? (Y/M)", "Suggested metric", "Key person to follow up", "Internal vs External",
            "Analysis notes", "Covered by Databook?", "Definition", "Reporting boundary", "Period", "Source system", "End-2026 figure",
            "Quality note / limitations", "Data owner", "Owner confirmed?", "Confirmation date"]
    hdr_row(ws, 5, cols)
    for k, q in enumerate(quan):
        rr = 6 + k
        vals = [q["ref"], q["pa"], q["indicator"], q["flag"], q["metric"], q["person"], q["intext"], q["notes"]] + [""] * 10
        for i, v in enumerate(vals):
            body(ws.cell(rr, 2 + i, v), REFF if i == 0 else (GOLD if 8 <= i else None), bold=i == 0)
            if i >= 8:
                ws.cell(rr, 2 + i).fill = PatternFill("solid", fgColor="FFF2CC")
        ws.row_dimensions[rr].height = 75
    last = 5 + len(quan)
    lists_tmp = {"yn": "=Lists!$E$2:$E$4"}
    dv_list(ws, lists_tmp["yn"], f"J6:J{last}"); dv_list(ws, lists_tmp["yn"], f"R6:R{last}")
    ws.freeze_panes = "E6"
    widths(ws, dict(zip("BCDEFGHIJKLMNOPQRS", [9, 18, 40, 9, 40, 18, 11, 36, 12, 24, 18, 12, 18, 14, 26, 18, 12, 14])))

    # 3. Stakeholder map & RASCI
    ws = wb.create_sheet("3. Stakeholder map & RASCI")
    title(ws, "End-2026 target status assessment and reporting", "3. Stakeholder map & RASCI")
    hdr_row(ws, 4, ["SR26 unit", "SR26 full name", "SR25 equivalent (2025 master)", "Section / team", "Key contact", "Status / notes"])
    units = [("Faculty", name, f"Faculties tab – {code}", "Associate Dean Sustainability / faculty sustainability lead", "TBC", "Pilot faculty") for code, name in PILOT_FACULTIES.items()]
    for code, name, sr25name, secs in CENTRAL:
        for label, _ in secs:
            units.append((code, name, sr25name, label, "TBC", ""))
    units += [
        ("Chancellery", "Advancement, Communications & Marketing; Global, Culture & Engagement; Education (incl. SASS); Indigenous", "Chancellery tab",
         "—", "TBC", "No SR26 qualitative indicator currently names these units. SASS (Melbourne Plus) referenced for EE1(c). Confirm whether requests are needed."),
        ("Legal & Risk", "Legal and Risk", "Legal and Risk tab", "—", "TBC", "No SR26 indicator currently assigned – may be relevant to climate resilience maturity (CL3(a)). TBC."),
        ("TBC", "CGOP / CDEP / CDSS", "Not in SR25 master", "—", "TBC", "Structure changes TBC – confirm before requests are issued."),
    ]
    for k, u in enumerate(units):
        for i, v in enumerate(u):
            body(ws.cell(5 + k, 2 + i, v))
    widths(ws, {"A": 4, "B": 14, "C": 40, "D": 44, "E": 40, "F": 18, "G": 50})
    r0 = 7 + len(units)
    ws.cell(r0 - 1, 2, "RASCI – who receives a request for each indicator (R = Responsible to report; TBC = to confirm)").font = Font(name="Aptos", bold=True, size=14, color=NAVY)
    mcols = ["Ref", "Priority area", "Indicator"] + list(PILOT_FACULTIES) + [c[0] for c in CENTRAL]
    hdr_row(ws, r0, mcols)
    for k, r in enumerate(qual):
        rr = r0 + 1 + k
        vals = [r["ref"], r["pa"], r["indicator"]]
        vals += ["R" if "Faculties" in r["stake"] else "" for _ in PILOT_FACULTIES]
        vals += ["R" if any(s[1](r) for s in c[3]) else "" for c in CENTRAL]
        for i, v in enumerate(vals):
            cell = ws.cell(rr, 2 + i, v); body(cell, REFF if i == 0 else None)
            if i >= 3:
                cell.alignment = CWRAP
                if v:
                    cell.fill = PatternFill("solid", fgColor="E2EFDA")
        ws.row_dimensions[rr].height = 45
    for i in range(len(mcols) - 3):
        ws.column_dimensions[get_column_letter(5 + i)].width = max(ws.column_dimensions[get_column_letter(5 + i)].width or 0, 10)
    ws.column_dimensions["D"].width = 44

    # 4. Request tracker
    ws = wb.create_sheet("4. Request tracker")
    title(ws, "End-2026 target status assessment and reporting", "4. Request tracker")
    cols = ["#", "Group", "Unit / section", "Key contact", "No. of indicators requested", "Request workbook", "Date sent",
            "Due date", "Reminder date", "Support meeting offered/held", "Response received", "Owner confirmed",
            "Status", "Days overdue", "Notes / follow-up"]
    hdr_row(ws, 9, cols)
    track = [(f"F{i+1}", "Faculty", f"{code} – {name}", len(fac_rows), f"{code} - end-2026 sustainability reporting request.xlsx")
             for i, (code, name) in enumerate(PILOT_FACULTIES.items())]
    n = 0
    for code, name, _, secs in CENTRAL:
        for label, rows in central_sections[code]:
            n += 1
            track.append((f"C{n}", "Central", label, len(rows), f"{code.replace('&', 'and')} - end-2026 sustainability reporting request.xlsx"))
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
    def unit_tab(tab, subtitle, sections, sr25_for):
        ws = wb.create_sheet(tab)
        title(ws, "End-2026 target status assessment and reporting", subtitle)
        setup_block_sheet(ws)
        ws["B3"] = "Consolidated responses – paste from returned request workbooks (identical column layout). Click [+] above Columns H–K for end-2025 responses."
        ws["B3"].font = Font(name="Aptos", italic=True)
        top = 5
        for idx, (label, rows, sr) in enumerate(sections):
            top = block(ws, top, f"1.{idx + 1}  {label}", rows, sr, LISTS)
        return ws

    global LISTS
    lists = lists_sheet(wb); LISTS = lists
    unit_tab("Faculties", "Faculties", [(f"{c} – {n}", fac_rows, sr25_all.get(c, {})) for c, n in PILOT_FACULTIES.items()], True)
    for code, name, _, secs in CENTRAL:
        unit_tab(code, name, [(label, rows, None) for label, rows in central_sections[code]], False)

    # Highlighted stories
    st = wb.create_sheet("Highlighted stories")
    stories_sheet(st, unit=None)

    # Evidence register
    ws = wb.create_sheet("Evidence register")
    title(ws, "End-2026 target status assessment and reporting", "Evidence register")
    ws["B3"] = ("Every material figure or claim needs a named owner, definition, source, reporting boundary, period, quality note and confirmation record (SR26 approach – Data quality).")
    ws["B3"].font = Font(name="Aptos", italic=True)
    cols = ["Evidence ID", "Ref", "Indicator", "Faculty/Portfolio", "Claim / figure used in report", "Evidence description",
            "Link / file location", "Source system", "Reporting period", "Reporting boundary", "Limitations / quality note",
            "Data owner (name, role)", "Owner confirmed?", "Confirmation date", "SST reviewer", "Report section", "Databook aligned?", "Status / notes"]
    hdr_row(ws, 5, cols)
    for r in range(6, 206):
        ws.cell(r, 2, f'=IF(C{r}="","","EV-"&TEXT(ROW()-5,"000"))')
        for c in range(2, 2 + len(cols)):
            body(ws.cell(r, c))
    dv_list(ws, "=Lists!$E$2:$E$4", "N6:N205"); dv_list(ws, "=Lists!$E$2:$E$4", "R6:R205")
    widths(ws, dict(zip("BCDEFGHIJKLMNOPQRS", [10, 9, 30, 18, 40, 36, 30, 16, 12, 16, 30, 22, 11, 13, 14, 18, 11, 26])))
    ws.freeze_panes = "D6"

    # Issues log
    ws = wb.create_sheet("Issues log")
    title(ws, "End-2026 target status assessment and reporting", "Issues log")
    hdr_row(ws, 4, ["#", "Date raised", "Issue / question", "Ref / unit", "Owner", "Proposed action", "Status"])
    issues = [
        ("Climate resilience maturity (CL3(a)) and Estate & infrastructure indicators – confirm owner and request with Gerard.", "CL3(a); EI1(a)–(d)", "Stefanus", "Meet Gerard", "Open"),
        ("Confirm whether CGOP, CDEP and CDSS structures have changed since the 2025 master spreadsheet stakeholder mapping.", "Stakeholder map", "Stefanus", "Confirm with team", "Open"),
        ("Business Services split into CIOG (AI) and ESG (nature & biodiversity, waste & circular economy, part of quant. climate leadership) – confirm contacts.", "CIOG; ESG", "TBC", "Confirm contacts", "Open"),
        ("Procurement sits with CFOG (RP1(a) procurement case studies routed to CFOG – Procurement although the Qual sheet names no stakeholder); estate planning centralised in CI&S – requests issued centrally, not to faculties.", "CFOG; CI&S", "—", "Noted", "Closed"),
        ("Chancellery units and Legal & Risk had SR25 requests but no SR26 qualitative indicator names them – confirm if requests are needed.", "Stakeholder map", "TBC", "Decide", "Open"),
        ("TR2(b) 'Documented progress of strategic initiatives, incl. Impact Accelerators' reuses the TR2(d) case-study wording and has no stakeholder in the Qual sheet – left as-is.", "TR2(b)", "TBC", "Review wording/owner", "Open"),
        ("Several rows in the Qual sheet are quantitative (GHG inventory, intensity, waste, procurement spend %, AUM %, biodiversity metrics) and have no stakeholder – routed to the Databook mapping tab, not requested from units.", "CL1(a)-(b); CL2(b); NB1(a); CE1-2; RP2; RI1(a)", "TBC", "Confirm Databook coverage", "Open"),
        ("Target status rating (Met/Partially met/Not met) retained pending revamp of the traffic-light framework.", "All", "Director, Sustainability", "Update templates once agreed", "Open"),
        ("Due dates are placeholders based on the SR26 approach timeline (stories early Dec; requests due January).", "All", "TBC", "Confirm dates", "Open"),
        ("20 indicators are requested from each faculty (SR25 sent 10). Monitor burden in pilot.", "Faculties", "TBC", "Review after pilot", "Open"),
    ]
    from datetime import date
    for k, it in enumerate(issues):
        vals = [k + 1, date(2026, 10, 5)] + list(it)
        for i, v in enumerate(vals):
            body(ws.cell(5 + k, 2 + i, v))
        ws.cell(5 + k, 3).number_format = "dd/mm/yyyy"
    widths(ws, {"A": 4, "B": 5, "C": 12, "D": 70, "E": 24, "F": 18, "G": 26, "H": 10})

    wb.move_sheet("Lists", offset=len(wb.sheetnames))
    OUT.mkdir(exist_ok=True)
    p = OUT / "SR26 end-year reporting master spreadsheet - V0.1.xlsx"
    wb.save(p)
    return p.name


def main():
    qual = read_qual()
    by_row = {r["row"]: r for r in qual}
    quan = read_quan(by_row)
    sr25 = read_sr25_faculty()
    fac_rows = [r for r in qual if "Faculties" in r["stake"]]
    central_sections = {code: [(label, [r for r in qual if m(r)]) for label, m in secs] for code, _, _, secs in CENTRAL}
    print("Qual indicators:", len(qual), "| faculty rows:", len(fac_rows), "| quant rows:", len(quan))
    for code, name in PILOT_FACULTIES.items():
        print(" ", request_workbook(code, name, [(f"{code} – {name}", fac_rows)], sr25.get(code, {})))
    for code, name, _, _ in CENTRAL:
        secs = [(l, rows) for l, rows in central_sections[code] if rows]
        print(" ", request_workbook(code, name, secs, None), [len(r) for _, r in secs])
    print(" ", master(qual, quan, fac_rows, central_sections, sr25))


if __name__ == "__main__":
    main()
