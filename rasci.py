"""SR26 RASCI matrix, rebuilt to the SR25 'RASCI' tab standard (Sustainability Plan 2030 Implementation Roles):
Domain > Priority area > Target > Indicator hierarchy, consolidated portfolio columns with grouped sub-units,
SST contact per area, A / R / S / C / I codes, colour coding and COUNTIF summary rows."""
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# SST contact per area, carried from the SR25 RASCI (row 3) – to confirm for SR26
SR25_CONTACT = {"ABP": "Gerard", "ARTS": "Rose", "FBE": "Katie", "SCI": "Katie",
                "CI&S – Estate Planning & Development": "Gerard and Davina", "CI&S – Sustainability Strategy": "Director",
                "CFOG – Procurement": "Chris", "CFOG – Treasury & Investments": "Chris", "ESG": "All", "CIOG": "TBC",
                "MRE": "Director", "L&R": "Gerard", "SST": "All", "FAC*": "Director", "COO*": "Director"}
CODES = [("A", "Accountable"), ("A and R", "Accountable and Responsible"), ("R", "Responsible – narrative request"),
         ("R (quant)", "Responsible – figures for the Databook (Tab 4)"), ("R (compile)", "SST compiles from other responses / no request"),
         ("S", "Support"), ("C", "Consulted"), ("I", "Informed")]
FILL = {"A and R": "00B050", "A": "92D050", "R": "C6EFCE", "R (quant)": "BDD7EE", "R (compile)": "D9D9D9",
        "S": "FFEB9C", "C": "F8CBAD", "I": "EDEDED"}


SR26_CODES = {}  # ref -> {SR26 area label: code}, filled by rasci_tab for the triangulation tab


def rasci_tab(wb, qual, coverage, faculties, central):
    import build as B
    import quant
    ws = wb.create_sheet("3b. RASCI matrix")
    ws.sheet_view.showGridLines = False
    ws["B1"] = "Sustainability 2030 – SR26 reporting roles (RASCI)"
    ws["B1"].font = Font(name="Aptos", size=20, bold=True, color=B.NAVY)
    ws["B2"] = ("Who is Accountable / Responsible / Support / Consulted / Informed for reporting each Sustainability 2030 indicator in SR26. "
                "Green columns are consolidated portfolios (click [+]/[–] above to show or hide their areas). Codes are pre-filled from the "
                "requests issued; edit with the dropdown. Accountable owners (A) are TBC – to confirm with the Director, Sustainability.")
    ws["B2"].alignment = B.WRAP; ws.merge_cells("B2:G2"); ws.row_dimensions[2].height = 48

    # ---- columns: (portfolio, area label, key, is_consolidated)
    cols = [("Faculty", "Faculties – pilot (all)", "FAC*", True)]
    cols += [("Faculty", f"{code} – {name.replace('Faculty of ', '')}", code, False) for code, name in faculties.items()]
    coo = [(c, n, secs) for c, n, _, secs in central if c != "MRE"]
    cols.append(("COO portfolio", "COO portfolio (all)", "COO*", True))
    for code, name, secs in coo:
        for label, _ in secs:
            key = label.split(" (")[0].split(" – ")[0] if len(secs) == 1 else label.split(" (")[0]
            cols.append(("COO portfolio", label.split(" (")[0], key, False))
    cols.append(("Chancellery", "MRE – Melbourne Research and Enterprise", "MRE", False))
    cols.append(("Sustainability Strategy", "SST (compiles / Databook)", "SST", False))
    first = 9  # first area column (I), as SR25
    col_of = {c[2]: first + i for i, c in enumerate(cols)}

    hdr = ["Number", "Category", "Domain", "Priority area", "Target / indicator", "Type", "Related SR25 ref", "Activity area >>"]
    for i, h in enumerate(hdr):
        ws.cell(6, 1 + i, h)
    ws.cell(3, 8, "SST contact (SR25 – confirm) >>"); ws.cell(4, 8, "Portfolio / Faculty >>")
    ws.merge_cells(start_row=5, start_column=first, end_row=5, end_column=first + len(cols) - 1)
    ws.cell(5, first, "Areas Accountable and/or Responsible for SR26 reporting (green = consolidated portfolio, white = individual area)")
    for i, (port, label, key, cons) in enumerate(cols):
        c = first + i
        ws.cell(3, c, SR25_CONTACT.get(key, SR25_CONTACT.get(key.split(" –")[0], "TBC")))
        ws.cell(4, c, port + (" (consolidated)" if cons else ""))
        ws.cell(6, c, label)
        for r in (3, 4, 6):
            x = ws.cell(r, c); x.alignment = B.CWRAP; x.border = B.BOX
            x.font = Font(name="Aptos", bold=r == 6, size=10 if r != 6 else 11)
        ws.cell(6, c).fill = PatternFill("solid", fgColor="00B050" if cons else "FFFFFF")
        ws.column_dimensions[get_column_letter(c)].width = 13
        if not cons and port in ("Faculty", "COO portfolio"):
            ws.column_dimensions[get_column_letter(c)].outlineLevel = 1
    for i in range(1, 9):
        x = ws.cell(6, i); x.fill = B.HDR; x.font = B.WHITE_B; x.alignment = B.CWRAP; x.border = B.BOX
    ws.cell(5, first).font = Font(name="Aptos", bold=True); ws.cell(5, first).alignment = B.CWRAP
    ws.row_dimensions[6].height = 60

    # ---- who does what per indicator
    cov = {c["ref"]: c for c in coverage}
    purged = quant.quant_only_rows(qual)

    def codes_for(q):
        d = {}
        if "Faculties" in q["stake"]:
            for code in faculties:
                d[code] = "R"
        for code, name, _, secs in central:
            for label, m in secs:
                if m(q):
                    key = "MRE" if code == "MRE" else next(c[2] for c in cols if c[1] == label.split(" (")[0])
                    d[key] = "R"
        c = cov.get(q["ref"])
        if c:
            own = c["owner"]
            if own == "SST":
                d.setdefault("SST", "R (compile)")
            else:
                if own == "CFOG":  # two CFOG sections: investments vs procurement
                    own = "CFOG – Treasury & Investments" if q["ref"].startswith("RI") else "CFOG – Procurement"
                key = own if own in col_of else next((k for k in col_of if k.startswith(own)), None)
                if key and d.get(key) != "R" and c["npts"]:
                    d[key] = "R (quant)"
                elif key and c["npts"]:
                    d[key] = "R"  # narrative + figures from the same area
        if not d:
            d["SST"] = "R (compile)"
        d["SST"] = d.get("SST", "S")  # SST supports every indicator
        if any(k in faculties for k in d):
            d["FAC*"] = "A and R"
        if any(cols[col_of[k] - first][0] == "COO portfolio" for k in d if k in col_of):
            d["COO*"] = "A and R"
        SR26_CODES[q["ref"]] = {cols[col_of[k] - first][1]: v for k, v in d.items() if k in col_of and not k.endswith("*")}
        return d

    r, num = 7, {"dom": 0}
    dom = pa = tgt = None
    data_rows = []
    dfill = PatternFill("solid", fgColor="D9D9D9"); pfill = PatternFill("solid", fgColor="F2F2F2")
    for q in qual:
        if q["domain"] != dom:
            dom = q["domain"]; num["dom"] += 1
            for i, v in enumerate([f"D{num['dom']}", "Domain", "", "", dom]):
                ws.cell(r, 1 + i, v)
            for c in range(1, first + len(cols)):
                ws.cell(r, c).fill = dfill; ws.cell(r, c).font = Font(name="Aptos", bold=True)
            r += 1
        if q["pa"] != pa:
            pa = q["pa"]
            for i, v in enumerate([q["ref"][:2], "Priority", "", "", pa]):
                ws.cell(r, 1 + i, v)
            for c in range(1, first + len(cols)):
                ws.cell(r, c).fill = pfill; ws.cell(r, c).font = Font(name="Aptos", bold=True)
            r += 1
        if q["target"] != tgt:
            tgt = q["target"]
            for i, v in enumerate([q["ref"].split("(")[0], "Target", dom, pa, tgt]):
                c = ws.cell(r, 1 + i, v); c.alignment = B.WRAP; c.font = Font(name="Aptos", bold=True, size=10)
            ws.row_dimensions[r].height = 45
            r += 1
        kind = "Quantitative" if q["row"] in purged else ("Mixed" if q["ref"] in cov else "Qualitative")
        for i, v in enumerate([q["ref"], "Indicator", dom, pa, q["indicator"], kind, B.SR25_MAP.get(q["row"], "")]):
            c = ws.cell(r, 1 + i, v); c.alignment = B.WRAP; c.border = B.BOX; c.font = Font(name="Aptos Narrow", size=10)
        for k, v in codes_for(q).items():
            if k in col_of:
                ws.cell(r, col_of[k], v)
        for c in range(first, first + len(cols)):
            x = ws.cell(r, c); x.alignment = B.CWRAP; x.border = B.BOX; x.font = Font(name="Aptos Narrow", size=10)
        ws.row_dimensions[r].height = 42
        data_rows.append(r)
        r += 1
    top, bot = data_rows[0], data_rows[-1]
    last_col = get_column_letter(first + len(cols) - 1)
    area = f"{get_column_letter(first)}{top}:{last_col}{bot}"
    B.dv_list(ws, '"' + ",".join(c for c, _ in CODES) + '"', area)
    for code, colour in FILL.items():
        ws.conditional_formatting.add(area, FormulaRule(formula=[f'EXACT({get_column_letter(first)}{top},"{code}")'],
                                                        fill=PatternFill("solid", fgColor=colour)))

    # ---- legend + summary (as SR25 rows 50–61)
    r = bot + 2
    ws.cell(r, 5, "Summary").font = Font(name="Aptos", bold=True, size=14, color=B.NAVY)
    r += 1
    ws.cell(r, 5, "Area >>").font = Font(name="Aptos", bold=True)
    for i, c in enumerate(cols):
        ws.cell(r, first + i, f"={get_column_letter(first + i)}6").font = Font(name="Aptos", bold=True, size=9)
        ws.cell(r, first + i).alignment = B.CWRAP
    ws.row_dimensions[r].height = 45
    rows_at = {}
    for code, desc in CODES:
        r += 1
        ws.cell(r, 5, desc); ws.cell(r, 6, code).fill = PatternFill("solid", fgColor=FILL[code])
        rows_at[code] = r
        for i in range(len(cols)):
            L = get_column_letter(first + i)
            ws.cell(r, first + i, f'=COUNTIF({L}{top}:{L}{bot},$F{r})').alignment = B.CWRAP
    for label, f in [("Total A", lambda L: f"={L}{rows_at['A']}+{L}{rows_at['A and R']}"),
                     ("Total R (all types)", lambda L: f"={L}{rows_at['A and R']}+{L}{rows_at['R']}+{L}{rows_at['R (quant)']}+{L}{rows_at['R (compile)']}"),
                     ("Total indicators involved", lambda L: f"=SUM({L}{rows_at['A']}:{L}{rows_at['I']})")]:
        r += 1
        ws.cell(r, 5, label).font = Font(name="Aptos", bold=True)
        for i in range(len(cols)):
            x = ws.cell(r, first + i, f(get_column_letter(first + i))); x.font = Font(name="Aptos", bold=True); x.alignment = B.CWRAP
    B.widths(ws, {"A": 8, "B": 9, "C": 14, "D": 16, "E": 50, "F": 12, "G": 10, "H": 4})
    ws.freeze_panes = f"{get_column_letter(first)}7"
    ws.sheet_properties.outlinePr.summaryRight = False
    ws.auto_filter.ref = f"A6:{last_col}{bot}"


# ---------------------------------------------------------------- SR25 -> SR26 triangulation
# S2030 Qual row -> equivalent SR25 (SP2030) target(s) in the SR25 RASCI. Judgement-based; [] = new in Sustainability 2030.
SR25_TARGETS = {8: ["4a"], 9: ["4b"], 10: ["3c"], 11: ["4b", "3e"], 12: ["5a", "2a"], 13: ["5a"], 14: ["5a"], 15: ["5a"],
                16: ["5b"], 17: ["5b"], 18: ["5b"], 19: ["5b"], 20: ["5c"], 21: ["5c"], 22: ["5c", "3d"], 23: ["5c"],
                24: ["1a"], 25: ["1a"], 26: ["1a", "1b"], 27: ["1b"], 28: ["1b"], 29: ["1b"], 30: ["8a"],
                31: ["9a", "9b"], 32: ["9a"], 33: ["9b"], 34: ["9b"], 35: ["11c"], 36: ["11c"], 37: ["11c"], 38: [], 39: [],
                40: ["7b", "8a"], 41: ["7b"], 42: ["7b"], 43: ["7b", "6a"], 44: ["11a"], 45: ["11a", "11b"], 46: ["11c"],
                47: ["11d"], 48: ["11d"], 49: ["11c"], 50: ["12a"], 51: ["12b"], 52: ["6a"], 53: ["2a"], 54: ["7a", "7c", "3a"],
                55: ["3a"]}
# SR25 RASCI area -> SR26 area (None = consolidated column, skipped; 'Not in SR26: …' = area with no SR26 request)
SR25_AREA = {
    "Faculties (all)": None, "Office of the Provost (all)": None, "COO Portfolio (all)": None,
    "Chancellery Global, Culture & Engagement (all)": None, "CD consolidated": None,
    "Architecture, Building and Planning": "ABP – Architecture, Building and Planning", "Arts": "ARTS – Arts",
    "Business and Economics / MBS": "FBE – Business and Economics", "Science": "SCI – Science",
    "Education": "Non-pilot faculties", "Engineering and IT": "Non-pilot faculties", "Fine Arts and Music": "Non-pilot faculties",
    "Law": "Non-pilot faculties", "MDHS": "Non-pilot faculties",
    "Academic": "Not in SR26: Provost / Chancellery Education", "Indigenous": "Not in SR26: Chancellery Indigenous",
    "People strategy?": "Not in SR26: People strategy", "Advancement, Communications & Marketing": "Not in SR26: Advancement (ACM)",
    "Chancellery Research and Enterprise": "MRE – Melbourne Research and Enterprise",
    "Chancellery Global": "Not in SR26: Chancellery Global, Culture & Engagement",
    "Community and cultural partnerships": "Not in SR26: Chancellery Global, Culture & Engagement",
    "CD Sustainability Strategy": "CI&S – Sustainability Strategy", "CD Treasury & Investments": "CFOG – Treasury & Investments",
    "CD Estate (development,strategy,planning and perfomance)": "CI&S – Estate Planning & Development",
    "CD EPMO and investment office": "Not in SR26: EPMO & investment office", "CFOG (Finance)": "Not in SR26: CFOG Finance",
    "CFOG (Procurement)": "CFOG – Procurement", "Campus Management - Sustainability Delivery": "ESG – Campus Operations & Sustainability Delivery",
    "Campus management - project delivery": "ESG – Campus Operations & Sustainability Delivery",
    "Student and scholarly services": "Not in SR26: Student & Scholarly Services", "RIC": "Not in SR26: RIC",
    "Legal and Risk": "Legal & Risk – Risk & resilience", "EPG": "Not in SR26: EPG",
    "University governance": "Not in SR26: University governance", "Academic Board": "Not in SR26: Academic Board",
}
STRONG = ("A", "R")  # SR25 codes that count as a reporting responsibility (A, R, A and R, R (air travel)...)


def read_sr25_rasci():
    """{SR25 target: (text, {SR26 area: SR25 code})}"""
    from openpyxl import load_workbook
    import build as B
    ws = load_workbook(B.SRC / "SR25_master.xlsx", data_only=True)["RASCI"]
    areas = {c: " ".join(str(ws.cell(6, c).value).split()) for c in range(9, ws.max_column + 1) if ws.cell(6, c).value}
    out = {}
    for r in range(7, ws.max_row + 1):
        if ws.cell(r, 2).value == "Target":
            d = {}
            for c, a in areas.items():
                v = ws.cell(r, c).value
                m = SR25_AREA.get(a)
                if v not in (None, "") and m:
                    v = str(v).strip()
                    if m not in d or (v.startswith(STRONG) and not d[m].startswith(STRONG)):
                        d[m] = v
            out[str(ws.cell(r, 1).value)] = (str(ws.cell(r, 5).value), d)
    return out


def triangulation_tab(wb, qual):
    import build as B
    from openpyxl.styles import Alignment, Font, PatternFill
    sr25 = read_sr25_rasci()
    ws = wb.create_sheet("3c. RASCI triangulation")
    B.title(ws, "End-2026 target status assessment and reporting", "3c. RASCI triangulation (SR25 → SR26)")
    ws["B3"] = ("Cross-check of the SR26 RASCI (3b) against the SR25 RASCI. Each SR26 indicator is mapped to its equivalent SR25 target(s) "
                "(judgement – see column D); SR25 areas are translated to the SR26 structure. 'SR25 only' flags an area that held a "
                "responsibility in SR25 but has no SR26 role – check whether it should be asked, consulted or informed. "
                "Non-pilot faculties are excluded from gaps (pilot only). SR26 codes are not changed by this tab.")
    ws["B3"].alignment = B.WRAP; ws.merge_cells("B3:J3"); ws.row_dimensions[3].height = 60
    cols = ["Ref", "Indicator", "Equivalent SR25 target(s)", "SR25 responsibilities (translated to SR26 areas)",
            "SR26 responsibilities (3b)", "In both", "SR25 only – check", "New in SR26", "Suggested action", "Decision / notes"]
    B.hdr_row(ws, 5, cols)
    gap_count = {}
    r = 6
    for q in qual:
        tg = SR25_TARGETS.get(q["row"], [])
        s25 = {}
        for t in tg:
            for a, v in sr25.get(t, ("", {}))[1].items():
                if a == "Non-pilot faculties":
                    continue
                if a not in s25 or (v.startswith(STRONG) and not s25[a].startswith(STRONG)):
                    s25[a] = v
        s26 = {a: v for a, v in SR26_CODES.get(q["ref"], {}).items() if not a.startswith("SST")}
        r25 = {a for a, v in s25.items() if v.startswith(STRONG)}
        both = sorted(r25 & set(s26))
        only25 = sorted(r25 - set(s26))
        new26 = sorted(set(s26) - set(s25))
        for a in only25:
            gap_count[a] = gap_count.get(a, 0) + 1
        if not tg:
            action = "New indicator – no SR25 equivalent; SR26 roles stand."
        elif not only25:
            action = "Consistent with SR25." + (" New areas added in SR26." if new26 else "")
        else:
            outside = [a for a in only25 if a.startswith("Not in SR26")]
            inside = [a for a in only25 if not a.startswith("Not in SR26")]
            parts = []
            if inside:
                parts.append("Check why " + ", ".join(x.split(" – ")[0] for x in inside) + " is not asked (was responsible in SR25)")
            if outside:
                parts.append("Consider request or 'C/I' for " + ", ".join(x.replace("Not in SR26: ", "") for x in outside))
            action = "; ".join(parts) + "."
        fmt = lambda d: "\n".join(f"{a.replace('Not in SR26: ', '⚠ ')}: {v}" for a, v in sorted(d.items()))
        tg_txt = "\n".join(f"{t} – {sr25.get(t, ('?',))[0][:70]}" for t in tg) or "None (new in Sustainability 2030)"
        vals = [q["ref"], q["indicator"], tg_txt, fmt(s25) or "–", fmt(s26) or "SST compiles", "\n".join(both),
                "\n".join(only25), "\n".join(new26), action, ""]
        for i, v in enumerate(vals):
            B.body(ws.cell(r, 2 + i, v), B.REFF if i == 0 else None, bold=i == 0)
        if only25:
            ws.cell(r, 8).fill = PatternFill("solid", fgColor="F8CBAD")
        if new26:
            ws.cell(r, 9).fill = PatternFill("solid", fgColor="DDEBF7")
        ws.row_dimensions[r].height = max(60, 14 * max(len(s25), len(s26), 1))
        r += 1
    last = r - 1
    # summary by area
    r += 2
    ws.cell(r, 2, "SR25 areas with responsibilities but no SR26 role (number of indicators)").font = Font(name="Aptos", bold=True, size=14, color=B.NAVY)
    r += 1
    B.hdr_row(ws, r, ["Area", "Indicators", "In SR26 structure?"])
    for a, n in sorted(gap_count.items(), key=lambda x: -x[1]):
        r += 1
        for i, v in enumerate([a.replace("Not in SR26: ", ""), n, "No – no SR26 request" if a.startswith("Not in SR26") else "Yes – check"]):
            B.body(ws.cell(r, 2 + i, v))
    B.widths(ws, dict(zip("BCDEFGHIJK", [9, 36, 30, 36, 30, 22, 26, 22, 38, 26])))
    ws.freeze_panes = "D6"
    ws.auto_filter.ref = f"B5:K{last}"
    return gap_count
