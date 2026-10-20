"""Build 06_Segment_Prioritization.xlsx — v2 (post partner review, 2026-10-19; review issue C5).

v2 principles:
  * one indicator per criterion (no input feeds two criteria);
  * Phase 4 findings fed back: revenue potential = usage-based first-year ACV per won account at list prices;
  * market attractiveness = reachable accounts (no separate headroom flag; existing Clay customers are excluded instead);
  * product-market fit = capability-fit index from the 05 pain-to-capability map (not a free judgment);
  * the coverage proxy enters only Competitive Position (via 04 factor F1).
All scores are live formulas. Random weights are generated in Python (seed 42) and stored as values; their totals are formulas.
v1 is archived in appendix/archive/.
"""
import math
import os
import sys
import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter as L

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(HERE)
OUT = f"{BASE}/06_Segment_Prioritization.xlsx"
sys.path.insert(0, HERE)
import clay_engine as ce  # noqa: E402

F = "Arial"
BLUE = Font(name=F, color="0000FF", size=10)
BLACK = Font(name=F, color="000000", size=10)
GREEN = Font(name=F, color="008000", size=10)
HDR = Font(name=F, bold=True, color="FFFFFF", size=10)
TITLE = Font(name=F, bold=True, size=14, color="1F2A44")
SUBF = Font(name=F, italic=True, size=9, color="555555")
HFILL = PatternFill("solid", fgColor="1F2A44")
SFILL = PatternFill("solid", fgColor="E8ECF3")
YFILL = PatternFill("solid", fgColor="FFF2CC")
KFILL = PatternFill("solid", fgColor="C6E0B4")
RFILL = PatternFill("solid", fgColor="F8CBAD")
thin = Side(style="thin", color="BFBFBF")
BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
USD = '$#,##0;($#,##0);"-"'
PCT = '0.0%'
D2 = '0.00'
N0 = '#,##0'
X2 = '0.00"x"'

SEGS = ["A. Industrial Distribution", "B. Industrial Manufacturing", "C. B2B Prof. Services"]
SHORT = ["Distribution", "Manufacturing", "Prof. services"]
CRIT = ["Market Attractiveness", "Customer Need", "Product-Market Fit", "Revenue Potential", "Competitive Position", "Ease of Execution"]
QUANT = [True, True, False, True, False, False]   # scaled from raw metric vs direct 1–5

wb = Workbook()


def put(ws, ref, val, font=BLACK, fmt=None, fill=None, bold=False, wrap=True, border=True, center=False, italic=False):
    c = ws[ref]
    c.value = val
    c.font = Font(name=F, size=10, bold=bold, italic=italic, color=font.color)
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    c.alignment = CENTER if center else (WRAP if wrap else Alignment(vertical="top"))
    if border:
        c.border = BOX
    return c


def hdr(ws, row, labels, col=1, widths=None, height=30):
    for i, t in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=t)
        c.font, c.fill, c.alignment, c.border = HDR, HFILL, CENTER, BOX
    if widths:
        for i, w in enumerate(widths):
            if w:
                ws.column_dimensions[L(col + i)].width = w
    ws.row_dimensions[row].height = height


def title(ws, t, s):
    ws["A1"].value, ws["A1"].font = t, TITLE
    ws["A2"].value, ws["A2"].font = s, SUBF
    ws.sheet_view.showGridLines = False


def note(ws, rng, text, height=36):
    a = rng.split(":")[0]
    put(ws, a, text, border=False)
    ws.merge_cells(rng)
    ws.row_dimensions[int("".join(ch for ch in a if ch.isdigit()))].height = height


cover = wb.active
cover.title = "Cover"
exs = wb.create_sheet("Executive_Summary")
wt = wb.create_sheet("Weights")
inp = wb.create_sheet("Inputs")
sc = wb.create_sheet("Scorecard")
brg = wb.create_sheet("Bridge_v1_v2")
wsn = wb.create_sheet("Weight_Sensitivity")
rw = wb.create_sheet("Random_Weights")
stt = wb.create_sheet("Stress_Tests")
eco = wb.create_sheet("Economic_Crosscheck")
dd = wb.create_sheet("Subsegment_Drilldown")
chg = wb.create_sheet("Change_Log")
srcs = wb.create_sheet("Sources")

# ======================================================================= WEIGHTS
title(wt, "Weights — the brief's suggested weights plus five alternative sets (unchanged from v1 for comparability)", "Blue = input. Each column must sum to 100%.")
WSETS = [("Base (brief)", [.20, .20, .20, .15, .15, .10]), ("Equal", [1 / 6] * 6), ("Market-size first", [.30, .15, .15, .25, .10, .05]),
         ("Execution first", [.15, .15, .20, .10, .10, .30]), ("Differentiation first", [.15, .20, .15, .10, .30, .10]), ("Near-term revenue", [.15, .10, .20, .30, .10, .15])]
hdr(wt, 4, ["Criterion"] + [n for n, _ in WSETS] + ["Single indicator used in v2"], widths=[24, 12, 12, 14, 14, 16, 14, 70])
IND = ["Reachable accounts: unique firms (20+) × CRM in use × (1 − already Clay customers)",
       "Sales intensity: sales occupations share of employment ÷ all-industry share",
       "Capability-fit index: mean fit of Clay's current product across the segment's core jobs (05 §3), builder-capacity gap excluded",
       "Revenue per account: usage-based first-year ACV per won account at list prices, account-mix weighted (Phase 4 method)",
       "Competitive position: 04 factors F1 data differentiation, F2 incumbent saturation, F3 specialist threat (the coverage proxy enters only here, via F1)",
       "Adoption / execution fit: 04 factors F4 builder capacity, F5 references and partner channel"]
for i, c in enumerate(CRIT):
    r = 5 + i
    put(wt, f"A{r}", c, bold=True)
    for j, (_, w) in enumerate(WSETS):
        put(wt, f"{L(2 + j)}{r}", w[i], font=BLUE, fmt='0.0%', fill=YFILL if j == 0 else None, center=True)
    put(wt, f"H{r}", IND[i])
    wt.row_dimensions[r].height = 30
put(wt, "A11", "Check: sum", bold=True, fill=SFILL)
for j in range(len(WSETS)):
    c = L(2 + j)
    put(wt, f"{c}11", f"=SUM({c}5:{c}10)", fmt='0.0%', bold=True, fill=SFILL, center=True)
W_BASE = "Weights!$B$5:$B$10"

# ======================================================================= INPUTS
title(inp, "Inputs — one indicator per criterion (v2)", "Blue = input copied from the named source; black = formula. Yellow = analyst judgment.")
hdr(inp, 4, ["ID", "Indicator", "Criterion", SEGS[0], SEGS[1], SEGS[2], "Unit", "Class", "Source"], widths=[7, 52, 20, 17, 17, 17, 10, 8, 70])
IR = {}
rows = [
    ("I1", "Unique U.S. firms with 20+ employees (de-duplicated)", "", [11494, 23785, 33035], N0, "firms", "C from A", "02_Market_Sizing › Firm_Counts (Census 2022, S14)", False),
    ("I2", "Share with a CRM in use (integrable today)", "", [.56, .60, .70], PCT, "%", "B / D", "Dist: Phocas/DSG 2026, 56% use CRM (S25). Mfg, PS: Phase 1 assumptions (02 A8, A9)", True),
    ("I3", "Share already Clay customers (excluded)", "", [.03, .03, .10], PCT, "%", "D", "07 v2 Assumptions; PS higher: Clay markets to PS (S06), IT services top detected industry (S32)", True),
    ("I4", "Reachable accounts = I1 × I2 × (1 − I3)", "Market Attractiveness", None, N0, "accounts", "C", "Formula", False),
    ("I5", "Sales occupations share of employment", "", [.21956, .027390, .061719], PCT, "%", "A → C", "BLS OEWS May 2025 (S15) via 02 Segment_Profile", False),
    ("I6", "Sales intensity index = I5 ÷ all-industry share", "Customer Need", None, X2, "index", "C", "Formula; all industries 8.633% (S15)", False),
    ("I7", "Capability-fit index (mean of pain-point fit, 1–3)", "", None, D2, "1–3", "C", "Pain-to-capability map below (05 §3)", False),
    ("I8", "Product-market fit score = 1 + 4 × (I7 − 1) ÷ 2", "Product-Market Fit", None, D2, "1–5", "C", "Formula", False),
    ("I9", "Usage-based first-year ACV per won account (Base usage)", "Revenue Potential", None, USD, "$", "C / D", "ACV block below (07 v2 method, published Growth prices S01)", False),
    ("I10", "Same, at Upside usage (1.5×) — stress test T10 only", "", None, USD, "$", "D", "ACV block below", False),
    ("I11", "Competitive position (mean of F1–F3)", "Competitive Position", None, D2, "1–5", "C", "04_Competitive_Analysis › Segment_Scoring", False),
    ("I12", "Adoption / execution fit (mean of F4–F5)", "Ease of Execution", None, D2, "1–5", "C", "04_Competitive_Analysis › Segment_Scoring", False),
]
r = 5
for rid, lab, crit, vals, fmt, unit, cls, src, judg in rows:
    IR[rid] = r
    put(inp, f"A{r}", rid, bold=True)
    put(inp, f"B{r}", lab, bold=bool(crit))
    put(inp, f"C{r}", crit)
    if vals:
        for j, v in enumerate(vals):
            put(inp, f"{L(4 + j)}{r}", v, font=BLUE, fmt=fmt, center=True, fill=YFILL if judg else None)
    put(inp, f"G{r}", unit, center=True)
    put(inp, f"H{r}", cls, center=True)
    put(inp, f"I{r}", src)
    r += 1
put(inp, f"A{r}", "All-industry sales occupations share", bold=True)
inp.merge_cells(f"A{r}:C{r}")
put(inp, f"D{r}", 13424300 / 155495730, font=BLUE, fmt='0.000%', center=True)
put(inp, f"I{r}", "BLS OEWS May 2025: 13,424,300 ÷ 155,495,730 (S15)")
ALLSALES = f"Inputs!$D${r}"
r += 2
for j in range(3):
    c = L(4 + j)
    put(inp, f"{c}{IR['I4']}", f"={c}{IR['I1']}*{c}{IR['I2']}*(1-{c}{IR['I3']})", fmt=N0, center=True, bold=True)
    put(inp, f"{c}{IR['I6']}", f"={c}{IR['I5']}/{ALLSALES}", fmt=X2, center=True, bold=True)
    put(inp, f"{c}{IR['I8']}", f"=1+4*({c}{IR['I7']}-1)/2", fmt=D2, center=True, bold=True)

# --- pain-to-capability map
PM0 = r
put(inp, f"A{r}", "Pain-to-capability map (05 §3). Fit: 3 = strong, 2 = partial, 1 = weak/gap; blank = not a core job for the segment. The builder-capacity gap is excluded here and scored under Ease of Execution.", bold=True, fill=SFILL)
inp.merge_cells(f"A{r}:I{r}")
inp.row_dimensions[r].height = 30
r += 1
hdr(inp, r, ["#", "Pain point / core job", "Clay capability", SEGS[0], SEGS[1], SEGS[2], "", "", "Gap"])
r += 1
PAINS = [
    ("Finding contractors and local accounts missing from B2B databases", "Google Maps, Shovels permits, waterfall, Claygent", [3, None, None], "Construction-project and licence data"),
    ("Prioritizing accounts by territory and branch", "Scoring, audiences, signals", [3, None, None], "Branch-level templates"),
    ("Keeping CRM complete when reps don't log", "CRM auto-sync, enrichment, job-change tracking", [3, 3, None], "Needs a CRM"),
    ("Vetting e-commerce and inbound registrations", "Webhook → AI verification → routing", [3, None, None], "Packaged template"),
    ("Cross-selling to existing accounts from order history", "Warehouse sync / HTTP API only", [1, None, None], "No ERP connector (Proton.ai, ERP CRMs own this)"),
    ("Construction project and bid timing", "Public bid signals, permits", [2, 2, None], "Private projects (Dodge) not in marketplace"),
    ("Capturing sourcing-platform demand (Thomasnet)", "—", [None, 1, None], "No API"),
    ("Researching prospects before pursuits / RFPs", "Claygent, account agents, pre-call prep", [None, None, 3], "—"),
    ("Targeted outbound for new clients", "Audiences, signals, sequencer", [None, None, 3], "Strong substitutes (Apollo, internal builds)"),
]
P1 = r
for i, (p, cap, fits, gap) in enumerate(PAINS):
    put(inp, f"A{r}", i + 1, center=True)
    put(inp, f"B{r}", p)
    put(inp, f"C{r}", cap)
    for j, v in enumerate(fits):
        put(inp, f"{L(4 + j)}{r}", v, font=BLUE, center=True, fill=YFILL if v is not None else None)
    put(inp, f"I{r}", gap)
    r += 1
P2 = r - 1
put(inp, f"B{r}", "Mean fit (I7) and number of core jobs mapped", bold=True)
for j in range(3):
    c = L(4 + j)
    put(inp, f"{c}{r}", f"=AVERAGE({c}{P1}:{c}{P2})", fmt=D2, center=True, bold=True)
    put(inp, f"{c}{IR['I7']}", f"={c}{r}", fmt=D2, center=True)
r += 1
put(inp, f"B{r}", "Core jobs mapped (small samples: read I8 with care)")
for j in range(3):
    c = L(4 + j)
    put(inp, f"{c}{r}", f"=COUNT({c}{P1}:{c}{P2})", center=True)
r += 2

# --- competitive and ease factors
put(inp, f"A{r}", "Competitive position and adoption factors (04 › Segment_Scoring; analyst judgments [C] tied to evidence there)", bold=True, fill=SFILL)
inp.merge_cells(f"A{r}:I{r}")
r += 1
FAC = [("F1", "Relevance of Clay's data differentiation (off-LinkedIn, multi-source)", [5, 3, 2], "I11"), ("F2", "Incumbent saturation (5 = little entrenched sales-intelligence use)", [4, 4, 2], "I11"),
       ("F3", "Threat from vertical specialists / systems of record (5 = low)", [3, 3, 3], "I11"), ("F4", "Builder capacity to self-serve Clay", [2, 1, 5], "I12"), ("F5", "Existing references and partner channel", [2, 2, 4], "I12")]
FR = {}
for fid, lab, vals, tgt in FAC:
    FR[fid] = r
    put(inp, f"A{r}", fid, bold=True)
    put(inp, f"B{r}", lab)
    put(inp, f"C{r}", "→ " + tgt)
    for j, v in enumerate(vals):
        put(inp, f"{L(4 + j)}{r}", v, font=BLUE, center=True, fill=YFILL)
    r += 1
for j in range(3):
    c = L(4 + j)
    put(inp, f"{c}{IR['I11']}", f"=AVERAGE({c}{FR['F1']}:{c}{FR['F3']})", fmt=D2, center=True, bold=True)
    put(inp, f"{c}{IR['I12']}", f"=AVERAGE({c}{FR['F4']}:{c}{FR['F5']})", fmt=D2, center=True, bold=True)
r += 1

# --- ACV block
put(inp, f"A{r}", "Usage-based ACV per won account (07 v2 method). Pricing and usage parameters from 07 › Assumptions (Base).", bold=True, fill=SFILL)
inp.merge_cells(f"A{r}:I{r}")
r += 1
P = ce.params("dist", 1)
PARS = [("dcprice", "Data-credit pack price ($/mo)"), ("dcsize", "Data-credit pack size (credits/yr)"), ("acprice", "Actions pack price ($/mo)"), ("acsize", "Actions pack size (actions/yr)"),
        ("aps", "Accounts tracked per seller"), ("cpa", "Contacts per account"), ("refresh", "Refreshes per year"), ("crc", "Credits per contact refresh"), ("ai", "AI credits per account"),
        ("newacc", "New accounts sourced per year (% of tracked)"), ("crnew", "Credits per new account"), ("apo", "Actions per row operation"), ("entcr", "Enterprise minimum credits / yr"),
        ("entac", "Enterprise minimum actions / yr"), ("entdisc", "Enterprise discount"), ("tiltM", "ICP tilt — Mid-Market"), ("tiltE", "ICP tilt — Enterprise"), ("propE", "Enterprise propensity")]
PR = {}
for k, lab in PARS:
    PR[k] = f"Inputs!$D${r}"
    put(inp, f"B{r}", lab)
    put(inp, f"D{r}", P[k], font=BLUE, center=True, fmt='0.00' if P[k] < 10 else N0)
    r += 1
put(inp, f"B{r}", "Usage intensity — Base / Upside")
put(inp, f"D{r}", 1.0, font=BLUE, center=True, fmt='0.00')
put(inp, f"E{r}", 1.5, font=BLUE, center=True, fmt='0.00')
USE_B, USE_U = f"Inputs!$D${r}", f"Inputs!$E${r}"
r += 2
CENSUS = {"4236": ([1408, 570, 350, 136, 114], [42698, 39297, 54183, 47492, 315093]), "4237": ([1013, 424, 242, 85, 57], [31260, 29270, 36995, 29008, 116895]),
          "4238": ([3622, 1218, 746, 240, 183], [110129, 83565, 115979, 83365, 220621]),
          "326": ([1505, 935, 760, 313, 277], [48064, 66254, 119875, 108128, 450163]), "332": ([7758, 3050, 1685, 423, 278], [238654, 210252, 253795, 147406, 354948]),
          "333": ([3291, 1484, 997, 358, 295], [102872, 103029, 151690, 123683, 513567]), "335": ([775, 422, 288, 108, 129], [24550, 29321, 44277, 37403, 220680]),
          "5413": ([5930, 2104, 1251, 429, 387], [179028, 145073, 189979, 149368, 727292]), "5415": ([5831, 2359, 1581, 561, 447], [179046, 164769, 240853, 195289, 1156449]),
          "5416": ([5789, 2151, 1244, 430, 339], [174222, 148901, 188821, 149492, 677207]), "5418": ([1655, 584, 371, 132, 102], [49648, 40398, 56587, 44231, 230304])}
SEGN = [["4236", "4237", "4238"], ["326", "332", "333", "335"], ["5413", "5415", "5416", "5418"]]
DEDUP = [[0.973, 1, 1, 1, 1], [0.970, 0.936, 0.900, 0.908, 0.913], [0.988, 0.979, 0.967, 0.951, 0.972]]
hdr(inp, r, ["NAICS", "Industry (Census 2022, S14)", "Segment", "Firms 20–49", "50–99", "100–249", "250–499", "500+", "Employees 20–49 · 50–99 · 100–249 · 250–499 · 500+ (cols J–N)"])
for j in range(5):
    inp.cell(row=r, column=10 + j, value=["Emp 20–49", "50–99", "100–249", "250–499", "500+"][j]).font = HDR
    inp.cell(row=r, column=10 + j).fill = HFILL
for c in "JKLMN":
    inp.column_dimensions[c].width = 11
r += 1
CR = {}
for s, ns in enumerate(SEGN):
    for n in ns:
        CR[n] = r
        put(inp, f"A{r}", n, bold=True)
        put(inp, f"B{r}", n)
        put(inp, f"C{r}", SHORT[s])
        for j, v in enumerate(CENSUS[n][0]):
            put(inp, f"{L(4 + j)}{r}", v, font=BLUE, fmt=N0, center=True)
        for j, v in enumerate(CENSUS[n][1]):
            put(inp, f"{L(10 + j)}{r}", v, font=BLUE, fmt=N0, center=True)
        r += 1
DR = {}
for s in range(3):
    DR[s] = r
    put(inp, f"B{r}", f"De-duplication factor by band — {SHORT[s]}")
    for j, v in enumerate(DEDUP[s]):
        put(inp, f"{L(4 + j)}{r}", v, font=BLUE, fmt='0.000', center=True)
    r += 1
put(inp, f"B{r}", "Manufacturing factors: sector 31-33 ÷ Σ subsectors (overlap_raw.csv); 250+ bands omit suppressed subsectors", italic=True, border=False)
r += 2
hdr(inp, r, ["Band", "Line", "", SEGS[0], SEGS[1], SEGS[2], "", "", "Formula"])
r += 1
AB = {}
BANDS = [("S", (0, 1), "SMB 20–99"), ("M", (2, 3), "Mid-Market 100–499"), ("E", (4,), "Enterprise 500+")]


def cp():
    return f"({PR['dcprice']}*12/{PR['dcsize']})"


def ap():
    return f"({PR['acprice']}*12/{PR['acsize']})"


for b, idx, blab in BANDS:
    for line in ["firms", "emp", "rawfirms", "sellers", "acvB", "acvU", "w"]:
        AB[(b, line)] = r
        labmap = {"firms": "Firms (de-duplicated)", "emp": "Employees", "rawfirms": "Firms (raw, for averages)", "sellers": "Sellers per firm = employees ÷ firms × sales share",
                  "acvB": "First-year ACV at Base usage", "acvU": "First-year ACV at Upside usage", "w": "Opportunity weight = firms × tilt × propensity"}
        put(inp, f"A{r}", blab if line == "firms" else "")
        put(inp, f"B{r}", labmap[line])
        for s in range(3):
            c = L(4 + s)
            ns = SEGN[s]
            if line == "firms":
                f = "=" + "+".join(f"{L(4 + i)}{CR[n]}*{L(4 + i)}${DR[s]}" for n in ns for i in idx)
                fmt = N0
            elif line == "emp":
                f = "=" + "+".join(f"{L(10 + i)}{CR[n]}" for n in ns for i in idx)
                fmt = N0
            elif line == "rawfirms":
                f = "=" + "+".join(f"{L(4 + i)}{CR[n]}" for n in ns for i in idx)
                fmt = N0
            elif line == "sellers":
                f = f"={c}{AB[(b, 'emp')]}/{c}{AB[(b, 'rawfirms')]}*{c}${IR['I5']}"
                fmt = '0.0'
            elif line in ("acvB", "acvU"):
                use = USE_B if line == "acvB" else USE_U
                acc = f"({c}{AB[(b, 'sellers')]}*{PR['aps']}*{use})"
                con = f"({acc}*{PR['cpa']})"
                cred = f"({con}*{PR['refresh']}*{PR['crc']}+{acc}*{PR['ai']}+{acc}*{PR['newacc']}*{PR['crnew']})"
                ops = f"({con}*{PR['refresh']}+{acc}+{acc}*{PR['newacc']})"
                usage = f"({cred}*{cp()}+{ops}*{PR['apo']}*{ap()})"
                if b == "E":
                    f = f"=MAX({PR['entcr']}*{cp()}+{PR['entac']}*{ap()},{usage})*(1-{PR['entdisc']})"
                else:
                    f = f"=MAX(({PR['dcprice']}+{PR['acprice']})*12,{usage})"
                fmt = USD
            else:
                mult = "1" if b == "S" else (PR["tiltM"] if b == "M" else f"{PR['tiltE']}*{PR['propE']}")
                f = f"={c}{AB[(b, 'firms')]}*{mult}"
                fmt = N0
            put(inp, f"{c}{r}", f, fmt=fmt, center=True)
        r += 1
for line, rid in (("acvB", "I9"), ("acvU", "I10")):
    put(inp, f"B{r}", f"Blended ACV per won account ({'Base' if line == 'acvB' else 'Upside'} usage) = Σ weight × ACV ÷ Σ weight", bold=True)
    for s in range(3):
        c = L(4 + s)
        num = "+".join(f"{c}{AB[(b, 'w')]}*{c}{AB[(b, line)]}" for b, _, _ in BANDS)
        den = "+".join(f"{c}{AB[(b, 'w')]}" for b, _, _ in BANDS)
        put(inp, f"{c}{r}", f"=({num})/({den})", fmt=USD, center=True, bold=True, fill=KFILL)
        put(inp, f"{c}{IR[rid]}", f"={c}{r}", fmt=USD, center=True, bold=True)
    r += 1
note(inp, f"A{r + 1}:I{r + 1}", "Distribution ACV uses 4236–4238 (4235 has suppressed Census cells); the account count I1 uses Phase 1's full four-industry segment. Manufacturing and professional-services SMB and Mid-Market ACVs sit at the Growth floor because they have few sellers per firm.", 40)
METRIC_REF = {0: IR["I4"], 1: IR["I6"], 2: IR["I8"], 3: IR["I9"], 4: IR["I11"], 5: IR["I12"]}


# ======================================================================= SCORING ENGINE
def block(ws, top, label, over=None, rule=None, weights=W_BASE, desc=""):
    """Metric rows, score rows, total, winner. over: {crit_index: [formula|value for A,B,C]}.
    rule: {crit_index: 'ratio'|'minmax'|'sqrt'} for quantitative criteria (default ratio)."""
    over = over or {}
    rule = rule or {}
    put(ws, f"A{top}", label, bold=True, fill=SFILL)
    for c in "BCDEF":
        put(ws, f"{c}{top}", None, fill=SFILL)
    if desc:
        put(ws, f"F{top}", desc, fill=SFILL, italic=True)
    r0 = top + 1
    hdr(ws, r0, ["Criterion", "Line", SHORT[0], SHORT[1], SHORT[2], "Scoring"], height=20)
    mrow, srow = {}, {}
    rr = r0 + 1
    for i, cname in enumerate(CRIT):
        mrow[i] = rr
        put(ws, f"A{rr}", cname)
        put(ws, f"B{rr}", "metric")
        for j in range(3):
            v = over.get(i, [None] * 3)[j] if i in over else None
            if v is None:
                v = f"=Inputs!{L(4 + j)}{METRIC_REF[i]}"
                fnt = GREEN
            else:
                fnt = BLUE if not str(v).startswith("=") else BLACK
            put(ws, f"{L(3 + j)}{rr}", v, font=fnt, fmt=USD if i == 3 else (N0 if i == 0 else D2), center=True, fill=YFILL if i in over else None)
        rr += 1
    for i, cname in enumerate(CRIT):
        srow[i] = rr
        put(ws, f"A{rr}", cname, bold=True)
        put(ws, f"B{rr}", "score 1–5")
        m = mrow[i]
        rg = f"$C${m}:$E${m}"
        kind = rule.get(i, "ratio") if QUANT[i] else "direct"
        for j in range(3):
            c = L(3 + j)
            if kind == "ratio":
                f = f"=1+4*{c}{m}/MAX({rg})"
            elif kind == "minmax":
                f = f"=1+4*({c}{m}-MIN({rg}))/(MAX({rg})-MIN({rg}))"
            elif kind == "sqrt":
                f = f"=1+4*SQRT({c}{m}/MAX({rg}))"
            else:
                f = f"={c}{m}"
            put(ws, f"{c}{rr}", f, fmt=D2, center=True)
        put(ws, f"F{rr}", {"ratio": "1 + 4 × v ÷ max", "minmax": "1 + 4 × (v − min) ÷ (max − min)", "sqrt": "1 + 4 × √(v ÷ max)", "direct": "already 1–5"}[kind])
        rr += 1
    tot = rr
    put(ws, f"A{tot}", "Weighted total", bold=True, fill=KFILL)
    put(ws, f"B{tot}", "Σ weight × score", fill=KFILL)
    for j in range(3):
        c = L(3 + j)
        put(ws, f"{c}{tot}", f"=SUMPRODUCT({weights},{c}{srow[0]}:{c}{srow[5]})", fmt=D2, center=True, bold=True, fill=KFILL)
    win = tot + 1
    put(ws, f"A{win}", "Winner / margin of distribution over professional services", bold=True)
    put(ws, f"C{win}", f'=INDEX({{"{SHORT[0]}","{SHORT[1]}","{SHORT[2]}"}},MATCH(MAX(C{tot}:E{tot}),C{tot}:E{tot},0))', center=True, bold=True)
    put(ws, f"E{win}", f"=C{tot}-E{tot}", fmt='+0.00;-0.00', center=True, bold=True)
    return dict(m=mrow, s=srow, tot=tot, win=win, end=win)


# ======================================================================= SCORECARD
title(sc, "Scorecard v2 — one indicator per criterion, Phase 4 economics fed back", "Base weights from the brief. Quantitative indicators scale 1 + 4 × v ÷ max; judgment indicators are already 1–5.")
for c, w in zip("ABCDEF", [26, 12, 15, 15, 15, 30]):
    sc.column_dimensions[c].width = w
SB = block(sc, 4, "Base case (brief weights, ratio-to-max scaling)")
TOT = SB["tot"]
SCORE = {i: SB["s"][i] for i in range(6)}
r = SB["end"] + 2
put(sc, f"A{r}", "Contribution to total (weight × score)", bold=True, fill=SFILL)
for c in "BCDEF":
    put(sc, f"{c}{r}", None, fill=SFILL)
r += 1
CONT0 = r
for i, cname in enumerate(CRIT):
    put(sc, f"A{r}", cname)
    put(sc, f"B{r}", f"=Weights!$B${5 + i}", fmt='0%', center=True)
    for j in range(3):
        c = L(3 + j)
        put(sc, f"{c}{r}", f"=$B{r}*{c}{SCORE[i]}", fmt=D2, center=True)
    r += 1
CONT1 = r - 1
ch = BarChart()
ch.type = "bar"
ch.grouping = "stacked"
ch.overlap = 100
ch.title = "Weighted contribution by criterion (v2)"
for i in range(6):
    from openpyxl.chart import Series
    sref = Reference(sc, min_col=3, max_col=5, min_row=CONT0 + i)
    s_ = Series(sref, title=CRIT[i])
    ch.series.append(s_)
ch.set_categories(Reference(sc, min_col=3, max_col=5, min_row=SB["m"][0] - 1))
ch.height, ch.width = 9, 18
ch.x_axis.delete = False
ch.y_axis.delete = False
sc.add_chart(ch, f"H4")
r += 1
EVID = [
    ("Market Attractiveness", "Reachable accounts after the CRM screen and excluding existing Clay customers: ~6,200 distributors vs ~13,800 manufacturers and ~20,800 professional-services firms. Scaled linearly (T11 tests concave scaling)."),
    ("Customer Need", "Sales occupations are 22.0% of distribution employment, 2.5× the economy (S15). The coverage proxy no longer enters here (moved to Competitive Position only)."),
    ("Product-Market Fit", "Mean fit across each segment's mapped core jobs (05 §3). Professional services' research and outbound jobs are Clay's horizontal sweet spot (fit 3.0); distribution 2.5 (ERP cross-sell gap); manufacturing 2.0."),
    ("Revenue Potential", "Usage-based ACV per won account at list prices (Phase 4): distribution ~$9.0K; manufacturing and professional services ~$7.1K because their SMB and Mid-Market ACVs sit at the Growth floor."),
    ("Competitive Position", "Unchanged from 04: off-LinkedIn data relevance, incumbent saturation, specialist threat. The only criterion that uses the LinkedIn-coverage proxy."),
    ("Ease of Execution", "Unchanged from 04: builder capacity and references/partners. Professional services 4.5 vs distribution 2.0."),
]
put(sc, f"A{r}", "Evidence behind each criterion", bold=True, fill=SFILL)
for c in "BCDEF":
    put(sc, f"{c}{r}", None, fill=SFILL)
r += 1
for cname, txt in EVID:
    put(sc, f"A{r}", cname, bold=True)
    put(sc, f"B{r}", txt)
    sc.merge_cells(f"B{r}:F{r}")
    sc.row_dimensions[r].height = 40
    r += 1

# ======================================================================= BRIDGE v1 -> v2
title(brg, "Bridge v1 → v2 — how each definitional change moves the totals", "v1 criterion scores are typed from the archived v1 workbook; v2 scores link to the Scorecard. Each step swaps one criterion from its v1 to its v2 definition.")
hdr(brg, 4, ["Step", "What changed", SHORT[0], SHORT[1], SHORT[2], "Distribution − Prof. services"], widths=[34, 70, 15, 15, 15, 18])
V1 = [[3.91971322815144, 4.02332795210303, 3.5], [5, 2.52981955268848, 2.80458193586941], [4, 2, 4], [3.89224941319643, 2.29227181029663, 3.67292418772563],
      [4, 3.3333, 2.3333], [2, 1.5, 4.5]]
put(brg, "A20", "v1 criterion scores (typed)", bold=True)
for i in range(6):
    put(brg, f"A{21 + i}", CRIT[i])
    for j in range(3):
        put(brg, f"{L(3 + j)}{21 + i}", V1[i][j], font=BLUE, fmt=D2, center=True)
V1R = {i: 21 + i for i in range(6)}
STEPS = [("v1 total (as published)", "Market = SAM $ + headroom flag; Need = sales intensity + coverage gap; PMF = judgment 4/2/4; Revenue = SOM + seller density", []),
         ("1. Market attractiveness", "SAM $ (uniform Phase 1 ACV) and a separate headroom flag → reachable accounts net of existing Clay customers", [0]),
         ("2. Customer need", "Drop the coverage-gap half (it already drives Competitive Position F1) → sales intensity only", [0, 1]),
         ("3. Product-market fit", "Free judgment 4/2/4 → capability-fit index from the 05 pain map", [0, 1, 2]),
         ("4. Revenue potential", "SOM + seller density (assumed ACV lift) → usage-based ACV per won account (Phase 4)", [0, 1, 2, 3]),
         ("v2 total", "Competitive position and ease unchanged", [0, 1, 2, 3])]
for k, (lab, what, swapped) in enumerate(STEPS):
    rr = 5 + k
    put(brg, f"A{rr}", lab, bold=(k in (0, len(STEPS) - 1)))
    put(brg, f"B{rr}", what)
    for j in range(3):
        c = L(3 + j)
        terms = []
        for i in range(6):
            src = f"Scorecard!{c}{SCORE[i]}" if i in swapped else f"{c}{V1R[i]}"
            terms.append(f"Weights!$B${5 + i}*{src}")
        put(brg, f"{c}{rr}", "=" + "+".join(terms), fmt=D2, center=True, bold=(k in (0, len(STEPS) - 1)))
    put(brg, f"F{rr}", f"=C{rr}-E{rr}", fmt='+0.00;-0.00', center=True, bold=True)
    brg.row_dimensions[rr].height = 32
put(brg, "A12", "Check: v2 total equals Scorecard (should be 0)")
put(brg, "C12", f"=ROUND(C10-Scorecard!C{TOT},6)+ROUND(E10-Scorecard!E{TOT},6)", fmt=D2, center=True)
note(brg, "A14:F14", "Reading: the largest swing is market attractiveness — counting reachable accounts instead of SAM dollars (which carried a uniform ACV and an outbound filter rationalized with the same sales share) moves ~0.3 points from distribution to professional services. Revenue potential moves the other way once Phase 4's usage-based ACV replaces the assumed seller-density lift.", 46)

# ======================================================================= WEIGHT SENSITIVITY
title(wsn, "Weight sensitivity — named weight sets and break-even weights", "Block A: totals under six named weight sets. Block B: the weight on each criterion at which distribution and professional services tie (others rescaled proportionally).")
hdr(wsn, 4, ["Weight set", SHORT[0], SHORT[1], SHORT[2], "Winner", "Dist − PS"], widths=[28, 15, 15, 15, 18, 14])
for k, (nm, _) in enumerate(WSETS):
    rr = 5 + k
    wcol = L(2 + k)
    put(wsn, f"A{rr}", nm, bold=(k == 0))
    for j in range(3):
        c = L(3 + j)
        put(wsn, f"{L(2 + j)}{rr}", f"=SUMPRODUCT(Weights!${wcol}$5:${wcol}$10,Scorecard!{c}{SCORE[0]}:{c}{SCORE[5]})", fmt=D2, center=True)
    put(wsn, f"E{rr}", f'=INDEX({{"{SHORT[0]}","{SHORT[1]}","{SHORT[2]}"}},MATCH(MAX(B{rr}:D{rr}),B{rr}:D{rr},0))', center=True)
    put(wsn, f"F{rr}", f"=B{rr}-D{rr}", fmt='+0.00;-0.00', center=True, bold=True)
WS_END = 5 + len(WSETS) - 1
put(wsn, f"A{WS_END + 1}", "Named sets won by distribution", bold=True)
put(wsn, f"B{WS_END + 1}", f'=COUNTIF(E5:E{WS_END},"{SHORT[0]}")', center=True, bold=True)
put(wsn, f"A{WS_END + 2}", "Named sets won by professional services", bold=True)
put(wsn, f"B{WS_END + 2}", f'=COUNTIF(E5:E{WS_END},"{SHORT[2]}")', center=True, bold=True)
bb = WS_END + 5
put(wsn, f"A{bb}", "Block B — break-even weight: distribution = professional services", bold=True, border=False)
hdr(wsn, bb + 1, ["Criterion", "Base weight", "Score gap (Dist − PS)", "Rest-of-card gap", "Break-even weight", "Reading"])
for i in range(6):
    rr = bb + 2 + i
    put(wsn, f"A{rr}", CRIT[i])
    put(wsn, f"B{rr}", f"=Weights!$B${5 + i}", fmt='0%', center=True)
    put(wsn, f"C{rr}", f"=Scorecard!C{SCORE[i]}-Scorecard!E{SCORE[i]}", fmt='+0.00;-0.00', center=True)
    put(wsn, f"D{rr}", f"=((Scorecard!C{TOT}-Scorecard!E{TOT})-B{rr}*C{rr})/(1-B{rr})", fmt='+0.00;-0.00', center=True)
    put(wsn, f"E{rr}", f'=IFERROR(IF(AND(D{rr}/(D{rr}-C{rr})>=0,D{rr}/(D{rr}-C{rr})<=1),D{rr}/(D{rr}-C{rr}),"none in 0–100%"),"none")', fmt='0.0%', center=True, bold=True)
    put(wsn, f"F{rr}", f'=IF(ISNUMBER(E{rr}),IF(C{rr}>0,"Distribution leads above this weight","Prof. services leads above this weight"),"")')
BE0 = bb + 2

# ======================================================================= RANDOM WEIGHTS
title(rw, "Random weights — 1,000 flat and 1,000 centered Dirichlet draws (numpy, seed 42; stored values), totals are live formulas",
      "Flat: any weighting equally likely. Centered: Dirichlet(30 × base weights), reasonable disagreement around the brief.")
rng = np.random.default_rng(42)
flat = rng.dirichlet([1] * 6, 1000)
cent = rng.dirichlet([30 * w for w in WSETS[0][1]], 1000)
hdr(rw, 4, [c[:12] for c in CRIT] + SHORT + ["Winner"], widths=[11] * 6 + [12, 12, 12, 14])
hdr(rw, 4, [c[:12] for c in CRIT] + SHORT + ["Winner"], col=12, widths=[11] * 6 + [12, 12, 12, 14])
put(rw, "A3", "Flat Dirichlet(1, …, 1)", bold=True, border=False)
put(rw, "L3", "Centered Dirichlet(30 × base weights)", bold=True, border=False)
for blockc, arr in ((1, flat), (12, cent)):
    for k in range(1000):
        rr = 5 + k
        for i in range(6):
            rw.cell(row=rr, column=blockc + i, value=round(float(arr[k][i]), 6)).number_format = '0.000'
        for j in range(3):
            c = L(3 + j)
            terms = "+".join(f"{L(blockc + i)}{rr}*Scorecard!${c}${SCORE[i]}" for i in range(6))
            rw.cell(row=rr, column=blockc + 6 + j, value="=" + terms).number_format = D2
        a, b_, c_ = (L(blockc + 6 + j) for j in range(3))
        rw.cell(row=rr, column=blockc + 9, value=f'=IF(AND({a}{rr}>={b_}{rr},{a}{rr}>={c_}{rr}),"{SHORT[0]}",IF({c_}{rr}>={b_}{rr},"{SHORT[2]}","{SHORT[1]}"))')
RW_END = 1004
put(wsn, f"A{BE0 + 8}", "Block C — random-weight simulation", bold=True, border=False)
hdr(wsn, BE0 + 9, ["Simulation", SHORT[0], SHORT[1], SHORT[2], "", ""])
for k, (lab, col) in enumerate((("Flat Dirichlet — share of 1,000 draws won", "J"), ("Centered Dirichlet — share of 1,000 draws won", "U"))):
    rr = BE0 + 10 + k
    put(wsn, f"A{rr}", lab)
    for j in range(3):
        put(wsn, f"{L(2 + j)}{rr}", f'=COUNTIF(Random_Weights!${col}$5:${col}${RW_END},"{SHORT[j]}")/1000', fmt='0.0%', center=True, bold=True)
RWS = BE0 + 10

# ======================================================================= STRESS TESTS
title(stt, "Stress tests — what has to be true for each segment to win", "Each block changes only the yellow inputs. Summary first; detail blocks below.")
for c, w in zip("ABCDEFG", [44, 50, 14, 14, 14, 18, 14]):
    stt.column_dimensions[c].width = w
TESTS = [
    ("T0 Base case", "Nothing (reproduces Scorecard)", {}, {}),
    ("T1 Partner channel works", "Distribution ease 2.0 → 3.5 (templates and partners close the builder gap)", {5: ["=3.5", None, None]}, {}),
    ("T2 More PS headroom", "Prof. services already-Clay share 10% → 5%", {0: [None, None, f"=Inputs!F{IR['I1']}*Inputs!F{IR['I2']}*(1-0.05)"]}, {}),
    ("T3 Less PS headroom", "Prof. services already-Clay share 10% → 25%", {0: [None, None, f"=Inputs!F{IR['I1']}*Inputs!F{IR['I2']}*(1-0.25)"]}, {}),
    ("T4 Coverage proxy wrong", "Distribution F1 data differentiation 5 → 3 (competitive 4.0 → 3.33)", {4: ["=10/3", None, None]}, {}),
    ("T5 ERP gap binds", "Distribution product-market fit 4.0 → 3.0", {2: ["=3", None, None]}, {}),
    ("T6 Apollo/HubSpot close the local-data gap", "Distribution competitive position 4.0 → 3.0", {4: ["=3", None, None]}, {}),
    ("T7 Distribution combined downside", "T5 + T6", {2: ["=3", None, None], 4: ["=3", None, None]}, {}),
    ("T8 Prof. services best case", "PS competitive 2.33 → 3.0; ease 4.5 → 5.0", {4: [None, None, "=3"], 5: [None, None, "=5"]}, {}),
    ("T9 Worst case for A + best case for C", "T7 + T8", {2: ["=3", None, None], 4: ["=3", None, "=3"], 5: [None, None, "=5"]}, {}),
    ("T10 Upside usage", "ACV at 1.5× usage for every segment", {3: [f"=Inputs!D{IR['I10']}", f"=Inputs!E{IR['I10']}", f"=Inputs!F{IR['I10']}"]}, {}),
    ("T11 Concave size scoring", "Market attractiveness scored 1 + 4 × √(v ÷ max): a pilot can serve only a few hundred accounts", {}, {0: "sqrt"}),
    ("T12 Min-max scaling", "All quantitative indicators scaled min-max instead of ratio-to-max", {}, {0: "minmax", 1: "minmax", 3: "minmax"}),
    ("T13 v1 market definition", "Reachable accounts × v1 outbound-motion filter (80% / 45% / 55%)",
     {0: [f"=Inputs!D{IR['I4']}*0.8", f"=Inputs!E{IR['I4']}*0.45", f"=Inputs!F{IR['I4']}*0.55"]}, {}),
]
hdr(stt, 4, ["Test", "What changes", SHORT[0], SHORT[1], SHORT[2], "Winner", "Dist − PS"])
TB = {}
top = 5 + len(TESTS) + 3
for k, (nm, what, over, rule) in enumerate(TESTS):
    b = block(stt, top, nm, over=over, rule=rule, desc=what)
    TB[nm] = b
    rr = 5 + k
    put(stt, f"A{rr}", nm, bold=(k == 0))
    put(stt, f"B{rr}", what)
    for j in range(3):
        put(stt, f"{L(3 + j)}{rr}", f"={L(3 + j)}{b['tot']}", fmt=D2, center=True)
    put(stt, f"F{rr}", f"=C{b['win']}", center=True)
    put(stt, f"G{rr}", f"=C{rr}-E{rr}", fmt='+0.00;-0.00', center=True, bold=True)
    top = b["end"] + 2
T_END = 5 + len(TESTS) - 1
stt.conditional_formatting.add(f"G5:G{T_END}", CellIsRule(operator="greaterThan", formula=["0"], fill=KFILL))
stt.conditional_formatting.add(f"G5:G{T_END}", CellIsRule(operator="lessThan", formula=["0"], fill=RFILL))
put(stt, f"A{T_END + 1}", "Tests won by distribution / professional services", bold=True)
put(stt, f"C{T_END + 1}", f'=COUNTIF(F5:F{T_END},"{SHORT[0]}")', center=True, bold=True)
put(stt, f"E{T_END + 1}", f'=COUNTIF(F5:F{T_END},"{SHORT[2]}")', center=True, bold=True)
STT_SUM = T_END + 1

# ======================================================================= ECONOMIC CROSS-CHECK (from 07 v2)
title(eco, "Economic cross-check — the same two segments on the 07 v2 engine", "Values copied from 07_Financial_Model.xlsx (v2, Scenarios and Monte Carlo sheets) at build time; they are not part of the weighted score, to avoid double counting revenue potential.")
m07 = load_workbook(f"{BASE}/07_Financial_Model.xlsx", data_only=True)
s07, mc07 = m07["Scenarios"], m07["Monte Carlo"]
hdr7 = [s07.cell(4, c).value for c in range(1, s07.max_column + 1)]


def s07v(label_row_key, col_label):
    row = next(rr for rr in range(5, 30) if s07.cell(rr, 1).value == label_row_key)
    col = hdr7.index(col_label) + 1
    return s07.cell(row, col).value


def mcv(label, col):
    row = next(rr for rr in range(5, 20) if mc07.cell(rr, 1).value == label)
    return mc07.cell(row, col).value


hdr(eco, 4, ["Metric (07 v2)", "Distribution", "Prof. services", "Better", "Source"], widths=[52, 16, 16, 16, 50])
ECO = [
    ("Lean pilot — Base NPV at 12%", s07v("NPV of the program at hurdle rate", "Dist · Lean partner-led pilot · Base"), s07v("NPV of the program at hurdle rate", "PS · Lean partner-led pilot · Base"), USD, "07 › Scenarios"),
    ("Lean pilot — Upside NPV at 12%", s07v("NPV of the program at hurdle rate", "Dist · Lean partner-led pilot · Upside"), s07v("NPV of the program at hurdle rate", "PS · Lean partner-led pilot · Upside"), USD, "07 › Scenarios"),
    ("Solution play — Base NPV", s07v("NPV of the program at hurdle rate", "Dist · Solution play · Base"), s07v("NPV of the program at hurdle rate", "PS · Solution play · Base"), USD, "07 › Scenarios"),
    ("Solution play — Upside NPV advantage over Option 0", s07v("NPV advantage over Option 0", "Dist · Solution play · Upside"), s07v("NPV advantage over Option 0", "PS · Solution play · Upside"), USD, "07 › Scenarios"),
    ("Lean pilot, staged — simulated mean NPV", mcv("Industrial distribution — Lean pilot — staged: Gate 1 with observation noise", 2), mcv("Professional services — Lean pilot — staged: Gate 1 with observation noise", 2), USD, "07 › Monte Carlo"),
    ("Lean pilot — Base average first-year ACV", s07v("Average first-year ACV", "Dist · Lean partner-led pilot · Base"), s07v("Average first-year ACV", "PS · Lean partner-led pilot · Base"), USD, "07 › Scenarios"),
]
for k, (lab, a, c_, fmt, srcx) in enumerate(ECO):
    rr = 5 + k
    put(eco, f"A{rr}", lab)
    put(eco, f"B{rr}", round(a), font=BLUE, fmt=fmt, center=True)
    put(eco, f"C{rr}", round(c_), font=BLUE, fmt=fmt, center=True)
    put(eco, f"D{rr}", f'=IF(B{rr}>C{rr},"{SHORT[0]}","{SHORT[2]}")', center=True, bold=True)
    put(eco, f"E{rr}", srcx)
ECO_END = 5 + len(ECO) - 1
put(eco, f"A{ECO_END + 1}", "Metrics on which distribution is better", bold=True)
put(eco, f"B{ECO_END + 1}", f'=COUNTIF(D5:D{ECO_END},"{SHORT[0]}")&" of {len(ECO)}"', center=True, bold=True)
note(eco, f"A{ECO_END + 3}:E{ECO_END + 3}", "Reading: on the full economics distribution is modestly better than professional services in every comparison — mainly because more of a professional-services program's wins would have come through Clay's horizontal motion anyway, and its list-price ACV sits at the Growth floor. Both are NPV-negative in Base.", 46)
ECO_VALS = [(e[1], e[2]) for e in ECO]

# ======================================================================= SUBSEGMENT DRILLDOWN
title(dd, "Sub-segment drilldown — which distribution industry first? (v2 indicators)", "Only the criteria that differ by sub-industry are scored (market, need, revenue; base weights renormalized). Product fit, competitive position and ease are assumed equal across the three. Reference customers are a tie-breaker, not a score.")
for c, w in zip("ABCDEFG", [46, 16, 16, 16, 14, 14, 50]):
    dd.column_dimensions[c].width = w
hdr(dd, 4, ["Indicator", "4236 Electrical", "4237 Plumbing / HVAC", "4238 Machinery / MRO", "Weight", "", "Source"])
SUBN = ["4236", "4237", "4238"]
SALES_SUB = [322290 / 1503920, 322290 / 1503920, 178040 / 774540]
rr = 5
put(dd, f"A{rr}", "Firms 20+ (de-duplicated, 20–49 × 0.973)")
for j, n in enumerate(SUBN):
    rowc = CR[n]
    put(dd, f"{L(2 + j)}{rr}", f"=Inputs!D{rowc}*0.973+SUM(Inputs!E{rowc}:H{rowc})", fmt=N0, center=True)
put(dd, f"G{rr}", "Census 2022 (S14)")
DDR = {"firms": rr}
rr += 1
put(dd, f"A{rr}", "Reachable accounts = firms × CRM in use × (1 − already Clay)")
for j in range(3):
    c = L(2 + j)
    put(dd, f"{c}{rr}", f"={c}{DDR['firms']}*Inputs!$D${IR['I2']}*(1-Inputs!$D${IR['I3']})", fmt=N0, center=True)
put(dd, f"E{rr}", f"=Weights!B5/(Weights!B5+Weights!B6+Weights!B8)", fmt='0%', center=True)
DDR["acc"] = rr
rr += 1
put(dd, f"A{rr}", "Sales occupations share of employment")
for j in range(3):
    put(dd, f"{L(2 + j)}{rr}", SALES_SUB[j], font=BLUE, fmt=PCT, center=True)
put(dd, f"E{rr}", f"=Weights!B6/(Weights!B5+Weights!B6+Weights!B8)", fmt='0%', center=True)
put(dd, f"G{rr}", "OEWS May 2025 (S15): 4236/4237 share the 4230A1 group (durable goods ex-machinery, 21.4%); 4238 23.0%. Cannot separate 4236 from 4237.")
DDR["sales"] = rr
rr += 1
put(dd, f"A{rr}", "Usage-based ACV per won account (Base)")
DDR["acv"] = rr
put(dd, f"E{rr}", f"=Weights!B8/(Weights!B5+Weights!B6+Weights!B8)", fmt='0%', center=True)
put(dd, f"G{rr}", "Same method as Inputs; each sub-industry's own employees per firm")
rr += 1
# ACV per sub-industry (inline formulas)
for j, n in enumerate(SUBN):
    rowc = CR[n]
    terms_w, terms_n = [], []
    for b, idx, _ in BANDS:
        firms = "+".join((f"Inputs!{L(4 + i)}{rowc}*0.973" if i == 0 else f"Inputs!{L(4 + i)}{rowc}") for i in idx)
        rawf = "+".join(f"Inputs!{L(4 + i)}{rowc}" for i in idx)
        emp = "+".join(f"Inputs!{L(10 + i)}{rowc}" for i in idx)
        sellers = f"(({emp})/({rawf})*{L(2 + j)}{DDR['sales']})"
        acc = f"({sellers}*{PR['aps']}*{USE_B})"
        con = f"({acc}*{PR['cpa']})"
        cred = f"({con}*{PR['refresh']}*{PR['crc']}+{acc}*{PR['ai']}+{acc}*{PR['newacc']}*{PR['crnew']})"
        ops = f"({con}*{PR['refresh']}+{acc}+{acc}*{PR['newacc']})"
        usage = f"({cred}*{cp()}+{ops}*{PR['apo']}*{ap()})"
        acv = f"MAX({PR['entcr']}*{cp()}+{PR['entac']}*{ap()},{usage})*(1-{PR['entdisc']})" if b == "E" else f"MAX(({PR['dcprice']}+{PR['acprice']})*12,{usage})"
        mult = "1" if b == "S" else (PR["tiltM"] if b == "M" else f"{PR['tiltE']}*{PR['propE']}")
        terms_w.append(f"({firms})*{mult}")
        terms_n.append(f"({firms})*{mult}*{acv}")
    put(dd, f"{L(2 + j)}{DDR['acv']}", f"=({'+'.join(terms_n)})/({'+'.join(terms_w)})", fmt=USD, center=True)
rr += 1
put(dd, f"A{rr}", "Scores (ratio-to-max)", bold=True, fill=SFILL)
for c in "BCDEFG":
    put(dd, f"{c}{rr}", None, fill=SFILL)
rr += 1
SR = {}
for key, lab in (("acc", "Market (reachable accounts)"), ("sales", "Need (sales share)"), ("acv", "Revenue (ACV per account)")):
    SR[key] = rr
    put(dd, f"A{rr}", lab)
    for j in range(3):
        c = L(2 + j)
        put(dd, f"{c}{rr}", f"=1+4*{c}{DDR[key]}/MAX($B${DDR[key]}:$D${DDR[key]})", fmt=D2, center=True)
    rr += 1
put(dd, f"A{rr}", "Weighted score (market, need, revenue)", bold=True, fill=KFILL)
for j in range(3):
    c = L(2 + j)
    put(dd, f"{c}{rr}", f"=$E${DDR['acc']}*{c}{SR['acc']}+$E${DDR['sales']}*{c}{SR['sales']}+$E${DDR['acv']}*{c}{SR['acv']}", fmt=D2, center=True, bold=True, fill=KFILL)
DDT = rr
rr += 1
put(dd, f"A{rr}", "Same, with concave market scoring (√)", bold=True)
for j in range(3):
    c = L(2 + j)
    put(dd, f"{c}{rr}", f"=$E${DDR['acc']}*(1+4*SQRT({c}{DDR['acc']}/MAX($B${DDR['acc']}:$D${DDR['acc']})))+$E${DDR['sales']}*{c}{SR['sales']}+$E${DDR['acv']}*{c}{SR['acv']}", fmt=D2, center=True, bold=True)
DDT2 = rr
rr += 2
put(dd, f"A{rr}", "Not scored — tie-breakers", bold=True, fill=SFILL)
for c in "BCDEFG":
    put(dd, f"{c}{rr}", None, fill=SFILL)
rr += 1
for lab, vals in (("Public Clay reference customer", ["Regency Supply (S07)", "None found", "None found"]),
                  ("Contractor-facing (permits / Maps sourcing relevance)", ["High (electrical contractors)", "Highest (plumbing, HVAC contractors)", "Mixed (plants, MRO buyers)"]),
                  ("Establishments per firm, full-year firms (territory complexity)", ["1.64", "2.21", "1.57"]),
                  ("Industry association channel", ["NAED (S53)", "ASA, 4,000+ branch locations (S54)", "ISA, PTDA [K]"])):
    put(dd, f"A{rr}", lab)
    for j, v in enumerate(vals):
        put(dd, f"{L(2 + j)}{rr}", v, center=True)
    rr += 1
note(dd, f"A{rr + 1}:G{rr + 1}", "Reading: on v2 indicators, 4238 machinery/MRO scores highest — the largest account base and the highest sales share, with ACV close to 4236's. 4236 electrical remains a defensible first wave only because it has the one public Clay reference (Regency) and a single association channel. A validation sprint should test both 4236 and 4238 rather than assume the beachhead.", 52)

# ======================================================================= PYTHON MIRROR (for summary text and the post-recalc check)
def score_py(over=None, rule=None, weights=None):
    over, rule = over or {}, rule or {}
    w = weights or WSETS[0][1]
    acc = [11494 * .56 * .97, 23785 * .60 * .97, 33035 * .70 * .90]
    need = [x / (13424300 / 155495730) for x in [.21956, .027390, .061719]]
    fit = [2.5, 2.0, 3.0]
    pmf = [1 + 4 * (f - 1) / 2 for f in fit]

    def acvseg(s, use):
        tot_w = tot_n = 0
        for b, idx, _ in BANDS:
            firms = sum(CENSUS[n][0][i] * DEDUP[s][i] for n in SEGN[s] for i in idx)
            rawf = sum(CENSUS[n][0][i] for n in SEGN[s] for i in idx)
            emp = sum(CENSUS[n][1][i] for n in SEGN[s] for i in idx)
            sellers = emp / rawf * [.21956, .027390, .061719][s]
            accn = sellers * P["aps"] * use
            con = accn * P["cpa"]
            cred = con * P["refresh"] * P["crc"] + accn * P["ai"] + accn * P["newacc"] * P["crnew"]
            ops = con * P["refresh"] + accn + accn * P["newacc"]
            cpv, apv = P["dcprice"] * 12 / P["dcsize"], P["acprice"] * 12 / P["acsize"]
            usage = cred * cpv + ops * P["apo"] * apv
            a = max(P["entcr"] * cpv + P["entac"] * apv, usage) * (1 - P["entdisc"]) if b == "E" else max((P["dcprice"] + P["acprice"]) * 12, usage)
            mult = 1 if b == "S" else (P["tiltM"] if b == "M" else P["tiltE"] * P["propE"])
            tot_w += firms * mult
            tot_n += firms * mult * a
        return tot_n / tot_w
    rev = [acvseg(s, 1.0) for s in range(3)]
    comp = [(5 + 4 + 3) / 3, (3 + 4 + 3) / 3, (2 + 2 + 3) / 3]
    ease = [2, 1.5, 4.5]
    M_ = [acc, need, pmf, rev, comp, ease]
    for i, v in over.items():
        M_[i] = [M_[i][j] if v[j] is None else v[j] for j in range(3)]
    S_ = []
    for i in range(6):
        v = M_[i]
        kind = rule.get(i, "ratio") if QUANT[i] else "direct"
        if kind == "ratio":
            S_.append([1 + 4 * x / max(v) for x in v])
        elif kind == "minmax":
            S_.append([1 + 4 * (x - min(v)) / (max(v) - min(v)) for x in v])
        elif kind == "sqrt":
            S_.append([1 + 4 * math.sqrt(x / max(v)) for x in v])
        else:
            S_.append(list(v))
    tot = [sum(w[i] * S_[i][j] for i in range(6)) for j in range(3)]
    return tot, S_, M_


BASE_TOT, BASE_S, BASE_M = score_py()
T11, _, _ = score_py(rule={0: "sqrt"})
T7, _, _ = score_py(over={2: [3, None, None], 4: [3, None, None]})
T1, _, _ = score_py(over={5: [3.5, None, None]})
WS_RES = [score_py(weights=w)[0] for _, w in WSETS]
FLAT = [sum(1 for k in range(1000) if (lambda t: t.index(max(t)))(score_py(weights=list(flat[k]))[0]) == j) / 1000 for j in range(3)]
CENT = [sum(1 for k in range(1000) if (lambda t: t.index(max(t)))(score_py(weights=list(cent[k]))[0]) == j) / 1000 for j in range(3)]

# ======================================================================= EXECUTIVE SUMMARY
title(exs, "Executive Summary — segment prioritization v2 (after partner review)", "Independent case study · not commissioned by or affiliated with Clay Labs Inc. · scores are analyst estimates [C] built from public data")
hdr(exs, 4, ["Result", SHORT[0], SHORT[1], SHORT[2], "Note"], widths=[52, 16, 16, 16, 60])
res_rows = [
    ("Weighted total — v2 base (brief weights)", [f"=Scorecard!{L(3 + j)}{TOT}" for j in range(3)], D2, "One indicator per criterion; Phase 4 ACV fed back"),
    ("Weighted total — display precision (1 decimal)", [f"=ROUND(Scorecard!{L(3 + j)}{TOT},1)" for j in range(3)], '0.0', "Judgment-based inputs do not support two-decimal precision (review m7)"),
    ("Score range — P10, centered random weights", [f"=PERCENTILE(Random_Weights!${c}$5:${c}${RW_END},0.1)" for c in "RST"], '0.0', "Dirichlet(30 × brief weights), 1,000 draws"),
    ("Score range — P90, centered random weights", [f"=PERCENTILE(Random_Weights!${c}$5:${c}${RW_END},0.9)" for c in "RST"], '0.0', "Distribution and professional services ranges overlap: a tie"),
    ("Weighted total — v1 (as published)", [f"=Bridge_v1_v2!{L(3 + j)}5" for j in range(3)], D2, "Archived v1"),
    ("Named weight sets won (of 6)", [f'=COUNTIF(Weight_Sensitivity!$E$5:$E${WS_END},"{SHORT[j]}")' for j in range(3)], '0', "Weight_Sensitivity Block A"),
    ("Random weights won — flat Dirichlet", [f"=Weight_Sensitivity!{L(2 + j)}{RWS}" for j in range(3)], '0.0%', "1,000 draws"),
    ("Random weights won — centered Dirichlet", [f"=Weight_Sensitivity!{L(2 + j)}{RWS + 1}" for j in range(3)], '0.0%', "1,000 draws"),
    ("Stress tests won (of 14)", [f'=COUNTIF(Stress_Tests!$F$5:$F${T_END},"{SHORT[j]}")' for j in range(3)], '0', "Stress_Tests summary"),
]
for k, (lab, fs, fmt, nt) in enumerate(res_rows):
    rr = 5 + k
    put(exs, f"A{rr}", lab, bold=(k == 0))
    for j, f in enumerate(fs):
        put(exs, f"{L(2 + j)}{rr}", f, fmt=fmt, font=GREEN, center=True, fill=KFILL if k == 0 else None)
    put(exs, f"E{rr}", nt)
rr = 5 + len(res_rows) + 1
put(exs, f"A{rr}", "What changed and what it means", bold=True, fill=SFILL)
for c in "BCDE":
    put(exs, f"{c}{rr}", None, fill=SFILL)
rr += 1
gap = BASE_TOT[0] - BASE_TOT[2]
lead = SHORT[0] if gap > 0 else SHORT[2]
W0 = WSETS[0][1]


def bridge_gap(swapped):
    t = [sum(W0[i] * (BASE_S[i][j] if i in swapped else V1[i][j]) for i in range(6)) for j in range(3)]
    return t[0] - t[2]


bg = [bridge_gap(set(range(k))) for k in range(5)]
d_ = [bg[k + 1] - bg[k] for k in range(4)]
sgn = lambda v: ("+" if v >= 0 else "−") + f"{abs(v):.2f}"
msgs = [
    f"1. Distribution's lead disappears. v1: 3.97 vs 3.41 (+0.56). v2: {BASE_TOT[0]:.2f} vs {BASE_TOT[2]:.2f} ({sgn(gap)}); manufacturing {BASE_TOT[1]:.2f} stays last. With one indicator per criterion and Phase 4's list-price ACV, the two leading segments are a statistical tie: moving a single criterion's weight by about 2–6 points flips the order (Weight_Sensitivity Block B).",
    f"2. Bridge from v1 (distribution − professional services gap): reachable accounts instead of SAM dollars {sgn(d_[0])}; dropping the duplicated coverage half of Need {sgn(d_[1])}; evidence-based product-fit index {sgn(d_[2])}; usage-based ACV {sgn(d_[3])}. v1's lead came mostly from scoring market size in dollars at a uniform ACV with a filter built on the same sales share that also drove Need and Revenue.",
    f"3. Robustness is now two-sided: professional services wins 5 of 6 named weight sets and {FLAT[2]:.0%} of flat random weightings; distribution wins under concave size scoring (T11: {T11[0]:.2f} vs {T11[2]:.2f}), min-max scaling, Upside usage, the v1 market definition and a working partner channel (T1). Neither segment wins robustly.",
    "4. The scorecard can no longer decide the segment on its own. Tie-breakers: (a) Clay's differentiation is real only in distribution (competitive position 4.0 vs 2.3); (b) on the 07 v2 economics distribution is better on all six comparisons (Economic_Crosscheck), mainly because more professional-services wins would arrive through Clay's horizontal motion anyway and its list-price ACV sits at the Growth floor.",
    "5. Implication: if Clay tests a vertical at all (07 v2 supports only a low-cost validation or solution play), distribution remains the better test bed, and professional services stays in the horizontal motion. The beachhead is reopened: 4238 machinery/MRO outscores 4236 electrical on v2 indicators (4.70 vs 4.07); 4236's edges are higher ACV per account and the only public reference (Regency) (Subsegment_Drilldown).",
]
for t in msgs:
    put(exs, f"A{rr}", t)
    exs.merge_cells(f"A{rr}:E{rr}")
    exs.row_dimensions[rr].height = 46
    rr += 1

# ======================================================================= CHANGE LOG
title(chg, "Change Log — v1 → v2 (review issue C5)", "Partner review: review/Partner_Review_2026-10-19.md. v1 archived in appendix/archive/.")
hdr(chg, 4, ["Criterion / item", "v1", "v2", "Why"], widths=[24, 60, 60, 50])
CH = [("Market Attractiveness", "50% SAM $ (uniform ACV, outbound filter) + 50% headroom flag (5/5/2)", "Reachable accounts = firms × CRM in use × (1 − already Clay)", "SAM $ embedded an ACV assumption and a sales-share-based filter; headroom double-penalized PS with F2"),
      ("Customer Need", "50% sales intensity + 50% LinkedIn coverage gap", "Sales intensity only", "Coverage proxy already drives Competitive Position F1"),
      ("Product-Market Fit", "Free judgment 4 / 2 / 4", "Capability-fit index from the 05 pain-to-capability map → 4.0 / 3.0 / 5.0", "Traceable to evidence; builder gap kept out (it is Ease)"),
      ("Revenue Potential", "50% SOM ARR + 50% sales staff per firm (assumed ACV lift)", "Usage-based first-year ACV per won account at list prices", "Phase 4 showed seller density lifts ACV only from Mid-Market up"),
      ("CRM readiness (distribution)", "65% (DSG 2022)", "56% (Phocas/DSG 2026, S25)", "Newest evidence (review M2)"),
      ("Robustness", "11 stress tests, 6 weight sets, random weights", "14 stress tests incl. concave size scoring and the v1 market definition; same weight sets; random weights; economic cross-check from 07 v2", ""),
      ("Precision (review m7)", "Totals shown to two decimals and read as a ranking", "Totals displayed to one decimal with P10–P90 ranges under centered random weights (Executive_Summary); deck shows the same", "Judgment scores cannot separate 3.79 from 3.85"),
      ("Drilldown", "Receipts/judgment indicators; 4236 first", "v2 indicators; 4238 scores first; Regency kept as a tie-breaker", "Beachhead reopened")]
for i, row in enumerate(CH):
    for j, v in enumerate(row):
        put(chg, f"{L(1 + j)}{5 + i}", v)
    chg.row_dimensions[5 + i].height = 42

# ======================================================================= COVER & SOURCES
title(cover, "06 — Segment Prioritization v2: which segment, if any, should Clay test first?", "Independent portfolio case study by Zhuoran (Zora) Wang · Not commissioned by or affiliated with Clay Labs Inc. · v2 rebuilt 2026-10-19 after partner review (issue C5)")
crow = [("Question", "Ranking of industrial distribution, industrial manufacturing and B2B professional services on six criteria with the brief's weights."),
        ("v2 rule", "One indicator per criterion. No input feeds two criteria. Phase 4's usage-based ACV replaces the assumed seller-density uplift."),
        ("Structure", "Executive_Summary → Weights → Inputs → Scorecard → Bridge_v1_v2 → Weight_Sensitivity → Random_Weights → Stress_Tests → Economic_Crosscheck → Subsegment_Drilldown → Change_Log → Sources"),
        ("Evidence classes", "A verified fact · B benchmark · C analyst estimate · D scenario assumption · E recommendation. Judgment inputs are yellow."),
        ("Not a forecast", "Scores are analyst estimates from public data. Economic values are illustrative scenario outputs, not Clay forecasts.")]
for i, (k, v) in enumerate(crow, start=4):
    put(cover, f"A{i}", k, bold=True, fill=SFILL)
    put(cover, f"B{i}", v)
    cover.row_dimensions[i].height = 40
cover.column_dimensions["A"].width = 18
cover.column_dimensions["B"].width = 125
title(srcs, "Sources", "Full register: Research_Sources.csv")
hdr(srcs, 4, ["ID", "Source", "Used for"], widths=[8, 70, 60])
for i, s_ in enumerate([("S14", "U.S. Census Bureau, 2022 Economic Census — Employment Size of Firms", "Firm and employee counts, de-duplication"),
                        ("S15", "U.S. BLS, OEWS May 2025 industry-specific estimates", "Sales occupations share"),
                        ("S01", "Clay pricing page", "Usage-based ACV"),
                        ("S25", "Phocas / DSG, State of Sales in Distribution 2026", "CRM in use 56%"),
                        ("S06 / S32", "Clay case-study filters; Bloomberry Clay user analysis", "Already-Clay share assumptions"),
                        ("S07", "Clay case study: Regency Supply", "Drilldown tie-breaker"),
                        ("S53 / S54", "NAED; American Supply Association", "Drilldown channels"),
                        ("04, 05, 07", "Project workbooks and memos", "Competitive factors, pain map, economic cross-check")], start=5):
    for j, v in enumerate(s_):
        put(srcs, f"{L(1 + j)}{i}", v)

order = ["Cover", "Executive_Summary", "Weights", "Inputs", "Scorecard", "Bridge_v1_v2", "Weight_Sensitivity", "Random_Weights", "Stress_Tests", "Economic_Crosscheck", "Subsegment_Drilldown", "Change_Log", "Sources"]
wb._sheets = [wb[n] for n in order]
wb.save(OUT)
print("saved", OUT)
print("python mirror base", [round(x, 4) for x in BASE_TOT], "T11", [round(x, 3) for x in T11], "T7", [round(x, 3) for x in T7], "T1", [round(x, 3) for x in T1])
print("weight sets", [[round(x, 2) for x in t] for t in WS_RES])
print("flat", FLAT, "cent", CENT)
print("metrics", [[round(x, 3) for x in m] for m in BASE_M])
for m in msgs:
    print(m)
