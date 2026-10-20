"""Build 07_Financial_Model.xlsx — Clay vertical expansion case study (Phase 4).
Quarterly model Q1 2027–Q4 2029. One generator ('Engine') writes the live model across GTM Funnel / Revenue Model /
Unit Economics and replicates it for every scenario and sensitivity run, so all outputs are live formulas."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, BarChart, Reference, Series
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as L

BASE = "/home/claude/Clay_Strategy_Consulting_Case"
OUT = f"{BASE}/07_Financial_Model.xlsx"
F = "Arial"
BLUE = Font(name=F, color="0000FF", size=10)
BLACK = Font(name=F, color="000000", size=10)
GREEN = Font(name=F, color="008000", size=10)
HDR = Font(name=F, bold=True, color="FFFFFF", size=10)
TITLE = Font(name=F, bold=True, size=14, color="1F2A44")
SUB = Font(name=F, italic=True, size=9, color="555555")
HFILL = PatternFill("solid", fgColor="1F2A44")
SFILL = PatternFill("solid", fgColor="E8ECF3")
YFILL = PatternFill("solid", fgColor="FFF2CC")
KFILL = PatternFill("solid", fgColor="C6E0B4")
thin = Side(style="thin", color="BFBFBF")
BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="top")
USD = '$#,##0;($#,##0);"-"'
USDK = '$#,##0.0,"K";($#,##0.0,"K");"-"'
USDM = '$#,##0.00,,"M";($#,##0.00,,"M");"-"'
PCT = '0.0%;(0.0%);"-"'
N1 = '#,##0.0;(#,##0.0);"-"'
N0 = '#,##0;(#,##0);"-"'
N3 = '0.0000'
MO = '0.0" mo";(0.0" mo");"-"'
X2 = '0.00"x"'

QCOLS = [L(3 + i) for i in range(12)]          # C..N
QLAB = [f"Q{(i % 4) + 1} {2027 + i // 4}" for i in range(12)]
ANN = [("P", "2027", 0, 4), ("Q", "2028", 4, 8), ("R", "2029", 8, 12)]
TOT = "S"

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
    ws["A2"].value, ws["A2"].font = s, SUB
    ws.sheet_view.showGridLines = False


def note(ws, rng, text, height=34):
    a = rng.split(":")[0]
    put(ws, a, text, border=False)
    ws.merge_cells(rng)
    ws.row_dimensions[int("".join(ch for ch in a if ch.isdigit()))].height = height


# ============================================================ sheet stubs (order fixed at end)
cover = wb.active
cover.title = "Cover"
exo = wb.create_sheet("Executive Output")
asm = wb.create_sheet("Assumptions")
ms = wb.create_sheet("Market Sizing")
acv = wb.create_sheet("ACV Build")
fun = wb.create_sheet("GTM Funnel")
rev = wb.create_sheet("Revenue Model")
ue = wb.create_sheet("Unit Economics")
scn = wb.create_sheet("Scenarios")
sen = wb.create_sheet("Sensitivity")
eng = wb.create_sheet("Engine Runs")
src = wb.create_sheet("Sources")

# ============================================================ ASSUMPTIONS
title(asm, "Assumptions — every input in one place (Downside / Base / Upside)",
      "Change blue cells only. 'Live' feeds the main model (GTM Funnel → Revenue Model → Unit Economics) via the two selectors below. Scenarios/Sensitivity run their own copies of the model.")
put(asm, "A4", "Scenario selector (1 = Downside, 2 = Base, 3 = Upside)", bold=True, fill=SFILL)
asm.merge_cells("A4:C4")
put(asm, "D4", 2, font=BLUE, fill=YFILL, center=True, bold=True)
put(asm, "E4", '=CHOOSE($D$4,"Downside","Base","Upside")', center=True, bold=True)
put(asm, "A5", "Motion selector (1 = Partner-led pilot [Option 4], 2 = Sales-led vertical team [Option 1])", bold=True, fill=SFILL)
asm.merge_cells("A5:C5")
put(asm, "D5", 1, font=BLUE, fill=YFILL, center=True, bold=True)
put(asm, "E5", '=CHOOSE($D$5,"Partner-led pilot","Sales-led vertical team")', center=True, bold=True)
dv1 = DataValidation(type="whole", operator="between", formula1="1", formula2="3")
dv2 = DataValidation(type="whole", operator="between", formula1="1", formula2="2")
asm.add_data_validation(dv1)
asm.add_data_validation(dv2)
dv1.add("D4")
dv2.add("D5")
SEL_S, SEL_M = "Assumptions!$D$4", "Assumptions!$D$5"

hdr(asm, 7, ["ID", "Input", "Unit", "Downside", "Base", "Upside", "Live", "Class", "Source / rationale"],
    widths=[7, 46, 12, 12, 12, 12, 12, 7, 92])
A = {}   # name -> row (scenario inputs)
SECTIONS = [
    ("MARKET", [
        ("f1", "SAM filter 1: share of firms with active outbound/territory motion", "%", .65, .80, .90, "D", "Phase 1 (02_Market_Sizing, Assumptions A4); sales occupations 22% of distribution employment (S15)"),
        ("f2", "SAM filter 2: share CRM-ready (integrable today)", "%", .52, .65, .76, "B/D", "Phase 1 A7; DSG 2022 distributor CRM survey (S17)"),
        ("exist", "Share of SAM already Clay customers (excluded)", "%", .03, .03, .03, "D", "No public data; small given ~40% software mix of detected Clay users (S32)"),
        ("w1", "Wave 1 launch quarter — 4236 electrical distributors (1 = Q1 2027)", "quarter", 1, 1, 1, "E", "Beachhead from 06_Segment_Prioritization (Subsegment_Drilldown)"),
        ("w2", "Wave 2 launch quarter — 4237 hardware, plumbing & HVAC", "quarter", 7, 5, 4, "D", "After first Go/No-Go gate; Downside delays"),
        ("w3", "Wave 3 launch quarter — 4238 machinery & MRO (99 = never)", "quarter", 99, 7, 6, "D", "Downside never launches wave 3"),
    ]),
    ("GTM FUNNEL", [
        ("cov", "Outbound coverage: share of active SAM accounts sequenced per quarter", "%", .30, .40, .50, "D", "Clay-run outbound (Clay's own product); rotating coverage"),
        ("meet", "Outbound meeting rate per account sequenced", "%", .020, .030, .045, "D", "Typical cold outbound meeting rates 1–5% [K]; 'lead lists/email' MQL→SQL are among the weakest sources (S49)"),
        ("inb", "Inbound / template-led leads per quarter (% of active SAM)", "%", .005, .010, .015, "D", "Templates, trade associations, events, Clay community"),
        ("p0", "Partners active at launch", "#", 1, 2, 3, "D", "No distribution-focused Clay partners visible today (S35) → recruitment is a pilot task"),
        ("padd", "Partners added per quarter", "#", 1, 1, 2, "D", ""),
        ("pcap", "Partner network cap", "#", 5, 10, 15, "D", ""),
        ("lpp", "Leads per active partner per quarter", "#", 2, 3, 4, "D", "Warm introductions from implementers / distribution consultancies"),
        ("qual", "Qualification rate — outbound & inbound leads → SQL", "%", .35, .45, .55, "B", "Norwest 2025 ~36% MQL→SQL; SMB/MM average 39% (S48, S49)"),
        ("qualp", "Qualification rate — partner leads → SQL", "%", .50, .60, .70, "D", "Warm, pre-vetted introductions"),
        ("s2o", "SQL → opportunity conversion", "%", .35, .45, .55, "B", "SMB/MM average 42% (S49)"),
        ("win", "Win rate (opportunity → closed-won)", "%", .20, .28, .35, "B", "HubSpot 2024 average 21%; opportunity→won 31–39% (S48, S49); ICP-matched deals win ~3.1× more (Ebsta, S48)"),
        ("lag", "Sales cycle: quarters from opportunity to close (1 or 2)", "qtrs", 2, 1, 1, "B", "Deals < $25K: 2–3 months (Norwest 2025, S48)"),
        ("mixS", "Win mix — SMB (20–99 employees)", "%", .50, .50, .50, "E", "ICP prioritizes 100+ employee distributors where ACV is above the Growth floor (ACV Build)"),
        ("mixM", "Win mix — Mid-Market (100–499)", "%", .38, .38, .38, "E", ""),
        ("mixE", "Win mix — Enterprise (500+)", "%", .12, .12, .12, "E", ""),
        ("ss", "Self-serve share of SMB wins (templates; no AE)", "%", .15, .30, .40, "D", "Low builder capacity limits self-serve (H4); templates raise it"),
    ]),
    ("ACV BUILD (usage-based pricing)", [
        ("salesshare", "Sales occupations share of employment (distribution)", "%", .22, .22, .22, "C", "OEWS May 2025 (S15) via 02_Market_Sizing Segment_Profile"),
        ("aps", "Accounts tracked per seller", "#", 150, 150, 150, "D", "Territory book of business; outside reps cover contractor and plant accounts"),
        ("usage", "Usage intensity factor (multiplies accounts tracked)", "x", .75, 1.0, 1.5, "D", "Range for workflow depth (signals, research, sourcing)"),
        ("cpa", "Contacts per account", "#", 2, 2, 2, "D", ""),
        ("refresh", "Contact refreshes per year", "#", 2, 2, 2, "D", ""),
        ("crc", "Data credits per contact refresh (email + phone waterfall)", "credits", 6, 6, 6, "B/D", "Third-party: email lookups 1–5 credits per record depending on provider (S51)"),
        ("ai", "AI research credits per account per year", "credits", 4, 4, 4, "D", "Fixed-price Claygent models (S01)"),
        ("newacc", "New accounts sourced per year (% of tracked)", "%", .25, .25, .25, "D", "Google Maps / permits sourcing of new contractors (S37, S26)"),
        ("crnew", "Data credits per new account sourced", "credits", 8, 8, 8, "D", ""),
        ("apo", "Actions per row operation", "actions", 6, 6, 6, "D", "Actions 'cost less than a cent each' (S01)"),
        ("dcprice", "Data-credit pack price ($/month, Growth 72K credits/yr)", "$/mo", 261, 261, 261, "A", "Clay pricing page (S01)"),
        ("dcsize", "Data-credit pack size (credits/yr)", "credits", 72000, 72000, 72000, "A", "S01"),
        ("acprice", "Actions pack price ($/month, Growth 480K actions/yr)", "$/mo", 185, 185, 185, "A", "S01"),
        ("acsize", "Actions pack size (actions/yr)", "actions", 480000, 480000, 480000, "A", "S01"),
        ("entcr", "Enterprise minimum data credits per year", "credits", 1200000, 1200000, 1200000, "A", "Enterprise includes 100,000+ data credits per month (S01 FAQ)"),
        ("entac", "Enterprise minimum actions per year", "actions", 2400000, 2400000, 2400000, "A", "Enterprise includes 200,000+ actions per month (S01 FAQ)"),
        ("entdisc", "Enterprise discount vs list", "%", .30, .30, .30, "B", "Vendr: Enterprise 25–35% below initial quotes (S16)"),
    ]),
    ("RETENTION & EXPANSION (annual)", [
        ("grrS", "Gross revenue retention — SMB", "%", .85, .88, .91, "B", "SaaS Capital 2025 median GRR 90% for ACV < $12K (S43); low builder capacity → base below median"),
        ("grrM", "Gross revenue retention — Mid-Market", "%", .86, .90, .93, "B", "SaaS Capital: 91% ($12–25K) (S43)"),
        ("grrE", "Gross revenue retention — Enterprise", "%", .88, .92, .95, "B", "SaaS Capital: 90–91% ($50–250K), 95% (> $250K) (S43)"),
        ("nrrS", "Net revenue retention — SMB", "%", .92, 1.00, 1.06, "B", "SaaS Capital median NRR 98% (< $12K) (S43)"),
        ("nrrM", "Net revenue retention — Mid-Market", "%", .98, 1.06, 1.15, "B", "SaaS Capital 103% ($12–25K) (S43)"),
        ("nrrE", "Net revenue retention — Enterprise", "%", 1.05, 1.15, 1.30, "B", "SaaS Capital 104% ($50–100K); Clay reports enterprise NRR > 200% overall (S02), not assumed for this segment"),
    ]),
    ("SALES CAPACITY & COST RATES", [
        ("quota", "AE annual quota (new ARR)", "$", 875000, 875000, 875000, "B", "Bridge Group 2026: SaaS median quota $875K (n=158) (S44)"),
        ("prod", "AE productivity (average attainment of quota)", "%", .55, .70, .85, "B", "Bridge Group 2026: 48% of reps at quota (S44)"),
        ("aeote", "AE on-target earnings (cash)", "$", 200000, 200000, 200000, "B", "Bridge Group 2026 median OTE $200K (S44)"),
        ("wshare", "Wages as share of total compensation (loading)", "%", .70, .70, .70, "A", "BLS ECEC, private industry, June 2026: wages 70.0% (S45)"),
        ("wmgr", "Sales manager wage (vertical lead, partner manager)", "$", 186910, 186910, 186910, "A", "OEWS May 2025, software publishers 11-2022 mean (S46)"),
        ("wmkt", "Marketing manager wage (product marketing)", "$", 206930, 206930, 206930, "A", "OEWS May 2025, software publishers 11-2021 mean (S46)"),
        ("wdev", "Software developer wage (GTM / solutions engineer)", "$", 158300, 158300, 158300, "A", "OEWS May 2025, software publishers 15-1252 mean (S46)"),
        ("wcs", "Business operations specialist wage (customer success)", "$", 112510, 112510, 112510, "A", "OEWS May 2025, software publishers 13-1199 mean (S46)"),
        ("csmM", "Mid-Market + Enterprise accounts per CSM", "#", 40, 40, 40, "D", ""),
        ("csmS", "SMB accounts per CSM (pooled / digital)", "#", 250, 250, 250, "D", ""),
        ("gm", "Gross margin on Clay revenue (data pass-through, AI, hosting)", "%", .55, .65, .72, "D", "Sacra: Clay margins resemble data-heavy SaaS, below 80%+ pure software (S52); no reported figure"),
        ("launchsp", "Launch marketing spend per wave", "$", 40000, 40000, 40000, "D", "Trade shows, content, association sponsorships"),
        ("ref", "Partner fee on partner-sourced first-year ACV", "%", .25, .20, .15, "B", "Referral 10–15%; VAR 20–30% (Freshworks, Bessemer via S50)"),
        ("enab", "Partner enablement per active partner per quarter", "$", 2000, 2000, 2000, "D", "Training, certification, co-marketing"),
        ("onb", "Onboarding / co-build subsidy, Mid-Market & Enterprise (% of first-year ACV)", "%", .20, .15, .10, "B", "Bessemer: budget 10–20% of purchase price for partner-delivered services (S50)"),
        ("build", "Template & ERP-connector build (one-time)", "$", 250000, 250000, 250000, "D", "Distribution templates (contractor sourcing, branch scoring, e-commerce vetting) + ERP connector gap (05 §3)"),
        ("buildq", "Quarters over which build is spent", "qtrs", 4, 4, 4, "D", ""),
        ("stop", "Program stop quarter (99 = runs all 12 quarters; used for gated cases)", "quarter", 99, 99, 99, "D", "Scenarios tab overrides to 4 for 'Downside, stopped at Q4 2027 gate'"),
        ("life", "Embedded-value horizon after 2029 (years)", "yrs", 5, 5, 5, "D", "Gross profit of the 2029 customer base over 5 more years, decaying at GRR, net of CS cost"),
    ]),
    ("DECISION ANALYSIS", [
        ("pD", "Probability weight — Downside", "%", .35, .35, .35, "D", "Judgment; break-even probability shown on Scenarios"),
        ("pB", "Probability weight — Base", "%", .45, .45, .45, "D", ""),
        ("pU", "Probability weight — Upside", "%", .20, .20, .20, "D", ""),
    ]),
]
r = 8
pct_names = set()
for sec, items in SECTIONS:
    put(asm, f"A{r}", sec, bold=True, fill=SFILL)
    for col in "BCDEFGHI":
        put(asm, f"{col}{r}", None, fill=SFILL)
    r += 1
    for name, label, unit, lo, ba, hi, cls, why in items:
        A[name] = r
        fmt = PCT if unit == "%" else (USD if unit.startswith("$") else ('0.00' if unit == "x" else N0))
        if unit == "%":
            pct_names.add(name)
        put(asm, f"A{r}", f"A{len(A)}")
        put(asm, f"B{r}", label)
        put(asm, f"C{r}", unit, center=True)
        for col, v in zip("DEF", (lo, ba, hi)):
            put(asm, f"{col}{r}", v, font=BLUE, fmt=fmt, center=True, fill=YFILL if cls in ("D", "E", "B/D") else None)
        put(asm, f"G{r}", f"=INDEX(D{r}:F{r},1,{SEL_S})", fmt=fmt, center=True, bold=True)
        put(asm, f"H{r}", cls, center=True)
        put(asm, f"I{r}", why)
        r += 1
    r += 1
put(asm, f"A{r}", "Check: win mix sums to 100% (each scenario)", bold=True)
for col in "DEF":
    put(asm, f"{col}{r}", f"={col}{A['mixS']}+{col}{A['mixM']}+{col}{A['mixE']}", fmt=PCT, center=True)
MIXCHK = r
put(asm, f"A{r+1}", "Check: probabilities sum to 100%", bold=True)
put(asm, f"E{r+1}", f"=E{A['pD']}+E{A['pB']}+E{A['pU']}", fmt=PCT, center=True)
PCHK = r + 1
r += 3
# Motion inputs
put(asm, f"A{r}", "MOTION DESIGN (team and cost structure)", bold=True, fill=SFILL)
for col in "BCDEFGHI":
    put(asm, f"{col}{r}", None, fill=SFILL)
r += 1
hdr(asm, r, ["ID", "Input", "Unit", "Partner-led pilot", "Sales-led team", "", "Live", "Class", "Rationale"])
r += 1
M = {}
MOTION = [
    ("fte_lead", "Vertical lead (player-coach; owns partners in year 1)", "FTE", 1, 1, "Single owner for the vertical P&L"),
    ("pm1", "Partner managers — 2027", "FTE", 0, 1, "Partner-led: vertical lead covers partners in year 1"),
    ("pm2", "Partner managers — 2028 onward", "FTE", 0.5, 1, ""),
    ("eng1", "GTM / solutions engineers — 2027", "FTE", 1, 1, "Builds templates and connector; pilot support"),
    ("eng2", "GTM / solutions engineers — 2028 onward", "FTE", 0.5, 0.5, ""),
    ("pmm", "Product marketing", "FTE", 0.25, 0.5, ""),
    ("prog", "Program marketing per quarter", "$", 25000, 60000, "Events, content, paid; sales-led needs more top-of-funnel"),
    ("shared", "Shared AE capacity borrowed from Clay's existing commercial team", "AE-eq.", 2, 0, "Partner-led closes deals with existing AEs, paying a variable cost per deal"),
    ("varae", "Charge variable AE cost per deal closed with shared capacity (1 = yes)", "flag", 1, 0, ""),
]
for name, label, unit, pl, sl, why in MOTION:
    M[name] = r
    fmt = USD if unit == "$" else '0.00'
    put(asm, f"A{r}", f"M{len(M)}")
    put(asm, f"B{r}", label)
    put(asm, f"C{r}", unit, center=True)
    put(asm, f"D{r}", pl, font=BLUE, fmt=fmt, center=True, fill=YFILL)
    put(asm, f"E{r}", sl, font=BLUE, fmt=fmt, center=True, fill=YFILL)
    put(asm, f"G{r}", f"=INDEX(D{r}:E{r},1,{SEL_M})", fmt=fmt, center=True, bold=True)
    put(asm, f"H{r}", "D", center=True)
    put(asm, f"I{r}", why)
    r += 1
r += 1
put(asm, f"A{r}", "Dedicated AE headcount by quarter (FTE, end of quarter)", bold=True, fill=SFILL)
asm.merge_cells(f"A{r}:C{r}")
r += 1
for i, q in enumerate(QLAB):
    put(asm, f"{L(4+i)}{r}", q, bold=True, center=True, fill=SFILL)
r += 1
HIRE = {}
for k, (nm, sched) in enumerate([("Partner-led pilot", [0] * 12), ("Sales-led team", [2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 6, 6])]):
    put(asm, f"B{r}", nm, bold=True)
    for i, v in enumerate(sched):
        put(asm, f"{L(4+i)}{r}", v, font=BLUE, center=True, fill=YFILL)
    HIRE[k] = r
    r += 1
put(asm, f"B{r}", "Live", bold=True)
for i in range(12):
    c = L(4 + i)
    put(asm, f"{c}{r}", f"=INDEX({c}{HIRE[0]}:{c}{HIRE[1]},{SEL_M},1)", center=True, bold=True)
HIRE["live"] = r
r += 1
put(asm, f"B{r}", "Partner-led + scale-up hires (used only in the Upside scale-up run)", bold=True)
for i, v in enumerate([0, 0, 0, 0, 1, 2, 2, 2, 2, 2, 2, 2]):
    put(asm, f"{L(4+i)}{r}", v, font=BLUE, center=True, fill=YFILL)
HIRE["scale"] = r
r += 1
put(asm, f"B{r}", "Trigger: in Upside, shared AE capacity binds from Q3 2028 (see 'Wins lost to capacity'); two dedicated AEs hired in Q1–Q2 2028 after the gate clear it.", italic=True, border=False)
asm.merge_cells(f"B{r}:O{r}")
for col in "JKLMNO":
    asm.column_dimensions[col].width = 9
asm.freeze_panes = "D8"

# ============================================================ MARKET SIZING
title(ms, "Market Sizing — the addressable opportunity for the distribution pilot",
      "Unique firms with 20+ employees, 2022 Economic Census (S14). Wave 1 = electrical (4236), wave 2 = PHVAC (4237), wave 3 = machinery/MRO (4238). Metals (4235) excluded per Phase 3.")
hdr(ms, 4, ["NAICS", "Sub-industry", "Firms 20–49", "Firms 50–99", "Firms 100–249", "Firms 250–499", "Firms 500+",
            "Emp. 20–49", "Emp. 50–99", "Emp. 100–249", "Emp. 250–499", "Emp. 500+"],
    widths=[9, 38, 11, 11, 12, 12, 11, 12, 12, 12, 12, 12])
cen = [("4236", "Electrical & electronic goods (wave 1)", [1408, 570, 350, 136, 114], [42698, 39297, 54183, 47492, 315093]),
       ("4237", "Hardware, plumbing & HVAC (wave 2)", [1013, 424, 242, 85, 57], [31260, 29270, 36995, 29008, 116895]),
       ("4238", "Machinery, equipment & supplies (wave 3)", [3622, 1218, 746, 240, 183], [110129, 83565, 115979, 83365, 220621])]
for i, (n, lab, fi, em) in enumerate(cen):
    rr = 5 + i
    put(ms, f"A{rr}", n, bold=True)
    put(ms, f"B{rr}", lab)
    for j, v in enumerate(fi + em):
        put(ms, f"{L(3+j)}{rr}", v, font=BLUE, fmt=N0, center=True)
put(ms, "A9", "De-duplication factor for 20–49 band (sector 423)", bold=True)
ms.merge_cells("A9:B9")
put(ms, "C9", 0.973, font=BLUE, fmt='0.000', center=True)
put(ms, "D9", "02_Market_Sizing › Firm_Counts Block 3 (larger bands: no measurable overlap)", border=False)
hdr(ms, 11, ["Wave", "Sub-industry", "SMB 20–99", "Mid-Market 100–499", "Enterprise 500+", "Firms 20+ total", "SAM firms (Live)", "SAM share"])
WAVE = {}
for i in range(3):
    rr = 12 + i
    s = 5 + i
    put(ms, f"A{rr}", f"W{i+1}", bold=True)
    put(ms, f"B{rr}", f"=B{s}")
    put(ms, f"C{rr}", f"=C{s}*$C$9+D{s}", fmt=N0, center=True)
    put(ms, f"D{rr}", f"=E{s}+F{s}", fmt=N0, center=True)
    put(ms, f"E{rr}", f"=G{s}", fmt=N0, center=True)
    put(ms, f"F{rr}", f"=SUM(C{rr}:E{rr})", fmt=N0, center=True, bold=True)
    put(ms, f"G{rr}", f"=F{rr}*Assumptions!$G${A['f1']}*Assumptions!$G${A['f2']}*(1-Assumptions!$G${A['exist']})", fmt=N0, center=True, font=GREEN)
    put(ms, f"H{rr}", f"=G{rr}/F{rr}", fmt=PCT, center=True)
    WAVE[i] = rr
put(ms, "A15", "Total", bold=True, fill=SFILL)
put(ms, "B15", None, fill=SFILL)
for c in "CDEFG":
    put(ms, f"{c}15", f"=SUM({c}12:{c}14)", fmt=N0, center=True, bold=True, fill=SFILL)
put(ms, "H15", "=G15/F15", fmt=PCT, center=True, fill=SFILL)
hdr(ms, 17, ["Band", "", "Firms (3 waves)", "Employees", "Avg employees per firm", "", "", ""])
EMP = {}
for i, (b, fc, ec) in enumerate([("SMB 20–99", ["C", "D"], ["H", "I"]), ("Mid-Market 100–499", ["E", "F"], ["J", "K"]), ("Enterprise 500+", ["G"], ["L"])]):
    rr = 18 + i
    put(ms, f"A{rr}", b, bold=True)
    ms.merge_cells(f"A{rr}:B{rr}")
    put(ms, f"C{rr}", "=" + "+".join(f"SUM({c}5:{c}7)" for c in fc), fmt=N0, center=True)
    put(ms, f"D{rr}", "=" + "+".join(f"SUM({c}5:{c}7)" for c in ec), fmt=N0, center=True)
    put(ms, f"E{rr}", f"=D{rr}/C{rr}", fmt=N1, center=True, bold=True)
    EMP[i] = f"'Market Sizing'!$E${rr}"
note(ms, "A22:L22", "Average employees per firm feed the ACV Build (sellers per firm = employees × sales share). The 20–49 band is not de-duplicated here because this is an average, not a count.")
hdr(ms, 24, ["Reconciliation to Phase 1", "", "Phase 1 (02)", "Phase 4 (this model)", "Difference / reason", "", "", ""])
recon = [
    ("Distribution SAM firms", 5976.6, "=G15", "Phase 4 excludes 4235 metals (Phase 3) and Clay's existing customers"),
    ("Base ACV — SMB", 7500, "=Assumptions!$E$" + "{acvS}", "Usage build: SMB usage sits below the Growth floor, so ACV = $5,352 list"),
    ("Base ACV — Mid-Market", 20000, "=Assumptions!$E$" + "{acvM}", "Usage build: ~45 sellers × 150 accounts → ~$9.8K at list"),
    ("Base ACV — Enterprise", 60000, "=Assumptions!$E$" + "{acvE}", "Usage build after 30% discount ≈ Vendr median ($59,996)"),
    ("Year-3 ARR (Base, Phase 1 SOM vs Phase 4 exit ARR)", 1864124, "=Scenarios!$D${exitarr}", "Funnel- and wave-based build with lower SMB/MM ACV"),
]
RECON_START = 25

# ============================================================ ACV formula generator
def acv_formula(band, g):
    """Return an Excel formula string for first-year ACV of band 0/1/2 using getter g(name)."""
    emp = EMP[band]
    cp = f"({g('dcprice')}*12/{g('dcsize')})"
    ap = f"({g('acprice')}*12/{g('acsize')})"
    acc = f"({emp}*{g('salesshare')}*{g('aps')}*{g('usage')})"
    con = f"({acc}*{g('cpa')})"
    cred = f"({con}*{g('refresh')}*{g('crc')}+{acc}*{g('ai')}+{acc}*{g('newacc')}*{g('crnew')})"
    ops = f"({con}*{g('refresh')}+{acc}+{acc}*{g('newacc')})"
    usage = f"({cred}*{cp}+{ops}*{g('apo')}*{ap})"
    floor = f"(({g('dcprice')}+{g('acprice')})*12)"
    if band < 2:
        return f"MAX({floor},{usage})"
    entfloor = f"({g('entcr')}*{cp}+{g('entac')}*{ap})"
    return f"MAX({entfloor},{usage})*(1-{g('entdisc')})"


def g_col(col):
    return lambda n: f"Assumptions!${col}${A[n]}"


# ACV rows in Assumptions (computed, per scenario) -> so engine getters can read ACV like any input
r_acv = r + 2
put(asm, f"A{r_acv}", "COMPUTED FROM ACV BUILD (do not edit)", bold=True, fill=SFILL)
for col in "BCDEFGHI":
    put(asm, f"{col}{r_acv}", None, fill=SFILL)
for k, (nm, lab) in enumerate([("acvS", "First-year ACV — SMB"), ("acvM", "First-year ACV — Mid-Market"), ("acvE", "First-year ACV — Enterprise")]):
    rr = r_acv + 1 + k
    A[nm] = rr
    put(asm, f"B{rr}", lab)
    put(asm, f"C{rr}", "$", center=True)
    for col in "DEF":
        put(asm, f"{col}{rr}", f"='ACV Build'!{col}{40+k}", fmt=USD, center=True, font=GREEN)
    put(asm, f"G{rr}", f"=INDEX(D{rr}:F{rr},1,{SEL_S})", fmt=USD, center=True, bold=True)
    put(asm, f"H{rr}", "C", center=True)
    put(asm, f"I{rr}", "Usage-based build on 'ACV Build' tab")

# ============================================================ ACV BUILD sheet
title(acv, "ACV Build — what would a distributor actually spend on Clay?",
      "Bottom-up from sellers per firm × accounts tracked × Clay's usage meters (Data Credits and Actions) at published Growth-tier rates. SMB and Mid-Market are floored at the Growth plan (CRM sync requires Growth). Enterprise is floored at the Enterprise minimum, less discount.")
hdr(acv, 4, ["Step", "Formula", "", "Downside", "Base", "Upside", "Live"], widths=[52, 46, 2, 15, 15, 15, 15])
steps = []
r2 = 5
GCOLS = {"D": g_col("D"), "E": g_col("E"), "F": g_col("F"), "G": g_col("G")}
put(acv, f"A{r2}", "Unit prices (from published plans)", bold=True, fill=SFILL)
r2 += 1
for lab, fn, fmt in [("Price per data credit", lambda g: f"={g('dcprice')}*12/{g('dcsize')}", '$0.0000'),
                     ("Price per action", lambda g: f"={g('acprice')}*12/{g('acsize')}", '$0.00000'),
                     ("Growth plan floor (annual)", lambda g: f"=({g('dcprice')}+{g('acprice')})*12", USD),
                     ("Enterprise minimum at list (annual)", lambda g: f"={g('entcr')}*{g('dcprice')}*12/{g('dcsize')}+{g('entac')}*{g('acprice')}*12/{g('acsize')}", USD)]:
    put(acv, f"A{r2}", lab)
    for col in "DEFG":
        put(acv, f"{col}{r2}", fn(GCOLS[col]), fmt=fmt, center=True)
    r2 += 1
r2 += 1
bands = ["SMB (20–99)", "Mid-Market (100–499)", "Enterprise (500+)"]
for b in range(3):
    put(acv, f"A{r2}", bands[b], bold=True, fill=SFILL)
    for col in "BCDEFG":
        put(acv, f"{col}{r2}", None, fill=SFILL)
    r2 += 1
    rows_b = {}
    defs = [
        ("emp", "Avg employees per firm", "Market Sizing", lambda g: f"={EMP[b]}", N1),
        ("sell", "Sellers per firm", "employees × sales share", lambda g: f"={{emp}}*{g('salesshare')}", N1),
        ("acc", "Accounts tracked", "sellers × accounts/seller × usage factor", lambda g: f"={{sell}}*{g('aps')}*{g('usage')}", N0),
        ("con", "Contacts tracked", "accounts × contacts/account", lambda g: f"={{acc}}*{g('cpa')}", N0),
        ("cred", "Data credits per year", "refreshes + AI research + new-account sourcing", lambda g: f"={{con}}*{g('refresh')}*{g('crc')}+{{acc}}*{g('ai')}+{{acc}}*{g('newacc')}*{g('crnew')}", N0),
        ("act", "Actions per year", "row operations × actions/op", lambda g: f"=({{con}}*{g('refresh')}+{{acc}}+{{acc}}*{g('newacc')})*{g('apo')}", N0),
        ("use", "Usage cost at list", "credits × price + actions × price", lambda g: f"={{cred}}*{g('dcprice')}*12/{g('dcsize')}+{{act}}*{g('acprice')}*12/{g('acsize')}", USD),
    ]
    for key, lab, how, fn, fmt in defs:
        put(acv, f"A{r2}", lab)
        put(acv, f"B{r2}", how, italic=True)
        for col in "DEFG":
            f = fn(GCOLS[col])
            for kk, rv in rows_b.items():
                f = f.replace("{" + kk + "}", f"{col}{rv}")
            put(acv, f"{col}{r2}", f, fmt=fmt, center=True)
        rows_b[key] = r2
        r2 += 1
    r2 += 1
# final ACV rows at fixed row 40..42
assert r2 <= 39, r2
put(acv, "A39", "First-year ACV (feeds Assumptions)", bold=True, fill=KFILL)
for col in "BCDEFG":
    put(acv, f"{col}39", None, fill=KFILL)
for b in range(3):
    rr = 40 + b
    put(acv, f"A{rr}", f"ACV — {bands[b]}", bold=True)
    put(acv, f"B{rr}", "MAX(floor, usage)" if b < 2 else "MAX(Enterprise minimum, usage) × (1 − discount)", italic=True)
    for col in "DEFG":
        put(acv, f"{col}{rr}", "=" + acv_formula(b, GCOLS[col]), fmt=USD, center=True, bold=True)
put(acv, "A44", "Check: Base Enterprise ACV vs Vendr median buyer ($59,996, S16)")
put(acv, "E44", "=E42/59996", fmt=X2, center=True)
note(acv, "A46:G46", "Finding: at published rates, an SMB distributor's usage (~$2K/yr) sits well below the Growth floor, so seller density does NOT lift SMB ACV. It matters only from mid-market upward. This tempers Phase 3's revenue-potential argument (stress test T4 there) and is why the pilot's ICP leans to 100+ employee distributors.", 48)

# ============================================================ ENGINE
class Engine:
    def __init__(self, stages, get, hire, label):
        self.stages = stages            # stage -> [ws, next_row]
        self.get = get                  # name -> cell ref / literal
        self.hire = hire                # i -> cell ref
        self.label = label
        self.reg = {}
        self.cur = None
        self.out = {}

    def _r(self, key, col, absolute=False):
        ws, row = self.reg[key]
        cur_ws = self.stages[self.cur][0]
        c = f"${col}${row}" if absolute else f"{col}{row}"
        return c if ws.title == cur_ws.title else f"'{ws.title}'!{c}"

    def S(self, key):
        return self._r(key, "C", absolute=True)

    def Q(self, key, i):
        return self._r(key, QCOLS[i])

    def section(self, stage, text):
        ws, row = self.stages[stage]
        put(ws, f"A{row}", text, bold=True, fill=SFILL)
        for col in ["B"] + QCOLS + ["P", "Q", "R", "S"]:
            put(ws, f"{col}{row}", None, fill=SFILL)
        self.stages[stage][1] += 1

    def header(self, stage, text):
        ws, row = self.stages[stage]
        put(ws, f"A{row}", text, bold=True, font=Font(color="FFFFFF"), fill=HFILL)
        put(ws, f"B{row}", "Unit / note", font=Font(color="FFFFFF"), fill=HFILL, bold=True)
        for i, q in enumerate(QLAB):
            put(ws, f"{QCOLS[i]}{row}", q, font=Font(color="FFFFFF"), fill=HFILL, bold=True, center=True)
        for col, y, _, _ in ANN:
            put(ws, f"{col}{row}", y, font=Font(color="FFFFFF"), fill=HFILL, bold=True, center=True)
        put(ws, f"{TOT}{row}", "2027–29", font=Font(color="FFFFFF"), fill=HFILL, bold=True, center=True)
        self.stages[stage][1] += 1

    def scalar(self, stage, key, label, formula, fmt, note_=""):
        self.cur = stage
        ws, row = self.stages[stage]
        put(ws, f"A{row}", label)
        put(ws, f"B{row}", note_, italic=True)
        self.reg[key] = (ws, row)
        put(ws, f"C{row}", formula if not callable(formula) else formula(), fmt=fmt, center=True)
        self.stages[stage][1] += 1

    def row(self, stage, key, label, unit, fn, fmt, agg="sum", bold=False, ann=None):
        self.cur = stage
        ws, row = self.stages[stage]
        self.reg[key] = (ws, row)
        put(ws, f"A{row}", label, bold=bold)
        put(ws, f"B{row}", unit, italic=True)
        for i in range(12):
            v = fn(i)
            put(ws, f"{QCOLS[i]}{row}", v, fmt=fmt, wrap=False, bold=bold)
            ws[f"{QCOLS[i]}{row}"].alignment = RIGHT
        for col, y, a, b in ANN:
            if ann:
                f = ann(col, a, b)
            elif agg == "sum":
                f = f"=SUM({QCOLS[a]}{row}:{QCOLS[b-1]}{row})"
            elif agg == "end":
                f = f"={QCOLS[b-1]}{row}"
            elif agg == "avg":
                f = f"=AVERAGE({QCOLS[a]}{row}:{QCOLS[b-1]}{row})"
            else:
                f = None
            put(ws, f"{col}{row}", f, fmt=fmt, bold=True, wrap=False)
            ws[f"{col}{row}"].alignment = RIGHT
        if ann:
            f = ann(TOT, 0, 12)
        elif agg == "sum":
            f = f"=SUM(P{row}:R{row})"
        elif agg == "end":
            f = f"=R{row}"
        else:
            f = None
        put(ws, f"{TOT}{row}", f, fmt=fmt, bold=True, wrap=False, fill=SFILL)
        ws[f"{TOT}{row}"].alignment = RIGHT
        self.stages[stage][1] += 1

    def outrow(self, stage, key, label, formula, fmt):
        self.cur = stage
        ws, row = self.stages[stage]
        put(ws, f"A{row}", label, bold=True)
        put(ws, f"C{row}", formula, fmt=fmt, bold=True, fill=KFILL, center=True)
        ws.merge_cells(f"C{row}:D{row}")
        self.reg[key] = (ws, row)
        self.out[key] = f"'{ws.title}'!$C${row}"
        self.stages[stage][1] += 1

    # ---------------------------------------------------------------- full model
    def build(self):
        g, S, Q = self.get, self.S, self.Q

        def prev(key, i, first="0"):
            return Q(key, i - 1) if i > 0 else first

        # ---------------- FUNNEL stage: scalars
        self.header("fun", self.label + " — inputs carried into this run")
        self.scalar("fun", "samw1", "SAM accounts — wave 1 (4236 electrical)", f"='Market Sizing'!$F${WAVE[0]}*{g('f1')}*{g('f2')}*(1-{g('exist')})", N0, "firms × filters × (1 − existing)")
        self.scalar("fun", "samw2", "SAM accounts — wave 2 (4237 PHVAC)", f"='Market Sizing'!$F${WAVE[1]}*{g('f1')}*{g('f2')}*(1-{g('exist')})", N0, "")
        self.scalar("fun", "samw3", "SAM accounts — wave 3 (4238 machinery/MRO)", f"='Market Sizing'!$F${WAVE[2]}*{g('f1')}*{g('f2')}*(1-{g('exist')})", N0, "")
        self.scalar("fun", "acvS", "First-year ACV — SMB", "=" + acv_formula(0, g), USD, "usage build (ACV Build)")
        self.scalar("fun", "acvM", "First-year ACV — Mid-Market", "=" + acv_formula(1, g), USD, "")
        self.scalar("fun", "acvE", "First-year ACV — Enterprise", "=" + acv_formula(2, g), USD, "")
        self.scalar("fun", "den", "Sales-assisted mix denominator", f"={g('mixS')}*(1-{g('ss')})+{g('mixM')}+{g('mixE')}", N3, "mixS(1−ss)+mixM+mixE")
        self.cur = "fun"
        self.scalar("fun", "acvsa", "Blended ACV of sales-assisted wins", lambda: f"=({g('mixS')}*(1-{g('ss')})*{S('acvS')}+{g('mixM')}*{S('acvM')}+{g('mixE')}*{S('acvE')})/{S('den')}", USD, "")
        self.scalar("fun", "aeload", "AE fully loaded annual cost", f"={g('aeote')}/{g('wshare')}", USD, "OTE ÷ wage share")
        self.scalar("fun", "dealsae", "Deals per ramped AE per quarter", lambda: f"={g('quota')}*{g('prod')}/4/{S('acvsa')}", N1, "quota × productivity ÷ 4 ÷ ACV")
        self.scalar("fun", "cpd", "Variable AE cost per deal (shared capacity)", lambda: f"={S('aeload')}/4/{S('dealsae')}", USD, "")
        # ---------------- FUNNEL stage: quarterly
        self.header("fun", "GTM funnel")
        self.row("fun", "qn", "Quarter number", "#", lambda i: i + 1, N0, agg=None)
        self.row("fun", "run", "Program running (1) / stopped (0)", "flag", lambda i: f"=IF({Q('qn', i)}<={g('stop')},1,0)", N0, agg=None)
        for w, nm in enumerate(["w1", "w2", "w3"]):
            self.row("fun", f"ramp{w}", f"Wave {w+1} active (0.5 in launch quarter)", "factor",
                     lambda i, nm=nm: f"=IF({Q('qn', i)}={g(nm)},0.5,IF({Q('qn', i)}>{g(nm)},1,0))", '0.0', agg=None)
        self.row("fun", "act", "Active SAM accounts in launched waves", "accounts",
                 lambda i: f"={S('samw1')}*{Q('ramp0', i)}+{S('samw2')}*{Q('ramp1', i)}+{S('samw3')}*{Q('ramp2', i)}", N0, agg="end")
        self.row("fun", "cont", "Accounts sequenced by Clay-run outbound", "accounts", lambda i: f"={Q('act', i)}*{g('cov')}*{Q('run', i)}", N0)
        self.row("fun", "lo", "Leads — outbound meetings", "leads", lambda i: f"={Q('cont', i)}*{g('meet')}", N1)
        self.row("fun", "li", "Leads — inbound / template-led", "leads", lambda i: f"={Q('act', i)}*{g('inb')}*{Q('run', i)}", N1)
        self.row("fun", "parts", "Active partners", "partners", lambda i: f"=IF({Q('run', i)}=1,MIN({g('pcap')},{g('p0')}+{g('padd')}*({Q('qn', i)}-1)),0)", N1, agg="end")
        self.row("fun", "lp", "Leads — partner-sourced", "leads", lambda i: f"={Q('parts', i)}*{g('lpp')}", N1)
        self.row("fun", "lt", "Total leads", "leads", lambda i: f"={Q('lo', i)}+{Q('li', i)}+{Q('lp', i)}", N1, bold=True)
        self.row("fun", "sqlnp", "Qualified (SQL) — outbound + inbound", "SQLs", lambda i: f"=({Q('lo', i)}+{Q('li', i)})*{g('qual')}", N1)
        self.row("fun", "sqlp", "Qualified (SQL) — partner", "SQLs", lambda i: f"={Q('lp', i)}*{g('qualp')}", N1)
        self.row("fun", "opp", "Opportunities created", "opps", lambda i: f"=({Q('sqlnp', i)}+{Q('sqlp', i)})*{g('s2o')}", N1, bold=True)
        self.row("fun", "oppp", "   of which partner-sourced", "opps", lambda i: f"={Q('sqlp', i)}*{g('s2o')}", N1)

        def lagged(key, i):
            if i == 0:
                return "0"
            if i == 1:
                return f"IF({g('lag')}=1,{Q(key, 0)},0)"
            return f"IF({g('lag')}=1,{Q(key, i-1)},{Q(key, i-2)})"
        self.row("fun", "dw", "Demand-limited wins (opps from 1–2 quarters earlier × win rate)", "wins", lambda i: f"={lagged('opp', i)}*{g('win')}", N1)
        self.row("fun", "psh", "Partner-sourced share of wins", "%", lambda i: f"=IFERROR({lagged('oppp', i)}/{lagged('opp', i)},0)", PCT, agg=None)
        self.row("fun", "ssw", "Self-serve SMB wins (no AE)", "wins", lambda i: f"={Q('dw', i)}*{g('mixS')}*{g('ss')}", N1)
        self.row("fun", "sad", "Sales-assisted demand", "wins", lambda i: f"={Q('dw', i)}-{Q('ssw', i)}", N1)
        self.row("fun", "ae", "Dedicated AEs (headcount)", "FTE", lambda i: f"={self.hire(i)}", N1, agg="end")

        def ramped(i):
            if i == 0:
                return "=0"
            if i == 1:
                return f"=0.5*{Q('ae', 0)}"
            return f"={Q('ae', i-2)}+0.5*({Q('ae', i-1)}-{Q('ae', i-2)})"
        self.row("fun", "aer", "Ramped dedicated AEs (≈6-month ramp)", "FTE", ramped, N1, agg="end")
        self.row("fun", "cap", "Sales capacity (dedicated + shared AEs)", "deals", lambda i: f"=({Q('aer', i)}+{g('shared')}*{Q('run', i)})*{S('dealsae')}", N1)
        self.row("fun", "saw", "Sales-assisted wins = MIN(demand, capacity)", "wins", lambda i: f"=MIN({Q('sad', i)},{Q('cap', i)})", N1)
        self.row("fun", "lost", "Wins lost to capacity constraint", "wins", lambda i: f"={Q('sad', i)}-{Q('saw', i)}", N1)
        self.row("fun", "new", "New customers", "customers", lambda i: f"={Q('ssw', i)}+{Q('saw', i)}", N1, bold=True)
        self.row("fun", "nS", "   New SMB", "customers", lambda i: f"={Q('ssw', i)}+{Q('saw', i)}*{g('mixS')}*(1-{g('ss')})/{S('den')}", N1)
        self.row("fun", "nM", "   New Mid-Market", "customers", lambda i: f"={Q('saw', i)}*{g('mixM')}/{S('den')}", N1)
        self.row("fun", "nE", "   New Enterprise", "customers", lambda i: f"={Q('saw', i)}*{g('mixE')}/{S('den')}", N1)
        self.row("fun", "conv", "Lead-to-win conversion (lagged)", "%", lambda i: f"=IFERROR({Q('new', i)}/{Q('lt', i-1)},0)" if i > 0 else "=0", PCT, agg=None)

        # ---------------- REVENUE stage
        self.header("rev", "Customers, ARR and revenue")
        for b, (sfx, bl) in enumerate([("S", "SMB"), ("M", "Mid-Market"), ("E", "Enterprise")]):
            self.row("rev", f"c{sfx}", f"Customers — {bl}", "customers",
                     lambda i, sfx=sfx: f"={prev('c'+sfx, i)}*{g('grr'+sfx)}^(1/4)+{Q('n'+sfx, i)}", N1, agg="end")
        self.row("rev", "cT", "Customers — total", "customers", lambda i: f"={Q('cS', i)}+{Q('cM', i)}+{Q('cE', i)}", N1, agg="end", bold=True)
        for sfx, bl in [("S", "SMB"), ("M", "Mid-Market"), ("E", "Enterprise")]:
            self.row("rev", f"a{sfx}", f"ARR — {bl} (end of quarter)", "$",
                     lambda i, sfx=sfx: f"={prev('a'+sfx, i)}*{g('nrr'+sfx)}^(1/4)+{Q('n'+sfx, i)}*{S('acv'+sfx)}", USD, agg="end")
        self.row("rev", "aT", "ARR — total (end of quarter)", "$", lambda i: f"={Q('aS', i)}+{Q('aM', i)}+{Q('aE', i)}", USD, agg="end", bold=True)
        self.row("rev", "na", "New ARR booked (first-year ACV)", "$",
                 lambda i: f"={Q('nS', i)}*{S('acvS')}+{Q('nM', i)}*{S('acvM')}+{Q('nE', i)}*{S('acvE')}", USD)
        self.row("rev", "ex", "Net expansion / (contraction) of existing ARR", "$",
                 lambda i: f"={Q('aT', i)}-{prev('aT', i)}-{Q('na', i)}", USD)
        self.row("rev", "rv", "Recognized revenue (average ARR ÷ 4)", "$", lambda i: f"=({prev('aT', i)}+{Q('aT', i)})/8", USD, bold=True)
        self.row("rev", "gp", "Gross profit", "$", lambda i: f"={Q('rv', i)}*{g('gm')}", USD)

        # ---------------- UNIT ECONOMICS stage
        self.header("ue", "Incremental cost of the vertical program")
        self.row("ue", "fl", "FTE — vertical lead", "FTE", lambda i: f"={g('fte_lead')}*{Q('run', i)}", N1, agg="avg")
        self.row("ue", "fa", "FTE — dedicated AEs", "FTE", lambda i: f"={Q('ae', i)}*{Q('run', i)}", N1, agg="avg")
        self.row("ue", "fp", "FTE — partner managers", "FTE", lambda i: f"=IF({Q('qn', i)}<=4,{g('pm1')},{g('pm2')})*{Q('run', i)}", N1, agg="avg")
        self.row("ue", "fe", "FTE — GTM / solutions engineers", "FTE", lambda i: f"=IF({Q('qn', i)}<=4,{g('eng1')},{g('eng2')})*{Q('run', i)}", N1, agg="avg")
        self.row("ue", "fm", "FTE — product marketing", "FTE", lambda i: f"={g('pmm')}*{Q('run', i)}", N1, agg="avg")
        self.row("ue", "fc", "FTE — customer success (coverage ratios)", "FTE", lambda i: f"=({Q('cM', i)}+{Q('cE', i)})/{g('csmM')}+{Q('cS', i)}/{g('csmS')}", N1, agg="avg")
        self.row("ue", "kl", "Cost — vertical lead + partner managers", "$", lambda i: f"=({Q('fl', i)}+{Q('fp', i)})*{g('wmgr')}/{g('wshare')}/4", USD)
        self.row("ue", "ka", "Cost — dedicated AEs", "$", lambda i: f"={Q('fa', i)}*{S('aeload')}/4", USD)
        self.row("ue", "kv", "Cost — variable AE cost (shared capacity deals)", "$", lambda i: f"=MAX(0,{Q('saw', i)}-{Q('aer', i)}*{S('dealsae')})*{S('cpd')}*{g('varae')}", USD)
        self.row("ue", "km", "Cost — product marketing", "$", lambda i: f"={Q('fm', i)}*{g('wmkt')}/{g('wshare')}/4", USD)
        self.row("ue", "kp", "Cost — program and launch marketing", "$",
                 lambda i: f"=({g('prog')}+{g('launchsp')}*(({Q('qn', i)}={g('w1')})+({Q('qn', i)}={g('w2')})+({Q('qn', i)}={g('w3')})))*{Q('run', i)}", USD)
        self.row("ue", "kr", "Cost — partner fees on partner-sourced ACV", "$", lambda i: f"={Q('na', i)}*{Q('psh', i)}*{g('ref')}", USD)
        self.row("ue", "kn", "Cost — partner enablement", "$", lambda i: f"={Q('parts', i)}*{g('enab')}", USD)
        self.row("ue", "ko", "Cost — onboarding / co-build subsidy (MM & Ent)", "$", lambda i: f"=({Q('nM', i)}*{S('acvM')}+{Q('nE', i)}*{S('acvE')})*{g('onb')}", USD)
        self.row("ue", "sm", "Sales & marketing cost (CAC basis)", "$",
                 lambda i: f"={Q('kl', i)}+{Q('ka', i)}+{Q('kv', i)}+{Q('km', i)}+{Q('kp', i)}+{Q('kr', i)}+{Q('kn', i)}+{Q('ko', i)}", USD, bold=True)
        self.row("ue", "ke", "Cost — GTM / solutions engineers", "$", lambda i: f"={Q('fe', i)}*{g('wdev')}/{g('wshare')}/4", USD)
        self.row("ue", "kb", "Cost — template & ERP-connector build (one-time)", "$", lambda i: f"=IF({Q('qn', i)}<={g('buildq')},{g('build')}/{g('buildq')},0)", USD)
        self.row("ue", "kc", "Cost — customer success", "$", lambda i: f"={Q('fc', i)}*{g('wcs')}/{g('wshare')}/4", USD)
        self.row("ue", "kt", "Total incremental program cost (excl. COGS)", "$", lambda i: f"={Q('sm', i)}+{Q('ke', i)}+{Q('kb', i)}+{Q('kc', i)}", USD, bold=True)
        self.header("ue", "Program P&L and unit economics")
        self.row("ue", "gp2", "Gross profit (from Revenue Model)", "$", lambda i: f"={Q('gp', i)}", USD)
        self.row("ue", "cm", "Contribution after customer success", "$", lambda i: f"={Q('gp', i)}-{Q('kc', i)}", USD)
        self.row("ue", "pl", "Program P&L (gross profit − all incremental cost)", "$", lambda i: f"={Q('gp', i)}-{Q('kt', i)}", USD, bold=True)
        self.row("ue", "cum", "Cumulative program P&L (investment position)", "$", lambda i: f"={prev('cum', i)}+{Q('pl', i)}", USD, agg="end", bold=True)
        self.row("ue", "pos", "Helper: quarter number if P&L > 0, else 99", "#", lambda i: f"=IF(AND({Q('pl', i)}>0,{Q('run', i)}=1),{Q('qn', i)},99)", N0, agg=None)

        def ratio_ann(num, den_, mult="1"):
            def f(col, a, b):
                return f"=IFERROR({self._r(num, col)}/{self._r(den_, col)}*{mult},0)"
            return f
        self.row("ue", "cac", "CAC (S&M ÷ new customers)", "$", lambda i: f"=IFERROR({Q('sm', i)}/{Q('new', i)},0)", USD, ann=ratio_ann("sm", "new"))
        self.row("ue", "acvn", "Average first-year ACV of new customers", "$", lambda i: f"=IFERROR({Q('na', i)}/{Q('new', i)},0)", USD, ann=ratio_ann("na", "new"))
        self.row("ue", "pb", "CAC payback (months of gross profit)", "months",
                 lambda i: f"=IFERROR({Q('cac', i)}/({Q('acvn', i)}*{g('gm')}/12),0)", MO,
                 ann=lambda col, a, b: f"=IFERROR({self._r('cac', col)}/({self._r('acvn', col)}*{g('gm')}/12),0)")
        # outputs
        self.scalar("ue", "grrb", "ARR-weighted blended GRR (new-customer mix)", lambda: (
            f"=({g('mixS')}*{S('acvS')}*{g('grrS')}+{g('mixM')}*{S('acvM')}*{g('grrM')}+{g('mixE')}*{S('acvE')}*{g('grrE')})"
            f"/({g('mixS')}*{S('acvS')}+{g('mixM')}*{S('acvM')}+{g('mixE')}*{S('acvE')})"), PCT, "")
        self.cur = "ue"
        self.header("ue", "Key outputs of this run")
        self.outrow("ue", "o_exit", "Exit ARR, Q4 2029", f"={self._r('aT', 'N')}", USD)
        self.outrow("ue", "o_cust", "Customers, Q4 2029", f"={self._r('cT', 'N')}", N0)
        self.outrow("ue", "o_newc", "New customers 2027–29", f"={self._r('new', TOT)}", N0)
        self.outrow("ue", "o_rev", "Recognized revenue 2027–29", f"={self._r('rv', TOT)}", USD)
        self.outrow("ue", "o_sm", "S&M cost 2027–29", f"={self._r('sm', TOT)}", USD)
        self.outrow("ue", "o_cost", "Total incremental cost 2027–29", f"={self._r('kt', TOT)}", USD)
        self.outrow("ue", "o_acv", "Average first-year ACV", f"={self._r('acvn', TOT)}", USD)
        self.outrow("ue", "o_cac", "CAC (2027–29)", f"={self._r('cac', TOT)}", USD)
        self.outrow("ue", "o_pb", "CAC payback, months (2027–29)", f"={self._r('pb', TOT)}", MO)
        self.outrow("ue", "o_pb29", "CAC payback, months (2029 cohort)", f"={self._r('pb', 'R')}", MO)
        self.outrow("ue", "o_ltv", "LTV : CAC (lifetime capped at 5 years)", f"=IFERROR({self._r('acvn', TOT)}*{g('gm')}*MIN(1/(1-{S('grrb')}),5)/{self._r('cac', TOT)},0)", X2)
        self.outrow("ue", "o_cum", "Cumulative program P&L 2027–29", f"={self._r('cum', 'N')}", USD)
        self.outrow("ue", "o_peak", "Peak investment (lowest cumulative P&L)", f"=MIN({self._r('cum', 'C')}:{self._r('cum', 'N').split('!')[-1]})", USD)
        self.outrow("ue", "o_be", "First quarter with positive program P&L (99 = not by Q4 2029)", f"=MIN({self._r('pos', 'C')}:{self._r('pos', 'N').split('!')[-1]})", N0)
        self.outrow("ue", "o_rr", "Q4 2029 program P&L, annualized", f"={self._r('pl', 'N')}*4", USD)
        self.outrow("ue", "o_emb", "Embedded value of 2029 ARR base (5 yrs, net of CS, no new spend)",
                    f"={self._r('aT', 'N')}*({g('gm')}-IFERROR({self._r('kc', 'N')}*4/{self._r('aT', 'N')},0))*{S('grrb')}*(1-{S('grrb')}^{g('life')})/(1-{S('grrb')})", USD)
        self.outrow("ue", "o_tv", "Total value = cumulative P&L + embedded value", f"={self.out['o_cum']}+{self.out['o_emb']}", USD)
        return self.out


# ============================================================ MAIN (LIVE) MODEL across 3 sheets
def g_live(n):
    if n in M:
        return f"Assumptions!$G${M[n]}"
    return f"Assumptions!$G${A[n]}"


for ws_, t, s_ in [(fun, "GTM Funnel — live model (selected scenario and motion)", "Lead generation → qualification → opportunities → wins, with a sales-capacity constraint. Quarters Q1 2027–Q4 2029; annual columns P–R; total S."),
                   (rev, "Revenue Model — customers, ARR and recognized revenue", "ARR = prior ARR × NRR^(1/4) + new customers × ACV. Recognized revenue = average ARR ÷ 4 (kept separate from ARR). Customers decay at GRR^(1/4)."),
                   (ue, "Unit Economics — incremental cost, program P&L, CAC and payback", "Incremental to Clay's existing business: only costs created by the vertical program. CAC basis = sales & marketing; engineering and CS shown separately.")]:
    title(ws_, t, s_)
    put(ws_, "A3", '="Live run: "&Assumptions!$E$4&" scenario · "&Assumptions!$E$5', bold=True, border=False)
    ws_.column_dimensions["A"].width = 52
    ws_.column_dimensions["B"].width = 16
    for c in QCOLS:
        ws_.column_dimensions[c].width = 11
    ws_.column_dimensions["O"].width = 2
    for c in "PQRS":
        ws_.column_dimensions[c].width = 13
    ws_.freeze_panes = "C5"

LIVE = Engine({"fun": [fun, 5], "rev": [rev, 5], "ue": [ue, 5]}, g_live,
              lambda i: f"Assumptions!${L(4+i)}${HIRE['live']}", "Live model").build()


# ============================================================ ENGINE RUNS (scenarios + sensitivity)
title(eng, "Engine Runs — full copies of the model for each scenario and sensitivity case",
      "Generated by the same code as the live model (identical row logic). Each block lists its inputs at the top; yellow = input changed for that run.")
eng.column_dimensions["A"].width = 52
eng.column_dimensions["B"].width = 16
for c in QCOLS:
    eng.column_dimensions[c].width = 11
for c in "PQRS":
    eng.column_dimensions[c].width = 13
eng_row = [4]


def make_get(scol, mcol, overrides):
    def g(n):
        if n in overrides:
            return overrides[n]
        if n in M:
            return f"Assumptions!${mcol}${M[n]}"
        return f"Assumptions!${scol}${A[n]}"
    return g


def run(label, scen=1, motion=0, overrides=None, hire_row=None):
    overrides = overrides or {}
    scol = "DEF"[scen]
    mcol = "DE"[motion]
    hr = HIRE[motion] if hire_row is None else hire_row
    start = eng_row[0]
    put(eng, f"A{start}", "RUN: " + label, bold=True, fill=YFILL if overrides else KFILL)
    eng.merge_cells(f"A{start}:N{start}")
    eng_row[0] += 1
    if overrides:
        put(eng, f"A{eng_row[0]}", "Overrides: " + "; ".join(f"{k} = {v}" for k, v in overrides.items()), italic=True, border=False)
        eng.merge_cells(f"A{eng_row[0]}:N{eng_row[0]}")
        eng_row[0] += 1
    e = Engine({"fun": [eng, eng_row[0]], "rev": [eng, 0], "ue": [eng, 0]}, make_get(scol, mcol, overrides),
               lambda i: f"Assumptions!${L(4+i)}${hr}", label)

    # chain stages on one sheet: run funnel, then rev, then ue sequentially
    class Chain(dict):
        pass
    st = e.stages
    st["rev"] = st["fun"]
    st["ue"] = st["fun"]
    out = e.build()
    eng_row[0] = st["fun"][1] + 2
    out["_reg"] = e.reg
    return out


RUNS = {}
RUNS["D"] = run("Partner-led pilot — Downside", 0, 0)
RUNS["B"] = run("Partner-led pilot — Base", 1, 0)
RUNS["U"] = run("Partner-led pilot — Upside", 2, 0)
RUNS["Dg"] = run("Partner-led pilot — Downside, stopped at Q4 2027 gate", 0, 0, {"stop": 4})
RUNS["Us"] = run("Partner-led pilot — Upside, with scale-up hires after gate", 2, 0, hire_row=HIRE["scale"])
RUNS["SB"] = run("Sales-led vertical team — Base", 1, 1)
RUNS["SU"] = run("Sales-led vertical team — Upside", 2, 1)

# Sensitivity runs (Base scenario, partner-led): one-at-a-time Downside / Upside values
TORNADO = [
    ("Outbound meeting rate", ["meet"]),
    ("Inbound / template lead rate", ["inb"]),
    ("Partner network (start, adds, cap, leads/partner)", ["p0", "padd", "pcap", "lpp"]),
    ("Qualification rates", ["qual", "qualp"]),
    ("SQL → opportunity", ["s2o"]),
    ("Win rate", ["win"]),
    ("Usage intensity (ACV)", ["usage"]),
    ("Gross revenue retention", ["grrS", "grrM", "grrE"]),
    ("Net revenue retention", ["nrrS", "nrrM", "nrrE"]),
    ("Gross margin", ["gm"]),
    ("Wave timing (waves 2 and 3)", ["w2", "w3"]),
    ("Partner fee and onboarding subsidy", ["ref", "onb"]),
    ("Self-serve share of SMB", ["ss"]),
]
TRUNS = []
for nm, keys in TORNADO:
    lo = run(f"Sensitivity — {nm}: Downside value", 1, 0, {k: f"Assumptions!$D${A[k]}" for k in keys})
    hi = run(f"Sensitivity — {nm}: Upside value", 1, 0, {k: f"Assumptions!$F${A[k]}" for k in keys})
    TRUNS.append((nm, keys, lo, hi))
WINS = [0.15, 0.20, 0.25, 0.28, 0.30, 0.35, 0.40, 0.50]
WRUNS = [run(f"Break-even — win rate {w:.0%}", 1, 0, {"win": w}) for w in WINS]
GW = [0.20, 0.28, 0.35, 0.45]
GU = [0.75, 1.0, 1.5, 2.0]
GRUNS = {(w, u): run(f"Grid — win {w:.0%} × usage {u}x", 1, 0, {"win": w, "usage": u}) for w in GW for u in GU}

# ============================================================ SCENARIOS sheet
title(scn, "Scenarios — partner-led pilot (Downside / Base / Upside), gated downside, and the sales-led alternative",
      "Every figure links to a full model run on 'Engine Runs'. Scenario outputs are illustrative, not Clay forecasts.")
cols = [("D", "Downside", "D"), ("Dg", "Downside — stopped at Q4 2027 gate", "E"), ("B", "Base", "F"), ("U", "Upside", "G"), ("Us", "Upside + scale-up", "H"), ("SB", "Sales-led team — Base", "I"), ("SU", "Sales-led team — Upside", "J")]
hdr(scn, 4, ["Output", "Unit", "", "Partner-led: Downside", "Partner-led: Downside, gated at Q4 2027", "Partner-led: Base", "Partner-led: Upside (shared AEs only)", "Partner-led: Upside + 2 AEs hired after gate", "Sales-led team: Base", "Sales-led team: Upside"],
    widths=[52, 10, 2, 16, 18, 16, 17, 18, 16, 16])
OUTS = [("o_exit", "Exit ARR, Q4 2029", "$", USD), ("o_cust", "Customers, Q4 2029", "#", N0), ("o_newc", "New customers 2027–29", "#", N0),
        ("o_rev", "Recognized revenue 2027–29", "$", USD), ("o_acv", "Average first-year ACV", "$", USD), ("o_cac", "CAC", "$", USD),
        ("o_pb", "CAC payback (2027–29 average)", "months", MO), ("o_pb29", "CAC payback (2029 cohort)", "months", MO), ("o_ltv", "LTV : CAC", "x", X2),
        ("o_cost", "Total incremental cost 2027–29", "$", USD), ("o_cum", "Cumulative program P&L 2027–29", "$", USD),
        ("o_peak", "Peak investment (lowest cumulative P&L)", "$", USD), ("o_be", "First quarter with positive program P&L (99 = none)", "#", N0),
        ("o_rr", "Q4 2029 program P&L, annualized", "$", USD), ("o_emb", "Embedded value of 2029 ARR base", "$", USD),
        ("o_tv", "Total value (cumulative P&L + embedded value)", "$", USD)]
SROW = {}
for i, (k, lab, unit, fmt) in enumerate(OUTS):
    r = 5 + i
    SROW[k] = r
    put(scn, f"A{r}", lab, bold=(k in ("o_exit", "o_pb", "o_cum", "o_tv")))
    put(scn, f"B{r}", unit, center=True)
    for key, _, col in cols:
        put(scn, f"{col}{r}", f"={RUNS[key][k]}", fmt=fmt, font=GREEN, center=True, fill=KFILL if key == "B" else None)
r = 5 + len(OUTS) + 1
put(scn, f"A{r}", "Reconciliation: live model vs matching engine run (should be 0)", bold=True, fill=SFILL)
for c in "BCDEFGHIJ":
    put(scn, f"{c}{r}", None, fill=SFILL)
r += 1
put(scn, f"A{r}", "Exit ARR difference")
put(scn, f"D{r}", f'=IF({SEL_M}=1,{LIVE["o_exit"]}-INDEX(D{SROW["o_exit"]}:G{SROW["o_exit"]},1,CHOOSE({SEL_S},1,3,4)),IF({SEL_S}=2,{LIVE["o_exit"]}-I{SROW["o_exit"]},IF({SEL_S}=3,{LIVE["o_exit"]}-J{SROW["o_exit"]},"n/a")))', fmt=USD, center=True)
RECON1 = r
r += 1
put(scn, f"A{r}", "Cumulative P&L difference")
put(scn, f"D{r}", f'=IF({SEL_M}=1,{LIVE["o_cum"]}-INDEX(D{SROW["o_cum"]}:G{SROW["o_cum"]},1,CHOOSE({SEL_S},1,3,4)),IF({SEL_S}=2,{LIVE["o_cum"]}-I{SROW["o_cum"]},IF({SEL_S}=3,{LIVE["o_cum"]}-J{SROW["o_cum"]},"n/a")))', fmt=USD, center=True)
RECON2 = r
r += 2
# Decision analysis
put(scn, f"A{r}", "Decision analysis — is the pilot worth running? (probabilities are judgment [D])", bold=True, fill=SFILL)
for c in "BCDEFGHI":
    put(scn, f"{c}{r}", None, fill=SFILL)
r += 1
hdr(scn, r, ["Item", "", "", "Downside", "Base", "Upside", "Expected value", "", ""])
r += 1
put(scn, f"A{r}", "Probability weight")
for col, nm in zip("DEF", ["pD", "pB", "pU"]):
    put(scn, f"{col}{r}", f"=Assumptions!$E${A[nm]}", fmt=PCT, font=GREEN, center=True)
PR = r
r += 1
put(scn, f"A{r}", "Total value — ungated (run all 12 quarters)")
put(scn, f"D{r}", f"=D{SROW['o_tv']}", fmt=USD, center=True)
put(scn, f"E{r}", f"=F{SROW['o_tv']}", fmt=USD, center=True)
put(scn, f"F{r}", f"=G{SROW['o_tv']}", fmt=USD, center=True)
put(scn, f"G{r}", f"=SUMPRODUCT(D{PR}:F{PR},D{r}:F{r})", fmt=USD, center=True, bold=True)
EV1 = r
r += 1
put(scn, f"A{r}", "Total value — staged (Downside stopped at Q4 2027 gate; Upside adds 2 AEs)")
put(scn, f"D{r}", f"=E{SROW['o_tv']}", fmt=USD, center=True)
put(scn, f"E{r}", f"=F{SROW['o_tv']}", fmt=USD, center=True)
put(scn, f"F{r}", f"=H{SROW['o_tv']}", fmt=USD, center=True)
put(scn, f"G{r}", f"=SUMPRODUCT(D{PR}:F{PR},D{r}:F{r})", fmt=USD, center=True, bold=True, fill=KFILL)
EV2 = r
r += 1
put(scn, f"A{r}", "Cumulative P&L 2027–29 only — staged (excludes post-2029 value)")
put(scn, f"D{r}", f"=E{SROW['o_cum']}", fmt=USD, center=True)
put(scn, f"E{r}", f"=F{SROW['o_cum']}", fmt=USD, center=True)
put(scn, f"F{r}", f"=H{SROW['o_cum']}", fmt=USD, center=True)
put(scn, f"G{r}", f"=SUMPRODUCT(D{PR}:F{PR},D{r}:F{r})", fmt=USD, center=True, bold=True)
EV3 = r
r += 1
put(scn, f"A{r}", "Value of staging (staged EV − ungated EV)")
put(scn, f"G{r}", f"=G{EV2}-G{EV1}", fmt=USD, center=True, bold=True)
r += 1
put(scn, f"A{r}", "Break-even probability of Upside for staged EV = 0 (Downside:Base ratio held)")
xr = f"((D{PR}*D{EV2}+E{PR}*E{EV2})/(D{PR}+E{PR}))"
put(scn, f"G{r}", f'=IF(AND({xr}<0,F{EV2}>0),-{xr}/(F{EV2}-{xr}),IF({xr}>=0,0,"never"))', fmt=PCT, center=True, bold=True)
BEP = r
r += 1
put(scn, f"A{r}", "Sales-led team vs partner-led: difference in Base cumulative P&L")
put(scn, f"G{r}", f"=I{SROW['o_cum']}-F{SROW['o_cum']}", fmt=USD, center=True, bold=True)
SLDIFF = r
r += 2
note(scn, f"A{r}:J{r}", "Reading: in Base the binding constraint is demand (pipeline), not sales capacity, so a dedicated sales team produces the same ARR as the partner-led pilot at much higher cost. Only in Upside does the borrowed AE capacity bind (from Q3 2028), so dedicated AEs should be hired on that evidence, after the gate, not up front. The gate assumes Downside is distinguishable from Base by Q4 2027 on leading indicators (win rate, ACV, partner leads); if it is not, use the ungated row.", 58)
# ARR trajectories for chart
r += 2
put(scn, f"A{r}", "ARR trajectory by quarter (for charts)", bold=True, fill=SFILL)
for i, q in enumerate(QLAB):
    put(scn, f"{L(4+i)}{r}", q, bold=True, center=True, fill=SFILL)
TRJ = r
traj = [("Partner-led: Downside", "D"), ("Partner-led: Base", "B"), ("Partner-led: Upside", "U"), ("Partner-led: Upside + scale-up", "Us"), ("Sales-led: Base", "SB")]
for k, (lab, key) in enumerate(traj):
    rr = r + 1 + k
    put(scn, f"A{rr}", lab)
    regw, regr = RUNS[key]["_reg"]["aT"]
    for i in range(12):
        put(scn, f"{L(4+i)}{rr}", f"='Engine Runs'!{QCOLS[i]}{regr}", fmt=USD, font=GREEN)
r = r + 7
put(scn, f"A{r}", "Cumulative program P&L by quarter (for charts)", bold=True, fill=SFILL)
for i, q in enumerate(QLAB):
    put(scn, f"{L(4+i)}{r}", q, bold=True, center=True, fill=SFILL)
TRJ2 = r
traj2 = [("Partner-led: Downside", "D"), ("Partner-led: Downside, gated", "Dg"), ("Partner-led: Base", "B"), ("Partner-led: Upside + scale-up", "Us"), ("Sales-led: Base", "SB")]
for k, (lab, key) in enumerate(traj2):
    rr = r + 1 + k
    put(scn, f"A{rr}", lab)
    regw, regr = RUNS[key]["_reg"]["cum"]
    for i in range(12):
        put(scn, f"{L(4+i)}{rr}", f"='Engine Runs'!{QCOLS[i]}{regr}", fmt=USD, font=GREEN)
for c in "KLMNO":
    scn.column_dimensions[c].width = 12

# ============================================================ SENSITIVITY sheet
title(sen, "Sensitivity — which assumptions matter, and what has to be true",
      "Base scenario, partner-led pilot. Block A: one-at-a-time swings using each input's Downside / Upside value. Block B: minimum win rate. Block C: win rate × usage intensity grid.")
hdr(sen, 4, ["Input swung (one at a time)", "Downside value", "Upside value", "Total value @ Downside", "Total value @ Upside", "Swing (total value)",
             "Exit ARR @ Downside", "Exit ARR @ Upside", "Payback @ Downside", "Payback @ Upside"],
    widths=[44, 13, 13, 16, 16, 16, 15, 15, 13, 13])
for i, (nm, keys, lo, hi) in enumerate(TRUNS):
    r = 5 + i
    put(sen, f"A{r}", nm, bold=True)
    k0 = keys[0]
    fmt = PCT if k0 in pct_names else ('0.00' if k0 == "usage" else N0)
    put(sen, f"B{r}", f"=Assumptions!$D${A[k0]}", fmt=fmt, font=GREEN, center=True)
    put(sen, f"C{r}", f"=Assumptions!$F${A[k0]}", fmt=fmt, font=GREEN, center=True)
    put(sen, f"D{r}", f"={lo['o_tv']}", fmt=USD, center=True)
    put(sen, f"E{r}", f"={hi['o_tv']}", fmt=USD, center=True)
    put(sen, f"F{r}", f"=ABS(E{r}-D{r})", fmt=USD, center=True, bold=True)
    put(sen, f"G{r}", f"={lo['o_exit']}", fmt=USD, center=True)
    put(sen, f"H{r}", f"={hi['o_exit']}", fmt=USD, center=True)
    put(sen, f"I{r}", f"={lo['o_pb']}", fmt=MO, center=True)
    put(sen, f"J{r}", f"={hi['o_pb']}", fmt=MO, center=True)
TL = 5 + len(TRUNS) - 1
sen.conditional_formatting.add(f"F5:F{TL}", DataBarRule(start_type="num", start_value=0, end_type="max", color="4472C4"))
r = TL + 1
put(sen, f"A{r}", "Base case (for reference)", bold=True, fill=SFILL)
put(sen, f"D{r}", f"={RUNS['B']['o_tv']}", fmt=USD, center=True, fill=SFILL)
put(sen, f"E{r}", f"={RUNS['B']['o_tv']}", fmt=USD, center=True, fill=SFILL)
put(sen, f"G{r}", f"={RUNS['B']['o_exit']}", fmt=USD, center=True, fill=SFILL)
put(sen, f"I{r}", f"={RUNS['B']['o_pb']}", fmt=MO, center=True, fill=SFILL)
put(sen, f"A{r+1}", "Note: the rank of each input = its swing; the bar chart below plots the swing.", italic=True, border=False)
chart = BarChart()
chart.type = "bar"
chart.title = "Swing in total value (Downside ↔ Upside value of one input)"
chart.add_data(Reference(sen, min_col=6, min_row=4, max_row=TL), titles_from_data=True)
chart.set_categories(Reference(sen, min_col=1, min_row=5, max_row=TL))
chart.height, chart.width = 9, 18
chart.legend = None
chart.x_axis.delete = False
chart.y_axis.delete = False
sen.add_chart(chart, f"L4")
# Block B: win-rate break-even
bb = TL + 4
put(sen, f"A{bb}", "Block B — What win rate does the pilot need? (all else Base)", bold=True, border=False)
hdr(sen, bb + 1, ["Win rate", "New customers", "Exit ARR", "CAC payback (2027–29)", "CAC payback (2029 cohort)", "Cumulative P&L", "Total value", "Payback ≤ 24 mo (2029 cohort)?", "Total value ≥ 0?", ""])
for i, w in enumerate(WINS):
    rr = bb + 2 + i
    o = WRUNS[i]
    put(sen, f"A{rr}", w, font=BLUE, fmt=PCT, center=True)
    for col, k, fmt in [("B", "o_newc", N0), ("C", "o_exit", USD), ("D", "o_pb", MO), ("E", "o_pb29", MO), ("F", "o_cum", USD), ("G", "o_tv", USD)]:
        put(sen, f"{col}{rr}", f"={o[k]}", fmt=fmt, center=True)
    put(sen, f"H{rr}", f'=IF(AND(E{rr}>0,E{rr}<=24),"Yes","No")', center=True)
    put(sen, f"I{rr}", f'=IF(G{rr}>=0,"Yes","No")', center=True)
WL = bb + 2 + len(WINS) - 1
sen.conditional_formatting.add(f"H{bb+2}:I{WL}", CellIsRule(operator="equal", formula=['"Yes"'], fill=KFILL))
put(sen, f"A{WL+1}", "Lowest tested win rate with 2029-cohort payback ≤ 24 months", bold=True)
put(sen, f"E{WL+1}", f'=IFERROR(INDEX(A{bb+2}:A{WL},MATCH("Yes",H{bb+2}:H{WL},0)),"none tested")', fmt=PCT, center=True, bold=True, fill=KFILL)
put(sen, f"A{WL+2}", "Lowest tested win rate with total value ≥ 0", bold=True)
put(sen, f"E{WL+2}", f'=IFERROR(INDEX(A{bb+2}:A{WL},MATCH("Yes",I{bb+2}:I{WL},0)),"none tested")', fmt=PCT, center=True, bold=True, fill=KFILL)
WBE1, WBE2 = WL + 1, WL + 2
# Block C: grid
gc = WL + 5
put(sen, f"A{gc}", "Block C — CAC payback, 2029 cohort (months): win rate (rows) × usage intensity (columns)", bold=True, border=False)
put(sen, f"A{gc+1}", "Win rate ↓ / usage →", bold=True, fill=SFILL)
for j, u in enumerate(GU):
    put(sen, f"{L(2+j)}{gc+1}", u, font=BLUE, fmt='0.00"x"', center=True, fill=SFILL)
for i, w in enumerate(GW):
    rr = gc + 2 + i
    put(sen, f"A{rr}", w, font=BLUE, fmt=PCT, center=True, fill=SFILL)
    for j, u in enumerate(GU):
        put(sen, f"{L(2+j)}{rr}", f"={GRUNS[(w, u)]['o_pb29']}", fmt=MO, center=True)
sen.conditional_formatting.add(f"B{gc+2}:E{gc+5}", ColorScaleRule(start_type="num", start_value=12, start_color="C6E0B4", mid_type="num", mid_value=24, mid_color="FFF2CC", end_type="num", end_value=60, end_color="F8CBAD"))
gc2 = gc + 8
put(sen, f"A{gc2}", "Block C2 — Total value ($): same grid", bold=True, border=False)
put(sen, f"A{gc2+1}", "Win rate ↓ / usage →", bold=True, fill=SFILL)
for j, u in enumerate(GU):
    put(sen, f"{L(2+j)}{gc2+1}", u, font=BLUE, fmt='0.00"x"', center=True, fill=SFILL)
for i, w in enumerate(GW):
    rr = gc2 + 2 + i
    put(sen, f"A{rr}", w, font=BLUE, fmt=PCT, center=True, fill=SFILL)
    for j, u in enumerate(GU):
        put(sen, f"{L(2+j)}{rr}", f"={GRUNS[(w, u)]['o_tv']}", fmt=USD, center=True)
sen.conditional_formatting.add(f"B{gc2+2}:E{gc2+5}", CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=KFILL))
sen.conditional_formatting.add(f"B{gc2+2}:E{gc2+5}", CellIsRule(operator="lessThan", formula=["0"], fill=PatternFill("solid", fgColor="F8CBAD")))
note(sen, f"A{gc2+7}:J{gc2+7}", "Reading: win rate alone cannot rescue the economics — at Base usage, even a 50% win rate leaves 2029-cohort payback above 24 months (Block B). Payback ≤ 24 months and positive total value appear only toward the bottom-right of the grid: win rates of 35%+ combined with usage ~1.5–2× the base build (i.e. mid-market ACVs of ~$15–20K). ACV per account is therefore the first thing the pilot must prove.", 40)

# ============================================================ EXECUTIVE OUTPUT
title(exo, "Executive Output — three-year business case for the distribution pilot (2027–2029)",
      "Scenario outputs are illustrative estimates built from public data and labeled assumptions. They are not Clay forecasts and do not use Clay internal data.")
put(exo, "A3", '="Live model currently set to: "&Assumptions!$E$4&" · "&Assumptions!$E$5&"  (change on Assumptions D4:D5)"', bold=True, border=False)
hdr(exo, 5, ["Metric", "Partner-led: Downside", "Partner-led: Downside, gated", "Partner-led: Base", "Partner-led: Upside", "Partner-led: Upside + scale-up", "Sales-led team: Base", "Note"],
    widths=[50, 16, 17, 16, 16, 17, 16, 46])
kpis = [("o_exit", "Exit ARR, Q4 2029", USD, ""), ("o_newc", "New customers 2027–29", N0, ""), ("o_acv", "Average first-year ACV", USD, "Usage-based build"),
        ("o_cac", "CAC", USD, "Sales & marketing ÷ new customers"), ("o_pb", "CAC payback, 2027–29 average (months)", MO, "Benchmark median: SMB 8–12, MM 14–18 (S47)"),
        ("o_pb29", "CAC payback, 2029 cohort (months)", MO, "Steady-state view after launch costs"), ("o_ltv", "LTV : CAC", X2, ">3x is the usual bar [K]"),
        ("o_cost", "Total incremental cost 2027–29", USD, ""), ("o_cum", "Cumulative program P&L 2027–29", USD, ""),
        ("o_peak", "Peak investment required", USD, "Lowest cumulative P&L"), ("o_rr", "Q4 2029 program P&L, annualized", USD, "Exit run-rate"),
        ("o_tv", "Total value incl. embedded value of ARR base", USD, "Cumulative P&L + 5-yr GP of 2029 base")]
for i, (k, lab, fmt, nt) in enumerate(kpis):
    r = 6 + i
    put(exo, f"A{r}", lab, bold=k in ("o_exit", "o_pb29", "o_cum", "o_tv"))
    for col, sc in zip("BCDEFG", ["D", "E", "F", "G", "H", "I"]):
        put(exo, f"{col}{r}", f"=Scenarios!{sc}{SROW[k]}", fmt=fmt, font=GREEN, center=True, fill=KFILL if col == "D" else None)
    put(exo, f"H{r}", nt)
r = 6 + len(kpis) + 1
put(exo, f"A{r}", "Decision metrics", bold=True, fill=SFILL)
for c in "BCDEFGH":
    put(exo, f"{c}{r}", None, fill=SFILL)
r += 1
dm = [("Expected total value — staged pilot (probability-weighted)", f"=Scenarios!G{EV2}", USD, "Weights 35/45/20 (Assumptions, judgment)"),
      ("Expected total value — ungated", f"=Scenarios!G{EV1}", USD, ""),
      ("Expected cumulative P&L 2027–29 — staged (no post-2029 value)", f"=Scenarios!G{EV3}", USD, "Every scenario is cash-negative through 2029; value depends on the ARR base after 2029"),
      ("Break-even probability of Upside (staged EV = 0)", f"=Scenarios!G{BEP}", PCT, "Pilot is worth running if P(Upside) exceeds this"),
      ("Sales-led team minus partner-led, Base cumulative P&L", f"=Scenarios!G{SLDIFF}", USD, "Same ARR; extra cost only"),
      ("Lowest tested win rate for ≤ 24-month payback (2029 cohort, Base usage)", f"=Sensitivity!E{WBE1}", PCT, "Benchmarks: 21% avg; 31–39% opp→won"),
      ("Lowest tested win rate for total value ≥ 0", f"=Sensitivity!E{WBE2}", PCT, ""),
      ("Model checks (0 = pass)", f"=ABS(Scenarios!D{RECON1})+ABS(Scenarios!D{RECON2})", USD, "Live model reconciles to engine runs")]
for lab, f, fmt, nt in dm:
    put(exo, f"A{r}", lab)
    put(exo, f"B{r}", f, fmt=fmt, font=GREEN, center=True, bold=True)
    put(exo, f"H{r}", nt)
    r += 1
r += 1
put(exo, f"A{r}", "What the model shows", bold=True, fill=SFILL)
for c in "BCDEFGH":
    put(exo, f"{c}{r}", None, fill=SFILL)
r += 1
msgs = [
    "1. The distribution vertical is demand-limited, not capacity-limited. A dedicated sales team produces the same ARR as the partner-led pilot in Base, at materially higher cost, so Option 1 is rejected on economics (H7 supported).",
    "2. In the Base case the partner-led pilot does not pay back within 2027–29. The serviceable market is ~5,200 firms, and list-price ACVs are ~$5K (SMB), ~$10K (Mid-Market) and ~$62K (Enterprise).",
    "3. The Upside case (higher win rates, usage and retention) pays back and is profitable at its 2029 exit run-rate; there, borrowed AE capacity binds from Q3 2028, which is the evidence-based trigger to hire two dedicated AEs. The pilot is therefore an option: a capped, gateable investment with an asymmetric payoff.",
    "4. A Q4 2027 Go/No-Go gate cuts the Downside loss by more than half. The sensitivity is led by usage intensity (ACV), then funnel conversion (win rate, SQL→opportunity, qualification), gross margin and partner productivity; these become the gate metrics.",
    "5. Implication for strategy: run the pilot lean (≈1 lead + 1 engineer + shared AEs + partners), focus on 100+ employee distributors, and do not hire a vertical sales team unless the gate metrics are met.",
]
for t in msgs:
    put(exo, f"A{r}", t)
    exo.merge_cells(f"A{r}:H{r}")
    exo.row_dimensions[r].height = 34
    r += 1
# charts
ch1 = LineChart()
ch1.title = "ARR by quarter, 2027–2029"
ch1.y_axis.title = "ARR ($)"
ch1.y_axis.numFmt = '$#,##0,"K"'
for k in range(len(traj)):
    s = Series(Reference(scn, min_col=4, max_col=15, min_row=TRJ + 1 + k), title=traj[k][0])
    ch1.series.append(s)
ch1.set_categories(Reference(scn, min_col=4, max_col=15, min_row=TRJ))
ch1.height, ch1.width = 9, 17
ch1.x_axis.delete = False
ch1.y_axis.delete = False
exo.add_chart(ch1, f"A{r+1}")
ch2 = LineChart()
ch2.title = "Cumulative program P&L (investment position)"
ch2.y_axis.numFmt = '$#,##0,"K"'
for k in range(len(traj2)):
    s = Series(Reference(scn, min_col=4, max_col=15, min_row=TRJ2 + 1 + k), title=traj2[k][0])
    ch2.series.append(s)
ch2.set_categories(Reference(scn, min_col=4, max_col=15, min_row=TRJ2))
ch2.height, ch2.width = 9, 17
ch2.x_axis.delete = False
ch2.y_axis.delete = False
exo.add_chart(ch2, f"E{r+1}")

# Market sizing reconciliation (needs Scenarios rows)
for i, (lab, p1, p4, why) in enumerate(recon):
    rr = RECON_START + i
    put(ms, f"A{rr}", lab, bold=True)
    ms.merge_cells(f"A{rr}:B{rr}")
    put(ms, f"C{rr}", p1, font=BLUE, fmt=USD if "ACV" in lab or "ARR" in lab else N0, center=True)
    f = p4.format(acvS=A["acvS"], acvM=A["acvM"], acvE=A["acvE"], exitarr=SROW["o_exit"])
    if "{" not in p4 and p4.startswith("=Scenarios") is False:
        f = p4
    put(ms, f"D{rr}", f.replace("$D$", "$F$") if "Scenarios" in f else f, fmt=USD if "ACV" in lab or "ARR" in lab else N0, center=True, font=GREEN)
    put(ms, f"E{rr}", why)
    ms.merge_cells(f"E{rr}:L{rr}")

# ============================================================ COVER
title(cover, "07 — Financial Model: three-year business case for a U.S. industrial-distribution pilot",
      "Independent portfolio case study by Zhuoran (Zora) Wang · Not commissioned by or affiliated with Clay Labs Inc. · Research reference date: 2026-10-18")
rows = [
    ("Question", "Does the Phase 3 recommendation (enter industrial distribution through a staged, partner-led pilot) create value over 2027–2029, and how should it be sized and gated?"),
    ("Structure", "Assumptions → Market Sizing → ACV Build → GTM Funnel → Revenue Model → Unit Economics (live model, selectors on Assumptions D4:D5) → Scenarios and Sensitivity (full model copies on Engine Runs) → Executive Output."),
    ("Time and units", "Quarterly, Q1 2027–Q4 2029; annual totals in columns P–R and 2027–29 in S. US$ nominal. ARR is end-of-quarter run-rate; recognized revenue = average ARR ÷ 4 (kept separate)."),
    ("Scope of costs", "Incremental to Clay only: vertical team, variable sales cost, programs, partner fees, enablement, onboarding subsidies, template/connector build, customer success. COGS via gross margin. No allocation of Clay's existing overhead."),
    ("Evidence classes", "A = verified public fact · B = external benchmark · C = analyst estimate · D = scenario assumption · E = recommendation. Every input row on Assumptions carries its class and source."),
    ("Color legend", "Blue = input · black = formula · green = link to another tab · yellow fill = judgment assumption · green fill = key output."),
    ("Not a forecast", "Outputs are scenario-based estimates. They are not Clay forecasts and imply no access to Clay internal financials. Probability weights in the decision analysis are explicit judgments."),
]
for i, (k, v) in enumerate(rows, start=4):
    put(cover, f"A{i}", k, bold=True, fill=SFILL)
    put(cover, f"B{i}", v)
    cover.row_dimensions[i].height = 44
cover.column_dimensions["A"].width = 18
cover.column_dimensions["B"].width = 125

# ============================================================ SOURCES
title(src, "Sources", "Full register with dates, reliability and limitations: Research_Sources.csv")
hdr(src, 4, ["ID", "Source", "Publisher", "Date / reference", "Used for", "URL"], widths=[7, 52, 28, 22, 44, 70])
SRCS = [
    ("S01", "Clay pricing page (plans, Actions & Data Credits packs, Enterprise minimums)", "Clay Labs Inc.", "Live, Oct 2026", "ACV Build prices and floors", "https://www.clay.com/pricing"),
    ("S02", "Clay tender offer release (enterprise NRR > 200%)", "Business Wire", "2026-01-28", "NRR context", "https://www.businesswire.com/news/home/20260128514638/en"),
    ("S14", "2022 Economic Census — Employment Size of Firms", "U.S. Census Bureau", "Ref. year 2022", "Firm counts, employees per firm", "https://data.census.gov/table/ECNSIZE2022.EC2200SIZEEMPFIRM"),
    ("S15", "OEWS May 2025 industry-specific estimates (sales share)", "U.S. BLS", "May 2025", "Sellers per firm", "https://data.bls.gov/oes/#/industry/423800/2025"),
    ("S16", "Clay buyer pricing data", "Vendr", "Feb 2026", "Enterprise discount, ACV check", "https://www.vendr.com/marketplace/clay"),
    ("S17", "State of CRM in Distribution", "Distribution Strategy Group", "2022", "CRM-ready filter", "https://archive.distributionstrategy.com/wp-content/uploads/2022/10/DSG-The-State-of-CRM-in-Distribution-Making-the-Software-Work-for-You.pdf"),
    ("S43", "2025 B2B SaaS Retention Benchmarks", "SaaS Capital", "2025-09", "GRR / NRR by ACV band", "https://www.saas-capital.com/wp-content/uploads/2025/09/RB32WS1-2025-B2B-SaaS-Retention-Benchmarks.pdf"),
    ("S44", "2026 AE Models, Motions & Metrics", "The Bridge Group", "2026-06-22", "AE quota, OTE, ramp, attainment", "https://www.bridgegroupinc.com/research/2026-ae-models-motions-metrics"),
    ("S45", "Employer Costs for Employee Compensation, June 2026", "U.S. BLS (TED)", "2026-09", "Wage share 70.0% → cost loading", "https://www.bls.gov/opub/ted/2026/compensation-costs-for-private-industry-workers-averaged-46-89-per-hour-worked-in-june-2026.htm"),
    ("S46", "OEWS May 2025 — Software Publishers (513200) wages", "U.S. BLS", "May 2025", "Role wage rates", "https://data.bls.gov/oes/#/industry/513200/2025"),
    ("S47", "CAC payback benchmarks (compilation of Maxio, Benchmarkit, MetricHQ)", "SaaS Hero", "Updated 2026-07-18", "Payback benchmarks", "https://www.saashero.net/strategy/maxio-cac-benchmarks-b2b-saas/"),
    ("S48", "Win rate statistics (HubSpot 2024; Ebsta/Pavilion 2024; Norwest 2025)", "Webtonic (compilation)", "2026", "Win rate, MQL→SQL, cycle", "https://www.webtonic.io/blog/win-rate-statistics"),
    ("S49", "Pipeline performance benchmarks 2025", "The Digital Bloom (compilation)", "2025-10-22", "Funnel stage conversion", "https://thedigitalbloom.com/learn/pipeline-performance-benchmarks-2025/"),
    ("S50", "What margin should we give partners? (Freshworks, Bessemer cited)", "Echai Ventures", "Undated", "Partner fees, services subsidy", "https://echai.ventures/gtm/gtm-channel-ecosystem/what-margin-should-we-give-partners-and-how-do-we-stop-channel-conflict"),
    ("S51", "Clay pricing guide 2026 (credits per provider)", "Abmatic AI", "2026-09-01", "Credits per lookup", "https://abmatic.ai/blog/clay-pricing-guide-2026"),
    ("S52", "Clay company profile (margin structure, ARR estimate)", "Sacra", "2026", "Gross-margin context", "https://sacra.com/c/clay/"),
]
for i, s in enumerate(SRCS, start=5):
    for j, v in enumerate(s, start=1):
        put(src, f"{L(j)}{i}", v)

order = ["Cover", "Executive Output", "Assumptions", "Market Sizing", "ACV Build", "GTM Funnel", "Revenue Model", "Unit Economics", "Scenarios", "Sensitivity", "Engine Runs", "Sources"]
wb._sheets = [wb[n] for n in order]
for w in wb.worksheets:
    w.sheet_properties.tabColor = {"Cover": "1F2A44", "Executive Output": "1F2A44", "Assumptions": "FFC000", "Engine Runs": "A6A6A6", "Sources": "A6A6A6"}.get(w.title, "4472C4")
wb.save(OUT)
print("saved", OUT, "engine rows", eng_row[0], "runs", 6 + 2 * len(TRUNS) + len(WRUNS) + len(GRUNS))
