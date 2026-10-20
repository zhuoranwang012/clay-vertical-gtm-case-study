"""Build 06_Segment_Prioritization.xlsx — Clay vertical expansion case study (Phase 3).
Metric inputs are copied (blue) from 02_Market_Sizing.xlsx and 04_Competitive_Analysis.xlsx with cell references;
every score, total, rank and sensitivity is a live formula."""
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule, FormulaRule, ColorScaleRule
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter as L

BASE = "/home/claude/Clay_Strategy_Consulting_Case"
OUT = f"{BASE}/06_Segment_Prioritization.xlsx"
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
WIN = PatternFill("solid", fgColor="C6E0B4")
thin = Side(style="thin", color="BFBFBF")
BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
USD = '$#,##0;($#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
D2 = '0.00'
X2 = '0.00"x"'
SEGS = ["A. Industrial Distribution", "B. Industrial Manufacturing", "C. B2B Prof. Services"]
SHORT = ["A. Distribution", "B. Manufacturing", "C. Prof. services"]

wb = Workbook()


def hdr(ws, row, labels, col=1, widths=None, height=30):
    for i, t in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=t)
        c.font, c.fill, c.alignment, c.border = HDR, HFILL, CENTER, BOX
    if widths:
        for i, w in enumerate(widths):
            ws.column_dimensions[L(col + i)].width = w
    ws.row_dimensions[row].height = height


def put(ws, ref, val, font=BLACK, fmt=None, fill=None, bold=False, wrap=True, border=True, center=False):
    c = ws[ref]
    c.value = val
    c.font = Font(name=F, size=10, bold=bold, color=font.color)
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    c.alignment = CENTER if center else (WRAP if wrap else Alignment(vertical="top"))
    if border:
        c.border = BOX
    return c


def title(ws, t, s):
    ws["A1"].value, ws["A1"].font = t, TITLE
    ws["A2"].value, ws["A2"].font = s, SUB
    ws.sheet_view.showGridLines = False


def note(ws, ref_range, text, height=36):
    a = ref_range.split(":")[0]
    put(ws, a, text, border=False)
    ws.merge_cells(ref_range)
    ws.row_dimensions[int("".join(ch for ch in a if ch.isdigit()))].height = height


# ======================================================================= Cover
ws = wb.active
ws.title = "Cover"
title(ws, "06 — Segment Prioritization: Which vertical deserves Clay's next GTM dollar?",
      "Independent portfolio case study by Zhuoran (Zora) Wang · Not commissioned by or affiliated with Clay Labs Inc. · Research reference date: 2026-10-18")
rows = [
    ("Decision", "Rank three candidate U.S. verticals for incremental go-to-market investment and choose how to enter: no vertical bet, full vertical build, or staged pilot."),
    ("Framework", "Six criteria with the suggested weights (Market Attractiveness 20%, Customer Need 20%, Product-Market Fit 20%, Revenue Potential 15%, Competitive Position 15%, Ease of Execution 10%). The weights are analytical assumptions, not management preferences."),
    ("Scoring rule", "Quantitative metrics → 1–5 by ratio-to-max: score = 1 + 4 × value ÷ max(value across segments). This preserves proportional gaps; a min-max alternative is tested on Stress_Tests. Judgment scores (PMF, and Competitive Position / Ease from Phase 2) carry written evidence."),
    ("Inputs", "Copied from 02_Market_Sizing.xlsx and 04_Competitive_Analysis.xlsx (cell references noted) so this file recalculates on its own. Blue = input · black = formula · yellow = judgment score or assumption."),
    ("Robustness", "Five alternative weight sets; break-even weight per criterion; 2 × 1,000 random weight vectors (seeded, stored as values); 11 input stress tests; scoring-rule test."),
    ("Not a forecast", "Scores rank options under stated assumptions. Revenue figures are scenario outputs from Phase 1, not Clay forecasts. No Clay internal data or interviews were used."),
    ("Tabs", "Executive_Summary → Weights → Inputs → Scorecard → Weight_Sensitivity → Random_Weights → Stress_Tests → Subsegment_Drilldown → Options_Tradeoffs → Change_Conditions → Sources"),
]
for i, (k, v) in enumerate(rows, start=4):
    put(ws, f"A{i}", k, bold=True, fill=SFILL)
    put(ws, f"B{i}", v)
    ws.row_dimensions[i].height = 44
ws.column_dimensions["A"].width = 18
ws.column_dimensions["B"].width = 125

# ======================================================================= Weights
wt = wb.create_sheet("Weights")
title(wt, "Criteria weights — analytical assumptions, not management preferences",
      "Base = suggested weights from the project brief. Alternative sets are used on Weight_Sensitivity. Each column must sum to 100%.")
CRIT = ["Market Attractiveness", "Customer Need", "Product-Market Fit", "Revenue Potential", "Competitive Position", "Ease of Execution"]
WSETS = [("Base (brief)", [.20, .20, .20, .15, .15, .10]),
         ("Equal", [1 / 6] * 6),
         ("Market-size first", [.30, .15, .15, .25, .10, .05]),
         ("Execution first", [.15, .15, .20, .10, .10, .30]),
         ("Differentiation first", [.15, .20, .15, .10, .30, .10]),
         ("Near-term revenue", [.15, .10, .20, .30, .10, .15])]
hdr(wt, 4, ["Criterion"] + [w[0] for w in WSETS] + ["What the criterion asks"], widths=[24] + [14] * 6 + [70])
ask = ["How large and reachable is the segment, and how much of it is still unpenetrated by Clay?",
       "How intense is the GTM pain, and how badly do standard tools serve it?",
       "How well does Clay's current product solve the core jobs without major new build?",
       "How much ARR could Clay realistically capture, and how large can each account become?",
       "Can Clay win against incumbents and alternatives, and keep winning?",
       "How easily can customers adopt Clay and can Clay sell and deliver into the segment?"]
for i, cname in enumerate(CRIT):
    r = 5 + i
    put(wt, f"A{r}", cname, bold=True)
    for j, (_, ws_) in enumerate(WSETS):
        put(wt, f"{L(2+j)}{r}", ws_[i], font=BLUE, fmt=PCT, fill=YFILL if j == 0 else None, center=True)
    put(wt, f"H{r}", ask[i])
put(wt, "A11", "Check: sum", bold=True, fill=SFILL)
for j in range(len(WSETS)):
    c = L(2 + j)
    put(wt, f"{c}11", f"=SUM({c}5:{c}10)", fmt=PCT, bold=True, fill=SFILL, center=True)
wt.conditional_formatting.add("B11:G11", FormulaRule(formula=["ABS(B11-1)>0.0001"], fill=PatternFill("solid", fgColor="F8CBAD")))
put(wt, "A13", "Sub-weights inside composite criteria", bold=True, fill=SFILL)
SUBW = [("Market Attractiveness: share on SAM size (rest on incremental headroom)", 0.5),
        ("Customer Need: share on sales intensity (rest on data-coverage gap)", 0.5),
        ("Revenue Potential: share on SOM ARR (rest on seller density / account expansion)", 0.5)]
for i, (n, v) in enumerate(SUBW):
    put(wt, f"A{14+i}", n)
    wt.merge_cells(f"A{14+i}:F{14+i}")
    put(wt, f"G{14+i}", v, font=BLUE, fmt=PCT, fill=YFILL, center=True)
W_BASE = "Weights!$B$5:$B$10"

# ======================================================================= Inputs
inp = wb.create_sheet("Inputs")
title(inp, "Metric inputs carried from Phases 1–2", "Blue values are copied from the named source cells so this workbook recalculates independently. Change them only if the upstream workbooks change.")
hdr(inp, 4, ["ID", "Metric", "Criterion", SEGS[0], SEGS[1], SEGS[2], "Unit", "Class", "Source (file › tab › cells)"],
    widths=[6, 46, 20, 16, 16, 16, 14, 8, 60])
METRICS = [
    ("M1", "SAM, Base case", "Market Attractiveness", [68163346.58, 75842616.73, 148227391.96], USD, "$ / yr", "C/D", "02_Market_Sizing › Executive_Summary › B10:D10"),
    ("M2", "Incremental headroom flag (5 = largely unpenetrated by Clay)", "Market Attractiveness", [5, 5, 2], D2, "1–5", "C", "04_Competitive_Analysis › Segment_Scoring › F6 row"),
    ("M3", "Sales intensity index (sales occupations share ÷ all industries)", "Customer Need", [2.54356, 0.31725, 0.71487], X2, "index", "C from A", "02_Market_Sizing › Segment_Profile › B22:D22"),
    ("M4", "Customer-persona LinkedIn coverage index vs Prof. services", "Customer Need", [0.62124, 0.97041, 1.0], X2, "index", "C", "04_Competitive_Analysis › Summary › B9:D9"),
    ("M5", "Product-market fit (judgment; see Scorecard evidence)", "Product-Market Fit", [4, 2, 4], D2, "1–5", "C", "This workbook › Scorecard (pain-to-capability map in 05_Customer_Needs_Assessment §3)"),
    ("M6", "SOM — year-3 ARR potential, Base case", "Revenue Potential", [1864123.57, 2149272.49, 4178480.91], USD, "$ / yr", "D", "02_Market_Sizing › Executive_Summary › B11:D11"),
    ("M7", "Implied sales staff per 20+ firm (usage / expansion driver)", "Revenue Potential", [27.70, 3.65, 9.32], '0.0', "people", "C from A", "02_Market_Sizing › Segment_Profile › B19:D19"),
    ("M8", "Competitive position score", "Competitive Position", [4.0, 3.3333, 2.3333], D2, "1–5", "C", "04_Competitive_Analysis › Summary › B5:D5"),
    ("M9", "Adoption / execution fit score", "Ease of Execution", [2.0, 1.5, 4.5], D2, "1–5", "C", "04_Competitive_Analysis › Summary › B6:D6"),
]
MROW = {}
for i, (mid, name, crit, vals, fmt, unit, cls, src) in enumerate(METRICS):
    r = 5 + i
    MROW[mid] = r
    put(inp, f"A{r}", mid, bold=True)
    put(inp, f"B{r}", name)
    put(inp, f"C{r}", crit)
    for j, v in enumerate(vals):
        put(inp, f"{L(4+j)}{r}", v, font=BLUE, fmt=fmt, center=True, fill=YFILL if mid in ("M2", "M5", "M8", "M9") else None)
    put(inp, f"G{r}", unit)
    put(inp, f"H{r}", cls)
    put(inp, f"I{r}", src)
    inp.row_dimensions[r].height = 30
note(inp, "A15:I15", "M4 enters inversely: a lower coverage index means a larger data gap, so data-gap score = 1 + 4 × MIN(M4) ÷ M4. M2 headroom, M5, M8 and M9 are already on the 1–5 scale and enter directly.")


# ======================================================================= Scoring engine (shared block builder)
def engine_block(ws, top, label, overrides=None, rule="ratio", weights=W_BASE, show_metrics=True):
    """Writes a metrics block + criteria block + total + rank. Returns dict of row refs.
    Columns: A label, B description/ID, C/D/E = segments A/B/C."""
    overrides = overrides or {}
    r = top
    put(ws, f"A{r}", label, bold=True, fill=SFILL)
    for col in "BCDE":
        put(ws, f"{col}{r}", None, fill=SFILL)
    r += 1
    met = {}
    # metric rows
    for mid, name, crit, vals, fmt, unit, cls, src in METRICS:
        put(ws, f"A{r}", f"{mid} {name}")
        put(ws, f"B{r}", "input")
        for j in range(3):
            col = L(3 + j)
            ov = overrides.get((mid, j))
            if ov is not None:
                put(ws, f"{col}{r}", ov, font=BLUE, fmt=fmt, fill=YFILL, center=True)
            else:
                put(ws, f"{col}{r}", f"=Inputs!{L(4+j)}{MROW[mid]}", font=GREEN, fmt=fmt, center=True)
        met[mid] = r
        r += 1
    sw = {}
    for k, (n, v) in enumerate(SUBW):
        key = ["SW_MA", "SW_N", "SW_R"][k]
        put(ws, f"A{r}", "Sub-weight: " + n.split(":")[0] + " (" + n.split(":")[1].strip().split(" (")[0] + ")")
        put(ws, f"B{r}", "sub-weight")
        ov = overrides.get(key)
        put(ws, f"C{r}", ov if ov is not None else f"=Weights!$G${14+k}", font=BLUE if ov is not None else GREEN, fmt=PCT, center=True,
            fill=YFILL if ov is not None else None)
        sw[key] = f"$C${r}"
        r += 1
    if not show_metrics:
        for rr in range(top + 1, r):
            ws.row_dimensions[rr].hidden = True

    def rng(mid):
        return f"$C${met[mid]}:$E${met[mid]}"

    def sc(mid, col, inverse=False):
        cell = f"{col}{met[mid]}"
        if rule == "ratio":
            return f"(1+4*MIN({rng(mid)})/{cell})" if inverse else f"(1+4*{cell}/MAX({rng(mid)}))"
        # min-max
        if inverse:
            return f"(1+4*(1/{cell}-1/MAX({rng(mid)}))/(1/MIN({rng(mid)})-1/MAX({rng(mid)})))"
        return f"(1+4*({cell}-MIN({rng(mid)}))/(MAX({rng(mid)})-MIN({rng(mid)})))"

    crit_rows = {}
    for k, cname in enumerate(CRIT):
        put(ws, f"A{r}", f"Score — {cname}", bold=True)
        put(ws, f"B{r}", "1–5")
        for j in range(3):
            col = L(3 + j)
            if cname == "Market Attractiveness":
                f = f"={sw['SW_MA']}*{sc('M1', col)}+(1-{sw['SW_MA']})*{col}{met['M2']}"
            elif cname == "Customer Need":
                f = f"={sw['SW_N']}*{sc('M3', col)}+(1-{sw['SW_N']})*{sc('M4', col, inverse=True)}"
            elif cname == "Product-Market Fit":
                f = f"={col}{met['M5']}"
            elif cname == "Revenue Potential":
                f = f"={sw['SW_R']}*{sc('M6', col)}+(1-{sw['SW_R']})*{sc('M7', col)}"
            elif cname == "Competitive Position":
                f = f"={col}{met['M8']}"
            else:
                f = f"={col}{met['M9']}"
            put(ws, f"{col}{r}", f, fmt=D2, center=True)
        crit_rows[cname] = r
        r += 1
    first, last = crit_rows[CRIT[0]], crit_rows[CRIT[-1]]
    put(ws, f"A{r}", "Weighted total score", bold=True, fill=WIN)
    put(ws, f"B{r}", "1–5", fill=WIN)
    for j in range(3):
        col = L(3 + j)
        put(ws, f"{col}{r}", f"=SUMPRODUCT({weights},{col}{first}:{col}{last})", fmt=D2, bold=True, center=True, fill=WIN)
    tot = r
    r += 1
    put(ws, f"A{r}", "Rank", bold=True)
    for j in range(3):
        col = L(3 + j)
        put(ws, f"{col}{r}", f"=RANK({col}{tot},$C${tot}:$E${tot})", center=True, bold=True)
    rank = r
    r += 1
    put(ws, f"A{r}", "Winner / margin over runner-up", bold=True)
    put(ws, f"C{r}", f'=INDEX($C$4:$E$4,MATCH(MAX(C{tot}:E{tot}),C{tot}:E{tot},0))', bold=True, center=True)
    put(ws, f"D{r}", f"=MAX(C{tot}:E{tot})-LARGE(C{tot}:E{tot},2)", fmt=D2, center=True)
    put(ws, f"E{r}", None)
    return dict(tot=tot, rank=rank, win=r, crit=crit_rows, met=met, end=r)


# ======================================================================= Scorecard
sc_ = wb.create_sheet("Scorecard")
title(sc_, "Scorecard — base weights, evidence for every score",
      "Scores 1–5 (5 = most attractive for Clay). Columns C–E = segments. Evidence for each criterion is in column G.")
hdr(sc_, 4, ["Item", "Type", SEGS[0], SEGS[1], SEGS[2], "", "Evidence (why these scores)"], widths=[52, 11, 17, 17, 17, 2, 110])
SB = engine_block(sc_, 5, "Base case (suggested weights)")
evid = {
    "Market Attractiveness": "Prof. services has the largest SAM ($148M vs $68–76M, 02_Market_Sizing) but the least incremental headroom: agencies and consultancies are already Clay customers and partners (S32, S34). The two industrial segments are largely unpenetrated: no industrial industry in the top 5 of Clay's detected users (S32); the Sep 2026 marketplace expansion skipped industrials (S09). The two effects nearly offset.",
    "Customer Need": "Distribution: sales occupations are 22% of employment (2.5× average) and 39% of surveyed distributors prioritize new-customer growth (S25). Its customers' personas have ~0.62× the estimated LinkedIn coverage of prof.-services buyers, and Regency reports ZoomInfo and Apollo failed to find contractors (S07). Manufacturing sells indirectly (0.32× sales intensity) and has no data gap. Prof. services has a moderate sales function and no data gap.",
    "Product-Market Fit": "Judgment from the pain-to-capability map (05 §3). Distribution has 4 strong fits (contractor sourcing via Google Maps and Shovels, territory scoring, CRM completeness, e-commerce vetting; all shown in the Regency case) but a gap on ERP-based cross-sell, so 4. Manufacturing: its core jobs (channel management, sourcing-platform demand such as Thomasnet with no API) are weak fits, so 2. Prof. services: account research, pursuit prep and targeted outbound are proven Clay use cases (OpenAI pre-call prep, agency usage), so 4.",
    "Revenue Potential": "Half on Phase 1 SOM ARR (prof. services ~2.2× distribution) and half on seller density: ~28 sales staff per distributor vs ~9 in prof. services and ~4 in manufacturing (OEWS × Census). Clay's usage pricing scales with records and seats, and enterprise NRR is >200% (S02).",
    "Competitive Position": "Phase 2 Segment_Scoring: data-differentiation relevance, incumbent saturation, vertical-specialist threat. Distribution 4.0 (real data gap; thin incumbents: 68% of distributors use spreadsheets, S25). Prof. services 2.3 (well-covered personas, Apollo/ZoomInfo and internal builds entrenched).",
    "Ease of Execution": "Phase 2 adoption fit: builder capacity (computer-occupation index 0.69 / 0.45 / 7.15) and existing references/partners (one distribution case study; no industrial partners found; prof. services has its own case-study filter). Prof. services is far easier to sell into and deliver.",
}
for cname, r in SB["crit"].items():
    put(sc_, f"G{r}", evid[cname])
    sc_.row_dimensions[r].height = 92
for k, w_ in enumerate(WSETS[0][1]):
    pass
rr = SB["end"] + 2
put(sc_, f"A{rr}", "Weighted contribution by criterion (base weights)", bold=True, fill=SFILL)
for col in "BCDE":
    put(sc_, f"{col}{rr}", None, fill=SFILL)
CONTRIB = rr + 1
for k, cname in enumerate(CRIT):
    r = rr + 1 + k
    put(sc_, f"A{r}", cname)
    put(sc_, f"B{r}", f"=Weights!$B${5+k}", fmt=PCT, font=GREEN, center=True)
    for j in range(3):
        col = L(3 + j)
        put(sc_, f"{col}{r}", f"=$B{r}*{col}{SB['crit'][cname]}", fmt=D2, center=True)
r = rr + 7
put(sc_, f"A{r}", "Total (= weighted total above)", bold=True)
for j in range(3):
    col = L(3 + j)
    put(sc_, f"{col}{r}", f"=SUM({col}{rr+1}:{col}{rr+6})", fmt=D2, bold=True, center=True)
CHECK_ROW = r + 1
put(sc_, f"A{CHECK_ROW}", "Check: contribution total − weighted total (should be 0)")
put(sc_, f"C{CHECK_ROW}", f"=ROUND(SUM(C{r}:E{r})-SUM(C{SB['tot']}:E{SB['tot']}),6)", fmt='0.000000', center=True)
put(sc_, "C4", SEGS[0])  # used by INDEX for winner names
ch = BarChart()
ch.type = "bar"
ch.grouping = "stacked"
ch.overlap = 100
ch.title = "Weighted score by criterion (base weights)"
ch = BarChart()
ch.type = "bar"
ch.grouping = "stacked"
ch.overlap = 100
ch.title = "Weighted score by criterion (base weights)"
for k in range(6):
    s_ref = Reference(sc_, min_col=3, max_col=5, min_row=rr + 1 + k, max_row=rr + 1 + k)
    from openpyxl.chart import Series
    ser = Series(s_ref, title=CRIT[k])
    ch.series.append(ser)
ch.set_categories(Reference(sc_, min_col=3, max_col=5, min_row=4, max_row=4))
ch.height, ch.width = 9, 22
ch.y_axis.title = "Weighted score (sum = total, max 5)"
ch.y_axis.delete = False
ch.x_axis.delete = False
sc_.add_chart(ch, f"A{CHECK_ROW+2}")

# ======================================================================= Weight_Sensitivity
wsn = wb.create_sheet("Weight_Sensitivity")
title(wsn, "Weight sensitivity — does the ranking survive different priorities?",
      "Block A: alternative weight sets. Block B: break-even weight per criterion (others rescaled proportionally). Block C: random-weight simulation summary (detail on Random_Weights).")
hdr(wsn, 4, ["Weight set", SEGS[0], SEGS[1], SEGS[2], "Winner", "Margin", "Ranking (1st > 2nd > 3rd)"], widths=[30, 17, 17, 17, 22, 10, 50])
crit_first = SB["crit"][CRIT[0]]
for j, (nm, _) in enumerate(WSETS):
    r = 5 + j
    wcol = L(2 + j)
    put(wsn, f"A{r}", nm, bold=True)
    for k in range(3):
        sc_col = L(3 + k)
        put(wsn, f"{L(2+k)}{r}", f"=SUMPRODUCT(Weights!${wcol}$5:${wcol}$10,Scorecard!{sc_col}${crit_first}:{sc_col}${crit_first+5})", fmt=D2, center=True)
    put(wsn, f"E{r}", f"=INDEX($B$4:$D$4,MATCH(MAX(B{r}:D{r}),B{r}:D{r},0))", center=True, bold=True)
    put(wsn, f"F{r}", f"=MAX(B{r}:D{r})-LARGE(B{r}:D{r},2)", fmt=D2, center=True)
    put(wsn, f"G{r}", f'=INDEX($B$4:$D$4,MATCH(LARGE(B{r}:D{r},1),B{r}:D{r},0))&"  >  "&INDEX($B$4:$D$4,MATCH(LARGE(B{r}:D{r},2),B{r}:D{r},0))&"  >  "&INDEX($B$4:$D$4,MATCH(LARGE(B{r}:D{r},3),B{r}:D{r},0))')
wsn.conditional_formatting.add("B5:D10", ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color="C6E0B4"))
# Block B break-even
bb = 13
put(wsn, f"A{bb}", "Block B — Break-even weight: how much weight on one criterion (others rescaled pro rata) flips the base winner (A) vs each rival?", bold=True, border=False)
wsn.merge_cells(f"A{bb}:G{bb}")
hdr(wsn, bb + 1, ["Criterion", "Base weight", "Score gap A − C (this criterion)", "A − C gap from other criteria (rescaled)", "Break-even weight vs C", "Break-even weight vs B", "Interpretation"])
for k, cname in enumerate(CRIT):
    r = bb + 2 + k
    cr = SB["crit"][cname]
    put(wsn, f"A{r}", cname, bold=True)
    put(wsn, f"B{r}", f"=Weights!$B${5+k}", font=GREEN, fmt=PCT, center=True)
    put(wsn, f"C{r}", f"=Scorecard!C{cr}-Scorecard!E{cr}", fmt=D2, center=True)
    # A_k = (sum_i w_i d_i - w_k d_k)/(1-w_k)
    put(wsn, f"D{r}", f"=(SUMPRODUCT({W_BASE},Scorecard!$C${crit_first}:$C${crit_first+5})-SUMPRODUCT({W_BASE},Scorecard!$E${crit_first}:$E${crit_first+5})-B{r}*C{r})/(1-B{r})", fmt=D2, center=True)
    put(wsn, f"E{r}", f'=IF(AND(D{r}>0,C{r}<0),D{r}/(D{r}-C{r}),"No flip")', fmt=PCT, center=True, bold=True)
    # vs B
    dB = f"(Scorecard!C{cr}-Scorecard!D{cr})"
    AB = f"((SUMPRODUCT({W_BASE},Scorecard!$C${crit_first}:$C${crit_first+5})-SUMPRODUCT({W_BASE},Scorecard!$D${crit_first}:$D${crit_first+5})-B{r}*{dB})/(1-B{r}))"
    put(wsn, f"F{r}", f'=IF(AND({AB}>0,{dB}<0),{AB}/({AB}-{dB}),"No flip")', fmt=PCT, center=True)
    put(wsn, f"G{r}", f'=IF(ISNUMBER(E{r}),"C overtakes A if this weight exceeds "&TEXT(E{r},"0%")&" (base "&TEXT(B{r},"0%")&")","A stays ahead of C at any weight on this criterion")')
EASE_BE = bb + 2 + 5
note(wsn, f"A{bb+9}:G{bb+9}", "Reading: Ease of Execution is the only criterion whose weight can flip the result. Prof. services overtakes distribution once Ease carries roughly 26% or more of the total weight (base 10%). Manufacturing never overtakes distribution.")
# Block C summary of random weights
bc = bb + 11
put(wsn, f"A{bc}", "Block C — Random-weight simulation (1,000 draws each; stored values, seed 42)", bold=True, border=False)
hdr(wsn, bc + 1, ["Weight distribution", "A wins", "B wins", "C wins", "Draws", "", "What it represents"])
sims = [("Flat Dirichlet(1,…,1) — any weighting equally likely", "D", "Agnostic: no view on which criteria matter"),
        ("Centered Dirichlet(30 × base weights)", "M", "Reasonable disagreement around the brief's weights")]
for i, (nm, col, what) in enumerate(sims):
    r = bc + 2 + i
    put(wsn, f"A{r}", nm, bold=True)
    for k, s in enumerate(SHORT):
        put(wsn, f"{L(2+k)}{r}", f'=COUNTIF(Random_Weights!{col}$6:{col}$1005,"{s}")/COUNTA(Random_Weights!{col}$6:{col}$1005)', fmt=PCT, center=True)
    put(wsn, f"E{r}", f"=COUNTA(Random_Weights!{col}$6:{col}$1005)", center=True)
    put(wsn, f"G{r}", what)
WS_SIM = bc + 2

# ======================================================================= Random_Weights
rw = wb.create_sheet("Random_Weights")
title(rw, "Random weight vectors — generated in Python (numpy, seed 42) and stored as values for reproducibility",
      "Columns A–F flat Dirichlet weights; G–I scores; D-col winner. Right block: centered Dirichlet. Scores use Scorecard criterion scores (live).")
rng_ = np.random.default_rng(42)
Wf = rng_.dirichlet(np.ones(6), 1000)
Wc = rng_.dirichlet(np.array(WSETS[0][1]) * 30, 1000)
# layout: flat block cols A..C? Use: A=draw #, then 6 weights (B..G)?  Keep simple:
# Flat: B-G weights, H-J scores, D?  To match Weight_Sensitivity references we put winners in col D (flat) and M (centered).
# Final layout: A draw, flat winner in D requires D free -> use: A draw, B..? Restructure below.
hdr(rw, 5, ["Draw", "Flat score A", "Flat score B", "Flat winner", "w MA", "w Need", "w PMF", "w Rev", "w CP", "w Ease",
            "Flat score C", "", "Centered winner", "w MA", "w Need", "w PMF", "w Rev", "w CP", "w Ease", "Centered score A", "Centered score B", "Centered score C"],
    widths=[7, 9, 9, 18, 8, 8, 8, 8, 8, 8, 3, 9, 18, 8, 8, 8, 8, 8, 8, 9, 9, 9])
sc_cols = ["C", "D", "E"]
for i in range(1000):
    r = 6 + i
    rw.cell(row=r, column=1, value=i + 1)
    for k in range(6):
        c1 = rw.cell(row=r, column=5 + k, value=float(Wf[i, k]))
        c1.font, c1.number_format = BLUE, '0.000'
        c2 = rw.cell(row=r, column=14 + k, value=float(Wc[i, k]))
        c2.font, c2.number_format = BLUE, '0.000'
    # flat scores in B, C and a hidden third? Need 3 score cells: use B,C and X? -> place flat scores in W,X,Y? keep: B=A score, C=B score, and compute C-seg inline in winner formula.
    fA = "+".join(f"{L(5+k)}{r}*Scorecard!$C${crit_first+k}" for k in range(6))
    fB = "+".join(f"{L(5+k)}{r}*Scorecard!$D${crit_first+k}" for k in range(6))
    fC = "+".join(f"{L(5+k)}{r}*Scorecard!$E${crit_first+k}" for k in range(6))
    rw.cell(row=r, column=2, value=f"={fA}").number_format = D2
    rw.cell(row=r, column=3, value=f"={fB}").number_format = D2
    rw.cell(row=r, column=11, value=f"={fC}").number_format = D2
    rw.cell(row=r, column=4, value=f'=IF(AND(B{r}>=C{r},B{r}>=K{r}),"{SHORT[0]}",IF(C{r}>=K{r},"{SHORT[1]}","{SHORT[2]}"))')
    gA = "+".join(f"{L(14+k)}{r}*Scorecard!$C${crit_first+k}" for k in range(6))
    gB = "+".join(f"{L(14+k)}{r}*Scorecard!$D${crit_first+k}" for k in range(6))
    gC = "+".join(f"{L(14+k)}{r}*Scorecard!$E${crit_first+k}" for k in range(6))
    rw.cell(row=r, column=20, value=f"={gA}").number_format = D2
    rw.cell(row=r, column=21, value=f"={gB}").number_format = D2
    rw.cell(row=r, column=22, value=f"={gC}").number_format = D2
    rw.cell(row=r, column=13, value=f'=IF(AND(T{r}>=U{r},T{r}>=V{r}),"{SHORT[0]}",IF(U{r}>=V{r},"{SHORT[1]}","{SHORT[2]}"))')
rw.freeze_panes = "A6"
# fix Weight_Sensitivity references: flat winner col D, centered winner col M
for i, (nm, col, what) in enumerate(sims):
    pass

# ======================================================================= Stress_Tests
st = wb.create_sheet("Stress_Tests")
title(st, "Stress tests — uncertain inputs and scoring choices (base weights)",
      "Each test changes the yellow input(s) only; everything else is linked to Inputs. Summary at top; full blocks below (metric rows are shown so every change is visible).")
hdr(st, 4, ["Test", "What changes", SEGS[0], SEGS[1], SEGS[2], "Winner", "Margin"], widths=[44, 52, 16, 16, 16, 22, 10])
tests = [
    ("T0 Base case", "Nothing (reproduces Scorecard)", {}, "ratio"),
    ("T1 Partner channel works", "Distribution ease of execution 2.0 → 3.5 (templates + partners close half the gap)", {("M9", 0): 3.5}, "ratio"),
    ("T2 Prof. services headroom larger", "Prof. services headroom flag 2 → 4 (Clay less penetrated than assumed)", {("M2", 2): 4}, "ratio"),
    ("T3 No data gap (H3 proxy wrong)", "Distribution persona coverage index 0.62 → 1.00 (standard databases cover contractors)", {("M4", 0): 1.0}, "ratio"),
    ("T4 Seller density doesn't lift ACV", "Revenue Potential = 100% SOM ARR (no account-expansion credit)", {"SW_R": 1.0}, "ratio"),
    ("T5 ERP gap binds", "Distribution PMF 4 → 3", {("M5", 0): 3}, "ratio"),
    ("T6 Apollo/HubSpot close local-data gap", "Distribution competitive position 4.0 → 3.0", {("M8", 0): 3.0}, "ratio"),
    ("T7 Distribution combined downside", "T3 + T5 + T6 together", {("M4", 0): 1.0, ("M5", 0): 3, ("M8", 0): 3.0}, "ratio"),
    ("T8 Prof. services best case", "PS headroom 4, competitive position 3.0, ease 5.0", {("M2", 2): 4, ("M8", 2): 3.0, ("M9", 2): 5.0}, "ratio"),
    ("T9 Worst case for A + best case for C", "T7 + T8 together", {("M4", 0): 1.0, ("M5", 0): 3, ("M8", 0): 3.0, ("M2", 2): 4, ("M8", 2): 3.0, ("M9", 2): 5.0}, "ratio"),
    ("T10 Market size only (no headroom)", "Market Attractiveness = 100% SAM", {"SW_MA": 1.0}, "ratio"),
    ("T11 Min-max normalization", "Quantitative metrics scaled min-max instead of ratio-to-max (exaggerates gaps)", {}, "minmax"),
]
blocks = []
top = 4 + len(tests) + 4
for i, (nm, what, ov, rule) in enumerate(tests):
    b = engine_block(st, top, nm + " — " + what, overrides=ov, rule=rule)
    blocks.append(b)
    top = b["end"] + 3
for i, (nm, what, ov, rule) in enumerate(tests):
    r = 5 + i
    b = blocks[i]
    put(st, f"A{r}", nm, bold=True)
    put(st, f"B{r}", what)
    for j in range(3):
        col = L(3 + j)
        put(st, f"{col}{r}", f"={col}{b['tot']}", fmt=D2, center=True)
    put(st, f"F{r}", f"=INDEX($C$4:$E$4,MATCH(MAX(C{r}:E{r}),C{r}:E{r},0))", bold=True, center=True)
    put(st, f"G{r}", f"=MAX(C{r}:E{r})-LARGE(C{r}:E{r},2)", fmt=D2, center=True)
    st.row_dimensions[r].height = 30
st.conditional_formatting.add(f"C5:E{4+len(tests)}", ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color="C6E0B4"))
put(st, f"A{4+len(tests)+2}", "Detail blocks follow. In each block the winner formula reads segment names from row 4.", border=False)
ST_SUM_FIRST, ST_SUM_LAST = 5, 4 + len(tests)

# ======================================================================= Subsegment_Drilldown
sd = wb.create_sheet("Subsegment_Drilldown")
title(sd, "Within distribution — which sub-industry is the beachhead?",
      "Census 2022 inputs [A] (Raw: appendix/raw_census_ecnsize_2022_firm_size.csv); judgment scores [C] in yellow. Same ratio-to-max scoring rule.")
hdr(sd, 4, ["Metric", "Weight", "4235 Metals & minerals", "4236 Electrical & electronic", "4237 Hardware, plumbing & HVAC", "4238 Machinery, equipment & supplies", "Basis"],
    widths=[44, 10, 17, 17, 17, 17, 80])
sub_in = [
    ("Firms with 20+ employees (incl. imputed 4235 bands)", [1268, 2578, 1821, 6009], '#,##0', "02_Market_Sizing › Firm_Counts (before cross-industry de-dup)"),
    ("Establishments per firm (all full-year firms)", [1.423, 1.640, 2.213, 1.574], '0.00', "Census EC2200SIZEEMPFIRM (S14)"),
    ("Avg receipts per firm, 20–99 employees ($M)", [87.6, 44.3, 32.2, 34.1], '0.0', "Census (S14); 4235 uses 20–49 band only (50–99 suppressed)"),
    ("Customer-persona data-gap relevance (judgment, 1–5)", [2, 5, 5, 4], '0', "4236/4237 sell to electrical and PHVAC contractors (est. ~19–22% LinkedIn use, 04 Persona_Coverage); 4238 mixes plant MRO (maintenance workers ~20%) with dealers; 4235 sells mainly to fabricators and manufacturers (white-collar purchasing)"),
    ("Published Clay reference in sub-industry (1 = none, 5 = yes)", [1, 5, 1, 1], '0', "Regency Supply = commercial lighting & electrical distributor (S07)"),
]
for i, (n, vals, fmt, basis) in enumerate(sub_in):
    r = 5 + i
    put(sd, f"A{r}", n)
    put(sd, f"B{r}", None)
    for j, v in enumerate(vals):
        put(sd, f"{L(3+j)}{r}", v, font=BLUE, fmt=fmt, center=True, fill=YFILL if i >= 3 else None)
    put(sd, f"G{r}", basis)
    sd.row_dimensions[r].height = 40
hdr(sd, 11, ["Score (1–5)", "Weight", "4235", "4236", "4237", "4238", "Rule"])
sub_scores = [("Size", 0.20, 5, "ratio"), ("Branch / territory intensity", 0.15, 6, "ratio"), ("Ability to pay", 0.15, 7, "ratio"),
              ("Data-gap relevance (H3)", 0.35, 8, "direct"), ("Reference proof", 0.15, 9, "direct")]
for i, (n, w, src, how) in enumerate(sub_scores):
    r = 12 + i
    put(sd, f"A{r}", n, bold=True)
    put(sd, f"B{r}", w, font=BLUE, fmt=PCT, fill=YFILL, center=True)
    for j in range(4):
        col = L(3 + j)
        f = f"=1+4*{col}{src}/MAX($C${src}:$F${src})" if how == "ratio" else f"={col}{src}"
        put(sd, f"{col}{r}", f, fmt=D2, center=True)
    put(sd, f"G{r}", "1 + 4 × value ÷ max" if how == "ratio" else "Judgment score used directly")
put(sd, "A17", "Weighted score", bold=True, fill=WIN)
put(sd, "B17", "=SUM(B12:B16)", fmt=PCT, center=True, fill=WIN)
for j in range(4):
    col = L(3 + j)
    put(sd, f"{col}17", f"=SUMPRODUCT($B$12:$B$16,{col}12:{col}16)", fmt=D2, bold=True, center=True, fill=WIN)
put(sd, "A18", "Rank", bold=True)
for j in range(4):
    col = L(3 + j)
    put(sd, f"{col}18", f"=RANK({col}17,$C$17:$F$17)", center=True, bold=True)
put(sd, "A19", "Rank without the Regency reference (reference weight → 0, rescaled)", bold=True)
for j in range(4):
    col = L(3 + j)
    put(sd, f"{col}19", f"=SUMPRODUCT($B$12:$B$15,{col}12:{col}15)/SUM($B$12:$B$15)", fmt=D2, center=True)
put(sd, "A20", "Rank (no-reference variant)", bold=True)
for j in range(4):
    col = L(3 + j)
    put(sd, f"{col}20", f"=RANK({col}19,$C$19:$F$19)", center=True)
note(sd, "A22:G22", "Reading: electrical distributors (4236) are the strongest beachhead. They combine contractor-facing buyers (largest data gap), multi-branch territories, healthy revenue per firm and the only published Clay proof point. PHVAC (4237) and machinery/MRO (4238) are the natural expansion waves. Without the Regency reference, 4236, 4237 and 4238 are close, so the beachhead choice partly rests on having a proof point to reuse. Metals service centers (4235) rank last because their buyers are white-collar purchasing teams already covered by standard data.", 58)

# ======================================================================= Options_Tradeoffs
ot = wb.create_sheet("Options_Tradeoffs")
title(ot, "Strategic options — what we recommend, what we reject, and why",
      "Options from 01_Project_Charter §4. Recommendation is labeled [E]; Phase 4 tests its economics and Phase 5 designs it.")
hdr(ot, 4, ["#", "Option", "Score evidence", "Main upside", "Main risk / trade-off", "Reversibility", "Verdict"], widths=[5, 34, 44, 40, 48, 16, 30])
opts = [
    ("0", "No vertical bet: keep horizontal, persona-led growth",
     "Comparison baseline.",
     "Zero incremental cost; focus stays on Clay's fast-growing core.",
     "Distribution's need does not get served: Clay's detected base is ~40% software, there are no industrial partners, and the marketplace expansion skipped industrials. The gap persists or a specialist fills it.",
     "Fully", "Rejected for distribution; retained for prof. services, which the horizontal motion already serves"),
    ("1", "Full industrial-distribution vertical build (dedicated sales team, vertical marketing)",
     "Distribution ranks #1 (base 3.97 vs 3.41).",
     "Fastest capture if the thesis is right; first-mover in an under-served niche.",
     "Ease of execution is 2.0/5: committing a sales org before proving adoption and ACV risks a high-CAC, high-churn segment. Payback is unproven (Phase 4).",
     "Low (hiring, brand)", "Deferred: becomes the scale-up path if pilot Go criteria are met"),
    ("2", "Industrial-manufacturing vertical",
     "Ranks #3 under every named weight set and wins none of the 2,000 random weight draws.",
     "Large firm count (23.8K firms with 20+ employees).",
     "Direct sales intensity 0.32× average, no data gap, lowest builder capacity. Manufacturers' selling runs through distributors and reps.",
     "Medium", "Rejected as a stand-alone vertical; reach manufacturers indirectly through distributors (H10)"),
    ("3", "Dedicated B2B professional-services vertical program",
     "Ranks #2 (3.41); wins only if Ease of Execution weight exceeds ~26% or in combined stress tests.",
     "Largest SAM, easiest to adopt, existing proof points.",
     "Little incremental headroom: agencies and consultancies are already Clay customers and partners. A dedicated program mostly re-buys demand the horizontal motion already captures. Low differentiation vs Apollo and internal builds.",
     "High", "Rejected as a dedicated program; continue serving through the core motion"),
    ("4", "Staged, partner-led pilot in industrial distribution (beachhead: electrical 4236, then PHVAC 4237 and MRO/machinery 4238)",
     "Uses the #1 segment while directly testing its #1 weakness (ease of execution, T1).",
     "Captures distribution's need and competitive advantage with limited capital. Builds the templates and partner channel that the scores say are missing.",
     "Slower than a full build; pilot learnings may not generalize across sub-industries; needs clear Go/No-Go gates.",
     "High", "RECOMMENDED [E], subject to Phase 4 unit economics and Phase 5 Go/No-Go design"),
]
for i, o in enumerate(opts):
    r = 5 + i
    for j, v in enumerate(o):
        put(ot, f"{L(1+j)}{r}", v, bold=(j in (1, 6)), fill=WIN if o[0] == "4" else None)
    ot.row_dimensions[r].height = 96
note(ot, "A11:G11", "Key trade-off: choosing distribution trades a smaller, harder-to-serve market for a larger, more defensible advantage. Professional services offers the opposite: an easy, larger market with little incremental headroom and weak differentiation. The staged pilot is how Clay buys the information needed to resolve this trade-off before committing a sales organization.", 52)

# ======================================================================= Change_Conditions
cc = wb.create_sheet("Change_Conditions")
title(cc, "What would change the recommendation — signposts to monitor",
      "Each condition links to the stress test or break-even that shows its effect. Thresholds marked [D] are proposals to be finalized in Phase 5 Go/No-Go criteria.")
hdr(cc, 4, ["#", "Condition (evidence that would change the call)", "How to observe it", "Linked test", "Effect on ranking", "Response"], widths=[5, 50, 46, 14, 26, 46])
conds = [
    ("1", "Management weights near-term ease of execution at ≥ ~26% of the decision",
     "Executive priorities; capacity constraints in sales/partnerships",
     "Weight_Sensitivity Block B", f"=Weight_Sensitivity!G{EASE_BE}",
     "Prioritize prof. services through product-led packaging; keep a lighter distribution experiment"),
    ("2", "Contractor contact match rates for Apollo/ZoomInfo come within ~10 pts of Clay's waterfall [D]",
     "Coverage test on ~200 licensed contractors (appendix/Interview_Guide.md)",
     "T3, T6, T7", "=Stress_Tests!F12&\" wins under T7 (margin \"&TEXT(Stress_Tests!G12,\"0.00\")&\")\"",
     "Drop the data-coverage message; re-score need and competitive position"),
    ("3", "Partners plus templates can't get pilot accounts to self-operate within ~60 days [D]",
     "Pilot time-to-value and active-workflow metrics",
     "T1 (reverse)", "Ease stays ≤ 2; risk of churn",
     "Pause scale-up; consider a managed-service offer or stop"),
    ("4", "Clay's penetration of prof. services proves low (headroom ≥ 4)",
     "Customer-base mix by NAICS (internal data); partner surveys",
     "T2, T8", "=Stress_Tests!F7&\" wins under T2; \"&Stress_Tests!F13&\" under T8\"",
     "Re-open Option 3"),
    ("5", "Distributor ACV stays at the Growth entry floor (~$5.4K) despite seller density",
     "Pilot contract values and usage (Actions, Data Credits)",
     "T4", "=Stress_Tests!F9&\" wins under T4 (margin \"&TEXT(Stress_Tests!G9,\"0.00\")&\")\"",
     "Phase 4 payback likely fails; keep distribution self-serve only"),
    ("6", "Apollo or HubSpot ships native off-LinkedIn local data (maps, permits, licenses)",
     "Competitor release notes and changelogs",
     "T6", "=Stress_Tests!F11&\" wins under T6\"",
     "Shift differentiation to workflows and templates; accelerate the partner moat"),
    ("7", "Distribution survey evidence fails to replicate (new-customer priority, CRM gaps)",
     "Interviews (appendix/Interview_Guide.md); larger industry surveys",
     "T3/T5", "Need and PMF scores fall",
     "Re-run Scorecard"),
]
for i, c in enumerate(conds):
    r = 5 + i
    for j, v in enumerate(c):
        put(cc, f"{L(1+j)}{r}", v)
    cc.row_dimensions[r].height = 58
note(cc, "A13:F13", "Robustness: distribution stays first under each single-input test that does not involve execution, including removing the data gap (T3). It loses first place only when several downside assumptions combine with a professional-services best case (T9), or when ease of execution dominates the weights.", 44)

# ======================================================================= Executive_Summary
ex = wb.create_sheet("Executive_Summary", 1)
title(ex, "Executive summary — Phase 3 segment prioritization", "All figures link to the analysis tabs.")
hdr(ex, 4, ["Metric", SEGS[0], SEGS[1], SEGS[2], "Note"], widths=[58, 18, 18, 18, 52])
r = 5
for cname in CRIT:
    put(ex, f"A{r}", f"{cname} (weight " + f"{int(round(WSETS[0][1][CRIT.index(cname)]*100))}%)")
    for j in range(3):
        put(ex, f"{L(2+j)}{r}", f"=Scorecard!{L(3+j)}{SB['crit'][cname]}", font=GREEN, fmt=D2, center=True)
    put(ex, f"E{r}", "Score 1–5")
    r += 1
put(ex, f"A{r}", "Weighted total (base weights)", bold=True, fill=WIN)
for j in range(3):
    put(ex, f"{L(2+j)}{r}", f"=Scorecard!{L(3+j)}{SB['tot']}", font=GREEN, fmt=D2, center=True, bold=True, fill=WIN)
put(ex, f"E{r}", "", fill=WIN)
EX_TOT = r
r += 1
put(ex, f"A{r}", "Rank", bold=True)
for j in range(3):
    put(ex, f"{L(2+j)}{r}", f"=Scorecard!{L(3+j)}{SB['rank']}", font=GREEN, center=True, bold=True)
r += 2
put(ex, f"A{r}", "Robustness", bold=True, fill=SFILL)
for col in "BCDE":
    put(ex, f"{col}{r}", None, fill=SFILL)
r += 1
robust = [
    ("Wins across 6 named weight sets", [f'=COUNTIF(Weight_Sensitivity!$E$5:$E$10,"{s}")' for s in SEGS], '0', "Weight_Sensitivity Block A"),
    ("Share of 1,000 flat random weightings won", [f"=Weight_Sensitivity!{L(2+k)}{WS_SIM}" for k in range(3)], PCT, "Any weighting equally likely"),
    ("Share of 1,000 weightings centered on base won", [f"=Weight_Sensitivity!{L(2+k)}{WS_SIM+1}" for k in range(3)], PCT, "Disagreement around base weights"),
    ("Wins across 12 input stress tests (incl. base)", [f'=COUNTIF(Stress_Tests!$F${ST_SUM_FIRST}:$F${ST_SUM_LAST},"{s}")' for s in SEGS], '0', "Stress_Tests"),
]
for n, fs, fmt, nt in robust:
    put(ex, f"A{r}", n)
    for j, f in enumerate(fs):
        put(ex, f"{L(2+j)}{r}", f, font=GREEN, fmt=fmt, center=True)
    put(ex, f"E{r}", nt)
    r += 1
put(ex, f"A{r}", "Ease-of-Execution weight at which C overtakes A")
put(ex, f"B{r}", f"=Weight_Sensitivity!E{EASE_BE}", font=GREEN, fmt=PCT, center=True)
put(ex, f"E{r}", "Base weight 10%; only criterion that can flip the result")
r += 1
put(ex, f"A{r}", "Distribution beachhead (highest-scoring sub-industry)")
put(ex, f"B{r}", '=INDEX(Subsegment_Drilldown!$C$4:$F$4,MATCH(1,Subsegment_Drilldown!$C$18:$F$18,0))', font=GREEN, center=True)
ex.merge_cells(f"B{r}:D{r}")
put(ex, f"E{r}", "Subsegment_Drilldown")
r += 2
put(ex, f"A{r}", "Recommendation [E] and rationale", bold=True, fill=SFILL)
for col in "BCDE":
    put(ex, f"{col}{r}", None, fill=SFILL)
r += 1
recs = [
    "1. Prioritize U.S. industrial distribution, entering through a staged, partner-led pilot (Option 4) rather than a full vertical build. Beachhead: electrical distributors (NAICS 4236); expansion: PHVAC (4237), then machinery/MRO (4238).",
    "2. Why distribution: it wins on customer need (5.0), competitive position (4.0) and product fit (4.0), the three things a vertical program can't buy later. It ranks first in 5 of 6 named weight sets (all except 'Execution first'), ~74% of agnostic random weightings, ~99% of weightings near the brief's weights, and 11 of 12 input stress tests. The margin is thin only when several distribution downsides combine (T7: +0.05) or when headroom is ignored (T10: +0.04).",
    "3. Why not professional services: it is larger and much easier to sell, but Clay already serves it through the horizontal motion and agency ecosystem. A dedicated program mostly re-buys existing demand, and its differentiation vs Apollo and internal builds is weakest. It wins only if ease of execution carries ≳26% of the weight.",
    "4. Why not manufacturing: last under every weighting. Its selling runs through distributors and reps, so it is best reached indirectly through the distribution beachhead.",
    "5. The pilot must answer the one question the scores can't: can templates plus partners raise distribution's ease of execution from 2.0 toward ~3.5 (T1)? Phase 4 sizes whether the resulting economics pay back.",
]
for t in recs:
    put(ex, f"A{r}", t)
    ex.merge_cells(f"A{r}:E{r}")
    ex.row_dimensions[r].height = 44
    r += 1

# ======================================================================= Sources
so = wb.create_sheet("Sources")
title(so, "Sources and upstream files", "Full register: Research_Sources.csv")
hdr(so, 4, ["ID / file", "Description", "Used for"], widths=[34, 80, 50])
srcs = [
    ("02_Market_Sizing.xlsx", "Phase 1 market sizing (Census 2022 firm counts; OEWS May 2025)", "M1, M3, M6, M7; sub-segment firm counts"),
    ("04_Competitive_Analysis.xlsx", "Phase 2 competitive analysis", "M2, M4, M8, M9"),
    ("05_Customer_Needs_Assessment.md", "Phase 2 needs and pain-to-capability map", "M5 product-market fit evidence"),
    ("S02", "Clay tender offer release: enterprise NRR >200%", "Revenue Potential evidence"),
    ("S07", "Regency Supply case study (electrical distributor)", "Need, PMF, sub-segment reference proof"),
    ("S09", "Clay integrations roundup, Sep 2026 (industries added)", "Market Attractiveness (headroom)"),
    ("S14", "2022 Economic Census, Employment Size of Firms", "Sub-segment drilldown inputs"),
    ("S25", "Phocas State of Sales in Distribution 2026 (via DSG)", "Customer Need, Competitive Position evidence"),
    ("S32 / S34", "Bloomberry: Clay users and solution partners", "Headroom flag evidence"),
]
for i, s in enumerate(srcs, start=5):
    for j, v in enumerate(s, start=1):
        put(so, f"{L(j)}{i}", v)

order = ["Cover", "Executive_Summary", "Weights", "Inputs", "Scorecard", "Weight_Sensitivity", "Random_Weights",
         "Stress_Tests", "Subsegment_Drilldown", "Options_Tradeoffs", "Change_Conditions", "Sources"]
wb._sheets = [wb[n] for n in order]
for w in wb.worksheets:
    w.sheet_properties.tabColor = {"Cover": "1F2A44", "Executive_Summary": "1F2A44", "Weights": "FFC000", "Inputs": "FFC000",
                                   "Random_Weights": "A6A6A6", "Sources": "A6A6A6"}.get(w.title, "4472C4")
wb.save(OUT)
print("saved", OUT, "Scorecard rows", SB)
