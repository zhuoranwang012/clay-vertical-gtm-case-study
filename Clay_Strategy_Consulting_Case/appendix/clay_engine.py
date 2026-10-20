"""Clay vertical-expansion model, v2 — pure-Python engine.

Mirrors, formula for formula, the engine written into 07_Financial_Model.xlsx by build_financial_model.py.
It is used for (1) the Monte Carlo simulation and gate-policy analysis, which Excel cannot run natively,
and (2) an independent reconciliation of the workbook (Monte Carlo sheet compares both engines run by run).

v2 changes vs v1 (see 07 › Change Log): derived win mix by size band; dedicated-AE outbound and
program-spend elasticity (sales-led can now create pipeline); outbound fatigue; meeting→opportunity
funnel stages; NPV at a hurdle rate with an incrementality haircut; Option 0 opportunity-cost benchmark;
professional-services segment on the same engine; gate policy with observation noise.
"""
import math

QN = 12
BANDS = ("S", "M", "E")

# ----------------------------------------------------------------------------- segment data
# Census 2022 Economic Census, Employment Size of Firms (S14). Bands: 20-49, 50-99, 100-249, 250-499, 500+.
CENSUS = {
    "4236": ([1408, 570, 350, 136, 114], [42698, 39297, 54183, 47492, 315093]),
    "4237": ([1013, 424, 242, 85, 57], [31260, 29270, 36995, 29008, 116895]),
    "4238": ([3622, 1218, 746, 240, 183], [110129, 83565, 115979, 83365, 220621]),
    "5413": ([5930, 2104, 1251, 429, 387], [179028, 145073, 189979, 149368, 727292]),
    "5416": ([5789, 2151, 1244, 430, 339], [174222, 148901, 188821, 149492, 677207]),
    "5415": ([5831, 2359, 1581, 561, 447], [179046, 164769, 240853, 195289, 1156449]),
}
# De-duplication factors by size band (sector firms ÷ Σ sub-industry firms; 02 Firm_Counts / overlap_raw.csv).
# 423: measurable only for 20-49 (0.973) and 500+ (≈1.0); 50-499 not measurable (suppressed 4235 cells) → 1.0.
DEDUP = {"dist": [0.973, 1.0, 1.0, 1.0, 1.0], "ps": [0.988, 0.979, 0.967, 0.951, 0.972]}
SEGMENTS = {
    "dist": {"label": "Industrial distribution", "waves": ["4236", "4237", "4238"]},
    "ps": {"label": "Professional services", "waves": ["5413", "5416", "5415"]},
}


def seg_firms(seg):
    """Firms by wave and band group (S=20-99, M=100-499, E=500+), de-duplicated."""
    dd = DEDUP[seg]
    out = []
    for n in SEGMENTS[seg]["waves"]:
        f = CENSUS[n][0]
        g = [f[i] * dd[i] for i in range(5)]
        out.append({"S": g[0] + g[1], "M": g[2] + g[3], "E": g[4]})
    return out


def seg_emp_per_firm(seg):
    """Average employees per firm by band group across the segment's waves (not de-duplicated: an average)."""
    tot = {b: [0, 0] for b in BANDS}
    for n in SEGMENTS[seg]["waves"]:
        f, e = CENSUS[n]
        for b, idx in (("S", (0, 1)), ("M", (2, 3)), ("E", (4,))):
            tot[b][0] += sum(f[i] for i in idx)
            tot[b][1] += sum(e[i] for i in idx)
    return {b: tot[b][1] / tot[b][0] for b in BANDS}


# ----------------------------------------------------------------------------- inputs
# name: (dist Downside, Base, Upside, ps Downside, Base, Upside, unit, class, mc_group, label, source)
# mc_group: None = fixed (not sampled); 'demand' | 'value' | 'cost' | 'timing' | 'indep'
INPUTS = [
    # ---- market (fixed across scenarios: the universe does not flex)
    ("f1", .80, .80, .80, .55, .55, .55, "%", "D", None, "SAM filter 1: active outbound / territory sales motion", "Phase 1 (02 A4/A6); sales occupations 22.0% vs 6.2% (S15)"),
    ("f2", .56, .56, .56, .70, .70, .70, "%", "B/D", None, "SAM filter 2: CRM in use (integrable today)", "Dist: 56% use CRM, Phocas/DSG 2026 (S25) — newest evidence (v1 used 65%). PS: no public rate; Phase 1 assumption (02 A9)"),
    ("exist", .03, .03, .03, .10, .10, .10, "%", "D", None, "Share of SAM already Clay customers (excluded)", "No public data. PS higher: Clay markets to PS (S06) and IT services is a top industry of detected Clay users (S32)"),
    ("w1", 1, 1, 1, 1, 1, 1, "qtr", "E", None, "Wave 1 launch quarter (Dist 4236 electrical · PS 5413 A&E)", "Beachheads: 06 drilldown (Dist); A&E pursuit-data gap, Deltek/Unanet ecosystem (PS, 05)"),
    ("w2", 7, 5, 4, 7, 5, 4, "qtr", "D", "timing", "Wave 2 launch quarter (Dist 4237 PHVAC · PS 5416 consulting)", "After Gate 1; Downside delays"),
    ("w3", 99, 7, 6, 99, 7, 6, "qtr", "D", "timing", "Wave 3 launch quarter, 99 = never (Dist 4238 MRO · PS 5415 IT services)", "PS excludes 5418 advertising/PR: Clay's agency partners (channel conflict) [E]"),
    # ---- demand
    ("cov", .30, .40, .50, .30, .40, .50, "%", "D", "demand", "Automated outbound coverage: share of active SAM sequenced per quarter", "Clay-run outbound with its own product"),
    ("tiltM", 2, 2, 2, 2, 2, 2, "x", "E", None, "ICP tilt: outbound weight on Mid-Market accounts (vs SMB = 1)", "Strategic lever (08 ICP); changes allocation of the same coverage, not the total"),
    ("tiltE", 2, 2, 2, 2, 2, 2, "x", "E", None, "ICP tilt: outbound weight on Enterprise accounts (vs SMB = 1)", "Strategic lever"),
    ("meet", .020, .030, .045, .015, .025, .035, "%", "D", "demand", "Meeting rate per account sequenced (first year)", "Cold outbound 1–5% [K]; PS buyers more heavily prospected → lower"),
    ("fatigue", .60, .75, .85, .60, .75, .85, "x/yr", "D", "indep", "Meeting-rate retention per year of re-sequencing the same accounts", "Outbound fatigue (review m3); applied by wave age"),
    ("inb", .005, .010, .015, .010, .015, .020, "%", "D", "demand", "Inbound / template MQLs per quarter at base program spend (% of active SAM)", "PS more digital → higher"),
    ("cpl", 600, 400, 250, 600, 400, 250, "$", "B/D", "cost", "Cost per extra MQL from program spend above base", "SaaS CPL commonly cited at $200–350 (S55); niche industrial audience costs more"),
    ("inbcap", .03, .03, .03, .03, .03, .03, "%", "D", None, "Ceiling on extra MQLs per quarter (% of active SAM)", "Diminishing returns: extra MQLs = cap × (1 − e^(−spend/(CPL × cap)))"),
    ("p0", 1, 2, 3, 2, 3, 4, "#", "D", "demand", "Partners active at launch", "Dist: no distribution partners visible today (S35). PS: agency ecosystem exists (S34) but competes"),
    ("padd", 1, 1, 2, 1, 1, 2, "#/qtr", "D", "demand", "Partners added per quarter", ""),
    ("pcap", 5, 10, 15, 6, 10, 14, "#", "D", "demand", "Partner network cap", ""),
    ("lpp", 2, 3, 4, 2, 3, 4, "#/qtr", "D", "demand", "Introductions per active partner per quarter", ""),
    ("m2o", .30, .40, .50, .30, .40, .50, "%", "B/D", "demand", "Outbound meeting → opportunity", "SQL→opportunity 42% (S49); a booked meeting ≈ SQL stage (fixes v1 stage mix-up, review M4)"),
    ("p2o", .40, .50, .60, .40, .50, .60, "%", "D", "demand", "Partner introduction → opportunity", "Warm, pre-vetted introductions"),
    ("i2o", .10, .16, .22, .10, .16, .22, "%", "B", "demand", "Inbound MQL → opportunity", "MQL→SQL 39% × SQL→opp 42% (S49)"),
    ("win", .15, .21, .30, .13, .18, .25, "%", "B/D", "demand", "Win rate (opportunity → won)", "HubSpot average 21% (S48) as Base for a new vertical with one reference (v1: 28%). PS lower: competitive position 2.3 vs 4.0 (04)"),
    ("propM", 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, "x", "D", None, "Mid-Market opportunity propensity per sequenced account (SMB = 1)", ""),
    ("propE", .5, .7, .9, .5, .7, .9, "x", "D", "value", "Enterprise opportunity propensity per sequenced account (SMB = 1)", "Enterprise: buying committees, incumbents; lower conversion (review C2)"),
    ("lagSM", 1, 1, 1, 1, 1, 1, "qtr", "B", None, "Sales cycle, SMB & Mid-Market (quarters from opportunity to close)", "Deals < $25K close in 2–3 months (S48)"),
    ("lagE", 2, 2, 2, 2, 2, 2, "qtr", "B/D", None, "Sales cycle, Enterprise (quarters)", "Longer enterprise cycles (review C2)"),
    ("ss", .15, .30, .40, .30, .45, .60, "%", "D", "value", "Self-serve share of SMB wins (no AE)", "Builder capacity: Dist 0.69× vs PS 7.2× technical-talent index (H4)"),
    ("aeacc", 100, 150, 200, 100, 150, 200, "#/qtr", "D", "demand", "Accounts a ramped dedicated AE works per quarter (beyond automated coverage)", "AE self-sourcing (review C1); practitioner guidance: AEs self-source < ~30% of pipeline [K]"),
    ("aemult", 1.25, 1.5, 1.75, 1.25, 1.5, 1.75, "x", "D", "demand", "AE outreach meeting-rate multiplier vs automated outbound", "Human, multi-channel prospecting"),
    # ---- ACV build
    ("salesshare", .22, .22, .22, .062, .062, .062, "%", "A→C", None, "Sales occupations share of employment", "OEWS May 2025 (S15)"),
    ("aps", 150, 150, 150, 150, 150, 150, "#", "D", None, "Accounts tracked per seller", "Same for both segments (neutrality)"),
    ("usage", .75, 1.0, 1.5, .75, 1.0, 1.5, "x", "D", "value", "Usage intensity factor", ""),
    ("cpa", 2, 2, 2, 2, 2, 2, "#", "D", None, "Contacts per account", ""),
    ("refresh", 2, 2, 2, 2, 2, 2, "#/yr", "D", None, "Contact refreshes per year", ""),
    ("crc", 6, 6, 6, 6, 6, 6, "credits", "B/D", None, "Data credits per contact refresh", "S51"),
    ("ai", 4, 4, 4, 4, 4, 4, "credits", "D", None, "AI research credits per account per year", "S01"),
    ("newacc", .25, .25, .25, .25, .25, .25, "%", "D", None, "New accounts sourced per year (% of tracked)", ""),
    ("crnew", 8, 8, 8, 8, 8, 8, "credits", "D", None, "Data credits per new account sourced", ""),
    ("apo", 6, 6, 6, 6, 6, 6, "actions", "D", None, "Actions per row operation", ""),
    ("dcprice", 261, 261, 261, 261, 261, 261, "$/mo", "A", None, "Data-credit pack price (Growth, 72K credits/yr)", "S01"),
    ("dcsize", 72000, 72000, 72000, 72000, 72000, 72000, "credits", "A", None, "Data-credit pack size", "S01"),
    ("acprice", 185, 185, 185, 185, 185, 185, "$/mo", "A", None, "Actions pack price (Growth, 480K actions/yr)", "S01"),
    ("acsize", 480000, 480000, 480000, 480000, 480000, 480000, "actions", "A", None, "Actions pack size", "S01"),
    ("entcr", 1200000, 1200000, 1200000, 1200000, 1200000, 1200000, "credits", "A", None, "Enterprise minimum data credits per year", "S01"),
    ("entac", 2400000, 2400000, 2400000, 2400000, 2400000, 2400000, "actions", "A", None, "Enterprise minimum actions per year", "S01"),
    ("entdisc", .30, .30, .30, .30, .30, .30, "%", "B", None, "Enterprise discount vs list", "Vendr (S16)"),
    # ---- retention
    ("grrS", .85, .88, .91, .85, .88, .91, "%", "B", "value", "GRR — SMB", "SaaS Capital 2025 (S43)"),
    ("grrM", .86, .90, .93, .86, .90, .93, "%", "B", "value", "GRR — Mid-Market", "S43"),
    ("grrE", .88, .92, .95, .88, .92, .95, "%", "B", "value", "GRR — Enterprise", "S43"),
    ("nrrS", .92, 1.00, 1.06, .92, 1.00, 1.06, "%", "B", "value", "NRR — SMB", "S43"),
    ("nrrM", .98, 1.06, 1.15, .98, 1.06, 1.15, "%", "B", "value", "NRR — Mid-Market", "S43"),
    ("nrrE", 1.05, 1.15, 1.30, 1.05, 1.15, 1.30, "%", "B", "value", "NRR — Enterprise", "S43; Clay enterprise NRR > 200% (S02) not assumed"),
    # ---- cost
    ("quota", 875000, 875000, 875000, 875000, 875000, 875000, "$", "B", None, "AE annual quota", "Bridge Group 2026 (S44)"),
    ("prod", .55, .70, .85, .55, .70, .85, "%", "B", "demand", "AE productivity (attainment)", "S44: 48% of reps at quota"),
    ("aeote", 200000, 200000, 200000, 200000, 200000, 200000, "$", "B", None, "AE on-target earnings", "S44"),
    ("wshare", .70, .70, .70, .70, .70, .70, "%", "A", None, "Wages as share of total compensation", "BLS ECEC June 2026 (S45)"),
    ("wmgr", 186910, 186910, 186910, 186910, 186910, 186910, "$", "A", None, "Sales manager wage", "OEWS 513200 (S46)"),
    ("wmkt", 206930, 206930, 206930, 206930, 206930, 206930, "$", "A", None, "Marketing manager wage", "S46"),
    ("wdev", 158300, 158300, 158300, 158300, 158300, 158300, "$", "A", None, "Software developer wage", "S46"),
    ("wcs", 112510, 112510, 112510, 112510, 112510, 112510, "$", "A", None, "Business operations specialist wage (CS)", "S46"),
    ("csmM", 40, 40, 40, 40, 40, 40, "#", "D", None, "Mid-Market + Enterprise accounts per CSM", ""),
    ("csmS", 250, 250, 250, 250, 250, 250, "#", "D", None, "SMB accounts per CSM", ""),
    ("gm", .55, .65, .72, .55, .65, .72, "%", "D", "value", "Gross margin on Clay revenue", "No reported figure; data-heavy margins (S52)"),
    ("launchsp", 40000, 40000, 40000, 40000, 40000, 40000, "$", "D", None, "Launch marketing per wave", ""),
    ("progbase", 25000, 25000, 25000, 25000, 25000, 25000, "$/qtr", "D", None, "Program spend embedded in the base inbound rate", "Spend above this buys extra MQLs at the CPL"),
    ("ref", .25, .20, .15, .25, .20, .15, "%", "B", "cost", "Partner fee on partner-sourced first-year ACV", "S50"),
    ("enab", 2000, 2000, 2000, 2000, 2000, 2000, "$/qtr", "D", None, "Enablement per active partner per quarter", ""),
    ("onb", .20, .15, .10, .10, .07, .05, "%", "B/D", "cost", "Onboarding subsidy, MM & Ent (% first-year ACV)", "S50; PS buyers need less help (adoption fit 4.5 vs 2.0)"),
    ("build", 250000, 250000, 250000, 250000, 250000, 250000, "$", "D", None, "Template & connector build (one-time)", ""),
    ("buildq", 4, 4, 4, 4, 4, 4, "qtrs", "D", None, "Quarters over which build is spent", ""),
    ("stop", 99, 99, 99, 99, 99, 99, "qtr", "D", None, "Program stop quarter (99 = never; 4 = stopped at Gate 1)", ""),
    ("life", 5, 5, 5, 5, 5, 5, "yrs", "D", None, "Years of post-2029 value from the 2029 ARR base", ""),
    # ---- valuation (review C3/C4)
    ("disc", .12, .12, .12, .12, .12, .12, "%", "D", None, "Discount rate (hurdle) for NPV", "Typical growth-software hurdle; tested 8–15% on Sensitivity"),
    ("hair", .30, .20, .10, .60, .45, .30, "%", "D", "indep", "Incrementality haircut: share of program wins that would have happened anyway", "Dist: few organic distributor wins (one case study). PS: Clay already markets to and sells in PS (S06, S32, S34)"),
    ("pb0", 16, 16, 16, 16, 16, 16, "months", "B", None, "Option 0: CAC payback of marginal spend in Clay's core motion", "Mid-market benchmark 14–18 months (S47)"),
    ("eff0", .60, .60, .60, .60, .60, .60, "x", "D", None, "Option 0: marginal efficiency vs benchmark (diminishing returns)", "Core already funded; marginal $ assumed 40% less efficient than average"),
    ("grr0", .90, .90, .90, .90, .90, .90, "%", "B", None, "Option 0: GRR of core customers", "SaaS Capital median 91% (S43)"),
    ("nrr0", 1.01, 1.01, 1.01, 1.01, 1.01, 1.01, "%", "B", None, "Option 0: NRR of core customers", "SaaS Capital median 101% (S43)"),
    ("csr0", .05, .05, .05, .05, .05, .05, "%", "D", None, "Option 0: customer-success cost as % of ARR", ""),
]
INPUT_KEYS = [x[0] for x in INPUTS]
MOTIONS = {
    # name: (fte_lead, pm1, pm2, eng1, eng2, pmm, prog, shared, varae)
    "lean": dict(fte_lead=1, pm1=0, pm2=0.5, eng1=1, eng2=0.5, pmm=0.25, prog=25000, shared=2, varae=1, covmult=1, partmult=1, buildmult=1),
    "sales": dict(fte_lead=1, pm1=1, pm2=1, eng1=1, eng2=0.5, pmm=0.5, prog=60000, shared=0, varae=0, covmult=1, partmult=1, buildmult=1),
    # Solution play: templates + data providers + partner listing inside the horizontal motion; no vertical lead
    "play": dict(fte_lead=0, pm1=0, pm2=0, eng1=0.5, eng2=0.25, pmm=0.1, prog=10000, shared=2, varae=1, covmult=0.5, partmult=0.5, buildmult=0.6),
}
HIRES = {
    "lean": [0] * 12,
    "sales": [2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 6, 6],
    "scale": [0, 0, 0, 0, 1, 2, 2, 2, 2, 2, 2, 2],
    "play": [0] * 12,
}


def params(seg, scen):
    """scen: 0 = Downside, 1 = Base, 2 = Upside."""
    off = 0 if seg == "dist" else 3
    return {x[0]: x[1 + off + scen] for x in INPUTS}


# ----------------------------------------------------------------------------- ACV
def acv(p, seg):
    emp = seg_emp_per_firm(seg)
    cp = p["dcprice"] * 12 / p["dcsize"]
    ap = p["acprice"] * 12 / p["acsize"]
    floor = (p["dcprice"] + p["acprice"]) * 12
    out = {}
    for b in BANDS:
        acc = emp[b] * p["salesshare"] * p["aps"] * p["usage"]
        con = acc * p["cpa"]
        cred = con * p["refresh"] * p["crc"] + acc * p["ai"] + acc * p["newacc"] * p["crnew"]
        ops = con * p["refresh"] + acc + acc * p["newacc"]
        use = cred * cp + ops * p["apo"] * ap
        if b == "E":
            out[b] = max(p["entcr"] * cp + p["entac"] * ap, use) * (1 - p["entdisc"])
        else:
            out[b] = max(floor, use)
    return out


# ----------------------------------------------------------------------------- engine
def run(seg, p, motion="lean", hires=None, stop=None, detail=False):
    m = MOTIONS[motion]
    hires = hires if hires is not None else HIRES[motion]
    stop = p["stop"] if stop is None else stop
    firms = seg_firms(seg)
    filt = p["f1"] * p["f2"] * (1 - p["exist"])
    sam = [{b: w[b] * filt for b in BANDS} for w in firms]
    samt = [sum(w.values()) for w in sam]
    A = acv(p, seg)
    waves = [p["w1"], p["w2"], p["w3"]]
    aeload = p["aeote"] / p["wshare"]
    capae = p["quota"] * p["prod"] / 4                     # ARR one ramped AE can close per quarter
    costper = aeload / p["quota"] / p["prod"]              # AE cost per $ of first-year ACV closed
    R = {k: [0.0] * QN for k in ["run", "act", "lo", "lae", "li", "parts", "lp", "opp", "op", "psh", "dwS", "dwM", "dwE",
                                 "ssw", "sadv", "capv", "fac", "lost", "nS", "nM", "nE", "new", "cS", "cM", "cE", "cT", "aS", "aM", "aE", "aT",
                                 "na", "rv", "gp", "ae", "aer", "kl", "ka", "kv", "km", "kp", "kr", "kn", "ko", "sm", "ke", "kb", "kc", "kt",
                                 "pl", "cum", "df", "pvi", "n0", "a0", "gp0", "pl0", "pv0", "mixS", "mixM", "mixE", "lpcum"]}
    for i in range(QN):
        qn = i + 1
        run_ = 1 if qn <= stop else 0
        ramps = [0.5 if qn == w else (1.0 if qn > w else 0.0) for w in waves]
        fats = [p["fatigue"] ** (max(0, qn - w) / 4) for w in waves]
        actb = {b: sum(sam[k][b] * ramps[k] for k in range(3)) for b in BANDS}
        act = sum(actb.values())
        efw = sum(samt[k] * ramps[k] * fats[k] for k in range(3))   # fatigue-weighted active accounts
        cov = p["cov"] * m["covmult"]
        lo = cov * p["meet"] * run_ * efw
        ae = hires[i]
        aer = 0.0 if i == 0 else (0.5 * hires[0] if i == 1 else hires[i - 2] + 0.5 * (hires[i - 1] - hires[i - 2]))
        avgfat = efw / act if act > 0 else 0
        aeacc_ = min(aer * p["aeacc"], act * (1 - cov))
        lae = aeacc_ * p["meet"] * p["aemult"] * avgfat * run_
        extra = max(0, m["prog"] - p["progbase"])
        xm = act * p["inbcap"] * (1 - math.exp(-extra / (p["cpl"] * act * p["inbcap"]))) if act > 0 and extra > 0 else 0
        li = (act * p["inb"] * min(1, m["prog"] / p["progbase"]) + xm) * run_
        parts = run_ * min(p["pcap"], p["p0"] + p["padd"] * (qn - 1)) * m["partmult"]
        lp = parts * p["lpp"]
        oo = (lo + lae) * p["m2o"]
        oi = li * p["i2o"]
        op = lp * p["p2o"]
        opp = oo + oi + op
        w = {"S": actb["S"], "M": actb["M"] * p["tiltM"] * p["propM"], "E": actb["E"] * p["tiltE"] * p["propE"]}
        ws = sum(w.values())
        for b in BANDS:
            R["mix" + b][i] = w[b] / ws if ws > 0 else 0
        R["opp"][i], R["op"][i] = opp, op
        R["run"][i], R["act"][i], R["lo"][i], R["lae"][i], R["li"][i], R["parts"][i], R["lp"][i] = run_, act, lo, lae, li, parts, lp
        R["lpcum"][i] = lp + (R["lpcum"][i - 1] if i else 0)

        def lagged(key_mix, lag):
            j = i - lag
            return R["opp"][j] * R[key_mix][j] * p["win"] if j >= 0 else 0
        dwS, dwM, dwE = lagged("mixS", p["lagSM"]), lagged("mixM", p["lagSM"]), lagged("mixE", p["lagE"])
        jj = i - p["lagSM"]
        psh = (R["op"][jj] / R["opp"][jj]) if jj >= 0 and R["opp"][jj] > 0 else 0
        ssw = dwS * p["ss"]
        sad = {"S": dwS * (1 - p["ss"]), "M": dwM, "E": dwE}
        sadv = sum(sad[b] * A[b] for b in BANDS)
        capv = (aer + m["shared"] * run_) * capae
        fac = min(1.0, capv / sadv) if sadv > 0 else 1.0
        nS, nM, nE = ssw + sad["S"] * fac, sad["M"] * fac, sad["E"] * fac
        new = nS + nM + nE
        prev = lambda k: R[k][i - 1] if i else 0.0
        cS = prev("cS") * p["grrS"] ** .25 + nS
        cM = prev("cM") * p["grrM"] ** .25 + nM
        cE = prev("cE") * p["grrE"] ** .25 + nE
        aS = prev("aS") * p["nrrS"] ** .25 + nS * A["S"]
        aM = prev("aM") * p["nrrM"] ** .25 + nM * A["M"]
        aE = prev("aE") * p["nrrE"] ** .25 + nE * A["E"]
        aT = aS + aM + aE
        na = nS * A["S"] + nM * A["M"] + nE * A["E"]
        rv = (prev("aT") + aT) / 8
        gp = rv * p["gm"]
        # costs
        fl = m["fte_lead"] * run_
        fa = ae * run_
        fp = (m["pm1"] if qn <= 4 else m["pm2"]) * run_
        fe = (m["eng1"] if qn <= 4 else m["eng2"]) * run_
        fm = m["pmm"] * run_
        fc = (cM + cE) / p["csmM"] + cS / p["csmS"]
        kl = (fl + fp) * p["wmgr"] / p["wshare"] / 4
        ka = fa * aeload / 4
        closedv = na - ssw * A["S"]                 # ARR closed by AEs (excludes self-serve)
        kv = max(0, closedv - aer * capae) * costper * m["varae"]
        km = fm * p["wmkt"] / p["wshare"] / 4
        nlaunch = sum(1 for wv in waves if qn == wv)
        kp = (m["prog"] + p["launchsp"] * nlaunch) * run_
        kr = na * psh * p["ref"]
        kn = parts * p["enab"]
        ko = (nM * A["M"] + nE * A["E"]) * p["onb"]
        sm = kl + ka + kv + km + kp + kr + kn + ko
        ke = fe * p["wdev"] / p["wshare"] / 4
        kb = p["build"] * m["buildmult"] / p["buildq"] if qn <= p["buildq"] else 0
        kc = fc * p["wcs"] / p["wshare"] / 4
        kt = sm + ke + kb + kc
        pl = gp - kt
        df = 1 / (1 + p["disc"]) ** (qn / 4)
        pvi = (gp * (1 - p["hair"]) - kt) * df
        # Option 0: redeploy the same non-CS program spend into the core motion
        n0 = (kt - kc) * 12 / (p["pb0"] * p["gm"]) * p["eff0"]
        a0 = prev("a0") * p["nrr0"] ** .25 + n0
        gp0 = (prev("a0") + a0) / 8 * p["gm"]
        pl0 = gp0 - (kt - kc) - a0 * p["csr0"] / 4
        pv0 = pl0 * df
        for k, v in dict(dwS=dwS, dwM=dwM, dwE=dwE, psh=psh, ssw=ssw, sadv=sadv, capv=capv, fac=fac, lost=sadv * (1 - fac), nS=nS, nM=nM, nE=nE, new=new,
                         cS=cS, cM=cM, cE=cE, cT=cS + cM + cE, aS=aS, aM=aM, aE=aE, aT=aT, na=na, rv=rv, gp=gp, ae=ae, aer=aer,
                         kl=kl, ka=ka, kv=kv, km=km, kp=kp, kr=kr, kn=kn, ko=ko, sm=sm, ke=ke, kb=kb, kc=kc, kt=kt, pl=pl,
                         cum=prev("cum") + pl, df=df, pvi=pvi, n0=n0, a0=a0, gp0=gp0, pl0=pl0, pv0=pv0).items():
            R[k][i] = v
    # outputs
    mixw = sum(R["n" + b][QN - 1] for b in BANDS)
    tot_new = sum(R["new"])
    grrb = (sum(sum(R["n" + b]) * A[b] * p["grr" + b] for b in BANDS) / sum(sum(R["n" + b]) * A[b] for b in BANDS)) if tot_new > 0 else 0.9
    aT = R["aT"][QN - 1]
    csr = (R["kc"][QN - 1] * 4 / aT) if aT > 0 else 0
    emb_flows = [aT * (p["gm"] - csr) * grrb ** k for k in range(1, int(p["life"]) + 1)]
    emb = sum(emb_flows)
    pvemb = sum(f * (1 - p["hair"]) / (1 + p["disc"]) ** (3 + k - 0.5) for k, f in enumerate(emb_flows, start=1))
    a0T = R["a0"][QN - 1]
    pvemb0 = sum(a0T * (p["gm"] - p["csr0"]) * p["grr0"] ** k / (1 + p["disc"]) ** (3 + k - 0.5) for k in range(1, int(p["life"]) + 1))
    npv = sum(R["pvi"]) + pvemb
    npv0 = sum(R["pv0"]) + pvemb0
    sm_tot = sum(R["sm"])
    na_tot = sum(R["na"])
    margcost = sum(R[k][i] for k in ("kv", "kp", "kr", "kn", "ko") for i in range(QN))
    cac = sm_tot / tot_new if tot_new else 0
    acvn = na_tot / tot_new if tot_new else 0
    sm29, new29, na29 = sum(R["sm"][8:]), sum(R["new"][8:]), sum(R["na"][8:])
    cac29 = sm29 / new29 if new29 else 0
    acv29 = na29 / new29 if new29 else 0
    out = dict(
        exit_arr=aT, cust=R["cT"][QN - 1], newc=tot_new, rev=sum(R["rv"]), acv=acvn, cac=cac, mcac=margcost / tot_new if tot_new else 0,
        pb=(cac / (acvn * p["gm"] / 12)) if acvn else 0, pb29=(cac29 / (acv29 * p["gm"] / 12)) if acv29 else 0,
        cost=sum(R["kt"]), cum=R["cum"][QN - 1], peak=min(R["cum"]), emb=emb, npv=npv, npv0=npv0, excess=npv - npv0,
        lost=sum(R["lost"]), acvS=A["S"], acvM=A["M"], acvE=A["E"], grrb=grrb,
        g_opps=sum(R["opp"][:4]), g_cust=sum(R["new"][:4]), g_pl=R["lpcum"][3], g_arr=R["aT"][3],
        g_acv=(sum(R["na"][:4]) / sum(R["new"][:4])) if sum(R["new"][:4]) else 0,
        g_dec=sum(R["opp"][:3]),
        g_mix=[sum(R["n" + b][:4]) for b in BANDS],
    )
    if detail:
        out["R"] = R
        out["A"] = A
        out["sam"] = sam
    return out


# ----------------------------------------------------------------------------- Monte Carlo
def _phi(z):
    return 0.5 * (1 + math.erf(z / math.sqrt(2)))


def _triq(u, a, c, b):
    """Inverse CDF of triangular(min a, mode c, max b)."""
    if b == a:
        return a
    fc = (c - a) / (b - a)
    if u < fc:
        return a + math.sqrt(u * (b - a) * (c - a))
    return b - math.sqrt((1 - u) * (b - a) * (b - c))


def sample_params(seg, rng, rho=0.5):
    """Draw one parameter set. Each sampled input is triangular(Downside, Base, Upside) in 'goodness' space,
    with a Gaussian copula: inputs in the same group share a latent factor (loading √rho)."""
    off = 0 if seg == "dist" else 3
    groups = {g: rng.gauss(0, 1) for g in ("demand", "value", "cost", "timing")}
    p = {}
    for x in INPUTS:
        name, grp = x[0], x[9]
        d, b, u_ = x[1 + off], x[2 + off], x[3 + off]
        if grp is None:
            p[name] = b
            continue
        if grp == "indep":
            z = rng.gauss(0, 1)
        else:
            z = math.sqrt(rho) * groups[grp] + math.sqrt(1 - rho) * rng.gauss(0, 1)
        q = _phi(z)
        if grp == "timing":
            p[name] = d if q < 1 / 3 else (b if q < 2 / 3 else u_)
            continue
        if d <= u_:
            p[name] = _triq(q, d, b, u_)
        else:
            p[name] = _triq(1 - q, u_, b, d)
    return p


def _poisson(rng, lam):
    if lam <= 0:
        return 0
    if lam > 50:
        return max(0, int(round(rng.gauss(lam, math.sqrt(lam)))))
    L, k, pr = math.exp(-lam), 0, 1.0
    while True:
        pr *= rng.random()
        if pr < L:
            return k
        k += 1


def gate_decision(out, p, th, rng):
    """Observe 2027 metrics with sampling noise and apply the Gate 1 rule. Returns 'stop' | 'scale' | 'continue'."""
    opps = _poisson(rng, out["g_opps"])
    cust = _poisson(rng, out["g_cust"])
    pl = _poisson(rng, out["g_pl"])
    mix = out["g_mix"]
    tm = sum(mix)
    if cust > 0 and tm > 0:
        counts = [0, 0, 0]
        for _ in range(cust):
            r, acc = rng.random() * tm, 0
            for k in range(3):
                acc += mix[k]
                if r <= acc:
                    counts[k] += 1
                    break
        acv_obs = (counts[0] * out["acvS"] + counts[1] * out["acvM"] + counts[2] * out["acvE"]) / cust
    else:
        acv_obs = 0
    dec = _poisson(rng, out["g_dec"])
    wins = sum(1 for _ in range(dec) if rng.random() < p["win"])
    win_obs = wins / dec if dec >= th["min_dec"] else None
    fails = sum([opps < th["opps"], cust < th["cust"], pl < th["pl"], acv_obs < th["acv"], (win_obs is not None and win_obs < th["win"])])
    if fails >= 2:
        return "stop"
    if acv_obs >= th["s_acv"] and cust >= th["s_cust"] and (win_obs is None or win_obs >= th["s_win"]):
        return "scale"
    return "continue"


def simulate(seg, n, seed, th, motion="lean"):
    import random
    rng = random.Random(seed)
    res = []
    for _ in range(n):
        p = sample_params(seg, rng)
        base = run(seg, p, motion=motion)
        rec = {"ungated": base["npv"], "ungated_x": base["excess"], "exit": base["exit_arr"]}
        if motion == "lean" and th is not None:
            d = gate_decision(base, p, th, rng)
            if d == "stop":
                o = run(seg, p, motion="lean", stop=4)
            elif d == "scale":
                o = run(seg, p, motion="lean", hires=HIRES["scale"])
            else:
                o = base
            rec.update(staged=o["npv"], staged_x=o["excess"], decision=d, staged_exit=o["exit_arr"])
        res.append(rec)
    return res


def summarize(vals):
    v = sorted(vals)
    n = len(v)
    q = lambda f: v[min(n - 1, int(f * n))]
    return dict(mean=sum(v) / n, p10=q(.10), p50=q(.50), p90=q(.90), ppos=sum(1 for x in v if x > 0) / n)


if __name__ == "__main__":
    for seg in ("dist", "ps"):
        for s, nm in enumerate(("Downside", "Base", "Upside")):
            o = run(seg, params(seg, s))
            print(seg, nm, {k: round(o[k]) if isinstance(o[k], float) and abs(o[k]) > 10 else (round(o[k], 2) if isinstance(o[k], float) else o[k]) for k in
                            ("exit_arr", "newc", "acv", "cac", "pb29", "cum", "npv", "npv0", "excess", "g_opps", "g_cust", "g_pl", "g_acv", "lost")})
        o = run(seg, params(seg, 1), motion="sales")
        print(seg, "SALES Base", {k: round(o[k]) for k in ("exit_arr", "newc", "acv", "cac", "cum", "npv", "excess", "lost")})
