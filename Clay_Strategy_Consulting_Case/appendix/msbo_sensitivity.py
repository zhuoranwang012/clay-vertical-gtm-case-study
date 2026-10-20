"""
M3 sensitivity — manufacturers' sales branches and offices (MSBOs) inside NAICS 423 "distributors".
U.S. Vertical Expansion & GTM Strategy for Clay - independent case study, not affiliated with Clay Labs Inc.

NAICS 423 includes manufacturers' sales branches (TYPOP 21) and sales offices (TYPOP 22). They are
manufacturers' own sales arms, not independent distributors. Census reports type of operation by
establishment only, not by firm size, so the share of 500+ employee "distributor" firms that are
really MSBOs can only be bounded:

  low  = MSBOs are spread like all establishments (MSBO share of all establishments / employees)
  high = every MSBO establishment belongs to a 500+ employee firm (MSBO ÷ 500+ firms' establishments / employees)

The script removes that share from the 500+ band of NAICS 4236–4238 (firms by the establishment share,
employment by the employee share), re-runs the v2 engine (clay_engine.py) and writes
appendix/msbo_sensitivity.csv. Base inputs in 07 are unchanged.

Sources: EC2242TYPOP (S58) → appendix/raw_census_ec2242typop_2022_423.csv; EC2200SIZEEMPFIRM (S14).
"""
import copy
import csv
import os

import clay_engine as ce

HERE = os.path.dirname(os.path.abspath(__file__))
CODES = ["4236", "4237", "4238"]

typ = {}
for r in csv.DictReader(open(os.path.join(HERE, "raw_census_ec2242typop_2022_423.csv"))):
    n, t = r["NAICS2022"], r["TYPOP"]
    typ.setdefault(n, {})[t] = (int(r["ESTAB"] or 0), int(r["EMP"] or 0))
size = {(r["naics2022"], r["size_band"]): r for r in csv.DictReader(open(os.path.join(HERE, "raw_census_ecnsize_2022_firm_size.csv")))}


def msbo(n):
    return tuple(sum(typ[n].get(t, (0, 0))[i] for t in ("21", "22")) for i in (0, 1))


def total(n):
    return tuple(sum(v[i] for v in typ[n].values()) for i in (0, 1))


def e500(n):
    r = size[(n, "500+")]
    return int(r["establishments"]), int(r["employees"])


def shares(mode):
    out = {}
    for n in CODES:
        m, d = msbo(n), (total(n) if mode == "low" else e500(n))
        out[n] = (min(1.0, m[0] / d[0]), min(1.0, m[1] / d[1]))
    return out


BASE = copy.deepcopy(ce.CENSUS)


def apply(mode):
    ce.CENSUS = copy.deepcopy(BASE)
    if mode == "base":
        return
    for n, (fs, es) in shares(mode).items():
        f, e = (list(x) for x in ce.CENSUS[n])
        f[4] *= 1 - fs
        e[4] *= 1 - es
        ce.CENSUS[n] = (f, e)


rows = []
for mode in ("base", "low", "high"):
    apply(mode)
    ent = sum(w["E"] for w in ce.seg_firms("dist"))
    for motion, mlabel in (("lean", "Lean pilot"), ("play", "Solution play"), ("sales", "Sales-led team")):
        for sc, sname in ((1, "Base"), (2, "Upside")):
            o = ce.run("dist", ce.params("dist", sc), motion)
            rows.append(dict(case=mode, enterprise_firms_4236_4238=round(ent, 1), motion=mlabel, scenario=sname,
                             npv=round(o["npv"]), excess_vs_option0=round(o["excess"]), exit_arr=round(o["exit_arr"]),
                             avg_first_year_acv=round(o["acv"]), enterprise_acv=round(o["acvE"])))
apply("base")

with open(os.path.join(HERE, "msbo_sensitivity.csv"), "w", newline="") as f:
    w = csv.DictWriter(f, fieldnames=list(rows[0]))
    w.writeheader()
    w.writerows(rows)

print("MSBO shares (establishments, employees) — low:", {k: tuple(round(x, 3) for x in v) for k, v in shares("low").items()})
print("MSBO shares (establishments, employees) — high:", {k: tuple(round(x, 3) for x in v) for k, v in shares("high").items()})
for r in rows:
    print(f"{r['case']:5} ent {r['enterprise_firms_4236_4238']:6} {r['motion']:15} {r['scenario']:6} "
          f"NPV {r['npv'] / 1e6:6.2f}M  excess {r['excess_vs_option0'] / 1e6:6.2f}M  exit ARR {r['exit_arr'] / 1e6:5.2f}M  ACV {r['avg_first_year_acv']:,}")
print("wrote appendix/msbo_sensitivity.csv")
