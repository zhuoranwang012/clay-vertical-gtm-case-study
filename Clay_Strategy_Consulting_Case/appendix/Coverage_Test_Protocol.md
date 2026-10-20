# Contractor Coverage Test — Protocol

**U.S. Vertical Expansion & Go-to-Market Strategy for Clay** · Review issue M5 · preview of sprint test T1
Independent portfolio case study by Zhuoran (Zora) Wang · Not commissioned by or affiliated with Clay Labs Inc. · Protocol written 2026-10-19, **before any data were collected**

> **Status: sample drawn 2026-10-20; data collection not started.** No results exist yet. `Coverage_Test_Log.xlsx` holds the 120 sampled businesses but no observations; its Results sheet reads "Insufficient sample" until at least 60 firms are completed. Nothing in this case study may cite a coverage rate until the log holds real observations.

---

## 1. Why this test

The case's coverage claim has thin evidence (partner review M5):
- **H3** ("industrial buyer personas are poorly covered by LinkedIn-derived databases") rests on an **education-based proxy**: about 20% estimated LinkedIn use for contractor personas (05, [C]). That proxy measures LinkedIn use, not database coverage.
- The only direct evidence is **one vendor case study** (Regency Supply, S07).
- 08 v2 already withdrew v1's claim that Clay finds "the ones ZoomInfo and Apollo can't see" until it is measured.

This test measures it on a small scale, using only free, public tools.

**Question.** For contractors that electrical, HVAC and plumbing distributors (NAICS 4236 and 4237) sell to:
1. How often does a standard B2B database (Apollo) list a decision-maker?
2. How much does adding public-web sources — the kind a Clay waterfall or AI research agent reaches — raise that share?

## 2. Sample

| Item | Design |
|---|---|
| Frame | CSLB "List by Classification" (S56): free Excel download of currently renewed California licenses. Fields: license number, business name, address, phone, status, classifications. No emails. |
| Classifications | **C-10** electrical, **C-20** HVAC (warm-air heating, ventilating and air-conditioning), **C-36** plumbing. Download all three in one request. |
| Draw | Seeded (2026), stratified random sample: **100 primary** (34 / 33 / 33) plus **20 reserve** (7 / 7 / 6). `python3 appendix/coverage_test.py sample --input <download>`. A license holding several classes is sampled at most once. |
| Precision | n = 100 gives a 95% interval of about ±10 points on a share near 50%; about ±16 points per stratum. |
| Minimum | 60 completed firms for a verdict (the Results sheet enforces it). |
| Replacement | A Skipped row is replaced by the next unused Reserve row **of the same stratum**, in order. Skipped shares are reported, not hidden. |

**Frame as downloaded** (`CSLBSearchData_215814037.xlsx`, CSLB data portal, downloaded October 2026):
- 52,821 current (CLEAR) licenses holding C-10, C-20 or C-36.
- 51,240 of them have a California mailing address; the other 1,581 (mostly NV, AZ, TX, OR) are excluded so the frame is California contractors.
- Eligible by stratum: C-10 26,923; C-20 11,387; C-36 17,278. Multi-class licenses appear in more than one pool but can be drawn only once.
- Business types in the primary sample: 53 corporations, 42 sole owners, 4 LLCs, 1 partnership.

**Amendments made after the download and before any observation** (they change no threshold):
1. Licenses with a non-California mailing address are excluded (`--state CA`).
2. The CSLB business type is carried into the log (column Z). The Results sheet reports Apollo DM coverage for sole owners vs corporations, LLCs and partnerships, which is a better size proxy than "has a website". The verdict still uses the full sample.

**Why contractors.** Contractors are the customers that 4236/4237 distributors' reps prospect, and they are the persona that the H3 proxy and the Regency case are about.

**Not covered.** Plant and maintenance buyers of 4238 MRO distributors would need a different frame (for example, manufacturing establishments). That is a follow-up.

## 3. What is measured per firm

Record **Y/N flags and title categories only**. Full definitions are in the workbook's Definitions sheet.

| # | Field | Source | Role in the analysis |
|---|---|---|---|
| 1 | Website found | Google | Public-web layer |
| 2 | Google Business Profile | Google Maps | Public-web layer; the Regency method |
| 3 | **Web DM identifiable** | Website / Google profile | Public-web DM (Clay-style waterfall input) |
| 4 | Registry person of record | CSLB Instant License Check | Upper-bound layer; reported separately |
| 5 | Apollo company record | Apollo free account | Database layer |
| 6 | Apollo contacts listed (0 / 1–2 / 3+) | Apollo | Depth |
| 7 | **Apollo DM listed** | Apollo | **Primary metric** |
| 8 | Apollo DM email shown available | Apollo (do not reveal) | Contactability |
| 9 | LinkedIn company page | LinkedIn / Google site search | LinkedIn layer |
| 10 | LinkedIn DM profile | LinkedIn / Google site search | Tests the 20% proxy |

**Decision-maker (DM):** owner, founder, president/CEO, principal or partner, VP or general manager, operations manager, purchasing manager, estimator, or project manager.

**Derived (formulas):**
- **Apollo ∪ web DM.** A DM is found in Apollo **or** on the website or Google profile.
- Apollo ∪ web ∪ registry.
- Apollo ∪ LinkedIn.

## 4. Procedure (about 4 minutes per firm; about 7 hours for 100)

1. **Download.** Get the CSLB list for C-10, C-20 and C-36, then run the sampler. If the .xls will not open, re-save it as .xlsx or .csv, or install `xlrd`.
2. **Pilot.** Work through IDs P001–P010.
   - Tighten any ambiguous definition now, and log the change with a date in the Definitions sheet.
   - **After the pilot, definitions are frozen.**
3. **Per firm, in this fixed order** (database first, so the web search cannot bias the Apollo check):
   1. **Apollo.** Search Companies by business name; confirm the match on city or domain. Open People and record the count and whether any listed person has a DM title. Do not reveal emails or phones (this keeps within free-plan credits and avoids collecting personal data).
   2. **LinkedIn.** Search for the company page, then for people at the company with a DM title. Searching Google with `site:linkedin.com` is acceptable.
   3. **Google.** Search for the business name and city. Record the website and the Google Business Profile. Record "Web DM identifiable" only if a DM's name **and** role appear on the website or profile.
   4. **CSLB license detail.** Record whether a qualifier, officer or owner is listed.
   5. **Time.** Enter the minutes spent, then set Status = Done.
4. **Skip rules.**
   - *Out of business:* no web presence **and** a closed or "permanently closed" listing.
   - *Cannot identify:* the name is too generic to match within 5 minutes.
5. **Re-check.** On a different day, re-record three fields for 15 random Done rows (Recheck sheet) without looking at the first entry. If agreement on any field is below 85%, that definition is ambiguous: fix it and re-run that field on all rows.

## 5. Pre-registered analysis and verdict

These are fixed now, before data collection. They are coded in the Results sheet.

| Metric | Definition |
|---|---|
| Primary | Apollo DM coverage = share of completed firms with Apollo DM listed = Y, with a Wilson 95% interval |
| **Uplift (T1 analogue)** | (Apollo ∪ web DM) ÷ Apollo DM |
| Conservative uplift | Lower bound of the union ÷ upper bound of Apollo (both Wilson 95%) |
| Upper-bound uplift | (Apollo ∪ web ∪ registry) ÷ Apollo. Reported, **not** used for the verdict, because CSLB lists a person of record for nearly every license by design |
| Proxy check | Measured LinkedIn DM coverage − the 20% proxy used in 05/06 |

| Result | Verdict | What changes in the case |
|---|---|---|
| n < 60 | Insufficient sample | Nothing |
| Uplift ≥ 1.5× **and** conservative uplift ≥ 1.5× | **Supported (robust)** | Keep coverage-led positioning and cite the measured rate; competitive data-coverage score (04) re-based on measurement; H3 "supported (measured)" |
| Uplift ≥ 1.5× only | **Supported (point estimate)** | Claim a measured gap; show the interval |
| 1.2× ≤ uplift < 1.5× | **Partial** | Soften to "adds coverage on a material share of contractors"; T1 must still run at scale |
| Uplift < 1.2× | **Not supported** | Withdraw coverage-led positioning (roadmap risk R4); H3 "not supported"; distribution's tie-breaker in 06 weakens |

**Interpretation limits.** This is a small-scale **preview** of T1, not a substitute for it:
- **Free tier.** The Apollo free plan gives basic search over the same database (S57); paid filters could find somewhat more.
- **Geography.** California only.
- **No firm size.** CSLB gives no employee counts, so the sample over-represents sole proprietors relative to distributors' 100–499-employee ICP. Coverage among firms with a website is reported as a rough size proxy.
- **Web search is not Clay.** "Web DM identifiable" approximates what a Clay waterfall or AI agent could reach. It does not measure Clay itself.

## 6. Data handling and compliance

- **Business data.** Only business-level registry fields, which CSLB publishes, are stored.
- **No personal data about individuals.** No names, emails, phone numbers or profile URLs are recorded; notes must not contain them.
- **Manual lookups only**, within each site's terms. No scraping or automation of LinkedIn, Apollo or Google. Use your own free Apollo account.
- **No outreach.** Nobody is contacted; this is a desk check.
- **Raw download.** Delete it after the test if it is no longer needed. The log keeps only the sampled businesses.

## 7. What gets updated after the test

| File | Update |
|---|---|
| `04_Competitive_Analysis.xlsx` | Data coverage of industrial buyers: replace the Clay 5 / Apollo 2 judgment with the measured rates (evidence class B→A/C) |
| `05_Customer_Needs_Assessment.md` | H3 verdict and the 20% proxy, with measured LinkedIn coverage beside it |
| `06_Segment_Prioritization.xlsx` | Competitive-position input for distribution, if the measured gap differs from the proxy |
| `08_GTM_Strategy.md` / deck slides 6, 7, A2 | Positioning wording per the verdict table above |
| `Research_Sources.csv` | Add the test as a primary dataset (date range, n, limitations) |
