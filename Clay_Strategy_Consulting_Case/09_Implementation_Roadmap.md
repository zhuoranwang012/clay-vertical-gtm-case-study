# 09 — Implementation Roadmap v2: Sprint, Decision Gates, Owners and KPIs

**U.S. Vertical Expansion & Go-to-Market Strategy for Clay** · Phase 5 (v2, rewritten 2026-10-19 after partner review)
Independent portfolio case study by Zhuoran (Zora) Wang · Not commissioned by or affiliated with Clay Labs Inc. · Research reference date: October 18, 2026

> **Scope.** Companion to `08_GTM_Strategy.md` v2. v1 (a $1.1M pilot with four gates) is archived in `appendix/archive/`.
>
> **Thresholds.** Derived from the v2 model paths in `07_Financial_Model.xlsx` and marked **[E]**. Model reference values are **[C/D]** scenario outputs, not Clay forecasts.
>
> **Owners.** Named by business function only; this case does not know Clay's reporting structure.

---

## 1. The plan on one page

| Period | Stage | Spend cap | Decision at the end | Default |
|---|---|---:|---|---|
| Nov 2026 – Mar 2027 | **A. Validation sprint** (distribution: 4236 and 4238) | ~$0.2M | **Gate 1 (end of Mar 2027):** solution play or stop | Stop |
| Apr – Dec 2027 | **B. Solution play** (only if Gate 1 = Go; no vertical team) | ~$0.3M | **Gate 2 (Q4 2027):** continue, consider a lean pilot, or stop | Stop unless on the Upside track |
| 2028 | **C. Continue the play** or, at most, a lean pilot (only if Gate 2 = scale) | Re-planned at Gate 2 | **Gate 3 (Q4 2028):** keep, fold into the horizontal motion, or stop | Fold into horizontal |

**What Clay keeps whatever happens:**
- Distribution templates and contractor/permit data sources in the horizontal library.
- A measured answer on contractor coverage versus Apollo and ZoomInfo.
- Calibrated usage and conversion data for any future vertical case, including the persona-vertical alternative (08 §6).

**Maximum loss if the sprint says stop:** ~$0.2M. In v1, the Downside loss was capped at ~$1.0M.

---

## 2. Stage A — validation sprint

### Workstreams and timing

| Weeks (from early Nov 2026) | Workstream | Output | Owner (function) |
|---|---|---|---|
| 1–2 | **Set-up.** Confirm executive sponsor, cap, decision rule and test protocols in writing | Signed sprint charter | Executive sponsor; Finance |
| 1–8 | **T1 Coverage test.** 200 contractor/plant accounts from 4236/4238 prospects' territories; Clay waterfall + Maps + permits vs best single database (incl. Apollo) | Match-rate table, verified by a second reviewer | GTM engineering; Data partnerships |
| 2–6 | **Template MVP.** Contractor discovery, branch scoring, e-commerce registration vetting | 3 templates in the library | GTM engineering |
| 3–16 | **T2/T3 Outbound test.** Clay-run sequences to ≥ 600 ICP accounts (300 per sub-industry); log meetings and opportunities | Meeting rate; meeting → opportunity rate, by sub-industry | Marketing ops; Sales (existing AEs) |
| 3–14 | **T4 Design partners.** 5+ distributors with 100+ employees and a CRM; paid Growth annual at list, sized by territory; promotional credits only | Committed credits → implied ACV | Sales leadership; Deal desk |
| 4–12 | **T5 Partners.** Approach distribution consultancies and ERP/CRM implementers | Partners committing named consultants | Partnerships |
| 2–10 | **Interviews.** 15–20 with distributor sales leaders, sales ops and reps (`appendix/Interview_Guide.md`) | Validated pains, buying group, willingness to pay | Product marketing |
| 17–20 | **Gate 1 pack.** Results vs thresholds, verified by Finance | Go / Stop memo | Product marketing (author); Finance (verifier) |

### Gate 1 — solution play or stop (end of March 2027)

| Test | Threshold [E] | Model anchor (Base → Upside, 07 v2) |
|---|---|---|
| T1 Coverage uplift vs best single database | ≥ 1.5× | Competitive position; meeting rate |
| T2 Meeting rate on ≥ 600 sequenced accounts | ≥ 4.0% | 3.0% → 4.5% |
| T3 Meeting → opportunity | ≥ 45% | 40% → 50% |
| **T4 Implied Mid-Market ACV from committed credits (≥ 5 design partners)** | **≥ $14.6K** | $9.8K → $14.6K (usage 1.0× → 1.5×) |
| T5 Implementation partners committing named consultants | ≥ 3 | 2 → 3 at launch |

**Rule [E]:**
- **Go** if T4 passes **and** at least three of T1, T2, T3 and T5 pass.
- **One more quarter, at no additional cap,** if T4 passes and only two others pass.
- **Stop** otherwise.

**Why T4 is mandatory:** usage intensity is the largest single driver of NPV (07 v2 › Sensitivity). Without it, no other result can make the economics work.

---

## 3. Stage B — solution play (April–December 2027, only after Go)

### Design

| Element | Design | Owner (function) |
|---|---|---|
| People | 0.5 GTM engineer (0.25 from 2028), 0.1 product marketing; no vertical lead, no dedicated AEs | Product (templates); Marketing |
| Demand | Clay-run outbound at half the pilot's coverage; templates; NAED / ASA content | Marketing ops |
| Closing | Existing AEs, variable cost per $ closed | Sales leadership |
| Delivery | Listed partners run a customer-paid quickstart | Partnerships |
| Spend | ≈ $0.3M for Apr–Dec 2027 (Base solution-play run-rate in 07 v2) | Finance (tracks against the cap) |

### Gate 2 — continue, scale or stop (Q4 2027)

The solution play is NPV-negative in Base (−$0.54M) and positive only on its Upside track (+$1.11M). So it continues only if it is on that track.

Model reference for the 2027 solution-play path [C/D]:

| Metric (2027) | Downside | Base | Upside | Threshold to continue [E] |
|---|---:|---:|---:|---:|
| Opportunities created | 8 | 21 | 57 | **≥ 40** |
| New customers | 1 | 3 | 11 | **≥ 7** |
| Average first-year ACV | $6.9K | $8.5K | $11.7K | **≥ $10K** |
| ARR, end of Q4 2027 | $5K | $26K | $127K | ≥ $80K |

Outcomes:
- **Continue the play** if all four thresholds are met.
- **Consider a lean pilot** (a vertical lead, still no dedicated AEs) only if the thresholds are met **and** borrowed AE capacity is binding: sales-assisted demand above ~80% of shared capacity for two quarters.
- **Stop** otherwise, and keep the templates in the library.

Re-run `07_Financial_Model.xlsx` with the measured values before any lean-pilot decision. Fund the pilot only if its re-estimated NPV beats Option 0.

### Gate 3 — keep, fold or stop (Q4 2028)

- **Keep** a distinct solution play only if 2028-cohort CAC payback is ≤ 18 months (the mid-market benchmark is 14–18, S47) and first-cohort net revenue retention is ≥ 100%.
- **Otherwise fold** the assets into the horizontal motion with no vertical cost line.

---

## 4. Decision owners by business function

| Decision | Accountable | Responsible | Consulted |
|---|---|---|---|
| Fund the sprint; set the cap and decision rule; all gates | Executive sponsor (GTM / revenue leadership) | Product marketing (sprint lead) | Finance; Product; Sales |
| Verify gate metrics and spend | Finance / FP&A | RevOps / analytics | Sprint lead |
| Test protocols (T1–T5), statistical reading | Product marketing | RevOps; GTM engineering | Finance |
| Templates, data-provider trials | Product | GTM engineering | Data partnerships; Legal |
| AE time for design partners, transfer cost | Sales leadership | Deal desk | Finance |
| Partner approach and terms | Partnerships | Partnerships | Legal |
| Data compliance (contact data, permits, scraping terms) | Legal / security | — | Product; Data partnerships |

---

## 5. KPIs

The required commercial KPIs, defined once and tracked from the sprint onward. Targets are the Upside-track values, because Base does not clear the hurdle [E].

| KPI | Definition | Sprint (Gate 1) | Solution play 2027 (Gate 2) | Source of truth / cadence |
|---|---|---|---|---|
| **Qualified pipeline** | Opportunities × expected first-year ACV | Directional only | ≥ 40 opps × ≥ $10K ≈ ≥ $0.4M | CRM / weekly |
| **Lead-to-opportunity conversion** | Opportunities ÷ meetings, introductions and MQLs | Meeting → opp ≥ 45% | ≥ 45% blended (Base 37%, Upside 48%) | CRM / monthly |
| **Win rate** | Won ÷ decided opportunities | Not measured (sample too small) | ≥ 30% on ≥ 10 decided opportunities (Upside 30%; Base 21%) | CRM / monthly |
| **CAC** | Program S&M ÷ new customers; fully loaded and marginal | n/a | Marginal ≤ ~$10K (Upside solution play) | Finance / quarterly |
| **Revenue per customer** | Average first-year ACV | Implied Mid-Market ACV ≥ $14.6K | Average ≥ $10K | Finance / monthly |
| **Payback** | Cohort CAC ÷ (first-year ACV × gross margin ÷ 12) | n/a | ≤ 24 months on the 2027 cohort | Finance / quarterly |
| **Adoption** | Time to first live workflow; workflows live at day 90; credit consumption vs commitment | Design partners: first workflow ≤ 30 days | ≥ 2 workflows by day 90; ≥ 80% of committed credits by month 6 | Product analytics / monthly |
| **Retention** | Logo retention and net revenue retention of first cohorts | n/a | Leading indicator only in 2027; tested at Gate 3 (NRR ≥ 100%) | Finance / quarterly |
| **Coverage** | Verified decision-maker contacts vs best single database | ≥ 1.5× (T1) | Maintained on new accounts | GTM engineering / per test |

---

## 6. Risk register

L = likelihood, I = impact; judgments [C].

| # | Risk | L | I | Early warning | Mitigation | Owner |
|---|---|:-:|:-:|---|---|---|
| R1 | Sprint too small to detect a real Upside (false stop) | M | M | Borderline T2/T3 with T4 passing | One extra quarter at no extra cap; several tests must pass together | Sprint lead |
| R2 | Sunk-cost drift: anecdotes override the rule | M | H | Requests to hire before Gate 2 | Rule and owners fixed in writing; Finance verifies | Executive sponsor |
| R3 | Design partners sign only at a discount | M | H | Price objections | Promotional credits only; a discounted deal does not count toward T4 | Deal desk |
| R4 | Coverage edge smaller than claimed | M | H | T1 below 1.5× | Stop; withdraw coverage-led positioning | Data partnerships |
| R5 | Shared AEs deprioritize small sprint deals | M | M | Slow meeting follow-up | Named AE pool; transfer-cost credit | Sales leadership |
| R6 | Channel conflict with existing agency partners | L | L | Account disputes | Deal registration | Partnerships |
| R7 | Data compliance (scraping, permits, contact data) | L | H | Provider or customer complaints | Licensed marketplace providers only; legal review of templates | Legal |
| R8 | A better adjacency is ignored (persona vertical, data play) | M | M | — | Size and model in parallel (08 §6) | Product marketing |

---

## 7. What survives from v1

- **Kept:** the hybrid motion's building blocks (Clay-run outbound, templates, partners, borrowed AEs); list pricing with territory-based credit sizing; the ICP tilt to 100–499-employee, CRM-using, contractor-facing distributors; and gate discipline with several indicators per gate.
- **Removed:** the $1.1M pilot budget; the vertical lead; AE hiring triggers before Gate 2; and the claim that Clay sees accounts ZoomInfo and Apollo cannot, until T1 measures it.
