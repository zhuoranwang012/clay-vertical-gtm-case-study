# 09 — Implementation Roadmap, KPIs and Go/No-Go Gates

**U.S. Vertical Expansion & Go-to-Market Strategy for Clay** · Phase 5
Independent portfolio case study by Zhuoran (Zora) Wang · Not commissioned by or affiliated with Clay Labs Inc. · Research reference date: October 18, 2026

> Companion to `08_GTM_Strategy.md`. Gate thresholds are derived from the scenario runs in `07_Financial_Model.xlsx` (Engine Runs, Scenarios) and are marked **[E]** recommendations. Model reference values are **[C/D]** scenario outputs, not Clay forecasts. Owners are named **by business function only**; this case does not know or assume Clay's actual reporting structure.

---

## 1. The plan on one page

| Period | Phase | What happens | Decision at the end |
|---|---|---|---|
| Oct 2026 – Jan 2027 | **Validate** (no-regret, staff time only) | Validation interviews, contractor match-rate test, 3 paid design partners, 2 implementation partners, budget cap | **Gate 0** (Jan 2027): launch the pilot or not |
| Q1 – Q4 2027 | **Pilot wave 1** — electrical (4236) | Templates v1–v2, Clay-run outbound, NAED channel, partner certification, first case studies; mid-year checkpoint | **Gate 1** (Q4 2027 review, decision by Jan 2028): stop / continue lean / scale |
| Q1 – Q4 2028 | **Extend** — PHVAC (4237) Q1, MRO (4238) Q3 | Wave launches; AE hires only if the capacity trigger fires; retention and expansion playbooks | **Gate 2** (Q4 2028): scale to a standing vertical / hold lean / sunset |
| 2029 | **Scale or harvest** | Scale track: dedicated team, native ERP connector decision, vertical data licensing. Hold track: no new hires, optimize retention | **Gate 3** (Q4 2029): graduate to standing vertical or fold into the horizontal motion |

**Budget discipline:**
- 2027 spend is capped at **$1.1M**, released quarterly. Base model: $1.04M; Upside: $1.31M if the scale branch is taken at Gate 1.
- The Downside loss is capped at **~$1.0M** by stopping at Gate 1 [C].

---

## 2. Pilot design

### Pilot objectives — what the pilot must prove

| # | Pilot hypothesis | Why it matters (from 07 sensitivity) | Measured by |
|---|---|---|---|
| P1 | Clay finds materially more contractor/plant accounts and contacts than a single database | Basis of differentiation (H3 is a proxy today) | Match-rate test at Gate 0; uplift ratio in live deployments |
| P2 | Mid-market distributors commit to ACVs of ~$15K+ by running several workflows | Usage intensity is the **largest** swing in total value ($1.3M) | Average first-year ACV; credit consumption vs commitment |
| P3 | Funnel converts near or above benchmarks (win rate ≥ 28%; lead→opp ≥ 22%) | Win rate, SQL→opp and qualification are the next-largest swings ($0.6–0.8M each) | CRM funnel reporting |
| P4 | Partners can source and deliver without Clay engineers | Builder capacity is the binding constraint (H4); partner network swing $0.6M | Producing partners; leads per partner; time to first live workflow |
| P5 | Customers retain and expand | GRR/NRR swing $0.3–0.4M; drives post-2029 value | Usage trend; 90-day health; first renewals |

### Scope and team (Base)

| Element | Specification |
|---|---|
| Market | Wave 1: ~1,280 serviceable electrical distributors (4236). ICP core: 100–499 employees with CRM |
| Team | 1 vertical lead (player-coach, owns partners in 2027); 1 GTM/solutions engineer (0.5 from 2028); 0.25 product marketing; shared AEs from the existing commercial team at a per-deal transfer cost; pooled CS |
| Partners | 2 at launch, +1 per quarter, cap ~10 |
| Assets | Distribution Territory Kit: contractor discovery (Maps + permits), branch territory scoring, e-commerce registration vetting, job-change alerts; credit-sizing calculator; ERP data path via warehouse/HTTP |
| Commercials | Growth annual at list; no vertical discount; partner Quickstart paid by the customer; Clay subsidy 10–20% of first-year ACV for mid-market/enterprise |
| Not in scope | Dedicated AEs, native ERP connector, manufacturing or professional-services programs |

---

## 3. First 90 days (Oct 2026 – Jan 2027): Gate 0 preparation

| Week | Activity | Function |
|---|---|---|
| 1–3 | Confirm executive sponsor; set the budget cap and gate calendar; agree a transfer-cost rule for shared AEs | Executive sponsor, Finance, Sales leadership |
| 1–6 | Run 15–20 validation interviews with distributor sales leaders, sales ops and reps (4236 first, some 4237) using `appendix/Interview_Guide.md` | Vertical lead, PMM |
| 2–8 | **Contractor match-rate test:** 200 contractor/plant accounts sampled from 3 prospects' territories. Compare Clay waterfall + Google Maps + permits against the best single database (verified email/phone/decision-maker) | GTM engineering, Data partnerships |
| 3–10 | Recruit design partners: 3 distributors with 100+ employees and CRM; paid Growth annual at list with 90-day success criteria | Vertical lead, Sales leadership |
| 3–10 | Recruit 2 implementation partners (distribution consultancy or ERP/CRM implementer); agree fees and certification | Partnerships, Legal |
| 6–12 | Build template v1 (contractor discovery, e-commerce vetting); instrument funnel and usage dashboards | GTM engineering, RevOps |
| 10–13 | Approach NAED for content/event participation (no contact made in this study) | Marketing |
| 13 | **Gate 0 review** | Executive sponsor (decides), Finance (verifies) |

---

## 4. Milestones by quarter

| Quarter | Milestones | Model reference, Base [C] |
|---|---|---|
| **Q4 2026** | Gate 0 work (§3) | — |
| **Q1 2027** | Wave 1 launch; template v1 live; Clay-run outbound to ~40% of active SAM per quarter; 2 partners certified; design partners live | Partners: 2 |
| **Q2 2027** | Template v2 (branch scoring, permit signals); first non-design-partner wins; **mid-year checkpoint** | H1: ~57 leads, ~13 opps, ~1 customer |
| **Q3 2027** | First public case study; 4–5 partners; credit-sizing calculator in every deal | — |
| **Q4 2027** | **Gate 1** evidence pack; wave-2 readiness (ASA channel, PHVAC templates) | 2027: ~141 leads, ~31 opps, ~6 customers, ARR ~$85K |
| **Q1 2028** | Wave 2 (4237) launch, if Gate 1 = continue/scale. Scale branch: hire AE #1 | Upside run: AE hires Q1–Q2 2028 |
| **Q2 2028** | Partner manager (0.5 FTE); retention playbook from first renewals. Scale branch: AE #2 | — |
| **Q3 2028** | Wave 3 (4238 MRO) launch | Active SAM ~5,170 accounts by Q4 2028 |
| **Q4 2028** | **Gate 2** evidence pack; ERP connector and vertical data-licensing business cases | ARR ~$334K; ~22 customers |
| **2029** | Scale or hold execution | Q4 2029 ARR: Base ~$0.86M; Upside + scale-up ~$6.0M |
| **Q4 2029** | **Gate 3** strategic review | — |

---

## 5. Decision owners by business function

Functions are generic. Map them to whatever roles hold these decision rights.

| Decision | Accountable (decides) | Responsible (does the work) | Consulted |
|---|---|---|---|
| Launch pilot; budget cap; all Go/No-Go gates | Executive sponsor (GTM / revenue leadership) | Vertical lead | Finance; Product; Partnerships |
| Gate metric verification and spend tracking | Finance / FP&A | RevOps / analytics | Vertical lead |
| ICP, positioning, messaging | Product marketing | Vertical lead | Sales leadership; CS |
| Shared AE allocation, transfer cost, AE hiring trigger | Sales leadership | Vertical lead; deal desk | Finance |
| Pricing exceptions, credit sizing, promotional credits | Pricing / deal desk (with Finance) | Vertical lead | Sales leadership |
| Partner recruitment, terms, certification, deal registration | Partnerships / ecosystem | Partner manager (vertical lead in 2027) | Legal; Finance; Sales |
| Template roadmap; ERP data path; native connector decision | Product | GTM / solutions engineering | CS; Partnerships |
| New data providers (e.g., construction projects, license registries) | Data partnerships | GTM engineering | Legal; Product |
| Onboarding, adoption, renewals, expansion | Customer success | CS team and partners | Vertical lead |
| Data privacy, partner and data-license contracts | Legal / security | — | Partnerships; Data partnerships |

---

## 6. KPI framework

**Reference values** are model outputs [C].
- **Scale signal:** Upside-track values.
- **Floor:** Downside-track values.
- **Targets** are recommendations [E].

### Commercial KPIs (required)

| KPI | Definition | Base 2027 | Base 2028 | Scale signal | Floor | Cadence / owner |
|---|---|---:|---:|---:|---:|---|
| **Qualified pipeline** | Opportunities created × expected first-year ACV | ~$0.43M | ~$1.12M | ≥ $1.5M in 2027 (Upside ~$2.1M) | < $0.2M (Downside ~$0.08M) | Weekly / RevOps |
| **Lead-to-opportunity conversion** | Opportunities ÷ leads (all sources) | 22% | 22% | ≥ 30% (Upside 33%) | < 15% (Downside 14%) | Monthly / RevOps |
| **Win rate** | Won ÷ decided opportunities (won + lost) | 28% | 28% | ≥ 32% (Upside 35%) | < 20% | Monthly; judge on ≥ 10 decided opps / Sales leadership |
| **Customer acquisition cost** | Program S&M cost ÷ new customers (includes variable AE, partner fees, onboarding subsidy) | ~$93K (launch year) | ~$51K | ≤ $25K by 2028 (Upside ~$20K) | > $150K | Quarterly / Finance |
| **Revenue per customer** | Average first-year ACV of new customers; ARR ÷ customers | $13.8K first-year ACV | $13.8K | ≥ $16K (Upside $19.4K) | < $10K | Monthly / Finance |
| **Payback period** | Cohort CAC ÷ (first-year ACV × gross margin ÷ 12) | ~124 mo (launch year) | ~68 mo | ≤ 24 mo (Upside 2028 ~17 mo) | > 120 mo | Quarterly / Finance |

### Partner KPIs

| KPI | Base reference | Target / trigger |
|---|---|---|
| Producing partners (≥ 1 lead in quarter) | 5 by Q4 2027; 10 by 2029 | ≥ 4 by Gate 1 |
| Leads per producing partner per quarter | 3 [D] | ≥ 2 (floor), ≥ 4 (scale) |
| Partner-sourced share of opportunities | ~36% in 2027 [C] | ≥ 30% |
| Time from partner kickoff to customer's first live workflow | — | ≤ 30 days [E] |

### Adoption and retention indicators

Targets are recommendations [E]. The model has no adoption inputs, so these are early proxies for the GRR/NRR assumptions.

| Indicator | Target | Why |
|---|---|---|
| Time to first live workflow | ≤ 30 days from signature | Builder-capacity risk shows up here first |
| Workflows live at day 90 | ≥ 2 per customer (≥ 3 for mid-market+) | Multi-workflow use is the ACV lever |
| Credit consumption vs commitment | ≥ 80% by month 6, without overage disputes | Leading indicator of renewal size and expansion |
| CRM sync active, and % of territory accounts enriched | Sync active at 100% of customers; ≥ 70% of target accounts enriched | Proves value pillar 3 |
| Reps working Clay-sourced accounts | ≥ 50% of reps with activity on ≥ 1 Clay-sourced account per month | Value reaches the field |
| Logo retention at first renewal | ≥ 88% (Base GRR SMB 88%, MM 90%) | Validates the retention assumptions |
| Net revenue retention, first cohorts | ≥ 100% | Base NRR 100–115% by band; SaaS Capital median 98–104% (S43) |
| Contractor match-rate uplift vs single database | ≥ 1.5× | Validates P1 / H3 in live data |

---

## 7. Go / No-Go criteria

### Gate 0 — Launch? (January 2027)

**Go if all of the following hold** [E]:
1. **Coverage proven.** On the 200-account match-rate test, Clay returns a verified decision-maker contact for ≥ 1.5× as many accounts as the best single database, and for ≥ 50% of accounts overall.
2. **Pain confirmed.** In ≥ 15 validation interviews, a majority of sales leaders name new-customer acquisition or contractor coverage as a top-3 priority and have a CRM in use.
3. **Willingness to pay.** ≥ 3 distributors with 100+ employees sign paid Growth annual at list, with credit sizing implying an average ACV of ≥ $12K.
4. **Delivery capacity.** ≥ 2 implementation partners commit named consultants to certification by end of Q1 2027.
5. **Funding.** Executive sponsor approves the 2027 cap ($1.1M) and the transfer-cost rule for shared AEs.

**No-Go:** do not spend the $250K template/connector build. Fold any templates into the horizontal library and record the learnings. Cost to this point is staff time only.

### Mid-year checkpoint — June 2027 (early warning, not a gate)

Model H1 2027: Downside ~20 leads / 6 partner leads / ~3 opps; Base ~57 / 15 / ~13 [C].

If **opportunities < 6 and partner leads < 8**, freeze program spend for Q3. Then pull Gate 1 forward to the end of Q3 2027.

### Gate 1 — Stop / continue lean / scale (Q4 2027 evidence; decide by January 2028)

Model reference for 2027 [C]:

| Metric (cumulative 2027) | Downside | Base | Upside |
|---|---:|---:|---:|
| Leads | 52 | 141 | 340 |
| Partner-sourced leads | 20 | 42 | 96 |
| Producing partners (end Q4) | 4 | 5 | 9 |
| Opportunities | 7 | 31 | 111 |
| New customers | 0.5 | 6 | 25 |
| ARR, end Q4 2027 | $6K | $85K | $500K |
| Average first-year ACV | $11.0K | $13.8K | $19.4K |
| Cumulative program P&L | −$0.98M | −$1.03M | −$1.20M |

**Decision rules** [E]:

| Outcome | Rule | Action |
|---|---|---|
| **STOP** | Any **two** of: opportunities < 15; new customers < 3; partner-sourced leads < 25; average first-year ACV < $10K; win rate < 20% on ≥ 10 decided opportunities | End vertical spend. Keep customers in horizontal CS and templates in the library. Loss capped at ~$1.0M |
| **CONTINUE LEAN** (Base track) | Opportunities ≥ 20 **and** new customers ≥ 4 **and** ARR ≥ $50K **and** average first-year ACV ≥ $12K **and** ≥ 4 producing partners | Launch wave 2 in Q1 2028. No AE hires. Focus on ACV (ICP tilt, multi-workflow kit) |
| **SCALE** (Upside track) | Opportunities ≥ 70 **and** new customers ≥ 15 **and** ARR ≥ $300K **and** average first-year ACV ≥ $16K **and** win rate ≥ 32% **and** ≥ 7 producing partners | Launch wave 2 in Q1 2028. Hire 2 dedicated AEs in H1 2028 (capacity binds from Q3 2028 in the Upside run). Bring wave 3 forward |
| Between CONTINUE and STOP | Mixed signals | Continue lean for two more quarters with a hard re-review in Q2 2028; no new spend categories |

**Why "continue lean" is the right Base response:**
- Forward total value of continuing in Base from Q1 2028 is ≈ −$0.02M [C], which is roughly neutral.
- The pilot continues in Base only to find out whether Upside is reachable.

**Statistical caution:**
- In Base, Gate 1 rests on about 6 customers and ~31 opportunities.
- That is why the rules combine several indicators ("any two of") rather than one ratio.
- Win rate is not used on its own with fewer than 10 decided opportunities.

### Gate 2 — Scale to a standing vertical / hold lean / sunset (Q4 2028)

Model reference end-2028 [C]: Base ARR ~$334K, ~22 customers, 2028 CAC ~$51K, 2028-cohort payback ~68 months. Upside + scale-up: ARR ~$2.5M, ~113 customers, CAC ~$20K, payback ~17 months.

| Outcome | Rule [E] | Action |
|---|---|---|
| **SCALE** | 2028-cohort payback ≤ 24 months **and** ARR ≥ $1.5M **and** first-cohort NRR ≥ 100% **and** GRR ≥ 88% | Standing vertical team. Native ERP connector business case. Vertical data licensing (e.g., construction projects). Consider manufacturer-channel programs (H10) |
| **HOLD LEAN** | ARR ≥ $250K **and** 2028 CAC ≤ $55K **and** GRR ≥ 85% | No new hires. Keep partners and templates. Optimize retention and expansion through 2029. In Base, continuing from Q1 2029 is worth ≈ +$0.9M [C] because most costs are already spent |
| **SUNSET** | Below HOLD thresholds | Stop vertical programs. Transfer customers to horizontal CS. Keep templates public |

### Gate 3 — Graduate or fold (Q4 2029)

- **Graduate** to a standing vertical if exit ARR ≥ $3M and the 2029-cohort payback is ≤ 18 months. That is in line with mid-market benchmarks of 14–18 months (S47).
- Otherwise, fold the vertical into the horizontal motion as a template-and-partner "solution", with no dedicated cost line.

---

## 8. Risk register

L = likelihood, I = impact (H/M/L). Assessments are judgment [C].

| # | Risk | L | I | Early warning | Mitigation | Owner (function) |
|---|---|:-:|:-:|---|---|---|
| R1 | Demand thinner than Base (Downside) | M | H | Mid-year checkpoint: < 6 opps | Gates 0/1; spend released quarterly; ~$1.0M loss cap | Executive sponsor |
| R2 | ACV stuck near the Growth floor | H | H | Average first-year ACV < $12K; MM consumption < plan | ICP tilt to 100+ employees; credit sizing; multi-workflow kit; no discounts | PMM; Deal desk |
| R3 | Coverage advantage smaller than the proxy suggests (H3) | M | H | Match-rate uplift < 1.5× at Gate 0 | Gate 0 No-Go; add data providers before launch | Data partnerships |
| R4 | Too few productive partners | H | M | < 4 producing partners; time to first workflow > 30 days | Recruit from ERP/CRM implementers; certification; Clay-delivered onboarding for enterprise only | Partnerships |
| R5 | Shared AEs deprioritize small vertical deals | M | M | SQL-to-meeting lag > 5 days; lost deals for "no follow-up" | Transfer-cost credit to AE quota; named AE pool; capacity trigger for dedicated hires | Sales leadership |
| R6 | Channel conflict with existing agency partners | M | L | Account disputes | Deal registration; certification open to existing partners | Partnerships |
| R7 | Pricing / overage friction hurts retention | M | M | Overage disputes; downgrades at month 6 | Annual commitments sized at signature; 30/60/90 usage reviews | CS; Finance |
| R8 | Competitive response (Apollo waterfall/spreadsheet, ZoomInfo discounting) | H | M | Losses to Apollo on price; incumbent bundling | Lead with off-LinkedIn coverage and unlimited seats; complementary positioning | PMM |
| R9 | ERP data gap limits the enterprise use case | M | M | Enterprise deals stall on ERP integration | Warehouse/HTTP path first; native connector only at Gate 2 | Product |
| R10 | Construction-cycle slowdown hits electrical/PHVAC distributors | M | M | Distributor sales and hiring news; slower cycles | Bring wave 3 (MRO) forward; hold program spend | Vertical lead |
| R11 | Pilot distracts from the core horizontal motion | L | M | Shared AE time > plan | Fixed transfer cost and capacity ceiling; separate P&L tracking | Finance |
| R12 | Data compliance (contact data, permits, scraping terms) | L | H | Customer or provider complaints | Use marketplace providers with licensed data; legal review of templates | Legal |

---

## 9. What would change the plan

| If we learn… | Then… |
|---|---|
| Match-rate uplift is large (≥ 2×) and design-partner ACVs exceed $20K | Treat Upside as more likely; consider launching wave 2 in Q4 2027 (Upside timing) |
| Small distributors self-serve templates successfully (self-serve share ≥ 40%) | Add a product-led SMB track; low effect on value (swing ~$0.02M in 07) but cheap reach |
| Partners prove unproductive but direct Clay onboarding works | Redesign to a small Clay-delivered onboarding team for mid-market+; re-run 07 with higher onboarding cost and lower partner fees |
| Manufacturers' channel teams adopt via distributors | Open a manufacturer-channel motion at Gate 2 (H10) |
| Professional-services incremental headroom proves larger than estimated | Revisit Phase 3 trade-off (stress test T9 flips to professional services) |

---

## 10. Open items before execution

- Validate ICP roles, pains and willingness to pay (Gate 0 interviews). None were conducted in this study.
- Confirm Clay's actual win rates, transfer costs, gross margin and partner economics. These are external benchmarks or assumptions in 07.
- Verify competitor capabilities rated [K] in 04 against current vendor documentation.
- Cross-check Census firm counts against SUSB 2022 (open QA item from Phase 1).
