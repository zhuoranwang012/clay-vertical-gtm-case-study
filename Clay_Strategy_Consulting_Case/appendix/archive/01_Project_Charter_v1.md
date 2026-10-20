# 01 — Project Charter

**U.S. Vertical Expansion & Go-to-Market Strategy for Clay**
Independent strategy consulting case study · Author: Zhuoran (Zora) Wang · Research reference date: October 18, 2026

> **Status disclaimer.** This is an independent portfolio case study. It is not commissioned by, affiliated with, or endorsed by Clay Labs Inc. It uses only public information. No Clay internal data, interviews, or financial records were accessed. Any revenue, adoption, or cost figure produced in later phases is a labeled scenario estimate, not a Clay forecast.

**Evidence labels used throughout** (see `Research_Sources.csv`):
**[A]** Verified public fact · **[B]** External benchmark / third-party estimate · **[C]** Analyst estimate (this project) · **[D]** Scenario assumption · **[E]** Strategic recommendation. Source IDs (e.g., S01) map to the register.

---

## 1. Executive problem statement

Clay has grown from a tool adopted by AI-native and software companies into a broad GTM data and workflow platform (company-reported >$100M ARR in December 2025 and 14,000 customers in January 2026 [A, S02]; 17,000+ customers reported at its September 2026 Series D [A*, S03]). Its public customer evidence remains heavily concentrated in software and AI, yet there are early, scattered signs of use in industrial settings — an electrical distributor case study (S07) and an enterprise manufacturer pilot run through a consulting partner (S08).

The question is **not** whether Clay *can* serve industrial companies — it already does in some cases. The question is whether **deliberately investing incremental GTM resources** in U.S. industrial distribution and manufacturing would generate better risk-adjusted growth over three years than (a) investing in an alternative vertical such as professional services, or (b) continuing a horizontal, persona-led strategy with no vertical bet.

**Governing question:**
> *Over 2027–2029, should Clay commit incremental go-to-market investment to U.S. industrial distribution and/or manufacturing — relative to professional services or a no-vertical baseline — and if so, through which go-to-market motion and at what scale?*

---

## 2. Client context (public information only)

### 2.1 What Clay sells
| Dimension | Finding | Evidence |
|---|---|---|
| Core positioning | GTM data + orchestration platform: multi-provider "waterfall" enrichment, AI research agents (Claygent, account agents, a new global agent), signals, audiences, workflows, sequencer and ad-audience activation | [A] S01 |
| Data model | Marketplace of 150–200+ third-party data providers accessed through Clay; customers can also bring their own API keys | [A] S01 |
| Monetization | Dual meter since a March 2026 restructuring: **Actions** (platform usage) and **Data Credits** (data and AI purchased from marketplace vendors) | [A] S01; restructuring date [B] S04 |
| List pricing (monthly billing) | Free; **Launch ~$185/mo**; **Growth ~$495/mo**; **Enterprise custom** (annual commitment, 100K+ data credits, 200K+ actions/mo). Annual billing ~10% lower ($167 / $446 per month displayed). | [A] S01 |
| Key tier gate | **CRM integrations, HTTP API, webhooks and data-warehouse sync begin at Growth**; SSO, RBAC and a dedicated "growth strategist" (CSM) are Enterprise-only | [A] S01 |
| Seats | Unlimited seats on all plans — pricing scales with usage, not headcount | [A] S01 |

*Implication for this project:* any industrial customer that needs CRM-synced workflows (most mid-market sales teams) enters at Growth at minimum, i.e., roughly **$5.3K–$5.9K/year at list** before usage expansion [C, arithmetic from S01]. This becomes a floor for ACV assumptions in Phase 4.

### 2.2 Scale and trajectory
| Metric | Value | Date | Class / source |
|---|---|---|---|
| ARR | >$100M | Dec 2025 | [A] company-reported, S02 |
| ARR (estimate) | ~$150M | May 2026 | [B] third-party estimate, S04 — *not company-confirmed* |
| Customers | 14,000 | Jan 2026 | [A] S02 |
| Customers | 16,000+ | undated website copy | [A] S06 |
| Customers | 17,000+ | Sep 2026 | [A*] S03 (company statement as reported by trade press) |
| Enterprise NRR | >200% | Jan 2026 | [A] S02 |
| Funding | $115M Series D at $7.1B valuation | Sep 2026 | [A*] S03 (secondary reporting of NYT account) |
| Sales model | "Roughly 50/50 self-serve and sales-led" | undated blog | [A] S05 |

\*Reported by secondary outlets; primary press release not yet located — to be confirmed in Phase 1 QA.

**Derived indicator [C]:** $100M ARR ÷ 14,000 customers ≈ **$7.1K blended ARR per customer**. Combined with >200% enterprise NRR, this implies a long tail of Launch/Growth accounts plus a smaller enterprise base that expands quickly. Dates are approximately but not exactly aligned (Dec 2025 vs. Jan 2026); treat as order-of-magnitude only.

### 2.3 Publicly visible market presence
- **Customer showcase is software/AI-heavy.** Featured case studies include OpenAI, Anthropic, Rippling, Figma, Intercom, Vanta, Verkada, Mistral AI, and similar [A, S06].
- **Industry taxonomy omits industrials.** Clay's own case-study filters list Software & AI, Healthcare & Biotech, Real Estate, Finance, **Professional Services**, and Media & Events — **no manufacturing, industrial, or distribution category** [A, S06].
- **Latest data-marketplace expansion targeted other verticals.** The September 29, 2026 integrations release added providers for **public sector, commercial real estate, and private equity** — not industrial or manufacturing [A, S09].
- **Enterprise partner channel is emerging.** Clay announced a strategic partnership with consultancy Marketbridge on September 29, 2026, aimed at large enterprises [A, S08].

---

## 3. What is known and unknown about industrial penetration

### Known (public evidence)
| Evidence | What it shows | Class |
|---|---|---|
| **Regency Supply** case study — commercial lighting & electrical distributor, Chatsworth, CA (S07) | Clay used to (1) find local electrical contractors via Google Maps data where ZoomInfo and Apollo fell short, (2) automate qualification of wholesale e-commerce registrations, (3) enrich and sync inside HubSpot, (4) monitor job changes across 5,000+ procurement/construction/facilities contacts | [A] |
| **LG Electronics USA** (Commercial Display & IT Solutions) piloting a Clay-powered workflow via Marketbridge (S08) | A large manufacturer is testing Clay — but through a partner, as a pilot, in a tech-adjacent division | [A] |
| **Waste Management** cited by Clay as a user (S05) | Clay itself argues it is not limited to tech/PLG companies; WM is industrial services, not manufacturing or distribution | [A] |
| Third-party agencies market "Clay for manufacturing" programs (S12) | An ecosystem is attempting to sell Clay-based services into manufacturing; signals perceived demand, not proven demand | [B, weak] |

### Unknown (and how the project will handle each)
| Unknown | Why it matters | Handling |
|---|---|---|
| Share of Clay customers / ARR in NAICS 31-33 or 42 | Baseline penetration; is this "penetration" or "near-greenfield"? | Not public. State the gap; do not estimate as fact. Use proxy indicators (case studies, partner listings, job postings) [C]. |
| Win rate, retention, ACV by vertical | Core inputs to the business case | Use external SaaS benchmarks [B] + scenario ranges [D]; sensitivity-test heavily |
| Quality of Clay's data coverage for industrial buyer personas | Regency case suggests key personas are thin on LinkedIn-derived databases — a pain point Clay can solve, but only with extra workflow-building effort | Test in Phase 2 via provider documentation and, if feasible, a small public-data coverage test (optional AI module) |
| Industrial firms' internal capacity to build Clay workflows (RevOps / GTM-engineering talent) | Clay's value is unlocked by builders; low builder density means low self-serve adoption | Proxy via public occupational/job-posting data in Phase 1–2 |
| Whether Clay already has a vertical sales team for industrials | Determines whether the decision is "start" or "scale" | Not public; framed as option-neutral |

---

## 4. The strategic decision

**Decision owner (functional, not organizational):** executive team responsible for GTM resource allocation (revenue, marketing, partnerships, product). This case does not presume Clay's actual reporting structure.

**Options to be evaluated:**

| # | Option | Description |
|---|---|---|
| 0 | **No vertical bet (baseline)** | Continue horizontal, persona-led growth; industrial customers arrive organically |
| 1 | **Industrial distribution focus** | Incremental investment targeting U.S. wholesale distributors of durable industrial goods |
| 2 | **B2B manufacturing focus** | Incremental investment targeting U.S. industrial/discrete manufacturers selling B2B |
| 3 | **Professional services focus** | Incremental investment in a segment Clay already markets to (has its own showcase filter) |
| 4 | **Staged pilot** | Limited, partner-led pilot in the most promising industrial sub-segment with explicit Go/No-Go gates before scaling |

Each option will be evaluated by **motion** (product-led, sales-led, partner-led, hybrid) and **scale** (pilot vs. full commitment). Option 0 is the comparison point for incremental value; a vertical investment must beat what the same resources would earn in Clay's core horizontal motion.

**Decision criteria (defined up front, weighted in Phase 3):** market attractiveness, customer need intensity, product-market fit, revenue potential, competitive position, ease of execution, and strategic risk.

---

## 5. Project objectives

1. Size the U.S. opportunity in each candidate segment bottom-up from Census firm counts, with explicit TAM/SAM/SOM definitions.
2. Assess whether Clay holds a defensible advantage over Apollo, ZoomInfo, HubSpot, specialist industrial data providers, and internal build in each segment.
3. Rank segments with a transparent, sensitivity-tested weighted framework.
4. Build a three-year scenario business case (funnel → customers → revenue → CAC, payback, gross profit) for the recommended option.
5. Translate findings into a GTM strategy, pilot design, KPIs, and Go/No-Go criteria.
6. Produce an executive deck that a consulting interviewer can interrogate slide by slide.

---

## 6. Scope and limitations

**In scope:** U.S. only; for-profit employer firms; segments A–C (refinable); three-year horizon (2027–2029); Clay's current product as publicly documented.

**Out of scope:** international markets; product roadmap design beyond packaging/positioning implications; M&A; valuation of Clay.

**Limitations (stated now, revisited in every phase):**
1. **No internal data.** Penetration, win rates, churn, CAC and gross margin by segment are unavailable; all are benchmarks [B] or assumptions [D].
2. **No primary research.** No customer interviews or surveys were conducted, and none will be fabricated. An interview guide for validating hypotheses will be provided in `appendix/` as a recommended next step.
3. **Selection bias in case studies.** Vendor case studies showcase successes; absence of an industrial case study is weak evidence of low penetration, not proof.
4. **Census data lag.** The most recent SUSB annual tables with receipts are the **2022 reference year, released 2025** [A, S10]. Firm counts are assumed to be broadly stable to 2026; this assumption will be noted where used.
5. **Firm vs. establishment.** SUSB reports firms and establishments separately; a multi-industry firm is counted in each industry it operates in, so **firm counts are not additive across NAICS codes**. Phase 1 will use enterprise-size bands and avoid summing overlapping firm counts without adjustment.
6. **Third-party financial estimates** (e.g., Sacra ARR estimates, S04) are labeled [B] and never presented as company fact.
7. **Pricing volatility.** Clay restructured pricing in March 2026 [B, S04]; community commentary suggests the change raised costs for small, high-usage users (S11, opinion). List prices may change again within the horizon.

---

## 7. Issue tree

**Governing question:** Should Clay commit incremental U.S. GTM investment to industrial distribution and/or manufacturing, relative to professional services or no vertical bet — and through which motion?

```
GOVERNING QUESTION
│
├── 1. MARKET ATTRACTIVENESS — Is the segment large and reachable enough?
│     1.1 How many U.S. firms (not establishments) fit each segment definition?   → SUSB 2022 by NAICS × enterprise size
│     1.2 How many have a B2B sales motion that Clay can improve?                 → size-band filters; receipts per firm
│     1.3 Is the segment growing, stable, or consolidating?                        → SUSB/CBP time series; industry reports
│
├── 2. CUSTOMER NEEDS — How intense and specific is the pain?
│     2.1 What GTM tasks are manual today (prospecting, territory, qualification)? → case studies, job postings, industry research
│     2.2 Are buyer personas poorly covered by standard B2B databases?             → Regency evidence; provider coverage claims
│     2.3 Is the pain budgeted (is there a buyer with authority + spend)?          → org structures, sales-ops role prevalence
│
├── 3. COMPETITIVE POSITION — Can Clay win, and keep winning?
│     3.1 Who serves this need today (Apollo, ZoomInfo, HubSpot, D&B, Thomasnet, ERP/CRM vendors, agencies, internal build)?
│     3.2 Where is Clay differentiated (multi-source, AI research, custom signals)?
│     3.3 Where are alternatives stronger (embedded data, simplicity, industry-native data, existing contracts)?
│
├── 4. PRODUCT-MARKET FIT — Does the product work for these buyers as-is?
│     4.1 Which Clay capabilities map to segment pain points?                      → pain-to-capability matrix
│     4.2 What capability gaps exist (industrial data providers, ERP integration, templates)?
│     4.3 Can customers build workflows themselves, or do they need services?      → builder-talent proxy
│
├── 5. REVENUE POTENTIAL — How much revenue is realistically capturable?
│     5.1 TAM / SAM / SOM (bottom-up)
│     5.2 Expected ACV by size band and tier                                       → pricing page + usage scenarios
│     5.3 Expansion potential (NRR) vs. core segments
│
├── 6. UNIT ECONOMICS — Is the growth profitable enough?
│     6.1 CAC by motion (PLG, sales-led, partner-led)
│     6.2 Gross margin (data-credit pass-through vs. Actions)
│     6.3 CAC payback and contribution over three years; break-even thresholds
│
├── 7. EXECUTION FEASIBILITY — Can Clay deliver it with reasonable effort?
│     7.1 New assets required: case studies, templates, data providers, partners, vertical sellers
│     7.2 Sales-cycle length and complexity vs. Clay's current motion
│     7.3 Partner ecosystem availability (agencies, distributor-focused consultancies)
│
└── 8. STRATEGIC RISK — What could go wrong, and what is the downside?
      8.1 Focus/dilution risk vs. core AI-native and enterprise segments
      8.2 Adoption risk (low technical capacity → churn)
      8.3 Competitive response (incumbents bundling industrial data)
      8.4 Macro/industrial cycle exposure
```

---

## 8. Initial hypotheses (falsifiable)

These are working hypotheses to be **tested, not defended**. Each lists the evidence that would disprove it.

| ID | Branch | Hypothesis | Test (phase) | Would be rejected if… |
|---|---|---|---|---|
| **H1** | Market | "B2B manufacturing suppliers" as written is too broad to be a coherent GTM segment; a narrower set of industrial/discrete manufacturing subsectors is more appropriate. | NAICS review; SUSB composition (P1) | Subsector-level firm profiles (size, receipts, B2B share) are homogeneous enough to target as one segment. |
| **H2** | Market | Industrial distribution has a larger pool of mid-market firms (≈20–499 employees) than the narrowed manufacturing segment, and a mid-market pool at least half the size of the target professional-services sub-industries. | SUSB 2022 firm counts by enterprise size (P1) | Distribution mid-market firm count is < 50% of target professional services or smaller than narrowed manufacturing. |
| **H3** | Customer need | Industrial buyer personas (contractors, plant/maintenance managers, procurement) are poorly covered by LinkedIn-derived B2B databases, creating a data-coverage pain that Clay's multi-source + AI research model addresses better than single-database tools. | Regency evidence (S07); competitor data-coverage documentation; optional coverage test (P2) | Major providers demonstrably offer strong industrial persona coverage, or specialist industrial databases already solve it cheaply. |
| **H4** | PMF / execution | The binding constraint in industrial segments is **builder capacity**, not need: few industrial firms employ RevOps/GTM-engineering talent, so self-serve adoption will be low and success will require partner- or service-led delivery. | Occupational / job-posting proxies by industry; partner directory (P1–P2) | Sales-ops/RevOps role prevalence in industrial firms is comparable to professional services. |
| **H5** | Competitive | Clay's advantage in industrial segments is **complementary** (orchestration on top of existing data/CRM) rather than **substitutive**, so the realistic entry is "layer on HubSpot/Salesforce/ZoomInfo," not "replace them." | Competitive feature/pricing comparison (P2) | Evidence that industrial buyers consolidate onto one all-in-one vendor and reject layered stacks. |
| **H6** | Revenue | Typical industrial mid-market ACV will sit in the Growth-tier range (~$5K–$25K/year), below enterprise but above Clay's ~$7K blended average only for the upper size bands. | Pricing (S01) + usage scenarios (P1, P4) | Public evidence of materially higher industrial deal sizes (e.g., enterprise-scale industrial contracts). |
| **H7** | Unit economics | A **fully sales-led** industrial vertical motion will not reach CAC payback < 24 months at Growth-tier ACVs in the base case; a **partner-led** motion can. | Phase 4 model | Base-case sales-led payback < 24 months under benchmark CAC. |
| **H8** | Strategic risk | Even in the upside case, an industrial vertical will contribute a small share of Clay's total ARR within three years, so it should be treated as an **option-value pilot**, not a core strategic pivot. | Phase 4 scenarios vs. public ARR baseline | Upside scenario reaches a materially meaningful share (threshold to be set in P4, e.g., >10% of projected ARR). |
| **H9** | Counter-hypothesis | **Professional services offers better risk-adjusted returns than either industrial segment** because buyers are more digitally mature, Clay already markets to them (S06), and execution assets partially exist. | Phase 3 scoring + Phase 4 comparison | Professional services scores lower on need intensity or competitive position, or proves saturated/low-ACV. |

**Day-1 lead hypothesis (stated to be challenged):** *Industrial distribution is the more attractive industrial sub-segment, but it justifies a partner-led, staged pilot rather than a full vertical build-out; professional services may deliver faster near-term returns.* Phase 3 must explicitly confirm or overturn this.

**Disconfirming evidence already observed (Phase 0):**
- Clay's September 2026 data-marketplace expansion prioritized public sector, commercial real estate and private equity, not industrials (S09) — suggesting either lower internal priority for industrials or an unaddressed gap. Both readings will be considered.
- The Regency case cuts both ways: it shows Clay winning where ZoomInfo/Apollo failed (supports H3), but the winning workflow required custom building by a technically capable marketing director (supports H4's adoption-risk concern).

---

## 9. Research methodology

**Approach:** hypothesis-driven, issue-tree-structured; each analysis answers a named sub-question.

| Step | Method | Primary data | Output |
|---|---|---|---|
| Segment definition | NAICS mapping (2017 NAICS, consistent with SUSB 2022) | Census NAICS manuals, SUSB | Segment definition table |
| Market sizing | Bottom-up: firms × eligibility filters × adoption × ACV; TAM ≠ SAM ≠ SOM | SUSB 2022 (firms by enterprise size), CBP cross-check | `02_Market_Sizing.xlsx` |
| Competitive analysis | Feature / pricing / data-coverage comparison from public vendor pages and filings (e.g., ZoomInfo 10-K, HubSpot 10-K) | Vendor sites, SEC EDGAR | `04_Competitive_Analysis.xlsx` |
| Customer needs | Pain-point synthesis from public case studies, job postings, industry research; pain-to-capability mapping | Public sources only | `05_Customer_Needs_Assessment.md` |
| Prioritization | Weighted scoring (suggested weights treated as assumptions) + weight sensitivity + rank-reversal tests | Phases 1–2 | `06_Segment_Prioritization.xlsx` |
| Business case | Funnel-driven 3-year model; Base/Upside/Downside; tornado sensitivity; break-even thresholds | Benchmarks [B], assumptions [D] | `07_Financial_Model.xlsx` |
| Strategy | ICP, motion choice, pricing/packaging, pilot design, KPIs, Go/No-Go | All prior | `08_GTM_Strategy.md`, `09_Implementation_Roadmap.md` |

**Proxy rule** (when data is unavailable): (1) state the gap, (2) choose the most defensible proxy, (3) label it [C] or [D], (4) run sensitivity, (5) report as a range, not a point estimate.

---

## 10. Research workplan

| Phase | Key question answered | Deliverables | Quality gate before next phase |
|---|---|---|---|
| **0 — Scoping** ✅ | What exactly are we deciding? | `01_Project_Charter.md`, seed `Research_Sources.csv`, `README.md` | Facts sourced; evidence classes applied; framing is "prioritization," not "entry" |
| **1 — Market** | How big and reachable is each segment? | `02_Market_Sizing.xlsx`, `03_Market_Research.md` | Firm vs. establishment check; no cross-NAICS double count; TAM/SAM/SOM definitions consistent |
| **2 — Competition & needs** | Can Clay win, and where does it fit? | `04_Competitive_Analysis.xlsx`, `05_Customer_Needs_Assessment.md` | Every competitor claim sourced; overlap verified, not assumed |
| **3 — Prioritization** | Which segment deserves investment? | `06_Segment_Prioritization.xlsx` | Every score has evidence; ranking survives weight sensitivity or reversal conditions are stated |
| **4 — Business case** | What does the investment return? | `07_Financial_Model.xlsx` | Formulas linked, no hardcoded outputs, scenarios recalculate, payback checked by hand |
| **5 — Strategy** | How should Clay act? | `08_GTM_Strategy.md`, `09_Implementation_Roadmap.md` | Every recommendation traces to an analysis; Go/No-Go measurable |
| *Optional — AI module* | Can AI-assisted account research support execution? | `appendix/` prototype + measured results | Measured, not invented, performance |
| **6 — Deck** | Can an executive decide from this? | `Executive_Strategy_Deck.pptx` | Charts match model; scenario outputs labeled; independent-case disclaimer present |

---

## 11. Evidence requirements

| Branch | Evidence needed | Preferred source | Fallback proxy | Expected class |
|---|---|---|---|---|
| Market | Firm counts by NAICS × enterprise employment size; receipts | Census SUSB 2022 | CBP 2023 establishment counts (cross-check only) | A |
| Market | Segment growth / consolidation | SUSB time series; Census Economic Census | Trade association reports | A / B |
| Need | Prevalence of manual prospecting, territory, qualification pain | Public case studies; industry surveys | Job-posting language analysis | B / C |
| Need | Coverage of industrial personas in B2B data | Vendor documentation; S07 | Small public-data coverage test | B / C |
| Competition | Pricing, features, segments, data coverage | Vendor pricing pages; SEC 10-Ks (ZoomInfo, HubSpot) | Reputable analyst / review sites | A / B |
| PMF | Builder talent density by industry | BLS OES occupational data | Job-board counts | A / C |
| Revenue | ACV ranges | Clay pricing page (S01) | Usage-based scenario | A + D |
| Unit economics | CAC, payback, gross margin, NRR benchmarks | Public SaaS benchmark reports; public-company filings | Ranges with sensitivity | B / D |
| Execution | Partner ecosystem, sales-cycle length | Clay partner pages; S08, S12 | Industry sales-cycle benchmarks | A / B |
| Risk | Pricing change impact; competitive moves | Company announcements; trade press | Commentary, labeled as opinion | A / B |

---

## 12. Phase 0 quality review

| Check | Result |
|---|---|
| Framing is prioritization/penetration, not first entry | ✅ Known industrial evidence documented in §3 |
| Independent-case status stated | ✅ Header disclaimer; no implied client relationship |
| Facts vs. estimates separated | ✅ Evidence classes applied; third-party ARR labeled [B] |
| Source discrepancies flagged | ✅ Customer counts (14K / 16K+ / 17K+) shown with dates; Series D reported via secondary sources pending primary confirmation |
| Pricing interpretation checked | ✅ Monthly vs. annual display reconciled ($185 × 0.9 ≈ $167; $495 × 0.9 ≈ $446) |
| Disconfirming evidence sought | ✅ §8 lists evidence that cuts against the lead hypothesis |
| Excluded as unreliable | One widely circulated "business model" article (fourweekmba.com) reporting unsourced ARR figures inconsistent with company disclosures was **not used**. |

**Open items carried into Phase 1:** confirm Series D via a primary source; confirm whether a 2023 SUSB vintage has been released; locate publication dates for undated Clay pages (S05, S07).

---

## 13. Interview framing (how to explain Phase 0)

*"I reframed the prompt from 'should Clay enter manufacturing' to 'where should Clay put its next marginal GTM dollar,' because Clay already has industrial customers. I defined a no-vertical baseline so any vertical bet had to beat Clay's core motion, built an eight-branch issue tree, and wrote nine falsifiable hypotheses — including a counter-hypothesis that professional services wins. Two early facts shaped the work: Clay's own industry taxonomy and latest data-provider additions skip industrials entirely, and its best industrial case study shows both strong need and high build complexity. That tension — high need, low builder capacity — became the central question for the rest of the analysis."*
