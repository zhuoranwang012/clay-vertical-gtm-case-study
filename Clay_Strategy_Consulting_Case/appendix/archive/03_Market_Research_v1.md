# 03 — Market Research & Sizing

**U.S. Vertical Expansion & Go-to-Market Strategy for Clay** · Phase 1
Independent portfolio case study by Zhuoran (Zora) Wang · Not commissioned by or affiliated with Clay Labs Inc. · Research reference date: October 18, 2026

> All figures trace to `02_Market_Sizing.xlsx` (every output is a live formula). Evidence classes: **[A]** verified public fact · **[B]** external benchmark · **[C]** analyst estimate derived from data · **[D]** scenario assumption. TAM/SAM/SOM values are scenario outputs, **not Clay forecasts**.

---

## 1. Answer first

**On market size alone, industrial distribution is the weakest of the three segments — but it is by far the most sales-intensive, and that tension is the central question for Phases 2–4.**

| | A. Industrial Distribution | B. Industrial Manufacturing | C. B2B Professional Services |
|---|---:|---:|---:|
| Unique U.S. firms, 20+ employees (2022) [C] | **11,494** | 23,785 | **33,035** |
| TAM, Base [C/D] | $131M | $281M | $385M |
| SAM, Base [C/D] | $68M | $76M | $148M |
| SOM — year-3 ARR potential, Base [D] | **$1.9M** | $2.1M | **$4.2M** |
| SOM outer bounds (compounded Low – High) [D] | $0.2M – $8.8M | $0.2M – $13.0M | $0.5M – $23.4M |
| Sales occupations, % of employment (all industries 8.6%) [C] | **22.0%** | 2.7% | 6.2% |
| Computer & math occupations, % of employment (all industries 3.4%) [C] | 2.3% | 1.5% | **24.2%** |

Five findings:

1. **Hypothesis H2 is rejected.** Distribution has the fewest mid-market firms: its 20–499-employee pool is 35% of professional services' (threshold was 50%) and 49% of manufacturing's (threshold was 100%).
2. **Distribution compensates with need intensity.** Sales occupations are ~22% of distribution employment, about 2.5× the economy-wide share, and an in-scope distributor employs ~28 sales staff on average. Because so many distributors run active sales motions, distribution's SAM ($68M) roughly equals manufacturing's ($76M) despite half the firm count.
3. **Manufacturing looks large but sells indirectly.** Sales occupations are only 2.7% of employment in the core manufacturing subsectors, consistent with selling through distributors and independent reps. This suggests reaching manufacturers *through* the distribution channel rather than as a stand-alone vertical (new hypothesis H10).
4. **Builder capacity mirrors need.** Professional services has near-software levels of technical talent; distribution and manufacturing sit below the economy-wide average. This is preliminary support for H4: in distribution, the binding constraint is the ability to build workflows, not the need for them.
5. **No vertical is a pivot.** At uniform fair-share penetration, base-case SOM equals ~1–2% of Clay's estimated ~$200M ARR (Sep 2026; [B] Sacra estimate, not company-confirmed). Any vertical program is an incremental growth option (context for H8).

---

## 2. Segment definitions (refined)

The brief's three segments were refined using NAICS 2022. Each "core" set is used in the model; "broad" sets appear only in the definition sensitivity.

| Segment | Core NAICS | Excluded / broad-only | Rationale |
|---|---|---|---|
| **A. Industrial Distribution** | 4235 metals; 4236 electrical/electronic; 4237 hardware, plumbing & HVAC; 4238 machinery, equipment & supplies (merchant wholesalers) | 4233 construction materials; 4234 professional & commercial equipment (includes computers and medical) | Outside-sales-led wholesale of industrial durables. 4236 matches the published Regency Supply case (S07). |
| **B. Industrial Manufacturing** *(narrowed from "B2B Manufacturing Suppliers")* | 326 plastics & rubber; 332 fabricated metal; 333 machinery; 335 electrical equipment | Rest of 31-33 (food, beverage, apparel, printing, furniture, wood, chemicals, transportation equipment, electronics) | **H1 supported.** The core subsectors are only ~39% of manufacturers with 20+ employees. The remainder mostly sells through retail/CPG or local-market channels rather than engineered B2B selling. |
| **C. B2B Professional Services** | 5413 architecture & engineering; 5415 IT services; 5416 consulting; 5418 advertising & PR | 5411 legal; 5412 accounting & tax prep | Firms that sell to businesses with some outbound or business-development function. Legal and tax prep are referral-driven or consumer-facing. |

**Size bands** use the employment of the whole firm: SMB 20–99, Mid-Market 100–499, Enterprise 500+.
- Firms under 20 employees are excluded from TAM. They rarely staff a dedicated sales or RevOps function, and Clay's CRM sync starts at the Growth tier (S01).
- Benchmark only, not scored: software publishers (5132), as a proxy for Clay's current core.
- Adjacency noted for future screening: staffing (5613).

---

## 3. Data and method

### 3.1 Sources
- **Firm counts:** 2022 Economic Census, *Establishment and Firm Size Statistics — Employment Size of Firms* (EC2200SIZEEMPFIRM), retrieved via data.census.gov (S14) [A].
  - Reference year 2022. It is the most recent official firm-by-size data with receipts for these sectors.
  - The brief named the Statistics of U.S. Businesses (SUSB) as a starting source. Its 2022 spreadsheets were not machine-accessible from this research environment, and the Census API now requires a key.
  - Both programs report 2022 firm counts by firm employment size. Re-running the firm counts against SUSB 2022 offline is logged as a QA follow-up.
- **Occupational structure:** BLS OEWS May 2025 national industry-specific estimates (S15) [A].
- **ACV benchmarks:**
  - Clay list pricing (S01) [A].
  - Vendr marketplace data on Clay (S16) [B]: median $59,996/yr across 96 purchases, range $16.5K–$170K, page updated February 2026.
- **CRM readiness:** Distribution Strategy Group, *State of CRM in Distribution* (2022; n = 96 respondents, 66 distributors, $10M–$500M revenue) (S17) [B].

### 3.2 Firms, not establishments
Establishments are never used as the counting unit. The gap is material in distribution:
- The four core distribution industries have 61,863 firms but 102,922 establishments (1.66 per firm).
- Hardware/plumbing/HVAC (4237) averages 2.21 branches per firm.
- Counting establishments would have overstated distribution's buyer universe by about two-thirds.
- Branch density is itself a finding: territory planning and branch-level account mapping are core workflows in distribution.

### 3.3 Two Census data-quality adjustments [C]
1. **Suppressed cells.** Census suppresses firm counts for NAICS 4235 at 50–99, 100–249 and 250–499 employees.
   - The suppressed total (475 firms) is recovered as a residual from the published full-year total.
   - It is allocated using the parent sector's (423) size mix.
   - Imputed cells are flagged yellow in the workbook.
2. **Double counting across sub-industries.** Census counts a firm in every industry where it has an establishment, so firm counts are not additive across NAICS codes.
   - Overlap is measured directly: each sector's firm count is compared with the sum of its sub-industries, by size band.

| Measured overlap (Σ sub-industries ÷ sector − 1) | 20–49 | 50–99 | 100–249 | 250–499 | 500+ |
|---|---:|---:|---:|---:|---:|
| Durable-goods wholesale (423) | 2.8% | n/m* | n/m* | n/m* | ~0% |
| Manufacturing (31-33) | 3.1% | 6.8% | 11.1% | 10.2% | 9.6% |
| Professional services (54) | 1.2% | 2.2% | 3.4% | 5.2% | 2.9% |

\*n/m: some sub-industry cells are suppressed, so overlap cannot be measured; no reduction applied.

How the overlap is used:
- Each segment's subtotal is scaled down by the sector's band-level overlap. Without this, mid-size manufacturing would be overstated by up to ~11%.
- Where sub-industry cells are suppressed, the factor removes less overlap than truly exists, so it is conservative.
- Cross-segment overlap (e.g., a manufacturer with a wholesale arm) is not removed, because the segments are compared, never added.

### 3.4 TAM / SAM / SOM definitions (applied identically to all segments)

| Layer | Definition | Formula (per size band) |
|---|---|---|
| **TAM** | Revenue if every in-scope 20+ firm bought at the size-band ACV | Unique firms × ACV |
| **SAM** | Firms with an active B2B outbound/territory sales motion **and** a CRM Clay can integrate with today | TAM firms × filter 1 × filter 2 × ACV |
| **SOM** | Year-3 ARR potential at uniform fair-share penetration | SAM firms × 3-yr penetration × ACV |

**Design choice:** Phase 1 uses the **same ACV and the same penetration for every segment.** This isolates market size and reachability. Segment-specific right-to-win (competition, data coverage, ease of adoption) is assessed in Phases 2–3, not hidden inside the sizing.

### 3.5 Key assumptions (full list on the `Assumptions` tab)

| Input | Low | Base | High | Basis |
|---|---:|---:|---:|---|
| ACV — SMB (20–99) | $5,352 | $7,500 | $12,000 | Low = Growth tier at list, annual billing [A]; CRM sync requires Growth+ |
| ACV — Mid-Market | $12,000 | $20,000 | $35,000 | Vendr mid-market range ~$10K–$30K [B] |
| ACV — Enterprise | $30,000 | $60,000 | $120,000 | Vendr median $59,996 (n=96) [B] |
| Filter 1: outbound motion — Distribution / Mfg / Prof. Services | 65 / 30 / 40% | 80 / 45 / 55% | 90 / 60 / 70% | Informed by OEWS sales intensity [A→D] |
| Filter 2: CRM-ready — Distribution | 52% | 65% | 76% | DSG 2022 distributor survey [B] |
| Filter 2: CRM-ready — Mfg / Prof. Services | 45 / 55% | 60 / 70% | 75 / 85% | **Data gap** — no reliable public source [D] |
| 3-yr penetration — SMB / MM / Enterprise | 0.5 / 1 / 2% | 1.5 / 3 / 6% | 3 / 6 / 12% | Fair-share assumption [D] |

---

## 4. Findings in detail

### 4.1 Market size: professional services is ~3× distribution

| Unique firms (2022) | SMB 20–99 | Mid-Market 100–499 | Enterprise 500+ | **Total 20+** |
|---|---:|---:|---:|---:|
| A. Industrial Distribution | 9,106 | 2,011 | 376 | **11,494** |
| B. Industrial Manufacturing | 18,442 | 4,449 | 893 | **23,785** |
| C. B2B Professional Services | 26,021 | 5,775 | 1,239 | **33,035** |
| *Ref: software publishers* | | | | *3,274* |

- Within distribution, machinery & equipment (4238) is 52% of in-scope firms before overlap removal (6,009 of 11,676), followed by electrical (2,578), hardware/plumbing/HVAC (1,821) and metals (1,268).
- Within manufacturing, fabricated metal (332) is over half the segment, and it is dominated by small job shops: average receipts are ~$9M per firm across all sizes.

### 4.2 Need intensity: distribution sells; manufacturing mostly doesn't (directly)

| OEWS May 2025 [C] | Distribution | Manufacturing | Prof. Services | Software (ref) | All industries |
|---|---:|---:|---:|---:|---:|
| Sales occupations, % of employment | **22.0%** | 2.7% | 6.2% | 10.8% | 8.6% |
| Wholesale & mfg sales reps per 1,000 employees | **156** | 23 | 6 | 28 | 10 |
| Sales managers per 1,000 employees | 18.9 | 6.0 | 10.5 | 24.8 | 4.1 |
| Implied sales staff per 20+ firm | **27.7** | 3.6 | 9.3 | 32.0 | — |

Clay's usage-based pricing scales with records enriched, signals tracked and sequences run, all of which grow with seller headcount.
- A typical in-scope distributor has ~3× the sellers of a professional-services firm.
- It also has ~4× the revenue: in the 20–99-employee band, average receipts are $40.6M for distributors vs $10.1M for professional services.
- **Phase 1 deliberately did not convert this into a higher distribution ACV**, because there is no public contract evidence. It is the most important open lever (see §5).

### 4.3 Builder capacity: the mirror image

| OEWS May 2025 [C] | Distribution | Manufacturing | Prof. Services | Software (ref) | All industries |
|---|---:|---:|---:|---:|---:|
| Computer & math occupations, % of employment | 2.3% | 1.5% | **24.2%** | 46.7% | 3.4% |
| Market research analysts & marketing specialists per 1,000 | 8.4 | 3.6 | **31.7** | 35.6 | 5.8 |
| Marketing managers per 1,000 | 4.4 | 1.8 | **11.6** | 19.7 | 2.5 |

Professional services looks like a smaller version of Clay's existing software customer base. Distribution has the opposite profile: dense sales teams with thin technical and marketing-ops benches. This matches the Regency Supply case, where a technically capable marketing director built every workflow himself (S07).
- **Implication for the GTM motion:** distribution likely needs packaged templates and/or implementation partners.
- **Implication for professional services:** it can plausibly adopt Clay's existing product-led motion as-is.

### 4.4 TAM → SAM → SOM (Base)

| Base case | Distribution | Manufacturing | Prof. Services |
|---|---:|---:|---:|
| TAM | $131.1M | $280.9M | $385.0M |
| SAM (% of TAM) | $68.2M (52%) | $75.8M (27%) | $148.2M (39%) |
| SOM year-3 ARR | $1.86M | $2.15M | $4.18M |
| SOM customers, year 3 | ~114 | ~125 | ~246 |
| SOM as % of Clay's est. ARR (~$200M, Sep 2026, [B]) | 0.9% | 1.1% | 2.1% |

The SAM conversion rate is the clearest single comparison:
- **Distribution converts 52% of TAM to SAM**, the highest of the three, because sales motions are near-universal.
- **Manufacturing converts only 27%.**

---

## 5. Sensitivity: what moves the answer

**One-at-a-time swings on Base SOM** (Sensitivity tab, Block A):
- **3-yr penetration** is the largest driver for every segment. For distribution it swings SOM from $0.6M to $3.7M.
- **ACV** is second ($1.1M–$3.4M for distribution).
- The two SAM filters matter less: each moves distribution's SOM by ~$0.6–0.7M between its low and high value.

**What would have to be true for distribution to match professional services?** (Block B)
- At identical assumptions, distribution's base SOM is 45% of professional services'.
- Distribution needs a **combined ~2.2× uplift in penetration × ACV** to reach parity.
- Two plausible paths:
  - **Higher ACV from seller density.** Distributors employ ~3× the sales staff per firm, and Clay's pricing scales with usage. If ACV scaled even partially with sellers, parity becomes plausible. → Test in Phase 4 with an explicit usage-driven ACV build.
  - **Higher win rates from a stronger right-to-win.** Data coverage of off-LinkedIn personas (H3) and weaker incumbent fit could lift win rates. → Test in Phase 2.

**Segment definition** (Block C):

| Definition | Distribution | Manufacturing | Prof. Services |
|---|---:|---:|---:|
| Core (model) | 11,494 | 23,785 | 33,035 |
| Broad alternative | 16,432 (+ 4233, 4234) | 60,710 (all 31-33) | 57,665 (all 54) |
| Broad ÷ core | 1.43× | 2.55× | 1.75× |

The ranking by firm count does not change under any broad definition. Distribution remains the smallest even after adding construction materials and commercial-equipment wholesalers.

---

## 6. Hypothesis scorecard after Phase 1

| ID | Hypothesis | Status | Evidence |
|---|---|---|---|
| H1 | "Manufacturing suppliers" too broad; narrow it | **Supported** | Core industrial subsectors ≈ 39% of 20+ manufacturers; the rest is retail/CPG or local-market GTM |
| H2 | Distribution mid-market pool ≥ manufacturing and ≥ 50% of prof. services | **Rejected** | 35% of prof. services; 49% of manufacturing |
| H3 | Industrial personas poorly covered by standard B2B databases | *Phase 2* | Regency evidence only (S07) |
| H4 | Builder capacity, not need, is the binding constraint | **Preliminary support (distribution)** | Highest sales intensity (2.5× avg) with below-average technical share (0.7× avg). Manufacturing is low on **both**, so its issue is need, not capacity |
| H6 | Industrial ACV in the Growth-tier range | *Untested* | Bounded by list price and the Vendr benchmark only |
| H8 | Industrial vertical is an option bet, not a pivot | **Context supports** | Base SOM ~1–2% of est. ARR; even the compounded-High outer bound is ≤ ~12% (prof. services) |
| H9 | Professional services offers better risk-adjusted returns | **Supported on size and readiness** | ~2.2× distribution's base SOM; ~10× the technical talent density. Competition and incumbency still untested |
| **H10 (new)** | Manufacturers are best reached *through* distribution, not as a stand-alone vertical | *Phase 2–3* | Manufacturing has 0.3× average direct-sales intensity; distributors are its sales channel |

**Updated working view** (to be challenged in Phases 2–3): the original question — "should Clay prioritize industrial distribution and manufacturing?" — is now sharper.
- **Manufacturing as a stand-alone vertical looks weak.** It has large firm counts, but little direct selling and the least technical capacity.
- **Distribution is a high-need, low-breadth, low-capacity niche.** It is worth pursuing only if Phase 2 confirms a real data-coverage advantage and Phase 4 shows that usage-driven ACV and a partner-led delivery model close the ~2.2× gap.
- **Professional services is the default to beat.**

---

## 7. Implications for the next phases

- **Phase 2 (competition and needs):**
  - Test H3 directly: how well Apollo, ZoomInfo, HubSpot and industrial-specialist databases cover distributor and contractor buyer personas.
  - Gather builder-capacity evidence beyond occupations: RevOps/"GTM engineer" job postings by industry, and Clay solution partners serving distribution.
  - For professional services, assess incumbency. Clay already markets to the segment (S06), so the *incremental* opportunity may be smaller than SOM suggests.
- **Phase 3 (prioritization):** score market attractiveness on SAM and need intensity together, not firm counts alone. Run the weight sensitivity explicitly against the 2.2× parity gap.
- **Phase 4 (business case):** build ACV bottom-up from sellers per firm and Clay's Actions / Data-Credit meters, so the distribution ACV question is answered with arithmetic rather than assumption.

---

## 8. Limitations

1. **Economic Census, not SUSB.** Same reference year (2022) and unit (firms by firm employment size), but the two programs differ in scope and methodology. Confirm counts against SUSB 2022 offline (QA follow-up).
2. **Data age.** Firm counts are 2022; assumed broadly stable to 2026. Occupational data is May 2025.
3. **OEWS proxies are approximate.**
   - The distribution OEWS proxy (BLS group 4230A1) also includes furniture, construction-materials and miscellaneous durable wholesalers.
   - Seller-doer business development in professional services is invisible in sales-occupation counts.
4. **CRM readiness is the weakest input.**
   - Distribution relies on a single 2022 survey with a small sample.
   - Manufacturing and professional-services values are assumptions. A 2023 vendor-sponsored manufacturing survey was reviewed and rejected: it reported usage frequency among CRM users, not adoption, and did not disclose a sample size.
5. **ACV benchmarks are thin.**
   - The Vendr data is vendor-reported (n=96) and skews toward negotiated, sales-led purchases.
   - Its plan names (Explorer, Pro) predate Clay's March 2026 pricing restructure.
6. **Penetration rates are judgment calls** [D] and the largest driver of SOM. They are ranged widely and will be replaced by a funnel-based build in Phase 4.
7. **No primary research** — no interviews or surveys were conducted, and none are implied.

---

## 9. Phase 1 quality review

| Check | Result |
|---|---|
| Firm vs establishment | ✅ Firms only; establishments shown for contrast (§3.2) |
| Double counting | ✅ Measured and removed within sectors; cross-segment overlap irrelevant because segments are not summed |
| Raw data transcription | ✅ Every size band sums exactly to the published full-year total for all 21 NAICS codes pulled (scripted check) |
| Suppressed data | ✅ Imputed as a residual, flagged yellow, documented |
| Formula integrity | ✅ 1,330 formulas recalculated, 0 errors; sensitivity base reconciles to the TAM_SAM_SOM base (check cell = 0) |
| Hand checks | ✅ Distribution SMB unique firms (6,814 × 0.973 + 2,475 = 9,106); SMB base SOM (9,106 × 0.80 × 0.65 × 1.5% × $7,500 = $0.53M); sales share (500,330 ÷ 2,278,460 = 22.0%) |
| TAM/SAM/SOM consistency | ✅ Same definitions, ACVs and penetration across segments; only the SAM filters differ, each sourced or labeled |
| Facts vs assumptions | ✅ Every input carries a class; scenario outputs labeled as not forecasts |
| Disconfirming evidence | ✅ H2 rejected; the counter-hypothesis (H9) strengthened and stated plainly |

**Open QA items carried forward:** SUSB 2022 cross-check; primary-source confirmation of the September 2026 Series D (carried from Phase 0); CRM-adoption evidence for manufacturing and professional services.
