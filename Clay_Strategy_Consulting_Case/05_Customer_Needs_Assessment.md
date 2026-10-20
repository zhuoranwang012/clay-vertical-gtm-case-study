# 05 — Customer Needs Assessment & Value Proposition

**U.S. Vertical Expansion & Go-to-Market Strategy for Clay** · Phase 2
Independent portfolio case study by Zhuoran (Zora) Wang · Not commissioned by or affiliated with Clay Labs Inc. · Research reference date: October 18, 2026

> Companion to `04_Competitive_Analysis.xlsx`. Evidence classes: **[A]** verified public fact · **[B]** external benchmark / third-party data · **[C]** analyst estimate or judgment · **[D]** scenario assumption · **[E]** recommendation. **No customer interviews or surveys were conducted.** Needs are inferred from public surveys, official statistics, vendor case studies and product documentation. A validation interview guide is provided in `appendix/Interview_Guide.md` as a recommended next step.

---

## 1. Answer first

**Clay has a real, evidence-backed advantage in industrial distribution: its buyers' customers are poorly covered by standard B2B databases. But distributors are poorly equipped to adopt Clay as it is sold today. In professional services the reverse holds: easy to adopt, but less differentiated and already contested, partly by Clay's own agency ecosystem.**

| | A. Industrial Distribution | B. Industrial Manufacturing | C. B2B Prof. Services |
|---|---:|---:|---:|
| Need intensity (sales share of employment; all industries 8.6%) [C] | **22.0%** | 2.7% | 6.2% |
| Est. LinkedIn presence of their customers' personas (proxy for database coverage) [C] | **26%** (contractor personas 20%) | 41% | 42% |
| Competitive position score, 1–5 [C] | **4.0** | 3.3 | 2.3 |
| Adoption / execution fit score, 1–5 [C] | 2.0 | 1.5 | **4.5** |
| Incremental headroom flag, 1–5 [C] | 5 | 5 | 2 |

Hypothesis updates:
- **H3 is supported for distribution and not for manufacturing.**
- **H4 is supported.**
- **H5 is supported.**
- **H9 is weakened:** professional services still leads on size and ease, but Clay's incremental headroom there is smaller than Phase 1 implied.
- **H10 holds directionally:** manufacturers are reached more efficiently through distribution.

---

## 2. Segment needs: what each buyer is trying to get done

### A. Industrial distribution — "find and win more contractor and plant accounts in my territories, with reps who live in the field"

| Need / pain point | Evidence | Class |
|---|---|---|
| **New-customer growth is the top priority**, and reps want help choosing which accounts to pursue | 39% name new-customer revenue as their top 12-month priority (vs 34% existing-account growth); 20% say better account prioritization would have the biggest impact; "identifying opportunities and prioritizing accounts" is the #1 AI use case | [B] S25 (Phocas/DSG 2026, ~100+ respondents) |
| **Target buyers are hard to find in standard databases** | Contractor personas (electricians, plumbers, HVAC, construction supervisors) are estimated at ~20% LinkedIn use vs ~42% for professional-services buyers. Regency Supply: "We never would have been able to do that in ZoomInfo" — contractors found via Google Maps listings instead | [C] Persona_Coverage; [A] S07 |
| **CRM data is thin and stale** | 56% use CRM; only 34% of CRM users log ≥75% of customer interactions; 68% rely on spreadsheets; 27% say better CRM adoption is their single biggest lever. 15% had no CRM at all (2022) | [B] S25, S17 |
| **Measurement is immature** | 34% don't measure new-business win rates; 22% don't track forecast accuracy | [B] S25 |
| **Inbound e-commerce sign-ups need vetting** | Regency verified wholesale e-commerce registrations one by one until Clay automated the check (commercial vs residential, address, business status) | [A] S07 |
| **Branch / territory complexity** | 1.69 establishments per unique firm across all sizes, 4.4 among firms with 20+ employees (2.21 per firm in plumbing/HVAC), so account mapping must work at branch level | [A→C] 02_Market_Sizing |
| **Lean technical bench** | Computer occupations are 2.3% of employment (0.69× the economy average); 76% of surveyed sales teams have fewer than 20 reps; 52% are family-owned | [A→C] S15; [B] S25 |
| **The customer system of record is the ERP** | Distributors run Epicor (Eclipse/Prophet 21), Infor and similar. Clay lists CRM and warehouse connectors but no ERP connector | [B] S40 (dated); [A] S36 |

### B. Industrial manufacturing — "keep the plant full; my channel does most of the selling"

| Need / pain point | Evidence | Class |
|---|---|---|
| **Direct selling is a small function** | Sales occupations are 2.7% of employment (0.32× average), ~3.6 sales staff per 20+ firm. Much of the selling runs through distributors and independent reps | [A→C] S15 |
| **Customers are reachable with standard data** | Their buyers (purchasing managers, production managers, engineers) are estimated at ~41% LinkedIn use, about the same as professional services | [C] Persona_Coverage |
| **Demand capture via sourcing platforms** | Job shops use Thomasnet (500K+ suppliers) as an inbound channel; it has no API or export | [B] S27 |
| **Lowest builder capacity of the three** | Computer occupations 1.5% (0.45× average); marketing-ops talent 0.62× average | [A→C] S15 |
| **Enterprise interest exists but runs through partners** | LG Electronics USA (Commercial Display & IT Solutions) is piloting a Clay workflow via Marketbridge | [A] S08 |

**Implication:** manufacturing's direct need is weak and its data gap is small, so it fails both of the tests that make distribution interesting. The stronger route to manufacturers is indirect, through the distributors and reps who sell for them (H10).

### C. B2B professional services — "keep the pipeline of good-fit clients full while billable people do business development"

| Need / pain point | Evidence | Class |
|---|---|---|
| **Business development by "seller-doers," not sales teams** | Sales occupations are only 6.2% of employment, but marketing specialists (31.7 per 1,000) and marketing managers (11.6 per 1,000) are dense | [A→C] S15 |
| **Technically capable buyers** | Computer occupations are 24.2% of employment (7.2× average), close to software publishers | [A→C] S15 |
| **RFP- and relationship-driven pursuits in A&E** | Median A&E proposal win rate 50%; proposal volume fell 38% while awarded value grew 52%; 53% of firms use AI tools (2025, n≈700) | [B] S39 |
| **Vertical systems of record** | Deltek Vantagepoint and Unanet CRM (3,400+ customers, 2022) hold pursuit data in A&E and GovCon | [B] S42 |
| **Already a Clay home market** | "Professional Services" is one of Clay's six case-study industry filters; the 174 companies listed as Clay solution partners are mostly marketing-services and consulting firms; IT services is among the top industries of Clay's detected users | [A] S06; [B] S34, S32 |

**Implication:** professional services can self-serve Clay, but the incremental opportunity is smaller than its firm count suggests. Agencies and consultancies (NAICS 5418/5416) are already Clay customers and partners. They were also the group most exposed to the March 2026 pricing change (S11, opinion), and they can substitute internal builds (n8n and Zapier appear in ~40% of GTM-engineer postings, S30).

---

## 3. Pain points mapped to Clay capabilities

Fit: ●●● strong · ●● partial · ● weak / gap

| Pain point | Segment(s) | Clay capability | Fit | Gap to close |
|---|---|---|---|---|
| Finding contractors and local accounts missing from B2B databases | A | Google Maps integration; Shovels permits and contractor profiles (via Clay); Claygent web research; waterfall across 150–200+ providers | ●●● | Add construction-project data (e.g., Dodge) and license or registry sources as marketplace providers |
| Prioritizing accounts by territory and branch | A | Audiences, scoring, territory-planning use case, signals | ●●● | Branch-level templates |
| Keeping CRM complete when reps don't log | A, B | CRM auto-sync and enrichment (Growth+), job-change tracking (Regency: 5,000+ contacts) | ●●● | Works only if a CRM exists (56% of distributors) |
| Vetting e-commerce and inbound registrations | A | Webhook → AI verification → routing (Regency) | ●●● | Packaged template |
| Cross-selling to existing accounts from order history | A | Warehouse sync / HTTP API only | ● | No ERP connector; Proton.ai and ERP CRMs own this job (S24) |
| Construction project and bid timing | A, B | Public-sector bid signals (GovSpend, Pursuit, Sep 2026), Shovels permits | ●● | Private commercial projects (Dodge) not in marketplace |
| Capturing sourcing-platform demand | B | — | ● | Thomasnet has no API (S27) |
| Researching prospects before pursuits / RFPs | C | Claygent, account agents, pre-call prep (OpenAI case) | ●●● | — |
| Targeted outbound for new clients | C | Audiences, signals, sequencer, ads audiences | ●●● | Strong substitutes (Apollo, internal builds) |
| Operating Clay without technical staff | A, B | Templates, Clay University, solution partners | ● | **The binding gap:** no industrial-focused partners found on the visible directory pages (S35) |

---

## 4. Hypothesis tests (Phase 2)

| ID | Hypothesis | Verdict | Evidence |
|---|---|---|---|
| **H3** | Industrial buyer personas are poorly covered by LinkedIn-derived databases; Clay's multi-source model solves this better | **Supported for distribution; not supported for manufacturing** | Contractor personas: est. 20% LinkedIn use (0.48× prof. services); all distribution personas 26% (0.62×). Manufacturing's customers 41% (0.97×). Corroborated by Regency's direct experience (S07). *Caveat:* education-based proxy, not a measured match rate |
| **H4** | Builder capacity, not need, is the binding constraint | **Supported** | Distribution: highest need (22% sales share) but 0.69× technical-talent index; 76% of surveyed teams under 20 reps (S25); no industrial specialists found on the visible partner-directory pages (S35); GTM-engineer postings concentrate in software (43.9%, S31) |
| **H5** | Clay's advantage is complementary (layered on CRM/data), not substitutive | **Supported** | Clay sells native CRM sync and bring-your-own-key enrichment (S01, S36); Regency layered Clay on HubSpot (S07); Clay has no ERP or CRM of its own to replace incumbents with |
| **H9** | Professional services offers better risk-adjusted returns | **Weakened, still open** | Wins on market size (Phase 1) and adoption fit (4.5 vs 2.0). Loses on competitive position (2.3 vs 4.0) and incremental headroom (2 vs 5). Phase 3 weighting decides |
| **H10** | Manufacturers are best reached through distribution | **Directionally supported** | Manufacturing direct-sales intensity 0.32× average; no data-coverage gap; distributors are its sales channel (S15) |

**Disconfirming evidence actively considered:**
- **Apollo is converging on Clay's model.** It added an 18-provider waterfall (Oct 2025), acquired Pocus and launched a spreadsheet beta in Oct 2026 (S23). Clay's moat in horizontal segments is narrowing, but none of these moves address off-LinkedIn industrial data.
- **"Industrials are low-tech, so they won't buy" is only partly true.** 49% of surveyed distribution sellers already use AI and another 29% plan to within six months (S25). Interest is high; capability is the constraint.
- **Clay's own investment signals still point elsewhere.** The Sep 2026 data expansion targeted public sector, commercial real estate and private equity (S09). This could mean industrials are a deliberate non-priority, or an unaddressed gap. Primary research with Clay's partners and customers would distinguish the two.

---

## 5. Value proposition by segment [E — draft, to be stress-tested in Phases 3–5]

**A. Industrial distribution (primary candidate)**
> *"Find every contractor and plant in your territories, including the ones ZoomInfo and Apollo can't see, and put them in your reps' CRM, scored and ready, without hiring a data team. Unlimited seats, so every outside rep gets the data."*

Proof points:
- Off-LinkedIn sourcing via Google Maps and Shovels.
- Regency Supply's results: contractor lists, automated registration vetting, 5,000+ contacts monitored for job changes.
- Seat economics: at 28 users, Apollo Professional costs ~2.7× Clay Growth's mid configuration at list. At 9 users it is ~0.87×, so this advantage is specific to seller-dense teams.

Required to deliver:
- Pre-built distribution templates (contractor discovery, branch territory scoring, e-commerce vetting, job-change alerts).
- Implementation partners.
- Optionally, an ERP connector.

**B. Industrial manufacturing (indirect)**
> *"Reach the distributors and reps who sell your products, and the OEM buyers they serve."* A secondary use case delivered through partners, not a stand-alone vertical program.

**C. B2B professional services (counter-option)**
> *"Turn your partners' business development into a researched, signal-driven pipeline: account research, pursuit prep and targeted outreach your marketing team can build themselves."*

Main risk: low differentiation versus Apollo and internal builds, plus overlap with segments Clay already serves.

---

## 6. Implications for Phase 3 and beyond

1. **Scoring inputs.** Use competitive position (4.0 / 3.3 / 2.3) and adoption fit (2.0 / 1.5 / 4.5) as separate criteria, and apply the headroom flag (5 / 5 / 2) as an adjustment to market attractiveness. This avoids double-counting ease of execution.
2. **The decision hinges on a trade-off.** Clay's advantage is strongest where adoption is hardest. Phase 3 must show how the ranking moves when weight shifts between competitive position and ease of execution.
3. **Phase 4 must cost the delivery gap.** Price templates, partners and onboarding explicitly. Test whether usage-driven ACV from seller-dense distributors (~28 sellers per firm) pays for them.
4. **Primary research would most improve the analysis on:** actual match rates for contractor contacts across Apollo, ZoomInfo and Clay waterfalls; distributor willingness to pay; and the extent of current Clay usage in agencies. See `appendix/Interview_Guide.md`.

---

## 7. Limitations

- **No primary research.** Needs are inferred, not heard from customers.
- **The H3 proxy is indirect.** It combines BLS educational attainment (S29) with Pew's 2023 LinkedIn usage by education (S28) and assumes usage depends only on education. Real database coverage also comes from web crawling and contributory networks.
- **Distribution survey evidence is small-sample.** Phocas 2026 (~100+ respondents, undisclosed sample size and dates) and DSG 2022 (n=96) are industry-sponsored.
- **Competitor capability ratings marked [K]** come from general product knowledge and need verification against current vendor documentation.
- **Clay usage detection by Bloomberry is narrow.** It covers only web-intent users (1,701 companies), not Clay's full 17,000+ customer base.
- **Pricing comparisons use list prices.** Discounts, add-ons and credit overages are excluded. ZoomInfo and HubSpot seat prices are not public and are not estimated.
