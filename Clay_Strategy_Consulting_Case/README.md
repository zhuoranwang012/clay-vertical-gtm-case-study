# U.S. Vertical Expansion & Go-to-Market Strategy for Clay

**Independent strategy consulting case study** by Zhuoran (Zora) Wang · Research reference date: October 18, 2026

> This is an independent portfolio project. It is not commissioned by, affiliated with, or endorsed by Clay Labs Inc. All inputs are public information; all forward-looking figures are labeled scenario estimates, not Clay forecasts.

## Strategic question
Over 2027–2029, should Clay commit incremental go-to-market investment to U.S. industrial distribution and/or manufacturing — relative to professional services or a no-vertical baseline — and if so, through which motion and at what scale?

## Evidence classification
| Code | Meaning |
|---|---|
| A | Verified public fact |
| B | External benchmark / third-party estimate |
| C | Analyst estimate (this project) |
| D | Scenario assumption |
| E | Strategic recommendation |

All sources are logged in `Research_Sources.csv` (Source ID, publisher, URL, publication date, data reference year, access date, data point, units, class, reliability, limitations).

## Project structure and status
| File | Phase | Status |
|---|---|---|
| `01_Project_Charter.md` | 0 — Scoping | ✅ v2 amendment (decision hurdles, Option 5, issue-tree branches) |
| `02_Market_Sizing.xlsx` | 1 — Market | ✅ v2 (segment-specific ACV, review M1/M2) |
| `03_Market_Research.md` | 1 — Market | ✅ v2 (numbers updated in place) |
| `04_Competitive_Analysis.xlsx` | 2 — Competition | ✅ Complete |
| `05_Customer_Needs_Assessment.md` | 2 — Needs | ✅ Complete |
| `06_Segment_Prioritization.xlsx` | 3 — Prioritization | ✅ v2 (post review) |
| `07_Financial_Model.xlsx` | 4 — Business case | ✅ v2 (post review) |
| `08_GTM_Strategy.md` | 5 — Strategy | ✅ v2 — "validate first, then decide" |
| `09_Implementation_Roadmap.md` | 5 — Strategy | ✅ v2 — sprint + 3 gates |
| `Executive_Strategy_Deck.pptx` | 6 — Presentation | ✅ v2 (12 main slides + appendix A1–A7) |
| `Research_Sources.csv` | All | 59 entries (S01–S58; S03 split into a/b) |
| `review/Partner_Review_2026-10-19.md` | QA | Skeptical partner review (60/100; issues C1–C5, M1–M7, m1–m8) |
| `appendix/` | All | Raw Census/BLS extracts; build scripts for every workbook and the deck; `clay_engine.py` (Python mirror of the 07 engine + Monte Carlo); `Coverage_Test_Protocol.md` + `Coverage_Test_Log.xlsx` + `coverage_test.py` (M5 contractor coverage test — 100 + 20 contractors sampled from CSLB; observations not yet collected); `export_deck_data.py` + `deck_data_from_model.json`; validation interview guide; `archive/` holds every v1 file |

## Method in one paragraph
Hypothesis-driven and issue-tree-structured. Segments are defined by NAICS and sized bottom-up from 2022 Economic Census firm counts by firm size (not establishments), de-duplicated across sub-industries, filtered to firms with a B2B sales motion Clay can improve. Competitive position is assessed from public vendor documentation and SEC filings. Segments are ranked with a weighted, sensitivity-tested framework, and the recommended option is tested in a funnel-driven three-year scenario model. No interviews, surveys, or internal Clay data were used or fabricated.

## Phase 1 headline — v2 (ACV fix, review M1/M2)
- ACV now follows each segment's sellers per firm and Clay's published usage meters instead of a uniform $7.5K / $20K / $60K. Distribution: $5.4K SMB / $9.7K mid-market / $61.7K enterprise (Base); manufacturing and professional services sit at the $5.4K Growth floor below enterprise.
- Base TAM / SAM: distribution $92M / $41M; manufacturing $162M / $44M; professional services $225M / $87M (v1 TAM: $131M / $281M / $385M).
- Year-3 SOM is ~0.6–1.2% of Clay's estimated ARR in every segment (High case below ~4%): a vertical is an option, not a pivot.

## Phase 3 headline — v2 (rebuilt 2026-10-19, review issue C5)
- One indicator per criterion; Phase 4's usage-based ACV fed back. Distribution 3.8 vs professional services 3.8 vs manufacturing 2.9, with overlapping P10–P90 ranges under random weights (3.5–4.1 vs 3.6–4.1); unrounded 3.79 / 3.85 / 2.91 (v1: 3.97 / 3.41 / 2.70).
- A statistical tie: a 2–6 point shift in any single weight flips the order. Professional services wins 5 of 6 weight sets and 57% of random weightings; distribution wins under concave size scoring, min-max scaling, Upside usage and a working partner channel.
- Tie-breakers favor distribution: Clay's differentiation (competitive position 4.0 vs 2.3) and better economics on all six 07 v2 comparisons. The beachhead is reopened: 4238 machinery/MRO outscores 4236 electrical on v2 indicators.

## Phase 4 headline — v2 (rebuilt 2026-10-19 after partner review; scenario estimates, not Clay forecasts)
- v2 fixes the review's critical issues C1–C4: derived win mix, sales-led pipeline creation, NPV at 12% with an incrementality haircut, Monte Carlo with a noisy gate, professional services and Option 0 on the same engine. v1 is archived in `appendix/archive/`.
- Distribution lean pilot, Base: NPV −$1.89M; exit ARR $0.51M; average first-year ACV $8.8K (v1: $13.8K). Upside NPV +$1.53M, but still below Option 0 (the same spend in Clay's core motion).
- Simulation (5,000 correlated draws): P(NPV > 0) ≈ 0% for every staffed option; ≤ 10% even with 2× wider ranges and strong correlation.
- Professional services is no better (Base NPV −$2.04M). The cheapest test, a solution play with no vertical team, loses least (−$0.54M Base) and is the only option that beats Option 0 in Upside.

## Phase 5 headline — v2 ("validate first, then decide")
- Recommendation [E]: **Option 0 by default** — no vertical cost line in 2027. Spend at most **~$0.2M on a two-quarter validation sprint** (Nov 2026 – Mar 2027) in distribution (4236 + 4238).
- Five pre-set tests: T1 coverage uplift ≥ 1.5×; T2 meeting rate ≥ 4.0% on ≥ 600 accounts; T3 meeting → opportunity ≥ 45%; **T4 implied mid-market ACV ≥ $14.6K from ≥ 5 paid design partners (mandatory)**; T5 ≥ 3 implementation partners.
- Gate 1 (end of Mar 2027): Go to a **no-team solution play** (~$0.3M, Apr–Dec 2027) only if T4 and three others pass. Gate 2 (Q4 2027) continues only on the Upside track (≥ 40 opps, ≥ 7 customers, ACV ≥ $10K, ARR ≥ $80K). Gate 3 (Q4 2028): payback ≤ 18 months and NRR ≥ 100%, else fold into the horizontal motion.
- No vertical lead or dedicated AEs before Gate 2. Maximum loss if the sprint says stop: ~$0.2M (v1 pilot: ~$1.0M).

## Phase 6 — executive deck v2
- Storyline: decision and options → answer first (Option 0 + capped sprint) → market and customers → competition and value → scorecard tie → options vs Option 0 → Monte Carlo → recommended path → sprint tests and decision rule. Appendix A1–A7: v1 → v2 changes, hypothesis verdicts, sizing funnel, sensitivity and what would have to be true, arguments for doing more, assumptions, limitations and sources.
- Every chart and model figure is exported from 07, 06 and 02 v2 by `appendix/export_deck_data.py`. v1 deck archived.
- Rebuild: `python3 appendix/export_deck_data.py && node appendix/build_deck.js appendix/deck_data_from_model.json Executive_Strategy_Deck.pptx`.

## Final QA summary
- All workbooks recalculate with 0 formula errors; the financial model's live sheets reconcile to its scenario engine runs under every selector setting.
- Firms (not establishments) are used throughout; TAM/SAM/SOM definitions are identical across segments; Phase 1 vs Phase 4 differences are reconciled on `07 › Market Sizing`.
- A skeptical partner review (`review/`) overturned the v1 recommendation; v2 fixes C1–C5 and M1/M2, and the deck's A1/A5 slides show what changed and what would change the call.
- Review issue M3 is bounded with Census type-of-operation data: manufacturers' sales branches are 4–22% of the 500+ distribution band; removing them changes no conclusion (02 Block 7, `appendix/msbo_sensitivity.py`).
- Review minor issues m7 (scorecard shown to one decimal with P10–P90 ranges) and m8 (establishments compared with de-duplicated, like-for-like firm counts: 4.4× within the 20+ scope, 1.69× across all sizes) are fixed in 02, 06, 03, 05 and the deck.
- The Python engine reconciles to the Excel engine run by run; deck figures are exported, not retyped.
- Open items: M5 coverage test (designed and sampled 2026-10-20; parked — Apollo/LinkedIn checks need manual logged-in lookups; no results exist); sizing and modeling Option 5; SUSB 2022 cross-check of firm counts; primary source for the Sep 2026 Series D figures; verification of competitor capabilities rated [K]; optional AI module (brief: secondary, not started).

## Key limitations
No internal data on Clay's vertical penetration, win rates, churn, or margins; vendor case studies are selection-biased; latest Census firm counts use the 2022 reference year (SUSB 2022 cross-check pending); Clay's pricing changed in March 2026 and may change again.
