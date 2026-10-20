// Build Executive_Strategy_Deck.pptx — v2 (after partner review, 2026-10-19). Storyline: validate first, then decide.
// Charts read the v2 workbook outputs exported by appendix/export_deck_data.py, so slides match 02, 06 and 07.
// Usage: python3 appendix/export_deck_data.py && node appendix/build_deck.js appendix/deck_data_from_model.json Executive_Strategy_Deck.pptx
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const { applyTheme } = require("/mnt/skills/public/pptx/scripts/apply_theme.js");

const D = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
const OUT = process.argv[3];

const THEME = {
  name: "Clay Case Study",
  headFontFace: "Cambria",
  bodyFontFace: "Arial",
  colors: {
    dk1: "1B2433", lt1: "FFFFFF", dk2: "3D4A5C", lt2: "F2F4F7",
    accent1: "B5532A", accent2: "5B6B82", accent3: "A9B4C2",
    accent4: "2E6F5E", accent5: "B7862B", accent6: "9B2C2C",
    hlink: "B5532A", folHlink: "5B6B82",
  },
};
const HEX = { ink: "1B2433", ink2: "3D4A5C", slate: "5B6B82", mist: "A9B4C2", copper: "B5532A", pale: "F2F4F7", green: "2E6F5E", red: "9B2C2C", grid: "E3E7EC" };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.author = "Zhuoran (Zora) Wang";
pres.company = "Independent case study";
pres.title = "U.S. Vertical Expansion and GTM Strategy for Clay — Executive Deck (v2)";
const C = pres.SchemeColor;
const ML = 0.55, CW = 12.23;

const FOOT = "Independent case study · not commissioned by or affiliated with Clay Labs Inc. · scenario figures are estimates, not Clay forecasts";
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: ML, y: 0.32, w: CW, h: 0.95, fontFace: "Cambria", fontSize: 24, bold: true, color: C.text1, valign: "top", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "source", type: "body", x: ML, y: 6.68, w: 11.4, h: 0.36, fontSize: 10, color: C.accent2, valign: "top", margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: ML, y: 7.1, w: 11.0, h: 0.24, fontSize: 10, color: C.accent2, margin: 0 } } },
  ],
  slideNumber: { x: 12.23, y: 7.1, w: 0.55, h: 0.24, fontSize: 10, color: C.accent2, align: "right" },
});
pres.defineSlideMaster({
  title: "TITLE",
  background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: ML, y: 0.85, w: 11.5, h: 1.5, fontFace: "Cambria", fontSize: 40, bold: true, color: C.text1, valign: "bottom", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "subtitle", type: "body", x: ML, y: 2.5, w: 11.5, h: 0.6, fontSize: 20, color: C.text2, valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "DIVIDER",
  background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: ML, y: 2.4, w: 11.5, h: 1.0, fontFace: "Cambria", fontSize: 40, bold: true, color: C.text1, valign: "bottom", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "subtitle", type: "body", x: ML, y: 3.6, w: 11.5, h: 2.6, fontSize: 16, color: C.text2, valign: "top", margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: ML, y: 7.1, w: 11.0, h: 0.24, fontSize: 10, color: C.accent2, margin: 0 } } },
  ],
  slideNumber: { x: 12.23, y: 7.1, w: 0.55, h: 0.24, fontSize: 10, color: C.accent2, align: "right" },
});

// ------------------------------------------------------------------ helpers
let objN = 0;
const nm = (p) => `${p}-${++objN}`;
function T(slide, text, o) {
  slide.addText(text, Object.assign({ isTextBox: true, margin: 0, color: C.text1, fontSize: 14, valign: "top", objectName: nm("text") }, o));
}
function box(slide, x, y, w, h, fill, transparency) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill, transparency: transparency || 0 }, line: { type: "none" }, objectName: nm("panel") });
}
function rbox(slide, x, y, w, h, fill, transparency) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.06, fill: { color: fill, transparency: transparency || 0 }, line: { type: "none" }, objectName: nm("card") });
}
function stat(slide, x, y, w, big, label, color, bigSize, labelH) {
  T(slide, big, { x, y, w, h: 0.5, fontFace: "Cambria", fontSize: bigSize || 26, bold: true, color: color || C.accent1 });
  T(slide, label, { x, y: y + 0.5, w, h: labelH || 0.6, fontSize: 12, color: C.text2 });
}
function chip(slide, x, y, w, text, fill) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.3, rectRadius: 0.08, fill: { color: fill || C.accent2 }, line: { type: "none" }, objectName: nm("chip") });
  T(slide, text, { x, y, w, h: 0.3, fontSize: 10, bold: true, color: C.background1, align: "center", valign: "middle" });
}
function content(title, source, section) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  s.addText(source, { placeholder: "source" });
  return s;
}
function sowhat(s, y, h, head, body) {
  box(s, ML, y, CW, h, C.accent1, 91);
  T(s, [{ text: head + "  ", options: { bold: true, color: C.accent1 } }, { text: body }], { x: ML + 0.25, y: y + 0.12, w: CW - 0.5, h: h - 0.2, fontSize: 14 });
}
const M = (v, d = 2) => (v < 0 ? "−" : "+") + "$" + Math.abs(v / 1e6).toFixed(d) + "M";
const Mu = (v, d = 1) => "$" + (v / 1e6).toFixed(d) + "M";
const K = (v) => "$" + (v / 1000).toFixed(1) + "K";
const P0 = (v) => (v * 100).toFixed(0) + "%";
const CHART_TXT = { catAxisLabelColor: HEX.ink2, valAxisLabelColor: HEX.ink2, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", catAxisLabelFontSize: 11, valAxisLabelFontSize: 11, dataLabelFontFace: "+mn-lt", dataLabelFontSize: 11, dataLabelColor: HEX.ink, titleFontFace: "+mn-lt", titleFontSize: 12, titleColor: HEX.ink, legendFontFace: "+mn-lt", legendFontSize: 11, legendColor: HEX.ink2 };
const QUIET = { valGridLine: { color: HEX.grid, size: 0.75 }, catGridLine: { style: "none" }, catAxisLineShow: true, valAxisLineShow: false };
const NOGRID = { valGridLine: { style: "none" } };
const TBL = { fontFace: "Arial", fontSize: 12, color: HEX.ink, border: { type: "solid", pt: 0.75, color: "D5DAE1" }, valign: "middle", margin: [3, 6, 3, 6] };
const th = (text, o) => ({ text, options: Object.assign({ bold: true, color: "FFFFFF", fill: { color: HEX.ink2 }, align: "center" }, o || {}) });
const td = (text, o) => ({ text, options: Object.assign({}, o || {}) });
const tdc = (text, o) => td(text, Object.assign({ align: "center" }, o || {}));

// model accessors
const SC = D.scen;
const col = (seg, mot, scn) => `${seg} · ${mot} · ${scn}`;
const MOT = { lean: "Lean partner-led pilot", sales: "Sales-led vertical team", play: "Solution play" };
const v = (metric, seg, mot, scn) => SC[metric][col(seg, MOT[mot], scn)];
const NPV = "NPV of the program at hurdle rate", XS = "NPV advantage over Option 0", EXIT = "Exit ARR, Q4 2029", ACV = "Average first-year ACV";
const MC = D.mc;
const mcRow = (seg, key) => MC[Object.keys(MC).find((k) => k.startsWith(seg) && k.includes(key))];
const SZ = D.sizing;
const S6 = D.score;
const R6 = D.robust6;
const N0 = (x) => Math.round(x).toLocaleString("en-US");

// ================================================================== 1. TITLE
{
  pres.addSection({ title: "Decision" });
  const s = pres.addSlide({ masterName: "TITLE", sectionTitle: "Decision" });
  s.addText("Should Clay build a U.S. industrial vertical?", { placeholder: "title" });
  s.addText("Vertical expansion and go-to-market strategy, 2027–2029 · revised after partner review", { placeholder: "subtitle" });
  T(s, "The decision: should Clay place incremental vertical GTM investment in 2027 — and if so, where and through which motion?", { x: ML, y: 3.4, w: 11.8, h: 0.4, fontSize: 16, bold: true });
  const opts = [
    ["Option 0", "No vertical bet — stay horizontal", 1],
    ["Option 1", "Industrial distribution — sales-led vertical team", 0],
    ["Option 2", "Manufacturing focus", 0],
    ["Option 3", "Professional services focus", 0],
    ["Option 4", "Distribution: validate first (≤ $0.2M), then a no-team solution play only on evidence", 2],
  ];
  const cw = 2.3, gap = 0.18;
  opts.forEach((o, i) => {
    const x = ML + i * (cw + gap);
    rbox(s, x, 4.0, cw, 1.6, o[2] === 1 ? C.accent1 : (o[2] === 2 ? C.accent1 : C.background2), o[2] === 2 ? 85 : 0);
    const dark = o[2] === 1;
    T(s, o[0], { x: x + 0.15, y: 4.12, w: cw - 0.3, h: 0.3, fontSize: 12, bold: true, color: dark ? C.background1 : C.accent2 });
    T(s, o[1], { x: x + 0.15, y: 4.45, w: cw - 0.3, h: 1.1, fontSize: 13, bold: o[2] > 0, color: dark ? C.background1 : C.text1 });
  });
  T(s, "Recommended default", { x: ML, y: 5.68, w: cw, h: 0.3, fontSize: 12, bold: true, color: C.accent1 });
  T(s, "Only as a capped test", { x: ML + 4 * (cw + gap), y: 5.68, w: cw, h: 0.3, fontSize: 12, bold: true, color: C.accent1 });
  T(s, "Independent strategy case study · Zhuoran (Zora) Wang · October 2026 · v2", { x: ML, y: 6.3, w: 11.5, h: 0.3, fontSize: 14, color: C.text2 });
  T(s, "Not commissioned by, affiliated with, or endorsed by Clay Labs Inc. Built entirely from public data; no interviews or internal Clay data were used. All forward-looking figures are scenario estimates, not Clay forecasts.", { x: ML, y: 6.68, w: 11.5, h: 0.5, fontSize: 10, color: C.accent2 });
  s.addNotes("Five options against an explicit baseline. The revised recommendation is Option 0 by default, with a capped validation sprint in distribution as the only vertical spend. v1 recommended a $1.1M pilot; a partner review found the analysis tilted toward that answer, and the corrected model reverses it.");
}

// ================================================================== 2. EXEC SUMMARY
{
  const s = content("Recommendation: don't fund a vertical in 2027 — spend at most ~$0.2M to test the assumptions that could change that",
    "Sources: 07_Financial_Model v2 (NPV at 12%, Monte Carlo; scenario estimates [C/D]); 06_Segment_Prioritization v2; 02_Market_Sizing v2; Census 2022 (S14); BLS OEWS May 2025 (S15).", "Decision");
  const dl = v(NPV, "Dist", "lean", "Base"), ds = v(NPV, "Dist", "sales", "Base");
  const st = mcRow("Industrial distribution", "staged");
  const rows = [
    ["Distribution is where Clay is most differentiated", "Distributors' customers are hard to find in standard databases (competitive position 4.0 vs 2.3 for professional services), and 22% of distribution staff sell."],
    ["But at list prices the economics don't work", `Usage-based ACV is ~$5.4K (SMB) and ~$9.7K (mid-market). Distribution lean pilot NPV ${M(dl)} at 12%; a sales-led team ${M(ds)}.`],
    ["No staffed option beats putting the money into Clay's core", `Every staffed option loses to Option 0, and P(NPV > 0) is ≈${P0(st.ppos)} across 5,000 correlated draws (≤ 10% even with 2× wider ranges).`],
    ["So validate first, then decide", "A ≤ $0.2M, two-quarter sprint in 4236 and 4238 measures ACV, demand, conversion, coverage and partners. Go to a no-team solution play only if pre-set thresholds pass; otherwise stop."],
  ];
  rows.forEach((r, i) => {
    const y = 1.5 + i * 1.27;
    s.addShape(pres.shapes.OVAL, { x: ML, y, w: 0.46, h: 0.46, fill: { color: C.text1 }, line: { type: "none" }, objectName: nm("num") });
    T(s, String(i + 1), { x: ML, y, w: 0.46, h: 0.46, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle" });
    T(s, r[0], { x: ML + 0.7, y: y - 0.02, w: 7.2, h: 0.35, fontSize: 16, bold: true });
    T(s, r[1], { x: ML + 0.7, y: y + 0.36, w: 7.2, h: 0.8, fontSize: 14, color: C.text2 });
  });
  box(s, 8.75, 1.45, 4.03, 5.05, C.accent1, 91);
  T(s, "The decision", { x: 9.0, y: 1.62, w: 3.6, h: 0.35, fontSize: 16, bold: true, color: C.accent1 });
  [["Option 0", "default: no vertical cost line in 2027"], ["≤ $0.2M", "validation sprint cap, Nov 2026 – Mar 2027"], ["≥ $14.6K", "mid-market ACV the sprint must show (mandatory test)"], ["~$0.2M", "maximum loss if the sprint says stop (v1 pilot: ~$1.0M)"]]
    .forEach((x, i) => stat(s, 9.0, 2.1 + i * 1.08, 3.6, x[0], x[1], C.accent1, 22, 0.5));
  s.addNotes("Answer first. Clay's edge is real in distribution, but list-price ACV is too low for any staffed program to clear the hurdle or beat the core. The only defensible spend is a small sprint that measures the inputs the model has to assume.");
}

// ================================================================== 3. MARKET
pres.addSection({ title: "Market & customers" });
{
  const s = content("Distribution has the fewest firms but by far the most sellers per firm",
    "Sources: U.S. Census Bureau, 2022 Economic Census, Employment Size of Firms (S14) — firms, not establishments, de-duplicated [C]; BLS OEWS May 2025 (S15). Cores: NAICS 4235–4238 · 326/332/333/335 · 5413/5415/5416/5418.", "Market & customers");
  const cats = ["Distribution", "Manufacturing", "Prof. services"];
  const cols = [HEX.copper, HEX.slate, HEX.slate];
  [{ t: "Unique U.S. firms, 20+ employees (2022)", v: SZ["Unique firms, 20+ employees"].map(Math.round), f: "#,##0" },
   { t: "Sales jobs, % of employment (all industries 8.6%)", v: SZ["Sales occupations share of employment (all industries ≈ 8.6%)"].map((x) => Math.round(x * 1000) / 10), f: '0.0"%"' },
   { t: "Sales staff per firm, 20+ employees", v: SZ["Implied sales staff per 20+ firm"].map((x) => Math.round(x * 10) / 10), f: "0.0" }].forEach((ch, i) => {
    s.addChart(pres.charts.BAR, [{ name: ch.t, labels: cats, values: ch.v }], Object.assign({
      x: ML + i * 4.15, y: 1.45, w: 3.95, h: 3.75, barDir: "col", chartColors: cols, showTitle: true, title: ch.t,
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: ch.f, valAxisHidden: true, showLegend: false, barGapWidthPct: 60, objectName: nm("chart"),
    }, CHART_TXT, QUIET, NOGRID));
  });
  sowhat(s, 5.45, 1.0, "So what", "Distribution trades market breadth for need intensity: an in-scope distributor employs ~28 sales staff, ~3× a professional-services firm. Tools priced per seat charge for every one of them; Clay's plans do not. Whether that density turns into revenue is the question the rest of the deck tests.");
  s.addNotes("Distribution is the smallest pool by firm count, but its employment is dominated by sellers. Counts are firms, not establishments.");
}

// ================================================================== 4. SIZING
{
  const s = content("At realistic ACVs, each vertical is worth ~0.6–1.2% of Clay's ARR — an option, not a pivot",
    "Sources: 02_Market_Sizing v2 — Census 2022 (S14); usage-based ACV at Clay list prices (S01); CRM in use 56% for distribution (S25). Clay ARR ~$200M is a Sacra estimate (S23). Penetration is uniform across segments by design, so SOM is an assumption-driven sizing [C/D].", "Market & customers");
  const cats = ["Distribution", "Manufacturing", "Prof. services"];
  s.addChart(pres.charts.BAR, [
    { name: "TAM", labels: cats, values: SZ["TAM (Base)"].map((x) => Math.round(x / 1e5) / 10) },
    { name: "SAM", labels: cats, values: SZ["SAM (Base)"].map((x) => Math.round(x / 1e5) / 10) },
  ], Object.assign({ x: ML, y: 1.45, w: 6.6, h: 4.3, barDir: "col", barGrouping: "clustered", chartColors: [HEX.mist, HEX.slate], showTitle: true, title: "TAM and SAM by segment, Base ($M per year)",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '"$"0', valAxisHidden: true, showLegend: true, legendPos: "b", barGapWidthPct: 70, objectName: nm("chart") }, CHART_TXT, QUIET, NOGRID));
  const som = SZ["SOM — year-3 ARR potential (Base)"];
  const rows = [[th("Segment", { align: "left" }), th("SOM, year-3 ARR"), th("% of Clay est. ARR"), th("ACV: SMB / MM / Ent")],
    [td("Distribution", { bold: true }), tdc(Mu(som[0])), tdc(P0(som[0] / 2e8 * 100 / 100).replace("%", "") === "1" ? "0.6%" : (som[0] / 2e6).toFixed(1) + "%"), tdc("$5.4K / $9.7K / $61.7K", { bold: true })],
    [td("Manufacturing"), tdc(Mu(som[1])), tdc((som[1] / 2e6).toFixed(1) + "%"), tdc("$5.4K / $5.4K / $44.3K")],
    [td("Professional services"), tdc(Mu(som[2])), tdc((som[2] / 2e6).toFixed(1) + "%"), tdc("$5.4K / $5.4K / $44.3K")]];
  rows[1][2] = tdc((som[0] / 2e6).toFixed(1) + "%");
  s.addTable(rows, Object.assign({ x: 7.45, y: 1.5, w: 5.33, colW: [1.55, 1.1, 1.05, 1.63], rowH: 0.42, objectName: nm("table") }, TBL, { fontSize: 11 }));
  box(s, 7.45, 3.45, 5.33, 2.3, C.background2, 0);
  T(s, [
    { text: "What changed (v2)  ", options: { bold: true, color: C.accent1, breakLine: true } },
    { text: "ACV now follows each segment's sellers per firm and Clay's published usage meters, instead of a uniform $7.5K / $20K / $60K. SMB sits at the $5,352 Growth floor everywhere; only distribution's mid-market rises above it.", options: { breakLine: true } },
    { text: " ", options: { breakLine: true, fontSize: 6 } },
    { text: "v1 TAM: $131M / $281M / $385M. Even the compounded High SOM stays below ~4% of Clay's ARR." },
  ], { x: 7.65, y: 3.58, w: 4.95, h: 2.1, fontSize: 12 });
  sowhat(s, 5.95, 0.55, "Refined in Phase 4", "Program SAM: ~4,450 distributors (NAICS 4236–4238, CRM in use, not yet Clay customers).");
  s.addNotes("Sized bottom-up with usage-based ACV. Professional services is the largest pool on size alone, but no vertical moves Clay's growth materially. SOM uses an assumed uniform penetration, so it is a sizing, not a finding.");
}

// ================================================================== 5. SEGMENTATION
{
  const s = content("Clay's edge is largest where adoption is hardest: distributors' customers are least visible online and their teams least technical",
    "Sources: BLS OEWS May 2025 (S15); persona coverage = BLS education mix (S29) × Pew LinkedIn use by education (S28), an indirect proxy [C]; scores (1–5) from 04_Competitive_Analysis [C].", "Market & customers");
  const G = { fill: { color: "E2EEE9" } }, R = { fill: { color: "F5E4E4" } }, N = {};
  const c = (t, o, b) => tdc(t, Object.assign({ bold: !!b }, o));
  const rows = [
    [th("Indicator", { align: "left" }), th("Distribution"), th("Manufacturing"), th("Prof. services")],
    [td("Need: sales occupations, % of employment (all industries 8.6%)"), c("22.0%", G, 1), c("2.7%", R), c("6.2%", N)],
    [td("Data gap: est. LinkedIn presence of customers' personas (lower = harder for databases)"), c("26% (contractors 20%)", G, 1), c("41%", N), c("42%", N)],
    [td("Competitive position (1–5)"), c("4.0", G, 1), c("3.3", N), c("2.3", R)],
    [td("Builder capacity: computer occupations, % of employment (all 3.4%)"), c("2.3%", R), c("1.5%", R), c("24.2%", G, 1)],
    [td("Adoption / execution fit (1–5)"), c("2.0", R), c("1.5", R), c("4.5", G, 1)],
    [td("Usage-based ACV per won account (list prices)"), c("$8.9K", G, 1), c("$7.1K", N), c("$7.1K", N)],
  ];
  s.addTable(rows, Object.assign({ x: ML, y: 1.5, w: 8.45, colW: [3.95, 1.6, 1.45, 1.45], rowH: 0.62, objectName: nm("table") }, TBL));
  T(s, "Green = favorable for Clay · red = unfavorable", { x: ML, y: 6.02, w: 8.45, h: 0.3, fontSize: 11, color: C.text2 });
  box(s, 9.3, 1.5, 3.48, 4.85, C.background2, 0);
  [["Fit: distribution", "Highest need, biggest data gap, strongest competitive position and the highest ACV per account."],
   ["Ease: professional services", "Technical buyers who can self-serve — but already a Clay home market, so more wins would come anyway."],
   ["Manufacturing", "Fails both tests; reach it through the distributors that sell for it (H10)."]].forEach((it, i) => {
    T(s, it[0], { x: 9.5, y: 1.7 + i * 1.5, w: 3.1, h: 0.32, fontSize: 14, bold: true, color: i === 0 ? C.accent1 : C.text1 });
    T(s, it[1], { x: 9.5, y: 2.05 + i * 1.5, w: 3.1, h: 1.05, fontSize: 13, color: C.text2 });
  });
  s.addNotes("The core tension: the segment where Clay is most differentiated is also the hardest to adopt in. Distribution's ACV per account is higher only because of its mid-market sellers.");
}

// ================================================================== 6. COMPETITION
pres.addSection({ title: "Competition & value" });
{
  const s = content("Only Clay combines off-LinkedIn data with workflow automation — but it is the hardest option to implement",
    "Sources: 04_Competitive_Analysis — scores are analyst judgments from public vendor documentation [C]; some capabilities rated from general knowledge [K], pending verification. Phocas/DSG 2026 (S25); Apollo pricing (S20); Sacra (S23).", "Competition & value");
  const heat = { 5: HEX.slate, 4: "8C99AB", 3: "BEC6D1", 2: "DDE2E8", 1: "F0F2F5" };
  const sc = (x) => tdc(String(x), { bold: true, fill: { color: heat[x] }, color: x >= 5 ? "FFFFFF" : HEX.ink });
  const data = [["Data coverage of industrial buyers", 5, 2, 2, 2, 4], ["Workflow automation", 5, 3, 3, 4, 1], ["Differentiation for industrial sellers", 5, 3, 3, 2, 4],
    ["Pricing architecture", 4, 3, 2, 3, 3], ["Ease of implementation (5 = easiest)", 2, 5, 3, 5, 4]];
  const rows = [[th("Dimension (score 1–5)", { align: "left" }), th("Clay", { fill: { color: HEX.copper } }), th("Apollo"), th("ZoomInfo"), th("HubSpot (Breeze)"), th("Vertical data (Dodge, Shovels)")]]
    .concat(data.map((d) => [td(d[0])].concat(d.slice(1).map(sc))));
  s.addTable(rows, Object.assign({ x: ML, y: 1.5, w: 8.3, colW: [2.9, 0.95, 0.95, 1.05, 1.15, 1.3], rowH: 0.7, objectName: nm("table") }, TBL));
  T(s, "Judgment scores, not measurements: the coverage advantage is what the validation sprint must prove (test T1).", { x: ML, y: 5.85, w: 8.3, h: 0.6, fontSize: 13, color: C.text2 });
  box(s, 9.15, 1.5, 3.63, 4.95, C.background2, 0);
  stat(s, 9.35, 1.65, 3.25, "68%", "of distributors rely on spreadsheets: the status quo is the main competitor (S25)", C.accent1, 26, 0.75);
  stat(s, 9.35, 3.2, 3.25, "2.7×", "Apollo Professional's list cost vs Clay Growth for a 28-user team; 0.87× at 9 users [C]", C.accent1, 26, 0.75);
  stat(s, 9.35, 4.75, 3.25, "6 of 11", "Clay use cases Apollo now also covers; off-LinkedIn contractor data is the gap it hasn't closed [C]", C.accent1, 26, 0.75);
  s.addNotes("In distribution the realistic alternative is spreadsheets plus a generic database. Clay's weakness is implementation, which shapes the motion; its claimed coverage edge is a judgment until measured.");
}

// ================================================================== 7. VALUE & ACV
{
  const s = content("Clay's edge is coverage, but at list prices it monetizes only from mid-market distributors upward",
    "Source: 07_Financial_Model v2, ACV Build — sellers per firm (OEWS sales share, S15) × 150 accounts × usage meters at published Growth rates (S01); Enterprise floor and 30% discount (S01, S16). Distribution, Base [C/D]. Needs: Phocas/DSG 2026 (S25).", "Competition & value");
  const lab = ["SMB (20–99)", "Mid-market (100–499)", "Enterprise (500+)"];
  s.addChart(pres.charts.BAR, [
    { name: "Usage at list price", labels: lab, values: [1.95, 9.77, 88.25] },
    { name: "First-year ACV (after plan floor / discount)", labels: lab, values: [5.35, 9.77, 61.78] },
  ], Object.assign({ x: ML, y: 1.45, w: 6.4, h: 4.95, barDir: "col", barGrouping: "clustered", chartColors: [HEX.mist, HEX.copper], showTitle: true, title: "Annual value per distributor, Base ($K)",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '"$"0.0"K"', valAxisHidden: true, showLegend: true, legendPos: "b", barGapWidthPct: 60, objectName: nm("chart") }, CHART_TXT, QUIET, NOGRID));
  box(s, 7.3, 1.45, 5.48, 1.4, C.accent1, 91);
  T(s, "“Find the contractors and plants in your territories that your current database misses, and put them in your reps' CRM, scored by branch. No data team, no per-seat fees.”",
    { x: 7.5, y: 1.57, w: 5.1, h: 1.2, fontSize: 14, italic: true });
  T(s, "What the value proposition still has to prove", { x: 7.3, y: 3.05, w: 5.48, h: 0.32, fontSize: 14, bold: true });
  T(s, [
    { text: "Coverage: measured uplift vs the best single database (v1's “ZoomInfo and Apollo can't see them” is withdrawn until measured)", options: { bullet: true, breakLine: true } },
    { text: "Usage: committed credits worth ≥ $14.6K for a mid-market distributor (1.5× the base build)", options: { bullet: true, breakLine: true } },
    { text: "Demand: Clay-run outbound meeting rate ≥ 4% with contractor-facing buyers", options: { bullet: true, breakLine: true } },
    { text: "Delivery: partners willing to specialize — builder capacity is the adoption constraint", options: { bullet: true } },
  ], { x: 7.3, y: 3.45, w: 5.48, h: 2.6, fontSize: 12, paraSpaceAfter: 5, color: C.text2 });
  T(s, "Not claimed: AI cross-sell from order history, replacing CRM/ERP, or efficiency gains before measurement.", { x: 7.3, y: 6.08, w: 5.48, h: 0.45, fontSize: 11, italic: true, color: C.text2 });
  s.addNotes("SMB distributors would use only about $2K of Clay a year at list, so they pay the Growth floor. Seller density matters only from mid-market up, which is why ACV is the swing factor and why the sprint's mandatory test is committed usage.");
}

// ================================================================== 8. PRIORITIZATION
pres.addSection({ title: "Prioritization & economics" });
{
  const s = content("Rebuilt without double counting, the segment scorecard is a tie — the tie-breakers favor distribution",
    "Source: 06_Segment_Prioritization v2 — one indicator per criterion; ratio-to-max scaling; brief weights; contribution = score × weight [C]. Totals to one decimal; (range) = P10–P90 across 1,000 centered random weightings. v1 (4.0 / 2.7 / 3.4) archived. Dirichlet, seed 42.", "Prioritization & economics");
  const crit = [["Market attractiveness (20%)", "Market Attractiveness", 0.20], ["Customer need (20%)", "Customer Need", 0.20], ["Product-market fit (20%)", "Product-Market Fit", 0.20],
    ["Revenue potential (15%)", "Revenue Potential", 0.15], ["Competitive position (15%)", "Competitive Position", 0.15], ["Ease of execution (10%)", "Ease of Execution", 0.10]];
  const tot = S6.total;
  const p10 = R6["Score range — P10, centered random weights"], p90 = R6["Score range — P90, centered random weights"];
  const lab = (n, j) => `${n} · ${tot[j].toFixed(1)} (${p10[j].toFixed(1)}–${p90[j].toFixed(1)})`;
  const labels = [lab("Manufacturing", 1), lab("Prof. services", 2), lab("Distribution", 0)];
  const order = [1, 2, 0];
  const series = crit.map((c) => ({ name: c[0], labels, values: order.map((j) => Math.round(S6[c[1]][j] * c[2] * 100) / 100) }));
  s.addChart(pres.charts.BAR, series, Object.assign({ x: ML, y: 1.45, w: 7.7, h: 4.95, barDir: "bar", barGrouping: "stacked", chartColors: [HEX.ink, HEX.slate, HEX.mist, HEX.copper, "D9A48A", "7E9C8F"],
    showTitle: true, title: "Weighted score by criterion, v2 (maximum 5.0)", showLegend: true, legendPos: "b", valAxisMaxVal: 5, valAxisMinVal: 0, valAxisMajorUnit: 1, barGapWidthPct: 55, objectName: nm("chart") }, CHART_TXT, QUIET));
  box(s, 8.6, 1.45, 4.18, 4.95, C.background2, 0);
  T(s, "Robustness is two-sided", { x: 8.82, y: 1.6, w: 3.8, h: 0.35, fontSize: 16, bold: true });
  const ws = R6["Named weight sets won (of 6)"], fl = R6["Random weights won — flat Dirichlet"], stw = R6["Stress tests won (of 14)"];
  [[`${ws[0]} vs ${ws[2]}`, "named weight sets won by distribution vs professional services"],
   [`${P0(fl[0])} / ${P0(fl[2])}`, "of random weightings won (flat Dirichlet)"],
   [`${stw[0]} vs ${stw[2]}`, "of 14 stress tests won; distribution wins with concave size scoring or a working partner channel"],
   ["2–6 pts", "shift in any single weight flips the order"]].forEach((x, i) => stat(s, 8.82, 2.08 + i * 1.07, 3.8, x[0], x[1], C.accent1, 20, 0.5));
  s.addNotes("v1 counted seller density three times and the coverage proxy twice. With one indicator per criterion, distribution and professional services tie. The tie-breakers are differentiation and economics, both favoring distribution.");
}

// ================================================================== 9. OPTIONS ON ONE ENGINE
{
  const s = content("On one engine, no staffed option clears a 12% hurdle or beats putting the same money into Clay's core",
    "Source: 07_Financial_Model v2 — quarterly funnel model 2027–29; NPV at 12% after an incrementality haircut; Option 0 = the same non-CS spend in Clay's core motion at a discounted benchmark payback (S47). Deterministic Base and Upside [C/D].", "Prioritization & economics");
  const opts = [["Dist", "lean", "Dist · lean pilot"], ["Dist", "sales", "Dist · sales-led team"], ["Dist", "play", "Dist · solution play"], ["PS", "lean", "PS · lean pilot"], ["PS", "sales", "PS · sales-led team"], ["PS", "play", "PS · solution play"]];
  s.addChart(pres.charts.BAR, [
    { name: "Base NPV", labels: opts.map((o) => o[2]), values: opts.map((o) => Math.round(v(NPV, o[0], o[1], "Base") / 1e4) / 100) },
    { name: "Upside NPV", labels: opts.map((o) => o[2]), values: opts.map((o) => Math.round(v(NPV, o[0], o[1], "Upside") / 1e4) / 100) },
  ], Object.assign({ x: ML, y: 1.45, w: 7.4, h: 4.95, barDir: "col", barGrouping: "clustered", chartColors: [HEX.slate, HEX.copper], showTitle: true, title: "NPV at 12% by option ($M)",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0.0', showLegend: true, legendPos: "b", barGapWidthPct: 50, valAxisLabelFormatCode: '0', objectName: nm("chart") }, CHART_TXT, QUIET));
  const rows = [[th("Option", { align: "left" }), th("Beats Option 0? Base"), th("Upside")]].concat(opts.map((o) => {
    const b = v(XS, o[0], o[1], "Base"), u = v(XS, o[0], o[1], "Upside");
    const f = (x) => tdc(x >= 0 ? "Yes " + M(x, 1) : "No " + M(x, 1), { color: x >= 0 ? HEX.green : HEX.red, bold: x >= 0 });
    return [td(o[2]), f(b), f(u)];
  }));
  s.addTable(rows, Object.assign({ x: 8.25, y: 1.5, w: 4.53, colW: [1.85, 1.34, 1.34], rowH: 0.42, objectName: nm("table") }, TBL, { fontSize: 11 }));
  box(s, 8.25, 4.65, 4.53, 1.75, C.accent1, 91);
  T(s, [{ text: "Reading  ", options: { bold: true, color: C.accent1 } },
    { text: `Fixed team cost is what sinks staffed options. The cheapest structure — templates, data providers, a partner listing, no vertical team — loses least (Base ${M(v(NPV, "Dist", "play", "Base"))}) and is the only one that beats Option 0, in distribution Upside.` }],
    { x: 8.45, y: 4.77, w: 4.15, h: 1.55, fontSize: 12 });
  s.addNotes("The sales-led team now creates its own pipeline and does add ARR, so the comparison is no longer true by construction — it just costs far more than it returns. Professional services is no better on the same engine.");
}

// ================================================================== 10. UNCERTAINTY
{
  const s = content("Correlated uncertainty doesn't rescue the case: P(NPV > 0) is ≈0% for every staffed option",
    "Source: 07_Financial_Model v2 › Monte Carlo — appendix/clay_engine.py (Python mirror reconciled to the workbook run by run); triangular(Downside, Base, Upside) inputs, Gaussian copula ρ = 0.5 within groups, 5,000 draws, seed 42; Gate 1 with observation noise [C/D].", "Prioritization & economics");
  const hist = D.hist.filter((h) => h[0] >= -6.0e6 && h[0] <= 0);
  const lab = hist.map((h) => (h[0] / 1e6).toFixed(1));
  s.addChart(pres.charts.BAR, [
    { name: "Sales-led team", labels: lab, values: hist.map((h) => h[2]) },
    { name: "Lean pilot, staged", labels: lab, values: hist.map((h) => h[1]) },
    { name: "Solution play", labels: lab, values: hist.map((h) => h[3]) },
  ], Object.assign({ x: ML, y: 1.45, w: 7.0, h: 4.95, barDir: "col", barGrouping: "clustered", chartColors: [HEX.mist, HEX.slate, HEX.copper], showTitle: true, title: "Simulated NPV, distribution — share of draws by $0.25M bin (NPV, $M)",
    showLegend: true, legendPos: "b", barGapWidthPct: 20, valAxisLabelFormatCode: "0%", catAxisLabelFontSize: 9, objectName: nm("chart") }, CHART_TXT, QUIET, { catAxisLabelFontSize: 9 }));
  const rows = [[th("Option (5,000 draws)", { align: "left" }), th("Mean NPV"), th("P90"), th("P(NPV > 0)")]];
  [["Industrial distribution", "staged", "Dist · lean pilot, staged"], ["Industrial distribution", "Sales-led", "Dist · sales-led team"], ["Industrial distribution", "Solution play", "Dist · solution play"],
   ["Professional services", "staged", "PS · lean pilot, staged"], ["Professional services", "Solution play", "PS · solution play"]].forEach((o) => {
    const r = mcRow(o[0], o[1]);
    rows.push([td(o[2]), tdc(M(r.mean)), tdc(M(r.p90)), tdc(P0(r.ppos))]);
  });
  s.addTable(rows, Object.assign({ x: 7.85, y: 1.5, w: 4.93, colW: [2.05, 0.98, 0.95, 0.95], rowH: 0.4, objectName: nm("table") }, TBL, { fontSize: 11 }));
  const rob = D.rob;
  const rr = [[th("Robustness (distribution lean pilot)", { align: "left" }), th("P(NPV > 0)")]].concat(
    [[0, "Base specification"], [2, "Near-comonotonic (≈ joint scenarios)"], [5, "Ranges 2× wider + strong correlation"]].map(([i, l]) => [td(l), tdc(P0(rob[i][6]))]));
  s.addTable(rr, Object.assign({ x: 7.85, y: 4.15, w: 4.93, colW: [3.68, 1.25], rowH: 0.4, objectName: nm("table") }, TBL, { fontSize: 11 }));
  T(s, "Upside requires every input to be favorable at once; v1 weighted that joint case at 20%.", { x: 7.85, y: 5.85, w: 4.93, h: 0.55, fontSize: 12, color: C.text2 });
  s.addNotes("The deterministic Upside is NPV-positive, but only when demand, ACV, retention and margin are all favorable together. With realistic correlation the probability of that is close to zero; even with ranges doubled it stays at about one in ten.");
}

// ================================================================== 11. RECOMMENDED PATH
pres.addSection({ title: "Recommendation & execution" });
{
  const s = content("Recommended path: validate in distribution, then at most a no-team solution play — never a vertical team up front",
    "Sources: 08_GTM_Strategy v2; 09_Implementation_Roadmap v2; 07_Financial_Model v2 (solution play, Base and Upside) [C/D/E]. Budget items are judgment estimates [D/E].", "Recommendation & execution");
  const stages = [
    ["A · Validation sprint", "Nov 2026 – Mar 2027", "≤ $0.2M", "4236 + 4238 · coverage, demand, conversion, committed usage, partners", "Gate 1: solution play or stop", "Default: stop"],
    ["B · Solution play", "Apr – Dec 2027 (only if Go)", "≈ $0.3M", "Templates, data providers, partner listing; existing AEs at a variable cost; no vertical lead", "Gate 2: continue only on the Upside track (≥ 40 opps, ≥ 7 customers, ACV ≥ $10K)", "Default: stop"],
    ["C · Lean pilot", "2028 at the earliest", "re-planned at Gate 2", "Only if the play is on track and borrowed AE capacity binds", "Re-run the model; fund only if NPV beats Option 0", "Default: fold into horizontal"],
  ];
  stages.forEach((st, i) => {
    const x = ML + i * 4.13, w = 3.97;
    rbox(s, x, 1.45, w, 2.95, i === 0 ? C.accent1 : C.background2, i === 0 ? 88 : 0);
    T(s, st[0], { x: x + 0.18, y: 1.55, w: w - 0.36, h: 0.32, fontSize: 15, bold: true, color: C.accent1 });
    T(s, [{ text: st[1] + "  ·  ", options: { bold: true } }, { text: "cap " + st[2] }], { x: x + 0.18, y: 1.92, w: w - 0.36, h: 0.3, fontSize: 12 });
    T(s, st[3], { x: x + 0.18, y: 2.3, w: w - 0.36, h: 0.85, fontSize: 12, color: C.text2 });
    T(s, st[4], { x: x + 0.18, y: 3.2, w: w - 0.36, h: 0.62, fontSize: 12, bold: true });
    T(s, st[5], { x: x + 0.18, y: 3.9, w: w - 0.36, h: 0.3, fontSize: 11, italic: true, color: C.accent1 });
  });
  T(s, "If the answer is Go — how to go to market", { x: ML, y: 4.6, w: 8, h: 0.32, fontSize: 14, bold: true });
  const gtm = [["ICP", "100–499 employees, multi-branch, 15+ territory sellers, CRM in place, selling to contractors and plants; SMB self-serve only"],
    ["Motion", "Clay-run outbound + templates + listed partners; existing AEs close; no dedicated AEs until capacity binds"],
    ["Pricing", "List price, Growth annual, no vertical discount; credit sizing by territory; promotional credits, not price cuts"],
    ["Not funded", "Vertical lead, dedicated AEs, native ERP connector, separate vertical marketing — before Gate 2"]];
  gtm.forEach((g, i) => {
    const x = ML + (i % 2) * 6.17, y = 5.0 + Math.floor(i / 2) * 0.78;
    T(s, g[0], { x, y, w: 1.2, h: 0.7, fontSize: 12, bold: true, color: C.accent1 });
    T(s, g[1], { x: x + 1.25, y, w: 4.8, h: 0.7, fontSize: 11, color: C.text2 });
  });
  s.addNotes("Each stage has a cap, a decision and a default of stopping. The GTM design survives from v1 — ICP, list pricing, partners, borrowed AEs — but it is only used if measured evidence supports it.");
}

// ================================================================== 12. SPRINT & GATE
{
  const s = content("The sprint tests five assumptions against thresholds set in advance — the default outcome is stop",
    "Sources: 09_Implementation_Roadmap v2; thresholds sit at the Upside values of 07_Financial_Model v2 [E]; owners are business functions, not Clay's actual org structure.", "Recommendation & execution");
  const rows = [[th("Test", { align: "left" }), th("What it measures", { align: "left" }), th("Base → Upside (model)"), th("Pass threshold"), th("Owner (function)", { align: "left" })],
    [td("T1 Coverage", { bold: true }), td("Verified decision-maker contacts on 200 contractor/plant accounts vs best single database"), tdc("—"), tdc("≥ 1.5×"), td("GTM eng., data partnerships")],
    [td("T2 Demand", { bold: true }), td("Clay-run outbound to ≥ 600 ICP accounts (300 per sub-industry)"), tdc("3.0% → 4.5%"), tdc("≥ 4.0% meetings"), td("Marketing ops")],
    [td("T3 Conversion", { bold: true }), td("Meetings that become qualified opportunities"), tdc("40% → 50%"), tdc("≥ 45%"), td("Sales leadership")],
    [td("T4 Value (mandatory)", { bold: true, color: HEX.copper }), td("Committed annual credits of ≥ 5 paid design partners at list"), tdc("$9.8K → $14.6K"), tdc("MM ACV ≥ $14.6K", { bold: true }), td("Deal desk")],
    [td("T5 Delivery", { bold: true }), td("Implementation partners committing named consultants"), tdc("2 → 3"), tdc("≥ 3"), td("Partnerships")]];
  s.addTable(rows, Object.assign({ x: ML, y: 1.45, w: CW, colW: [2.1, 4.85, 1.85, 1.6, 1.83], rowH: 0.5, objectName: nm("table") }, TBL));
  box(s, ML, 4.65, 6.0, 1.85, C.accent1, 91);
  T(s, [{ text: "Decision rule (end of March 2027)", options: { bold: true, color: C.accent1, breakLine: true } },
    { text: "Go to the solution play if T4 passes and at least three of T1, T2, T3, T5 pass. One more quarter (no extra cap) if T4 passes and only two others do. Otherwise stop: templates and data sources stay in the horizontal library.", options: {} }],
    { x: ML + 0.2, y: 4.77, w: 5.6, h: 1.65, fontSize: 12 });
  box(s, 6.78, 4.65, 6.0, 1.85, C.background2, 0);
  T(s, [{ text: "Why these tests  ", options: { bold: true, breakLine: true } },
    { text: "They are the inputs public data cannot pin down and the largest NPV swings in the model (usage, win and conversion rates, demand). Samples are small — 600 accounts give ~18 vs ~27 meetings at 3% vs 4.5% — so several tests must pass together.", options: { color: C.text2 } }],
    { x: 6.98, y: 4.77, w: 5.6, h: 1.65, fontSize: 12 });
  s.addNotes("The thresholds are the model's Upside values, because only the Upside case clears the hurdle. Committed usage is mandatory because nothing else can make the economics work without it. Win rate is deliberately not a gate test: five design partners cannot measure it.");
}

// ================================================================== APPENDIX
pres.addSection({ title: "Appendix" });
{
  const s = pres.addSlide({ masterName: "DIVIDER", sectionTitle: "Appendix" });
  s.addText("Appendix", { placeholder: "title" });
  s.addText([
    { text: "A1  What changed after the partner review (v1 → v2)", options: { breakLine: true } },
    { text: "A2  Hypothesis verdicts", options: { breakLine: true } },
    { text: "A3  How the market was sized: firms, not establishments", options: { breakLine: true } },
    { text: "A4  What moves NPV: sensitivity", options: { breakLine: true } },
    { text: "A5  Arguments for doing more — and what would change the recommendation", options: { breakLine: true } },
    { text: "A6  Key scenario assumptions", options: { breakLine: true } },
    { text: "A7  Limitations and principal sources" },
  ], { placeholder: "subtitle" });
}
{
  const s = content("A1 · A partner review found five critical flaws; fixing them reversed the recommendation",
    "Sources: review/Partner_Review_2026-10-19.md; Change Log sheets in 06 and 07 v2. v1 deck, workbooks and documents archived in appendix/archive/.", "Appendix");
  const rows = [[th("Issue", { align: "left" }), th("v1", { align: "left" }), th("v2", { align: "left" }), th("Effect", { align: "left" })],
    [td("C1 Sales-led comparison", { bold: true }), td("AEs could not create pipeline — result true by construction"), td("AEs prospect; program spend buys MQLs"), td("Sales-led adds ARR ($0.82M vs $0.51M) but NPV −$5.4M")],
    [td("C2 Win mix", { bold: true }), td("Fixed 50/38/12; enterprise accounts converting 5.4× SMB"), td("Derived from the account base"), td("ACV $13.8K → $8.8K")],
    [td("C3 Valuation", { bold: true }), td("Undiscounted; 20% weight on an all-favorable case; perfect gate"), td("NPV at 12%; incrementality; Monte Carlo; noisy gate"), td("P(NPV > 0) ≈ 0% for staffed options")],
    [td("C4 Alternatives", { bold: true }), td("Distribution vs zero only"), td("Prof. services, Option 0, solution play on one engine"), td("Every staffed option loses to Option 0")],
    [td("C5 Scorecard", { bold: true }), td("Seller density counted 3×, coverage proxy 2×"), td("One indicator per criterion"), td("Lead +0.56 → tie (−0.06)")],
    [td("M1/M2 Sizing", { bold: true }), td("Uniform ACV; CRM-ready 65%"), td("Usage-based ACV; CRM in use 56%"), td("TAM/SAM −30% to −45%")]];
  s.addTable(rows, Object.assign({ x: ML, y: 1.45, w: CW, colW: [2.2, 3.6, 3.2, 3.23], rowH: 0.62, objectName: nm("table") }, TBL));
}
{
  const s = content("A2 · Hypothesis verdicts after v2: the case for a vertical now fails on economics, not on need",
    "Sources: 01_Project_Charter (hypotheses set before research); verdicts from 03, 05, 06 and 07 v2. Evidence classes: A fact · B benchmark · C estimate · D assumption · E recommendation.", "Appendix");
  const V = (t, c) => td(t, { bold: true, color: c });
  const rows = [[th("ID"), th("Hypothesis", { align: "left" }), th("Verdict", { align: "left" }), th("Key evidence", { align: "left" })],
    [tdc("H1"), td("“B2B manufacturing” is too broad; narrow to core subsectors"), V("Supported", HEX.green), td("Core 4 subsectors ≈ 39% of 20+ manufacturers")],
    [tdc("H2"), td("Distribution's mid-market pool is ≥ 50% of professional services'"), V("Rejected", HEX.red), td("35% of professional services' 20–499 pool")],
    [tdc("H3"), td("Industrial buyer personas are poorly covered by LinkedIn-based data"), V("Supported (proxy)", HEX.green), td("26% vs 42% est. presence; sprint T1 to measure")],
    [tdc("H4"), td("Builder capacity, not need, is the binding constraint"), V("Supported", HEX.green), td("Computer occupations 0.69× average")],
    [tdc("H5"), td("Clay's edge is complementary, not substitutive"), V("Supported", HEX.green), td("Layers on CRM; no CRM/ERP of its own")],
    [tdc("H6"), td("Industrial ACV sits in the Growth-tier range (~$5–25K)"), V("Supported", HEX.green), td("$5.4K SMB · $9.7K mid-market (usage build)")],
    [tdc("H7"), td("Sales-led fails 24-month payback; partner-led can"), V("Rejected (2nd half)", HEX.red), td("Both fail: lean pilot NPV −$1.89M; sales-led −$5.44M")],
    [tdc("H8"), td("Even Upside is a small share of Clay ARR"), V("Supported", HEX.green), td("Upside exit ARR $3.0M ≈ 1.5% of est. $200M")],
    [tdc("H9"), td("Professional services has better risk-adjusted returns"), V("Tie on score; worse on economics", HEX.copper), td("Both 3.8, ranges overlap; NPV −$2.04M vs −$1.89M")],
    [tdc("H10"), td("Manufacturers are best reached through distribution"), V("Directional", HEX.copper), td("Manufacturing sales share 0.32× average")]];
  s.addTable(rows, Object.assign({ x: ML, y: 1.45, w: CW, colW: [0.6, 4.9, 2.5, 4.23], rowH: 0.45, objectName: nm("table") }, TBL, { fontSize: 11 }));
}
{
  const E = D.estab;
  const est20 = E["Establishments of firms with 20+ employees, 4235–4238 [C]"], uq20 = E["Unique firms with 20+ employees (Block 5)"];
  const estAll = E["Establishments, all sizes (FY), NAICS 4235–4238"], uqAll = E["Unique firms, all sizes [C]"];
  const s = content(`A3 · Sizing counts firms, not establishments — among distributors with 20+ employees, establishments would overstate buyers ${(est20 / uq20).toFixed(1)}×`,
    "Sources: Census 2022 Economic Census, Employment Size of Firms (S14) and Type of Operation (S58); 02_Market_Sizing v2 › Firm_Counts Blocks 3–7 (de-duplication, imputation, establishments, sales branches) and TAM_SAM_SOM [C/D]; 07_Financial_Model v2 Market Sizing.", "Appendix");
  s.addChart(pres.charts.BAR, [{ name: "Distribution buyer universe", labels: ["Establishments of 20+ employee firms", "Unique firms, 20+ employees", "SAM firms (4235–4238)", "Program SAM (4236–4238, excl. Clay customers)"], values: [Math.round(est20), Math.round(uq20), 5149, 4452] }],
    Object.assign({ x: ML, y: 1.45, w: 7.3, h: 4.95, barDir: "col", chartColors: [HEX.mist, HEX.slate, HEX.slate, HEX.copper], showTitle: true, title: "Industrial distribution (NAICS 4235–4238), firms with 20+ employees: from establishments to the program's market",
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0", valAxisHidden: true, showLegend: false, barGapWidthPct: 45, objectName: nm("chart") }, CHART_TXT, QUIET, NOGRID));
  box(s, 8.15, 1.45, 4.63, 4.95, C.background2, 0);
  T(s, "Method choices", { x: 8.35, y: 1.6, w: 4.2, h: 0.32, fontSize: 14, bold: true });
  T(s, [
    { text: `Like for like: both bars cover firms with 20+ employees, the sizing scope. All sizes: ${N0(estAll)} establishments vs ${N0(uqAll)} unique firms (${(estAll / uqAll).toFixed(2)}×)`, options: { bullet: true, breakLine: true } },
    { text: "A de-duplication factor (sector ÷ Σ sub-industries) removes firms counted in several sub-industries, by size band (~1–3% in distribution)", options: { bullet: true, breakLine: true } },
    { text: "Suppressed 4235 cells (475 firms; their establishments) recovered as residuals from published totals [C]", options: { bullet: true, breakLine: true } },
    { text: "SAM filters: active outbound motion (80%) × CRM in use (56%, 2026 survey)", options: { bullet: true, breakLine: true } },
    { text: "Manufacturers' sales branches are 4–22% of the 500+ band (Census type of operation); removing them changes no conclusion", options: { bullet: true } },
  ], { x: 8.35, y: 2.0, w: 4.25, h: 4.3, fontSize: 12, paraSpaceAfter: 6, color: C.text2 });
}
{
  const s = content("A4 · Usage intensity and win rate move NPV most — but no single input turns it positive",
    "Source: 07_Financial_Model v2 › Sensitivity — distribution lean pilot, Base; one input at a time between its Downside and Upside values (ICP tilt 1×↔3×, discount 15%↔8%) [C/D].", "Appendix");
  const short = (n) => n.replace(" (start, adds, cap, intros)", "").replace(" (meeting / intro / MQL → opp)", "").replace(" (15% ↔ 8%)", "").replace(" to Mid-Market / Enterprise (1× ↔ 3×)", "");
  const tor = D.tornado.slice().sort((a, b) => b[1] - a[1]).slice(0, 10).reverse();
  s.addChart(pres.charts.BAR, [{ name: "Swing", labels: tor.map((t) => short(t[0])), values: tor.map((t) => Math.round(t[1] / 1e3) / 1e3) }],
    Object.assign({ x: ML, y: 1.45, w: 7.0, h: 4.95, barDir: "bar", chartColors: [HEX.slate], showTitle: true, title: "Swing in NPV, low ↔ high value of one input ($M; Base = −$1.89M)",
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '"$"0.00"M"', valAxisHidden: true, showLegend: false, barGapWidthPct: 45, objectName: nm("chart") }, CHART_TXT, QUIET, NOGRID));
  box(s, 7.9, 1.45, 4.88, 4.95, C.background2, 0);
  T(s, "What would have to be true", { x: 8.1, y: 1.6, w: 4.5, h: 0.32, fontSize: 14, bold: true });
  T(s, [
    { text: "Even a 40% win rate with 3× the base usage (ACV ~$20K) leaves the lean pilot at −$0.44M and the solution play at +$0.16M — still below Option 0", options: { bullet: true, breakLine: true } },
    { text: "Usage alone would need ~6.4× the base build for the lean pilot to break even", options: { bullet: true, breakLine: true } },
    { text: "The Option 0 hurdle would have to fall to ~0.08× benchmark efficiency for the Base pilot to tie it", options: { bullet: true, breakLine: true } },
    { text: "So the sprint tests several inputs together; no single good number can carry the case", options: { bullet: true } },
  ], { x: 8.1, y: 2.0, w: 4.5, h: 4.3, fontSize: 12, paraSpaceAfter: 8, color: C.text2 });
}
{
  const s = content("A5 · Arguments for doing more — and what would change the recommendation",
    "Sources: Business Wire / Clay (S02); Clay case study (S07); 07_Financial_Model v2 Sensitivity and Monte Carlo; review/Partner_Review_2026-10-19.md.", "Appendix");
  const rows = [[th("Argument for a bigger bet", { align: "left" }), th("Why it does not change the call (yet)", { align: "left" }), th("What would change it", { align: "left" })],
    [td("Clay reports enterprise NRR > 200% (S02)"), td("Company-wide figure, driven by software and AI customers; the model already assumes 115% for enterprise distributors"), td("Design partners expanding usage ≥ 1.5× within two quarters")],
    [td("Regency Supply shows real distributor value (S07)"), td("One self-selected case study; its builder was unusually technical"), td("T1 coverage uplift and T4 committed usage across 5+ design partners")],
    [td("A vertical data moat (contractor, permit, project data) has strategic value beyond 2029"), td("Not priced in the model; also available through a data-led play at near-zero GTM cost"), td("Evidence that vertical data drives retention or pricing power in Clay's core")],
    [td("The model's negatives rest on assumed inputs"), td("True — which is exactly why the recommendation is a sprint, not a flat no"), td("Sprint results at the Upside thresholds")],
    [td("Option 0's efficiency (0.6×) may be too generous to the core"), td("The Base pilot ties Option 0 only at ~0.08×; even against zero it is NPV-negative"), td("Clay data showing marginal core spend is far below benchmark")]];
  s.addTable(rows, Object.assign({ x: ML, y: 1.45, w: CW, colW: [3.9, 4.6, 3.73], rowH: 0.78, objectName: nm("table") }, TBL));
}
{
  const s = content("A6 · Key scenario assumptions (distribution: Downside / Base / Upside)",
    "Source: 07_Financial_Model v2 › Assumptions — every input carries its evidence class, Monte Carlo group and source. Benchmarks: SaaS Capital 2025 (S43), Bridge Group 2026 (S44), HubSpot via S48, Digital Bloom (S49), partner margins (S50), CPL (S55).", "Appendix");
  const A = [["SAM filters (outbound × CRM in use) — fixed", "80% × 56%", "80% × 56%", "80% × 56%", "D/B", "Phase 1; Phocas/DSG 2026 (S25)"],
    ["Meeting rate (automated outbound)", "2.0%", "3.0%", "4.5%", "D", "Cold outbound 1–5% [K]"],
    ["Meeting → opportunity / win rate", "30% / 15%", "40% / 21%", "50% / 30%", "B/D", "SQL→opp 42% (S49); HubSpot avg win 21% (S48)"],
    ["Enterprise opportunity propensity (SMB = 1)", "0.5", "0.7", "0.9", "D", "Win mix derived, not assumed"],
    ["Usage intensity (sets ACV)", "0.75×", "1.0×", "1.5×", "D", "Usage build at list prices (S01)"],
    ["GRR / NRR, mid-market", "86% / 98%", "90% / 106%", "93% / 115%", "B", "SaaS Capital 2025 (S43)"],
    ["Gross margin", "55%", "65%", "72%", "D", "Data-heavy margins below 80% (S52)"],
    ["Incrementality haircut (wins that would happen anyway)", "30%", "20%", "10%", "D", "Prof. services 60% / 45% / 30%"],
    ["Discount rate / Option 0 marginal efficiency", "12% / 0.6×", "12% / 0.6×", "12% / 0.6×", "D", "Payback benchmark 14–18 months (S47)"],
    ["Cost per extra MQL (program spend above base)", "$600", "$400", "$250", "B/D", "SaaS CPL $200–350 (S55)"]];
  const rows = [[th("Input", { align: "left" }), th("Downside"), th("Base"), th("Upside"), th("Class"), th("Basis", { align: "left" })]]
    .concat(A.map((a) => [td(a[0]), tdc(a[1]), tdc(a[2], { bold: true }), tdc(a[3]), tdc(a[4]), td(a[5])]));
  s.addTable(rows, Object.assign({ x: ML, y: 1.45, w: CW, colW: [3.6, 1.5, 1.5, 1.5, 0.7, 3.43], rowH: 0.44, objectName: nm("table") }, TBL, { fontSize: 11 }));
}
{
  const s = content("A7 · Limitations and principal sources",
    "Full register (59 entries with dates, reliability and limitations): Research_Sources.csv. Workbooks, build scripts and the Python engine: 02, 04, 06, 07 and appendix/.", "Appendix");
  T(s, "Limitations", { x: ML, y: 1.45, w: 5.8, h: 0.35, fontSize: 16, bold: true });
  T(s, [
    { text: "No interviews, surveys or Clay internal data; win rates, margins, partner economics and Option 0 efficiency are benchmarks or assumptions", options: { bullet: true, breakLine: true } },
    { text: "Database-coverage gap is an education-based proxy; a 100-firm CSLB coverage test is designed and sampled but not yet run", options: { bullet: true, breakLine: true } },
    { text: "ACV is a list-price usage build; Clay pricing changed in March 2026 and may change again", options: { bullet: true, breakLine: true } },
    { text: "Monte Carlo results depend on the input ranges and correlation (robustness reported)", options: { bullet: true, breakLine: true } },
    { text: "Census firm counts are for 2022; manufacturers' sales branches (4–22% of 500+ distributors) are bounded, not removed", options: { bullet: true, breakLine: true } },
    { text: "Persona-vertical and data-led alternatives are named but not yet modeled", options: { bullet: true } },
  ], { x: ML, y: 1.9, w: 5.8, h: 4.6, fontSize: 13, paraSpaceAfter: 6, color: C.text2 });
  const src = [[th("ID"), th("Source", { align: "left" }), th("Date")],
    ["S01", "Clay pricing page (plans, credits, actions)", "Oct 2026"], ["S14", "U.S. Census Bureau, 2022 Economic Census — Employment Size of Firms", "Ref. 2022"],
    ["S15", "U.S. BLS, OEWS industry-specific estimates", "May 2025"], ["S16", "Vendr, Clay buyer pricing data", "Feb 2026"], ["S23", "Sacra, Apollo and Clay revenue estimates", "Sep 2026"],
    ["S25", "Phocas / DSG, State of Sales in Distribution", "2026"], ["S43", "SaaS Capital, B2B SaaS Retention Benchmarks", "Sep 2025"], ["S44", "Bridge Group, AE Models, Motions & Metrics", "Jun 2026"],
    ["S47", "SaaS Hero (Maxio et al.), CAC payback benchmarks", "Jul 2026"]]
    .map((r, i) => (i === 0 ? r : [tdc(r[0]), td(r[1]), tdc(r[2])]));
  s.addTable(src, Object.assign({ x: 6.75, y: 1.5, w: 6.03, colW: [0.6, 4.13, 1.3], rowH: 0.45, objectName: nm("table") }, TBL, { fontSize: 11 }));
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
