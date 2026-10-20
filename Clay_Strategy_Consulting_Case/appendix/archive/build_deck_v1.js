// Build Executive_Strategy_Deck.pptx (Phase 6) — Clay vertical expansion case study.
// Charts read model outputs from deckdata.json (dumped from 07_Financial_Model.xlsx) so slides match the model.
// Usage: python3 appendix/export_deck_data.py && node appendix/build_deck.js appendix/deck_data_from_model.json Executive_Strategy_Deck.pptx
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const { applyTheme } = require("/mnt/skills/public/pptx/scripts/apply_theme.js");

const DATA = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
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
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.author = "Zhuoran (Zora) Wang";
pres.company = "Independent case study";
pres.title = "U.S. Vertical Expansion and GTM Strategy for Clay — Executive Deck";
const C = pres.SchemeColor;
const ML = 0.55, CW = 12.23;

// ------------------------------------------------------------------ layouts
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
    { placeholder: { options: { name: "title", type: "title", x: ML, y: 0.95, w: 11.5, h: 1.5, fontFace: "Cambria", fontSize: 40, bold: true, color: C.text1, valign: "bottom", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "subtitle", type: "body", x: ML, y: 2.6, w: 11.5, h: 0.6, fontSize: 20, color: C.text2, valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "DIVIDER",
  background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: ML, y: 2.4, w: 11.5, h: 1.0, fontFace: "Cambria", fontSize: 40, bold: true, color: C.text1, valign: "bottom", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "subtitle", type: "body", x: ML, y: 3.6, w: 11.5, h: 2.4, fontSize: 16, color: C.text2, valign: "top", margin: 0 }, text: "" } },
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
function box(slide, x, y, w, h, fill, transparency, extra) {
  slide.addShape(pres.shapes.RECTANGLE, Object.assign({ x, y, w, h, fill: { color: fill, transparency: transparency || 0 }, line: { type: "none" }, objectName: nm("panel") }, extra || {}));
}
function rbox(slide, x, y, w, h, fill, transparency) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.06, fill: { color: fill, transparency: transparency || 0 }, line: { type: "none" }, objectName: nm("card") });
}
function stat(slide, x, y, w, big, label, color, bigSize, labelH) {
  T(slide, big, { x, y, w, h: 0.5, fontFace: "Cambria", fontSize: bigSize || 26, bold: true, color: color || C.accent1 });
  T(slide, label, { x, y: y + 0.52, w, h: labelH || 0.6, fontSize: 12, color: C.text2 });
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
const usdM = (v, d = 2) => (v < 0 ? "−" : "") + "$" + Math.abs(v / 1e6).toFixed(d) + "M";
const usdK = (v) => (v < 0 ? "−" : "") + "$" + Math.round(Math.abs(v) / 1000) + "K";
const CHART_TXT = { catAxisLabelColor: HEX.ink2, valAxisLabelColor: HEX.ink2, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", catAxisLabelFontSize: 11, valAxisLabelFontSize: 11, dataLabelFontFace: "+mn-lt", dataLabelFontSize: 11, dataLabelColor: HEX.ink, titleFontFace: "+mn-lt", titleFontSize: 12, titleColor: HEX.ink, legendFontFace: "+mn-lt", legendFontSize: 11, legendColor: HEX.ink2 };
const QUIET = { valGridLine: { color: HEX.grid, size: 0.75 }, catGridLine: { style: "none" }, catAxisLineShow: true, valAxisLineShow: false };
const TBL = { fontFace: "Arial", fontSize: 12, color: HEX.ink, border: { type: "solid", pt: 0.75, color: "D5DAE1" }, valign: "middle", margin: [3, 6, 3, 6] };
function th(text, opts) { return { text, options: Object.assign({ bold: true, color: "FFFFFF", fill: { color: HEX.ink2 }, align: "center" }, opts || {}) }; }
function td(text, opts) { return { text, options: Object.assign({}, opts || {}) }; }

// ================================================================== 1. TITLE
{
  pres.addSection({ title: "Decision" });
  const s = pres.addSlide({ masterName: "TITLE", sectionTitle: "Decision" });
  s.addText("Should Clay go deeper into U.S. industrial distribution?", { placeholder: "title" });
  s.addText("U.S. vertical expansion and go-to-market strategy, 2027–2029", { placeholder: "subtitle" });
  T(s, "The decision: where should Clay place incremental vertical GTM investment, and through which motion?", { x: ML, y: 3.55, w: 11.5, h: 0.4, fontSize: 16, bold: true });
  const opts = [
    ["Option 0", "No vertical bet — stay horizontal"],
    ["Option 1", "Industrial distribution — full sales-led commitment"],
    ["Option 2", "Manufacturing focus"],
    ["Option 3", "Professional services focus"],
    ["Option 4", "Staged, partner-led distribution pilot with Go/No-Go gates"],
  ];
  const cw = 2.3, gap = 0.18;
  opts.forEach((o, i) => {
    const x = ML + i * (cw + gap), rec = i === 4;
    rbox(s, x, 4.15, cw, 1.45, rec ? C.accent1 : C.background2, 0);
    T(s, o[0], { x: x + 0.15, y: 4.27, w: cw - 0.3, h: 0.3, fontSize: 12, bold: true, color: rec ? C.background1 : C.accent2 });
    T(s, o[1], { x: x + 0.15, y: 4.6, w: cw - 0.3, h: 0.9, fontSize: 14, bold: rec, color: rec ? C.background1 : C.text1 });
  });
  T(s, "Recommended", { x: ML + 4 * (cw + gap), y: 5.68, w: cw, h: 0.3, fontSize: 12, bold: true, color: C.accent1 });
  T(s, "Independent strategy case study · Zhuoran (Zora) Wang · October 2026", { x: ML, y: 6.3, w: 11.5, h: 0.3, fontSize: 14, color: C.text2 });
  T(s, "Not commissioned by, affiliated with, or endorsed by Clay Labs Inc. Built entirely from public data; no interviews or internal Clay data were used. All forward-looking figures are scenario estimates, not Clay forecasts.", { x: ML, y: 6.68, w: 11.5, h: 0.5, fontSize: 10, color: C.accent2 });
  s.addNotes("Frame the decision first. Five options were evaluated against an explicit baseline (Option 0). The recommendation is Option 4: a staged, partner-led pilot in distribution, with gates that decide whether it ever becomes Option 1.");
}

// ================================================================== 2. EXEC SUMMARY
{
  const s = content("Recommendation: test industrial distribution through a capped, gated pilot — not a vertical sales build",
    "Sources: 06_Segment_Prioritization; 07_Financial_Model (scenario estimates [C/D], not Clay forecasts); Census 2022 Economic Census (S14); BLS OEWS May 2025 (S15).", "Decision");
  const rows = [
    ["Distribution is the right place to test", "Weighted score 3.97 vs 3.41 (professional services) and 2.70 (manufacturing); first in 5 of 6 weight sets and 11 of 12 stress tests."],
    ["Clay has a real edge there", "Distributors' customers are hard to find in standard databases (est. 26% LinkedIn presence vs 42% for professional services), and 22% of distribution staff are in sales."],
    ["But the prize is small at list prices", "~5,200 serviceable firms; usage-based ACV ≈ $5.4K (SMB) and $9.8K (mid-market). Base case: $0.86M ARR by 2029, 37-month payback, −$2.7M cumulative."],
    ["So buy an option, not a program", "A dedicated sales team adds ~$4.5M of cost and no ARR. The lean pilot reaches $6.0M ARR and 12-month payback in Upside, and stops at a ~$1.0M loss in Downside."],
  ];
  rows.forEach((r, i) => {
    const y = 1.5 + i * 1.27;
    s.addShape(pres.shapes.OVAL, { x: ML, y, w: 0.46, h: 0.46, fill: { color: C.text1 }, line: { type: "none" }, objectName: nm("num") });
    T(s, String(i + 1), { x: ML, y, w: 0.46, h: 0.46, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle" });
    T(s, r[0], { x: ML + 0.7, y: y - 0.02, w: 7.2, h: 0.35, fontSize: 16, bold: true });
    T(s, r[1], { x: ML + 0.7, y: y + 0.36, w: 7.2, h: 0.8, fontSize: 14, color: C.text2 });
  });
  box(s, 8.75, 1.45, 4.03, 5.05, C.accent1, 91);
  T(s, "What we recommend", { x: 9.0, y: 1.62, w: 3.6, h: 0.35, fontSize: 16, bold: true, color: C.accent1 });
  const st = [["Jan 2027", "Gate 0: prove coverage and willingness to pay before any build spend"],
    ["$1.1M", "2027 budget cap, released quarterly"],
    ["~$1.0M", "maximum Downside loss, by stopping at the Q4 2027 gate"],
    ["$16K+ ACV", "and 32%+ win rate: the evidence needed to scale (Gate 1)"]];
  st.forEach((x, i) => stat(s, 9.0, 2.1 + i * 1.08, 3.6, x[0], x[1], C.accent1, 22, 0.5));
  s.addNotes("Answer first. Distribution wins the segment comparison and Clay has a genuine data-coverage edge there. The financial model shows the market is small at list prices, so the right commitment is an option with a hard stop, not a vertical build-out.");
}

pres.addSection({ title: "Market & customers" });
// ================================================================== 3. MARKET OPPORTUNITY
{
  const s = content("Distribution has the fewest firms but by far the most sellers per firm",
    "Sources: U.S. Census Bureau, 2022 Economic Census, Employment Size of Firms (S14) — firms, not establishments, de-duplicated [C]; BLS OEWS May 2025 (S15). Cores: NAICS 4235–4238 · 326/332/333/335 · 5413/5415/5416/5418.", "Market & customers");
  const cats = ["Distribution", "Manufacturing", "Prof. services"];
  const cols = [HEX.copper, HEX.slate, HEX.slate];
  const charts = [
    { t: "Unique U.S. firms, 20+ employees (2022)", v: [11494, 23785, 33035], f: "#,##0", c: cats, cc: cols },
    { t: "Sales jobs, % of employment (all industries 8.6%)", v: [22.0, 2.7, 6.2], f: '0.0"%"', c: cats, cc: cols },
    { t: "Sales staff per firm, 20+ employees", v: [27.7, 3.65, 9.32], f: "0.0", c: cats, cc: cols },
  ];
  charts.forEach((ch, i) => {
    s.addChart(pres.charts.BAR, [{ name: ch.t, labels: ch.c, values: ch.v }], Object.assign({
      x: ML + i * 4.15, y: 1.45, w: 3.95, h: 3.75, barDir: "col", chartColors: ch.cc, showTitle: true, title: ch.t,
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: ch.f, valAxisHidden: true, showLegend: false, barGapWidthPct: 60,
      objectName: nm("chart"),
    }, CHART_TXT, QUIET, { valGridLine: { style: "none" } }));
  });
  box(s, ML, 5.45, CW, 1.0, C.accent1, 91);
  T(s, [{ text: "So what  ", options: { bold: true, color: C.accent1 } },
    { text: "Distribution trades market breadth for need intensity: an in-scope distributor employs ~28 sales staff, ~3× a professional-services firm and ~8× a manufacturer. Tools priced per seat charge for every one of them; Clay's plans do not. Hypothesis H2 (distribution has the larger mid-market pool) is rejected." }],
    { x: ML + 0.25, y: 5.57, w: CW - 0.5, h: 0.8, fontSize: 14 });
  s.addNotes("Distribution is the smallest pool by firm count, but its employment is dominated by sellers. That density is what makes a seat-free data and workflow tool attractive. Counts are firms, not establishments; establishments would overstate distribution by about two-thirds.");
}

// ================================================================== 4. TAM SAM SOM
{
  const s = content("Sized bottom-up, each vertical is an incremental option worth ~1–2% of Clay's ARR — not a pivot",
    "Sources: Census 2022 (S14); Clay pricing (S01); Vendr (S16); DSG 2022 CRM survey (S17). Clay ARR ~$200M is a Sacra estimate (S23), not company-confirmed. Uniform ACV and penetration across segments by design [C/D].", "Market & customers");
  s.addChart(pres.charts.BAR, [
    { name: "TAM", labels: ["Distribution", "Manufacturing", "Prof. services"], values: [131.1, 280.9, 385.0] },
    { name: "SAM", labels: ["Distribution", "Manufacturing", "Prof. services"], values: [68.2, 75.8, 148.2] },
  ], Object.assign({ x: ML, y: 1.45, w: 6.9, h: 4.2, barDir: "col", barGrouping: "clustered", chartColors: [HEX.mist, HEX.slate], showTitle: true, title: "TAM and SAM by segment, Base case ($M per year)",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '"$"0', valAxisHidden: true, showLegend: true, legendPos: "b", barGapWidthPct: 70, objectName: nm("chart") }, CHART_TXT, QUIET, { valGridLine: { style: "none" } }));
  const rows = [[th("Segment", { align: "left" }), th("SOM: year-3 ARR"), th("% of Clay est. ARR")],
    [td("Distribution", { bold: true }), td("$1.9M", { align: "center" }), td("0.9%", { align: "center" })],
    [td("Manufacturing"), td("$2.1M", { align: "center" }), td("1.1%", { align: "center" })],
    [td("Professional services"), td("$4.2M", { align: "center" }), td("2.1%", { align: "center" })]];
  s.addTable(rows, Object.assign({ x: 7.8, y: 1.5, w: 4.98, colW: [2.1, 1.44, 1.44], rowH: 0.38, objectName: nm("table") }, TBL));
  T(s, "Even the compounded High case stays at or below ~12% of Clay's estimated ARR.", { x: 7.8, y: 3.15, w: 4.98, h: 0.5, fontSize: 12, color: C.text2 });
  box(s, 7.8, 3.75, 4.98, 1.9, C.background2, 0);
  T(s, [
    { text: "TAM  ", options: { bold: true } }, { text: "unique firms with 20+ employees × size-band ACV", options: { breakLine: true } },
    { text: "SAM  ", options: { bold: true } }, { text: "TAM firms with an outbound/territory sales motion and a CRM Clay can integrate today", options: { breakLine: true } },
    { text: "SOM  ", options: { bold: true } }, { text: "year-3 ARR at uniform fair-share penetration (1.5–6% of SAM firms by size band)" },
  ], { x: 7.98, y: 3.85, w: 4.65, h: 1.75, fontSize: 12, paraSpaceAfter: 4 });
  box(s, ML, 5.85, CW, 0.6, C.accent1, 91);
  T(s, [{ text: "Refined in Phase 4  ", options: { bold: true, color: C.accent1 } }, { text: "The pilot's serviceable market (NAICS 4236–4238, excluding metals and existing Clay customers) is ~5,170 firms." }],
    { x: ML + 0.25, y: 5.98, w: CW - 0.5, h: 0.4, fontSize: 14 });
  s.addNotes("Professional services is the largest pool on size alone, which is why size cannot decide the question. Every segment's three-year ARR potential is one to two percent of Clay's estimated ARR, so this is an option-value decision, consistent with H8.");
}

// ================================================================== 5. SEGMENTATION
{
  const s = content("Clay's edge is largest where adoption is hardest: distributors' customers are least visible online and their teams least technical",
    "Sources: BLS OEWS May 2025 (S15); persona coverage = BLS education mix (S29) × Pew LinkedIn use by education (S28), an indirect proxy [C]; scores (1–5) from 04_Competitive_Analysis [C].", "Market & customers");
  const G = { fill: { color: "E2EEE9" } }, R = { fill: { color: "F5E4E4" } }, N = {};
  const c = (t, o, b) => td(t, Object.assign({ align: "center", bold: !!b }, o));
  const rows = [
    [th("Indicator", { align: "left" }), th("Distribution"), th("Manufacturing"), th("Prof. services")],
    [td("Need: sales occupations, % of employment (all industries 8.6%)"), c("22.0%", G, 1), c("2.7%", R), c("6.2%", N)],
    [td("Data gap: est. LinkedIn presence of customers' personas (lower = harder for databases)"), c("26% (contractors 20%)", G, 1), c("41%", N), c("42%", N)],
    [td("Competitive position (1–5)"), c("4.0", G, 1), c("3.3", N), c("2.3", R)],
    [td("Builder capacity: computer occupations, % of employment (all 3.4%)"), c("2.3%", R), c("1.5%", R), c("24.2%", G, 1)],
    [td("Adoption / execution fit (1–5)"), c("2.0", R), c("1.5", R), c("4.5", G, 1)],
    [td("Incremental headroom for Clay (1–5)"), c("5", G), c("5", G), c("2", R)],
  ];
  s.addTable(rows, Object.assign({ x: ML, y: 1.5, w: 8.45, colW: [3.95, 1.6, 1.45, 1.45], rowH: 0.62, objectName: nm("table") }, TBL));
  T(s, "Green = favorable for Clay · red = unfavorable", { x: ML, y: 6.02, w: 8.45, h: 0.3, fontSize: 11, color: C.text2 });
  box(s, 9.3, 1.5, 3.48, 4.85, C.background2, 0);
  const items = [["Fit: distribution", "Highest need, biggest data gap and strongest competitive position (H3 supported)."],
    ["Ease: professional services", "Technical buyers who can self-serve, but already a Clay home market with little headroom."],
    ["Manufacturing", "Fails both tests; reach it through the distributors that sell for it (H10)."]];
  items.forEach((it, i) => {
    T(s, it[0], { x: 9.5, y: 1.7 + i * 1.5, w: 3.1, h: 0.32, fontSize: 14, bold: true, color: i === 0 ? C.accent1 : C.text1 });
    T(s, it[1], { x: 9.5, y: 2.05 + i * 1.5, w: 3.1, h: 1.05, fontSize: 13, color: C.text2 });
  });
  s.addNotes("The core tension of the case: the segment where Clay is most differentiated is also the hardest to adopt in. That is why the prioritization had to be stress-tested on the weight given to ease of execution, and why delivery partners sit at the center of the GTM motion.");
}

pres.addSection({ title: "Competition & value proposition" });
// ================================================================== 6. COMPETITION
{
  const s = content("Only Clay combines off-LinkedIn data with workflow automation — but it is the hardest option to implement",
    "Sources: 04_Competitive_Analysis — scores are analyst judgments from public vendor documentation [C]; some capabilities rated from general knowledge [K] and pending verification. Phocas/DSG 2026 (S25); Apollo pricing (S20); Sacra (S23).", "Competition & value proposition");
  const heat = { 5: HEX.slate, 4: "8C99AB", 3: "BEC6D1", 2: "DDE2E8", 1: "F0F2F5" };
  const sc = (v) => td(String(v), { align: "center", bold: true, fill: { color: heat[v] }, color: v >= 5 ? "FFFFFF" : HEX.ink });
  const data = [
    ["Data coverage of industrial buyers", 5, 2, 2, 2, 4],
    ["Workflow automation", 5, 3, 3, 4, 1],
    ["Differentiation for industrial sellers", 5, 3, 3, 2, 4],
    ["Pricing architecture", 4, 3, 2, 3, 3],
    ["Ease of implementation (5 = easiest)", 2, 5, 3, 5, 4],
  ];
  const rows = [[th("Dimension (score 1–5)", { align: "left" }), th("Clay", { fill: { color: HEX.copper } }), th("Apollo"), th("ZoomInfo"), th("HubSpot (Breeze)"), th("Vertical data (Dodge, Shovels)")]]
    .concat(data.map((d) => [td(d[0])].concat(d.slice(1).map(sc))));
  s.addTable(rows, Object.assign({ x: ML, y: 1.5, w: 8.3, colW: [2.9, 0.95, 0.95, 1.05, 1.15, 1.3], rowH: 0.7, objectName: nm("table") }, TBL));
  T(s, "Clay leads on capability and industrial differentiation, and trails on ease of implementation — the gap partners and templates must close.", { x: ML, y: 5.85, w: 8.3, h: 0.6, fontSize: 13, color: C.text2 });
  box(s, 9.15, 1.5, 3.63, 4.95, C.background2, 0);
  stat(s, 9.35, 1.65, 3.25, "68%", "of distributors rely on spreadsheets: the status quo is the main competitor (S25)", C.accent1, 26, 0.75);
  stat(s, 9.35, 3.2, 3.25, "2.7×", "Apollo Professional's list cost vs Clay Growth for a 28-user team; 0.87× at 9 users [C]", C.accent1, 26, 0.75);
  stat(s, 9.35, 4.75, 3.25, "6 of 11", "Clay use cases Apollo now also covers; it still lacks off-LinkedIn contractor data [C]", C.accent1, 26, 0.75);
  s.addNotes("In distribution, the realistic alternative is spreadsheets plus a generic database. Clay is the only option that combines broad data with automation, and its unlimited-seat pricing beats per-seat tools for seller-dense teams. Its weakness is implementation, which shapes the motion.");
}

// ================================================================== 7. NEEDS & VALUE PROP
{
  const s = content("Lead with coverage: “the contractors and plants your database can't see”",
    "Sources: Phocas/DSG State of Sales in Distribution 2026 (S25; industry-sponsored, ~100+ respondents); BLS (S15, S29); Pew (S28); Regency Supply case study (S07). No customer interviews were conducted — needs are inferred.", "Competition & value proposition");
  T(s, "What distributors need", { x: ML, y: 1.45, w: 5.3, h: 0.35, fontSize: 16, bold: true });
  const needs = [["39%", "name new-customer revenue as their top 12-month priority (S25)"],
    ["34%", "of CRM users log at least 75% of customer interactions — CRM data is thin (S25)"],
    ["~20%", "est. LinkedIn presence of contractor personas, vs ~42% for professional-services buyers [C]"],
    ["2.3%", "of distribution employment is in computer occupations — few in-house builders (S15)"]];
  needs.forEach((n, i) => {
    const y = 1.95 + i * 1.12;
    T(s, n[0], { x: ML, y, w: 1.3, h: 0.5, fontFace: "Cambria", fontSize: 26, bold: true, color: C.accent1 });
    T(s, n[1], { x: ML + 1.4, y: y + 0.04, w: 4.0, h: 0.9, fontSize: 13, color: C.text2 });
  });
  box(s, 6.45, 1.45, 6.33, 1.35, C.accent1, 91);
  T(s, "“Find every contractor and plant in your territories — including the ones ZoomInfo and Apollo can't see — and put them in your reps' CRM, scored by branch. No data team, and no per-seat fees.”",
    { x: 6.65, y: 1.57, w: 5.95, h: 1.15, fontSize: 14, italic: true, color: C.text1 });
  const pillars = [["Coverage", "Maps sourcing, permits, multi-provider waterfall. Proof: Regency Supply (S07)"],
    ["Prioritization", "Branch scoring and signals. Proof to create: meeting-rate lift at design partners"],
    ["CRM completeness", "Auto-sync and job-change tracking. Proof: 5,000+ contacts monitored at Regency"],
    ["Economics", "Unlimited seats: Apollo costs ~2.7× more for a 28-user team [C]"]];
  pillars.forEach((p, i) => {
    const x = 6.45 + (i % 2) * 3.22, y = 3.0 + Math.floor(i / 2) * 1.42;
    rbox(s, x, y, 3.11, 1.3, C.background2, 0);
    T(s, p[0], { x: x + 0.15, y: y + 0.12, w: 2.8, h: 0.3, fontSize: 14, bold: true, color: C.accent1 });
    T(s, p[1], { x: x + 0.15, y: y + 0.45, w: 2.8, h: 0.8, fontSize: 12, color: C.text2 });
  });
  T(s, [{ text: "Not claimed: ", options: { bold: true } }, { text: "AI cross-sell from order history (no ERP data), replacing the CRM or ERP, or efficiency gains before pilots measure them." }],
    { x: 6.45, y: 5.95, w: 6.33, h: 0.55, fontSize: 12, color: C.text2 });
  s.addNotes("The value proposition leads with what generic databases cannot do, and avoids the crowded AI-copilot story where Proton.ai and ERP vendors own the order-history data. Each pillar lists the proof available today and the proof the pilot must create.");
}

pres.addSection({ title: "Prioritization & strategy" });
// ================================================================== 8. PRIORITIZATION
{
  const s = content("Distribution ranks first on the weighted scorecard and holds under most weightings and stress tests",
    "Source: 06_Segment_Prioritization — scores 1–5 with ratio-to-max scaling; weights as suggested in the brief; contribution = score × weight [C]. Random-weight test: Dirichlet draws (seed 42).", "Prioritization & strategy");
  const crit = [["Market attractiveness (20%)", [4.02, 3.50, 3.92], 0.20], ["Customer need (20%)", [2.53, 2.81, 5.00], 0.20], ["Product-market fit (20%)", [2, 4, 4], 0.20],
    ["Revenue potential (15%)", [2.29, 3.67, 3.89], 0.15], ["Competitive position (15%)", [3.33, 2.33, 4.00], 0.15], ["Ease of execution (10%)", [1.5, 4.5, 2.0], 0.10]];
  const labels = ["Manufacturing · 2.70", "Prof. services · 3.41", "Distribution · 3.97"];
  const series = crit.map((c) => ({ name: c[0], labels, values: c[1].map((v) => Math.round(v * c[2] * 100) / 100) }));
  s.addChart(pres.charts.BAR, series, Object.assign({ x: ML, y: 1.45, w: 7.7, h: 4.95, barDir: "bar", barGrouping: "stacked", chartColors: [HEX.ink, HEX.slate, HEX.mist, HEX.copper, "D9A48A", "7E9C8F"],
    showTitle: true, title: "Weighted score by criterion (maximum 5.0)", showLegend: true, legendPos: "b", valAxisMaxVal: 5, valAxisMinVal: 0, valAxisMajorUnit: 1, barGapWidthPct: 55, objectName: nm("chart") }, CHART_TXT, QUIET));
  box(s, 8.6, 1.45, 4.18, 4.95, C.background2, 0);
  T(s, "Robustness checks", { x: 8.82, y: 1.6, w: 3.8, h: 0.35, fontSize: 16, bold: true });
  const r = [["5 of 6", "named weight sets rank distribution first (it loses only “execution first”)"],
    ["74% / 99%", "of random weightings rank it first (flat / centered Dirichlet)"],
    ["11 of 12", "stress tests; it loses only when its worst case meets professional services' best case (T9)"],
    ["26%", "weight on ease of execution needed to flip the ranking (base 10%)"]];
  r.forEach((x, i) => stat(s, 8.82, 2.08 + i * 1.07, 3.8, x[0], x[1], C.accent1, 20, 0.5));
  s.addNotes("Distribution wins on customer need, competitive position and revenue potential, and loses on ease of execution. The ranking flips only when ease carries more than about a quarter of the weight, or when every distribution downside coincides with every professional-services upside.");
}

// ================================================================== 9. GTM STRATEGY
{
  const s = content("Enter through a lean hybrid motion: partners deliver, templates simplify, existing AEs close",
    "Sources: 08_GTM_Strategy [E]; 07_Financial_Model, Base case [C/D]; NAED (S53); American Supply Association (S54); partner-economics benchmarks (S50). Serviceable firms = Census 2022 firms × SAM filters [C/D].", "Prioritization & strategy");
  const steps = [["1  Create demand", "Clay-run outbound to ~40% of each wave's market per quarter; NAED and ASA channels; partner referrals"],
    ["2  Convert", "Existing AEs close at ~$7.1K per deal; no dedicated AEs until capacity binds"],
    ["3  Deliver", "Certified partners run a fixed-scope Quickstart using the Distribution Territory Kit"],
    ["4  Expand", "Add workflows (permits, e-commerce vetting, job changes): usage is the ACV lever"]];
  steps.forEach((st, i) => {
    const x = ML + i * 3.08;
    s.addShape(pres.shapes.CHEVRON, { x, y: 1.45, w: 2.98, h: 0.5, fill: { color: i === 2 ? C.accent1 : C.text2 }, line: { type: "none" }, objectName: nm("step") });
    T(s, st[0], { x: x + 0.3, y: 1.45, w: 2.4, h: 0.5, fontSize: 14, bold: true, color: C.background1, valign: "middle" });
    T(s, st[1], { x: x + 0.05, y: 2.05, w: 2.85, h: 1.05, fontSize: 12, color: C.text2 });
  });
  // ICP card
  rbox(s, ML, 3.3, 3.9, 3.2, C.background2, 0);
  T(s, "Ideal customer profile", { x: ML + 0.18, y: 3.42, w: 3.5, h: 0.32, fontSize: 14, bold: true, color: C.accent1 });
  T(s, [
    { text: "100–499 employees, multi-branch", options: { bullet: true, breakLine: true } },
    { text: "15+ territory sellers", options: { bullet: true, breakLine: true } },
    { text: "CRM in place with a named owner", options: { bullet: true, breakLine: true } },
    { text: "Sells to contractors, plants, local firms", options: { bullet: true, breakLine: true } },
    { text: "New-customer growth is a priority", options: { bullet: true } },
  ], { x: ML + 0.18, y: 3.82, w: 3.55, h: 1.75, fontSize: 12, paraSpaceAfter: 3 });
  T(s, "Avoid: no CRM, under ~5 sellers, or a need for ERP-based cross-sell", { x: ML + 0.18, y: 5.68, w: 3.55, h: 0.7, fontSize: 11, italic: true, color: C.text2 });
  // waves
  T(s, "Sequenced entry (Base case)", { x: 4.65, y: 3.35, w: 4.0, h: 0.32, fontSize: 14, bold: true });
  const wrows = [[th("Wave", { align: "left" }), th("Launch"), th("Firms")],
    [td("4236 Electrical", { bold: true }), td("Q1 2027", { align: "center" }), td("~1,280", { align: "center" })],
    [td("4237 Plumbing / HVAC"), td("Q1 2028", { align: "center" }), td("~900", { align: "center" })],
    [td("4238 Machinery / MRO"), td("Q3 2028", { align: "center" }), td("~2,980", { align: "center" })]];
  s.addTable(wrows, Object.assign({ x: 4.65, y: 3.75, w: 4.0, colW: [1.9, 1.05, 1.05], rowH: 0.4, objectName: nm("table") }, TBL));
  T(s, "Price at list, Growth annual, no vertical discount; size credits by territory at signature.", { x: 4.65, y: 5.5, w: 4.0, h: 0.9, fontSize: 12, color: C.text2 });
  // motion verdicts
  T(s, "Motion options", { x: 8.95, y: 3.35, w: 3.8, h: 0.32, fontSize: 14, bold: true });
  const mv = [["Sales-led team", "Reject: +$4.5M cost, same ARR", C.accent6], ["Product-led only", "Entry path only: few builders", C.accent5],
    ["Partner-led only", "Partners must be recruited; can't close enterprise", C.accent5], ["Hybrid", "Recommended", C.accent4]];
  mv.forEach((m, i) => {
    const y = 3.78 + i * 0.66;
    s.addShape(pres.shapes.OVAL, { x: 8.95, y: y + 0.06, w: 0.2, h: 0.2, fill: { color: m[2] }, line: { type: "none" }, objectName: nm("dot") });
    T(s, m[0], { x: 9.3, y, w: 3.45, h: 0.28, fontSize: 13, bold: true });
    T(s, m[1], { x: 9.3, y: y + 0.28, w: 3.45, h: 0.32, fontSize: 11, color: C.text2 });
  });
  s.addNotes("Each element answers a constraint found earlier: Clay-run outbound answers thin demand, borrowed AEs answer the small market, partners and templates answer low builder capacity, and the ICP tilt to mid-market answers the low SMB ACV.");
}

// ================================================================== 10. FINANCIAL CASE
{
  const s = content("A dedicated sales team adds ~$4.5M of cost and no ARR; the lean pilot pays back only in the Upside case",
    "Source: 07_Financial_Model.xlsx — quarterly funnel model, Q1 2027–Q4 2029; usage-based ACV from Clay list prices (S01); benchmarks S43–S50. Payback = CAC ÷ monthly gross profit per new customer. Scenario estimates [C/D].", "Prioritization & strategy");
  const q = ["Q1 27", "Q2 27", "Q3 27", "Q4 27", "Q1 28", "Q2 28", "Q3 28", "Q4 28", "Q1 29", "Q2 29", "Q3 29", "Q4 29"];
  const m = (a) => a.map((v) => Math.round(v / 1e4) / 100);
  s.addChart(pres.charts.LINE, [
    { name: "Upside + 2 AEs after gate", labels: q, values: m(DATA.arr["Partner-led: Upside + scale-up"]) },
    { name: "Base (sales-led team: identical)", labels: q, values: m(DATA.arr["Partner-led: Base"]) },
    { name: "Downside (ungated)", labels: q, values: m(DATA.arr["Partner-led: Downside"]) },
  ], Object.assign({ x: ML, y: 1.45, w: 5.9, h: 4.6, chartColors: [HEX.copper, HEX.ink, HEX.mist], lineSize: 2.5, lineDataSymbol: "none", showTitle: true, title: "Program ARR by quarter ($M)",
    showLegend: true, legendPos: "b", valAxisLabelFormatCode: '"$"0.0"M"', objectName: nm("chart") }, CHART_TXT, QUIET));
  chip(s, ML, 6.15, 4.2, "Illustrative scenario outputs [C/D] — not Clay forecasts", C.accent2);
  const S = DATA.scen;
  const col = [1, 2, 4, 5]; // Downside gated, Base, Upside+scale, Sales-led Base
  const fmt = {
    "Exit ARR, Q4 2029": (v) => (Math.abs(v) < 1e5 ? usdK(v) : usdM(v, 2)),
    "Customers, Q4 2029": (v) => (v < 1 ? "<1" : String(Math.round(v))),
    "Average first-year ACV": (v) => "$" + (v / 1000).toFixed(1) + "K",
    "CAC": (v, i) => (i === 0 ? "n/m" : usdK(v)),
    "CAC payback (2029 cohort)": (v, i) => (i === 0 ? "n/m" : Math.round(v) + " mo"),
    "Cumulative program P&L 2027–29": (v) => usdM(v, 2),
    "Peak investment (lowest cumulative P&L)": (v) => usdM(v, 2),
  };
  const lab = { "Exit ARR, Q4 2029": "Exit ARR, Q4 2029", "Customers, Q4 2029": "Customers, Q4 2029", "Average first-year ACV": "Avg first-year ACV", "CAC": "CAC",
    "CAC payback (2029 cohort)": "CAC payback, 2029 cohort", "Cumulative program P&L 2027–29": "Cumulative P&L 2027–29", "Peak investment (lowest cumulative P&L)": "Peak investment" };
  const rows = [[th("Metric", { align: "left" }), th("Downside, stopped Q4 27"), th("Base"), th("Upside + 2 AEs", { fill: { color: HEX.copper } }), th("Sales-led team, Base")]];
  Object.keys(fmt).forEach((k) => {
    rows.push([td(lab[k], { bold: k.startsWith("Exit") || k.startsWith("Cumulative") })].concat(col.map((c, i) => td(fmt[k](S[k][c], i), { align: "center", bold: k.startsWith("Exit") || k.startsWith("Cumulative") }))));
  });
  s.addTable(rows, Object.assign({ x: 6.75, y: 1.5, w: 6.03, colW: [1.95, 1.02, 0.94, 1.06, 1.06], rowH: 0.47, objectName: nm("table") }, TBL, { fontSize: 11 }));
  box(s, 6.75, 5.45, 6.03, 1.0, C.accent1, 91);
  T(s, [{ text: "Why  ", options: { bold: true, color: C.accent1 } }, { text: "Demand, not sales capacity, is the constraint: extra AEs win nothing in Base. Mid-market ACV of ~$10K at list keeps Base payback near three years." }],
    { x: 6.95, y: 5.55, w: 5.65, h: 0.85, fontSize: 13 });
  s.addNotes("This slide carries the key surprise of the analysis. The sales-led build produces exactly the same ARR as the lean pilot in Base, because pipeline, not closing capacity, binds. The lean pilot still does not pay back in Base within three years; it pays back in roughly a year only in Upside.");
}

pres.addSection({ title: "Risks & execution" });
// ================================================================== 11. RISKS & SCENARIOS
{
  const s = content("ACV per account is the swing factor — the pilot is worth running only with a hard stop",
    "Source: 07_Financial_Model (Sensitivity: one input at a time between its Downside and Upside values; Scenarios: decision analysis) [C/D]; 09_Implementation_Roadmap risk register. Total value = cumulative P&L + 5 years of gross profit from the 2029 base.", "Risks & execution");
  const short = { "Usage intensity (ACV)": "Usage intensity (ACV)", "Win rate": "Win rate", "Gross margin": "Gross margin", "SQL → opportunity": "SQL → opportunity", "Qualification rates": "Qualification rates",
    "Partner network (start, adds, cap, leads/partner)": "Partner network", "Wave timing (waves 2 and 3)": "Wave timing", "Outbound meeting rate": "Outbound meeting rate", "Inbound / template lead rate": "Inbound lead rate", "Net revenue retention": "Net revenue retention" };
  const tor = DATA.tornado.filter((t) => short[t[0]]).sort((a, b) => a[1] - b[1]);
  s.addChart(pres.charts.BAR, [{ name: "Swing", labels: tor.map((t) => short[t[0]]), values: tor.map((t) => Math.round(t[1] / 1e4) / 100) }],
    Object.assign({ x: ML, y: 1.45, w: 6.3, h: 5.05, barDir: "bar", chartColors: [HEX.slate], showTitle: true, title: "Swing in total value, Downside → Upside value of one input ($M; Base = −$1.04M)",
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '"$"0.00"M"', valAxisHidden: true, showLegend: false, barGapWidthPct: 45, objectName: nm("chart") }, CHART_TXT, QUIET, { valGridLine: { style: "none" } }));
  const D = DATA.dec;
  const ev = D["Total value — staged (Downside stopped at Q4 2027 gate; Upside adds 2 AEs)"][3];
  const be = D["Break-even probability of Upside for staged EV = 0 (Downside:Base ratio held)"][3];
  const cash = D["Cumulative P&L 2027–29 only — staged (excludes post-2029 value)"][3];
  box(s, 7.15, 1.45, 5.63, 2.45, C.background2, 0);
  T(s, "Decision analysis (staged pilot)", { x: 7.35, y: 1.58, w: 5.2, h: 0.32, fontSize: 14, bold: true });
  stat(s, 7.35, 2.0, 1.75, (ev >= 0 ? "+" : "") + usdM(ev, 1), "expected total value", C.accent4, 22, 0.5);
  stat(s, 9.2, 2.0, 1.75, Math.round(be * 100) + "%", "break-even probability of Upside", C.accent1, 22, 0.5);
  stat(s, 11.05, 2.0, 1.65, usdM(cash, 1), "expected cash P&L, 2027–29", C.accent6, 22, 0.5);
  T(s, "Weights 35 / 45 / 20 (Downside / Base / Upside) are judgments [D]. Every scenario is cash-negative through 2029; value rests on the post-2029 customer base and on stopping early.",
    { x: 7.35, y: 3.1, w: 5.25, h: 0.75, fontSize: 11, color: C.text2 });
  const rr = [[th("Top risk", { align: "left" }), th("Early signal", { align: "left" }), th("Response", { align: "left" })],
    [td("ACV stuck at the plan floor"), td("Avg ACV < $12K"), td("Mid-market ICP, credit sizing, multi-workflow kit")],
    [td("Demand thinner than Base"), td("< 15 opps by Q4 2027"), td("Gate 1 stop; loss capped ~$1.0M")],
    [td("Too few productive partners"), td("< 4 producing partners"), td("Recruit ERP/CRM implementers")],
    [td("Coverage edge smaller than proxy"), td("Match uplift < 1.5×"), td("Gate 0 No-Go")]];
  s.addTable(rr, Object.assign({ x: 7.15, y: 4.1, w: 5.63, colW: [1.95, 1.55, 2.13], rowH: 0.47, objectName: nm("table") }, TBL, { fontSize: 11 }));
  s.addNotes("Usage intensity, which sets ACV, moves total value more than any funnel rate, so the pilot must prove ACV first. With explicit probability weights the staged pilot is positive in expectation, but only because the gate truncates the Downside and the Upside is large. If the gate cannot separate Downside from Base, the case weakens sharply.");
}

// ================================================================== 12. ROADMAP
{
  const s = content("Four gates cap the downside at ~$1.0M and tie every scale-up step to evidence",
    "Source: 09_Implementation_Roadmap — gate thresholds derived from the 07 scenario runs [E]; owners are business functions, not Clay's actual org structure.", "Risks & execution");
  const phases = [["Q4 2026", "Validate", "staff time only"], ["2027", "Pilot wave 1: electrical", "budget cap $1.1M"], ["2028", "Extend: PHVAC, MRO", "AEs only if capacity binds"], ["2029", "Scale or harvest", "per Gate 2 outcome"]];
  const gates = [["Gate 0 · Jan 2027", "Launch?", "Match rate ≥ 1.5× best single database · ≥ 3 paid design partners (100+ staff) · ≥ 2 implementation partners"],
    ["Gate 1 · Q4 2027", "Stop / lean / scale", "Stop if two of: opps < 15, customers < 3, partner leads < 25, ACV < $10K, win rate < 20%. Scale if ACV ≥ $16K, win ≥ 32%, 15+ customers"],
    ["Gate 2 · Q4 2028", "Scale / hold / sunset", "Scale if 2028-cohort payback ≤ 24 mo and ARR ≥ $1.5M · hold if ARR ≥ $250K and CAC ≤ $55K"],
    ["Gate 3 · Q4 2029", "Graduate or fold", "Standing vertical if ARR ≥ $3M and payback ≤ 18 months; otherwise fold into the horizontal motion"]];
  gates.forEach((g, i) => {
    const x = ML + i * 3.08, p = phases[i];
    box(s, x, 1.45, 2.95, 0.72, i === 1 ? C.accent1 : C.background2, i === 1 ? 88 : 0);
    T(s, [{ text: p[0] + "  ", options: { bold: true, color: C.accent1 } }, { text: p[1], options: { bold: true, breakLine: true } }, { text: p[2], options: { fontSize: 11, color: C.text2 } }],
      { x: x + 0.12, y: 1.52, w: 2.71, h: 0.6, fontSize: 13 });
    const y = 2.35;
    s.addShape(pres.shapes.DIAMOND, { x, y: y + 0.02, w: 0.32, h: 0.32, fill: { color: C.accent1 }, line: { type: "none" }, objectName: nm("gate") });
    T(s, g[0], { x: x + 0.42, y, w: 2.5, h: 0.36, fontSize: 13, bold: true, valign: "middle" });
    rbox(s, x, y + 0.48, 2.95, 1.85, C.background2, 0);
    T(s, g[1], { x: x + 0.15, y: y + 0.6, w: 2.65, h: 0.3, fontSize: 13, bold: true, color: C.accent1 });
    T(s, g[2], { x: x + 0.15, y: y + 0.95, w: 2.65, h: 1.3, fontSize: 12, color: C.text2 });
  });
  T(s, "Next 90 days", { x: ML, y: 5.15, w: 4, h: 0.32, fontSize: 14, bold: true });
  const nx = [["15–20 validation interviews", "Vertical lead, PMM"], ["200-account contractor match-rate test", "GTM eng., data partnerships"], ["3 paid design partners", "Vertical lead, sales"],
    ["2 implementation partners", "Partnerships"], ["Budget cap and AE transfer-cost rule", "Exec sponsor, finance"]];
  nx.forEach((n, i) => {
    const x = ML + i * 2.45;
    s.addShape(pres.shapes.OVAL, { x, y: 5.6, w: 0.36, h: 0.36, fill: { color: C.text1 }, line: { type: "none" }, objectName: nm("num") });
    T(s, String(i + 1), { x, y: 5.6, w: 0.36, h: 0.36, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle" });
    T(s, n[0], { x: x + 0.45, y: 5.56, w: 1.9, h: 0.6, fontSize: 12, bold: true });
    T(s, n[1], { x: x + 0.45, y: 6.15, w: 1.9, h: 0.3, fontSize: 10, color: C.text2 });
  });
  s.addNotes("The roadmap converts the business case into decisions with thresholds. Gate 0 costs only staff time. Gate 1 thresholds sit between the model's Downside and Base paths for stopping, and at the Upside path for scaling, using several indicators because early samples will be small.");
}

pres.addSection({ title: "Appendix" });
// ================================================================== APPENDIX DIVIDER
{
  const s = pres.addSlide({ masterName: "DIVIDER", sectionTitle: "Appendix" });
  s.addText("Appendix", { placeholder: "title" });
  s.addText([
    { text: "A1  Approach and hypothesis verdicts", options: { breakLine: true } },
    { text: "A2  How the market was sized: firms, not establishments", options: { breakLine: true } },
    { text: "A3  Usage-based ACV build", options: { breakLine: true } },
    { text: "A4  Opposing evidence: what would change the recommendation", options: { breakLine: true } },
    { text: "A5  Key scenario assumptions", options: { breakLine: true } },
    { text: "A6  Limitations and principal sources" },
  ], { placeholder: "subtitle" });
}

// ================================================================== A1 HYPOTHESES
{
  const s = content("A1 · Hypothesis-driven, public data only: H2 and H9 were rejected and H7 held only in part",
    "Sources: 01_Project_Charter (hypotheses set before research); verdicts from 03, 05, 06 and 07. Evidence classes: A verified fact · B benchmark · C analyst estimate · D scenario assumption · E recommendation.", "Appendix");
  const V = (t, col) => td(t, { bold: true, color: col });
  const rows = [[th("ID"), th("Hypothesis", { align: "left" }), th("Verdict", { align: "left" }), th("Key evidence", { align: "left" })],
    [td("H1"), td("“B2B manufacturing” is too broad; narrow to core subsectors"), V("Supported", HEX.green), td("Core 4 subsectors ≈ 39% of 20+ manufacturers")],
    [td("H2"), td("Distribution's mid-market pool is ≥ 50% of professional services'"), V("Rejected", HEX.red), td("35% of professional services' 20–499 pool")],
    [td("H3"), td("Industrial buyer personas are poorly covered by LinkedIn-based data"), V("Supported (distribution)", HEX.green), td("26% vs 42% est. presence; proxy only")],
    [td("H4"), td("Builder capacity, not need, is the binding constraint"), V("Supported", HEX.green), td("Computer occupations 0.69× average")],
    [td("H5"), td("Clay's edge is complementary, not substitutive"), V("Supported", HEX.green), td("Layers on CRM; no CRM/ERP of its own")],
    [td("H6"), td("Industrial ACV sits in the Growth-tier range (~$5–25K)"), V("Supported", HEX.green), td("$5.4K SMB · $9.8K mid-market (Base)")],
    [td("H7"), td("Sales-led fails 24-month payback; partner-led can"), V("Partly", HEX.copper), td("Sales-led 108 mo; partner-led 37 mo Base, 12 Upside")],
    [td("H8"), td("Even Upside is a small share of Clay ARR"), V("Supported", HEX.green), td("$6.0M ≈ 3% of est. $200M ARR")],
    [td("H9"), td("Professional services has better risk-adjusted returns"), V("Not supported", HEX.red), td("Wins only in stress test T9")],
    [td("H10"), td("Manufacturers are best reached through distribution"), V("Directional", HEX.copper), td("Manufacturing sales share 0.32× average")]];
  s.addTable(rows, Object.assign({ x: ML, y: 1.45, w: CW, colW: [0.6, 5.0, 2.2, 4.43], rowH: 0.45, objectName: nm("table") }, TBL));
}

// ================================================================== A2 SIZING METHOD
{
  const s = content("A2 · Sizing counts firms, not establishments — establishments would overstate distribution's buyers by ~66%",
    "Sources: Census 2022 Economic Census, Employment Size of Firms (S14); 02_Market_Sizing (de-duplication, imputation, filters) [C/D]; 07_Financial_Model Market Sizing.", "Appendix");
  s.addChart(pres.charts.BAR, [{ name: "Distribution buyer universe", labels: ["Establishments, all sizes", "Firms, all sizes", "Unique firms, 20+ employees", "SAM firms (Phase 1, 4235–4238)", "Pilot SAM (Phase 4, 4236–4238)"], values: [102922, 61863, 11494, 5977, 5168] }],
    Object.assign({ x: ML, y: 1.45, w: 7.3, h: 4.95, barDir: "col", chartColors: [HEX.mist, HEX.mist, HEX.slate, HEX.slate, HEX.copper], showTitle: true, title: "Industrial distribution (NAICS 4235–4238): from establishments to the pilot's market",
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0", valAxisHidden: true, showLegend: false, barGapWidthPct: 45, objectName: nm("chart") }, CHART_TXT, QUIET, { valGridLine: { style: "none" } }));
  box(s, 8.15, 1.45, 4.63, 4.95, C.background2, 0);
  T(s, "Method choices", { x: 8.35, y: 1.6, w: 4.2, h: 0.32, fontSize: 14, bold: true });
  T(s, [
    { text: "Firms are counted in every sub-industry where they operate, so a de-duplication factor (sector ÷ Σ sub-industries) is applied by size band", options: { bullet: true, breakLine: true } },
    { text: "Suppressed 4235 cells (475 firms) recovered as a residual and allocated by the NAICS 423 size mix", options: { bullet: true, breakLine: true } },
    { text: "SAM filters: active outbound motion (80%) × CRM-ready (65%), Base", options: { bullet: true, breakLine: true } },
    { text: "Pilot SAM also excludes metals (4235, ranked last) and an assumed 3% already on Clay", options: { bullet: true, breakLine: true } },
    { text: "Reference year 2022; SUSB cross-check is an open QA item", options: { bullet: true } },
  ], { x: 8.35, y: 2.0, w: 4.25, h: 4.3, fontSize: 12, paraSpaceAfter: 6, color: C.text2 });
}

// ================================================================== A3 ACV BUILD
{
  const s = content("A3 · At list prices, seller density lifts ACV only from mid-market upward",
    "Source: 07_Financial_Model, ACV Build — sellers per firm (OEWS sales share, S15) × 150 accounts × usage meters at published Growth rates (S01); Enterprise floor and 30% discount (S01, S16). Base case [C/D].", "Appendix");
  s.addChart(pres.charts.BAR, [
    { name: "Usage at list price", labels: ["SMB (20–99)", "Mid-market (100–499)", "Enterprise (500+)"], values: [1.95, 9.77, 88.25] },
    { name: "First-year ACV used in model", labels: ["SMB (20–99)", "Mid-market (100–499)", "Enterprise (500+)"], values: [5.35, 9.77, 61.78] },
  ], Object.assign({ x: ML, y: 1.45, w: 6.9, h: 4.95, barDir: "col", barGrouping: "clustered", chartColors: [HEX.mist, HEX.copper], showTitle: true, title: "Annual value per distributor, Base ($K)",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '"$"0.0"K"', valAxisHidden: true, showLegend: true, legendPos: "b", barGapWidthPct: 60, objectName: nm("chart") }, CHART_TXT, QUIET, { valGridLine: { style: "none" } }));
  const rows = [[th("", { align: "left" }), th("SMB"), th("Mid-market"), th("Enterprise")],
    [td("Sellers per firm"), td("9.0", { align: "center" }), td("44.9", { align: "center" }), td("405.6", { align: "center" })],
    [td("Pricing rule"), td("Growth floor $5,352", { align: "center" }), td("Usage", { align: "center" }), td("Usage − 30%", { align: "center" })],
    [td("Phase 1 ACV assumption"), td("$7.5K", { align: "center" }), td("$20K", { align: "center" }), td("$60K", { align: "center" })],
    [td("Phase 4 ACV, Base"), td("$5.4K", { align: "center", bold: true }), td("$9.8K", { align: "center", bold: true }), td("$61.8K", { align: "center", bold: true })],
    [td("Downside / Upside"), td("$5.4K / $5.4K", { align: "center" }), td("$7.3K / $14.7K", { align: "center" }), td("$46K / $93K", { align: "center" })]];
  s.addTable(rows, Object.assign({ x: 7.75, y: 1.5, w: 5.03, colW: [1.55, 1.16, 1.16, 1.16], rowH: 0.52, objectName: nm("table") }, TBL, { fontSize: 11 }));
  box(s, 7.75, 4.85, 5.03, 1.55, C.accent1, 91);
  T(s, [{ text: "Implication  ", options: { bold: true, color: C.accent1 } }, { text: "Phase 1 overstated SMB and mid-market ACV. The pilot's ICP therefore tilts to 100+ employee distributors, and Enterprise ACV checks out against Vendr's $60K median buyer." }],
    { x: 7.95, y: 4.97, w: 4.65, h: 1.35, fontSize: 13 });
}

// ================================================================== A4 OPPOSING EVIDENCE
{
  const s = content("A4 · Opposing evidence: what would change the recommendation",
    "Sources: Clay changelog (S09); Sacra on Apollo (S23); 06_Segment_Prioritization stress tests T6, T9; 07_Financial_Model decision analysis [C/D].", "Appendix");
  const rows = [[th("Evidence against", { align: "left" }), th("Why it does not (yet) overturn the call", { align: "left" }), th("What would overturn it", { align: "left" })],
    [td("Clay's Sep 2026 data expansion targeted public sector, real estate and PE — not industrials (S09)"), td("Could be a deliberate non-priority or an unaddressed gap; public data cannot tell which"), td("Clay or partner evidence that industrial deals underperform")],
    [td("Apollo is converging: waterfall enrichment, spreadsheet beta (S23)"), td("None of its moves address off-LinkedIn contractor data; T6 (position 4 → 3) still ranks distribution first"), td("Apollo or ZoomInfo adding Maps- or permit-based contractor data")],
    [td("The Base case loses money within three years"), td("The gate truncates Downside and Upside is large, so staged EV is positive whenever P(Upside) exceeds ~7%"), td("Downside indistinguishable from Base by Q4 2027 (ungated EV falls to +$0.65M)")],
    [td("Professional services wins if ease of execution weighs ≥ 26% (or in T9)"), td("Distribution leads in 11 of 12 stress tests and 5 of 6 weight sets"), td("Evidence that partners cannot make distribution adoption easier")],
    [td("No primary research; the data-gap measure is a proxy"), td("Regency Supply corroborates it directly (S07)"), td("Gate 0 match-rate uplift below 1.5×")]];
  s.addTable(rows, Object.assign({ x: ML, y: 1.45, w: CW, colW: [4.2, 4.4, 3.63], rowH: 0.85, objectName: nm("table") }, TBL));
}

// ================================================================== A5 ASSUMPTIONS
{
  const s = content("A5 · Key scenario assumptions (Downside / Base / Upside)",
    "Source: 07_Financial_Model, Assumptions — every input carries its evidence class and source. Benchmarks: SaaS Capital 2025 (S43), Bridge Group 2026 (S44), HubSpot/Ebsta/Norwest via S48, Digital Bloom (S49), partner margins (S50), Sacra (S52).", "Appendix");
  const A = [["SAM filters (outbound × CRM-ready)", "65% × 52%", "80% × 65%", "90% × 76%", "D/B", "Phase 1; DSG 2022 (S17)"],
    ["Win rate (opportunity → won)", "20%", "28%", "35%", "B", "HubSpot avg 21%; opp→won 31–39% (S48, S49)"],
    ["Qualification / SQL → opportunity", "35% / 35%", "45% / 45%", "55% / 55%", "B", "MQL→SQL 36–39%; SQL→opp 42% (S48, S49)"],
    ["Partners: cap / leads per partner per quarter", "5 / 2", "10 / 3", "15 / 4", "D", "No distribution partners visible today (S35)"],
    ["Usage intensity (sets ACV)", "0.75×", "1.0×", "1.5×", "D", "Usage build at list prices (S01)"],
    ["Gross revenue retention, mid-market", "86%", "90%", "93%", "B", "SaaS Capital: 91% at $12–25K ACV (S43)"],
    ["Net revenue retention, mid-market", "98%", "106%", "115%", "B", "SaaS Capital: 103% at $12–25K ACV (S43)"],
    ["Gross margin", "55%", "65%", "72%", "D", "Data-heavy margins below 80% (S52)"],
    ["AE quota / attainment", "$875K / 55%", "$875K / 70%", "$875K / 85%", "B", "Bridge Group 2026 (S44)"],
    ["Partner fee on partner-sourced ACV", "25%", "20%", "15%", "B", "Referral 10–15%, VAR 20–30% (S50)"],
    ["Waves 2 / 3 launch", "Q3 2028 / never", "Q1 2028 / Q3 2028", "Q4 2027 / Q2 2028", "D", "Sequenced after gates"]];
  const rows = [[th("Input", { align: "left" }), th("Downside"), th("Base"), th("Upside"), th("Class"), th("Basis", { align: "left" })]]
    .concat(A.map((a) => [td(a[0]), td(a[1], { align: "center" }), td(a[2], { align: "center", bold: true }), td(a[3], { align: "center" }), td(a[4], { align: "center" }), td(a[5])]));
  s.addTable(rows, Object.assign({ x: ML, y: 1.45, w: CW, colW: [3.3, 1.6, 1.6, 1.6, 0.7, 3.43], rowH: 0.42, objectName: nm("table") }, TBL, { fontSize: 11 }));
}

// ================================================================== A6 LIMITATIONS & SOURCES
{
  const s = content("A6 · Limitations and principal sources",
    "Full register (55 sources with dates, reliability and limitations): Research_Sources.csv. Workbooks and build scripts: 02, 04, 06, 07 and appendix/.", "Appendix");
  T(s, "Limitations", { x: ML, y: 1.45, w: 5.8, h: 0.35, fontSize: 16, bold: true });
  T(s, [
    { text: "No interviews, surveys or Clay internal data; win rates, margins and partner economics are benchmarks or assumptions", options: { bullet: true, breakLine: true } },
    { text: "Database-coverage gap is an education-based proxy, not a measured match rate", options: { bullet: true, breakLine: true } },
    { text: "Census firm counts are for 2022; Clay pricing changed in March 2026 and may change again", options: { bullet: true, breakLine: true } },
    { text: "Distribution surveys are small and industry-sponsored", options: { bullet: true, breakLine: true } },
    { text: "Scenario probabilities are explicit judgments; Clay's ARR (~$200M) is a third-party estimate", options: { bullet: true, breakLine: true } },
    { text: "Some competitor capabilities rated from general knowledge [K] need verification", options: { bullet: true } },
  ], { x: ML, y: 1.9, w: 5.8, h: 4.6, fontSize: 13, paraSpaceAfter: 6, color: C.text2 });
  const src = [[th("ID"), th("Source", { align: "left" }), th("Date")],
    ["S01", "Clay pricing page (plans, credits, actions)", "Oct 2026"], ["S07", "Clay case study: Regency Supply", "Live page"], ["S14", "U.S. Census Bureau, 2022 Economic Census — Employment Size of Firms", "Ref. 2022"],
    ["S15", "U.S. BLS, OEWS industry-specific estimates", "May 2025"], ["S16", "Vendr, Clay buyer pricing data", "Feb 2026"], ["S23", "Sacra, Apollo and Clay revenue estimates", "Sep 2026"],
    ["S25", "Phocas / DSG, State of Sales in Distribution", "2026"], ["S43", "SaaS Capital, B2B SaaS Retention Benchmarks", "Sep 2025"], ["S44", "Bridge Group, AE Models, Motions & Metrics", "Jun 2026"]]
    .map((r, i) => (i === 0 ? r : [td(r[0], { align: "center" }), td(r[1]), td(r[2], { align: "center" })]));
  s.addTable(src, Object.assign({ x: 6.75, y: 1.5, w: 6.03, colW: [0.6, 4.13, 1.3], rowH: 0.45, objectName: nm("table") }, TBL, { fontSize: 11 }));
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
