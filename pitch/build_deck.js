// Stock pitch deck: Endurance Technologies. Numbers come from dcf/model_outputs.json
// (the recalculated DCF workbook) and from the sourced figures noted on each slide.
//   node pitch/build_deck.js  ->  pitch/Endurance_Stock_Pitch.pptx

const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/mnt/skills/public/pptx/scripts/apply_theme.js");

const OUT = path.join(__dirname, "Endurance_Stock_Pitch.pptx");
const M = JSON.parse(fs.readFileSync(path.join(__dirname, "..", "dcf", "model_outputs.json"), "utf8"));
const PRICE = 2406.4;
const fmt = (n) => Math.round(n).toLocaleString("en-IN");
const pct = (x) => (x >= 0 ? "+" : "") + (x * 100).toFixed(0) + "%";
const BASE = M.Base.value, BEAR = M.Bear.value, BULL = M.Bull.value;

const THEME = {
  name: "Graphite Amber",
  headFontFace: "Calibri",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1B2530", lt1: "FFFFFF", dk2: "2C3E50", lt2: "F2F4F6",
    accent1: "E39B2D", // molten-aluminium amber
    accent2: "5D7183", // steel
    accent3: "B23A3A", // downside red
    accent4: "2E7D5B", // upside green
    accent5: "A9B6C2", accent6: "D9DEE3",
    hlink: "1F5F99", folHlink: "5D7183",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625 in
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Endurance Technologies: Stock Pitch";
const C = pres.SchemeColor;

const FOOT = "Student equity research for learning, not investment advice. Prices as of 16 Oct 2026 (NSE close).";

pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: C.text1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.2, w: 8.8, h: 1.5, fontSize: 36, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 2.75, w: 8.8, h: 0.8, fontSize: 18, color: C.accent5, valign: "top", align: "left", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  margin: [0.5, 0.5, 0.6, 0.5],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 9.0, h: 0.6, fontSize: 28, bold: true, color: C.text1, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.5, y: 0.92, w: 9.0, h: 0.4, fontSize: 14, color: C.accent2, valign: "top", align: "left", margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: 0.5, y: 5.25, w: 8.2, h: 0.25, fontSize: 8, color: C.accent2, margin: 0 } } },
  ],
  slideNumber: { x: 9.1, y: 5.25, w: 0.4, h: 0.25, fontSize: 8, color: C.accent2, align: "right" },
});

let n = 0;
function content(title, sub, section) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  if (sub) s.addText(sub, { placeholder: "body" });
  n += 1;
  return s;
}
const tb = (s, text, o) => s.addText(text, { isTextBox: true, margin: 0, fontSize: 14, color: C.text1, valign: "top", ...o });
function card(s, x, y, w, h, name) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: C.background2 }, line: { color: C.background2 }, objectName: name });
}
function stat(s, x, y, w, big, label, color, name) {
  card(s, x, y, w, 1.35, name + "-card");
  tb(s, big, { x: x + 0.2, y: y + 0.15, w: w - 0.4, h: 0.65, fontSize: 34, bold: true, color: color || C.text1, objectName: name + "-num" });
  tb(s, label, { x: x + 0.2, y: y + 0.82, w: w - 0.4, h: 0.45, fontSize: 11, color: C.accent2, objectName: name + "-label" });
}
const chartBase = (title) => ({
  showTitle: true, title, titleFontSize: 12, titleColor: HEX.dk2, titleFontFace: "+mn-lt",
  catAxisLabelColor: HEX.accent2, valAxisLabelColor: HEX.accent2, catAxisLabelFontSize: 10, valAxisLabelFontSize: 9,
  catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
  valGridLine: { color: HEX.accent6, size: 0.5 }, catGridLine: { style: "none" },
  dataLabelFontSize: 10, dataLabelColor: HEX.dk1, dataLabelFontFace: "+mn-lt",
});

// 1. Title ----------------------------------------------------------------
pres.addSection({ title: "Summary" });
{
  const s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Summary" });
  s.addText("Endurance Technologies: a quality franchise, priced for the bull case", { placeholder: "title" });
  s.addText("NSE: ENDURANCE  ·  Auto components  ·  Stock pitch, October 2026", { placeholder: "body" });
  const items = [["Rating", "UNDERWEIGHT", HEX.accent1], ["Price (16 Oct 26)", "₹" + fmt(PRICE), HEX.lt1],
                 ["DCF fair value", "₹" + fmt(BASE), HEX.lt1], ["Implied return", pct(BASE / PRICE - 1), HEX.accent1]];
  items.forEach(([k, v, col], i) => {
    const x = 0.6 + i * 2.25;
    tb(s, k, { x, y: 3.85, w: 2.1, h: 0.3, fontSize: 11, color: C.accent5, objectName: "kpi-label-" + i });
    tb(s, v, { x, y: 4.15, w: 2.1, h: 0.5, fontSize: 22, bold: true, color: col, objectName: "kpi-value-" + i });
  });
  tb(s, "Student research for learning purposes, not investment advice.", { x: 0.6, y: 5.05, w: 8.8, h: 0.3, fontSize: 9, color: C.accent5 });
  s.addNotes("Open with the call: great company, wrong price. The DCF base case is about 35% below the market price; even the bull case is below it.");
}

// 2. Investment summary ----------------------------------------------------
{
  const s = content("The call in one slide", "We like the business. We don't like what the price already assumes.", "Summary");
  stat(s, 0.5, 1.5, 2.85, "₹" + fmt(BASE), "Base-case DCF value per share (Bear ₹" + fmt(BEAR) + " / Bull ₹" + fmt(BULL) + ")", C.accent3, "s2a");
  stat(s, 3.575, 1.5, 2.85, "~18%", "10-yr revenue CAGR needed to justify ₹2,406 (11% delivered last decade)", C.text1, "s2b");
  stat(s, 6.65, 1.5, 2.85, "15.0x", "EV / FY27E EBITDA at today's price, vs 9.8x at our base-case value", C.text1, "s2c");
  const pts = [
    { text: "Good business: ", options: { bold: true } }, { text: "34 plants, leader in 2W suspension and castings, Bajaj share falling (55-60% → 50% of standalone), ₹1,806 Cr EV orders.", options: { breakLine: true } },
    { text: "But margins are slipping: ", options: { bold: true } }, { text: "Q1 FY27 standalone EBITDA margin 11.2% vs 13.0% a year ago, on input costs.", options: { breakLine: true } },
    { text: "And growth is capital-hungry: ", options: { bold: true } }, { text: "capex ran at ~9% of sales in FY26; free cash flow was only 58% of net profit.", options: {} },
  ];
  tb(s, pts, { x: 0.5, y: 3.1, w: 9.0, h: 1.9, fontSize: 14, paraSpaceAfter: 8, objectName: "s2-points" });
  s.addNotes("Three numbers: fair value, the growth the market is pricing in, and the multiple gap. Then the three reasons.");
}

// 3. Company snapshot ------------------------------------------------------
pres.addSection({ title: "Business" });
{
  const s = content("What Endurance makes", "Aluminium castings, suspension, brakes and wheels, mainly for Indian two-wheelers and European cars.", "Business");
  s.addChart(pres.charts.DOUGHNUT, [{ name: "Mix", labels: ["Die casting", "Suspension", "Disc brakes", "Alloy wheels", "Transmission & aftermarket", "Others"], values: [44.5, 24.2, 11.0, 8.5, 6.5, 5.3] }], {
    x: 0.4, y: 1.4, w: 4.6, h: 3.7, holeSize: 55, showLegend: true, legendPos: "r", legendFontSize: 10, legendFontFace: "+mn-lt", legendColor: HEX.dk2,
    showPercent: true, showValue: false, dataLabelColor: HEX.lt1, dataLabelFontSize: 9, dataLabelFontFace: "+mn-lt",
    chartColors: [HEX.accent1, HEX.dk2, HEX.accent2, HEX.accent5, HEX.accent6, "8A6A3A"],
    showTitle: true, title: "Consolidated revenue by product, Q1 FY27", titleFontSize: 12, titleColor: HEX.dk2, titleFontFace: "+mn-lt",
  });
  const facts = [
    ["₹14,720 Cr", "FY26 total income, +26% (consolidated)"],
    ["73% / 27%", "India / Europe (Italy, Germany) split of FY26 income"],
    ["34 plants", "20 India, 9 Italy, 5 Germany"],
    ["~83%", "of India revenue from 2W (motorcycles 68%, scooters 14%)"],
  ];
  facts.forEach(([a, b], i) => {
    const y = 1.5 + i * 0.9;
    card(s, 5.3, y, 4.2, 0.75, "s3-card-" + i);
    tb(s, a, { x: 5.45, y: y + 0.1, w: 1.55, h: 0.55, fontSize: 16, bold: true, color: C.accent1, valign: "middle" });
    tb(s, b, { x: 7.0, y: y + 0.1, w: 2.4, h: 0.55, fontSize: 11, color: C.text1, valign: "middle" });
  });
  s.addNotes("Sources: FY26 results press release (NSE, 14 May 2026); Q1 FY27 investor presentation (NSE, 13 Aug 2026).");
}

// 4. Bull points we accept --------------------------------------------------
{
  const s = content("The bull case is real", "Revenue nearly doubled in four years, and the order book is diversifying away from one customer.", "Business");
  s.addChart(pres.charts.BAR, [{ name: "Revenue", labels: ["FY22", "FY23", "FY24", "FY25", "FY26"], values: [7549, 8804, 10241, 11561, 14596] }], {
    ...chartBase("Consolidated revenue (₹ Cr)"), x: 0.4, y: 1.4, w: 5.0, h: 3.7, barDir: "col", chartColors: [HEX.dk2],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0", valAxisHidden: true, valGridLine: { style: "none" }, showLegend: false,
  });
  const rows = [
    ["Order wins (ex-Bajaj)", "₹5,270 Cr over 5 years; ₹1,205 Cr in FY26"],
    ["EV orders", "₹1,806 Cr cumulative incl. Bajaj; castings, brakes, battery packs"],
    ["Customer mix", "Bajaj 50% of standalone revenue, down from 55-60% (ICRA)"],
    ["New products", "Dual-channel ABS, alloy wheels, 4W castings, BMS electronics"],
  ];
  rows.forEach(([k, v], i) => {
    const y = 1.5 + i * 0.9;
    card(s, 5.7, y, 3.8, 0.75, "s4-card-" + i);
    tb(s, k, { x: 5.85, y: y + 0.08, w: 3.5, h: 0.28, fontSize: 12, bold: true, color: C.text1 });
    tb(s, v, { x: 5.85, y: y + 0.36, w: 3.5, h: 0.35, fontSize: 11, color: C.accent2 });
  });
  s.addNotes("Revenue: StockAnalysis.com (S&P Global). Orders: Q1 FY27 investor presentation. Bajaj share: ICRA rationale, 30 Sep 2025.");
}

// 5. Thesis 1: margins ------------------------------------------------------
pres.addSection({ title: "Thesis" });
{
  const s = content("Thesis 1: margins are going the wrong way", "Revenue grew 30% in Q1 FY27, but EBITDA grew only 19%.", "Thesis");
  s.addChart(pres.charts.BAR, [
    { name: "Q1 FY26", labels: ["Consolidated", "India (standalone)", "Europe"], values: [0.143, 0.130, 0.174] },
    { name: "Q1 FY27", labels: ["Consolidated", "India (standalone)", "Europe"], values: [0.131, 0.112, 0.182] },
  ], {
    ...chartBase("EBITDA margin, Q1 FY26 vs Q1 FY27"), x: 0.4, y: 1.4, w: 5.4, h: 3.7, barDir: "col", barGapWidthPct: 60,
    chartColors: [HEX.accent5, HEX.accent1], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0%",
    valAxisHidden: true, valGridLine: { style: "none" }, showLegend: true, legendPos: "b", legendFontSize: 10, legendFontFace: "+mn-lt", legendColor: HEX.dk2,
  });
  tb(s, [
    { text: "India is 72% of revenue, and its margin fell 1.8 pts.", options: { bold: true, breakLine: true } },
    { text: "Management blames higher input costs. Aluminium and energy are pass-through with a lag, but new plants (ABS, alloy wheels, 4W castings) also carry start-up costs.", options: { breakLine: true } },
    { text: "Europe's margin held, but PAT fell 23% on accelerated depreciation of ICE programmes being phased out.", options: { breakLine: true } },
    { text: "Our base case still assumes recovery to 13.8% by FY30. If margins stay near 12.5% (bear), fair value is ₹" + fmt(BEAR) + ".", options: {} },
  ], { x: 6.1, y: 1.5, w: 3.4, h: 3.5, fontSize: 12, paraSpaceAfter: 8, objectName: "s5-text" });
  s.addNotes("Source: Q1 FY27 investor presentation (NSE, 13 Aug 2026).");
}

// 6. Thesis 2: cash ---------------------------------------------------------
{
  const s = content("Thesis 2: growth eats the cash", "Capex has grown faster than operating cash flow; free cash flow is only ~58% of profit.", "Thesis");
  const yrs = ["FY22", "FY23", "FY24", "FY25", "FY26"];
  s.addChart(pres.charts.BAR, [
    { name: "Operating cash flow", labels: yrs, values: [742, 862, 1057, 1532, 1851] },
    { name: "Capex", labels: yrs, values: [529, 636, 835, 1053, 1296] },
    { name: "Free cash flow", labels: yrs, values: [213, 226, 222, 478, 555] },
  ], {
    ...chartBase("₹ Cr"), x: 0.4, y: 1.4, w: 6.0, h: 3.7, barDir: "col", barGapWidthPct: 50,
    chartColors: [HEX.dk2, HEX.accent5, HEX.accent1], showValue: false, showLegend: true, legendPos: "b",
    legendFontSize: 10, legendFontFace: "+mn-lt", legendColor: HEX.dk2, valAxisLabelFormatCode: "#,##0",
  });
  stat(s, 6.7, 1.5, 2.8, "58%", "FY26 free cash flow / net profit (₹555 Cr / ₹952 Cr)", C.accent3, "s6a");
  stat(s, 6.7, 3.05, 2.8, "8.9%", "FY26 capex / revenue, plus ₹358 Cr spent on acquisitions", C.text1, "s6b");
  s.addNotes("Source: StockAnalysis.com cash flow statement (S&P Global). Mgmt expects FY27 capex similar to FY26.");
}

// 7. Thesis 3: what the price implies --------------------------------------
{
  const s = content("Thesis 3: the price assumes a lot", "Reverse DCF: what has to be true for ₹2,406 to be fair value, holding everything else at our base case.", "Thesis");
  const cols = [
    ["17.9%", "10-year revenue CAGR", "vs 11.3% in our base case and 11% actually delivered over the last 10 years"],
    ["9.2%", "WACC", "vs 11.2% from CAPM (Rf 5.3%, beta 0.87, India ERP 7.1%)"],
    ["17.4%", "EBITDA margin, every year", "vs 13.2% in FY26 and a five-year high of 13.2%"],
  ];
  cols.forEach(([big, k, v], i) => {
    const x = 0.5 + i * 3.075;
    card(s, x, 1.5, 2.85, 2.6, "s7-card-" + i);
    tb(s, big, { x: x + 0.2, y: 1.65, w: 2.45, h: 0.75, fontSize: 40, bold: true, color: C.accent1 });
    tb(s, k, { x: x + 0.2, y: 2.45, w: 2.45, h: 0.35, fontSize: 14, bold: true, color: C.text1 });
    tb(s, v, { x: x + 0.2, y: 2.85, w: 2.45, h: 1.1, fontSize: 11, color: C.accent2 });
  });
  tb(s, "Any one of these alone would justify today's price. We think none of them is the most likely outcome.", { x: 0.5, y: 4.35, w: 9.0, h: 0.5, fontSize: 14, bold: true, color: C.text1 });
  s.addNotes("Computed by solving the DCF for each input separately (dcf/build_dcf.py model, base case otherwise). The 10-year history CAGR is from Screener.");
}

// 8. Valuation --------------------------------------------------------------
pres.addSection({ title: "Valuation" });
{
  const s = content("DCF: fair value ₹" + fmt(BASE) + " in the base case", "Even the bull case is below today's price.", "Valuation");
  s.addChart(pres.charts.BAR, [{ name: "Value per share", labels: ["Bear", "Base", "Bull", "Price today"], values: [BEAR, BASE, BULL, PRICE].map(Math.round) }], {
    ...chartBase("Value per share (₹)"), x: 0.4, y: 1.4, w: 5.0, h: 3.7, barDir: "col", barGapWidthPct: 45,
    chartColors: [HEX.accent5, HEX.accent1, HEX.accent2, HEX.dk1], showValue: true, dataLabelPosition: "outEnd",
    dataLabelFormatCode: "#,##0", valAxisHidden: true, valGridLine: { style: "none" }, showLegend: false,
  });
  const rows = [
    ["", "Bear", "Base", "Bull"],
    ["10-yr revenue CAGR", "7.8%", "11.3%", "13.9%"],
    ["EBITDA margin (steady)", "12.5%", "13.8%", "14.8%"],
    ["Value / share", "₹" + fmt(BEAR), "₹" + fmt(BASE), "₹" + fmt(BULL)],
    ["vs price", pct(BEAR / PRICE - 1), pct(BASE / PRICE - 1), pct(BULL / PRICE - 1)],
  ];
  s.addTable(rows.map((r, i) => r.map((c, j) => ({ text: c, options: { bold: i === 0 || j === 0, color: HEX.dk1, fill: { color: i === 0 ? HEX.lt2 : HEX.lt1 }, align: j ? "right" : "left" } }))), {
    x: 5.7, y: 1.5, w: 3.8, colW: [1.6, 0.73, 0.74, 0.73], fontSize: 11, fontFace: "Calibri", border: { type: "solid", pt: 0.5, color: HEX.accent6 }, rowH: 0.36,
  });
  tb(s, "WACC 11.2% · terminal growth 5% · FCFF, mid-year discounting · terminal value ~66% of EV. Full model: dcf/Endurance_DCF.xlsx", { x: 5.7, y: 3.6, w: 3.8, h: 0.8, fontSize: 11, color: C.accent2 });
  s.addNotes("All values from the recalculated workbook with the scenario switch set to 1, 2 and 3.");
}

// 9. Sensitivity ------------------------------------------------------------
{
  const s = content("How sure are we? Sensitivity", "Value per share (₹) by WACC and terminal growth. Only the most optimistic corner reaches today's price.", "Valuation");
  const g = M.grid; // [ [null, g1..g5], [w, v..], ...]
  const header = [{ text: "WACC \\ g", options: { bold: true, fill: { color: HEX.lt2 }, color: HEX.dk1 } }]
    .concat(g[0].slice(1).map((x) => ({ text: (x * 100).toFixed(1) + "%", options: { bold: true, fill: { color: HEX.lt2 }, color: HEX.dk1, align: "right" } })));
  const body = g.slice(1).map((r) => [{ text: (r[0] * 100).toFixed(2) + "%", options: { bold: true, color: HEX.dk1, fill: { color: HEX.lt2 } } }]
    .concat(r.slice(1).map((v) => ({ text: fmt(v), options: { align: "right", color: v >= PRICE ? HEX.lt1 : HEX.dk1, bold: v >= PRICE, fill: { color: v >= PRICE ? HEX.accent4 : HEX.lt1 } } }))));
  s.addTable([header, ...body], { x: 0.5, y: 1.5, w: 5.6, colW: [1.1, 0.9, 0.9, 0.9, 0.9, 0.9], fontSize: 12, fontFace: "Calibri", rowH: 0.4, border: { type: "solid", pt: 0.5, color: HEX.accent6 } });
  tb(s, [
    { text: "Green = at or above ₹2,406.", options: { bold: true, breakLine: true } },
    { text: "Our WACC is 11.2%. Getting to today's price needs WACC under ~10% and 6% growth forever.", options: { breakLine: true } },
    { text: "Exit-multiple check: a 12x terminal EV/EBITDA gives ₹" + fmt(M.exit[1][2]) + ", but implies " + (M.exit[2][2] * 100).toFixed(1) + "% perpetual growth.", options: {} },
  ], { x: 6.4, y: 1.5, w: 3.1, h: 3.4, fontSize: 12, paraSpaceAfter: 8 });
  s.addNotes("Grid computed live in the Sensitivity tab of the workbook.");
}

// 10. Relative valuation: the counter-argument ------------------------------
{
  const s = content("Counterpoint: cheapest in its peer group", "On multiples, Endurance trades at a large discount to listed peers. That is the main risk to our call.", "Valuation");
  const peers = [["Gabriel India", 84.2], ["Sona BLW", 71.6], ["Uno Minda", 53.3], ["Endurance", 34.9], ["Minda Corp", 31.0]];
  s.addChart(pres.charts.BAR, [{ name: "P/E", labels: peers.map((p) => p[0]), values: peers.map((p) => p[1]) }], {
    ...chartBase("Trailing P/E (x), Oct 2026"), x: 0.4, y: 1.4, w: 5.2, h: 3.7, barDir: "bar", barGapWidthPct: 45,
    chartColors: [HEX.accent5, HEX.accent5, HEX.accent5, HEX.accent1, HEX.accent5], showValue: true, dataLabelPosition: "outEnd",
    dataLabelFormatCode: '0.0"x"', valAxisHidden: true, valGridLine: { style: "none" }, showLegend: false, catAxisOrientation: "maxMin",
  });
  tb(s, [
    { text: "Peer median P/E is ~62x vs 35x for Endurance.", options: { bold: true, breakLine: true } },
    { text: "Applying the peer median to TTM EPS (₹68.95) gives ~₹4,300.", options: { breakLine: true } },
    { text: "Why we still lean on the DCF: ", options: { bold: true } },
    { text: "the whole sector trades on growth that few companies deliver for a decade. Relative value tells us Endurance is cheaper than its peers, not that it is cheap.", options: {} },
  ], { x: 5.9, y: 1.5, w: 3.6, h: 3.4, fontSize: 12, paraSpaceAfter: 8 });
  s.addNotes("Peer data: Value Research and 5paisa stock pages, 5-16 Oct 2026. Bharat Forge excluded (126x P/E, defence mix). Gabriel and Bharat Forge prices are as of 5 Oct 2026.");
}

// 11. Risks and catalysts ---------------------------------------------------
pres.addSection({ title: "Risks" });
{
  const s = content("What would make us wrong", "Signals to watch, and what each would mean for the call.", "Risks");
  const risks = [
    ["Margin rebound", "Standalone EBITDA back above 13% for two quarters as ABS and alloy-wheel plants fill up."],
    ["Faster non-Bajaj wins", "Non-Bajaj orders running well above the ~₹1,100 Cr/yr average of FY24-FY26."],
    ["Sector re-rating", "Peers stay at 50-80x and money rotates into the cheapest quality name."],
    ["Lower rates", "RBI raised to 5.5% in Oct 2026; a cutting cycle would lower our WACC (each 0.5 pt is worth ~₹130-150/share)."],
  ];
  risks.forEach(([k, v], i) => {
    const x = 0.5 + (i % 2) * 4.6, y = 1.5 + Math.floor(i / 2) * 1.7;
    card(s, x, y, 4.4, 1.5, "s11-card-" + i);
    s.addShape(pres.shapes.OVAL, { x: x + 0.2, y: y + 0.22, w: 0.42, h: 0.42, fill: { color: C.accent1 }, line: { color: C.accent1 } });
    tb(s, String(i + 1), { x: x + 0.2, y: y + 0.22, w: 0.42, h: 0.42, fontSize: 14, bold: true, color: C.background1, align: "center", valign: "middle" });
    tb(s, k, { x: x + 0.8, y: y + 0.22, w: 3.4, h: 0.4, fontSize: 15, bold: true, color: C.text1, valign: "middle" });
    tb(s, v, { x: x + 0.8, y: y + 0.7, w: 3.4, h: 0.7, fontSize: 11, color: C.accent2 });
  });
  s.addNotes("Rate move: Trading Economics news item, 13 Oct 2026. Sensitivity per 0.5pt WACC from the Sensitivity tab.");
}

// 12. Conclusion ------------------------------------------------------------
{
  const s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Risks" });
  s.addText("Great business, wrong price: underweight at ₹2,406", { placeholder: "title" });
  s.addText("We would turn positive near our base-case value (~₹" + fmt(BASE) + "), or sooner if standalone margins recover above 13%.", { placeholder: "body" });
  const src = "Sources: Endurance FY26 results press release and Q1 FY27 investor presentation (NSE filings); StockAnalysis.com (S&P Global) financials; " +
    "ICRA rating rationale (Sep 2025); Damodaran country risk premiums (Jan 2026); Trading Economics (India 10Y, Sep 2026); " +
    "Value Research, 5paisa, Dhan, sharpely market data (Oct 2026); ACMA and SIAM FY26 industry data.";
  tb(s, src, { x: 0.6, y: 4.2, w: 8.8, h: 0.9, fontSize: 10, color: C.accent5 });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("Wrote", OUT);
})();
