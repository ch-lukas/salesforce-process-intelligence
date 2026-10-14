// Builds deck/salesforce-process-intelligence.pptx from data/insights.json.
// Colour theme follows the Salesforce CRM handbook: navy, cloud blue, light-blue panels.
// Usage (from deck/): npm ci && npm run build
//   DECK_KIND=tech builds the tech setup deck; DECK_LOGO=path/to/logo.svg adds a logo to the closing slide.
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");

const { applyTheme } = require("./theme");

const ROOT = path.join(__dirname, "..");
const I = JSON.parse(fs.readFileSync(path.join(ROOT, "data/insights.json"), "utf8"));
const LIVE = JSON.parse(fs.readFileSync(path.join(ROOT, "data/live/insights.json"), "utf8"));
// DECK_LOGO is optional: the shared repo version is built without a logo.
// DECK_KIND=tech builds the separate Tech setup deck; anything else builds the main story deck.
const KIND = process.env.DECK_KIND === "tech" ? "tech" : "main";
const OUT = process.env.DECK_OUT || path.join(__dirname, KIND === "tech" ? "salesforce-process-intelligence-tech-setup.pptx" : "salesforce-process-intelligence.pptx");
const LOGO = process.env.DECK_LOGO || "";
const REPO = "github.com/ch-lukas/salesforce-process-intelligence";
const SITE = "salesforce-process-intelligence.vercel.app";

const THEME = {
  name: "Process Intelligence",
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  colors: {
    dk1: "181818", lt1: "FFFFFF", dk2: "032D60", lt2: "EAF5FE",
    accent1: "0B9DE0", accent2: "4499D8", accent3: "CFE9FE",
    accent4: "FE9339", accent5: "2E844A", accent6: "0B5CAB",
    hlink: "0B5CAB", folHlink: "032D60",
  },
};
const HEX = THEME.colors;

// ---------- numbers ----------
const K = I.kpis;
const D = Object.fromEntries(I.value_drivers.map(d => [d.id, d]));
const winPair = d => d.evidence.match(/win (\d+)% vs (\d+)%/).slice(1).map(Number);
const [rwBad, rwGood] = winPair(D.rework);
const [qBad, qGood] = winPair(D.qualification);
const apWin = winPair(D.approval); // [waited >7 days, waited <=7 days]
const regions = Object.entries(I.by_region).filter(([, v]) => v.approval_wait_days != null);
const negotiated = D.rework.negotiated_per_year; // exact, unrounded
const framed = D.rework.annual_value * 0.5; // same calculation as the dashboard
const usd = v => (v >= 1e6 ? `$${(v / 1e6).toFixed(1)}M` : `$${Math.round(v / 1e3)}k`);
const usd2 = v => (v >= 1e6 ? `$${(v / 1e6).toFixed(2)}M` : `$${Math.round(v / 1e3)}k`);
const pct = v => `${Math.round(v * 100)}%`;
const days = v => Number(v).toFixed(1);
const headline = I.value_drivers.reduce((s, d) => s + d.annual_value * d.default_capture, 0);

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
  pres.title = KIND === "tech" ? "Tech setup: from a Salesforce org to Process Intelligence" : "From Salesforce data to a framed value case";
  pres.author = "Luke Audie";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;
  const hasLogo = Boolean(LOGO) && fs.existsSync(LOGO);

  // ---------- layouts ----------
  pres.defineSlideMaster({
    title: "TITLE_DARK",
    background: { color: C.text2 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.2, w: 9.2, h: 1.9, fontSize: 44, bold: true, color: C.background1, valign: "bottom", align: "left", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.3, w: 8.6, h: 1.1, fontSize: 20, color: C.accent3, valign: "top", align: "left", margin: 0 }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: C.background1 },
    margin: [0.5, 0.6, 0.6, 0.6],
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 0.9, fontSize: 30, bold: true, color: C.text2, valign: "middle", align: "left", margin: 0 }, text: "" } },
      { text: { text: "Salesforce → event log → process mining → value case", options: { x: 0.6, y: 7.0, w: 8, h: 0.3, fontSize: 10, color: C.accent6, margin: 0 } } },
    ],
    slideNumber: { x: 12.2, y: 7.0, w: 0.5, h: 0.3, fontSize: 10, color: C.accent6, align: "right" },
  });
  pres.defineSlideMaster({
    title: "SECTION_DARK",
    background: { color: C.text2 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.6, w: 8.6, h: 1.2, fontSize: 44, bold: true, color: C.background1, valign: "bottom", align: "left", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.0, w: 8.0, h: 1.0, fontSize: 20, color: C.accent3, valign: "top", align: "left", margin: 0 }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "CLOSING_DARK",
    background: { color: C.text2 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 0.6, w: 11.7, h: 1.0, fontSize: 36, bold: true, color: C.background1, valign: "middle", align: "left", margin: 0 }, text: "" } },
    ],
  });

  const logo = (slide, x, y, h) => {
    if (!hasLogo) return;
    const svg = fs.readFileSync(LOGO, "utf8");
    const vb = (svg.match(/viewBox="([^"]+)"/) || [])[1];
    const [, , vw, vh] = vb ? vb.split(/[\s,]+/).map(Number) : [0, 0, 4, 1];
    slide.addImage({ path: LOGO, x, y, h, w: h * (vw / vh), objectName: "Logo", altText: "Logo" });
  };

  const circle = (slide, x, y, d, color, transparency, name) =>
    slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color, transparency }, line: { type: "none" }, objectName: name });

  const card = (slide, x, y, w, h, name, fill = C.background2) =>
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { type: "none" }, objectName: name });

  const text = (slide, value, opts) => slide.addText(value, { isTextBox: true, margin: 0, color: C.text1, fontSize: 14, valign: "top", ...opts });

  let s;
  if (KIND === "main") {
  // ---------- Opening: title ----------
  pres.addSection({ title: "Opening" });
  s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Opening" });
  circle(s, 9.4, 3.4, 6.6, C.accent1, 70, "Decor circle large");
  circle(s, 10.6, 2.6, 0.75, C.accent3, 0, "Decor circle small");
  s.addText("From Salesforce data to a framed value case", { placeholder: "title" });
  s.addText("Simulated value case with a working Salesforce extraction pipeline", { placeholder: "body" });
  text(s, "Luke Audie  ·  October 2026", { x: 0.8, y: 6.4, w: 6, h: 0.4, fontSize: 14, color: C.accent3, objectName: "Byline" });
  s.addNotes("I wanted to show how I'd open with a Salesforce-heavy prospect: start from their own CRM data, find where the process leaks, and turn it into a number the CRO cares about.");

  // ---------- Opening: answer first ----------
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Opening" });
  s.addText("Illustrative annual upside: " + usd(headline) + " at 50% capture", { placeholder: "title" });
  text(s, usd(headline), { x: 0.6, y: 1.55, w: 5.4, h: 1.3, fontSize: 72, bold: true, color: C.accent1, fontFace: THEME.headFontFace, objectName: "Headline value" });
  text(s, `Gross scenario across three overlapping opportunities, before deduplication. Full gap: ${usd(I.total_annual_value)}.`, { x: 0.6, y: 2.95, w: 5.4, h: 1.0, fontSize: 16, color: C.text2, objectName: "Headline caption" });
  card(s, 0.6, 4.0, 5.6, 1.25, "Data note", C.accent3);
  text(s, `Illustrative: ${K.cases} simulated opportunities in Salesforce's object shape (${I.period.months} months). The same pipeline runs live on ${LIVE.kpis.cases} seeded opportunities in a Salesforce trial org.`, { x: 0.8, y: 4.08, w: 5.2, h: 1.1, fontSize: 13, color: C.text2, valign: "middle", objectName: "Headline basis" });

  const findings = [
    ["Accelerate bookings", `EMEA discount approvals take ${days(D.approval.baseline)} days vs ${days(D.approval.target)} in the fastest region.`],
    ["Improve win rate", `${pct(D.rework.baseline)} of negotiated deals loop back for a re-quote; they win ${rwBad}% vs ${rwGood}%.`],
    ["Sales productivity", `${pct(D.qualification.baseline)} of deals skip Qualification and win ${qBad}% vs ${qGood}%.`],
  ];
  findings.forEach(([chip, body], i) => {
    const y = 1.5 + i * 1.3;
    card(s, 6.6, y, 6.1, 1.15, `Finding ${i + 1}`);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.85, y: y + 0.15, w: 2.3, h: 0.34, rectRadius: 0.17, fill: { color: C.accent1 }, line: { type: "none" }, objectName: `Finding ${i + 1} chip` });
    text(s, chip, { x: 6.85, y: y + 0.15, w: 2.3, h: 0.34, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle", objectName: `Finding ${i + 1} label` });
    text(s, body, { x: 6.85, y: y + 0.56, w: 5.65, h: 0.5, fontSize: 14, color: C.text2, objectName: `Finding ${i + 1} text` });
  });

  const kpis = [
    [K.cases.toLocaleString(), "deals analysed"],
    [pct(K.win_rate), "win rate"],
    [`${K.median_cycle_days}d`, "median cycle time"],
    [String(K.variants), "distinct paths"],
    [pct(K.happy_path_conformance), "match the winning trace"],
  ];
  kpis.forEach(([v, l], i) => {
    const x = 0.6 + i * 2.45;
    card(s, x, 5.55, 2.3, 1.2, `KPI ${i + 1}`, C.accent3);
    text(s, v, { x: x + 0.2, y: 5.68, w: 1.9, h: 0.55, fontSize: 28, bold: true, color: C.text2, objectName: `KPI ${i + 1} value` });
    text(s, l, { x: x + 0.2, y: 6.25, w: 1.9, h: 0.35, fontSize: 12, color: C.accent6, objectName: `KPI ${i + 1} label` });
  });
  s.addNotes("Lead with the number, then the three findings. Everything after this slide is how I got there and how you would check it.");

  // ---------- Opening: approach overview ----------
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Opening" });
  s.addText("Approach: from finding value to realising it", { placeholder: "title" });
  const stages = [
    ["Identify", "Mine the CRM event log: process map, variants and bottlenecks.", "3 opportunities found", true],
    ["Frame", "Size each opportunity from its KPI baseline and illustrative assumptions to validate with process owners.", `${usd(I.total_annual_value)} gross scenario`, true],
    ["Commit", "Validate KPI definitions and values with each opportunity owner.", "Signed-off committed value", false],
    ["Realize", "Act on the root cause, then track the KPI against its baseline.", "Value realised, reported quarterly", false],
  ];
  stages.forEach(([head, body, out, done], i) => {
    const x = 0.6 + i * 3.1, w = 2.85;
    const fill = done ? C.text2 : C.background2, ink = done ? C.background1 : C.text2;
    card(s, x, 1.85, w, 3.55, `Stage ${i + 1} card`, fill);
    s.addShape(pres.shapes.OVAL, { x: x + 0.25, y: 2.1, w: 0.5, h: 0.5, fill: { color: done ? C.accent1 : C.accent3 }, line: { type: "none" }, objectName: `Stage ${i + 1} marker` });
    text(s, String(i + 1), { x: x + 0.25, y: 2.1, w: 0.5, h: 0.5, fontSize: 16, bold: true, color: done ? C.background1 : C.text2, align: "center", valign: "middle", objectName: `Stage ${i + 1} number` });
    text(s, head, { x: x + 0.25, y: 2.8, w: w - 0.5, h: 0.5, fontSize: 22, bold: true, color: ink, objectName: `Stage ${i + 1} title` });
    text(s, body, { x: x + 0.25, y: 3.35, w: w - 0.5, h: 1.2, fontSize: 14, color: ink, objectName: `Stage ${i + 1} body` });
    text(s, [{ text: "Output: ", options: { bold: true } }, { text: out }], { x: x + 0.25, y: 4.65, w: w - 0.5, h: 0.6, fontSize: 13, color: done ? C.accent3 : C.accent6, objectName: `Stage ${i + 1} output` });
    if (i < stages.length - 1) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + w + 0.03, y: 3.45, w: 0.19, h: 0.34, fill: { color: C.accent2 }, line: { type: "none" }, objectName: `Stage ${i + 1} arrow` });
  });
  // "Today" marker sits over the Frame stage.
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6 + 3.1 + 0.85, y: 1.38, w: 1.15, h: 0.36, rectRadius: 0.18, fill: { color: C.accent4 }, line: { type: "none" }, objectName: "Today pill" });
  text(s, "Today", { x: 0.6 + 3.1 + 0.85, y: 1.38, w: 1.15, h: 0.36, fontSize: 13, bold: true, color: C.background1, align: "center", valign: "middle", objectName: "Today label" });
  text(s, "This demo covers Identify and Frame on Salesforce data. Commit and Realize run with the customer's process owners.", { x: 0.6, y: 5.75, w: 12.1, h: 0.5, fontSize: 16, italic: true, color: C.accent6, objectName: "Approach note" });
  s.addNotes("Framed value isn't committed value. Today's work stops at Frame on purpose: the numbers are illustrative and the improvement assumption belongs to the opportunity owner. Commit means signing off KPI definitions and values together; Realize means tracking the KPI against the baseline on a fixed cadence.");

  // ---------- Evidence: discover ----------
  pres.addSection({ title: "Evidence" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Evidence" });
  s.addText("Discover: paths in the simulated event log", { placeholder: "title" });
  const discover = [
    [pct(K.happy_path_conformance), `match the selected winning trace (${pct(K.winning_trace_share_of_won)} of won deals)`],
    [String(K.variants), `distinct paths across ${K.cases} deals`],
    [`${days(D.approval.baseline)} days`, "median EMEA discount approval, the slowest region"],
  ];
  discover.forEach(([v, l], i) => {
    const y = 1.5 + i * 1.3;
    text(s, v, { x: 0.6, y, w: 2.9, h: 0.9, fontSize: 44, bold: true, color: C.accent1, valign: "middle", objectName: `Discover stat ${i + 1}` });
    text(s, l, { x: 3.6, y, w: 3.9, h: 0.9, fontSize: 18, color: C.text2, valign: "middle", objectName: `Discover label ${i + 1}` });
  });
  text(s, "Approval steps and re-quote loops show up as detours from the winning trace.", { x: 0.6, y: 5.35, w: 6.9, h: 0.45, fontSize: 14, italic: true, color: C.accent6, objectName: "Discover note" });
  // Highlighted, clickable link to the live dashboard.
  s.addText([
    { text: "Open the live process map  →", options: { bold: true, hyperlink: { url: `https://${SITE}/dashboard.html`, tooltip: "Open the live dashboard" } } },
  ], { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.15, x: 0.6, y: 6.0, w: 6.9, h: 0.6, fill: { color: C.accent3 }, line: { color: C.accent1, width: 1.5 },
    fontSize: 14, align: "center", valign: "middle", margin: 0.08, objectName: "Dashboard link" });
  const mapSize = require("image-size")(path.join(__dirname, "process-map.png"));
  const mapH = 5.25, mapW = mapH * (mapSize.width / mapSize.height);
  s.addShape(pres.shapes.RECTANGLE, { x: 12.7 - mapW - 0.05, y: 1.4, w: mapW + 0.1, h: mapH + 0.1, fill: { color: C.background1 }, line: { color: C.accent3, width: 1 }, objectName: "Map frame" });
  s.addImage({ path: path.join(__dirname, "process-map.png"), x: 12.7 - mapW, y: 1.45, w: mapW, h: mapH, objectName: "Process map", altText: "Approval and re-quote section of the process map" });
  s.addNotes("This is the directly-follows map from the hosted dashboard. Fourteen percent conformance is the hook: the CRM shows the stage, process mining shows the path.");

  // ---------- Evidence: diagnose ----------
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Evidence" });
  s.addText("Diagnose: three patterns associated with lower win rates", { placeholder: "title" });
  const muted = "9DC9EE", hot = HEX.accent4;
  const charts = [
    { title: "Median approval wait (days)", labels: regions.map(([r]) => r), values: regions.map(([, v]) => v.approval_wait_days),
      colors: regions.map(([, v]) => (v.approval_wait_days === Math.max(...regions.map(([, w]) => w.approval_wait_days)) ? hot : muted)), fmt: "0.0",
      take: `${(D.approval.baseline - D.approval.target).toFixed(1)} days slower than the fastest region. EMEA deals waiting >7 days won ${apWin[0]}% vs ${apWin[1]}% (n = ${D.approval.cohorts.problem} vs ${D.approval.cohorts.comparison}).` },
    { title: "Win rate of negotiated deals", labels: ["Clean path", "Re-quoted"], values: [rwGood / 100, rwBad / 100], colors: [muted, hot], fmt: "0%",
      take: `Re-quoted deals win ${rwGood - rwBad} points less often (n = ${D.rework.cohorts.problem} re-quoted vs ${D.rework.cohorts.comparison} clean).` },
    { title: "Win rate by qualification", labels: ["Qualified", "Skipped"], values: [qGood / 100, qBad / 100], colors: [muted, hot], fmt: "0%",
      take: `Deals that skip Qualification win ${qGood - qBad} points less often (n = ${D.qualification.cohorts.problem} vs ${D.qualification.cohorts.comparison}).` },
  ];
  charts.forEach((c, i) => {
    const x = 0.6 + i * 4.12;
    s.addChart(pres.charts.BAR, [{ name: c.title, labels: c.labels, values: c.values }], {
      x, y: 1.45, w: 3.86, h: 3.7, barDir: "col", objectName: `Chart ${i + 1}`,
      showTitle: true, title: c.title, titleFontSize: 14, titleColor: HEX.dk2, titleFontFace: "+mn-lt",
      chartColors: c.colors, barGapWidthPct: 60,
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: c.fmt, dataLabelFontSize: 14, dataLabelColor: HEX.dk2, dataLabelFontFace: "+mn-lt",
      catAxisLabelColor: HEX.dk2, catAxisLabelFontSize: 12, catAxisLabelFontFace: "+mn-lt",
      valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
      valAxisMinVal: 0, valAxisMaxVal: c.fmt === "0%" ? 0.7 : 8.5, showLegend: false,
    });
    card(s, x, 5.3, 3.86, 1.35, `Takeaway ${i + 1}`);
    text(s, c.take, { x: x + 0.2, y: 5.42, w: 3.46, h: 1.1, fontSize: 14, color: C.text2, valign: "middle", objectName: `Takeaway ${i + 1} text` });
  });
  s.addNotes("Orange marks the problem cohort in each chart. These are associations, not proven causes: they don't control for deal size, discount, region or rep. Win-rate gaps compare like with like, for example only deals that reached Negotiation for re-quotes.");

  // ---------- Value: business case overview ----------
  pres.addSection({ title: "Value" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Value" });
  s.addText("Business case overview: three opportunities, one deep-dive", { placeholder: "title" });
  const hdr = t => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 } } });
  const opps = [
    ["Accelerate bookings", "1  EMEA discount approvals", "Deal desk lead", "Short", D.approval, false],
    ["Improve win rate", "2  Re-quotes (deep-dive)", "VP Revenue Ops", "Medium", D.rework, true],
    ["Sales productivity", "3  Skipped qualification", "Sales operations", "Short", D.qualification, false],
  ];
  const cell = (t, focus, extra = {}) => ({ text: t, options: { fill: { color: focus ? C.accent3 : C.background1 }, bold: focus, ...extra } });
  const rows = [
    ["Objective", "Opportunity", "Proposed owner", "Time to value", "Full gap", "Committed", "Realized"].map(hdr),
    ...opps.map(([obj, opp, owner, ttv, d, focus]) => [
      cell(obj, focus), cell(opp, focus), cell(owner, focus), cell(ttv, focus),
      cell(usd2(d.annual_value), focus, { align: "right" }), cell("TBD", focus, { align: "right", color: C.accent6 }), cell("—", focus, { align: "right", color: C.accent6 }),
    ]),
    [cell("Total (gross)", false, { bold: true }), cell(""), cell(""), cell(""), cell(usd2(I.total_annual_value), false, { bold: true, align: "right" }), cell("TBD", false, { align: "right", color: C.accent6 }), cell("—", false, { align: "right", color: C.accent6 })],
  ];
  s.addTable(rows, { x: 0.6, y: 1.45, w: 12.1, colW: [2.05, 2.95, 1.9, 1.3, 1.2, 1.45, 1.25], fontSize: 14, color: C.text1, valign: "middle", rowH: 0.48,
    border: { type: "solid", pt: 0.75, color: HEX.accent3 }, margin: [0.05, 0.12, 0.05, 0.12], objectName: "Business case table" });
  text(s, "Full gap = illustrative scenario from simulated Salesforce data. Committed is agreed with each owner after validation; realized is tracked against a frozen baseline.", { x: 0.6, y: 4.08, w: 12.1, h: 0.35, fontSize: 12, italic: true, color: C.accent6, objectName: "Table note" });

  // Before any value is committed: what still has to be true.
  const ov = I.overlaps;
  card(s, 0.6, 4.6, 6.0, 1.85, "Overlap card");
  text(s, [
    { text: "Opportunities overlap", options: { bold: true, color: C.text2, breakLine: true } },
    { text: `${ov.rework_and_qualification} deals were re-quoted and skipped Qualification, ${ov.rework_and_approval} were re-quoted with slow approvals, ${ov.qualification_and_approval} skipped Qualification with slow approvals.`, options: { breakLine: true } },
    { text: " ", options: { breakLine: true, fontSize: 6 } },
    { text: "So the total is gross: count each incremental deal win once across all three improvement opportunities before committing.", options: { italic: true, color: C.accent6 } },
  ], { x: 0.85, y: 4.72, w: 5.5, h: 1.95, fontSize: 14, color: C.text1, objectName: "Overlap text" });
  text(s, "Before value is committed", { x: 6.95, y: 4.6, w: 5.75, h: 0.35, fontSize: 14, bold: true, color: C.text2, objectName: "Checklist heading" });
  const checks = [
    "Deduplicate overlapping deals across opportunities",
    "Agree cohort, period and Amount basis",
    "Owner signs off KPI definition and improvement assumption",
    "Finance signs off the value formula",
  ];
  checks.forEach((c, i) => {
    const y = 5.02 + i * 0.43;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.95, y: y + 0.06, w: 0.24, h: 0.24, rectRadius: 0.04, fill: { color: C.background1 }, line: { color: C.accent1, width: 1.5 }, objectName: `Check ${i + 1} box` });
    text(s, c, { x: 7.32, y, w: 5.4, h: 0.38, fontSize: 14, color: C.text1, valign: "middle", objectName: `Check ${i + 1} text` });
  });
  s.addNotes("Portfolio view first, then one deep-dive. Full-gap values are illustrative and gross; the overlap counts are real and show why the total can't simply be added. Owners are proposed, not committed. Committed stays TBD until the owner and Finance sign off.");

  // ---------- Value: value case ----------
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Value" });
  s.addText("Value case: reduce re-quotes in quote-to-close", { placeholder: "title" });
  card(s, 0.6, 1.35, 12.1, 0.45, "Filter bar");
  text(s, `${I.period.from.slice(0, 7)} to ${I.period.to.slice(0, 7)}, annualised  ·  Reached Negotiation/Review  ·  Amount = bookings  ·  Proposed owner: VP Revenue Ops  ·  Simulated data`, { x: 0.85, y: 1.35, w: 11.6, h: 0.45, fontSize: 12, color: C.text2, valign: "middle", objectName: "Filter text" });
  const cols = [
    ["Challenge", [`${Math.round(negotiated)} negotiated deals a year; ${pct(D.rework.baseline)} loop back for a re-quote`, `Re-quoted deals win ${rwBad}% vs ${rwGood}%`, `EMEA approvals take ${days(D.approval.baseline)} days vs ${days(D.approval.target)}`]],
    ["Root cause (to capture)", ["Reason codes: quote expired, discount, scope, pricing error", "Hotspots by approver queue, region and rep", "Test whether approval delay causes expiry and re-quotes"]],
    ["Solution with process intelligence", ["Performance views show where quotes wait and loop", "Conformance checking flags skipped steps", "Simulate a pilot before any policy change"]],
  ];
  cols.forEach(([head, items], i) => {
    const x = 0.6 + i * 4.12;
    text(s, head, { x, y: 2.0, w: 3.86, h: 0.4, fontSize: 16, bold: true, color: C.accent6, objectName: `Value column ${i + 1} heading` });
    text(s, items.map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < items.length - 1 } })), { x, y: 2.45, w: 3.86, h: 2.0, fontSize: 14, color: C.text1, paraSpaceAfter: 6, objectName: `Value column ${i + 1} body` });
  });
  const chain = [
    ["Negotiated deals / yr", Math.round(negotiated).toLocaleString()],
    ["Share re-quoted", `${(D.rework.baseline * 100).toFixed(1)}%`],
    ["Win-rate gap", `${(D.rework.win_rate_lift * 100).toFixed(1)} pts`],
    ["Average won deal", `$${(K.avg_won_deal / 1000).toFixed(1)}k`],
    ["Re-quotes prevented", "50%"],
  ];
  chain.forEach(([l, v], i) => {
    const x = 0.6 + i * 1.95;
    text(s, l, { x, y: 4.65, w: 1.65, h: 0.45, fontSize: 12, bold: true, color: C.text2, align: "center", valign: "bottom", objectName: `Chain ${i + 1} label` });
    card(s, x, 5.15, 1.65, 0.85, `Chain ${i + 1} box`);
    text(s, v, { x, y: 5.15, w: 1.65, h: 0.85, fontSize: 22, bold: true, color: C.text2, align: "center", valign: "middle", objectName: `Chain ${i + 1} value` });
    text(s, i < chain.length - 1 ? "×" : "=", { x: x + 1.65, y: 5.15, w: 0.3, h: 0.85, fontSize: 20, color: C.accent6, align: "center", valign: "middle", objectName: `Chain ${i + 1} operator` });
  });
  card(s, 10.4, 4.65, 2.3, 1.35, "Framed value pill", C.text2);
  text(s, [
    { text: `≈ ${usd(framed)}`, options: { fontSize: 30, bold: true, color: C.background1, breakLine: true } },
    { text: "modelled incremental bookings p.a.", options: { fontSize: 11, color: C.accent3 } },
  ], { x: 10.4, y: 4.65, w: 2.3, h: 1.35, align: "center", valign: "middle", objectName: "Framed value" });
  text(s, [
    { text: "Scenario to test: ", options: { bold: true, color: C.text2 } },
    { text: `cut re-quotes from ${(D.rework.baseline * 100).toFixed(1)}% to ${(D.rework.baseline * 50).toFixed(1)}% of negotiated deals, assuming each prevented re-quote recovers the ${(D.rework.win_rate_lift * 100).toFixed(1)}-point win-rate gap. Measured against a comparison group; Finance validates incremental bookings, revenue recognition assessed separately.` },
  ], { x: 0.6, y: 6.12, w: 12.1, h: 0.75, fontSize: 14, color: C.text1, objectName: "Scenario" });
  s.addNotes(`This follows my value framing format: challenge, root cause, solution, then a multiplicative value chain. The customer owns the improvement assumption; we own the evidence. Root cause is labelled 'to capture' because it needs reason codes the CRM doesn't hold yet. The chain multiplies to the displayed value using unrounded inputs. Annualised over the ${I.period.months}-month event window as an illustrative run rate; the win-rate gap is an association among deals that reached Negotiation/Review.`);

  }

  if (KIND === "tech") {
  // ---------- Tech setup: title ----------
  pres.addSection({ title: "Tech setup" });
  s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Tech setup" });
  circle(s, 9.4, 3.4, 6.6, C.accent1, 70, "Decor circle large");
  circle(s, 10.6, 2.6, 0.75, C.accent3, 0, "Decor circle small");
  s.addText("Tech setup", { placeholder: "title" });
  s.addText("The pipeline behind the demo: Salesforce data to an event log", { placeholder: "body" });
  text(s, "Luke Audie  ·  October 2026", { x: 0.8, y: 6.4, w: 6, h: 0.4, fontSize: 14, color: C.accent3, objectName: "Byline" });
  s.addNotes("Companion to the main deck for a technical audience: how the data gets out of Salesforce and where to see it running.");

  // ---------- Tech setup: evidence ----------
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Tech setup" });
  s.addText("From a Salesforce org to an event log", { placeholder: "title" });

  // Left: the org itself (screenshot with its URL), seeded by seed.py.
  const orgShot = fs.existsSync(path.join(ROOT, "docs/img/org/salesforce-with-url.jpg"))
    ? { file: "salesforce-with-url.jpg" } : { file: "opportunities.jpg" };
  const { width: ow, height: oh } = require("image-size")(path.join(ROOT, "docs/img/org", orgShot.file));
  const boxW = 5.55, boxH = 4.7, r = ow / oh;
  const sw = Math.min(boxW, boxH * r), sh = sw / r;
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6 - 0.04, y: 1.4 - 0.04, w: sw + 0.08, h: sh + 0.08, fill: { color: C.background1 }, line: { color: C.accent3, width: 1 }, objectName: "Org screenshot frame" });
  s.addImage({ path: path.join(ROOT, "docs/img/org", orgShot.file), x: 0.6, y: 1.4, w: sw, h: sh, objectName: "Org screenshot", altText: "Salesforce trial org with seeded demo records" });
  text(s, [
    { text: "Salesforce trial org: ", options: { bold: true } },
    { text: `seed.py loaded ${LIVE.kpis.cases} [Demo] accounts and opportunities to simulate a client's CRM.` },
  ], { x: 0.6, y: 6.2, w: 5.55, h: 0.75, fontSize: 14, color: C.text2, objectName: "Org caption" });

  // Right: one query from extract.py, readable, plus how each object becomes events.
  const src = fs.readFileSync(path.join(ROOT, "pipeline/extract.py"), "utf8").split("\n");
  const q0 = src.findIndex(l => l.includes('"stage_history"'));
  const q1 = src.findIndex((l, i) => i > q0 && l.includes('""",'));
  const codeLines = src.slice(q0 + 1, q1 + 1).map(l => l.trim().replace(/""",?$/, ""));
  const KW = /\b(SELECT|FROM|WHERE|ORDER BY|AND)\b/;
  const runs = [];
  codeLines.forEach((line, li) => {
    const parts = line.split(/(\b(?:SELECT|FROM|WHERE|ORDER BY|AND)\b)/).filter(Boolean);
    parts.forEach((part, pi) => {
      runs.push({ text: (pi === 0 && !KW.test(part) ? "       " : "") + part, options: { color: KW.test(part) ? HEX.accent1 : HEX.accent3, bold: KW.test(part), breakLine: pi === parts.length - 1 && li < codeLines.length - 1 } });
    });
  });
  const px = 6.45, pw = 6.25;
  card(s, px, 1.4, pw, 2.25, "Code panel", C.text2);
  [HEX.accent4, "F6C344", HEX.accent5].forEach((dot, i) =>
    s.addShape(pres.shapes.OVAL, { x: px + 0.2 + i * 0.2, y: 1.55, w: 0.12, h: 0.12, fill: { color: dot }, line: { type: "none" }, objectName: `Code panel dot ${i + 1}` }));
  text(s, "pipeline/extract.py · stage history", { x: px + 0.85, y: 1.48, w: 3.4, h: 0.26, fontSize: 11, color: C.accent3, objectName: "Code panel filename" });
  text(s, runs, { x: px + 0.25, y: 1.9, w: pw - 0.4, h: 1.6, fontSize: 14, fontFace: "Courier New", lineSpacingMultiple: 1.1, objectName: "SOQL code" });

  const hdr2 = t => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 } } });
  const mapRows = [
    ["Opportunity", "Case: Id, Amount, IsWon, region"],
    ["OpportunityHistory", "Stage events: StageName, CreatedDate"],
    ["Task (WhatId)", "Activities: Subject, created → completed"],
    ["Lead (ConvertedOpportunityId)", "Lead Created event"],
  ].map(([o, e], i) => [o, e].map(t => ({ text: t, options: { fill: { color: i % 2 ? C.background1 : C.background2 } } })));
  s.addTable([[hdr2("Salesforce object"), hdr2("Becomes in the event log")], ...mapRows], {
    x: px, y: 3.85, w: pw, colW: [2.75, 3.5], fontSize: 13, color: C.text1, valign: "middle", rowH: 0.42,
    border: { type: "solid", pt: 0.75, color: HEX.accent3 }, margin: [0.04, 0.1, 0.04, 0.1], objectName: "Object to event mapping",
  });
  text(s, [
    { text: "extract.py: ", options: { bold: true } },
    { text: "four SOQL queries over the REST API, signed by the Salesforce CLI session. Output: CSV and XES." },
  ], { x: px, y: 6.2, w: pw, h: 0.75, fontSize: 14, color: C.text2, objectName: "Code caption" });
  // Numbered popups walk the viewer through the flow, left to right.
  const popup = (n, x, y, w, label) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.4, rectRadius: 0.08, fill: { color: C.background1 }, line: { color: C.accent1, width: 1.5 },
      shadow: { type: "outer", color: "000000", opacity: 0.25, blur: 4, offset: 2, angle: 90 }, objectName: `Popup ${n} box` });
    s.addShape(pres.shapes.OVAL, { x: x + 0.06, y: y + 0.05, w: 0.3, h: 0.3, fill: { color: C.accent4 }, line: { type: "none" }, objectName: `Popup ${n} badge` });
    text(s, String(n), { x: x + 0.06, y: y + 0.05, w: 0.3, h: 0.3, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle", objectName: `Popup ${n} number` });
    text(s, label, { x: x + 0.44, y, w: w - 0.5, h: 0.4, fontSize: 12, bold: true, color: C.text2, valign: "middle", objectName: `Popup ${n} label` });
  };
  popup(1, 3.15, 1.6, 2.85, "Trial org = the client's CRM");
  popup(2, 3.3, 3.85, 2.7, `seed.py: ${LIVE.kpis.cases} [Demo] records`);
  popup(3, px + pw - 2.75, 1.43, 2.65, "SOQL via the CLI session");
  popup(4, px + pw - 2.45, 3.43, 2.35, "Objects → events");
  s.addNotes("Left: the trial org standing in for a client's CRM, seeded by seed.py so there's real OpportunityHistory to read. Right: the stage-history query and how each object maps to events; Opportunity Id is the case. Known limits: Task created → completed can include queueing time; open tasks have no completion time and show as zero duration; history can't be backdated in a trial org, so the 10-month dataset is simulated in the same shape. The export produces CSV and XES; an import into a process mining tool is the next step to verify. Production work would add incremental refresh, stable event IDs and reconciliation.");

  // ---------- Tech setup: everything is live ----------
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Tech setup" });
  s.addText("Hosted demo, live extraction proof and source code", { placeholder: "title" });
  const proof = [
    ["Overview: value items and business case", SITE, `https://${SITE}`],
    ["Process analysis", `${SITE}/dashboard.html`, `https://${SITE}/dashboard.html`],
    ["Business case (re-quotes)", `${SITE}/#business-case`, `https://${SITE}/#business-case`],
    ["Data assets and downloads", `${SITE}/#live-run`, `https://${SITE}/#live-run`],
    ["Source code and event logs", REPO, `https://${REPO}`],
  ];
  proof.forEach(([label, shown, url], i) => {
    const y = 1.5 + i * 1.0;
    card(s, 0.6, y, 6.6, 0.88, `Proof ${i + 1}`);
    s.addShape(pres.shapes.OVAL, { x: 0.82, y: y + 0.2, w: 0.45, h: 0.45, fill: { color: C.accent1 }, line: { type: "none" }, objectName: `Proof ${i + 1} marker` });
    text(s, "→", { x: 0.82, y: y + 0.2, w: 0.45, h: 0.45, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle", objectName: `Proof ${i + 1} glyph` });
    text(s, label, { x: 1.45, y: y + 0.08, w: 5.6, h: 0.36, fontSize: 16, bold: true, color: C.text2, objectName: `Proof ${i + 1} label` });
    text(s, url ? [{ text: shown, options: { hyperlink: { url } } }] : shown, { x: 1.45, y: y + 0.46, w: 5.6, h: 0.34, fontSize: 14, color: C.accent6, objectName: `Proof ${i + 1} link` });
  });
  const ovSize = require("image-size")(path.join(ROOT, "docs/img/overview-top.png"));
  const shotW = 5.1, shotH = shotW * (ovSize.height / ovSize.width), shotY = 1.5 + (4.88 - shotH) / 2;
  s.addShape(pres.shapes.RECTANGLE, { x: 7.6 - 0.05, y: shotY - 0.05, w: shotW + 0.1, h: shotH + 0.1, fill: { color: C.background1 }, line: { color: C.accent3, width: 1 }, objectName: "Screenshot frame" });
  s.addImage({ path: path.join(ROOT, "docs/img/overview-top.png"), x: 7.6, y: shotY, w: shotW, h: shotH, objectName: "Overview screenshot", altText: "Overview page of the hosted site" });
  s.addNotes("All of this is hosted. The overview carries the value items and business case; Data assets has the live run and every event log; the repo has the pipeline, the seeding script and the README with the setup gotchas from a new trial org.");

  }

  if (KIND === "main") {
  // ---------- Close: next steps ----------
  pres.addSection({ title: "Close" });
  s = pres.addSlide({ masterName: "CLOSING_DARK", sectionTitle: "Close" });
  circle(s, 10.2, 4.4, 5.4, C.accent1, 75, "Decor circle large");
  s.addText("Turning the finding into a decision", { placeholder: "title" });
  const next = [
    ["Validate", "Run the extract on the production org; confirm cohorts, causes and overlap.", "RevOps + SE · weeks 1–2"],
    ["Pilot", "Agree a controlled pilot with RevOps and Finance once re-quote causes are validated.", "RevOps + Finance · weeks 3–8"],
    ["Measure", "Track share re-quoted and win rate against a frozen baseline and a comparison group.", "Finance · monthly, signed off quarterly"],
  ];
  next.forEach(([head, body, owner], i) => {
    const x = 0.8 + i * 3.95;
    card(s, x, 2.0, 3.65, 2.65, `Next ${i + 1} card`, C.accent6);
    text(s, head, { x: x + 0.3, y: 2.3, w: 3.05, h: 0.5, fontSize: 22, bold: true, color: C.accent3, objectName: `Next ${i + 1} heading` });
    text(s, body, { x: x + 0.3, y: 2.95, w: 3.05, h: 1.3, fontSize: 16, color: C.background1, objectName: `Next ${i + 1} body` });
    text(s, owner, { x: x + 0.3, y: 4.0, w: 3.05, h: 0.4, fontSize: 12, color: C.accent3, objectName: `Next ${i + 1} owner` });
  });
  text(s, "Illustrative schedule; owners are proposed.", { x: 0.8, y: 4.75, w: 8, h: 0.3, fontSize: 12, italic: true, color: C.accent3, objectName: "Schedule note" });
  text(s, [
    { text: "Decision today: ", options: { bold: true, color: C.background1 } },
    { text: "agree the validation scope, named owners and pilot success criteria.", options: { color: C.accent3 } },
  ], { x: 0.8, y: 5.25, w: 9.2, h: 0.6, fontSize: 18, objectName: "Decision" });
  text(s, `${SITE}  ·  Luke Audie`, { x: 0.8, y: 6.5, w: 8, h: 0.4, fontSize: 14, color: C.accent3, objectName: "Closing contact" });
  logo(s, 10.35, 6.35, 0.55);
  s.addNotes("Close on the decision: validate on production data, run a controlled pilot only after the causes are confirmed, then measure against a frozen baseline with Finance sign-off.");

  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log(`Wrote ${OUT}${hasLogo ? " (with logo)" : ""}`);
})().catch(e => { console.error(e); process.exit(1); });
