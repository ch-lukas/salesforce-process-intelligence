// Builds deck/salesforce-process-intelligence.pptx from data/insights.json.
// Colour theme follows the Salesforce CRM handbook: navy, cloud blue, light-blue panels.
// Usage: node deck/build_deck.js   (logo picked up from deck/logo-white.svg if present)
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");

const SKILL = process.env.PPTX_SKILL_DIR;
const { applyTheme } = require(path.join(SKILL, "scripts/apply_theme.js"));

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
  name: "Salesforce Handbook",
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
const regions = Object.entries(I.by_region).filter(([, v]) => v.approval_wait_days != null);
const negotiated = Math.round(D.rework.affected_deals_per_year / D.rework.baseline);
const framed = negotiated * D.rework.baseline * D.rework.win_rate_lift * K.avg_won_deal * 0.5;
const usd = v => (v >= 1e6 ? `$${(v / 1e6).toFixed(1)}M` : `$${Math.round(v / 1e3)}k`);
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
    slide.addImage({ path: LOGO, x, y, h, w: h * (vw / vh), objectName: "Logo", altText: "Apromore logo" });
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
  s.addText("Lead-to-close process intelligence: a working demo build", { placeholder: "body" });
  text(s, "Luke Audie  ·  October 2026", { x: 0.8, y: 6.4, w: 6, h: 0.4, fontSize: 14, color: C.accent3, objectName: "Byline" });
  s.addNotes("I wanted to show how I'd open with a Salesforce-heavy prospect: start from their own CRM data, find where the process leaks, and turn it into a number the CRO cares about.");

  // ---------- Opening: answer first ----------
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Opening" });
  s.addText("The answer first: three process fixes worth " + usd(headline) + " a year", { placeholder: "title" });
  text(s, usd(headline), { x: 0.6, y: 1.55, w: 5.4, h: 1.3, fontSize: 72, bold: true, color: C.accent1, fontFace: THEME.headFontFace, objectName: "Headline value" });
  text(s, `Annual revenue upside if half of the gap is closed on three issues found in the CRM data. Full gap: ${usd(I.total_annual_value)}.`, { x: 0.6, y: 2.95, w: 5.4, h: 1.0, fontSize: 16, color: C.text2, objectName: "Headline caption" });
  text(s, `Sized from ${K.cases} opportunities over ${I.period.months} months.`, { x: 0.6, y: 4.05, w: 5.4, h: 0.4, fontSize: 14, italic: true, color: C.accent6, objectName: "Headline basis" });

  const findings = [
    ["Accelerate revenue", `EMEA discount approvals take ${days(D.approval.baseline)} days vs ${days(D.approval.target)} in the fastest region.`],
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
    [pct(K.happy_path_conformance), "follow the happy path"],
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
    ["Frame", "Size each opportunity from its KPI baseline and customer-owned assumptions.", `${usd(I.total_annual_value)} framed`, true],
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
  s.addNotes("Framed value isn't committed value. Today's work stops at Frame on purpose: the numbers are evidence-backed but the improvement assumption belongs to the opportunity owner. Commit means signing off KPI definitions and values together; Realize means tracking the KPI against the baseline on a fixed cadence.");

  // ---------- Evidence: discover ----------
  pres.addSection({ title: "Evidence" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Evidence" });
  s.addText("Discover: the process as it actually runs", { placeholder: "title" });
  const discover = [
    [pct(K.happy_path_conformance), "of deals follow the designed happy path"],
    [String(K.variants), `distinct paths across ${K.cases} deals`],
    [`${days(D.approval.baseline)} days`, "median EMEA approval wait, the slowest hand-off"],
  ];
  discover.forEach(([v, l], i) => {
    const y = 1.5 + i * 1.3;
    text(s, v, { x: 0.6, y, w: 2.9, h: 0.9, fontSize: 44, bold: true, color: C.accent1, valign: "middle", objectName: `Discover stat ${i + 1}` });
    text(s, l, { x: 3.6, y, w: 3.9, h: 0.9, fontSize: 18, color: C.text2, valign: "middle", objectName: `Discover label ${i + 1}` });
  });
  text(s, "Approval steps and re-quote loops show up as detours from the designed path.", { x: 0.6, y: 5.35, w: 6.9, h: 0.45, fontSize: 14, italic: true, color: C.accent6, objectName: "Discover note" });
  // Highlighted, clickable link to the live dashboard.
  s.addText([
    { text: "Open the live process map  →", options: { bold: true, hyperlink: { url: `https://${SITE}/dashboard.html`, tooltip: "Open the live dashboard" } } },
  ], { shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.15, x: 0.6, y: 6.0, w: 6.9, h: 0.6, fill: { color: C.accent3 }, line: { color: C.accent1, width: 1.5 },
    fontSize: 14, align: "center", valign: "middle", margin: 0.08, objectName: "Dashboard link" });
  const mapH = 5.25, mapW = mapH * (1080 / 1250);
  s.addShape(pres.shapes.RECTANGLE, { x: 12.7 - mapW - 0.05, y: 1.4, w: mapW + 0.1, h: mapH + 0.1, fill: { color: C.background1 }, line: { color: C.accent3, width: 1 }, objectName: "Map frame" });
  s.addImage({ path: path.join(__dirname, "process-map.png"), x: 12.7 - mapW, y: 1.45, w: mapW, h: mapH, objectName: "Process map", altText: "Process map from the hosted dashboard" });
  s.addNotes("This is the directly-follows map from the hosted dashboard. Fourteen percent conformance is the hook: the CRM shows the stage, process mining shows the path.");

  // ---------- Evidence: diagnose ----------
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Evidence" });
  s.addText("Diagnose: three deviations explain most of the lost value", { placeholder: "title" });
  const muted = "9DC9EE", hot = HEX.accent4;
  const charts = [
    { title: "Median approval wait (days)", labels: regions.map(([r]) => r), values: regions.map(([, v]) => v.approval_wait_days),
      colors: regions.map(([, v]) => (v.approval_wait_days === Math.max(...regions.map(([, w]) => w.approval_wait_days)) ? hot : muted)), fmt: "0.0",
      take: `EMEA's single approver adds about ${(D.approval.baseline - D.approval.target).toFixed(0)} days per discounted deal.` },
    { title: "Win rate of negotiated deals", labels: ["Clean path", "Re-quoted"], values: [rwGood / 100, rwBad / 100], colors: [muted, hot], fmt: "0%",
      take: `${pct(D.rework.baseline)} of negotiated deals are re-quoted, and they win ${rwGood - rwBad} points less often.` },
    { title: "Win rate by qualification", labels: ["Qualified", "Skipped"], values: [qGood / 100, qBad / 100], colors: [muted, hot], fmt: "0%",
      take: `${pct(D.qualification.baseline)} of deals skip Qualification and win less than half as often.` },
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
  s.addNotes("Orange marks the problem cohort in each chart. Win-rate gaps compare like with like, for example only deals that reached Negotiation for re-quotes.");

  // ---------- Value: business case overview ----------
  pres.addSection({ title: "Value" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Value" });
  s.addText("Business case overview: three opportunities, one deep-dive", { placeholder: "title" });
  const hdr = t => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 } } });
  const opps = [
    ["Accelerate revenue", "1  EMEA discount approvals", "Deal desk lead", "Short", D.approval, false],
    ["Improve win rate", "2  Re-quotes (deep-dive)", "VP Revenue Ops", "Medium", D.rework, true],
    ["Sales productivity", "3  Skipped qualification", "Sales operations", "Short", D.qualification, false],
  ];
  const cell = (t, focus, extra = {}) => ({ text: t, options: { fill: { color: focus ? C.accent3 : C.background1 }, bold: focus, ...extra } });
  const rows = [
    ["Objective", "Opportunity", "Owner", "Time to value", "Framed", "Committed", "Realized"].map(hdr),
    ...opps.map(([obj, opp, owner, ttv, d, focus]) => [
      cell(obj, focus), cell(opp, focus), cell(owner, focus), cell(ttv, focus),
      cell(usd(d.annual_value), focus, { align: "right" }), cell("TBD", focus, { align: "right", color: C.accent6 }), cell("—", focus, { align: "right", color: C.accent6 }),
    ]),
    [cell("Total", false, { bold: true }), cell(""), cell(""), cell(""), cell(usd(I.total_annual_value), false, { bold: true, align: "right" }), cell("TBD", false, { align: "right", color: C.accent6 }), cell("—", false, { align: "right", color: C.accent6 })],
  ];
  s.addTable(rows, { x: 0.6, y: 1.45, w: 12.1, colW: [2.05, 2.95, 1.9, 1.3, 1.2, 1.45, 1.25], fontSize: 14, color: C.text1, valign: "middle", rowH: 0.48,
    border: { type: "solid", pt: 0.75, color: HEX.accent3 }, margin: [0.05, 0.12, 0.05, 0.12], objectName: "Business case table" });
  text(s, "Framed = full gap sized from Salesforce data. Committed is agreed with each owner after validation; realized is tracked against the baseline.", { x: 0.6, y: 4.08, w: 12.1, h: 0.35, fontSize: 12, italic: true, color: C.accent6, objectName: "Table note" });

  // Value vs time-to-value grid: where each opportunity sits.
  const gx = 0.6, gy = 4.55, gw = 5.4, gh = 2.05;
  card(s, gx, gy, gw, gh, "Priority grid");
  s.addShape(pres.shapes.LINE, { x: gx + gw / 2, y: gy + 0.15, w: 0, h: gh - 0.3, line: { color: C.accent3, width: 1 }, objectName: "Grid vertical" });
  s.addShape(pres.shapes.LINE, { x: gx + 0.15, y: gy + gh / 2, w: gw - 0.3, h: 0, line: { color: C.accent3, width: 1 }, objectName: "Grid horizontal" });
  text(s, "Quick wins", { x: gx + 0.15, y: gy + 0.1, w: 2, h: 0.3, fontSize: 11, bold: true, color: C.accent6, objectName: "Grid quick wins" });
  text(s, "Strategic", { x: gx + gw / 2 + 0.15, y: gy + 0.1, w: 2, h: 0.3, fontSize: 11, bold: true, color: C.accent6, objectName: "Grid strategic" });
  text(s, "Faster  ←  time to value  →  slower", { x: gx, y: gy + gh + 0.02, w: gw, h: 0.3, fontSize: 11, color: C.accent6, align: "center", objectName: "Grid x axis" });
  text(s, "Higher value ↑", { x: gx + gw + 0.1, y: gy + 0.1, w: 1.5, h: 0.3, fontSize: 11, color: C.accent6, objectName: "Grid y axis" });
  const dots = [["1", 0.25, 0.75, false], ["2", 0.72, 0.22, true], ["3", 0.3, 0.25, false]]; // [label, x frac, y frac, focus]
  dots.forEach(([n, fx, fy, focus], i) => {
    const d = focus ? 0.5 : 0.42;
    s.addShape(pres.shapes.OVAL, { x: gx + fx * gw - d / 2, y: gy + fy * gh, w: d, h: d, fill: { color: focus ? C.accent4 : C.accent1 }, line: { type: "none" }, objectName: `Grid dot ${n}` });
    text(s, n, { x: gx + fx * gw - d / 2, y: gy + fy * gh, w: d, h: d, fontSize: 14, bold: true, color: C.background1, align: "center", valign: "middle", objectName: `Grid dot ${n} label` });
  });
  text(s, [
    { text: "Why re-quotes first", options: { bold: true, color: C.text2, breakLine: true } },
    { text: "Largest framed value, visible directly in OpportunityHistory, and a clear owner in RevOps.", options: { breakLine: true } },
    { text: " ", options: { breakLine: true, fontSize: 6 } },
    { text: "Quick wins run alongside", options: { bold: true, color: C.text2, breakLine: true } },
    { text: "Delegated EMEA approvals (1) and a required Qualification gate (3) are policy changes, not projects." },
  ], { x: 7.6, y: 4.55, w: 5.1, h: 2.05, fontSize: 14, color: C.text1, objectName: "Priority rationale" });
  s.addNotes("Portfolio view first, then one deep-dive. Framed values come straight from the data; Committed stays TBD until the owner signs off the KPI definition and the improvement assumption. Numbers on the grid match the table rows.");

  // ---------- Value: value case ----------
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Value" });
  s.addText("Value case: reduce re-quotes in quote-to-close", { placeholder: "title" });
  card(s, 0.6, 1.35, 12.1, 0.45, "Filter bar");
  text(s, `Time frame: ${I.period.from.slice(0, 7)} to ${I.period.to.slice(0, 7)}   ·   New business   ·   Deals that reached Negotiation   ·   Owner: VP Revenue Operations`, { x: 0.85, y: 1.35, w: 11.6, h: 0.45, fontSize: 12, color: C.text2, valign: "middle", objectName: "Filter text" });
  const cols = [
    ["Challenge", [`${negotiated} negotiated deals a year; ${pct(D.rework.baseline)} loop back for a re-quote`, `Re-quoted deals win ${rwBad}% vs ${rwGood}%`, `EMEA approvals take ${days(D.approval.baseline)} days vs ${days(D.approval.target)}`]],
    ["Root cause (to capture)", ["Reason codes: quote expired, discount, scope, pricing error", "Hotspots by approver queue, region and rep", "Confirm: slow approval → expired quote → re-quote"]],
    ["Solution with Apromore", ["Performance views show where quotes wait and loop", "Conformance checking flags skipped steps", "Simulate a delegated EMEA approval threshold first"]],
  ];
  cols.forEach(([head, items], i) => {
    const x = 0.6 + i * 4.12;
    text(s, head, { x, y: 2.0, w: 3.86, h: 0.4, fontSize: 16, bold: true, color: C.accent6, objectName: `Value column ${i + 1} heading` });
    text(s, items.map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < items.length - 1 } })), { x, y: 2.45, w: 3.86, h: 2.0, fontSize: 14, color: C.text1, paraSpaceAfter: 6, objectName: `Value column ${i + 1} body` });
  });
  const chain = [
    ["Negotiated deals / yr", negotiated.toLocaleString()],
    ["Re-quotes per deal", D.rework.baseline.toFixed(2)],
    ["Win-rate gap", `${(D.rework.win_rate_lift * 100).toFixed(1)} pts`],
    ["Average won deal", `$${Math.round(K.avg_won_deal / 1000)}k`],
    ["Improvement", "50%"],
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
    { text: "revenue upside p.a.", options: { fontSize: 12, color: C.accent3 } },
  ], { x: 10.4, y: 4.65, w: 2.3, h: 1.35, align: "center", valign: "middle", objectName: "Framed value" });
  text(s, "Assumptions: annualised from 10 months of history; win-rate gap compares only deals that reached Negotiation; cycle-time gains not counted; half of the gap closed.", { x: 0.6, y: 6.2, w: 12.1, h: 0.5, fontSize: 11, italic: true, color: C.accent6, objectName: "Assumptions" });
  s.addNotes("This follows my value framing format: challenge, root cause, solution, then a multiplicative value chain. The customer owns the improvement assumption; we own the evidence. Root cause is labelled 'to capture' because it needs reason codes the CRM doesn't hold yet.");

  }

  if (KIND === "tech") {
  // ---------- Tech setup: title ----------
  pres.addSection({ title: "Tech setup" });
  s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Tech setup" });
  circle(s, 9.4, 3.4, 6.6, C.accent1, 70, "Decor circle large");
  circle(s, 10.6, 2.6, 0.75, C.accent3, 0, "Decor circle small");
  s.addText("Tech setup", { placeholder: "title" });
  s.addText("From a Salesforce org to Process Intelligence: the pipeline behind the demo", { placeholder: "body" });
  text(s, "Luke Audie  ·  October 2026", { x: 0.8, y: 6.4, w: 6, h: 0.4, fontSize: 14, color: C.accent3, objectName: "Byline" });
  s.addNotes("Companion to the main deck for a technical audience: how the data gets out of Salesforce and where to see it running.");

  // ---------- Tech setup: evidence ----------
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Tech setup" });
  s.addText("Tech setup: from a Salesforce org to an event log", { placeholder: "title" });

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
    { text: `seed.py loaded ${LIVE.kpis.cases} [Demo] deals through real stage changes to simulate a client's CRM.` },
  ], { x: 0.6, y: 6.2, w: 5.55, h: 0.75, fontSize: 14, color: C.text2, objectName: "Org caption" });

  // Right: the SOQL in extract.py, shown verbatim with light syntax colouring.
  const src = fs.readFileSync(path.join(ROOT, "pipeline/extract.py"), "utf8").split("\n");
  const q0 = src.findIndex(l => l.startsWith("QUERIES = {"));
  const q1 = src.findIndex((l, i) => i > q0 && l.startsWith("}"));
  const codeLines = src.slice(q0, q1 + 1).map(l => l.replace(/^ {4}(?= {4})/, ""));
  const KW = /\b(SELECT|FROM|WHERE|ORDER BY|AND)\b/;
  const runs = [];
  codeLines.forEach((line, li) => {
    const parts = line.split(/(\b(?:SELECT|FROM|WHERE|ORDER BY|AND)\b|"[a-z_]+"|'[^']*')/).filter(Boolean);
    parts.forEach((part, pi) => {
      const color = KW.test(part) ? HEX.accent1 : /^["']/.test(part) ? HEX.accent4 : HEX.accent3;
      runs.push({ text: part, options: { color, bold: KW.test(part), breakLine: pi === parts.length - 1 && li < codeLines.length - 1 } });
    });
    if (!parts.length && li < codeLines.length - 1) runs.push({ text: " ", options: { breakLine: true } });
  });
  const px = 6.45, pw = 6.25;
  card(s, px, 1.4, pw, 4.7, "Code panel", C.text2);
  [HEX.accent4, "F6C344", HEX.accent5].forEach((dot, i) =>
    s.addShape(pres.shapes.OVAL, { x: px + 0.2 + i * 0.2, y: 1.55, w: 0.12, h: 0.12, fill: { color: dot }, line: { type: "none" }, objectName: `Code panel dot ${i + 1}` }));
  text(s, "pipeline/extract.py", { x: px + 0.85, y: 1.48, w: 3, h: 0.26, fontSize: 11, color: C.accent3, objectName: "Code panel filename" });
  text(s, runs, { x: px + 0.2, y: 1.82, w: pw - 0.3, h: 4.2, fontSize: 10, fontFace: "Courier New", lineSpacingMultiple: 1.0, objectName: "SOQL code" });
  text(s, [
    { text: "extract.py: ", options: { bold: true } },
    { text: "four SOQL queries over the REST API, signed by the Salesforce CLI session. Opportunity Id becomes the case ID." },
  ], { x: px, y: 6.2, w: pw, h: 0.75, fontSize: 14, color: C.text2, objectName: "Code caption" });
  // Numbered popups walk the viewer through the flow, left to right.
  const popup = (n, x, y, w, label) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.4, rectRadius: 0.08, fill: { color: C.background1 }, line: { color: C.accent1, width: 1.5 },
      shadow: { type: "outer", color: "000000", opacity: 0.25, blur: 4, offset: 2, angle: 90 }, objectName: `Popup ${n} box` });
    s.addShape(pres.shapes.OVAL, { x: x + 0.06, y: y + 0.05, w: 0.3, h: 0.3, fill: { color: C.accent4 }, line: { type: "none" }, objectName: `Popup ${n} badge` });
    text(s, String(n), { x: x + 0.06, y: y + 0.05, w: 0.3, h: 0.3, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle", objectName: `Popup ${n} number` });
    text(s, label, { x: x + 0.44, y, w: w - 0.5, h: 0.4, fontSize: 12, bold: true, color: C.text2, valign: "middle", objectName: `Popup ${n} label` });
  };
  const lineY = i => 1.89 + i * 0.1667; // centre of code line i in the panel
  const lineOf = re => codeLines.findIndex(l => re.test(l));
  popup(1, 3.15, 1.6, 2.85, "Trial org = the client's CRM");
  popup(2, 3.3, 3.85, 2.7, `seed.py: ${LIVE.kpis.cases} [Demo] records`);
  popup(3, px + pw - 2.95, 1.43, 2.85, "SOQL over REST via CLI");
  popup(4, px + pw - 2.2, lineY(lineOf(/FROM OpportunityHistory/)) - 0.2, 2.1, "History → events");
  popup(5, px + pw - 2.45, lineY(lineOf(/FROM Task/)) - 0.2, 2.35, "Tasks join on WhatId");
  s.addNotes("Left: the trial org standing in for a client's CRM, seeded by seed.py so there's real OpportunityHistory to read. Right: the exact SOQL the extractor runs. Opportunity is the case; leads join via ConvertedOpportunityId and tasks via WhatId. Trial orgs can't backdate history, so the 10-month dataset is simulated in the same object shape.");

  // ---------- Tech setup: everything is live ----------
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Tech setup" });
  s.addText("Everything is live and open to inspect", { placeholder: "title" });
  const proof = [
    ["Landing page and value items", SITE, `https://${SITE}`],
    ["Process dashboard", `${SITE}/dashboard.html`, `https://${SITE}/dashboard.html`],
    ["Value case one-pager", `${SITE}/value-case.html`, `https://${SITE}/value-case.html`],
    ["Live Salesforce extract", `${SITE}/live.html`, `https://${SITE}/live.html`],
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
  const shotH = 4.75, shotW = shotH * (1280 / 1180);
  s.addShape(pres.shapes.RECTANGLE, { x: 7.6 - 0.05, y: 1.5, w: shotW + 0.1, h: shotH + 0.1, fill: { color: C.background1 }, line: { color: C.accent3, width: 1 }, objectName: "Screenshot frame" });
  s.addImage({ path: path.join(ROOT, "docs/img/value-case.png"), x: 7.6, y: 1.55, w: shotW, h: shotH, objectName: "Value case screenshot", altText: "Value case one-pager on the hosted site" });
  s.addNotes("All of this is hosted. The repo has the pipeline, the seeding script and the README with the setup gotchas I hit on a new trial org.");

  }

  if (KIND === "main") {
  // ---------- Close: next steps ----------
  pres.addSection({ title: "Close" });
  s = pres.addSlide({ masterName: "CLOSING_DARK", sectionTitle: "Close" });
  circle(s, 10.2, 4.4, 5.4, C.accent1, 75, "Decor circle large");
  s.addText("Turning the finding into a decision", { placeholder: "title" });
  const next = [
    ["Validate", "Run the same extract on the production org and confirm approval waits and re-quote rates with RevOps."],
    ["Act", "Pilot a delegated approval threshold for EMEA discounts up to 30% and make Qualification a required gate."],
    ["Monitor", "Schedule the extract into Apromore, track the KPIs against target and report value realised quarterly."],
  ];
  next.forEach(([head, body], i) => {
    const x = 0.8 + i * 3.95;
    card(s, x, 2.0, 3.65, 2.5, `Next ${i + 1} card`, C.accent6);
    text(s, head, { x: x + 0.3, y: 2.3, w: 3.05, h: 0.5, fontSize: 22, bold: true, color: C.accent3, objectName: `Next ${i + 1} heading` });
    text(s, body, { x: x + 0.3, y: 2.95, w: 3.05, h: 1.4, fontSize: 16, color: C.background1, objectName: `Next ${i + 1} body` });
  });
  text(s, `${SITE}  ·  Luke Audie`, { x: 0.8, y: 6.5, w: 8, h: 0.4, fontSize: 14, color: C.accent3, objectName: "Closing contact" });
  logo(s, 10.35, 6.35, 0.55);
  s.addNotes("Close on the decision: validate on production data, pilot one policy change, then monitor value realised in Apromore.");

  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log(`Wrote ${OUT}${hasLogo ? " (with logo)" : ""}`);
})().catch(e => { console.error(e); process.exit(1); });
