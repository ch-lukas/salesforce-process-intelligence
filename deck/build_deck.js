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
const OUT = process.env.DECK_OUT || path.join(__dirname, "salesforce-process-intelligence.pptx");
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
  pres.title = "From Salesforce data to a framed value case";
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

  // ---------- 1. Title ----------
  pres.addSection({ title: "Opening" });
  let s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Opening" });
  circle(s, 9.4, 3.4, 6.6, C.accent1, 70, "Decor circle large");
  circle(s, 10.6, 2.6, 0.75, C.accent3, 0, "Decor circle small");
  s.addText("From Salesforce data to a framed value case", { placeholder: "title" });
  s.addText("Lead-to-close process intelligence: a working demo built for Apromore", { placeholder: "body" });
  text(s, "Luke Audie  ·  October 2026", { x: 0.8, y: 6.4, w: 6, h: 0.4, fontSize: 14, color: C.accent3, objectName: "Byline" });
  logo(s, 0.8, 0.7, 0.6);
  s.addNotes("I wanted to show how I'd open with a Salesforce-heavy prospect: start from their own CRM data, find where the process leaks, and turn it into a number the CRO cares about.");

  // ---------- 2. Answer first ----------
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

  // ---------- 3. How it was built ----------
  pres.addSection({ title: "Evidence" });
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Evidence" });
  s.addText("Salesforce to Apromore in four steps", { placeholder: "title" });
  const steps = [
    ["Extract", "SOQL over REST", "Lead, Opportunity, OpportunityHistory and Task, signed by the Salesforce CLI session. No stored credentials."],
    ["Correlate", "One case per deal", "Case ID = Opportunity Id. Tasks join on WhatId, Leads on ConvertedOpportunityId."],
    ["Load", "Event log", "CSV and IEEE XES with activity, timestamps, resource and case attributes, ready for Apromore."],
    ["Mine and size", "Map and value", "Directly-follows graph, variant analysis, cohort win rates and a driver-based value model."],
  ];
  steps.forEach(([eyebrow, head, body], i) => {
    const x = 0.6 + i * 3.12;
    card(s, x, 1.5, 2.75, 3.6, `Step ${i + 1} card`);
    s.addShape(pres.shapes.OVAL, { x: x + 0.25, y: 1.75, w: 0.55, h: 0.55, fill: { color: C.accent1 }, line: { type: "none" }, objectName: `Step ${i + 1} marker` });
    text(s, String(i + 1), { x: x + 0.25, y: 1.75, w: 0.55, h: 0.55, fontSize: 18, bold: true, color: C.background1, align: "center", valign: "middle", objectName: `Step ${i + 1} number` });
    text(s, eyebrow, { x: x + 0.25, y: 2.5, w: 2.3, h: 0.3, fontSize: 12, bold: true, color: C.accent6, objectName: `Step ${i + 1} label` });
    text(s, head, { x: x + 0.25, y: 2.82, w: 2.3, h: 0.45, fontSize: 18, bold: true, color: C.text2, objectName: `Step ${i + 1} title` });
    text(s, body, { x: x + 0.25, y: 3.35, w: 2.3, h: 1.65, fontSize: 14, color: C.text1, objectName: `Step ${i + 1} body` });
    if (i < steps.length - 1) {
      s.addShape(pres.shapes.RIGHT_ARROW, { x: x + 2.82, y: 3.12, w: 0.24, h: 0.36, fill: { color: C.accent2 }, line: { type: "none" }, objectName: `Step ${i + 1} arrow` });
    }
  });
  card(s, 0.6, 5.45, 12.1, 1.15, "Live run band", C.text2);
  text(s, [
    { text: "Live run on a Salesforce trial org:  ", options: { bold: true, color: C.accent3 } },
    { text: `${LIVE.kpis.cases} opportunities  ·  ${LIVE.kpis.events} events  ·  ${Object.keys(LIVE.kpis.events_by_object).length} objects joined  ·  CSV + XES for Apromore`, options: { color: C.background1 } },
  ], { x: 0.9, y: 5.45, w: 11.5, h: 1.15, fontSize: 16, valign: "middle", objectName: "Live run text" });
  s.addNotes("Opportunity is the case. Trial orgs can't backdate history, so the 10-month dataset is simulated in the exact Salesforce shape; the extractor runs unchanged against the live org, as the band at the bottom shows.");

  // ---------- 4. Discover ----------
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Evidence" });
  s.addText("Discover: the process as it actually runs", { placeholder: "title" });
  const discover = [
    [pct(K.happy_path_conformance), "of deals follow the designed happy path"],
    [String(K.variants), `distinct paths across ${K.cases} deals`],
    [`${days(D.approval.baseline)} days`, "median EMEA approval wait, the slowest hand-off"],
  ];
  discover.forEach(([v, l], i) => {
    const y = 1.65 + i * 1.55;
    text(s, v, { x: 0.6, y, w: 2.9, h: 0.9, fontSize: 44, bold: true, color: C.accent1, valign: "middle", objectName: `Discover stat ${i + 1}` });
    text(s, l, { x: 3.6, y, w: 3.9, h: 0.9, fontSize: 18, color: C.text2, valign: "middle", objectName: `Discover label ${i + 1}` });
  });
  text(s, "Approval steps and re-quote loops show up as detours from the designed path.", { x: 0.6, y: 6.25, w: 6.9, h: 0.5, fontSize: 14, italic: true, color: C.accent6, objectName: "Discover note" });
  const mapH = 5.25, mapW = mapH * (1080 / 1250);
  s.addShape(pres.shapes.RECTANGLE, { x: 12.7 - mapW - 0.05, y: 1.4, w: mapW + 0.1, h: mapH + 0.1, fill: { color: C.background1 }, line: { color: C.accent3, width: 1 }, objectName: "Map frame" });
  s.addImage({ path: path.join(__dirname, "process-map.png"), x: 12.7 - mapW, y: 1.45, w: mapW, h: mapH, objectName: "Process map", altText: "Process map from the hosted dashboard" });
  s.addNotes("This is the directly-follows map from the hosted dashboard. Fourteen percent conformance is the hook: the CRM shows the stage, process mining shows the path.");

  // ---------- 5. Diagnose ----------
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

  // ---------- 6. Value case ----------
  pres.addSection({ title: "Value" });
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

  // ---------- 7. Proof ----------
  s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Value" });
  s.addText("Everything is live and open to inspect", { placeholder: "title" });
  const proof = [
    ["Process dashboard", SITE, `https://${SITE}`],
    ["Value case one-pager", `${SITE}/value-case.html`, `https://${SITE}/value-case.html`],
    ["Live Salesforce extract", `${SITE}/live.html`, `https://${SITE}/live.html`],
    ["Source code and event logs", REPO, `https://${REPO}`],
  ];
  proof.forEach(([label, shown, url], i) => {
    const y = 1.55 + i * 1.22;
    card(s, 0.6, y, 6.6, 1.05, `Proof ${i + 1}`);
    s.addShape(pres.shapes.OVAL, { x: 0.82, y: y + 0.3, w: 0.45, h: 0.45, fill: { color: C.accent1 }, line: { type: "none" }, objectName: `Proof ${i + 1} marker` });
    text(s, "→", { x: 0.82, y: y + 0.3, w: 0.45, h: 0.45, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle", objectName: `Proof ${i + 1} glyph` });
    text(s, label, { x: 1.45, y: y + 0.14, w: 5.6, h: 0.38, fontSize: 16, bold: true, color: C.text2, objectName: `Proof ${i + 1} label` });
    text(s, url ? [{ text: shown, options: { hyperlink: { url } } }] : shown, { x: 1.45, y: y + 0.54, w: 5.6, h: 0.38, fontSize: 14, color: C.accent6, objectName: `Proof ${i + 1} link` });
  });
  const shotH = 4.75, shotW = shotH * (1280 / 1180);
  s.addShape(pres.shapes.RECTANGLE, { x: 7.6 - 0.05, y: 1.5, w: shotW + 0.1, h: shotH + 0.1, fill: { color: C.background1 }, line: { color: C.accent3, width: 1 }, objectName: "Screenshot frame" });
  s.addImage({ path: path.join(ROOT, "docs/img/value-case.png"), x: 7.6, y: 1.55, w: shotW, h: shotH, objectName: "Value case screenshot", altText: "Value case one-pager on the hosted site" });
  s.addNotes("All of this is hosted. The repo has the pipeline, the seeding script and the README with the setup gotchas I hit on a new trial org.");

  // ---------- 8. Next steps ----------
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

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log(`Wrote ${OUT}${hasLogo ? " (with logo)" : ""}`);
})().catch(e => { console.error(e); process.exit(1); });
