// Builds the two-page revenue-targets dashboard (مؤشرات الإستدامة المالية + حالة المديونيات).
//   node build.js            -> dashboard.pptx            (full slides with numbers and native charts)
//   node build.js --bg       -> powerbi-background.pptx   (same frame, numbers/charts left empty, for Power BI page backgrounds)
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require(process.env.APPLY_THEME || "/root/.claude/skills/synced/642995fb-6fb0-4612-9fc1-a71063ac95e2_2132651f-4df5-4687-9c61-978ed9bca072/pptx/scripts/apply_theme.js");

const BG = process.argv.includes("--bg");
const OUT = path.join(__dirname, BG ? "powerbi-background.pptx" : "dashboard.pptx");

// ---------------------------------------------------------------- colours (sampled from the original Amana pages)
const THEME = {
  name: "Amana Revenue Targets", headFontFace: "Vazirmatn", bodyFontFace: "Vazirmatn",
  colors: { dk1: "1E2B2A", lt1: "FFFFFF", dk2: "2D6A66", lt2: "F2F5F4", accent1: "2D6A66", accent2: "86B19A", accent3: "8A8F8E", accent4: "3B7F7C", accent5: "2E3F7F", accent6: "5E9467", hlink: "2D6A66", folHlink: "5E6B6A" },
};
const X = { teal: "2D6A66", teal2: "3B7F7C", sage: "86B19A", grey: "8A8F8E", track: "E4E9E8", ink: "1E2B2A", mute: "6B7675", bg: "F2F5F4", line: "DCE3E1" };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 — same 16:9 ratio as a Power BI page (1280 x 720)
pres.rtlMode = true;
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "مؤشرات الإستدامة المالية";
pres.author = "أمانة منطقة الرياض — إدارة تخطيط و مراقبة الإيرادات";
const C = pres.SchemeColor;

pres.defineSlideMaster({
  title: "DASH_T",
  background: { color: X.bg },
  objects: [
    { image: { x: 0.5, y: 0.22, w: 0.68, h: 0.8, path: path.join(__dirname, "logo.png") } },
    { text: { text: "إدارة تخطيط و مراقبة الإيرادات", options: { x: 1.3, y: 0.42, w: 3.2, h: 0.4, fontSize: 11, color: X.mute, align: "left", valign: "middle", margin: 0, rtlMode: true } } },
    { placeholder: { options: { name: "title", type: "title", x: 6.0, y: 0.3, w: 6.83, h: 0.66, fontSize: 30, bold: true, color: C.accent1, align: "right", valign: "middle", margin: 0, rtlMode: true }, text: "" } },
  ],
});

// ---------------------------------------------------------------- helpers
const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, rtlMode: true, align: "right", valign: "middle", margin: 0, fontSize: 12, color: X.ink }, o));
const V = (s, text, o) => { if (!BG) T(s, text, o); }; // data value: omitted on the Power BI background
const shadow = () => ({ type: "outer", color: "1E2B2A", opacity: 0.1, blur: 10, offset: 2, angle: 90 });
const card = (s, x, y, w, h, name, fill) => s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.14, fill: { color: fill || "FFFFFF" }, line: { type: "none" }, shadow: shadow(), objectName: name });
const MIL = "مليون ريال";

async function icon(Comp, color) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: 256 }));
  return "image/png;base64," + (await sharp(Buffer.from(svg)).png().toBuffer()).toString("base64");
}
// icon in a filled circle — the repeated motif on every card
async function badge(s, Comp, x, y, d, fill, fg) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: fill || X.teal }, line: { type: "none" } });
  const p = d * 0.27;
  s.addImage({ data: await icon(Comp, fg || "FFFFFF"), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
}
// card heading: badge at the right edge, title to its left
async function heading(s, Comp, title, right, y, w) {
  await badge(s, Comp, right - 0.42, y, 0.42);
  T(s, title, { x: right - 0.42 - 0.12 - w, y, w, h: 0.42, fontSize: 15, bold: true, color: X.teal });
}
// value + unit on one line, right-aligned: "3,399 مليون ريال"
function amount(s, value, x, y, w, h, size, color) {
  V(s, [
    { text: value, options: { fontSize: size, bold: true, color: color || X.ink } },
    { text: "  " + MIL, options: { fontSize: Math.max(10, Math.round(size * 0.36)), color: X.mute } },
  ], { x, y, w, h });
}

function doughnut(s, pct, x, y, d, color, track) {
  if (BG) return;
  s.addChart(pres.ChartType.doughnut, [{ name: "نسبة", labels: ["المنجز", "المتبقي"], values: [pct, 100 - pct] }], {
    x, y, w: d, h: d, holeSize: 74, chartColors: [color, track || X.track], dataBorder: { pt: 0.5, color: "FFFFFF" },
    showLegend: false, showValue: false, showPercent: false, showLabel: false, showTitle: false, firstSliceAng: 0,
  });
}
// half-ring gauge: the bottom half is a slice painted the card colour
function gauge(s, pct, x, y, d, color) {
  if (BG) return;
  s.addChart(pres.ChartType.doughnut, [{ name: "نسبة", labels: ["المنجز", "المتبقي", "-"], values: [pct, 100 - pct, 100] }], {
    x, y, w: d, h: d, holeSize: 68, chartColors: [color, X.track, "FFFFFF"], dataBorder: { pt: 0.5, color: "FFFFFF" },
    showLegend: false, showValue: false, showPercent: false, showLabel: false, showTitle: false, firstSliceAng: 270,
  });
}
// progress bar filling right-to-left
function bar(s, pct, x, y, w, color) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h: 0.13, rectRadius: 0.065, fill: { color: X.track }, line: { type: "none" } });
  if (!BG) s.addShape(pres.ShapeType.roundRect, { x: x + w * (1 - pct / 100), y, w: w * pct / 100, h: 0.13, rectRadius: 0.065, fill: { color }, line: { type: "none" } });
}

// ================================================================ PAGE 1 — مؤشرات الإستدامة المالية
async function page1() {
  pres.addSection({ title: "مستهدفات الإيرادات" });
  const s = pres.addSlide({ masterName: "DASH_T", sectionTitle: "مستهدفات الإيرادات" });
  s.addText("مؤشرات الإستدامة المالية", { placeholder: "title" });

  // ---------------- card A: billed revenue vs target (right)
  const ax = 8.23, ay = 1.15, aw = 4.6, ah = 3.25, aR = ax + aw - 0.25;
  card(s, ax, ay, aw, ah, "card-billed");
  await heading(s, fa.FaFileInvoiceDollar, "الإيرادات المفوترة", aR, ay + 0.2, 3.0);
  T(s, "قيمة المفوتر", { x: aR - 1.4, y: ay + 0.66, w: 1.4, h: 0.55, fontSize: 13, color: X.mute });
  amount(s, "3,399", ax + 0.25, ay + 0.61, aw - 1.9, 0.6, 32, X.ink);

  const halves = [
    { pct: 55, lab: "نسبة المفوتر من المستهدف (الفترة الحالية)", tgt: "3,368", tl: "المستهدف (الفترة الحالية)", col: X.teal },
    { pct: 45, lab: "نسبة المفوتر من المستهدف (الإجمالي)", tgt: "3,364", tl: "المستهدف (الإجمالي)", col: X.sage },
  ];
  const hw = (aw - 0.5) / 2;
  halves.forEach((h, i) => {
    const hx = aR - hw * (i + 1), dy = ay + 1.18, d = 0.85;
    // donut on the right of each half, labels to its left
    s.addShape(pres.ShapeType.ellipse, { x: hx + hw - d, y: dy, w: d, h: d, fill: { color: "FFFFFF" }, line: { color: X.track, width: BG ? 9 : 0.5 } });
    doughnut(s, h.pct, hx + hw - d, dy, d, h.col);
    V(s, h.pct + "%", { x: hx + hw - d, y: dy + 0.22, w: d, h: 0.4, fontSize: 16, bold: true, align: "center", color: h.col === X.sage ? "4E7F64" : X.teal });
    T(s, h.tl, { x: hx, y: dy + 0.9, w: hw - 0.05, h: 0.25, fontSize: 10, color: X.mute });
    amount(s, h.tgt, hx, dy + 1.12, hw - 0.05, 0.32, 15);
  });
  T(s, "يقيس المؤشر إجمالي الإيرادات المفوترة من قبل الأمانة خلال فترة زمنية محددة، لدعم متابعة الأداء المالي وتعزيز مصادر دخل الوزارة",
    { x: ax + 0.25, y: ay + ah - 0.56, w: aw - 0.5, h: 0.46, fontSize: 10, color: X.mute, valign: "top" });

  // ---------------- card B: collection rate (left)
  const bx = 0.5, by = 1.15, bw = 7.48, bh = 3.25, bR = bx + bw - 0.25;
  card(s, bx, by, bw, bh, "card-collection");
  await heading(s, fa.FaCoins, "مؤشر معدل التحصيل", bR, by + 0.2, 4.0);

  const gw = 2.6, d = 2.0;
  const gauges = [
    { x: bR - gw, pct: 78, lab: "إجمالي المحصل (السنة الحالية)", val: "3,305", cap: "نسبة محصل السنة الحالية من اجمالي المفوتر", col: X.teal },
    { x: bx + 0.25, pct: 76, lab: "إجمالي المحصل (السنوات السابقة)", val: "466", cap: "نسبة محصل السنوات السابقة من اجمالي المحصل", col: X.grey },
  ];
  for (const g of gauges) {
    const gx = g.x + (gw - d) / 2, gy = by + 0.75;
    gauge(s, g.pct, gx, gy, d, g.col);
    if (BG) s.addShape(pres.ShapeType.blockArc, { x: gx + 0.1, y: gy + 0.1, w: d - 0.2, h: d - 0.2, angleRange: [180, 0], arcThicknessRatio: 0.2, fill: { color: X.track }, line: { type: "none" } });
    V(s, g.pct + "%", { x: gx, y: gy + 0.55, w: d, h: 0.42, fontSize: 24, bold: true, align: "center", color: g.col === X.grey ? "5E6463" : X.teal });
    T(s, g.cap, { x: g.x, y: gy + 1.02, w: gw, h: 0.22, fontSize: 10, align: "center", color: X.mute });
    T(s, g.lab, { x: g.x, y: gy + 1.3, w: gw, h: 0.25, fontSize: 11, bold: true, align: "center", color: X.ink });
    V(s, [{ text: g.val, options: { fontSize: 26, bold: true } }, { text: "  " + MIL, options: { fontSize: 10, color: X.mute } }],
      { x: g.x, y: gy + 1.54, w: gw, h: 0.45, align: "center" });
  }
  // centre column: two stacked stat tiles
  const mx = bx + 0.25 + gw + 0.2, mw = bR - gw - 0.2 - mx;
  const mids = [
    { val: "2,955", lab: "إجمالي المحصل", unit: true, fill: X.bg, fg: X.ink },
    { val: "13%", lab: "نسبة التحصيل", unit: false, fill: X.teal, fg: "FFFFFF" },
  ];
  mids.forEach((m, i) => {
    const my = by + 0.85 + i * 1.08;
    s.addShape(pres.ShapeType.roundRect, { x: mx, y: my, w: mw, h: 0.92, rectRadius: 0.12, fill: { color: m.fill }, line: { type: "none" } });
    V(s, m.unit ? [{ text: m.val, options: { fontSize: 22, bold: true } }, { text: "  " + MIL, options: { fontSize: 9, color: X.mute } }] : m.val,
      { x: mx, y: my + 0.08, w: mw, h: 0.48, align: "center", fontSize: 24, bold: true, color: m.fg });
    T(s, m.lab, { x: mx, y: my + 0.56, w: mw, h: 0.26, align: "center", fontSize: 11, color: m.fg === "FFFFFF" ? "DCEBE8" : X.mute });
  });
  T(s, "يقيس المؤشر نسبة الإيرادات المفوترة التي تم تحصيلها خلال فترة زمنية محددة، بما يعكس كفاءة الوزارة في إدارة المستحقات ومنظومة التحصيل",
    { x: bx + 0.25, y: by + bh - 0.42, w: bw - 0.5, h: 0.32, fontSize: 10, color: X.mute, align: "center" });

  // ---------------- card C: operating-expense coverage (full width)
  const cx = 0.5, cy = 4.6, cw = 12.33, ch = 2.62, cR = cx + cw - 0.25;
  card(s, cx, cy, cw, ch, "card-opex");
  await heading(s, fa.FaBalanceScale, "مؤشر تغطية النفقات التشغيلية", cR, cy + 0.18, 5.0);
  const cols = [
    { t: "الإيرادات المفوترة (الإعتماد الأصلي)", cap: "تغطية النفقات التشغيلية باستخدام الإيرادات المفوترة", vl: "قيمة المفوتر", v: "3,399", rows: [["الفترة الحالية", 49, "3,993"], ["الإجمالي", 70, "3,593"]] },
    { t: "الإيرادات المفوترة (الإعتماد بعد التعديل)", cap: "تغطية النفقات التشغيلية المعدلة باستخدام الإيرادات المفوترة (بعد التعزيز)", vl: "قيمة المفوتر", v: "3,399", rows: [["الفترة الحالية", 52, "3,911"], ["الإجمالي", 50, "3,259"]] },
    { t: "الإيرادات المحصلة (الإعتماد بعد التعديل)", cap: "تغطية النفقات التشغيلية المعدلة باستخدام الإيرادات المحصلة", vl: "قيمة المحصل", v: "2,955", rows: [["الفترة الحالية", 71, "3,911"], ["الإجمالي", 76, "3,259"]] },
  ];
  const colW = (cw - 0.5) / 3, pad = 0.22;
  cols.forEach((c, i) => {
    const x0 = cR - colW * (i + 1), x = x0 + pad, w = colW - 2 * pad, top = cy + 0.72;
    if (i > 0) s.addShape(pres.ShapeType.line, { x: x0 + colW, y: top + 0.05, w: 0, h: ch - 0.95, line: { color: X.line, width: 0.75 } });
    T(s, c.t, { x, y: top, w, h: 0.26, fontSize: 12, bold: true });
    T(s, c.cap, { x, y: top + 0.27, w, h: 0.22, fontSize: 10, color: X.mute });
    T(s, c.vl, { x: x + w - 1.2, y: top + 0.55, w: 1.2, h: 0.4, fontSize: 11, color: X.mute });
    amount(s, c.v, x, top + 0.55, w - 1.25, 0.4, 20);
    c.rows.forEach(([lab, pct, tgt], j) => {
      const ry = top + 1.04 + j * 0.42, col = j === 0 ? X.teal2 : X.sage;
      T(s, lab, { x: x + w - 1.1, y: ry, w: 1.1, h: 0.22, fontSize: 11, bold: true, color: j === 0 ? X.teal : "4E7F64" });
      V(s, [{ text: "المستهدف  ", options: { color: X.mute } }, { text: tgt, options: { bold: true, color: X.ink } }, { text: " " + MIL, options: { color: X.mute } }],
        { x: x + 0.55, y: ry, w: w - 1.7, h: 0.22, fontSize: 10, align: "center" });
      V(s, pct + "%", { x, y: ry, w: 0.6, h: 0.22, fontSize: 13, bold: true, align: "left", color: j === 0 ? X.teal : "4E7F64" });
      bar(s, pct, x, ry + 0.25, w, col);
    });
  });
  s.addNotes("مؤشرات الإستدامة المالية — جميع القيم بالمليون ريال.");
}

// ================================================================ PAGE 2 — حالة المديونيات
async function page2() {
  pres.addSection({ title: "المديونيات" });
  const s = pres.addSlide({ masterName: "DASH_T", sectionTitle: "المديونيات" });
  s.addText("حالة المديونيات", { placeholder: "title" });

  const kpis = [
    { ic: fa.FaFileInvoiceDollar, lab: "رصيد المديونية بداية السنة الحالية", v: "121" },
    { ic: fa.FaTimes, lab: "الملغي من المديونية في السنة الحالية", v: "119" },
    { ic: fa.FaCheck, lab: "ما تم سداده من المديونية في السنة الحالية", v: "328" },
    { ic: fa.FaHandshake, lab: "نسبة السداد", v: "62%", hero: true },
    { ic: fa.FaExclamationTriangle, lab: "المتبقي من المديونية", v: "125" },
  ];
  const gap = 0.2, kw = (12.33 - gap * 4) / 5, ky = 1.15, kh = 1.2;
  for (let i = 0; i < kpis.length; i++) {
    const k = kpis[i], x = 12.83 - kw * (i + 1) - gap * i, fg = k.hero ? "FFFFFF" : X.ink;
    card(s, x, ky, kw, kh, "kpi-" + (i + 1), k.hero ? X.teal : "FFFFFF");
    await badge(s, k.ic, x + kw - 0.2 - 0.46, ky + 0.2, 0.46, k.hero ? "FFFFFF" : X.teal, k.hero ? X.teal : "FFFFFF");
    T(s, k.lab, { x: x + 0.18, y: ky + 0.14, w: kw - 0.9, h: 0.56, fontSize: 10.5, color: k.hero ? "DCEBE8" : X.mute, valign: "middle" });
    V(s, k.hero ? k.v : [{ text: k.v, options: { fontSize: 28, bold: true } }, { text: "  " + MIL, options: { fontSize: 10, color: X.mute } }],
      { x: x + 0.18, y: ky + 0.68, w: kw - 0.36, h: 0.44, fontSize: 28, bold: true, color: fg });
  }

  const cats = ["الإيرادات الإستثمارية", "رسوم وتراخيص", "جزاءات وغرامات", "التبغ والإيواء", "أخرى"];
  const periods = [
    { name: "2007 - 2016", tot: ["116", "299", "122"], bal: [116, 116, 128, 110, 117], paid: [318, 276, 295, 302, 334] },
    { name: "2017 - 2025", tot: ["126", "327", "112"], bal: [108, 121, 124, 130, 108], paid: [277, 273, 279, 283, 279] },
  ];
  const pw = (12.33 - 0.25) / 2, py = 2.6, ph = 4.62;
  periods.forEach((p, i) => {
    const px = 12.83 - pw * (i + 1) - 0.25 * i, pR = px + pw - 0.25;
    card(s, px, py, pw, ph, "period-" + p.name);
    // period pill + legend
    s.addShape(pres.ShapeType.roundRect, { x: pR - 1.75, y: py + 0.2, w: 1.75, h: 0.44, rectRadius: 0.22, fill: { color: X.teal }, line: { type: "none" } });
    T(s, p.name, { x: pR - 1.75, y: py + 0.2, w: 1.75, h: 0.44, fontSize: 15, bold: true, color: "FFFFFF", align: "center" });
    [["رصيد المديونية", X.teal], ["ما تم سداده", X.sage]].forEach(([lab, col], j) => {
      const lx = px + 0.25 + (1 - j) * 1.55;
      s.addShape(pres.ShapeType.roundRect, { x: lx + 1.2, y: py + 0.36, w: 0.14, h: 0.14, rectRadius: 0.03, fill: { color: col }, line: { type: "none" } });
      T(s, lab, { x: lx, y: py + 0.3, w: 1.13, h: 0.26, fontSize: 10, color: X.mute });
    });
    // three totals
    const tl = ["إجمالي رصيد المديونية", "إجمالي ما تم سداده", "إجمالي المتبقي"], tw = (pw - 0.5 - 0.3) / 3;
    tl.forEach((lab, j) => {
      const tx = pR - tw * (j + 1) - 0.15 * j, ty = py + 0.82;
      s.addShape(pres.ShapeType.roundRect, { x: tx, y: ty, w: tw, h: 0.86, rectRadius: 0.1, fill: { color: X.bg }, line: { type: "none" } });
      T(s, lab, { x: tx + 0.12, y: ty + 0.08, w: tw - 0.24, h: 0.26, fontSize: 10.5, color: j === 1 ? "4E7F64" : X.mute });
      V(s, [{ text: p.tot[j], options: { fontSize: 22, bold: true, color: j === 1 ? "4E7F64" : X.ink } }, { text: "  " + MIL, options: { fontSize: 9, color: X.mute } }],
        { x: tx + 0.12, y: ty + 0.36, w: tw - 0.24, h: 0.42 });
    });
    // categories: balance vs paid
    if (!BG) s.addChart(pres.ChartType.bar, [
      { name: "رصيد المديونية", labels: cats, values: p.bal },
      { name: "ما تم سداده", labels: cats, values: p.paid },
    ], {
      x: px + 0.15, y: py + 1.82, w: pw - 0.3, h: ph - 1.95, barDir: "bar", barGrouping: "clustered", barGapWidthPct: 55, barOverlapPct: -8,
      chartColors: [X.teal, X.sage], showLegend: false, showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11, dataLabelFontBold: true,
      dataLabelColor: X.ink, dataLabelFontFace: "+mn-lt",
      valAxisHidden: true, valGridLine: { style: "none" }, valAxisMaxVal: 400, valAxisMinVal: 0, valAxisOrientation: "maxMin",
      catAxisOrientation: "maxMin", catAxisLabelColor: X.ink, catAxisLabelFontSize: 11, catAxisLabelFontFace: "+mn-lt", catAxisLineShow: false, catGridLine: { style: "none" },
    });
  });
  s.addNotes("حالة المديونيات — جميع القيم بالمليون ريال. الرسم يقارن رصيد المديونية بما تم سداده لكل تبويب في كل فترة.");
}

// ================================================================ chart pages (2x2 grid of native charts)
const GRID = { y1: 1.15, h1: 2.9, y2: 4.25, h2: 2.97, w: (12.33 - 0.25) / 2 };
const cellX = (col) => (col === 0 ? 12.83 - GRID.w : 0.5); // col 0 = right
async function chartCard(s, col, row, Comp, title, name) {
  const x = cellX(col), y = row === 0 ? GRID.y1 : GRID.y2, h = row === 0 ? GRID.h1 : GRID.h2;
  card(s, x, y, GRID.w, h, name);
  await heading(s, Comp, title, x + GRID.w - 0.25, y + 0.18, GRID.w - 1.0);
  return { x: x + 0.2, y: y + 0.68, w: GRID.w - 0.4, h: h - 0.8 };
}
// one look for every column/bar chart: no gridlines, hidden value axis, values on the bars, categories right-to-left
function chart(s, type, data, box, o) {
  if (BG) return;
  s.addChart(type, data, Object.assign({
    x: box.x, y: box.y, w: box.w, h: box.h, barDir: "col", barGrouping: "clustered", barGapWidthPct: 70, barOverlapPct: -10,
    showLegend: data.length > 1, legendPos: "t", legendFontSize: 11, legendColor: X.mute, legendFontFace: "+mn-lt",
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11, dataLabelFontBold: true, dataLabelColor: X.ink, dataLabelFontFace: "+mn-lt", dataLabelFormatCode: "#,##0",
    valAxisHidden: true, valGridLine: { style: "none" }, valAxisMinVal: 0,
    catAxisOrientation: "maxMin", catAxisLabelColor: X.ink, catAxisLabelFontSize: 11, catAxisLabelFontFace: "+mn-lt", catAxisLineShow: true, catGridLine: { style: "none" },
  }, o));
}

async function page3() {
  const s = pres.addSlide({ masterName: "DASH_T", sectionTitle: "مستهدفات الإيرادات" });
  s.addText("تحليل مستهدفات الإيرادات", { placeholder: "title" });

  let b = await chartCard(s, 0, 0, fa.FaFileInvoiceDollar, "الإيرادات المفوترة مقابل المستهدف", "chart-billed");
  chart(s, pres.ChartType.bar, [
    { name: "قيمة المفوتر", labels: ["الفترة الحالية", "الإجمالي"], values: [3399, 3399] },
    { name: "المستهدف", labels: ["الفترة الحالية", "الإجمالي"], values: [3368, 3364] },
  ], b, { chartColors: [X.teal, X.sage], valAxisMaxVal: 4000, barGapWidthPct: 110 });

  b = await chartCard(s, 1, 0, fa.FaCoins, "إجمالي المحصل", "chart-collected");
  chart(s, pres.ChartType.bar, [
    { name: "إجمالي المحصل", labels: ["السنة الحالية", "السنوات السابقة", "إجمالي المحصل"], values: [3305, 466, 2955] },
  ], b, { chartColors: [X.teal], valAxisMaxVal: 3900, barGapWidthPct: 90 });

  const sc = ["المفوترة (الإعتماد الأصلي)", "المفوترة (بعد التعديل)", "المحصلة (بعد التعديل)"];
  b = await chartCard(s, 0, 1, fa.FaBalanceScale, "نسبة تغطية النفقات التشغيلية", "chart-coverage");
  chart(s, pres.ChartType.bar, [
    { name: "الفترة الحالية", labels: sc, values: [49, 52, 71] },
    { name: "الإجمالي", labels: sc, values: [70, 50, 76] },
  ], b, { chartColors: [X.teal, X.sage], valAxisMaxVal: 100, dataLabelFormatCode: '0"%"' });

  b = await chartCard(s, 1, 1, fa.FaBullseye, "القيمة مقابل مستهدف النفقات التشغيلية", "chart-opex-targets");
  chart(s, pres.ChartType.bar, [
    { name: "القيمة", labels: sc, values: [3399, 3399, 2955] },
    { name: "المستهدف (الفترة الحالية)", labels: sc, values: [3993, 3911, 3911] },
    { name: "المستهدف (الإجمالي)", labels: sc, values: [3593, 3259, 3259] },
  ], b, { chartColors: [X.teal, X.sage, X.grey], valAxisMaxVal: 4700, dataLabelFontSize: 10, barGapWidthPct: 50, barOverlapPct: -5 });
  s.addNotes("تحليل مستهدفات الإيرادات — القيم بالمليون ريال، ونسب التغطية بالنسبة المئوية.");
}

async function page4() {
  const s = pres.addSlide({ masterName: "DASH_T", sectionTitle: "المديونيات" });
  s.addText("تحليل المديونيات", { placeholder: "title" });

  let b = await chartCard(s, 0, 0, fa.FaExchangeAlt, "حركة المديونية في السنة الحالية", "chart-debt-flow");
  chart(s, pres.ChartType.bar, [
    { name: "مليون ريال", labels: ["رصيد بداية السنة", "الملغي", "ما تم سداده", "المتبقي"], values: [121, 119, 328, 125] },
  ], b, { chartColors: [X.teal], valAxisMaxVal: 380, barGapWidthPct: 80 });

  // payment ratio: large ring + the two amounts behind it
  b = await chartCard(s, 1, 0, fa.FaHandshake, "نسبة السداد", "chart-pay-ratio");
  const d = 1.95, dx = b.x + 0.35, dy = b.y + (b.h - d) / 2;
  if (BG) s.addShape(pres.ShapeType.ellipse, { x: dx + 0.12, y: dy + 0.12, w: d - 0.24, h: d - 0.24, fill: { color: "FFFFFF" }, line: { color: X.track, width: 18 } });
  doughnut(s, 62, dx, dy, d, X.teal);
  V(s, "62%", { x: dx, y: dy + d / 2 - 0.25, w: d, h: 0.5, fontSize: 24, bold: true, align: "center", color: X.teal });
  [["ما تم سداده من المديونية في السنة الحالية", "328", X.teal], ["المتبقي من المديونية", "125", X.grey]].forEach(([lab, v, col], j) => {
    const ty = b.y + 0.3 + j * 1.0, tx = dx + d + 0.4, tw = b.x + b.w - tx;
    s.addShape(pres.ShapeType.roundRect, { x: tx + tw - 0.14, y: ty + 0.07, w: 0.14, h: 0.14, rectRadius: 0.03, fill: { color: col }, line: { type: "none" } });
    T(s, lab, { x: tx, y: ty, w: tw - 0.24, h: 0.28, fontSize: 11, color: X.mute });
    amount(s, v, tx, ty + 0.32, tw - 0.24, 0.45, 24);
  });

  b = await chartCard(s, 0, 1, fa.FaCalendarAlt, "مقارنة الفترتين", "chart-periods");
  const tl = ["إجمالي رصيد المديونية", "إجمالي ما تم سداده", "إجمالي المتبقي"];
  chart(s, pres.ChartType.bar, [
    { name: "2007 - 2016", labels: tl, values: [116, 299, 122] },
    { name: "2017 - 2025", labels: tl, values: [126, 327, 112] },
  ], b, { chartColors: [X.teal, X.sage], valAxisMaxVal: 380 });

  b = await chartCard(s, 1, 1, fa.FaLayerGroup, "رصيد المديونية حسب التبويب", "chart-tabs");
  const cats = ["الإيرادات الإستثمارية", "رسوم وتراخيص", "جزاءات وغرامات", "التبغ والإيواء", "أخرى"];
  chart(s, pres.ChartType.bar, [
    { name: "2007 - 2016", labels: cats, values: [116, 116, 128, 110, 117] },
    { name: "2017 - 2025", labels: cats, values: [108, 121, 124, 130, 108] },
  ], b, { chartColors: [X.teal, X.sage], valAxisMaxVal: 150, barGapWidthPct: 60, barOverlapPct: -5 });
  s.addNotes("تحليل المديونيات — جميع القيم بالمليون ريال.");
}

(async () => {
  await page1();
  await page3();
  await page2();
  await page4();
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
