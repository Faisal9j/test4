// Builds revenue-report-2026.pptx from data/data.json (see scripts/extract_data.py).
// Design follows the Ministry of Tourism statistical report: deep green, mint accents,
// top tab strip, year chips. Data is organised by YEAR and TAB (التبويب), never by line item.
// Run: NODE_PATH=<dir with pptxgenjs> node scripts/build_deck.js [out.pptx]
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.APPLY_THEME || "/root/.claude/skills/synced/642995fb-6fb0-4612-9fc1-a71063ac95e2_2132651f-4df5-4687-9c61-978ed9bca072/pptx/scripts/apply_theme.js");

const ROOT = path.join(__dirname, "..");
const OUT = process.argv[2] || path.join(ROOT, "revenue-report-2026.pptx");
const LOGO = path.join(ROOT, "assets", "logo.png");
const D = require(path.join(ROOT, "data", "data.json"));

// ---------------------------------------------------------------- theme
const THEME = {
  name: "Ministry of Tourism Revenue",
  headFontFace: "Vazirmatn",
  bodyFontFace: "Vazirmatn",
  colors: {
    dk1: "032F2E", lt1: "FFFFFF", dk2: "4F5B73", lt2: "F1F6F3",
    accent1: "048F1D", accent2: "077948", accent3: "8CDB4F", accent4: "018A9C",
    accent5: "F0648C", accent6: "5B6B7B", hlink: "077948", folHlink: "5B6B7B",
  },
};
const HEX = { dk: "032F2E", dg: "01523B", g: "048F1D", lt: "F1F6F3", sl: "5B6B7B", grid: "DCE6E1", slate: "4F5B73", card: "0A3E36", edge: "2F8F4F", tab: "D2E8DF" };
const SERIES = ["01523B", "048F1D", "8CDB4F", "018A9C", "A0AAB8", "F0648C"];
const A = (f) => path.join(ROOT, "assets", f);

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.rtlMode = true;
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "تقرير الإيرادات 2026";
pres.author = "وزارة السياحة";
const C = pres.SchemeColor;
const K = { dk: C.text1, dg: C.text2, wh: C.background1, lt: C.background2, g: C.accent1, pan: C.accent2, lime: C.accent3, teal: C.accent4, pink: C.accent5, sl: C.accent6 };

// ---------------------------------------------------------------- data helpers
const rows = (k) => D[k].slice(1).filter((r) => r.some((c) => c !== null));
const n = (v) => (typeof v === "number" ? v : 0);
const sum = (a, f) => a.reduce((t, r) => t + n(f(r)), 0);
const TABS = ["إيرادات إستثمارية", "الرسوم والتراخيص", "جزاءات وغرامات", "التبغ والإيواء", "أخرى"];
const ADM = ["منطقة الرياض", "مدينة الرياض", "بلديات المنطقة"];
const MONTHS = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"];
const M = (v) => Math.round((n(v) / 1e6) * 10) / 10;
const fm = (v) => (n(v) / 1e6).toLocaleString("en-US", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
const delta = (a, b) => (b ? (a / b - 1) * 100 : null);
const YEAR = 2026;

// ---------------------------------------------------------------- slide plan (two-pass so tabs can link to slides)
const plan = [];
const ids = {};
const add = (id, fn) => { ids[id] = plan.length + 1; plan.push({ id, fn }); };
const num = (id) => ids[id];
const seenSections = new Set();
let cur = 0;

const FOOT = "تقرير الإيرادات 2026  |  المبالغ بالمليون ريال، والأرقام مقرّبة وقد لا تتطابق المجاميع";
const footer = (c) => ({ text: { text: FOOT, options: { x: 3.2, y: 7.02, w: 9.63, h: 0.3, fontSize: 10, color: c, align: "right", rtlMode: true, margin: 0 } } });
pres.defineSlideMaster({ title: "LIGHT", background: { path: A("bg-dots.png") },
  objects: [{ image: { x: 0.5, y: 0.25, w: 1.95, h: 0.75, path: LOGO } }, footer(HEX.sl)],
  slideNumber: { x: 0.5, y: 7.02, w: 0.6, h: 0.3, fontSize: 10, color: HEX.sl, align: "left" } });
pres.defineSlideMaster({ title: "DARK", background: { path: A("bg-dark.png") },
  objects: [{ rect: { x: 0.5, y: 0.25, w: 1.95, h: 0.75, fill: { color: "FFFFFF" }, rectRadius: 0.12 } }, { image: { x: 0.5, y: 0.25, w: 1.95, h: 0.75, path: LOGO } }, footer("9FB8AE")],
  slideNumber: { x: 0.5, y: 7.02, w: 0.6, h: 0.3, fontSize: 10, color: "9FB8AE", align: "left" } });
pres.defineSlideMaster({ title: "BLANK", background: { path: A("bg-dots.png") }, objects: [] });

// ---------------------------------------------------------------- building blocks
const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, rtlMode: true, align: "right", margin: 0, valign: "middle", fontSize: 14, color: K.dg }, o));

function chips(s, labels, active, targets, y, right, dark, minW) {
  const gap = 0.1, avail = minW ? 3.6 : 12.33;
  let ws = labels.map((lb) => Math.max(minW || 0.95, lb.length * 0.092 + 0.5));
  const f = Math.min(1, avail / (ws.reduce((a, b) => a + b, 0) + gap * (ws.length - 1)));
  ws = ws.map((w) => w * f);
  const fs = f < 1 ? 11 : 12;
  let x = right;
  labels.forEach((lb, i) => {
    const w = ws[i];
    x -= w;
    const on = i === active;
    const o = { x, y, w, h: 0.4, shape: pres.ShapeType.roundRect, rectRadius: 0.12, align: "center", fontSize: fs, bold: on };
    if (on) Object.assign(o, dark ? { color: K.wh, fill: { color: HEX.g } } : { color: K.dk, fill: { color: HEX.tab }, line: { color: HEX.edge, width: 0.75 } });
    else Object.assign(o, { color: dark ? "CFE3DA" : K.sl });
    T(s, lb, o);
    if (targets && targets[i] && !on) {
      s.addShape(pres.ShapeType.roundRect, { x, y, w, h: 0.4, rectRadius: 0.12, fill: { color: "FFFFFF", transparency: 100 }, line: { type: "none" }, hyperlink: { slide: num(targets[i]), tooltip: lb } });
    }
    x -= gap;
  });
}

// navigation buttons (index / previous / next) like the report header
function navButtons(s, dark) {
  const bs = 0.5, defs = [["الفهرس", "index"], ["السابق", cur - 1], ["التالي", cur + 1]];
  defs.forEach(([lb, tg], i) => {
    const x = 12.33 - i * 0.78, y = 0.2;
    const line = dark ? "9FB8AE" : "5B6B7B";
    s.addShape(pres.ShapeType.roundRect, { x, y, w: bs, h: bs, rectRadius: 0.1, fill: { color: "FFFFFF", transparency: 100 }, line: { color: line, width: 1 } });
    if (i === 0) {
      [[0, 0, "6C5CE7"], [1, 0, "048F1D"], [0, 1, "048F1D"], [1, 1, "8FB4F0"]].forEach(([a, b, c]) => s.addShape(pres.ShapeType.roundRect, { x: x + 0.09 + a * 0.17, y: y + 0.09 + b * 0.17, w: 0.14, h: 0.14, rectRadius: 0.03, fill: { color: "FFFFFF", transparency: 100 }, line: { color: c, width: 1.5 } }));
    } else {
      T(s, i === 1 ? "↑" : "↓", { x, y, w: bs, h: bs, align: "center", fontSize: 18, bold: true, color: dark ? K.wh : K.dk });
    }
    T(s, lb, { x: x - 0.1, y: y + 0.52, w: bs + 0.2, h: 0.22, align: "center", fontSize: 9, bold: true, color: dark ? "CFE3DA" : K.dk });
    const t = typeof tg === "string" ? num(tg) : tg;
    if (t >= 1 && t <= plan.length) s.addShape(pres.ShapeType.roundRect, { x, y, w: bs, h: bs, rectRadius: 0.1, fill: { color: "FFFFFF", transparency: 100 }, line: { type: "none" }, hyperlink: { slide: t, tooltip: lb } });
  });
}

// o: {section,title,sub,tabs,tabIdx,tabIds,years,yearIdx,yearIds,dark}
function frame(o) {
  if (!seenSections.has(o.section)) { seenSections.add(o.section); pres.addSection({ title: o.section }); }
  const s = pres.addSlide({ masterName: o.dark ? "DARK" : "LIGHT", sectionTitle: o.section });
  const txt = o.dark ? K.wh : K.dg;
  navButtons(s, o.dark);
  T(s, o.section, { x: 6.4, y: 0.3, w: 3.9, h: 0.4, fontSize: 13, bold: true, color: o.dark ? "9FB8AE" : K.sl });
  if (o.years) chips(s, o.years.map(String), o.yearIdx, o.yearIds, 0.42, 6.1, o.dark, 0.8);
  if (o.tabs) chips(s, o.tabs, o.tabIdx, o.tabIds, 1.05, 12.83, o.dark);
  if (o.title) T(s, o.title, { x: 0.5, y: 1.55, w: 12.33, h: 0.65, fontSize: 30, bold: true, color: txt });
  if (o.sub) T(s, o.sub, { x: 0.5, y: 2.2, w: 12.33, h: 0.32, fontSize: 14, color: o.dark ? "9FB8AE" : K.sl });
  return s;
}

const dTxt = (v) => (v === null ? "—" : (v >= 0 ? "▲ +" : "▼ ") + Math.abs(v).toFixed(1) + "%");

// card with label, big value, unit and up to two deltas; o.hero = gradient green card, o.icon = key in ICONS, o.more = slide id
function card(s, x, y, w, h, o) {
  const dark = o.dark, hero = o.hero, onDark = dark || hero;
  if (hero) s.addImage({ path: A("card-green.png"), x, y, w, h });
  else s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.14, fill: { color: dark ? HEX.card : HEX.lt }, line: dark ? { color: HEX.edge, width: 0.75 } : { type: "none" } });
  const big = o.big || 28;
  const hasD = o.deltas && o.deltas.length;
  const sub = onDark ? "CFE3DA" : K.sl;
  const iconW = o.icon && ICONS[o.icon] ? 0.75 : 0;
  if (iconW) {
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.15, y: y + 0.15, w: 0.6, h: 0.6, rectRadius: 0.1, fill: { color: "FFFFFF", transparency: 100 }, line: { color: HEX.edge, width: 1 } });
    s.addImage({ data: ICONS[o.icon], x: x + 0.27, y: y + 0.27, w: 0.36, h: 0.36 });
  }
  const moreW = o.more ? 1.15 : 0;
  T(s, o.label, { x: x + 0.15 + iconW, y: y + 0.08, w: w - 0.3 - iconW - (hasD ? moreW : 0), h: 0.34, fontSize: 13, color: sub });
  T(s, [{ text: o.value, options: { fontSize: big, bold: true, color: onDark ? K.wh : K.dg } }, { text: o.unit ? "  " + o.unit : "", options: { fontSize: 12, color: sub } }],
    { x: x + 0.15 + iconW, y: y + 0.42, w: w - 0.3 - iconW, h: hasD ? h - 0.95 : h - 0.55, fontSize: big });
  if (o.more) {
    const my = hasD ? y + 0.1 : y + h - 0.42;
    T(s, "للتفاصيل", { x: x + 0.15, y: my, w: moreW, h: 0.3, fontSize: 11, underline: { style: "sng" }, color: onDark ? "8CDB4F" : K.pan, align: "left", rtlMode: true });
    s.addShape(pres.ShapeType.rect, { x: x + 0.15, y: my, w: moreW, h: 0.3, fill: { color: "FFFFFF", transparency: 100 }, line: { type: "none" }, hyperlink: { slide: num(o.more), tooltip: "للتفاصيل" } });
  }
  if (hasD) {
    const area = w - 0.3, dw = area / o.deltas.length;
    o.deltas.forEach((d, i) => {
      const up = d.v !== null && d.v >= 0;
      const col = d.v === null ? sub : up ? (onDark ? "8CDB4F" : K.g) : onDark ? "F7A1BC" : K.pink;
      T(s, [{ text: dTxt(d.v), options: { bold: true, color: col, fontSize: 13 } }, { text: "  " + d.label, options: { color: sub, fontSize: 11 } }],
        { x: x + w - 0.15 - dw * (i + 1), y: y + h - 0.5, w: dw, h: 0.36, fontSize: 12 });
    });
  }
}

function cardRow(s, y, h, items, dark, big) {
  const gap = 0.2, w = (12.33 - gap * (items.length - 1)) / items.length;
  items.forEach((it, i) => card(s, 12.83 - w * (i + 1) - gap * i, y, w, h, Object.assign({ dark, big, hero: !dark && i === 0 }, it)));
}

function chartOpts(o) {
  const stacked = !!o.stacked;
  return {
    x: o.x, y: o.y, w: o.w, h: o.h,
    barDir: o.horiz ? "bar" : "col", barGrouping: stacked ? "stacked" : "clustered", barGapWidthPct: o.gap || 55,
    chartColors: (o.colors || SERIES).slice(0, o.nSeries),
    showLegend: o.legend !== false && o.nSeries > 1, legendPos: "b", legendFontSize: 11, legendColor: HEX.sl, legendFontFace: "+mn-lt",
    catAxisLabelFontSize: 11, catAxisLabelColor: HEX.sl, catAxisLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin",
    valAxisLabelFontSize: 10, valAxisLabelColor: HEX.sl, valAxisLabelFontFace: "+mn-lt", valAxisLabelFormatCode: "#,##0",
    valGridLine: { color: HEX.grid, size: 0.5 }, catGridLine: { style: "none" },
    showValue: o.labels !== false, dataLabelFontSize: 10, dataLabelColor: HEX.dk, dataLabelFontFace: "+mn-lt",
    dataLabelFormatCode: (o.fmt || "#,##0.0") + ";-" + (o.fmt || "#,##0.0") + ";", dataLabelPosition: stacked ? "ctr" : "outEnd",
    showTitle: false,
  };
}

// series: [{name, values(raw riyals)}]
function chart(s, title, cats, series, o) {
  T(s, title, { x: o.x, y: o.y - 0.38, w: o.w, h: 0.34, fontSize: 15, bold: true, color: K.dg });
  const data = series.map((se) => ({ name: se.name, labels: cats, values: se.values.map(M) }));
  s.addChart(pres.charts.BAR, data, chartOpts(Object.assign({ nSeries: series.length }, o)));
}

function table(s, header, body, o) {
  const rev = (a) => a.slice().reverse();
  const bd = { type: "solid", pt: 0.5, color: "D5E0DA" };
  const fs = o.fontSize || 12;
  const mk = (r, hdr, i, tot) => rev(r).map((c, j) => ({ text: String(c), options: {
    rtlMode: true, align: j === r.length - 1 ? "right" : "center", valign: "middle", fontSize: fs, bold: hdr || tot,
    color: hdr ? K.wh : K.dg, fill: { color: hdr ? K.pan : tot ? K.lt : K.wh }, border: [bd, bd, bd, bd], margin: [0.02, 0.08, 0.02, 0.08] } }));
  const all = [mk(header, true)].concat(body.map((r, i) => mk(r, false, i, o.lastTotal && i === body.length - 1)));
  s.addTable(all, { x: o.x, y: o.y, w: o.w, colW: rev(o.colW), rowH: o.rowH || 0.32 });
}

// ---------------------------------------------------------------- icons (react-icons -> png data URIs)
const ICONS = {};
async function loadIcons() {
  try {
    const React = require("react"), { renderToStaticMarkup } = require("react-dom/server"), sharp = require("sharp"), fa = require("react-icons/fa");
    const map = { money: "FaMoneyBillWave", invoice: "FaFileInvoiceDollar", history: "FaHistory", chart: "FaChartLine", balance: "FaBalanceScale", hand: "FaHandHoldingUsd", percent: "FaPercentage", target: "FaBullseye" };
    for (const [k, n] of Object.entries(map)) {
      const svg = renderToStaticMarkup(React.createElement(fa[n], { color: "#8CDB4F", size: 256 }));
      ICONS[k] = "image/png;base64," + (await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer()).toString("base64");
    }
  } catch (e) { console.warn("icons skipped:", e.message); }
}

// section divider in the report's style: green panel, notched title, dotted list of sub-pages
let secNo = 1;
function divider(key, name, items) {
  const no = ++secNo;
  add("dv-" + key, () => {
    const s = frame({ section: name, title: "" });
    s.addImage({ path: A("panel-green.png"), x: 0.5, y: 1.3, w: 12.33, h: 5.55 });
    s.addShape(pres.ShapeType.roundRect, { x: 6.5, y: 1.0, w: 6.8, h: 1.85, rectRadius: 0.4, fill: { color: "FFFFFF" }, line: { type: "none" } });
    T(s, name, { x: 6.8, y: 1.45, w: 5.7, h: 1.2, fontSize: name.length > 22 ? 26 : 36, bold: true, color: K.dg });
    T(s, String(no).padStart(2, "0"), { x: 0.9, y: 4.5, w: 3.5, h: 2.2, fontSize: 110, bold: true, color: K.lime, align: "left", rtlMode: false });
    const y0 = 3.15, st = 0.5;
    s.addShape(pres.ShapeType.line, { x: 12.25, y: y0 + 0.25, w: 0, h: st * (items.length - 1), line: { color: "FFFFFF", width: 1.25 } });
    items.forEach(([lb, id], i) => {
      const y = y0 + i * st;
      s.addShape(pres.ShapeType.ellipse, { x: 12.18, y: y + 0.18, w: 0.14, h: 0.14, fill: { color: "FFFFFF" }, line: { type: "none" } });
      T(s, lb, { x: 5.2, y, w: 6.7, h: 0.5, fontSize: 18, color: K.wh });
      s.addShape(pres.ShapeType.rect, { x: 5.2, y, w: 7.2, h: 0.5, fill: { color: "FFFFFF", transparency: 100 }, line: { type: "none" }, hyperlink: { slide: num(id), tooltip: lb } });
    });
  });
}

const LC = ["048F1D", "A0AAB8", "01523B", "018A9C", "F0648C"];
function lineChart(s, title, cats, series, o) {
  T(s, title, { x: o.x, y: o.y - 0.38, w: o.w, h: 0.34, fontSize: 15, bold: true, color: K.dg });
  // keep only the months that have data for the first series (empty points are not written safely)
  let last = series[0].values.length; while (last > 1 && series[0].values[last - 1] === null) last--;
  cats = cats.slice(0, last);
  const data = series.map((se) => ({ name: se.name, labels: cats, values: se.values.slice(0, last).map((v) => M(v)) }));
  s.addChart(pres.charts.LINE, data, {
    x: o.x, y: o.y, w: o.w, h: o.h, chartColors: (o.colors || LC).slice(0, series.length), lineSize: 3, lineSmooth: true, valAxisMinVal: 0, lineDataSymbol: "circle", lineDataSymbolSize: 7,
    showLegend: series.length > 1, legendPos: "b", legendFontSize: 11, legendColor: HEX.sl, legendFontFace: "+mn-lt",
    catAxisLabelFontSize: 11, catAxisLabelColor: HEX.sl, catAxisLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin",
    valAxisLabelFontSize: 10, valAxisLabelColor: HEX.sl, valAxisLabelFontFace: "+mn-lt", valAxisLabelFormatCode: "#,##0",
    valGridLine: { color: HEX.grid, size: 0.5 }, catGridLine: { style: "none" }, showValue: !!o.labels, dataLabelFontSize: 10, dataLabelColor: HEX.dk,
    dataLabelFontFace: "+mn-lt", dataLabelPosition: "t", dataLabelFormatCode: "#,##0.0", showTitle: false,
  });
}

// ================================================================= 1) cover
add("cover", () => {
  const s = pres.addSlide({ masterName: "BLANK" });
  pres.addSection({ title: "الغلاف" });
  s.addImage({ path: A("panel-teal.png"), x: 6.2, y: 0.4, w: 6.63, h: 6.7 });
  s.addImage({ path: A("cover-photo.png"), x: 0.5, y: 0.4, w: 5.44, h: 6.7 });
  T(s, "الإيرادات", { x: 6.6, y: 1.5, w: 5.8, h: 1.1, fontSize: 60, bold: true, color: K.wh });
  T(s, "حسب السنة والتبويب", { x: 6.6, y: 2.65, w: 5.8, h: 0.8, fontSize: 32, bold: true, color: K.wh });
  T(s, "التقرير السنوي والشهري للإيراد  " + YEAR + "م", { x: 6.6, y: 3.6, w: 5.8, h: 0.5, fontSize: 18, color: "E6F4F5" });
  T(s, "الفعلي مقابل المقدر والمستهدف  •  المفوتر والمحصل  •  المديونية", { x: 6.6, y: 4.2, w: 5.8, h: 0.8, fontSize: 14, color: "E6F4F5", valign: "top" });
  T(s, "المصدر: وزارة السياحة", { x: 6.6, y: 6.2, w: 5.8, h: 0.4, fontSize: 13, color: "E6F4F5" });
});

// ================================================================= 2) index
const SECTIONS = [
  ["المؤشرات الرئيسة", "kpi"], ["الإيراد حتى تاريخه", "dv-td"], ["الإيراد لنهاية السنة", "dv-ye"], ["الإيراد للشهر الحالي", "dv-mo"],
  ["الإيراد حسب القطاع", "dv-sec"], ["ملخص التبويبات: الاستثمارية والجزاءات والتبغ", "dv-sum"], ["الاستقطاع", "dv-cut"],
  ["سيناريوهات الإقفال", "dv-scn"], ["الجزاءات والغرامات الشهرية", "dv-pen"], ["التنفيذ والاسترداد", "dv-enf"],
  ["المفوتر والميزانية والمستهدف", "dv-inv"], ["تبويبات الوزارة: المفوتر والمستهدف", "dv-min"], ["المديونية", "dv-debt"], ["ملاحظات منهجية", "notes"],
];
add("index", () => {
  const s = frame({ section: "الفهرس", title: "الفهرس" });
  const cols = 4, gap = 0.2, w = (12.33 - gap * (cols - 1)) / cols, h = 1.05;
  SECTIONS.forEach(([name, id], i) => {
    const r = Math.floor(i / cols), c = i % cols;
    const x = 12.83 - w * (c + 1) - gap * c, y = 2.4 + r * (h + 0.18);
    s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.14, fill: { color: K.lt }, line: { type: "none" } });
    T(s, String(i + 1), { x: x + w - 0.62, y: y + 0.27, w: 0.5, h: 0.5, shape: pres.ShapeType.rect, fill: { color: K.lime }, color: K.dk, bold: true, fontSize: 18, align: "center" });
    T(s, name, { x: x + 0.12, y: y + 0.1, w: w - 0.9, h: h - 0.2, fontSize: 13, bold: true, color: K.dg });
    s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.14, fill: { color: "FFFFFF", transparency: 100 }, line: { type: "none" }, hyperlink: { slide: num(id), tooltip: name } });
  });
});

// ================================================================= 3) key indicators (dark)
add("kpi", () => {
  const td = rows("الايراد حتى تاريخه").map((r) => ({ adm: r[0], tab: r[1], inv: r[2], act: r[3], prior: r[4], est: r[5], tgt: r[6], ly: r[7] })).filter((r) => r.adm === ADM[0]);
  const T0 = (k) => sum(td, (r) => r[k]);
  const debt = rows("المديونية"), enf = rows("التنفيذ_والاسترداد"), cut = rows("استقطاع").filter((r) => r[6] === YEAR);
  const s = frame({ section: "المؤشرات الرئيسة", title: "أبرز مؤشرات الإيراد للعام " + YEAR, sub: "إجمالي منطقة الرياض (مدينة الرياض + بلديات المنطقة) حتى تاريخه", dark: true });
  cardRow(s, 2.7, 1.9, [
    { icon: "money", more: "td-all", label: "الفعلي حتى تاريخه", value: fm(T0("act")), unit: "مليون ريال", deltas: [{ v: delta(T0("act"), T0("est")), label: "عن المقدر" }, { v: delta(T0("act"), T0("tgt")), label: "عن المستهدف" }] },
    { icon: "invoice", more: "inv-0", label: "المفوتر", value: fm(T0("inv")), unit: "مليون ريال", deltas: [{ v: delta(T0("act"), T0("ly")), label: "الفعلي عن السنة الماضية" }] },
    { icon: "history", more: "td-all", label: "محصل سنوات سابقة", value: fm(T0("prior")), unit: "مليون ريال" },
    { icon: "chart", more: "td-all", label: "محصل السنة الماضية", value: fm(T0("ly")), unit: "مليون ريال" },
  ], true, 24);
  cardRow(s, 4.85, 1.9, [
    { icon: "balance", more: "debt-all", label: "المديونية المتبقية", value: fm(sum(debt, (r) => r[3])), unit: "مليون ريال" },
    { icon: "hand", more: "enf-0", label: "استرداد فعلي (التنفيذ)", value: fm(sum(enf, (r) => r[5])), unit: "مليون ريال" },
    { icon: "percent", more: "cut-2026", label: "الاستقطاع " + YEAR, value: fm(sum(cut, (r) => r[4])), unit: "مليون ريال", deltas: [{ v: (sum(cut, (r) => r[4]) / Math.max(1, sum(cut, (r) => r[3]))) * 100, label: "من الإيراد" }] },
    { icon: "target", more: "td-all", label: "المستهدف حتى تاريخه", value: fm(T0("tgt")), unit: "مليون ريال", deltas: [{ v: delta(T0("est"), T0("tgt")), label: "المقدر عن المستهدف" }] },
  ], true, 24);
});

// ================================================================= 4-6) revenue sections (overview + one slide per tab)
const MEAS = [{ k: "inv", n: "مفوتر" }, { k: "act", n: "فعلي" }, { k: "est", n: "مقدر" }, { k: "tgt", n: "مستهدف" }, { k: "ly", n: "محصل السنة الماضية" }];
function revSection(o) {
  const R = rows(o.key).map(o.map);
  const meas = MEAS.filter((m) => o.inv || m.k !== "inv");
  const tabNames = ["نظرة عامة"].concat(TABS);
  const tabIds = [o.id + "-all"].concat(TABS.map((_, i) => o.id + "-" + i));
  const get = (adm, tab) => R.find((r) => r.adm === adm && r.tab === tab) || {};
  const tot = (adm, k) => sum(TABS, (t) => get(adm, t)[k]);

  divider(o.id, o.name, [["نظرة عامة", tabIds[0]]].concat(TABS.map((t, i) => [t, tabIds[i + 1]])));
  add(o.id + "-all", () => {
    const s = frame({ section: o.name, title: o.name + " — نظرة عامة", sub: o.sub + " — منطقة الرياض", tabs: tabNames, tabIdx: 0, tabIds, years: [o.year], yearIdx: 0 });
    cardRow(s, 2.65, 1.45, meas.map((m, i) => {
      const it = { label: m.n, value: fm(tot(ADM[0], m.k)), unit: "م.ر", big: 26 };
      if (m.k === "act") it.deltas = [{ v: delta(tot(ADM[0], "act"), tot(ADM[0], "est")), label: "عن المقدر" }, { v: delta(tot(ADM[0], "act"), tot(ADM[0], "tgt")), label: "عن المستهدف" }];
      return it;
    }));
    const ms = meas.filter((m) => m.k !== "inv");
    chart(s, "الإيراد حسب التبويب — منطقة الرياض (مليون ريال)", TABS, ms.map((m) => ({ name: m.n, values: TABS.map((t) => get(ADM[0], t)[m.k]) })), { x: 0.5, y: 4.5, w: 12.33, h: 2.45, fmt: "#,##0", colors: SERIES.slice(1) });
  });

  TABS.forEach((tab, ti) => add(o.id + "-" + ti, () => {
    const s = frame({ section: o.name, title: tab, sub: o.name + " — " + o.sub, tabs: tabNames, tabIdx: ti + 1, tabIds, years: [o.year], yearIdx: 0 });
    ADM.forEach((adm, ai) => {
      const r = get(adm, tab);
      card(s, 9.63, 2.65 + ai * 1.45, 3.2, 1.3, { hero: ai === 0, label: adm, value: fm(r.act), unit: "م.ر", big: 24,
        deltas: [{ v: delta(n(r.act), n(r.est)), label: "المقدر" }, { v: delta(n(r.act), n(r.tgt)), label: "المستهدف" }] });
    });
    chart(s, "الفعلي مقابل المقدر والمستهدف حسب التنظيم الإداري (مليون ريال)", ADM, meas.map((m) => ({ name: m.n, values: ADM.map((a) => get(a, tab)[m.k]) })), { x: 0.5, y: 3.0, w: 8.8, h: 3.9, fmt: "#,##0.0", colors: SERIES.slice(o.inv ? 0 : 1) });
  }));
}
revSection({ id: "td", key: "الايراد حتى تاريخه", name: "الإيراد حتى تاريخه", sub: "الفترة: يناير – أغسطس " + YEAR, year: YEAR, inv: true,
  map: (r) => ({ adm: r[0], tab: r[1], inv: r[2], act: r[3], prior: r[4], est: r[5], tgt: r[6], ly: r[7] }) });
revSection({ id: "ye", key: "الايراد لنهاية السنة", name: "الإيراد لنهاية السنة", sub: "تقدير نهاية السنة " + YEAR, year: YEAR, inv: false,
  map: (r) => ({ adm: r[1], tab: r[2], act: r[3], est: r[4], tgt: r[5], ly: r[6] }) });
revSection({ id: "mo", key: "الايراد الشهر الحالي", name: "الإيراد للشهر الحالي", sub: "شهر أغسطس " + YEAR, year: YEAR, inv: false,
  map: (r) => ({ adm: r[1], tab: r[2], act: r[3], est: r[4], tgt: r[5], ly: r[6] }) });

// ================================================================= 7) by sector
{
  const R = rows("القطاع").map((r) => ({ adm: r[0], sec: r[1], tab: r[2], inv: r[3], act: r[4], prior: r[5], col: r[6] }));
  const tabs = ["نظرة عامة"].concat(TABS.slice(0, 4));
  const tids = ["sec-all"].concat([0, 1, 2, 3].map((i) => "sec-" + i));
  const secs = [...new Set(R.map((r) => r.sec))];
  divider("sec", "الإيراد حسب القطاع", [["نظرة عامة", "sec-all"]].concat(TABS.slice(0, 4).map((t, i) => [t, "sec-" + i])));
  add("sec-all", () => {
    const s = frame({ section: "الإيراد حسب القطاع", title: "الإيراد حسب القطاع — نظرة عامة", sub: "إجمالي المحصل لكل قطاع موزعاً على التبويبات", tabs, tabIdx: 0, tabIds: tids, years: [YEAR], yearIdx: 0 });
    chart(s, "إجمالي المحصل حسب القطاع والتبويب (مليون ريال)", secs, TABS.slice(0, 4).map((t) => ({ name: t, values: secs.map((c) => sum(R.filter((r) => r.sec === c && r.tab === t), (r) => r.col)) })),
      { x: 0.5, y: 3.0, w: 12.33, h: 3.9, horiz: true, stacked: true, labels: false, colors: SERIES, gap: 40 });
  });
  TABS.slice(0, 4).forEach((tab, ti) => add("sec-" + ti, () => {
    const s = frame({ section: "الإيراد حسب القطاع", title: tab, sub: "الفواتير والفعلي والمحصل حسب القطاع", tabs, tabIdx: ti + 1, tabIds: tids, years: [YEAR], yearIdx: 0 });
    const f = (c, k) => sum(R.filter((r) => r.sec === c && r.tab === tab), (r) => r[k]);
    chart(s, "إجمالي الفواتير والفعلي والمحصل (مليون ريال)", secs, [{ name: "إجمالي الفواتير", values: secs.map((c) => f(c, "inv")) }, { name: "فعلي", values: secs.map((c) => f(c, "act")) }, { name: "محصل سنوات سابقة", values: secs.map((c) => f(c, "prior")) }, { name: "إجمالي المحصل", values: secs.map((c) => f(c, "col")) }],
      { x: 0.5, y: 3.0, w: 12.33, h: 3.9, horiz: true, labels: false, gap: 35, colors: ["01523B", "048F1D", "8CDB4F", "018A9C"] });
  }));
}

// ================================================================= 8) tab summaries from account sheets
{
  const inv = rows("الايرادات_الاستثمارية");
  const pen = rows("الجزاءات والتبغ والايواء").map((r) => ({ tab: r[0], inv: r[2], act: r[3], prior: r[4], est: r[5], tgt: r[6], ly: r[7] }));
  const L = (t) => pen.filter((r) => r.tab.replace(/\s/g, "").startsWith(t));
  const sums = [
    { name: "إيرادات استثمارية", inv: null, act: sum(inv, (r) => r[2]), prior: n(inv[0][9]), cur: n(inv[0][8]), est: sum(inv, (r) => r[3]), tgt: sum(inv, (r) => r[4]), ly: sum(inv, (r) => r[5]) },
    { name: "جزاءات وغرامات", rr: L("الجزاءات") },
    { name: "التبغ والإيواء", rr: L("التبغ") },
  ].map((o) => (o.rr ? Object.assign(o, { inv: sum(o.rr, (r) => r.inv), act: sum(o.rr, (r) => r.act), prior: sum(o.rr, (r) => r.prior), est: sum(o.rr, (r) => r.est), tgt: sum(o.rr, (r) => r.tgt), ly: sum(o.rr, (r) => r.ly) }) : o));
  const names = sums.map((x) => x.name), tids = sums.map((_, i) => "sum-" + i);
  divider("sum", "ملخص التبويبات", sums.map((x, i) => [x.name, "sum-" + i]));
  sums.forEach((x, i) => add("sum-" + i, () => {
    const s = frame({ section: "ملخص التبويبات", title: x.name, sub: "ملخص التبويب حتى تاريخه — " + YEAR, tabs: names, tabIdx: i, tabIds: tids, years: [YEAR], yearIdx: 0 });
    card(s, 9.63, 2.65, 3.2, 1.45, { hero: true, label: "الفعلي", value: fm(x.act), unit: "م.ر", deltas: [{ v: delta(x.act, x.est), label: "المقدر" }, { v: delta(x.act, x.tgt), label: "المستهدف" }] });
    card(s, 9.63, 4.3, 3.2, 1.3, { label: x.inv === null ? "محصل السنة الحالية" : "المفوتر", value: fm(x.inv === null ? x.cur : x.inv), unit: "م.ر", big: 24 });
    card(s, 9.63, 5.75, 3.2, 1.15, { label: "الفعلي عن السنة الماضية", value: dTxt(delta(x.act, x.ly)), big: 22 });
    const cats = (x.inv === null ? ["محصل السنة الحالية"] : ["مفوتر"]).concat(["فعلي", "محصل سنوات سابقة", "مقدر", "مستهدف", "محصل السنة الماضية"]);
    const vals = [x.inv === null ? x.cur : x.inv, x.act, x.prior, x.est, x.tgt, x.ly];
    chart(s, "مؤشرات التبويب (مليون ريال)", cats, [{ name: x.name, values: vals }], { x: 0.5, y: 3.0, w: 8.8, h: 3.9, colors: ["048F1D"], fmt: "#,##0.0" });
  }));
}

// ================================================================= 9) deductions
{
  const R = rows("استقطاع").map((r) => ({ adm: r[0], tab: r[1], rev: r[3], cut: r[4], year: r[6] }));
  const AM = rows("الاستقطاع_الامانات_الاخرى");
  const years = [...new Set(R.map((r) => r.year))].sort();
  const yids = years.map((y) => "cut-" + y);
  divider("cut", "الاستقطاع", years.map((y) => ["سنة " + y, "cut-" + y]));
  years.forEach((y, yi) => add("cut-" + y, () => {
    const s = frame({ section: "الاستقطاع", title: "الاستقطاع من الإيراد — " + y, sub: "حسب التنظيم الإداري والتبويب", years, yearIdx: yi, yearIds: yids });
    const am = AM.find((r) => r[0] === y);
    const cats = [], rv = [], ct = [];
    ["مدينة الرياض", "بلديات المنطقة"].forEach((a) => ["الرسوم والتراخيص", "التبغ والإيواء"].forEach((t) => {
      const f = R.filter((r) => r.year === y && r.adm === a && r.tab === t);
      cats.push(a + " — " + t); rv.push(sum(f, (r) => r.rev)); ct.push(sum(f, (r) => r.cut));
    }));
    if (am) {
      card(s, 9.63, 2.65, 3.2, 1.0, { hero: true, label: "إيرادات كافة الأمانات", value: fm(am[1]), unit: "م.ر", big: 22 });
      card(s, 9.63, 3.8, 3.2, 1.0, { label: "الإيرادات الفعلية للأمانة", value: fm(am[2]), unit: "م.ر", big: 22 });
      card(s, 9.63, 4.95, 3.2, 1.0, { label: "الاستقطاع لكافة الأمانات", value: fm(am[3]), unit: "م.ر", big: 22 });
      card(s, 9.63, 6.1, 3.2, 0.85, { label: "استقطاع الأمانة", value: fm(am[4]), unit: "م.ر", big: 20 });
    }
    chart(s, "الإيراد والمستقطع (مليون ريال)", cats, [{ name: "الإيراد", values: rv }, { name: "المستقطع", values: ct }], { x: 0.5, y: 3.0, w: 8.8, h: 3.9, horiz: true, colors: ["01523B", "048F1D"], gap: 45 });
  }));
}

// ================================================================= 10) closing scenarios
{
  const mon = ["سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"];
  const names = ["السيناريو الأول", "السيناريو الثاني"], tids = ["scn-0", "scn-1"];
  const keys = ["السناريو_الاول_للاقفال", "السناريو_الثاني_للاقفال"];
  divider("scn", "سيناريوهات الإقفال", names.map((nm, i) => [nm, "scn-" + i]));
  keys.forEach((k, i) => add("scn-" + i, () => {
    const R = rows(k).map((r) => ({ y: r[0], m: r[1], tab: r[2], col: r[3], tgt: r[4] }));
    const s = frame({ section: "سيناريوهات الإقفال", title: names[i] + " للإقفال", sub: "المحصل المتوقع مقابل المستهدف للأشهر المتبقية من " + R[0].y, tabs: names, tabIdx: i, tabIds: tids, years: [R[0].y], yearIdx: 0 });
    const tc = sum(R, (r) => r.col), tt = sum(R, (r) => r.tgt);
    card(s, 9.63, 2.65, 3.2, 1.05, { hero: true, label: "المحصل المتوقع", value: fm(tc), unit: "م.ر", big: 24 });
    card(s, 9.63, 3.85, 3.2, 1.05, { label: "المستهدف", value: fm(tt), unit: "م.ر", big: 24 });
    card(s, 9.63, 5.05, 3.2, 1.0, { label: "المحصل مقابل المستهدف", value: dTxt(delta(tc, tt)), big: 22 });
    chart(s, "المحصل والمستهدف الشهري (مليون ريال)", mon, [{ name: "محصل", values: mon.map((m) => sum(R.filter((r) => r.m === m), (r) => r.col)) }, { name: "مستهدف", values: mon.map((m) => sum(R.filter((r) => r.m === m), (r) => r.tgt)) }], { x: 0.5, y: 3.0, w: 8.8, h: 2.1, colors: ["048F1D", "01523B"], fmt: "#,##0" });
    const tabs6 = TABS.concat(["نقل الملكية"]);
    table(s, ["التبويب", "محصل", "مستهدف", "المحصل عن المستهدف"], tabs6.map((t) => { const f = R.filter((r) => r.tab === t); return [t, fm(sum(f, (r) => r.col)), fm(sum(f, (r) => r.tgt)), dTxt(delta(sum(f, (r) => r.col), sum(f, (r) => r.tgt)))]; }),
      { x: 0.5, y: 5.2, w: 8.8, colW: [2.6, 2.0, 2.0, 2.2], rowH: 0.245, fontSize: 10 });
  }));
}

// ================================================================= 11) monthly penalties
{
  const R = rows("الجزائات والغرامات").map((r) => ({ y: r[0], m: r[1], inv: r[2], past: r[3], same: r[4] }));
  const years = [...new Set(R.map((r) => r.y))].sort(), yids = years.map((y) => "pen-" + y);
  divider("pen", "الجزاءات والغرامات الشهرية", years.map((y) => ["سنة " + y, "pen-" + y]));
  years.forEach((y, yi) => add("pen-" + y, () => {
    const s = frame({ section: "الجزاءات والغرامات الشهرية", title: "الجزاءات والغرامات — " + y, sub: "المفوتر والمحصل شهرياً", years, yearIdx: yi, yearIds: yids });
    const f = R.filter((r) => r.y === y).sort((a, b) => a.m - b.m);
    card(s, 9.63, 2.65, 3.2, 1.3, { hero: true, label: "إجمالي مفوتر السنة", value: fm(sum(f, (r) => r.inv)), unit: "م.ر", big: 24 });
    card(s, 9.63, 4.1, 3.2, 1.3, { label: "محصل عن الفترة الماضية", value: fm(sum(f, (r) => r.past)), unit: "م.ر", big: 24 });
    card(s, 9.63, 5.55, 3.2, 1.3, { label: "محصل عن ذات الشهر", value: fm(sum(f, (r) => r.same)), unit: "م.ر", big: 24 });
    lineChart(s, "الجزاءات والغرامات الشهرية (مليون ريال)", f.map((r) => MONTHS[r.m - 1]), [{ name: "مفوتر الشهر", values: f.map((r) => r.inv) }, { name: "محصل عن الفترة الماضية", values: f.map((r) => r.past) }, { name: "محصل عن ذات الشهر", values: f.map((r) => r.same) }],
      { x: 0.5, y: 3.0, w: 8.8, h: 3.9, labels: false, colors: ["01523B", "8CDB4F", "048F1D"], gap: 40 });
  }));
}

// ================================================================= 12) enforcement & recovery
{
  const R = rows("التنفيذ_والاسترداد").map((r) => ({ adm: r[0], tab: (r[1] || "").replace("الإيرادات الاستثمارية", "إيرادات إستثمارية").replace("جزاءات والغرامات", "جزاءات وغرامات"), inv: r[2], ex: r[3], rec: r[4], real: r[5] }));
  const names = ["الإجمالي", "مدينة الرياض", "بلديات المنطقة"], tids = ["enf-0", "enf-1", "enf-2"];
  divider("enf", "التنفيذ والاسترداد", names.map((nm, i) => [nm, "enf-" + i]));
  names.forEach((nm, i) => add("enf-" + i, () => {
    const s = frame({ section: "التنفيذ والاسترداد", title: "التنفيذ والاسترداد — " + nm, sub: "حسب التبويب — " + YEAR, tabs: names, tabIdx: i, tabIds: tids, years: [YEAR], yearIdx: 0 });
    const F = R.filter((r) => i === 0 || r.adm === nm);
    cardRow(s, 2.65, 1.2, [{ label: "قيمة الفواتير", value: fm(sum(F, (r) => r.inv)), unit: "م.ر", big: 24 }, { label: "المنفذ", value: fm(sum(F, (r) => r.ex)), unit: "م.ر", big: 24 }, { label: "إجمالي الاسترداد", value: fm(sum(F, (r) => r.rec)), unit: "م.ر", big: 24 }, { label: "استرداد فعلي", value: fm(sum(F, (r) => r.real)), unit: "م.ر", big: 24 }]);
    const by = (t, k) => sum(F.filter((r) => r.tab === t), (r) => r[k]);
    chart(s, "قيمة الفواتير والمنفذ والاسترداد حسب التبويب (مليون ريال)", TABS, [{ name: "قيمة الفواتير", values: TABS.map((t) => by(t, "inv")) }, { name: "المنفذ", values: TABS.map((t) => by(t, "ex")) }, { name: "إجمالي استرداد", values: TABS.map((t) => by(t, "rec")) }, { name: "استرداد فعلي", values: TABS.map((t) => by(t, "real")) }],
      { x: 0.5, y: 4.4, w: 12.33, h: 2.55, fmt: "#,##0.0", colors: ["01523B", "048F1D", "8CDB4F", "018A9C"] });
  }));
}

// ================================================================= 13) invoiced & budget
{
  const mn = (a, k) => MONTHS.map((_, i) => { const r = a.find((x) => x[1] === i + 1); return r ? r[k] : null; });
  const inv = rows("المفوتر"), orig = rows("الميزانية الاعتماد الأصلي").filter((r) => r[1]), post = rows("الميزانية بعد  الاعتماد").filter((r) => r[1]), tg = rows("مستهدف المفوتر");
  const names = ["المفوتر والمحصل", "الميزانية والمستهدف"], tids = ["inv-0", "inv-1"];
  divider("inv", "المفوتر والميزانية والمستهدف", names.map((nm, i) => [nm, "inv-" + i]));
  add("inv-0", () => {
    const s = frame({ section: "المفوتر والميزانية والمستهدف", title: "المفوتر والمحصل الشهري", sub: "الفترة: يناير – " + MONTHS[Math.max(...inv.map((r) => r[1])) - 1] + " " + YEAR, tabs: names, tabIdx: 0, tabIds: tids, years: [YEAR], yearIdx: 0 });
    cardRow(s, 2.65, 1.2, [{ label: "إجمالي المفوتر", value: fm(sum(inv, (r) => r[2])), unit: "م.ر", big: 24 }, { label: "صافي المفوتر", value: fm(sum(inv, (r) => r[3])), unit: "م.ر", big: 24 }, { label: "محصل السنة الحالية", value: fm(sum(inv, (r) => r[4])), unit: "م.ر", big: 24 }, { label: "محصل السنوات السابقة", value: fm(sum(inv, (r) => r[5])), unit: "م.ر", big: 24 }]);
    lineChart(s, "المفوتر والمحصل مقابل مستهدف المفوتر (مليون ريال)", MONTHS, [{ name: "مفوتر", values: mn(inv, 2) }, { name: "صافي المفوتر", values: mn(inv, 3) }, { name: "محصل السنة الحالية", values: mn(inv, 4) }, { name: "محصل السنوات السابقة", values: mn(inv, 5) }, { name: "مستهدف المفوتر", values: mn(tg, 2) }],
      { x: 0.5, y: 4.4, w: 12.33, h: 2.55, labels: false, colors: ["01523B", "048F1D", "8CDB4F", "018A9C", "A0AAB8"], gap: 30 });
  });
  add("inv-1", () => {
    const s = frame({ section: "المفوتر والميزانية والمستهدف", title: "الميزانية ومستهدف المفوتر", sub: "الاعتماد الأصلي وبعد الاعتماد " + YEAR, tabs: names, tabIdx: 1, tabIds: tids, years: [YEAR], yearIdx: 0 });
    cardRow(s, 2.65, 1.2, [{ label: "الاعتماد الأصلي (سنوي)", value: fm(sum(orig, (r) => r[2])), unit: "م.ر", big: 24 }, { label: "بعد الاعتماد (سنوي)", value: fm(sum(post, (r) => r[2])), unit: "م.ر", big: 24 }, { label: "مستهدف المفوتر (سنوي)", value: fm(sum(tg, (r) => r[2])), unit: "م.ر", big: 24 }]);
    lineChart(s, "الميزانية الشهرية ومستهدف المفوتر (مليون ريال)", MONTHS, [{ name: "الاعتماد الأصلي", values: mn(orig, 2) }, { name: "بعد الاعتماد", values: mn(post, 2) }, { name: "مستهدف المفوتر", values: mn(tg, 2) }],
      { x: 0.5, y: 4.4, w: 12.33, h: 2.55, labels: false, colors: ["01523B", "048F1D", "A0AAB8"], gap: 40 });
  });
}

// ================================================================= 14) ministry tabs
{
  const MT = [
    ["إيرادات استثمار العقارات البلدية", "استثمار العقارات"], ["إيرادات الرسوم والمقابلات المالية", "الرسوم والمقابلات"], ["إيرادات الجزاءات والغرامات", "الجزاءات والغرامات"],
    ["إيرادات رسم إشغال مرافق الإيواء", "إشغال الإيواء"], ["إيرادات رسم تقديم منتجات التبغ", "منتجات التبغ"], ["الإيرادات المختلفة", "إيرادات مختلفة"],
  ];
  const A = rows("مفوتر تبويبات الوزارة").map((r) => ({ m: r[1], t: r[2], v: r[4] }));
  const B = rows("مستهدف الوزارة").map((r) => ({ m: r[1], t: r[2], v: r[4] }));
  const mv = (L, t) => MONTHS.map((_, i) => { const f = L.filter((r) => r.m === i + 1 && (!t || r.t === t)); return f.length ? sum(f, (r) => r.v) : null; });
  const LM = Math.max(...A.map((r) => r.m)), LN = MONTHS[LM - 1];
  const upTo = (arr) => sum(arr.slice(0, LM), (v) => v);
  const names = ["نظرة عامة"].concat(MT.map((x) => x[1])), tids = ["min-all"].concat(MT.map((_, i) => "min-" + i));
  divider("min", "تبويبات الوزارة", names.map((nm, i) => [nm, tids[i]]));
  add("min-all", () => {
    const s = frame({ section: "تبويبات الوزارة", title: "تبويبات الوزارة — نظرة عامة", sub: "المفوتر حتى " + LN + " مقابل المستهدف", tabs: names, tabIdx: 0, tabIds: tids, years: [YEAR], yearIdx: 0 });
    const ta = upTo(mv(A)), tb = upTo(mv(B));
    cardRow(s, 2.65, 1.2, [{ label: "إجمالي المفوتر (يناير–" + LN + ")", value: fm(ta), unit: "م.ر", big: 24 }, { label: "المستهدف (يناير–" + LN + ")", value: fm(tb), unit: "م.ر", big: 24 }, { label: "المفوتر عن المستهدف", value: dTxt(delta(ta, tb)), big: 24 }, { label: "المستهدف السنوي", value: fm(sum(B, (r) => r.v)), unit: "م.ر", big: 24 }]);
    chart(s, "المفوتر والمستهدف حسب التبويب (مليون ريال)", MT.map((x) => x[1]), [{ name: "مفوتر (يناير–" + LN + ")", values: MT.map((x) => upTo(mv(A, x[0]))) }, { name: "مستهدف (يناير–" + LN + ")", values: MT.map((x) => upTo(mv(B, x[0]))) }, { name: "مستهدف السنة", values: MT.map((x) => sum(B.filter((r) => r.t === x[0]), (r) => r.v)) }],
      { x: 0.5, y: 4.4, w: 12.33, h: 2.55, fmt: "#,##0", colors: ["048F1D", "01523B", "A0AAB8"] });
  });
  MT.forEach((t, i) => add("min-" + i, () => {
    const s = frame({ section: "تبويبات الوزارة", title: t[0], sub: "المفوتر والمستهدف الشهري", tabs: names, tabIdx: i + 1, tabIds: tids, years: [YEAR], yearIdx: 0 });
    const a = mv(A, t[0]), b = mv(B, t[0]);
    cardRow(s, 2.65, 1.2, [{ label: "المفوتر (يناير–" + LN + ")", value: fm(upTo(a)), unit: "م.ر", big: 24 }, { label: "المستهدف (يناير–" + LN + ")", value: fm(upTo(b)), unit: "م.ر", big: 24 }, { label: "المفوتر عن المستهدف", value: dTxt(delta(upTo(a), upTo(b))), big: 24 }, { label: "المستهدف السنوي", value: fm(sum(b, (v) => v)), unit: "م.ر", big: 24 }]);
    lineChart(s, "المفوتر والمستهدف الشهري (مليون ريال)", MONTHS, [{ name: "مفوتر", values: a }, { name: "مستهدف", values: b }], { x: 0.5, y: 4.4, w: 12.33, h: 2.55, fmt: "#,##0.0", colors: ["048F1D", "01523B"], gap: 45 });
  }));
}

// ================================================================= 15) debt
{
  const R = rows("المديونية").map((r) => ({ y: r[0], tab: String(r[1]).trim(), tot: r[2], rem: r[3], paid: r[4], canc: r[5] }));
  const names = ["نظرة عامة"].concat(TABS.length ? ["إيرادات استثمارية", "جزاءات وغرامات", "التبغ والإيواء", "أخرى"] : []);
  const tids = ["debt-all"].concat([0, 1, 2, 3].map((i) => "debt-" + i));
  const cards = (s, F) => cardRow(s, 2.65, 1.2, [{ label: "رصيد المديونية الإجمالي", value: fm(sum(F, (r) => r.tot)), unit: "م.ر", big: 24 }, { label: "المتبقي حتى تاريخه", value: fm(sum(F, (r) => r.rem)), unit: "م.ر", big: 24 }, { label: "ما تم سداده السنة الحالية", value: fm(sum(F, (r) => r.paid)), unit: "م.ر", big: 24 }, { label: "الملغي من المديونية", value: fm(sum(F, (r) => r.canc)), unit: "م.ر", big: 24 }]);
  const SER = (g) => [{ name: "المتبقي", values: g.map((f) => sum(f, (r) => r.rem)) }, { name: "ما تم سداده", values: g.map((f) => sum(f, (r) => r.paid)) }, { name: "الملغي", values: g.map((f) => sum(f, (r) => r.canc)) }];
  divider("debt", "المديونية", names.map((nm, i) => [nm, tids[i]]));
  add("debt-all", () => {
    const s = frame({ section: "المديونية", title: "المديونية — نظرة عامة", sub: "رصيد المديونية حسب التبويب", tabs: names, tabIdx: 0, tabIds: tids, years: [YEAR], yearIdx: 0 });
    cards(s, R);
    const t4 = names.slice(1);
    chart(s, "المديونية حسب التبويب (مليون ريال)", t4, SER(t4.map((t) => R.filter((r) => r.tab === t))), { x: 0.5, y: 4.4, w: 12.33, h: 2.55, stacked: true, labels: false, fmt: "#,##0.0", colors: ["01523B", "048F1D", "A0AAB8"] });
  });
  names.slice(1).forEach((t, i) => add("debt-" + i, () => {
    const s = frame({ section: "المديونية", title: t, sub: "المديونية حسب سنة النشوء", tabs: names, tabIdx: i + 1, tabIds: tids, years: [YEAR], yearIdx: 0 });
    const F = R.filter((r) => r.tab === t).sort((a, b) => a.y - b.y);
    cards(s, F);
    chart(s, "المديونية حسب سنة النشوء (مليون ريال)", F.map((r) => String(r.y)), SER(F.map((r) => [r])), { x: 0.5, y: 4.4, w: 12.33, h: 2.55, stacked: true, labels: false, colors: ["01523B", "048F1D", "A0AAB8"], gap: 35 });
  }));
}

// ================================================================= 16) notes
add("notes", () => {
  const s = frame({ section: "ملاحظات منهجية", title: "ملاحظات منهجية" });
  const notes = [
    ["التنظيم", "جميع البيانات مسجّلة حسب السنة والتبويب؛ البنود التفصيلية (الحسابات) مجمّعة على مستوى التبويب."],
    ["السنة", "الشيتات التي لا تحتوي عمود سنة نُسبت إلى " + YEAR + " (الإيراد للشهر الحالي، القطاع، التنفيذ والاسترداد، الاستثمارية)."],
    ["النسب", "نسب الفعلي عن المقدر والمستهدف والسنة الماضية محسوبة من الأرقام الخام (الفعلي ÷ القاعدة − 1)."],
    ["المجاميع", "إجمالي منطقة الرياض = مدينة الرياض + بلديات المنطقة، ولا تُجمع الثلاثة معاً تفادياً للتكرار."],
    ["السيناريوهات", "بيانات سيناريو الإقفال الأول والثاني متطابقة في الملف المصدر."],
    ["الوحدة", "المبالغ بالمليون ريال سعودي، والأرقام مقرّبة لأقرب خانة عشرية وقد لا تتطابق المجاميع."],
  ];
  const w = 6.0, h = 1.25;
  notes.forEach(([k, v], i) => {
    const c = i % 2, r = Math.floor(i / 2), x = 12.83 - w * (c + 1) - 0.33 * c, y = 2.5 + r * (h + 0.2);
    s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.14, fill: { color: K.lt }, line: { type: "none" } });
    T(s, String(i + 1), { x: x + w - 0.7, y: y + 0.35, w: 0.5, h: 0.5, shape: pres.ShapeType.ellipse, fill: { color: K.g }, color: K.wh, bold: true, fontSize: 16, align: "center" });
    T(s, k, { x: x + 0.15, y: y + 0.1, w: w - 1.0, h: 0.32, fontSize: 14, bold: true, color: K.dg });
    T(s, v, { x: x + 0.15, y: y + 0.43, w: w - 1.0, h: 0.75, fontSize: 12, color: K.sl, valign: "top" });
  });
});

// ================================================================= build
loadIcons()
  .then(() => { plan.forEach((p, i) => { cur = i + 1; p.fn(); }); return pres.writeFile({ fileName: OUT }); })
  .then(() => applyTheme(OUT, THEME)).then(() => console.log("wrote", OUT, plan.length, "slides"));
