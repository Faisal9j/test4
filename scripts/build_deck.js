// Builds revenue-report-2026.pptx (portrait pages, like the Ministry of Tourism statistical report)
// from data/data.json (see scripts/extract_data.py). Data is organised by YEAR and TAB (التبويب), never by line item.
// Palette: Riyadh Municipality green/blue. Run: NODE_PATH=<dir with pptxgenjs> node scripts/build_deck.js [out.pptx]
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.APPLY_THEME || "/root/.claude/skills/synced/642995fb-6fb0-4612-9fc1-a71063ac95e2_2132651f-4df5-4687-9c61-978ed9bca072/pptx/scripts/apply_theme.js");

const ROOT = path.join(__dirname, "..");
const OUT = process.argv[2] || path.join(ROOT, "revenue-report-2026.pptx");
const A = (f) => path.join(ROOT, "assets", f);
const LOGO = A("logo.png");
const D = require(path.join(ROOT, "data", "data.json"));

// ---------------------------------------------------------------- theme
const THEME = {
  name: "Revenue Report", headFontFace: "Vazirmatn", bodyFontFace: "Vazirmatn",
  colors: { dk1: "242424", lt1: "FFFFFF", dk2: "006A46", lt2: "EAF3EF", accent1: "006A46", accent2: "147454", accent3: "74A4D4", accent4: "E4A42C", accent5: "F28B82", accent6: "8CBCAC", hlink: "006A46", folHlink: "5B6B7B" },
};
const HEX = { dk: "242424", g: "006A46", deep: "004D33", lt: "EAF3EF", sl: "5B6B7B", grid: "DCE6E1", card: "147454", edge: "3C9170", blue: "74A4D4", sage: "A4CCBC", gold: "E4A42C" };
const SERIES = ["004D33", "006A46", "A4CCBC", "74A4D4", "E4A42C", "F28B82"];
const LC = ["006A46", "A0AAB8", "74A4D4", "E4A42C", "F28B82"];

const pres = new pptxgen();
const PW = 7.5, PH = 14.88;
pres.defineLayout({ name: "PORTRAIT", width: PW, height: PH });
pres.layout = "PORTRAIT";
pres.rtlMode = true;
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "تقرير الإيرادات 2026";
pres.author = "وزارة السياحة";
const C = pres.SchemeColor;
const K = { dk: C.text1, wh: C.background1, lt: C.background2, g: C.accent1, pan: C.accent2, lime: C.accent6, blue: C.accent3, pink: C.accent5, sl: C.text2 };

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
const dTxt = (v) => (v === null ? "—" : (v >= 0 ? "▲ +" : "▼ ") + Math.abs(v).toFixed(1) + "%");
const pTxt = (v) => (v === null ? "" : (v >= 0 ? "بزيادة " : "بانخفاض ") + Math.abs(v).toFixed(1) + "%");
const YEAR = 2026;

// ---------------------------------------------------------------- plan (two-pass so tabs can link to slides)
const plan = [];
const ids = {};
const add = (id, fn) => { ids[id] = plan.length + 1; plan.push({ id, fn }); };
const num = (id) => ids[id];
const seenSections = new Set();
let cur = 0;

// main groups of the top navigation (like the report's section bar)
const GROUPS = [["المؤشرات", "kpi"], ["حتى تاريخه", "dv-td"], ["نهاية السنة", "dv-ye"], ["الشهر الحالي", "dv-mo"], ["القطاعات", "dv-sec"], ["التبويبات", "dv-sum"], ["الاستقطاع", "dv-cut"],
  ["الإقفال", "dv-scn"], ["الجزاءات", "dv-pen"], ["الاسترداد", "dv-enf"], ["المفوتر", "dv-inv"], ["الوزارة", "dv-min"], ["المديونية", "dv-debt"], ["ملاحظات", "notes"]];
const SEC_GROUP = { td: 1, ye: 2, mo: 3, sec: 4, sum: 5, cut: 6, scn: 7, pen: 8, enf: 9, inv: 10, min: 11, debt: 12 };

const NOTE = "التقرير الإحصائي للإيرادات 2026م\nالمصدر: ملف الإيرادات\nتم تقريب الأرقام لأقرب خانة عشرية، وقد لا تتطابق المجاميع نتيجة لعمليات التقريب";
const footerObjs = (c) => [
  { line: { x: 6.55, y: 13.78, w: 0.55, h: 0, line: { color: c, width: 0.75 } } },
  { text: { text: NOTE, options: { x: 1.4, y: 13.9, w: 5.7, h: 0.8, fontSize: 10, color: c, align: "right", valign: "top", rtlMode: true, margin: 0 } } },
];
pres.defineSlideMaster({ title: "LIGHT", background: { path: A("bg-dots-p.png") }, objects: [{ image: { x: 0.4, y: 0.3, w: 1.75, h: 0.67, path: LOGO } }].concat(footerObjs(HEX.sl)),
  slideNumber: { x: 6.5, y: 13.45, w: 0.6, h: 0.3, fontSize: 10, color: HEX.sl, align: "right" } });
pres.defineSlideMaster({ title: "DARK", background: { path: A("bg-dark-p.png") },
  objects: [{ rect: { x: 0.4, y: 0.3, w: 1.75, h: 0.67, fill: { color: "FFFFFF" }, rectRadius: 0.12 } }, { image: { x: 0.4, y: 0.3, w: 1.75, h: 0.67, path: LOGO } }].concat(footerObjs("CFE3DA")),
  slideNumber: { x: 6.5, y: 13.45, w: 0.6, h: 0.3, fontSize: 10, color: "CFE3DA", align: "right" } });
pres.defineSlideMaster({ title: "BLANK", background: { path: A("bg-dots-p.png") }, objects: [] });

// ---------------------------------------------------------------- building blocks
const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, rtlMode: true, align: "right", margin: 0, valign: "middle", fontSize: 14, color: K.dk }, o));
const hit = (s, x, y, w, h, id, tip, rad) => s.addShape(rad ? pres.ShapeType.roundRect : pres.ShapeType.rect, { x, y, w, h, rectRadius: rad || 0, fill: { color: "FFFFFF", transparency: 100 }, line: { type: "none" }, hyperlink: { slide: typeof id === "string" ? num(id) : id, tooltip: tip } });

// pills; opt: {avail, minW, fs, h}
function chips(s, labels, active, targets, y, right, dark, opt) {
  opt = opt || {};
  const gap = 0.08, avail = opt.avail || 7.1, h = opt.h || 0.36;
  let ws = labels.map((lb) => Math.max(opt.minW || 0.8, lb.length * 0.1 + 0.4));
  const f = Math.min(1, avail / (ws.reduce((a, b) => a + b, 0) + gap * (ws.length - 1)));
  ws = ws.map((w) => w * f);
  const fs = opt.fs || (f < 0.9 ? 10 : 11);
  let x = right;
  labels.forEach((lb, i) => {
    const w = ws[i]; x -= w;
    const on = i === active;
    const o = { x, y, w, h, shape: pres.ShapeType.roundRect, rectRadius: 0.12, align: "center", fontSize: fs, bold: on };
    if (on) Object.assign(o, dark ? { color: K.g, fill: { color: "FFFFFF" } } : { color: K.wh, fill: { color: HEX.g } });
    else Object.assign(o, { color: dark ? "E3F0EA" : K.dk });
    T(s, lb, o);
    if (targets && targets[i] && !on) hit(s, x, y, w, h, targets[i], lb, 0.12);
    x -= gap;
  });
}

// index / previous / next buttons, like the report header
function navButtons(s, dark) {
  const bs = 0.46, defs = [["الفهرس", "index"], ["السابق", cur - 1], ["التالي", cur + 1]];
  defs.forEach(([lb, tg], i) => {
    const x = 6.64 - i * 0.62, y = 0.22, line = dark ? "CFE3DA" : HEX.sl;
    s.addShape(pres.ShapeType.roundRect, { x, y, w: bs, h: bs, rectRadius: 0.1, fill: { color: "FFFFFF", transparency: 100 }, line: { color: line, width: 1 } });
    if (i === 0) [[0, 0, "74A4D4"], [1, 0, "006A46"], [0, 1, "006A46"], [1, 1, "8CBCAC"]].forEach(([a, b, c]) => s.addShape(pres.ShapeType.roundRect, { x: x + 0.085 + a * 0.155, y: y + 0.085 + b * 0.155, w: 0.13, h: 0.13, rectRadius: 0.03, fill: { color: "FFFFFF", transparency: 100 }, line: { color: c, width: 1.5 } }));
    else T(s, i === 1 ? "↑" : "↓", { x, y, w: bs, h: bs, align: "center", fontSize: 16, bold: true, color: dark ? K.wh : K.dk });
    T(s, lb, { x: x - 0.1, y: y + 0.5, w: bs + 0.2, h: 0.22, align: "center", fontSize: 9, bold: true, color: dark ? "CFE3DA" : K.dk });
    const t = typeof tg === "string" ? num(tg) : tg;
    if (t >= 1 && t <= plan.length) hit(s, x, y, bs, bs, t, lb, 0.1);
  });
}

// o: {section, group(0-13), tabs, tabIdx, tabIds, years, yearIdx, yearIds, dark}
function frame(o) {
  if (!seenSections.has(o.section)) { seenSections.add(o.section); pres.addSection({ title: o.section }); }
  const s = pres.addSlide({ masterName: o.dark ? "DARK" : "LIGHT", sectionTitle: o.section });
  navButtons(s, o.dark);
  // section bar: two rows of main groups
  const half = 7;
  [GROUPS.slice(0, half), GROUPS.slice(half)].forEach((row, r) => {
    const ix = row.map((_, i) => i + r * half);
    chips(s, row.map((g) => g[0]), ix.indexOf(o.group), row.map((g) => g[1]), 1.12 + r * 0.4, 7.1, o.dark, { avail: 6.7, minW: 0.7, fs: 10, h: 0.33 });
  });
  if (o.tabs) chips(s, o.tabs, o.tabIdx, o.tabIds, 1.98, 7.1, o.dark, { avail: 6.7, fs: 11, h: 0.34 });
  if (o.years) chips(s, o.years.map(String), o.yearIdx, o.yearIds, 2.4, 7.1, o.dark, { avail: 3, minW: 0.75, fs: 11, h: 0.34 });
  return s;
}

const ICONS = {};
async function loadIcons() {
  try {
    const React = require("react"), { renderToStaticMarkup } = require("react-dom/server"), sharp = require("sharp"), fa = require("react-icons/fa");
    const map = { money: "FaMoneyBillWave", invoice: "FaFileInvoiceDollar", history: "FaHistory", chart: "FaChartLine", balance: "FaBalanceScale", hand: "FaHandHoldingUsd", percent: "FaPercentage", target: "FaBullseye" };
    for (const [k, nm] of Object.entries(map)) {
      const svg = renderToStaticMarkup(React.createElement(fa[nm], { color: "#A4CCBC", size: 256 }));
      ICONS[k] = "image/png;base64," + (await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer()).toString("base64");
    }
  } catch (e) { console.warn("icons skipped:", e.message); }
}

// stat card: o = {label,value,unit,deltas,hero,tone('blue'),dark,icon,more,big}
function card(s, x, y, w, h, o) {
  const dark = o.dark, hero = o.hero, blue = o.tone === "blue", onDark = dark || hero || blue;
  if (hero) s.addImage({ path: A("card-green.png"), x, y, w, h });
  else s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.16, fill: { color: blue ? HEX.blue : dark ? HEX.card : HEX.lt }, line: dark ? { color: HEX.edge, width: 0.75 } : { type: "none" } });
  const sub = blue ? "FFFFFF" : onDark ? "D5E8DF" : K.sl, hasD = o.deltas && o.deltas.length, big = o.big || 28;
  const iconW = o.icon && ICONS[o.icon] ? 0.7 : 0;
  if (iconW) {
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.15, y: y + 0.15, w: 0.5, h: 0.5, rectRadius: 0.1, fill: { color: "FFFFFF", transparency: 100 }, line: { color: HEX.edge, width: 1 } });
    s.addImage({ data: ICONS[o.icon], x: x + 0.26, y: y + 0.26, w: 0.28, h: 0.28 });
  }
  T(s, o.label, { x: x + 0.15 + iconW, y: y + 0.1, w: w - 0.3 - iconW, h: 0.34, fontSize: 12, color: sub });
  T(s, [{ text: o.value, options: { fontSize: big, bold: true, color: onDark ? K.wh : K.dk } }, { text: o.unit ? "  " + o.unit : "", options: { fontSize: 11, color: sub } }],
    { x: x + 0.15, y: y + 0.46, w: w - 0.3, h: Math.max(0.5, h - (hasD ? 1.15 : 0.6)), fontSize: big });
  if (o.more) { T(s, "للتفاصيل", { x: x + 0.15, y: y + h - 0.38, w: 1.0, h: 0.28, fontSize: 10, underline: { style: "sng" }, color: "A4CCBC", align: "left" }); hit(s, x + 0.15, y + h - 0.38, 1.0, 0.28, o.more, "للتفاصيل"); }
  if (hasD) {
    const dw = (w - 0.3) / o.deltas.length;
    o.deltas.forEach((d, i) => {
      const up = d.v !== null && d.v >= 0;
      const col = d.v === null ? sub : blue ? "FFFFFF" : up ? (onDark ? "B9E8CF" : K.g) : onDark ? "F4A29A" : K.pink;
      T(s, [{ text: dTxt(d.v), options: { bold: true, color: col, fontSize: 12 } }, { text: " " + d.label, options: { color: sub, fontSize: 10 } }],
        { x: x + w - 0.15 - dw * (i + 1), y: y + h - 0.5, w: dw, h: 0.36, fontSize: 11 });
    });
  }
}

function chartOpts(o) {
  const stacked = !!o.stacked;
  return {
    x: o.x, y: o.y, w: o.w, h: o.h, barDir: o.horiz ? "bar" : "col", barGrouping: stacked ? "stacked" : "clustered", barGapWidthPct: o.gap || 55,
    chartColors: (o.colors || SERIES).slice(0, o.nSeries),
    showLegend: o.nSeries > 1, legendPos: "b", legendFontSize: 11, legendColor: HEX.sl, legendFontFace: "+mn-lt",
    catAxisLabelFontSize: 11, catAxisLabelColor: HEX.sl, catAxisLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin",
    valAxisLabelFontSize: 10, valAxisLabelColor: HEX.sl, valAxisLabelFontFace: "+mn-lt", valAxisLabelFormatCode: "#,##0",
    valGridLine: { color: HEX.grid, size: 0.5 }, catGridLine: { style: "none" },
    showValue: o.labels !== false, dataLabelFontSize: 10, dataLabelColor: HEX.dk, dataLabelFontFace: "+mn-lt",
    dataLabelFormatCode: (o.fmt || "#,##0.0") + ";-" + (o.fmt || "#,##0.0") + ";", dataLabelPosition: stacked ? "ctr" : "outEnd", showTitle: false,
  };
}
function barChart(s, cats, series, o) {
  s.addChart(pres.charts.BAR, series.map((se) => ({ name: se.name, labels: cats, values: se.values.map(M) })), chartOpts(Object.assign({ nSeries: series.length }, o)));
}
function lineChart(s, cats, series, o) {
  let last = series[0].values.length; while (last > 1 && series[0].values[last - 1] === null) last--; // empty points are not written safely
  cats = cats.slice(0, last);
  s.addChart(pres.charts.LINE, series.map((se) => ({ name: se.name, labels: cats, values: se.values.slice(0, last).map((v) => M(v)) })), {
    x: o.x, y: o.y, w: o.w, h: o.h, chartColors: LC.slice(0, series.length), lineSize: 3, lineSmooth: true, valAxisMinVal: 0, lineDataSymbol: "circle", lineDataSymbolSize: 7,
    showLegend: series.length > 1, legendPos: "b", legendFontSize: 11, legendColor: HEX.sl, legendFontFace: "+mn-lt",
    catAxisLabelFontSize: 10, catAxisLabelColor: HEX.sl, catAxisLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin",
    valAxisLabelFontSize: 10, valAxisLabelColor: HEX.sl, valAxisLabelFontFace: "+mn-lt", valAxisLabelFormatCode: "#,##0",
    valGridLine: { color: HEX.grid, size: 0.5 }, catGridLine: { style: "none" }, showValue: !!o.labels, dataLabelFontSize: 10, dataLabelColor: HEX.dk,
    dataLabelFontFace: "+mn-lt", dataLabelPosition: "t", dataLabelFormatCode: "#,##0.0", showTitle: false,
  });
}
function table(s, header, body, o) {
  const rev = (a) => a.slice().reverse(), bd = { type: "solid", pt: 0.5, color: "D5E0DA" };
  const mk = (r, hdr) => rev(r).map((c, j) => ({ text: String(c), options: { rtlMode: true, align: j === r.length - 1 ? "right" : "center", valign: "middle", fontSize: o.fontSize || 11, bold: hdr,
    color: hdr ? K.wh : K.dk, fill: { color: hdr ? K.pan : K.wh }, border: [bd, bd, bd, bd], margin: [0.02, 0.06, 0.02, 0.06] } }));
  s.addTable([mk(header, true)].concat(body.map((r) => mk(r, false))), { x: o.x, y: o.y, w: o.w, colW: rev(o.colW), rowH: o.rowH || 0.3 });
}

// ---------------------------------------------------------------- generic portrait page (report layout)
// c: frame opts + {photo(0-4), title, hero:{label,value,unit,deltas}, desc, cards:[...], charts:[{title,cats,series,type,opts}], table:{header,body,colW}}
const PHOTO = (i) => "ph" + (i % 5);
function page(c) {
  const s = frame(c);
  const y0 = 2.95;
  s.addImage({ path: A(PHOTO(c.photo || 0) + ".png"), x: 0.4, y: y0, w: 6.7, h: 2.6 });
  T(s, c.title, { x: 3.95, y: y0 + 0.08, w: 3.0, h: 0.85, fontSize: c.title.length > 26 ? 14 : c.title.length > 15 ? 17 : 21, bold: true, color: K.dk });
  let y = y0 + 3.8;
  if (c.hero) card(s, 0.95, y0 + 1.85, 5.9, 1.75, Object.assign({ hero: true, big: 34 }, c.hero));
  if (c.desc) { T(s, c.desc, { x: 0.45, y, w: 6.6, h: 1.05, fontSize: 12, color: K.sl, valign: "top" }); y += 1.1; }
  if (c.cards && c.cards.length) {
    const cw = 3.25, ch = 1.3;
    c.cards.forEach((it, i) => card(s, 7.1 - cw * ((i % 2) + 1) - 0.2 * (i % 2), y + Math.floor(i / 2) * (ch + 0.15), cw, ch, Object.assign({ big: 24, tone: i === 1 || i === 2 ? "blue" : undefined }, it)));
    y += Math.ceil(c.cards.length / 2) * (ch + 0.15) + 0.1;
  }
  const bottom = 13.35;
  if (c.table) { table(s, c.table.header, c.table.body, Object.assign({ x: 0.4, y, w: 6.7 }, c.table)); y += (c.table.body.length + 1) * (c.table.rowH || 0.3) + 0.2; }
  const ch = (c.charts || []);
  ch.forEach((g, i) => {
    const h = Math.max(2.4, (bottom - y) / (ch.length - i) - 0.5);
    T(s, g.title, { x: 0.4, y, w: 6.7, h: 0.36, fontSize: 13, bold: true, color: K.dk });
    const o = Object.assign({ x: 0.4, y: y + 0.4, w: 6.7, h }, g.opts);
    (g.type === "line" ? lineChart : barChart)(s, g.cats, g.series, o);
    y += h + 0.5;
  });
  return s;
}

// section divider: green panel with notched title and a dotted list of sub-pages
let secNo = 1;
function divider(key, name, items) {
  const no = ++secNo;
  add("dv-" + key, () => {
    const s = frame({ section: name, group: SEC_GROUP[key] });
    s.addImage({ path: A("panel-green.png"), x: 0.4, y: 3.0, w: 6.7, h: 10.2 });
    s.addShape(pres.ShapeType.roundRect, { x: 2.6, y: 2.6, w: 5.2, h: 1.7, rectRadius: 0.4, fill: { color: "FFFFFF" }, line: { type: "none" } });
    T(s, name, { x: 2.9, y: 2.95, w: 4.0, h: 1.2, fontSize: name.length > 18 ? 20 : 26, bold: true, color: K.dk });
    T(s, String(no).padStart(2, "0"), { x: 0.8, y: 10.6, w: 3.4, h: 2.4, fontSize: 100, bold: true, color: K.lime, align: "left", rtlMode: false });
    const y0 = 4.9, st = 0.62;
    s.addShape(pres.ShapeType.line, { x: 6.55, y: y0 + 0.28, w: 0, h: st * (items.length - 1), line: { color: "FFFFFF", width: 1.25 } });
    items.forEach(([lb, id], i) => {
      const y = y0 + i * st;
      s.addShape(pres.ShapeType.ellipse, { x: 6.48, y: y + 0.21, w: 0.14, h: 0.14, fill: { color: "FFFFFF" }, line: { type: "none" } });
      T(s, lb, { x: 1.0, y, w: 5.3, h: 0.56, fontSize: 16, color: K.wh });
      hit(s, 0.8, y, 5.7, 0.56, id, lb);
    });
  });
}

// ================================================================= cover
add("cover", () => {
  const s = pres.addSlide({ masterName: "BLANK" });
  pres.addSection({ title: "الغلاف" });
  s.addImage({ path: A("cover-photo.png"), x: 0.3, y: 0.3, w: 6.9, h: 8.5 });
  s.addImage({ path: A("panel-teal.png"), x: 0.3, y: 8.55, w: 6.9, h: 5.9 });
  T(s, "الإيرادات", { x: 0.8, y: 9.3, w: 6.0, h: 1.2, fontSize: 56, bold: true, color: K.wh });
  T(s, "حسب السنة والتبويب", { x: 0.8, y: 10.55, w: 6.0, h: 0.8, fontSize: 30, bold: true, color: K.wh });
  T(s, "التقرير السنوي والشهري للإيراد " + YEAR + "م", { x: 0.8, y: 11.5, w: 6.0, h: 0.5, fontSize: 16, color: "E6F4EF" });
  T(s, "الفعلي مقابل المقدر والمستهدف • المفوتر والمحصل • المديونية", { x: 0.8, y: 12.05, w: 6.0, h: 0.8, fontSize: 13, color: "E6F4EF", valign: "top" });
});

// ================================================================= index
const SECTIONS = [["المؤشرات الرئيسة", "kpi"], ["الإيراد حتى تاريخه", "dv-td"], ["الإيراد لنهاية السنة", "dv-ye"], ["الإيراد للشهر الحالي", "dv-mo"], ["الإيراد حسب القطاع", "dv-sec"],
  ["ملخص التبويبات: الاستثمارية والجزاءات والتبغ", "dv-sum"], ["الاستقطاع", "dv-cut"], ["سيناريوهات الإقفال", "dv-scn"], ["الجزاءات والغرامات الشهرية", "dv-pen"], ["التنفيذ والاسترداد", "dv-enf"],
  ["المفوتر والميزانية والمستهدف", "dv-inv"], ["تبويبات الوزارة: المفوتر والمستهدف", "dv-min"], ["المديونية", "dv-debt"], ["ملاحظات منهجية", "notes"]];
add("index", () => {
  const s = frame({ section: "الفهرس", group: -1 });
  T(s, "الفهرس", { x: 0.4, y: 2.5, w: 6.7, h: 0.7, fontSize: 30, bold: true, color: K.dk });
  s.addShape(pres.ShapeType.line, { x: 6.58, y: 3.5, w: 0, h: 9.6, line: { color: HEX.sage, width: 1.25 } });
  SECTIONS.forEach(([name, id], i) => {
    const y = 3.35 + i * 0.7;
    T(s, String(i + 1), { x: 6.2, y, w: 0.7, h: 0.55, shape: pres.ShapeType.rect, fill: { color: HEX.sage }, color: K.dk, bold: true, fontSize: 16, align: "center" });
    T(s, name, { x: 0.6, y, w: 5.4, h: 0.55, fontSize: 15, bold: true, color: K.dk });
    hit(s, 0.4, y, 6.7, 0.6, id, name);
  });
});

// ================================================================= key indicators (dark)
add("kpi", () => {
  const td = rows("الايراد حتى تاريخه").map((r) => ({ adm: r[0], inv: r[2], act: r[3], prior: r[4], est: r[5], tgt: r[6], ly: r[7] })).filter((r) => r.adm === ADM[0]);
  const T0 = (k) => sum(td, (r) => r[k]);
  const debt = rows("المديونية"), enf = rows("التنفيذ_والاسترداد"), cut = rows("استقطاع").filter((r) => r[6] === YEAR);
  const s = frame({ section: "المؤشرات الرئيسة", group: 0, dark: true });
  T(s, "أبرز مؤشرات الإيراد للعام " + YEAR + "م", { x: 0.4, y: 2.5, w: 6.7, h: 0.9, fontSize: 26, bold: true, color: K.wh });
  T(s, "إجمالي منطقة الرياض (مدينة الرياض + بلديات المنطقة) حتى تاريخه", { x: 0.4, y: 3.35, w: 6.7, h: 0.4, fontSize: 12, color: "CFE3DA" });
  const items = [
    { icon: "money", more: "td-all", label: "الفعلي حتى تاريخه", value: fm(T0("act")), unit: "مليون ريال", deltas: [{ v: delta(T0("act"), T0("est")), label: "عن المقدر" }, { v: delta(T0("act"), T0("tgt")), label: "عن المستهدف" }] },
    { icon: "invoice", more: "inv-0", label: "المفوتر", value: fm(T0("inv")), unit: "مليون ريال", deltas: [{ v: delta(T0("act"), T0("ly")), label: "الفعلي عن السنة الماضية" }] },
    { icon: "history", more: "td-all", label: "محصل سنوات سابقة", value: fm(T0("prior")), unit: "مليون ريال" },
    { icon: "chart", more: "td-all", label: "محصل السنة الماضية", value: fm(T0("ly")), unit: "مليون ريال" },
    { icon: "balance", more: "debt-all", label: "المديونية المتبقية", value: fm(sum(debt, (r) => r[3])), unit: "مليون ريال" },
    { icon: "hand", more: "enf-0", label: "استرداد فعلي (التنفيذ)", value: fm(sum(enf, (r) => r[5])), unit: "مليون ريال" },
    { icon: "percent", more: "cut-2026", label: "الاستقطاع " + YEAR, value: fm(sum(cut, (r) => r[4])), unit: "مليون ريال", deltas: [{ v: (sum(cut, (r) => r[4]) / Math.max(1, sum(cut, (r) => r[3]))) * 100, label: "من الإيراد" }] },
    { icon: "target", more: "td-all", label: "المستهدف حتى تاريخه", value: fm(T0("tgt")), unit: "مليون ريال", deltas: [{ v: delta(T0("est"), T0("tgt")), label: "المقدر عن المستهدف" }] },
  ];
  items.forEach((it, i) => card(s, 7.1 - 3.25 * ((i % 2) + 1) - 0.2 * (i % 2), 3.95 + Math.floor(i / 2) * 2.35, 3.25, 2.2, Object.assign({ dark: true, big: 26 }, it)));
});

// ================================================================= revenue sections (overview + one page per tab)
const MEAS = [{ k: "inv", n: "مفوتر" }, { k: "act", n: "فعلي" }, { k: "est", n: "مقدر" }, { k: "tgt", n: "مستهدف" }, { k: "ly", n: "محصل السنة الماضية" }];
function revSection(o) {
  const R = rows(o.key).map(o.map);
  const meas = MEAS.filter((m) => o.inv || m.k !== "inv");
  const tabNames = ["نظرة عامة"].concat(TABS), tabIds = [o.id + "-all"].concat(TABS.map((_, i) => o.id + "-" + i));
  const get = (adm, tab) => R.find((r) => r.adm === adm && r.tab === tab) || {};
  const tot = (adm, k) => sum(TABS, (t) => get(adm, t)[k]);
  const base = { section: o.name, group: SEC_GROUP[o.id], tabs: tabNames, tabIds, years: [o.year], yearIdx: 0 };
  divider(o.id, o.name, [["نظرة عامة", tabIds[0]]].concat(TABS.map((t, i) => [t, tabIds[i + 1]])));
  add(o.id + "-all", () => {
    const a = tot(ADM[0], "act");
    page(Object.assign({}, base, { tabIdx: 0, photo: SEC_GROUP[o.id], title: o.name,
      hero: { label: "الفعلي — منطقة الرياض", value: fm(a), unit: "مليون ريال", deltas: [{ v: delta(a, tot(ADM[0], "est")), label: "عن المقدر" }, { v: delta(a, tot(ADM[0], "tgt")), label: "عن المستهدف" }] },
      desc: o.sub + ". بلغ الفعلي " + fm(a) + " مليون ريال، " + pTxt(delta(a, tot(ADM[0], "est"))) + " عن المقدر و" + pTxt(delta(a, tot(ADM[0], "tgt"))) + " عن المستهدف.",
      cards: meas.filter((m) => m.k !== "act").map((m) => ({ label: m.n, value: fm(tot(ADM[0], m.k)), unit: "م.ر" })),
      charts: [{ title: "الإيراد حسب التبويب — منطقة الرياض (مليون ريال)", cats: TABS, series: meas.filter((m) => m.k !== "inv").map((m) => ({ name: m.n, values: TABS.map((t) => get(ADM[0], t)[m.k]) })), opts: { horiz: true, labels: false, colors: SERIES.slice(1), gap: 40 } }] }));
  });
  TABS.forEach((tab, ti) => add(o.id + "-" + ti, () => {
    const r0 = get(ADM[0], tab);
    page(Object.assign({}, base, { tabIdx: ti + 1, photo: ti, title: tab,
      hero: { label: "الفعلي — منطقة الرياض", value: fm(r0.act), unit: "مليون ريال", deltas: [{ v: delta(n(r0.act), n(r0.est)), label: "المقدر" }, { v: delta(n(r0.act), n(r0.tgt)), label: "المستهدف" }] },
      desc: tab + " — " + o.sub + ". بلغ فعلي منطقة الرياض " + fm(r0.act) + " مليون ريال، منها " + fm(get(ADM[1], tab).act) + " مليون لمدينة الرياض و" + fm(get(ADM[2], tab).act) + " مليون لبلديات المنطقة.",
      cards: ADM.slice(1).map((a) => ({ label: a, value: fm(get(a, tab).act), unit: "م.ر", deltas: [{ v: delta(n(get(a, tab).act), n(get(a, tab).est)), label: "المقدر" }] })),
      charts: [{ title: "الفعلي مقابل المقدر والمستهدف حسب التنظيم الإداري (مليون ريال)", cats: ADM, series: meas.map((m) => ({ name: m.n, values: ADM.map((a) => get(a, tab)[m.k]) })), opts: { fmt: "#,##0", colors: SERIES.slice(o.inv ? 0 : 1), gap: 40 } }] }));
  }));
}
revSection({ id: "td", key: "الايراد حتى تاريخه", name: "الإيراد حتى تاريخه", sub: "الفترة: يناير – أغسطس " + YEAR, year: YEAR, inv: true, map: (r) => ({ adm: r[0], tab: r[1], inv: r[2], act: r[3], prior: r[4], est: r[5], tgt: r[6], ly: r[7] }) });
revSection({ id: "ye", key: "الايراد لنهاية السنة", name: "الإيراد لنهاية السنة", sub: "تقدير نهاية السنة " + YEAR, year: YEAR, inv: false, map: (r) => ({ adm: r[1], tab: r[2], act: r[3], est: r[4], tgt: r[5], ly: r[6] }) });
revSection({ id: "mo", key: "الايراد الشهر الحالي", name: "الإيراد للشهر الحالي", sub: "شهر أغسطس " + YEAR, year: YEAR, inv: false, map: (r) => ({ adm: r[1], tab: r[2], act: r[3], est: r[4], tgt: r[5], ly: r[6] }) });

// ================================================================= by sector
{
  const R = rows("القطاع").map((r) => ({ adm: r[0], sec: r[1], tab: r[2], inv: r[3], act: r[4], prior: r[5], col: r[6] }));
  const tabs = ["نظرة عامة"].concat(TABS.slice(0, 4)), tids = ["sec-all"].concat([0, 1, 2, 3].map((i) => "sec-" + i));
  const secs = [...new Set(R.map((r) => r.sec))];
  const base = { section: "الإيراد حسب القطاع", group: 4, tabs, tabIds: tids, years: [YEAR], yearIdx: 0 };
  divider("sec", "الإيراد حسب القطاع", [["نظرة عامة", "sec-all"]].concat(TABS.slice(0, 4).map((t, i) => [t, "sec-" + i])));
  add("sec-all", () => page(Object.assign({}, base, { tabIdx: 0, photo: 4, title: "الإيراد حسب القطاع",
    hero: { label: "إجمالي المحصل", value: fm(sum(R, (r) => r.col)), unit: "مليون ريال" },
    desc: "إجمالي المحصل لكل قطاع موزعاً على التبويبات الأربعة، لقطاعات الأمانة وبلديات المنطقة.",
    cards: TABS.slice(0, 4).map((t) => ({ label: t, value: fm(sum(R.filter((r) => r.tab === t), (r) => r.col)), unit: "م.ر" })),
    charts: [{ title: "إجمالي المحصل حسب القطاع والتبويب (مليون ريال)", cats: secs, series: TABS.slice(0, 4).map((t) => ({ name: t, values: secs.map((c) => sum(R.filter((r) => r.sec === c && r.tab === t), (r) => r.col)) })), opts: { horiz: true, stacked: true, labels: false, gap: 40 } }] })));
  TABS.slice(0, 4).forEach((tab, ti) => add("sec-" + ti, () => {
    const F = R.filter((r) => r.tab === tab), f = (c, k) => sum(F.filter((r) => r.sec === c), (r) => r[k]);
    page(Object.assign({}, base, { tabIdx: ti + 1, photo: ti, title: tab,
      hero: { label: "إجمالي المحصل", value: fm(sum(F, (r) => r.col)), unit: "مليون ريال" },
      desc: "الفواتير والفعلي والمحصل لـ" + tab + " حسب القطاع.",
      cards: [{ label: "إجمالي الفواتير", value: fm(sum(F, (r) => r.inv)), unit: "م.ر" }, { label: "الفعلي", value: fm(sum(F, (r) => r.act)), unit: "م.ر" }, { label: "محصل سنوات سابقة", value: fm(sum(F, (r) => r.prior)), unit: "م.ر" }],
      charts: [{ title: "إجمالي الفواتير والفعلي والمحصل (مليون ريال)", cats: secs, series: [{ name: "إجمالي الفواتير", values: secs.map((c) => f(c, "inv")) }, { name: "فعلي", values: secs.map((c) => f(c, "act")) }, { name: "محصل سنوات سابقة", values: secs.map((c) => f(c, "prior")) }, { name: "إجمالي المحصل", values: secs.map((c) => f(c, "col")) }], opts: { horiz: true, labels: false, gap: 35, colors: ["004D33", "006A46", "A4CCBC", "74A4D4"] } }] }));
  }));
}

// ================================================================= tab summaries from account sheets
{
  const inv = rows("الايرادات_الاستثمارية");
  const pen = rows("الجزاءات والتبغ والايواء").map((r) => ({ tab: r[0], inv: r[2], act: r[3], prior: r[4], est: r[5], tgt: r[6], ly: r[7] }));
  const L = (t) => pen.filter((r) => r.tab.replace(/\s/g, "").startsWith(t));
  const sums = [
    { name: "إيرادات استثمارية", inv: null, act: sum(inv, (r) => r[2]), prior: n(inv[0][9]), cur: n(inv[0][8]), est: sum(inv, (r) => r[3]), tgt: sum(inv, (r) => r[4]), ly: sum(inv, (r) => r[5]) },
    { name: "جزاءات وغرامات", rr: L("الجزاءات") }, { name: "التبغ والإيواء", rr: L("التبغ") },
  ].map((o) => (o.rr ? Object.assign(o, { inv: sum(o.rr, (r) => r.inv), act: sum(o.rr, (r) => r.act), prior: sum(o.rr, (r) => r.prior), est: sum(o.rr, (r) => r.est), tgt: sum(o.rr, (r) => r.tgt), ly: sum(o.rr, (r) => r.ly) }) : o));
  const names = sums.map((x) => x.name), tids = sums.map((_, i) => "sum-" + i);
  divider("sum", "ملخص التبويبات", sums.map((x, i) => [x.name, "sum-" + i]));
  sums.forEach((x, i) => add("sum-" + i, () => {
    const first = x.inv === null ? x.cur : x.inv, firstName = x.inv === null ? "محصل السنة الحالية" : "مفوتر";
    page({ section: "ملخص التبويبات", group: 5, tabs: names, tabIdx: i, tabIds: tids, years: [YEAR], yearIdx: 0, photo: i, title: x.name,
      hero: { label: "الفعلي", value: fm(x.act), unit: "مليون ريال", deltas: [{ v: delta(x.act, x.est), label: "المقدر" }, { v: delta(x.act, x.tgt), label: "المستهدف" }] },
      desc: "ملخص " + x.name + " حتى تاريخه: بلغ الفعلي " + fm(x.act) + " مليون ريال، " + pTxt(delta(x.act, x.ly)) + " عن محصل السنة الماضية.",
      cards: [{ label: firstName, value: fm(first), unit: "م.ر" }, { label: "الفعلي عن السنة الماضية", value: dTxt(delta(x.act, x.ly)) }],
      charts: [{ title: "مؤشرات التبويب (مليون ريال)", cats: [firstName, "فعلي", "محصل سنوات سابقة", "مقدر", "مستهدف", "محصل السنة الماضية"], series: [{ name: x.name, values: [first, x.act, x.prior, x.est, x.tgt, x.ly] }], opts: { horiz: true, colors: ["006A46"], fmt: "#,##0.0" } }] });
  }));
}

// ================================================================= deductions
{
  const R = rows("استقطاع").map((r) => ({ adm: r[0], tab: r[1], rev: r[3], cut: r[4], year: r[6] }));
  const AM = rows("الاستقطاع_الامانات_الاخرى");
  const years = [...new Set(R.map((r) => r.year))].sort(), yids = years.map((y) => "cut-" + y);
  divider("cut", "الاستقطاع", years.map((y) => ["سنة " + y, "cut-" + y]));
  years.forEach((y, yi) => add("cut-" + y, () => {
    const am = AM.find((r) => r[0] === y) || [];
    const cats = [], rv = [], ct = [];
    ["مدينة الرياض", "بلديات المنطقة"].forEach((a) => ["الرسوم والتراخيص", "التبغ والإيواء"].forEach((t) => {
      const f = R.filter((r) => r.year === y && r.adm === a && r.tab === t); cats.push(a + " — " + t); rv.push(sum(f, (r) => r.rev)); ct.push(sum(f, (r) => r.cut));
    }));
    page({ section: "الاستقطاع", group: 6, years, yearIdx: yi, yearIds: yids, photo: yi + 1, title: "الاستقطاع من الإيراد " + y,
      hero: { label: "استقطاع الأمانة", value: fm(am[4]), unit: "مليون ريال" },
      desc: "الاستقطاع من الإيراد حسب التنظيم الإداري والتبويب لعام " + y + "؛ استقطاع كافة الأمانات " + fm(am[3]) + " مليون ريال.",
      cards: [{ label: "إيرادات كافة الأمانات", value: fm(am[1]), unit: "م.ر" }, { label: "الإيرادات الفعلية للأمانة", value: fm(am[2]), unit: "م.ر" }, { label: "الاستقطاع لكافة الأمانات", value: fm(am[3]), unit: "م.ر" }],
      charts: [{ title: "الإيراد والمستقطع (مليون ريال)", cats, series: [{ name: "الإيراد", values: rv }, { name: "المستقطع", values: ct }], opts: { horiz: true, colors: ["004D33", "74A4D4"], gap: 45 } }] });
  }));
}

// ================================================================= closing scenarios
{
  const mon = ["سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"], names = ["السيناريو الأول", "السيناريو الثاني"], tids = ["scn-0", "scn-1"];
  divider("scn", "سيناريوهات الإقفال", names.map((nm, i) => [nm, "scn-" + i]));
  ["السناريو_الاول_للاقفال", "السناريو_الثاني_للاقفال"].forEach((k, i) => add("scn-" + i, () => {
    const R = rows(k).map((r) => ({ y: r[0], m: r[1], tab: r[2], col: r[3], tgt: r[4] })), tc = sum(R, (r) => r.col), tt = sum(R, (r) => r.tgt);
    page({ section: "سيناريوهات الإقفال", group: 7, tabs: names, tabIdx: i, tabIds: tids, years: [R[0].y], yearIdx: 0, photo: 2 + i, title: names[i] + " للإقفال",
      hero: { label: "المحصل المتوقع", value: fm(tc), unit: "مليون ريال", deltas: [{ v: delta(tc, tt), label: "عن المستهدف" }] },
      desc: "المحصل المتوقع مقابل المستهدف للأشهر المتبقية من " + R[0].y + ": " + fm(tc) + " مليون ريال مقابل مستهدف " + fm(tt) + " مليون ريال.",
      table: { header: ["التبويب", "محصل", "مستهدف", "عن المستهدف"], body: TABS.concat(["نقل الملكية"]).map((t) => { const f = R.filter((r) => r.tab === t); return [t, fm(sum(f, (r) => r.col)), fm(sum(f, (r) => r.tgt)), dTxt(delta(sum(f, (r) => r.col), sum(f, (r) => r.tgt)))]; }), colW: [2.0, 1.5, 1.5, 1.7], rowH: 0.3 },
      charts: [{ title: "المحصل والمستهدف الشهري (مليون ريال)", cats: mon, series: [{ name: "محصل", values: mon.map((m) => sum(R.filter((r) => r.m === m), (r) => r.col)) }, { name: "مستهدف", values: mon.map((m) => sum(R.filter((r) => r.m === m), (r) => r.tgt)) }], opts: { colors: ["006A46", "74A4D4"], fmt: "#,##0" } }] });
  }));
}

// ================================================================= monthly penalties
{
  const R = rows("الجزائات والغرامات").map((r) => ({ y: r[0], m: r[1], inv: r[2], past: r[3], same: r[4] }));
  const years = [...new Set(R.map((r) => r.y))].sort(), yids = years.map((y) => "pen-" + y);
  divider("pen", "الجزاءات والغرامات الشهرية", years.map((y) => ["سنة " + y, "pen-" + y]));
  years.forEach((y, yi) => add("pen-" + y, () => {
    const f = R.filter((r) => r.y === y).sort((a, b) => a.m - b.m);
    page({ section: "الجزاءات والغرامات الشهرية", group: 8, years, yearIdx: yi, yearIds: yids, photo: yi, title: "الجزاءات والغرامات " + y,
      hero: { label: "إجمالي مفوتر السنة", value: fm(sum(f, (r) => r.inv)), unit: "مليون ريال" },
      desc: "المفوتر والمحصل شهرياً لعام " + y + ": محصل عن الفترة الماضية " + fm(sum(f, (r) => r.past)) + " مليون ريال، ومحصل عن ذات الشهر " + fm(sum(f, (r) => r.same)) + " مليون ريال.",
      cards: [{ label: "محصل عن الفترة الماضية", value: fm(sum(f, (r) => r.past)), unit: "م.ر" }, { label: "محصل عن ذات الشهر", value: fm(sum(f, (r) => r.same)), unit: "م.ر" }],
      charts: [{ title: "الجزاءات والغرامات الشهرية (مليون ريال)", type: "line", cats: f.map((r) => MONTHS[r.m - 1]), series: [{ name: "مفوتر الشهر", values: f.map((r) => r.inv) }, { name: "محصل عن الفترة الماضية", values: f.map((r) => r.past) }, { name: "محصل عن ذات الشهر", values: f.map((r) => r.same) }] }] });
  }));
}

// ================================================================= enforcement & recovery
{
  const R = rows("التنفيذ_والاسترداد").map((r) => ({ adm: r[0], tab: (r[1] || "").replace("الإيرادات الاستثمارية", "إيرادات إستثمارية").replace("جزاءات والغرامات", "جزاءات وغرامات"), inv: r[2], ex: r[3], rec: r[4], real: r[5] }));
  const names = ["الإجمالي", "مدينة الرياض", "بلديات المنطقة"], tids = ["enf-0", "enf-1", "enf-2"];
  divider("enf", "التنفيذ والاسترداد", names.map((nm, i) => [nm, "enf-" + i]));
  names.forEach((nm, i) => add("enf-" + i, () => {
    const F = R.filter((r) => i === 0 || r.adm === nm), by = (t, k) => sum(F.filter((r) => r.tab === t), (r) => r[k]);
    page({ section: "التنفيذ والاسترداد", group: 9, tabs: names, tabIdx: i, tabIds: tids, years: [YEAR], yearIdx: 0, photo: i, title: "التنفيذ والاسترداد — " + nm,
      hero: { label: "قيمة الفواتير", value: fm(sum(F, (r) => r.inv)), unit: "مليون ريال" },
      desc: "المنفذ " + fm(sum(F, (r) => r.ex)) + " مليون ريال، وبلغ الاسترداد الفعلي " + fm(sum(F, (r) => r.real)) + " مليون ريال من إجمالي استرداد " + fm(sum(F, (r) => r.rec)) + " مليون ريال.",
      cards: [{ label: "المنفذ", value: fm(sum(F, (r) => r.ex)), unit: "م.ر" }, { label: "إجمالي الاسترداد", value: fm(sum(F, (r) => r.rec)), unit: "م.ر" }, { label: "استرداد فعلي", value: fm(sum(F, (r) => r.real)), unit: "م.ر" }],
      charts: [{ title: "قيمة الفواتير والمنفذ والاسترداد حسب التبويب (مليون ريال)", cats: TABS, series: [{ name: "قيمة الفواتير", values: TABS.map((t) => by(t, "inv")) }, { name: "المنفذ", values: TABS.map((t) => by(t, "ex")) }, { name: "إجمالي استرداد", values: TABS.map((t) => by(t, "rec")) }, { name: "استرداد فعلي", values: TABS.map((t) => by(t, "real")) }], opts: { horiz: true, labels: false, colors: ["004D33", "006A46", "A4CCBC", "74A4D4"] } }] });
  }));
}

// ================================================================= invoiced & budget
{
  const mn = (a, k) => MONTHS.map((_, i) => { const r = a.find((x) => x[1] === i + 1); return r ? r[k] : null; });
  const inv = rows("المفوتر"), orig = rows("الميزانية الاعتماد الأصلي").filter((r) => r[1]), post = rows("الميزانية بعد  الاعتماد").filter((r) => r[1]), tg = rows("مستهدف المفوتر");
  const names = ["المفوتر والمحصل", "الميزانية والمستهدف"], tids = ["inv-0", "inv-1"], LN = MONTHS[Math.max(...inv.map((r) => r[1])) - 1];
  const base = { section: "المفوتر والميزانية والمستهدف", group: 10, tabs: names, tabIds: tids, years: [YEAR], yearIdx: 0 };
  divider("inv", "المفوتر والميزانية والمستهدف", names.map((nm, i) => [nm, "inv-" + i]));
  add("inv-0", () => page(Object.assign({}, base, { tabIdx: 0, photo: 0, title: "المفوتر والمحصل الشهري",
    hero: { label: "إجمالي المفوتر", value: fm(sum(inv, (r) => r[2])), unit: "مليون ريال" },
    desc: "الفترة: يناير – " + LN + " " + YEAR + ". صافي المفوتر " + fm(sum(inv, (r) => r[3])) + " مليون ريال، ومحصل السنة الحالية " + fm(sum(inv, (r) => r[4])) + " مليون ريال.",
    cards: [{ label: "صافي المفوتر", value: fm(sum(inv, (r) => r[3])), unit: "م.ر" }, { label: "محصل السنة الحالية", value: fm(sum(inv, (r) => r[4])), unit: "م.ر" }, { label: "محصل السنوات السابقة", value: fm(sum(inv, (r) => r[5])), unit: "م.ر" }],
    charts: [{ title: "المفوتر والمحصل الشهري (مليون ريال)", type: "line", cats: MONTHS, series: [{ name: "مفوتر", values: mn(inv, 2) }, { name: "صافي المفوتر", values: mn(inv, 3) }, { name: "محصل السنة الحالية", values: mn(inv, 4) }, { name: "محصل السنوات السابقة", values: mn(inv, 5) }] }] })));
  add("inv-1", () => page(Object.assign({}, base, { tabIdx: 1, photo: 1, title: "الميزانية ومستهدف المفوتر",
    hero: { label: "الاعتماد الأصلي (سنوي)", value: fm(sum(orig, (r) => r[2])), unit: "مليون ريال" },
    desc: "الميزانية بعد الاعتماد " + fm(sum(post, (r) => r[2])) + " مليون ريال، ومستهدف المفوتر السنوي " + fm(sum(tg, (r) => r[2])) + " مليون ريال.",
    cards: [{ label: "بعد الاعتماد (سنوي)", value: fm(sum(post, (r) => r[2])), unit: "م.ر" }, { label: "مستهدف المفوتر (سنوي)", value: fm(sum(tg, (r) => r[2])), unit: "م.ر" }],
    charts: [{ title: "الميزانية الشهرية ومستهدف المفوتر (مليون ريال)", type: "line", cats: MONTHS, series: [{ name: "مستهدف المفوتر", values: mn(tg, 2) }, { name: "الاعتماد الأصلي", values: mn(orig, 2) }, { name: "بعد الاعتماد", values: mn(post, 2) }] }] })));
}

// ================================================================= ministry tabs
{
  const MT = [["إيرادات استثمار العقارات البلدية", "استثمار العقارات"], ["إيرادات الرسوم والمقابلات المالية", "الرسوم والمقابلات"], ["إيرادات الجزاءات والغرامات", "الجزاءات والغرامات"],
    ["إيرادات رسم إشغال مرافق الإيواء", "إشغال الإيواء"], ["إيرادات رسم تقديم منتجات التبغ", "منتجات التبغ"], ["الإيرادات المختلفة", "إيرادات مختلفة"]];
  const A_ = rows("مفوتر تبويبات الوزارة").map((r) => ({ m: r[1], t: r[2], v: r[4] })), B_ = rows("مستهدف الوزارة").map((r) => ({ m: r[1], t: r[2], v: r[4] }));
  const mv = (L, t) => MONTHS.map((_, i) => { const f = L.filter((r) => r.m === i + 1 && (!t || r.t === t)); return f.length ? sum(f, (r) => r.v) : null; });
  const LM = Math.max(...A_.map((r) => r.m)), LN = MONTHS[LM - 1], upTo = (arr) => sum(arr.slice(0, LM), (v) => v);
  const names = ["نظرة عامة"].concat(MT.map((x) => x[1])), tids = ["min-all"].concat(MT.map((_, i) => "min-" + i));
  const base = { section: "تبويبات الوزارة", group: 11, tabs: names, tabIds: tids, years: [YEAR], yearIdx: 0 };
  divider("min", "تبويبات الوزارة", names.map((nm, i) => [nm, tids[i]]));
  add("min-all", () => {
    const ta = upTo(mv(A_)), tb = upTo(mv(B_));
    page(Object.assign({}, base, { tabIdx: 0, photo: 0, title: "تبويبات الوزارة",
      hero: { label: "إجمالي المفوتر (يناير–" + LN + ")", value: fm(ta), unit: "مليون ريال", deltas: [{ v: delta(ta, tb), label: "عن المستهدف" }] },
      desc: "المفوتر حتى " + LN + " " + fm(ta) + " مليون ريال مقابل مستهدف الفترة " + fm(tb) + " مليون ريال، ومستهدف السنة " + fm(sum(B_, (r) => r.v)) + " مليون ريال.",
      cards: [{ label: "المستهدف (يناير–" + LN + ")", value: fm(tb), unit: "م.ر" }, { label: "المستهدف السنوي", value: fm(sum(B_, (r) => r.v)), unit: "م.ر" }],
      charts: [{ title: "المفوتر والمستهدف حسب التبويب (مليون ريال)", cats: MT.map((x) => x[1]), series: [{ name: "مفوتر", values: MT.map((x) => upTo(mv(A_, x[0]))) }, { name: "مستهدف الفترة", values: MT.map((x) => upTo(mv(B_, x[0]))) }, { name: "مستهدف السنة", values: MT.map((x) => sum(B_.filter((r) => r.t === x[0]), (r) => r.v)) }], opts: { horiz: true, fmt: "#,##0", colors: ["006A46", "004D33", "74A4D4"] } }] }));
  });
  MT.forEach((t, i) => add("min-" + i, () => {
    const a = mv(A_, t[0]), b = mv(B_, t[0]);
    page(Object.assign({}, base, { tabIdx: i + 1, photo: i, title: t[0],
      hero: { label: "المفوتر (يناير–" + LN + ")", value: fm(upTo(a)), unit: "مليون ريال", deltas: [{ v: delta(upTo(a), upTo(b)), label: "عن المستهدف" }] },
      desc: t[0] + ": المفوتر " + fm(upTo(a)) + " مليون ريال مقابل مستهدف الفترة " + fm(upTo(b)) + " مليون ريال.",
      cards: [{ label: "المستهدف (يناير–" + LN + ")", value: fm(upTo(b)), unit: "م.ر" }, { label: "المستهدف السنوي", value: fm(sum(b, (v) => v)), unit: "م.ر" }],
      charts: [{ title: "المفوتر والمستهدف الشهري (مليون ريال)", type: "line", cats: MONTHS, series: [{ name: "مفوتر", values: a }, { name: "مستهدف", values: b }] }] }));
  }));
}

// ================================================================= debt
{
  const R = rows("المديونية").map((r) => ({ y: r[0], tab: String(r[1]).trim(), tot: r[2], rem: r[3], paid: r[4], canc: r[5] }));
  const names = ["نظرة عامة", "إيرادات استثمارية", "جزاءات وغرامات", "التبغ والإيواء", "أخرى"], tids = ["debt-all"].concat([0, 1, 2, 3].map((i) => "debt-" + i));
  const base = { section: "المديونية", group: 12, tabs: names, tabIds: tids, years: [YEAR], yearIdx: 0 };
  const cards = (F) => [{ label: "المتبقي حتى تاريخه", value: fm(sum(F, (r) => r.rem)), unit: "م.ر" }, { label: "ما تم سداده السنة الحالية", value: fm(sum(F, (r) => r.paid)), unit: "م.ر" }, { label: "الملغي من المديونية", value: fm(sum(F, (r) => r.canc)), unit: "م.ر" }];
  const SER = (g) => [{ name: "المتبقي", values: g.map((f) => sum(f, (r) => r.rem)) }, { name: "ما تم سداده", values: g.map((f) => sum(f, (r) => r.paid)) }, { name: "الملغي", values: g.map((f) => sum(f, (r) => r.canc)) }];
  divider("debt", "المديونية", names.map((nm, i) => [nm, tids[i]]));
  add("debt-all", () => page(Object.assign({}, base, { tabIdx: 0, photo: 0, title: "المديونية",
    hero: { label: "رصيد المديونية الإجمالي", value: fm(sum(R, (r) => r.tot)), unit: "مليون ريال" },
    desc: "رصيد المديونية حسب التبويب: المتبقي " + fm(sum(R, (r) => r.rem)) + " مليون ريال، وسُدد خلال السنة الحالية " + fm(sum(R, (r) => r.paid)) + " مليون ريال.",
    cards: cards(R),
    charts: [{ title: "المديونية حسب التبويب (مليون ريال)", cats: names.slice(1), series: SER(names.slice(1).map((t) => R.filter((r) => r.tab === t))), opts: { horiz: true, stacked: true, labels: false, colors: ["004D33", "74A4D4", "E4A42C"] } }] })));
  names.slice(1).forEach((t, i) => add("debt-" + i, () => {
    const F = R.filter((r) => r.tab === t).sort((a, b) => a.y - b.y);
    page(Object.assign({}, base, { tabIdx: i + 1, photo: i + 1, title: t,
      hero: { label: "رصيد المديونية الإجمالي", value: fm(sum(F, (r) => r.tot)), unit: "مليون ريال" },
      desc: "المديونية حسب سنة النشوء لتبويب " + t + ": المتبقي " + fm(sum(F, (r) => r.rem)) + " مليون ريال.",
      cards: cards(F),
      charts: [{ title: "المديونية حسب سنة النشوء (مليون ريال)", cats: F.map((r) => String(r.y)), series: SER(F.map((r) => [r])), opts: { stacked: true, labels: false, colors: ["004D33", "74A4D4", "E4A42C"], gap: 35 } }] }));
  }));
}

// ================================================================= notes
add("notes", () => {
  const s = frame({ section: "ملاحظات منهجية", group: 13 });
  T(s, "ملاحظات منهجية", { x: 0.4, y: 2.5, w: 6.7, h: 0.8, fontSize: 28, bold: true, color: K.dk });
  const notes = [
    ["التنظيم", "جميع البيانات مسجّلة حسب السنة والتبويب؛ البنود التفصيلية (الحسابات) مجمّعة على مستوى التبويب."],
    ["السنة", "الشيتات التي لا تحتوي عمود سنة نُسبت إلى " + YEAR + " (الإيراد للشهر الحالي، القطاع، التنفيذ والاسترداد، الاستثمارية)."],
    ["النسب", "نسب الفعلي عن المقدر والمستهدف والسنة الماضية محسوبة من الأرقام الخام (الفعلي ÷ القاعدة − 1)."],
    ["المجاميع", "إجمالي منطقة الرياض = مدينة الرياض + بلديات المنطقة، ولا تُجمع الثلاثة معاً تفادياً للتكرار."],
    ["السيناريوهات", "بيانات سيناريو الإقفال الأول والثاني متطابقة في الملف المصدر."],
    ["الوحدة", "المبالغ بالمليون ريال سعودي، والأرقام مقرّبة لأقرب خانة عشرية وقد لا تتطابق المجاميع."],
  ];
  notes.forEach(([k, v], i) => {
    const y = 3.5 + i * 1.6;
    s.addShape(pres.ShapeType.roundRect, { x: 0.4, y, w: 6.7, h: 1.45, rectRadius: 0.16, fill: { color: HEX.lt }, line: { type: "none" } });
    T(s, String(i + 1), { x: 6.35, y: y + 0.4, w: 0.6, h: 0.6, shape: pres.ShapeType.ellipse, fill: { color: HEX.g }, color: K.wh, bold: true, fontSize: 16, align: "center" });
    T(s, k, { x: 0.6, y: y + 0.12, w: 5.5, h: 0.35, fontSize: 14, bold: true, color: K.dk });
    T(s, v, { x: 0.6, y: y + 0.5, w: 5.5, h: 0.85, fontSize: 12, color: K.sl, valign: "top" });
  });
});

// ================================================================= build
loadIcons()
  .then(() => { plan.forEach((p, i) => { cur = i + 1; p.fn(); }); return pres.writeFile({ fileName: OUT }); })
  .then(() => applyTheme(OUT, THEME)).then(() => console.log("wrote", OUT, plan.length, "slides"));
