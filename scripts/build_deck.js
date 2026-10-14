// Builds revenue-report-2026.pptx — the Riyadh Municipality (Amana) report layout, in the Ministry of Tourism report style,
// from data/data.json (see scripts/extract_data.py). Data is organised by YEAR and TAB (التبويب), never by line item.
// Run: NODE_PATH=<dir with pptxgenjs> node scripts/build_deck.js [out.pptx]
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.APPLY_THEME || "/root/.claude/skills/synced/642995fb-6fb0-4612-9fc1-a71063ac95e2_2132651f-4df5-4687-9c61-978ed9bca072/pptx/scripts/apply_theme.js");

const ROOT = path.join(__dirname, "..");
const OUT = process.argv[2] || path.join(ROOT, "revenue-report-2026.pptx");
const A = (f) => path.join(ROOT, "assets", f);
const D = require(path.join(ROOT, "data", "data.json"));

// ---------------------------------------------------------------- theme (colours sampled from the Amana presentation)
const THEME = {
  name: "Amana Revenue Report", headFontFace: "Vazirmatn", bodyFontFace: "Vazirmatn",
  colors: { dk1: "1F2A37", lt1: "FFFFFF", dk2: "394B5F", lt2: "F1F5F3", accent1: "006A46", accent2: "74A4D4", accent3: "E4A42C", accent4: "0B6B73", accent5: "E5736B", accent6: "8CBCAC", hlink: "006A46", folHlink: "5B6B7B" },
};
const HEX = { g: "006A46", blue: "74A4D4", gold: "E4A42C", teal: "0B6B73", pink: "E5736B", slate: "394B5F", grey: "6B7785", track: "EEF2F4", card: "147454", edge: "3C9170", sage: "A4CCBC", lt: "F1F5F3" };
const CARD = { g: HEX.g, blue: HEX.blue, gold: HEX.gold, teal: HEX.teal };
const LC = ["006A46", "9AA5B1", "74A4D4", "E4A42C", "E5736B"];

const pres = new pptxgen();
const W = 13.33, H = 9.43;
pres.defineLayout({ name: "A4L", width: W, height: H });
pres.layout = "A4L";
pres.rtlMode = true;
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "تقرير الإيرادات 2026";
pres.author = "أمانة منطقة الرياض";
const C = pres.SchemeColor;
const K = { dk: C.text1, tx: C.text2, wh: C.background1, g: C.accent1 };

// ---------------------------------------------------------------- data helpers
const rows = (k) => D[k].slice(1).filter((r) => r.some((c) => c !== null));
const n = (v) => (typeof v === "number" ? v : 0);
const sum = (a, f) => a.reduce((t, r) => t + n(f(r)), 0);
const TABS = ["إيرادات إستثمارية", "الرسوم والتراخيص", "جزاءات وغرامات", "التبغ والإيواء", "أخرى"];
const ADM = ["منطقة الرياض", "مدينة الرياض", "بلديات المنطقة"];
const MONTHS = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"];
const M = (v) => Math.round((n(v) / 1e6) * 10) / 10;
const fm = (v) => (n(v) / 1e6).toLocaleString("en-US", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
const fk = (v) => Math.round(n(v) / 1e3).toLocaleString("en-US");
const delta = (a, b) => (b ? (a / b - 1) * 100 : null);
const dTxt = (v) => (v === null ? "—" : (v >= 0 ? "▲ " : "▼ ") + Math.abs(v).toFixed(1) + "%");
const pTxt = (v) => (v === null ? "" : (v >= 0 ? "بزيادة " : "بانخفاض ") + Math.abs(v).toFixed(1) + "%");
const YEAR = 2026;

// ---------------------------------------------------------------- plan (two-pass so pills can link to slides)
const plan = [], ids = {}, seen = new Set();
const add = (id, fn) => { ids[id] = plan.length + 1; plan.push({ id, fn }); };
const num = (id) => ids[id];
let cur = 0;

const GROUPS = [["المقدمة", "intro"], ["المؤشرات الرئيسة", "kpi"], ["الإيراد", "dv-rev"], ["القطاعات", "dv-sec"], ["التبويبات", "dv-tab"], ["التحصيل", "dv-col"], ["المديونية", "dv-debt"], ["إخلاء المسؤولية", "disc"]];
const NOTE = "التقرير الإحصائي للإيرادات 2026م | أمانة منطقة الرياض\nالمصدر: ملف الإيرادات\nتم تقريب الأرقام لأقرب خانة عشرية، وقد لا تتطابق المجاميع نتيجة لعمليات التقريب";
const footer = (c) => [
  { line: { x: 0.5, y: 8.93, w: 0.4, h: 0, line: { color: c, width: 0.75 } } },
  { text: { text: NOTE, options: { x: 5.2, y: 8.62, w: 7.63, h: 0.62, fontSize: 9, color: c, align: "right", valign: "top", rtlMode: true, margin: 0 } } },
];
pres.defineSlideMaster({ title: "LIGHT", background: { path: A("bg-dots.png") }, objects: [{ image: { x: 12.0, y: 0.22, w: 0.82, h: 0.89, path: A("logo-amana.png") } }].concat(footer("8A95A3")),
  slideNumber: { x: 0.5, y: 8.97, w: 0.4, h: 0.25, fontSize: 9, color: "5B6B7B", align: "left", bold: true } });
pres.defineSlideMaster({ title: "GREEN", background: { path: A("bg-green.png") },
  objects: [{ rect: { x: 12.0, y: 0.22, w: 0.82, h: 0.89, fill: { color: "FFFFFF" }, rectRadius: 0.14 } }, { image: { x: 12.0, y: 0.22, w: 0.82, h: 0.89, path: A("logo-amana.png") } }].concat(footer("CFE3DA")),
  slideNumber: { x: 0.5, y: 8.97, w: 0.4, h: 0.25, fontSize: 9, color: "CFE3DA", align: "left", bold: true } });
pres.defineSlideMaster({ title: "BLANK", background: { path: A("bg-dots.png") }, objects: [] });

// ---------------------------------------------------------------- building blocks
const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, rtlMode: true, align: "right", margin: 0, valign: "middle", fontSize: 14, color: K.dk }, o));
const hit = (s, x, y, w, h, id, tip, rad) => s.addShape(rad ? pres.ShapeType.roundRect : pres.ShapeType.rect, { x, y, w, h, rectRadius: rad || 0, fill: { color: "FFFFFF", transparency: 100 }, line: { type: "none" }, hyperlink: { slide: typeof id === "string" ? num(id) : id, tooltip: tip } });
const shadow = () => ({ type: "outer", color: "000000", opacity: 0.16, blur: 9, offset: 3, angle: 90 });

// pill row, laid out right-to-left
function chips(s, labels, active, targets, y, right, dark, o) {
  o = o || {};
  const gap = o.gap || 0.1, h = o.h || 0.42, avail = o.avail || 9.6;
  let ws = labels.map((lb) => Math.max(o.minW || 0.9, lb.length * (o.cw || 0.095) + (o.pad || 0.34)));
  const f = Math.min(1, avail / (ws.reduce((a, b) => a + b, 0) + gap * (ws.length - 1)));
  ws = ws.map((w) => w * f);
  const fs = Math.round((o.fs || 13) * Math.max(f, 0.85));
  let x = right;
  labels.forEach((lb, i) => {
    const w = ws[i]; x -= w;
    const on = i === active;
    const st = { x, y, w, h, shape: pres.ShapeType.roundRect, rectRadius: h / 2, align: "center", fontSize: fs, bold: on };
    if (on) Object.assign(st, dark ? { color: K.g, fill: { color: "FFFFFF" } } : { color: K.wh, fill: { color: HEX.g } });
    else Object.assign(st, { color: dark ? "E3F0EA" : HEX.slate, bold: !o.soft });
    T(s, lb, st);
    if (targets && targets[i] && !on) hit(s, x, y, w, h, targets[i], lb, h / 2);
    x -= gap;
  });
}

// mail / down / up / index buttons (left side of the header)
function navButtons(s, dark) {
  const bs = 0.46, defs = [["mail", "disc"], ["↓", cur + 1], ["↑", cur - 1], ["grid", "index"]];
  defs.forEach(([g, tg], i) => {
    const x = 0.5 + i * 0.58, y = 0.5, line = dark ? "CFE3DA" : "8A95A3";
    s.addShape(pres.ShapeType.roundRect, { x, y, w: bs, h: bs, rectRadius: 0.1, fill: { color: "FFFFFF", transparency: 100 }, line: { color: line, width: 1 } });
    if (g === "grid") [[0, 0], [1, 0], [0, 1], [1, 1]].forEach(([a, b]) => s.addShape(pres.ShapeType.rect, { x: x + 0.12 + a * 0.12, y: y + 0.12 + b * 0.12, w: 0.1, h: 0.1, fill: { color: dark ? "E3F0EA" : HEX.slate }, line: { type: "none" } }));
    else if (g === "mail") { s.addShape(pres.ShapeType.rect, { x: x + 0.11, y: y + 0.14, w: 0.24, h: 0.17, fill: { color: "FFFFFF", transparency: 100 }, line: { color: dark ? "E3F0EA" : HEX.slate, width: 1 } }); }
    else T(s, g, { x, y, w: bs, h: bs, align: "center", fontSize: 14, bold: true, color: dark ? K.wh : K.tx });
    const t = typeof tg === "string" ? num(tg) : tg;
    if (t >= 1 && t <= plan.length) hit(s, x, y, bs, bs, t, "", 0.1);
  });
}

// o: {section, group, title, sub, switcher:{labels,idx,ids}, tabs, tabIdx, tabIds, years, yearIdx, yearIds, dark}
function frame(o) {
  if (o.section && !seen.has(o.section)) { seen.add(o.section); pres.addSection({ title: o.section }); }
  const s = pres.addSlide({ masterName: o.dark ? "GREEN" : "LIGHT", sectionTitle: o.section || "الغلاف" });
  navButtons(s, o.dark);
  chips(s, GROUPS.map((g) => g[0]), o.group, GROUPS.map((g) => g[1]), 0.48, 11.85, o.dark, { avail: 8.9, minW: 0.8, cw: 0.095, fs: 13, h: 0.46, soft: false });
  if (o.title) T(s, o.title, { x: 3.2, y: 1.2, w: 9.63, h: 0.85, fontSize: 36, bold: true, color: o.dark ? K.wh : K.tx });
  if (o.sub) T(s, o.sub, { x: 3.2, y: 2.0, w: 9.63, h: 0.32, fontSize: 12, color: o.dark ? "CFE3DA" : HEX.grey });
  let y = 2.42;
  const left = o.years ? (o.years.length > 1 ? 3.8 : 1.1) : 0;
  if (o.tabs) { chips(s, o.tabs, o.tabIdx, o.tabIds, y, 12.83, o.dark, { avail: 12.33 - left - 0.2, fs: 12, h: 0.38, minW: 0.7 }); y += 0.5; }
  else if (o.years) y += 0.5;
  const rowY = 2.42;
  if (o.switcher) chips(s, o.switcher.labels, o.switcher.idx, o.switcher.ids, 1.32, 3.55, o.dark, { avail: 3.05, fs: 11, h: 0.36, minW: 0.7, pad: 0.26 });
  if (o.years) chips(s, o.years.map(String), o.yearIdx, o.yearIds, rowY, 0.5 + (o.years.length > 1 ? Math.min(3.6, o.years.length * 1.05) : 0.9), o.dark, { avail: o.years.length > 1 ? 3.6 : 0.9, fs: 11, h: 0.38, minW: 0.8, pad: 0.3 });
  s._y = y + 0.08;
  return s;
}

const ICONS = {};

// solid card (deck style): colored rounded card, label / big number / delta; o.tone g|blue|gold|teal, o.light = white body with coloured header
function hero(s, x, y, w, h, o) {
  const col = CARD[o.tone || "g"];
  const hasD0 = !!o.delta, hasN0 = !!o.note;
  const big = Math.min(o.big || 52, Math.max(18, Math.floor((h - 0.62 - (hasD0 ? 0.45 : 0) - (hasN0 ? 0.35 : 0)) * 58)));
  if (o.light) {
    s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.22, fill: { color: "FFFFFF" }, line: { color: "E3E9E6", width: 0.75 }, shadow: shadow() });
    s.addShape(pres.ShapeType.round2SameRect, { x, y, w, h: 0.85, rectRadius: 0.22, fill: { color: col }, line: { type: "none" } });
    T(s, o.label, { x: x + 0.25, y: y + 0.08, w: w - 0.5, h: 0.45, fontSize: 18, bold: true, color: K.wh });
    if (o.sub) T(s, o.sub, { x: x + 0.25, y: y + 0.5, w: w - 0.5, h: 0.28, fontSize: 10.5, color: "E8F2EE" });
    T(s, [{ text: o.value, options: { fontSize: big, bold: true, color: K.tx } }, { text: o.unit ? "  " + o.unit : "", options: { fontSize: 14, color: K.tx } }], { x: x + 0.25, y: y + 1.0, w: w - 0.5, h: 0.95, fontSize: big });
    if (o.delta) T(s, [{ text: dTxt(o.delta.v), options: { bold: true, color: o.delta.v === null || o.delta.v >= 0 ? HEX.g : HEX.pink } }, { text: "  " + o.delta.label, options: { color: HEX.grey } }], { x: x + 0.25, y: y + 1.95, w: w - 0.5, h: 0.32, fontSize: 12 });
    if (o.note) T(s, o.note, { x: x + 0.25, y: y + 2.27, w: w - 0.5, h: 0.3, fontSize: 11, color: HEX.grey });
    return;
  }
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.28, fill: { color: o.dark ? HEX.card : col }, line: o.dark ? { color: HEX.edge, width: 0.75 } : { type: "none" }, shadow: o.dark ? undefined : shadow() });
  T(s, o.label, { x: x + 0.3, y: y + 0.18, w: w - 0.6, h: 0.5, fontSize: o.lfs || 20, bold: true, color: K.wh });
  const hasD = !!o.delta, hasN = !!o.note;
  T(s, [{ text: o.value, options: { fontSize: big, bold: true, color: K.wh } }, { text: o.unit ? "  " + o.unit : "", options: { fontSize: 15, color: "E8F2EE" } }],
    { x: x + 0.3, y: y + 0.62, w: w - 0.6, h: Math.max(0.5, h - 0.62 - (hasD ? 0.45 : 0) - (hasN ? 0.35 : 0)), fontSize: big });
  if (hasD) T(s, [{ text: dTxt(o.delta.v), options: { bold: true, color: o.dark ? (o.delta.v >= 0 ? "B9E8CF" : "F4A29A") : K.wh } }, { text: "  " + o.delta.label, options: { color: "E8F2EE" } }], { x: x + 0.3, y: y + h - 0.45 - (hasN ? 0.35 : 0), w: w - 0.6, h: 0.35, fontSize: 12.5 });
  if (hasN) T(s, o.note, { x: x + 0.3, y: y + h - 0.45, w: w - 0.6, h: 0.32, fontSize: 11, color: "D5E8DF" });
}

// vertical stack of hero cards that always fits above the footer (y from s._y to ~8.5)
function stack(s, x, w, items) {
  const hs = { 2: [2.4, 1.85], 3: [2.0, 1.5, 1.5], 4: [1.4, 1.3, 1.3, 1.3] }[items.length], gap = 0.15;
  let y = s._y + 0.1;
  items.forEach((it, i) => { hero(s, x, y, w, hs[i], Object.assign({ big: i === 0 ? 44 : 34, lfs: i === 0 ? 18 : 15 }, it)); y += hs[i] + gap; });
}

// horizontal pill bars: items = [{label, sub, value(number), text, color}]
function barList(s, x, y, w, items, o) {
  o = o || {};
  const rh = o.rh || 0.58, mx = Math.max(...items.map((i) => i.value), 1), lw = o.lw || 3.0, vw = 0.95;
  items.forEach((it, i) => {
    const yy = y + i * rh;
    T(s, it.label, { x: x + w - lw, y: yy, w: lw, h: it.sub ? rh * 0.55 : rh - 0.1, fontSize: o.fs || 13, bold: true, color: K.tx, valign: it.sub ? "bottom" : "middle" });
    if (it.sub) T(s, it.sub, { x: x + w - lw, y: yy + rh * 0.55, w: lw, h: rh * 0.35, fontSize: 10, color: "9AA5B1", valign: "top" });
    const tx = x + vw, tw = w - lw - vw - 0.15, bh = 0.24, by = yy + (rh - 0.1) / 2 - bh / 2;
    s.addShape(pres.ShapeType.roundRect, { x: tx, y: by, w: tw, h: bh, rectRadius: bh / 2, fill: { color: HEX.track }, line: { type: "none" } });
    const fw = Math.max(0.2, tw * (it.value / mx));
    s.addShape(pres.ShapeType.roundRect, { x: tx + tw - fw, y: by, w: fw, h: bh, rectRadius: bh / 2, fill: { color: it.color || HEX.g }, line: { type: "none" } });
    T(s, it.text, { x, y: yy, w: vw - 0.05, h: rh - 0.1, fontSize: o.fs || 13, bold: true, color: K.dk, align: "left" });
  });
}

// heat-map table (green header, shaded cells); header/body arrays are in reading order (first column = row label)
function heat(s, header, body, o) {
  const rev = (a) => a.slice().reverse(), bd = { type: "solid", pt: 0.5, color: "FFFFFF" };
  const fs = o.fs || 11, nc = header.length;
  const lo = [0xEA, 0xF4, 0xEF], hi = [0x6B, 0xAA, 0x94];
  const shade = (t) => [0, 1, 2].map((k) => Math.round(lo[k] + (hi[k] - lo[k]) * t).toString(16).padStart(2, "0")).join("").toUpperCase();
  const num_ = (r, j) => (typeof r[j] === "object" ? r[j].v : null);
  // per-row range over shaded columns (o.shade = [first,last]); otherwise per column
  const cells = body.map((r, i) => rev(r.map((c, j) => {
    const txt = typeof c === "object" ? c.t : String(c), v = typeof c === "object" ? c.v : null;
    let fill = "FFFFFF", bold = j === 0 || (o.totalRow && i === body.length - 1), color = K.dk;
    if (v !== null && o.shade && j >= o.shade[0] && j <= o.shade[1]) {
      const vals = o.byCol ? body.map((b) => b[j].v) : r.slice(o.shade[0], o.shade[1] + 1).map((c2) => c2.v);
      const mn = Math.min(...vals), mxv = Math.max(...vals); fill = shade(mxv > mn ? (v - mn) / (mxv - mn) : 0.4);
    } else if (o.totalRow && i === body.length - 1) fill = "E3F1EA"; else if (j === 0) fill = "F6F9F7";
    return { text: txt, options: { rtlMode: true, align: j === 0 ? "right" : "center", valign: "middle", fontSize: fs, bold, color, fill: { color: fill }, border: [bd, bd, bd, bd], margin: [0.02, 0.07, 0.02, 0.07] } };
  })));
  const head = rev(header).map((c, j) => ({ text: c, options: { rtlMode: true, align: j === nc - 1 ? "right" : "center", valign: "middle", fontSize: fs, bold: true, color: K.wh, fill: { color: HEX.g }, border: [bd, bd, bd, bd], margin: [0.02, 0.07, 0.02, 0.07] } }));
  s.addTable([head].concat(cells), { x: o.x, y: o.y, w: o.w, colW: rev(o.colW), rowH: o.rowH || 0.34 });
}
const V = (v, t) => ({ v, t: t === undefined ? fm(v) : t });

// line chart: series [{name, values(riyals|null), dash}]
function lineChart(s, cats, series, o) {
  let last = series[0].values.length; while (last > 1 && series[0].values[last - 1] === null) last--;
  cats = cats.slice(0, last);
  s.addChart(pres.charts.LINE, series.map((se) => ({ name: se.name, labels: cats, values: se.values.slice(0, last).map((v) => M(v)) })), {
    x: o.x, y: o.y, w: o.w, h: o.h, chartColors: (o.colors || LC).slice(0, series.length), lineSize: 3, lineSmooth: true, valAxisMinVal: 0, lineDataSymbol: "circle", lineDataSymbolSize: 7,
    lineDash: series.map((se, i) => (se.dash ? "dash" : "solid")),
    showLegend: series.length > 1, legendPos: "t", legendFontSize: 11, legendColor: HEX.grey, legendFontFace: "+mn-lt",
    catAxisLabelFontSize: 10, catAxisLabelColor: HEX.grey, catAxisLabelFontFace: "+mn-lt", catAxisOrientation: "maxMin",
    valAxisHidden: !!o.hideVal, valAxisLabelFontSize: 10, valAxisLabelColor: HEX.grey, valAxisLabelFontFace: "+mn-lt", valAxisLabelFormatCode: "#,##0",
    valGridLine: { color: "E6ECE9", size: 0.5 }, catGridLine: { style: "none" }, showValue: o.labels !== false, dataLabelFontSize: 9, dataLabelColor: "394B5F",
    dataLabelFontFace: "+mn-lt", dataLabelPosition: "t", dataLabelFormatCode: "#,##0.0", showTitle: false,
  });
}
const chartTitle = (s, text, x, y, w) => T(s, text, { x, y, w, h: 0.4, fontSize: 14, bold: true, color: K.tx });
const para = (s, text, x, y, w, h) => T(s, text, { x, y, w, h, fontSize: 14, color: HEX.grey, valign: "top", align: "justify" });

// green notched-corner list box (divider pages)
function listBox(s, x, y, w, items) {
  const rh = 0.52, h = items.length * rh + 0.5;
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: 0.3, fill: { color: HEX.g }, line: { type: "none" } });
  items.forEach(([lb, id], i) => {
    const yy = y + 0.25 + i * rh;
    s.addShape(pres.ShapeType.ellipse, { x: x + w - 0.45, y: yy + 0.2, w: 0.11, h: 0.11, fill: { color: HEX.gold }, line: { type: "none" } });
    T(s, lb, { x: x + 0.3, y: yy, w: w - 0.95, h: rh, fontSize: 16, bold: true, color: K.wh });
    hit(s, x + 0.2, yy, w - 0.4, rh, id, lb);
  });
}

// divider: rounded photo left, title + list right
function divider(key, group, name, photo, items) {
  add("dv-" + key, () => {
    const s = frame({ section: name, group });
    s.addImage({ path: A("ph-" + photo + ".png"), x: 0.7, y: 1.5, w: 5.9, h: 6.6 });
    T(s, name, { x: 6.9, y: 2.6, w: 5.93, h: 1.2, fontSize: 48, bold: true, color: K.tx });
    listBox(s, 6.9, 4.0, 5.93, items);
  });
}

// ================================================================= cover
add("cover", () => {
  const s = pres.addSlide({ masterName: "BLANK" });
  pres.addSection({ title: "الغلاف" });
  s.addImage({ path: A("cover.png"), x: 0.4, y: 0.4, w: 12.53, h: 7.93 });
  T(s, "الإيرادات حسب", { x: 4.9, y: 5.35, w: 7.6, h: 0.8, fontSize: 38, bold: true, color: K.wh });
  T(s, "السنة والتبويب", { x: 4.9, y: 6.1, w: 7.6, h: 0.8, fontSize: 38, bold: true, color: K.wh });
  T(s, "في منطقة الرياض", { x: 4.9, y: 6.95, w: 7.6, h: 0.5, fontSize: 22, bold: true, color: K.wh });
  T(s, "التقرير الإحصائي للإيرادات للعام " + YEAR + "م", { x: 4.9, y: 7.5, w: 7.6, h: 0.45, fontSize: 16, color: "E3F0EA" });
  T(s, "أمانة منطقة الرياض | الإيرادات حسب السنة والتبويب", { x: 0.4, y: 8.7, w: 12.5, h: 0.4, fontSize: 13, bold: true, color: K.tx });
});

// ================================================================= intro, index, definitions
add("intro", () => {
  const s = frame({ section: "المقدمة", group: 0, title: "المقدمة" });
  s.addImage({ path: A("ph-tower.png"), x: 0.7, y: 1.3, w: 4.9, h: 5.5 });
  const ps = [
    "يقدّم هذا التقرير الإحصائي نظرة شاملة على إيرادات أمانة منطقة الرياض للعام " + YEAR + "م، اعتماداً على ملف الإيرادات الذي يجمع الإيراد الفعلي والمقدر والمستهدف والمفوتر والمحصل والمديونية.",
    "تُعرض البيانات منظمة حسب السنة والتبويب (إيرادات استثمارية، الرسوم والتراخيص، جزاءات وغرامات، التبغ والإيواء، أخرى) بدلاً من البنود التفصيلية، مع إمكانية التنقل بين التبويبات والسنوات من أعلى كل صفحة.",
    "يعرض التقرير المؤشرات الرئيسة، ثم الإيراد حتى تاريخه ولنهاية السنة وللشهر الحالي، والإيراد حسب القطاع، وتبويبات الوزارة، ثم التحصيل (المفوتر والجزاءات والتنفيذ والاسترداد وسيناريوهات الإقفال)، وأخيراً المديونية والاستقطاع.",
    "يهدف التقرير إلى دعم متخذي القرار بمؤشرات واضحة تساعد على متابعة الإيرادات وتحسين التحصيل.",
  ];
  ps.forEach((p, i) => para(s, p, 5.95, 2.3 + i * 1.55, 6.88, 1.5));
});

const SECTIONS = [["المقدمة", "intro", []], ["التعريفات", "defs", []], ["المؤشرات الرئيسة", "kpi", ["أبرز مؤشرات الإيراد", "الفعلي مقابل المقدر والمستهدف", "المفوتر والتحصيل"]],
  ["الإيراد", "dv-rev", ["الإيراد حتى تاريخه", "الإيراد لنهاية السنة", "الإيراد للشهر الحالي"]], ["القطاعات", "dv-sec", ["الإيراد حسب القطاع والتبويب"]], ["التبويبات", "dv-tab", ["ملخص التبويبات", "تبويبات الوزارة: المفوتر والمستهدف"]],
  ["التحصيل", "dv-col", ["المفوتر والميزانية", "الجزاءات والغرامات الشهرية", "سيناريوهات الإقفال", "التنفيذ والاسترداد"]], ["المديونية والاستقطاع", "dv-debt", ["المديونية حسب التبويب وسنة النشوء", "الاستقطاع"]], ["إخلاء المسؤولية وتواصل معنا", "disc", []]];
add("index", () => {
  const s = frame({ section: "المقدمة", group: 0, title: "الفهرس" });
  const cw = 3.95, cols = [SECTIONS.slice(0, 3), SECTIONS.slice(3, 6), SECTIONS.slice(6)];
  cols.forEach((list, ci) => {
    const x0 = 12.83 - cw * (ci + 1) - 0.25 * ci;
    let y = 2.45;
    list.forEach(([name, id, subs]) => {
      const no = SECTIONS.findIndex((q) => q[0] === name) + 1;
      T(s, String(no), { x: x0 + cw - 0.5, y, w: 0.5, h: 0.5, shape: pres.ShapeType.roundRect, rectRadius: 0.1, fill: { color: HEX.gold }, color: K.wh, bold: true, fontSize: 16, align: "center" });
      T(s, name, { x: x0, y, w: cw - 0.65, h: 0.5, fontSize: 18, bold: true, color: K.tx });
      hit(s, x0, y, cw, 0.5, id, name);
      y += 0.62;
      if (subs.length) {
        s.addShape(pres.ShapeType.line, { x: x0 + cw - 0.25, y: y - 0.05, w: 0, h: subs.length * 0.38 - 0.1, line: { color: HEX.blue, width: 1.5 } });
        subs.forEach((sb) => { T(s, sb, { x: x0, y, w: cw - 0.6, h: 0.34, fontSize: 12, color: HEX.grey }); y += 0.38; });
      }
      y += 0.3;
    });
  });
});

add("defs", () => {
  const s = frame({ section: "المقدمة", group: 0, title: "التعريفات" });
  const defs = [
    ["المفوتر", "قيمة الفواتير الصادرة خلال الفترة على المنشآت، ويُحتسب حسب تاريخ الإصدار."], ["الفعلي (المحقق)", "المبالغ المحصّلة فعلياً حسب تاريخ السداد، وهو الأساس المعتمد لدى الأمانة."],
    ["المقدّر", "الإيراد المتوقع للفترة المقدّر من الأمانة بالاتفاق مع وزارة المالية."], ["المستهدف", "المستهدف السنوي المعتمد للإيراد، ويُوزَّع شهرياً بحسب المؤشرات والسجلات المتاحة."],
    ["محصل سنوات سابقة", "المبالغ المحصّلة خلال السنة الحالية عن فواتير صادرة في سنوات سابقة."], ["التبويب", "تصنيف الإيراد الرئيس: إيرادات استثمارية، الرسوم والتراخيص، جزاءات وغرامات، التبغ والإيواء، أخرى."],
    ["الاستقطاع", "النسبة المقتطعة من إيراد التبويبات المحددة لصالح جهات أخرى وفق الأنظمة."], ["المديونية", "رصيد المبالغ المستحقة غير المسددة، موزعاً حسب التبويب وسنة نشوء المديونية."],
  ];
  defs.forEach(([h, t], i) => {
    const c = i % 2, r = Math.floor(i / 2), x = c === 0 ? 6.95 : 0.6, y = 2.35 + r * 1.55;
    T(s, h, { x, y, w: 5.9, h: 0.4, fontSize: 16, bold: true, color: HEX.g });
    T(s, t, { x, y: y + 0.42, w: 5.9, h: 0.9, fontSize: 14, color: HEX.grey, valign: "top" });
  });
});

// ================================================================= key indicators (green page)
add("kpi", () => {
  const td = rows("الايراد حتى تاريخه").map((r) => ({ adm: r[0], inv: r[2], act: r[3], prior: r[4], est: r[5], tgt: r[6], ly: r[7] })).filter((r) => r.adm === ADM[0]);
  const T0 = (k) => sum(td, (r) => r[k]);
  const ye = rows("الايراد لنهاية السنة").filter((r) => r[1] === ADM[0]), yEst = sum(ye, (r) => r[4]), yTgt = sum(ye, (r) => r[5]);
  const debt = rows("المديونية"), enf = rows("التنفيذ_والاسترداد"), cut = rows("استقطاع").filter((r) => r[6] === YEAR);
  const s = frame({ section: "المؤشرات الرئيسة", group: 1, dark: true, title: "أبرز مؤشرات الإيراد للعام " + YEAR + "م", sub: "إجمالي منطقة الرياض (مدينة الرياض + بلديات المنطقة) حتى تاريخه" });
  const items = [
    { label: "الإيراد المحقق (الفعلي) حتى تاريخه", value: fm(T0("act")), unit: "مليون ريال", delta: { v: delta(T0("act"), T0("ly")), label: "مقارنة بمحصل السنة الماضية" }, note: fm(T0("act") / yTgt * 1e6 / 1e6 * 1e6 / 1e6) + "% من المستهدف السنوي" },
    { label: "إجمالي المفوتر", value: fm(T0("inv")), unit: "مليون ريال", delta: { v: delta(T0("act"), T0("est")), label: "الفعلي عن المقدر" }, note: "المحصل سنوات سابقة " + fm(T0("prior")) + " مليون" },
    { label: "المستهدف حتى تاريخه", value: fm(T0("tgt")), unit: "مليون ريال", delta: { v: delta(T0("act"), T0("tgt")), label: "الفعلي عن المستهدف" }, note: "المقدر " + fm(T0("est")) + " مليون ريال" },
    { label: "المديونية المتبقية", value: fm(sum(debt, (r) => r[3])), unit: "مليون ريال", note: "سُدد خلال السنة " + fm(sum(debt, (r) => r[4])) + " مليون ريال" },
    { label: "استرداد فعلي (التنفيذ)", value: fm(sum(enf, (r) => r[5])), unit: "مليون ريال", note: "من إجمالي استرداد " + fm(sum(enf, (r) => r[4])) + " مليون" },
    { label: "الاستقطاع " + YEAR, value: fm(sum(cut, (r) => r[4])), unit: "مليون ريال", delta: { v: (sum(cut, (r) => r[4]) / Math.max(1, sum(cut, (r) => r[3]))) * 100, label: "من الإيراد" } },
  ];
  items[0].note = (T0("act") / yTgt * 100).toFixed(1) + "% من المستهدف السنوي";
  items.forEach((it, i) => hero(s, 12.83 - 4.0 * ((i % 3) + 1) - 0.165 * (i % 3), 2.7 + Math.floor(i / 3) * 2.95, 4.0, 2.75, Object.assign({ dark: true, big: 46, lfs: 16 }, it)));
});

// ================================================================= revenue sections (3 sheets x overview + one page per tab)
const MEAS = [{ k: "inv", n: "مفوتر", c: HEX.teal }, { k: "act", n: "فعلي", c: HEX.g }, { k: "est", n: "مقدر", c: HEX.gold }, { k: "tgt", n: "مستهدف", c: HEX.blue }, { k: "ly", n: "محصل السنة الماضية", c: "9AA5B1" }];
const REV = [
  { id: "td", key: "الايراد حتى تاريخه", name: "الإيراد حتى تاريخه", sub: "الفترة: يناير – أغسطس " + YEAR, inv: true, map: (r) => ({ adm: r[0], tab: r[1], inv: r[2], act: r[3], prior: r[4], est: r[5], tgt: r[6], ly: r[7] }) },
  { id: "ye", key: "الايراد لنهاية السنة", name: "الإيراد لنهاية السنة", sub: "تقدير نهاية السنة " + YEAR, inv: false, map: (r) => ({ adm: r[1], tab: r[2], act: r[3], est: r[4], tgt: r[5], ly: r[6] }) },
  { id: "mo", key: "الايراد الشهر الحالي", name: "الإيراد للشهر الحالي", sub: "شهر أغسطس " + YEAR, inv: false, map: (r) => ({ adm: r[1], tab: r[2], act: r[3], est: r[4], tgt: r[5], ly: r[6] }) },
];
divider("rev", 2, "الإيراد", "hotel", REV.map((o) => [o.name, o.id + "-all"]));
REV.forEach((o, oi) => {
  const R = rows(o.key).map(o.map), meas = MEAS.filter((m) => o.inv || m.k !== "inv");
  const tabNames = ["نظرة عامة"].concat(TABS), tabIds = [o.id + "-all"].concat(TABS.map((_, i) => o.id + "-" + i));
  const get = (adm, tab) => R.find((r) => r.adm === adm && r.tab === tab) || {};
  const tot = (adm, k) => sum(TABS, (t) => get(adm, t)[k]);
  const base = { section: "الإيراد", group: 2, switcher: { labels: REV.map((q) => q.name.replace("الإيراد ", "")), idx: oi, ids: REV.map((q) => q.id + "-all") }, tabs: tabNames, tabIds, years: [YEAR], yearIdx: 0 };
  add(o.id + "-all", () => {
    const s = frame(Object.assign({}, base, { tabIdx: 0, title: o.name, sub: o.sub + " — منطقة الرياض" }));
    const cards = [{ label: "إجمالي الإيراد", value: fm(tot(ADM[0], "act")), t: "g" }].concat(TABS.map((t) => ({ label: t, value: fm(get(ADM[0], t).act), t: null, tab: t })));
    cards.forEach((c, i) => {
      const x = 12.83 - 4.0 * ((i % 3) + 1) - 0.165 * (i % 3), y = s._y + 0.15 + Math.floor(i / 3) * 3.05;
      const a = c.tab ? n(get(ADM[0], c.tab).act) : tot(ADM[0], "act"), e = c.tab ? n(get(ADM[0], c.tab).est) : tot(ADM[0], "est"), g = c.tab ? n(get(ADM[0], c.tab).tgt) : tot(ADM[0], "tgt");
      hero(s, x, y, 4.0, 2.85, { light: true, tone: ["g", "gold", "blue", "teal", "g", "gold"][i], label: c.label, sub: "منطقة الرياض", value: fm(a), unit: "مليون ريال", delta: { v: delta(a, e), label: "عن المقدر" }, note: "من المستهدف: " + (g ? ((a / g) * 100).toFixed(1) : "—") + "%" });
      if (c.tab) hit(s, x, y, 4.0, 2.85, o.id + "-" + TABS.indexOf(c.tab), c.label, 0.2);
    });
  });
  TABS.forEach((tab, ti) => add(o.id + "-" + ti, () => {
    const r0 = get(ADM[0], tab), s = frame(Object.assign({}, base, { tabIdx: ti + 1, title: tab, sub: o.name + " — " + o.sub }));
    stack(s, 8.45, 4.38, [{ tone: "g", label: "الفعلي — منطقة الرياض", value: fm(r0.act), unit: "مليون ريال", delta: { v: delta(n(r0.act), n(r0.est)), label: "عن المقدر" }, note: "عن المستهدف " + dTxt(delta(n(r0.act), n(r0.tgt))) }]
      .concat(ADM.slice(1).map((a, ai) => ({ tone: ai === 0 ? "blue" : "gold", label: a, value: fm(get(a, tab).act), unit: "مليون ريال", delta: { v: delta(n(get(a, tab).act), n(get(a, tab).est)), label: "عن المقدر" } }))));
    chartTitle(s, "الفعلي مقابل المقدر والمستهدف — منطقة الرياض (مليون ريال)", 0.6, s._y + 0.1, 7.6);
    barList(s, 0.6, s._y + 0.65, 7.6, meas.map((m) => ({ label: m.n, value: n(r0[m.k]), text: fm(r0[m.k]), color: m.c })), { rh: 0.6, lw: 2.6 });
    para(s, tab + " — " + o.sub + ". بلغ فعلي منطقة الرياض " + fm(r0.act) + " مليون ريال، منها " + fm(get(ADM[1], tab).act) + " مليون لمدينة الرياض و" + fm(get(ADM[2], tab).act) + " مليون لبلديات المنطقة، " + pTxt(delta(n(r0.act), n(r0.est))) + " عن المقدر و" + pTxt(delta(n(r0.act), n(r0.tgt))) + " عن المستهدف.", 0.6, s._y + 0.8 + meas.length * 0.6, 7.6, 1.5);
  }));
});

// ================================================================= by sector
divider("sec", 3, "القطاعات", "plaza", [["الإيراد حسب القطاع والتبويب", "sec-all"]]);
{
  const R = rows("القطاع").map((r) => ({ adm: r[0], sec: r[1], tab: r[2], inv: r[3], act: r[4], prior: r[5], col: r[6] }));
  const T4 = TABS.slice(0, 4), tabs = ["نظرة عامة"].concat(T4), tids = ["sec-all"].concat([0, 1, 2, 3].map((i) => "sec-" + i));
  const secs = [...new Set(R.map((r) => r.sec))];
  const base = { section: "القطاعات", group: 3, tabs, tabIds: tids, years: [YEAR], yearIdx: 0 };
  add("sec-all", () => {
    const s = frame(Object.assign({}, base, { tabIdx: 0, title: "الإيراد حسب القطاع", sub: "إجمالي المحصل لكل قطاع موزعاً على التبويبات — مليون ريال" }));
    const body = secs.map((c) => { const f = T4.map((t) => sum(R.filter((r) => r.sec === c && r.tab === t), (r) => r.col)); return [c].concat(f.map((v) => V(v))).concat([V(f.reduce((a, b) => a + b, 0))]); });
    body.push(["الإجمالي"].concat(T4.map((t) => V(sum(R.filter((r) => r.tab === t), (r) => r.col)))).concat([V(sum(R, (r) => r.col))]));
    heat(s, ["القطاع"].concat(T4).concat(["المجموع"]), body, { x: 0.6, y: s._y + 0.2, w: 12.23, colW: [3.0, 2.0, 2.0, 2.0, 2.0, 1.23], rowH: 0.43, shade: [1, 4], byCol: true, totalRow: true, fs: 12 });
    T(s, "القيم بالمليون ريال، والأعمق = الأعلى ضمن التبويب نفسه.", { x: 0.6, y: s._y + 0.35 + 0.43 * (body.length + 1), w: 12.23, h: 0.3, fontSize: 10.5, color: "9AA5B1" });
  });
  T4.forEach((tab, ti) => add("sec-" + ti, () => {
    const F = R.filter((r) => r.tab === tab), s = frame(Object.assign({}, base, { tabIdx: ti + 1, title: tab, sub: "الفواتير والفعلي والمحصل حسب القطاع" }));
    stack(s, 8.85, 3.98, [{ tone: "g", label: "إجمالي المحصل", value: fm(sum(F, (r) => r.col)), unit: "مليون ريال" }, { tone: "blue", label: "إجمالي الفواتير", value: fm(sum(F, (r) => r.inv)), unit: "مليون ريال" }, { tone: "gold", label: "الفعلي", value: fm(sum(F, (r) => r.act)), unit: "مليون ريال" }]);
    chartTitle(s, "إجمالي المحصل حسب القطاع (مليون ريال)", 0.6, s._y + 0.1, 8.0);
    const its = secs.map((c) => ({ label: c, value: sum(F.filter((r) => r.sec === c), (r) => r.col), color: HEX.g })).sort((a, b) => b.value - a.value);
    barList(s, 0.6, s._y + 0.6, 8.0, its.map((i) => Object.assign(i, { text: fm(i.value) })), { rh: 0.5, lw: 3.2, fs: 12 });
  }));
}

// ================================================================= tab summaries from the account sheets + ministry tabs
divider("tab", 4, "التبويبات", "towers", [["ملخص التبويبات: الاستثمارية والجزاءات والتبغ", "sum-0"], ["تبويبات الوزارة: المفوتر والمستهدف", "min-all"]]);
{
  const inv = rows("الايرادات_الاستثمارية");
  const pen = rows("الجزاءات والتبغ والايواء").map((r) => ({ tab: r[0], inv: r[2], act: r[3], prior: r[4], est: r[5], tgt: r[6], ly: r[7] }));
  const L = (t) => pen.filter((r) => r.tab.replace(/\s/g, "").startsWith(t));
  const sums = [
    { name: "إيرادات استثمارية", inv: null, act: sum(inv, (r) => r[2]), prior: n(inv[0][9]), cur: n(inv[0][8]), est: sum(inv, (r) => r[3]), tgt: sum(inv, (r) => r[4]), ly: sum(inv, (r) => r[5]) },
    { name: "جزاءات وغرامات", rr: L("الجزاءات") }, { name: "التبغ والإيواء", rr: L("التبغ") },
  ].map((o) => (o.rr ? Object.assign(o, { inv: sum(o.rr, (r) => r.inv), act: sum(o.rr, (r) => r.act), prior: sum(o.rr, (r) => r.prior), est: sum(o.rr, (r) => r.est), tgt: sum(o.rr, (r) => r.tgt), ly: sum(o.rr, (r) => r.ly) }) : o));
  const names = ["ملخص التبويبات"].concat([]), swLabels = ["ملخص التبويبات", "تبويبات الوزارة"];
  const sw = (i) => ({ labels: swLabels, idx: i, ids: ["sum-0", "min-all"] });
  sums.forEach((x, i) => add("sum-" + i, () => {
    const first = x.inv === null ? x.cur : x.inv, firstName = x.inv === null ? "محصل السنة الحالية" : "مفوتر";
    const s = frame({ section: "التبويبات", group: 4, switcher: sw(0), tabs: sums.map((q) => q.name), tabIdx: i, tabIds: sums.map((_, k) => "sum-" + k), years: [YEAR], yearIdx: 0, title: x.name, sub: "ملخص التبويب حتى تاريخه — مقارنة بالمقدر والمستهدف" });
    stack(s, 8.45, 4.38, [{ tone: "g", label: "الفعلي", value: fm(x.act), unit: "مليون ريال", delta: { v: delta(x.act, x.est), label: "عن المقدر" }, note: "عن المستهدف " + dTxt(delta(x.act, x.tgt)) }, { tone: "blue", label: firstName, value: fm(first), unit: "مليون ريال" }, { tone: "gold", label: "الفعلي عن محصل السنة الماضية", value: dTxt(delta(x.act, x.ly)).replace(/[▲▼] /, "") }]);
    chartTitle(s, "مؤشرات التبويب (مليون ريال)", 0.6, s._y + 0.1, 7.6);
    barList(s, 0.6, s._y + 0.65, 7.6, [[firstName, first, HEX.teal], ["فعلي", x.act, HEX.g], ["محصل سنوات سابقة", x.prior, "9AA5B1"], ["مقدر", x.est, HEX.gold], ["مستهدف", x.tgt, HEX.blue], ["محصل السنة الماضية", x.ly, "9AA5B1"]].map(([l, v, c]) => ({ label: l, value: n(v), text: fm(v), color: c })), { rh: 0.62, lw: 2.6 });
    para(s, "بلغ فعلي " + x.name + " حتى تاريخه " + fm(x.act) + " مليون ريال، " + pTxt(delta(x.act, x.est)) + " عن المقدر و" + pTxt(delta(x.act, x.ly)) + " عن محصل السنة الماضية.", 0.6, s._y + 0.9 + 6 * 0.62, 7.6, 1.2);
  }));

  const MT = [["إيرادات استثمار العقارات البلدية", "استثمار العقارات"], ["إيرادات الرسوم والمقابلات المالية", "الرسوم والمقابلات"], ["إيرادات الجزاءات والغرامات", "الجزاءات والغرامات"],
    ["إيرادات رسم إشغال مرافق الإيواء", "إشغال الإيواء"], ["إيرادات رسم تقديم منتجات التبغ", "منتجات التبغ"], ["الإيرادات المختلفة", "إيرادات مختلفة"]];
  const A_ = rows("مفوتر تبويبات الوزارة").map((r) => ({ m: r[1], t: r[2], v: r[4] })), B_ = rows("مستهدف الوزارة").map((r) => ({ m: r[1], t: r[2], v: r[4] }));
  const mv = (Ls, t) => MONTHS.map((_, i) => { const f = Ls.filter((r) => r.m === i + 1 && (!t || r.t === t)); return f.length ? sum(f, (r) => r.v) : null; });
  const LM = Math.max(...A_.map((r) => r.m)), LN = MONTHS[LM - 1], upTo = (arr) => sum(arr.slice(0, LM), (v) => v);
  const mnames = ["نظرة عامة"].concat(MT.map((x) => x[1])), mids = ["min-all"].concat(MT.map((_, i) => "min-" + i));
  const mbase = { section: "التبويبات", group: 4, switcher: sw(1), tabs: mnames, tabIds: mids, years: [YEAR], yearIdx: 0 };
  add("min-all", () => {
    const s = frame(Object.assign({}, mbase, { tabIdx: 0, title: "المفوتر الشهري لكل تبويب", sub: "من يناير إلى " + LN + " " + YEAR + "م — مليون ريال" }));
    const body = MT.map((t) => { const a = mv(A_, t[0]); return [t[1]].concat(a.slice(0, LM).map((v) => V(v))).concat([V(upTo(a))]); });
    body.push(["الإجمالي"].concat(MONTHS.slice(0, LM).map((_, i) => V(sum(A_.filter((r) => r.m === i + 1), (r) => r.v)))).concat([V(upTo(mv(A_)))]));
    heat(s, ["التبويب"].concat(MONTHS.slice(0, LM)).concat(["المجموع"]), body, { x: 0.6, y: s._y + 0.2, w: 12.23, colW: [2.3].concat(MONTHS.slice(0, LM).map(() => 0.93)).concat([1.55]), rowH: 0.45, shade: [1, LM], totalRow: true, fs: 11 });
    const ta = upTo(mv(A_)), tb = upTo(mv(B_));
    para(s, "المفوتر حتى " + LN + " " + fm(ta) + " مليون ريال مقابل مستهدف الفترة " + fm(tb) + " مليون ريال (" + pTxt(delta(ta, tb)) + ")، ومستهدف السنة " + fm(sum(B_, (r) => r.v)) + " مليون ريال. الأعمق = الأعلى شهرياً داخل التبويب نفسه.", 0.6, s._y + 0.5 + 0.45 * (body.length + 1), 12.23, 1.0);
  });
  MT.forEach((t, i) => add("min-" + i, () => {
    const a = mv(A_, t[0]), b = mv(B_, t[0]);
    const s = frame(Object.assign({}, mbase, { tabIdx: i + 1, title: t[0], sub: "المفوتر والمستهدف الشهري" }));
    stack(s, 8.85, 3.98, [{ tone: "g", label: "المفوتر (يناير–" + LN + ")", value: fm(upTo(a)), unit: "مليون ريال", delta: { v: delta(upTo(a), upTo(b)), label: "عن المستهدف" } }, { tone: "blue", label: "المستهدف (يناير–" + LN + ")", value: fm(upTo(b)), unit: "مليون ريال" }, { tone: "gold", label: "المستهدف السنوي", value: fm(sum(b, (v) => v)), unit: "مليون ريال" }]);
    chartTitle(s, "المفوتر والمستهدف الشهري (مليون ريال)", 0.6, s._y + 0.1, 8.0);
    lineChart(s, MONTHS, [{ name: "مفوتر " + YEAR, values: a }, { name: "مستهدف", values: b, dash: true }], { x: 0.6, y: s._y + 0.5, w: 8.0, h: 4.1, colors: [HEX.g, "9AA5B1"] });
    para(s, t[0] + ": المفوتر " + fm(upTo(a)) + " مليون ريال مقابل مستهدف الفترة " + fm(upTo(b)) + " مليون ريال (" + pTxt(delta(upTo(a), upTo(b))) + ").", 0.6, s._y + 4.75, 8.0, 1.1);
  }));
}

// ================================================================= collection: invoiced/budget, penalties, scenarios, enforcement
divider("col", 5, "التحصيل", "building", [["المفوتر والميزانية", "inv-0"], ["الجزاءات والغرامات الشهرية", "pen-all"], ["سيناريوهات الإقفال", "scn-0"], ["التنفيذ والاسترداد", "enf-0"]]);
const colSwitch = (i) => ({ labels: ["المفوتر", "الجزاءات", "الإقفال", "الاسترداد"], idx: i, ids: ["inv-0", "pen-all", "scn-0", "enf-0"] });
{
  const mn = (a, k) => MONTHS.map((_, i) => { const r = a.find((x) => x[1] === i + 1); return r ? r[k] : null; });
  const inv = rows("المفوتر"), orig = rows("الميزانية الاعتماد الأصلي").filter((r) => r[1]), post = rows("الميزانية بعد  الاعتماد").filter((r) => r[1]), tg = rows("مستهدف المفوتر");
  const names = ["المفوتر والمحصل", "الميزانية والمستهدف"], tids = ["inv-0", "inv-1"], LN = MONTHS[Math.max(...inv.map((r) => r[1])) - 1];
  const base = { section: "التحصيل", group: 5, switcher: colSwitch(0), tabs: names, tabIds: tids, years: [YEAR], yearIdx: 0 };
  add("inv-0", () => {
    const s = frame(Object.assign({}, base, { tabIdx: 0, title: "الفوترة والتحصيل", sub: "الفترة: يناير – " + LN + " " + YEAR }));
    stack(s, 8.45, 4.38, [{ tone: "blue", label: "إجمالي المفوتر", value: fm(sum(inv, (r) => r[2])), unit: "مليون ريال" }, { tone: "g", label: "محصل السنة الحالية", value: fm(sum(inv, (r) => r[4])), unit: "مليون ريال" }, { tone: "gold", label: "محصل السنوات السابقة", value: fm(sum(inv, (r) => r[5])), unit: "مليون ريال" }]);
    chartTitle(s, "المفوتر وصافي المفوتر شهرياً (مليون ريال)", 0.6, s._y + 0.1, 7.6);
    lineChart(s, MONTHS, [{ name: "مفوتر", values: mn(inv, 2) }, { name: "صافي المفوتر", values: mn(inv, 3), dash: true }, { name: "مستهدف المفوتر", values: mn(tg, 2) }], { x: 0.6, y: s._y + 0.5, w: 7.6, h: 2.5, labels: false, colors: [HEX.g, "9AA5B1", HEX.gold] });
    chartTitle(s, "المحصل شهرياً (مليون ريال)", 0.6, s._y + 3.1, 7.6);
    lineChart(s, MONTHS, [{ name: "محصل السنة الحالية", values: mn(inv, 4) }, { name: "محصل السنوات السابقة", values: mn(inv, 5), dash: true }], { x: 0.6, y: s._y + 3.5, w: 7.6, h: 2.5, labels: false, colors: [HEX.blue, "9AA5B1"] });
  });
  add("inv-1", () => {
    const s = frame(Object.assign({}, base, { tabIdx: 1, title: "الميزانية ومستهدف المفوتر", sub: "الاعتماد الأصلي وبعد الاعتماد " + YEAR }));
    stack(s, 8.85, 3.98, [{ tone: "g", label: "الاعتماد الأصلي (سنوي)", value: fm(sum(orig, (r) => r[2])), unit: "مليون ريال" }, { tone: "blue", label: "بعد الاعتماد (سنوي)", value: fm(sum(post, (r) => r[2])), unit: "مليون ريال" }, { tone: "gold", label: "مستهدف المفوتر (سنوي)", value: fm(sum(tg, (r) => r[2])), unit: "مليون ريال" }]);
    chartTitle(s, "الميزانية الشهرية ومستهدف المفوتر (مليون ريال)", 0.6, s._y + 0.1, 8.0);
    lineChart(s, MONTHS, [{ name: "مستهدف المفوتر", values: mn(tg, 2) }, { name: "الاعتماد الأصلي", values: mn(orig, 2), dash: true }, { name: "بعد الاعتماد", values: mn(post, 2), dash: true }], { x: 0.6, y: s._y + 0.5, w: 8.0, h: 4.4, colors: [HEX.g, "9AA5B1", HEX.gold], labels: false });
    para(s, "الميزانية بعد الاعتماد " + fm(sum(post, (r) => r[2])) + " مليون ريال مقابل مستهدف مفوتر سنوي " + fm(sum(tg, (r) => r[2])) + " مليون ريال.", 0.6, s._y + 5.05, 8.0, 0.9);
  });

  // penalties by year (years are the tabs here)
  const P = rows("الجزائات والغرامات").map((r) => ({ y: r[0], m: r[1], inv: r[2], past: r[3], same: r[4] }));
  const years = [...new Set(P.map((r) => r.y))].sort(), yids = years.map((y) => "pen-" + y);
  const pbase = { section: "التحصيل", group: 5, switcher: colSwitch(1) };
  add("pen-all", () => {
    const s = frame(Object.assign({}, pbase, { title: "الجزاءات والغرامات — مقارنة السنوات", sub: "المفوتر الشهري لكل سنة", years: ["مقارنة"].concat(years), yearIdx: 0, yearIds: [null].concat(yids) }));
    const ser = years.slice().reverse().map((y, i) => ({ name: String(y), values: P.filter((r) => r.y === y).sort((a, b) => a.m - b.m).map((r) => r.inv), dash: y !== years[years.length - 1] }));
    chartTitle(s, "مفوتر الجزاءات والغرامات شهرياً حسب السنة (مليون ريال)", 0.6, s._y + 0.1, 8.0);
    lineChart(s, MONTHS, ser.reverse(), { x: 0.6, y: s._y + 0.5, w: 8.0, h: 4.6, colors: [HEX.g, HEX.blue, HEX.gold, "9AA5B1"], labels: false });
    stack(s, 8.85, 3.98, years.slice().reverse().map((y, i) => ({ tone: ["g", "blue", "gold", "teal"][i], label: "مفوتر سنة " + y, value: fm(sum(P.filter((r) => r.y === y), (r) => r.inv)), unit: "مليون ريال" })));
  });
  years.forEach((y, yi) => add("pen-" + y, () => {
    const f = P.filter((r) => r.y === y).sort((a, b) => a.m - b.m), prev = P.filter((r) => r.y === y - 1).sort((a, b) => a.m - b.m);
    const s = frame(Object.assign({}, pbase, { title: "الجزاءات والغرامات — " + y, sub: "المفوتر والمحصل شهرياً", years: ["مقارنة"].concat(years), yearIdx: yi + 1, yearIds: ["pen-all"].concat(yids) }));
    stack(s, 8.45, 4.38, [{ tone: "g", label: "إجمالي مفوتر السنة", value: fm(sum(f, (r) => r.inv)), unit: "مليون ريال", delta: prev.length ? { v: delta(sum(f, (r) => r.inv), sum(prev, (r) => r.inv)), label: "عن " + (y - 1) } : null }, { tone: "blue", label: "محصل عن الفترة الماضية", value: fm(sum(f, (r) => r.past)), unit: "مليون ريال" }, { tone: "gold", label: "محصل عن ذات الشهر", value: fm(sum(f, (r) => r.same)), unit: "مليون ريال" }]);
    chartTitle(s, "الجزاءات والغرامات الشهرية (مليون ريال)", 0.6, s._y + 0.1, 7.6);
    const ser = [{ name: "مفوتر الشهر " + y, values: f.map((r) => r.inv) }].concat(prev.length ? [{ name: "مفوتر " + (y - 1), values: prev.map((r) => r.inv), dash: true }] : []);
    lineChart(s, MONTHS, ser, { x: 0.6, y: s._y + 0.5, w: 7.6, h: 2.6, labels: false, colors: [HEX.g, "9AA5B1"] });
    chartTitle(s, "المحصل الشهري (مليون ريال)", 0.6, s._y + 3.2, 7.6);
    lineChart(s, MONTHS, [{ name: "محصل عن الفترة الماضية", values: f.map((r) => r.past) }, { name: "محصل عن ذات الشهر", values: f.map((r) => r.same), dash: true }], { x: 0.6, y: s._y + 3.6, w: 7.6, h: 2.4, labels: false, colors: [HEX.blue, "9AA5B1"] });
  }));

  // closing scenarios
  const mon = ["سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"], snames = ["السيناريو الأول", "السيناريو الثاني"], stids = ["scn-0", "scn-1"];
  ["السناريو_الاول_للاقفال", "السناريو_الثاني_للاقفال"].forEach((k, i) => add("scn-" + i, () => {
    const R = rows(k).map((r) => ({ y: r[0], m: r[1], tab: r[2], col: r[3], tgt: r[4] })), tc = sum(R, (r) => r.col), tt = sum(R, (r) => r.tgt);
    const s = frame({ section: "التحصيل", group: 5, switcher: colSwitch(2), tabs: snames, tabIdx: i, tabIds: stids, years: [R[0].y], yearIdx: 0, title: snames[i] + " للإقفال", sub: "المحصل المتوقع مقابل المستهدف للأشهر المتبقية من " + R[0].y });
    stack(s, 9.05, 3.78, [{ tone: "g", label: "المحصل المتوقع", value: fm(tc), unit: "مليون ريال", delta: { v: delta(tc, tt), label: "عن المستهدف" } }, { tone: "blue", label: "المستهدف", value: fm(tt), unit: "مليون ريال" }]);
    chartTitle(s, "المحصل والمستهدف الشهري (مليون ريال)", 0.6, s._y + 0.1, 8.2);
    lineChart(s, mon, [{ name: "محصل", values: mon.map((m) => sum(R.filter((r) => r.m === m), (r) => r.col)) }, { name: "مستهدف", values: mon.map((m) => sum(R.filter((r) => r.m === m), (r) => r.tgt)), dash: true }], { x: 0.6, y: s._y + 0.5, w: 8.2, h: 2.3, colors: [HEX.g, "9AA5B1"] });
    const body = TABS.concat(["نقل الملكية"]).map((t) => { const f = R.filter((r) => r.tab === t), a = sum(f, (r) => r.col), b = sum(f, (r) => r.tgt); return [t, V(a), V(b), V(delta(a, b), dTxt(delta(a, b)))]; });
    heat(s, ["التبويب", "محصل", "مستهدف", "المحصل عن المستهدف"], body, { x: 0.6, y: s._y + 3.0, w: 8.2, colW: [2.4, 1.8, 1.8, 2.2], rowH: 0.38, shade: [1, 2], byCol: true, fs: 11.5 });
  }));

  // enforcement & recovery
  const E = rows("التنفيذ_والاسترداد").map((r) => ({ adm: r[0], tab: (r[1] || "").replace("الإيرادات الاستثمارية", "إيرادات إستثمارية").replace("جزاءات والغرامات", "جزاءات وغرامات"), inv: r[2], ex: r[3], rec: r[4], real: r[5] }));
  const enames = ["الإجمالي", "مدينة الرياض", "بلديات المنطقة"], etids = ["enf-0", "enf-1", "enf-2"];
  enames.forEach((nm, i) => add("enf-" + i, () => {
    const F = E.filter((r) => i === 0 || r.adm === nm), by = (t, k) => sum(F.filter((r) => r.tab === t), (r) => r[k]);
    const s = frame({ section: "التحصيل", group: 5, switcher: colSwitch(3), tabs: enames, tabIdx: i, tabIds: etids, years: [YEAR], yearIdx: 0, title: "التنفيذ والاسترداد — " + nm, sub: "حسب التبويب — مليون ريال" });
    const c4 = [["قيمة الفواتير", "inv", "g"], ["المنفذ", "ex", "blue"], ["إجمالي الاسترداد", "rec", "gold"], ["استرداد فعلي", "real", "teal"]];
    c4.forEach(([l, k, t], j) => hero(s, 12.83 - 3.0 * (j + 1) - 0.11 * j, s._y + 0.1, 3.0, 1.75, { tone: t, label: l, value: fm(sum(F, (r) => r[k])), unit: "مليون ريال", big: 32, lfs: 15 }));
    const body = TABS.map((t) => [t, V(by(t, "inv")), V(by(t, "ex")), V(by(t, "rec")), V(by(t, "real"))]);
    body.push(["الإجمالي", V(sum(F, (r) => r.inv)), V(sum(F, (r) => r.ex)), V(sum(F, (r) => r.rec)), V(sum(F, (r) => r.real))]);
    heat(s, ["التبويب", "قيمة الفواتير", "المنفذ", "إجمالي الاسترداد", "استرداد فعلي"], body, { x: 0.6, y: s._y + 2.1, w: 12.23, colW: [3.4, 2.2, 2.2, 2.2, 2.23], rowH: 0.5, shade: [1, 4], byCol: true, totalRow: true, fs: 13 });
  }));
}

// ================================================================= debt & deductions
divider("debt", 6, "المديونية والاستقطاع", "night", [["المديونية حسب التبويب وسنة النشوء", "debt-all"], ["الاستقطاع", "cut-2026"]]);
{
  const R = rows("المديونية").map((r) => ({ y: r[0], tab: String(r[1]).trim(), tot: r[2], rem: r[3], paid: r[4], canc: r[5] }));
  const dn = ["نظرة عامة", "إيرادات استثمارية", "جزاءات وغرامات", "التبغ والإيواء", "أخرى"], dids = ["debt-all"].concat([0, 1, 2, 3].map((i) => "debt-" + i));
  const dsw = (i) => ({ labels: ["المديونية", "الاستقطاع"], idx: i, ids: ["debt-all", "cut-2026"] });
  const dbase = { section: "المديونية والاستقطاع", group: 6, switcher: dsw(0), tabs: dn, tabIds: dids, years: [YEAR], yearIdx: 0 };
  const side = (s, F) => {
    stack(s, 9.05, 3.78, [{ tone: "g", label: "رصيد المديونية الإجمالي", value: fm(sum(F, (r) => r.tot)), unit: "مليون ريال", note: "المتبقي " + fm(sum(F, (r) => r.rem)) + " مليون ريال" }, { tone: "blue", label: "ما تم سداده السنة الحالية", value: fm(sum(F, (r) => r.paid)), unit: "مليون ريال" }, { tone: "gold", label: "الملغي من المديونية", value: fm(sum(F, (r) => r.canc)), unit: "مليون ريال" }]);
  };
  add("debt-all", () => {
    const s = frame(Object.assign({}, dbase, { tabIdx: 0, title: "المديونية — نظرة عامة", sub: "رصيد المديونية حسب التبويب — مليون ريال" }));
    side(s, R);
    const body = dn.slice(1).map((t) => { const f = R.filter((r) => r.tab === t); return [t, V(sum(f, (r) => r.tot)), V(sum(f, (r) => r.rem)), V(sum(f, (r) => r.paid)), V(sum(f, (r) => r.canc))]; });
    body.push(["الإجمالي", V(sum(R, (r) => r.tot)), V(sum(R, (r) => r.rem)), V(sum(R, (r) => r.paid)), V(sum(R, (r) => r.canc))]);
    heat(s, ["التبويب", "الرصيد الإجمالي", "المتبقي", "ما تم سداده", "الملغي"], body, { x: 0.6, y: s._y + 0.2, w: 8.2, colW: [2.4, 1.5, 1.4, 1.5, 1.4], rowH: 0.6, shade: [1, 4], byCol: true, totalRow: true, fs: 12 });
    para(s, "المتبقي من المديونية " + fm(sum(R, (r) => r.rem)) + " مليون ريال من رصيد إجمالي " + fm(sum(R, (r) => r.tot)) + " مليون ريال؛ سُدّد منها خلال السنة الحالية " + fm(sum(R, (r) => r.paid)) + " مليون ريال.", 0.6, s._y + 0.45 + 0.6 * 6, 8.2, 1.2);
  });
  dn.slice(1).forEach((t, i) => add("debt-" + i, () => {
    const F = R.filter((r) => r.tab === t).sort((a, b) => b.y - a.y), s = frame(Object.assign({}, dbase, { tabIdx: i + 1, title: t, sub: "المديونية حسب سنة النشوء — ألف ريال" }));
    side(s, F);
    const rh = Math.min(0.34, 5.3 / (F.length + 1));
    heat(s, ["سنة النشوء", "الرصيد الإجمالي", "المتبقي", "ما تم سداده", "الملغي"], F.map((r) => [String(r.y), V(r.tot, fk(r.tot)), V(r.rem, fk(r.rem)), V(r.paid, fk(r.paid)), V(r.canc, fk(r.canc))]), { x: 0.6, y: s._y + 0.2, w: 8.2, colW: [1.6, 1.9, 1.7, 1.6, 1.4], rowH: rh, shade: [1, 4], byCol: true, fs: 10.5 });
  }));

  const Rc = rows("استقطاع").map((r) => ({ adm: r[0], tab: r[1], rev: r[3], cut: r[4], year: r[6] }));
  const AM = rows("الاستقطاع_الامانات_الاخرى"), years = [...new Set(Rc.map((r) => r.year))].sort(), yids = years.map((y) => "cut-" + y);
  years.forEach((y, yi) => add("cut-" + y, () => {
    const am = AM.find((r) => r[0] === y) || [];
    const s = frame({ section: "المديونية والاستقطاع", group: 6, switcher: dsw(1), years, yearIdx: yi, yearIds: yids, title: "الاستقطاع من الإيراد — " + y, sub: "حسب التنظيم الإداري والتبويب" });
    stack(s, 8.85, 3.98, [{ tone: "g", label: "استقطاع الأمانة", value: fm(am[4]), unit: "مليون ريال" }, { tone: "blue", label: "الإيرادات الفعلية للأمانة", value: fm(am[2]), unit: "مليون ريال" }, { tone: "gold", label: "الاستقطاع لكافة الأمانات", value: fm(am[3]), unit: "مليون ريال" }, { tone: "teal", label: "إيرادات كافة الأمانات", value: fm(am[1]), unit: "مليون ريال" }]);
    const body = [];
    ["مدينة الرياض", "بلديات المنطقة"].forEach((a) => ["الرسوم والتراخيص", "التبغ والإيواء"].forEach((t) => { const f = Rc.filter((r) => r.year === y && r.adm === a && r.tab === t), rv = sum(f, (r) => r.rev), ct = sum(f, (r) => r.cut); body.push([a + " — " + t, V(rv), V(ct), V(rv ? (ct / rv) * 100 : 0, (rv ? ((ct / rv) * 100).toFixed(1) : "0.0") + "%")]); }));
    chartTitle(s, "الإيراد والمستقطع حسب التنظيم الإداري والتبويب (مليون ريال)", 0.6, s._y + 0.1, 8.0);
    heat(s, ["التنظيم الإداري — التبويب", "الإيراد", "المستقطع", "نسبة الاستقطاع"], body, { x: 0.6, y: s._y + 0.6, w: 8.0, colW: [3.5, 1.5, 1.5, 1.5], rowH: 0.55, shade: [1, 2], byCol: true, fs: 12 });
    barList(s, 0.6, s._y + 0.8 + 0.55 * 5, 8.0, body.map((r) => ({ label: r[0], value: r[2].v, text: r[2].t, color: HEX.blue })), { rh: 0.45, lw: 3.5, fs: 11 });
  }));
}

// ================================================================= disclaimer, notes, end
add("disc", () => {
  const s = frame({ section: "إخلاء المسؤولية", group: 7, title: "إخلاء المسؤولية" });
  s.addImage({ path: A("ph-night.png"), x: 0.7, y: 1.3, w: 4.6, h: 5.2 });
  const ps = ["أُعدّ هذا التقرير من بيانات ملف الإيرادات لأمانة منطقة الرياض لأغراض المتابعة الداخلية ودعم اتخاذ القرار، وهو يعكس البيانات المتاحة وقت إعداده وقد تتغير بعد ورود تحصيلات أو فواتير جديدة.",
    "ولا تتحمل الأمانة مسؤولية أي قرار يُتخذ بالاعتماد على هذا التقرير وحده دون مراجعة المصادر الأصلية.", "لا يجوز نسخ هذا التقرير أو تداوله خارج الجهات المصرّح لها دون إذن مسبق من الأمانة."];
  ps.forEach((p, i) => para(s, p, 5.7, 2.3 + i * 1.0, 7.13, 0.95));
  s.addShape(pres.ShapeType.roundRect, { x: 5.7, y: 5.4, w: 7.13, h: 2.3, rectRadius: 0.25, fill: { color: "FFFFFF" }, line: { color: "E3E9E6", width: 0.75 }, shadow: shadow() });
  T(s, "تواصل معنا", { x: 6.0, y: 5.5, w: 6.5, h: 0.5, fontSize: 24, bold: true, color: K.tx });
  T(s, "أمانة منطقة الرياض — إدارة متابعة الإيرادات. للاستفسار عن بيانات هذا التقرير أو طلب نسخة محدّثة، يرجى التواصل مع الفريق المختص.", { x: 6.0, y: 6.1, w: 6.5, h: 1.4, fontSize: 13, color: HEX.grey, valign: "top" });
});
add("notes", () => {
  const s = frame({ section: "إخلاء المسؤولية", group: 7, title: "مصادر البيانات والملاحظات" });
  const notes = [
    "جميع البيانات مسجّلة حسب السنة والتبويب؛ البنود التفصيلية (الحسابات) مجمّعة على مستوى التبويب.",
    "الشيتات التي لا تحتوي عمود سنة نُسبت إلى " + YEAR + " (الإيراد للشهر الحالي، القطاع، التنفيذ والاسترداد، الاستثمارية).",
    "نسب الفعلي عن المقدر والمستهدف والسنة الماضية محسوبة من الأرقام الخام (الفعلي ÷ القاعدة − 1).",
    "إجمالي منطقة الرياض = مدينة الرياض + بلديات المنطقة، ولا تُجمع الثلاثة معاً تفادياً للتكرار.",
    "بيانات سيناريو الإقفال الأول والثاني متطابقة في الملف المصدر.",
    "المبالغ بالمليون ريال سعودي ما لم يُذكر خلاف ذلك، والأرقام مقرّبة وقد لا تتطابق المجاميع نتيجة لعمليات التقريب.",
    "تواريخ الفترات: الإيراد حتى تاريخه لشهر أغسطس، والمفوتر وتبويبات الوزارة حتى أحدث شهر في الملف.",
  ];
  notes.forEach((t, i) => { s.addShape(pres.ShapeType.ellipse, { x: 12.65, y: 2.5 + i * 0.85 + 0.17, w: 0.12, h: 0.12, fill: { color: HEX.slate }, line: { type: "none" } }); T(s, t, { x: 0.6, y: 2.5 + i * 0.85, w: 11.85, h: 0.75, fontSize: 14, color: HEX.grey, valign: "top" }); });
});
add("end", () => {
  const s = pres.addSlide({ masterName: "BLANK" });
  pres.addSection({ title: "الختام" });
  s.addShape(pres.ShapeType.rect, { x: 0, y: 0, w: W, h: H, fill: { color: HEX.g }, line: { type: "none" } });
  s.addShape(pres.ShapeType.roundRect, { x: 5.6, y: 3.0, w: 2.13, h: 2.3, rectRadius: 0.3, fill: { color: "FFFFFF" }, line: { type: "none" } });
  s.addImage({ path: A("logo-amana.png"), x: 5.77, y: 3.15, w: 1.8, h: 1.95 });
  T(s, "أمانة منطقة الرياض | الإيرادات حسب السنة والتبويب", { x: 2.0, y: 5.7, w: 9.33, h: 0.6, fontSize: 24, bold: true, color: K.wh, align: "center" });
});

// ================================================================= build
plan.forEach((p, i) => { cur = i + 1; p.fn(); });
pres.writeFile({ fileName: OUT }).then(() => applyTheme(OUT, THEME)).then(() => console.log("wrote", OUT, plan.length, "slides"));
