// Builds the executive deck for the Revenue & Finance Directorate strategy 2026–2030,
// on the Amana visual identity used by build_deck.js (same canvas, backgrounds, logo, palette).
// Run: node scripts/build_strategy_deck.js [out.pptx]
const path = require("path");
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.env.APPLY_THEME || "/root/.claude/skills/synced/642995fb-6fb0-4612-9fc1-a71063ac95e2_2132651f-4df5-4687-9c61-978ed9bca072/pptx/scripts/apply_theme.js");

const ROOT = path.join(__dirname, "..");
const OUT = process.argv[2] || path.join(ROOT, "strategy", "revenue-finance-strategy-2026-2030.pptx");
const A = (f) => path.join(ROOT, "assets", f);

const THEME = {
  name: "Amana Strategy", headFontFace: "Arial", bodyFontFace: "Arial",
  colors: { dk1: "1F2A37", lt1: "FFFFFF", dk2: "394B5F", lt2: "F1F5F3", accent1: "006A46", accent2: "74A4D4", accent3: "E4A42C", accent4: "0B6B73", accent5: "E5736B", accent6: "8CBCAC", hlink: "006A46", folHlink: "5B6B7B" },
};
const HEX = { g: "006A46", teal: "0B6B73", gold: "E4A42C", pink: "E5736B", blue: "74A4D4", slate: "394B5F", grey: "6B7785", lt: "F1F5F3", mint: "E3F0EA", line: "D5DFDA", dk: "1F2A37" };

const pres = new pptxgen();
const W = 13.33, H = 9.43, L = 0.5, R = 12.83, CW = R - L;
pres.defineLayout({ name: "A4L", width: W, height: H });
pres.layout = "A4L";
pres.rtlMode = true;
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "استراتيجية إدارة الإيرادات والتمويل 2026–2030";
pres.author = "أمانة منطقة الرياض";
const C = pres.SchemeColor;

const NOTE = "استراتيجية إدارة الإيرادات والتمويل 2026–2030 | أمانة منطقة الرياض — مسودة للمراجعة";
const footer = (c) => [{ text: { text: NOTE, options: { x: 4.2, y: 8.85, w: 8.63, h: 0.3, fontSize: 10, color: c, align: "right", valign: "middle", rtlMode: true, margin: 0 } } }];
pres.defineSlideMaster({ title: "LIGHT", background: { path: A("bg-dots.png") },
  objects: [{ image: { x: 12.0, y: 0.22, w: 0.82, h: 0.89, path: A("logo-amana.png") } }].concat(footer("8A95A3")),
  slideNumber: { x: 0.5, y: 8.88, w: 0.5, h: 0.25, fontSize: 10, color: "5B6B7B", align: "left", bold: true } });
pres.defineSlideMaster({ title: "GREEN", background: { path: A("bg-green.png") },
  objects: [{ rect: { x: 12.0, y: 0.22, w: 0.82, h: 0.89, fill: { color: "FFFFFF" }, rectRadius: 0.14 } }, { image: { x: 12.0, y: 0.22, w: 0.82, h: 0.89, path: A("logo-amana.png") } }].concat(footer("CFE3DA")),
  slideNumber: { x: 0.5, y: 8.88, w: 0.5, h: 0.25, fontSize: 10, color: "CFE3DA", align: "left", bold: true } });
pres.defineSlideMaster({ title: "COVER", background: { color: "1F2A37" }, objects: [] });

// ---------- building blocks ----------
let secSeen = new Set();
function slide(section, master = "LIGHT") {
  if (!secSeen.has(section)) { secSeen.add(section); pres.addSection({ title: section }); }
  return pres.addSlide({ masterName: master, sectionTitle: section });
}
const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, rtlMode: true, align: "right", margin: 0, valign: "middle", fontSize: 15, color: HEX.dk, fontFace: "Arial" }, o));
const shadow = () => ({ type: "outer", color: "000000", opacity: 0.12, blur: 8, offset: 2, angle: 90 });
function card(s, x, y, w, h, o = {}) {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, rectRadius: o.r ?? 0.16, fill: { color: o.fill || "FFFFFF" }, line: o.line ? { color: o.line, width: 1 } : { type: "none" }, shadow: o.flat ? undefined : shadow() });
}
function title(s, t, sub, dark) {
  T(s, t, { x: L, y: 0.32, w: 11.2, h: 0.8, fontSize: 32, bold: true, color: dark ? "FFFFFF" : HEX.dk });
  if (sub) T(s, sub, { x: L, y: 1.1, w: 11.2, h: 0.42, fontSize: 15, color: dark ? "CFE3DA" : HEX.grey });
}
function badge(s, x, y, d, txt, fill = HEX.g, color = "FFFFFF", fs = 14) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" } });
  T(s, txt, { x, y, w: d, h: d, align: "center", fontSize: fs, bold: true, color });
}
// RTL grid: item i placed from the right edge
const colX = (i, n, gap, x0 = L, w0 = CW) => { const w = (w0 - gap * (n - 1)) / n; return { x: x0 + w0 - (i + 1) * w - i * gap, w }; };
const bullets = (items, o = {}) => items.map((t, i) => ({ text: t, options: { bullet: { indent: 18 }, breakLine: i < items.length - 1, paraSpaceAfter: o.space ?? 6 } }));

// ---------- 1. cover ----------
{
  const s = slide("الغلاف", "COVER");
  s.addImage({ path: A("cover.png"), x: 0, y: 0, w: W, h: H, sizing: { type: "cover", w: W, h: H } });
  T(s, "استراتيجية إدارة الإيرادات والتمويل", { x: 4.3, y: 5.85, w: 8.3, h: 0.95, fontSize: 38, bold: true, color: "FFFFFF" });
  T(s, "2026 – 2030", { x: 4.3, y: 6.8, w: 8.3, h: 0.7, fontSize: 30, bold: true, color: HEX.gold });
  T(s, "وكالة الخدمات المؤسسية | أمانة منطقة الرياض", { x: 4.3, y: 7.6, w: 8.3, h: 0.45, fontSize: 17, color: HEX.mint });
  T(s, "مسودة للمراجعة والاعتماد", { x: 4.3, y: 8.15, w: 8.3, h: 0.4, fontSize: 13, color: "CFE3DA" });
  s.addNotes("استراتيجية إدارة الإيرادات والتمويل وإداراتها الثلاث للفترة 2026–2030.");
}

// ---------- 2. agenda ----------
{
  const s = slide("المقدمة");
  title(s, "محتويات العرض");
  const items = ["السياق والمستهدف", "الرؤية والأهداف", "الهيكل التنظيمي", "تخطيط ومراقبة الإيرادات", "الائتمان والتمويل", "الاستدامة المالية", "خارطة الطريق والمؤشرات", "القرارات المطلوبة"];
  items.forEach((t, i) => {
    const r = Math.floor(i / 4), { x, w } = colX(i % 4, 4, 0.35);
    const y = 2.0 + r * 3.1;
    card(s, x, y, w, 2.6);
    T(s, String(i + 1).padStart(2, "0"), { x: x + 0.3, y: y + 0.3, w: w - 0.6, h: 0.9, fontSize: 40, bold: true, color: i % 2 ? HEX.teal : HEX.g });
    T(s, t, { x: x + 0.3, y: y + 1.35, w: w - 0.6, h: 0.95, fontSize: 19, bold: true, valign: "top" });
  });
}

// ---------- 3. context ----------
{
  const s = slide("السياق والمستهدف");
  title(s, "إيرادات نمت 30% خلال سنتين", "تطور إيرادات الأمانة (مليون ريال)");
  s.addChart(pres.ChartType.bar, [{ name: "الإيرادات", labels: ["قبل سنتين", "السنة السابقة", "السنة الأخيرة"], values: [2800, 3100, 3650] }], {
    x: L, y: 1.8, w: 6.4, h: 6.6, barDir: "col", chartColors: [HEX.g], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0",
    dataLabelFontSize: 16, dataLabelFontBold: true, dataLabelColor: HEX.dk, dataLabelFontFace: "+mn-lt", catAxisLabelFontSize: 14, catAxisLabelColor: HEX.slate, catAxisLabelFontFace: "+mn-lt",
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false, barGapWidthPct: 70, valAxisMinVal: 0, valAxisMaxVal: 4200,
  });
  const facts = [
    ["10.7%", "نمو السنة السابقة", HEX.teal],
    ["17.7%", "نمو السنة الأخيرة", HEX.g],
    ["الفائض لنا", "الإيرادات تُورَّد للخزينة، وما يتجاوز المستهدف يعود لميزانية الأمانة", HEX.gold],
    ["2030 · 2034", "إكسبو الرياض وكأس العالم يرفعان الطلب على المشاريع والبنية التحتية", HEX.slate],
  ];
  facts.forEach(([big, txt, col], i) => {
    const y = 1.8 + i * 1.68, x = 7.3, w = R - 7.3;
    card(s, x, y, w, 1.45);
    T(s, big, { x: x + 0.3, y: y + 0.15, w: w - 0.6, h: 0.6, fontSize: 26, bold: true, color: col });
    T(s, txt, { x: x + 0.3, y: y + 0.75, w: w - 0.6, h: 0.6, fontSize: 14, color: HEX.slate, valign: "top" });
  });
}

// ---------- 4. vision ----------
{
  const s = slide("الإطار الاستراتيجي", "GREEN");
  T(s, "الرؤية", { x: L, y: 1.2, w: 11.2, h: 0.5, fontSize: 18, bold: true, color: HEX.gold });
  T(s, "النموذج البلدي الرائد إقليمياً في تنمية الإيرادات والتمويل المبتكر", { x: L, y: 1.75, w: 11.2, h: 1.5, fontSize: 36, bold: true, color: "FFFFFF", valign: "top" });
  T(s, "الرسالة", { x: L, y: 3.6, w: 11.2, h: 0.5, fontSize: 18, bold: true, color: HEX.gold });
  T(s, "تعظيم إيرادات الأمانة وتنويعها، وتوفير حلول تمويل مبتكرة من خارج الميزانية بالشراكة مع القطاعين العام والخاص، بما يدعم مشاريع مدينة الرياض وجودة الحياة فيها ويضمن الاستدامة المالية، وفق أعلى معايير الحوكمة والشفافية.", { x: L, y: 4.1, w: 12.0, h: 1.5, fontSize: 18, color: HEX.mint, valign: "top" });
  T(s, "القيم", { x: L, y: 5.95, w: 11.2, h: 0.5, fontSize: 18, bold: true, color: HEX.gold });
  const vals = ["الحوكمة والشفافية", "الابتكار", "الشراكة", "الكفاءة", "الاستدامة", "القرار المبني على البيانات"];
  vals.forEach((v, i) => {
    const { x, w } = colX(i, 6, 0.2);
    s.addShape(pres.ShapeType.roundRect, { x, y: 6.55, w, h: 0.8, rectRadius: 0.4, fill: { color: "FFFFFF", transparency: 88 }, line: { color: "CFE3DA", width: 1 } });
    T(s, v, { x, y: 6.55, w, h: 0.8, align: "center", fontSize: 15, bold: true, color: "FFFFFF" });
  });
}

// ---------- 5. objectives ----------
{
  const s = slide("الإطار الاستراتيجي");
  title(s, "سبعة أهداف استراتيجية", "لكل هدف إدارة مسؤولة ومؤشرات قياس");
  const goals = [
    ["تعظيم الإيرادات البلدية وتجاوز المستهدف السنوي", "تخطيط ومراقبة الإيرادات"],
    ["تعزيز الرقابة على التحصيل والحد من المتأخرات والفاقد", "تخطيط ومراقبة الإيرادات"],
    ["تنويع مصادر الإيرادات وتقليل الاعتماد على التأجير", "تخطيط ومراقبة الإيرادات"],
    ["تمويل مشاريع المدينة من خارج الميزانية", "الائتمان والتمويل"],
    ["ضمان الاستدامة المالية ورفع كفاءة الإنفاق", "الاستدامة المالية"],
    ["التحول الرقمي وتكامل البيانات المالية", "مشترك"],
    ["التميز المؤسسي وبناء القدرات", "الإدارة العامة"],
  ];
  goals.forEach(([g, own], i) => {
    const r = i < 4 ? 0 : 1, n = r ? 3 : 4, idx = r ? i - 4 : i;
    const { x, w } = colX(idx, n, 0.3);
    const y = 1.9 + r * 3.35, h = 3.0;
    card(s, x, y, w, h);
    badge(s, x + w - 0.95, y + 0.3, 0.65, String(i + 1), [HEX.g, HEX.g, HEX.g, HEX.teal, HEX.gold, HEX.slate, HEX.slate][i]);
    T(s, g, { x: x + 0.3, y: y + 1.1, w: w - 0.6, h: 1.2, fontSize: 17, bold: true, valign: "top" });
    T(s, own, { x: x + 0.3, y: y + h - 0.6, w: w - 0.6, h: 0.4, fontSize: 13, color: HEX.grey });
  });
}

// ---------- 6. 2030 target ----------
{
  const s = slide("السياق والمستهدف");
  title(s, "5 مليارات ريال بحلول 2030", "المسار السنوي المستهدف للإيرادات (مليون ريال)");
  const g = Math.pow(5000 / 3650, 1 / 5);
  const vals = [3650, ...[1, 2, 3, 4, 5].map((k) => Math.round(3650 * Math.pow(g, k)))];
  s.addChart(pres.ChartType.bar, [{ name: "الإيرادات", labels: ["الأساس", "2026", "2027", "2028", "2029", "2030"], values: vals }], {
    x: L, y: 1.8, w: 7.6, h: 6.6, barDir: "col", chartColors: ["9AA5B1", HEX.g, HEX.g, HEX.g, HEX.g, HEX.teal], invertedColors: undefined, showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "#,##0",
    dataLabelFontSize: 14, dataLabelFontBold: true, dataLabelColor: HEX.dk, dataLabelFontFace: "+mn-lt", catAxisLabelFontSize: 14, catAxisLabelColor: HEX.slate, catAxisLabelFontFace: "+mn-lt",
    valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false, barGapWidthPct: 60, valAxisMinVal: 0, valAxisMaxVal: 5600,
  });
  const x = 8.5, w = R - x;
  card(s, x, 1.8, w, 2.4, { fill: HEX.g });
  T(s, `${((g - 1) * 100).toFixed(1)}%`, { x: x + 0.3, y: 1.95, w: w - 0.6, h: 1.1, fontSize: 54, bold: true, color: "FFFFFF" });
  T(s, "نمو سنوي مركّب مطلوب — أقل من النمو الأخير (14%) لأن جزءاً منه مرتبط بالطفرة الإنشائية", { x: x + 0.3, y: 3.05, w: w - 0.6, h: 1.0, fontSize: 14, color: HEX.mint, valign: "top" });
  card(s, x, 4.45, w, 3.95);
  T(s, "سيناريوهات 2030", { x: x + 0.3, y: 4.6, w: w - 0.6, h: 0.5, fontSize: 17, bold: true });
  [["متحفظ · 7%", "5.1 مليار"], ["أساسي · 10%", "5.9 مليار"], ["طموح · 14%", "7.0 مليار"]].forEach(([a, b], i) => {
    const y = 5.25 + i * 1.0;
    s.addShape(pres.ShapeType.line, { x: x + 0.3, y: y - 0.05, w: w - 0.6, h: 0, line: { color: HEX.line, width: 1 } });
    T(s, a, { x: x + 1.9, y, w: w - 2.2, h: 0.8, fontSize: 16, color: HEX.slate });
    T(s, b, { x: x + 0.3, y, w: 1.7, h: 0.8, fontSize: 17, bold: true, color: HEX.g, align: "left" });
  });
}

// ---------- 7. org structure ----------
{
  const s = slide("الهيكل التنظيمي");
  title(s, "إدارة واحدة، ثلاث إدارات، تسعة أقسام");
  const chain = ["سمو الأمين", "مساعد الأمين لإدارة المدينة", "وكالة الخدمات المؤسسية", "إدارة الإيرادات والتمويل"];
  chain.forEach((t, i) => {
    const { x, w } = colX(i, 4, 0.45);
    const on = i === 3;
    card(s, x, 1.65, w, 0.85, { fill: on ? HEX.g : "FFFFFF", line: on ? undefined : HEX.line, flat: !on });
    T(s, t, { x, y: 1.65, w, h: 0.85, align: "center", fontSize: 15, bold: true, color: on ? "FFFFFF" : HEX.slate });
    if (i < 3) T(s, "←", { x: x - 0.45, y: 1.65, w: 0.45, h: 0.85, align: "center", fontSize: 20, color: HEX.grey });
  });
  const depts = [
    ["تخطيط ومراقبة الإيرادات", HEX.g, ["تخطيط الإيرادات والتنبؤ بها", "مراقبة الإيرادات والتحصيل", "تنمية الإيرادات وتنويعها"]],
    ["الائتمان والتمويل", HEX.teal, ["أمانة اللجنة والمحفظة", "التمويل الاستثماري والحكومي", "التمويل البنكي وأسواق المال"]],
    ["الاستدامة المالية", "B7801A", ["السياسات والقواعد المالية", "المخاطر والالتزامات المالية", "الإسقاطات المالية وكفاءة الإنفاق"]],
  ];
  depts.forEach(([d, col, secs], i) => {
    const { x, w } = colX(i, 3, 0.4);
    card(s, x, 2.85, w, 0.95, { fill: col });
    T(s, "إدارة " + d, { x: x + 0.2, y: 2.85, w: w - 0.4, h: 0.95, align: "center", fontSize: 18, bold: true, color: "FFFFFF" });
    secs.forEach((t, j) => {
      const y = 3.95 + j * 0.95;
      card(s, x, y, w, 0.8, { line: HEX.line, flat: true });
      T(s, "قسم " + t, { x: x + 0.25, y, w: w - 0.5, h: 0.8, fontSize: 15, color: HEX.dk });
    });
  });
  card(s, L, 7.0, CW, 1.15, { fill: HEX.lt, flat: true });
  T(s, [{ text: "اللجنة العليا للتمويل: ", options: { bold: true, color: HEX.teal } }, { text: "برئاسة مساعد الأمين لإدارة المدينة، وإدارة الائتمان والتمويل أمين سرها وذراعها التمكيني" }], { x: L + 0.3, y: 7.0, w: CW - 0.6, h: 1.15, fontSize: 16 });
}

// ---------- 8. revenue dept overview ----------
{
  const s = slide("تخطيط ومراقبة الإيرادات");
  title(s, "إدارة تخطيط ومراقبة الإيرادات", "برج المراقبة المالي للأمانة ومصدرها الموحد لبيانات الإيرادات");
  const x = 8.4, w = R - x;
  card(s, x, 1.8, w, 6.6, { fill: HEX.g });
  T(s, "دورها", { x: x + 0.35, y: 2.0, w: w - 0.7, h: 0.5, fontSize: 18, bold: true, color: HEX.gold });
  T(s, "تخطط إيرادات الأمانة وتتنبأ بها وتوزّع مستهدفاتها، وتراقب أداء جهات التحصيل الـ54 وتقارنها، وتكتشف الفاقد، وتنوّع مصادر الإيرادات.", { x: x + 0.35, y: 2.55, w: w - 0.7, h: 2.4, fontSize: 17, color: "FFFFFF", valign: "top" });
  T(s, "أقسامها", { x: x + 0.35, y: 5.1, w: w - 0.7, h: 0.5, fontSize: 18, bold: true, color: HEX.gold });
  T(s, bullets(["تخطيط الإيرادات والتنبؤ بها", "مراقبة الإيرادات والتحصيل", "تنمية الإيرادات وتنويعها"]), { x: x + 0.35, y: 5.6, w: w - 0.7, h: 2.5, fontSize: 16, color: "FFFFFF", valign: "top" });
  const goals = [
    ["تخطيط دقيق بمستهدف لكل جهة", "دقة التنبؤ · نسبة الجهات ذات المستهدف"],
    ["رقابة موحّدة على التحصيل", "الجهات على اللوحة · نسبة التحصيل · المتأخرات"],
    ["تنمية الإيرادات واكتشاف الفاقد", "الفاقد المكتشف والمحصّل · تجاوز المستهدف"],
    ["تنويع مصادر الإيرادات", "الإيرادات من غير التأجير · نمو الإيواء"],
  ];
  goals.forEach(([g, k], i) => {
    const r = Math.floor(i / 2), { x: gx, w: gw } = colX(i % 2, 2, 0.3, L, 7.6);
    const y = 1.8 + r * 3.4;
    card(s, gx, y, gw, 3.15);
    badge(s, gx + gw - 0.95, y + 0.3, 0.65, String(i + 1));
    T(s, g, { x: gx + 0.3, y: y + 1.1, w: gw - 0.6, h: 1.0, fontSize: 18, bold: true, valign: "top" });
    T(s, k, { x: gx + 0.3, y: y + 2.1, w: gw - 0.6, h: 0.85, fontSize: 13, color: HEX.grey, valign: "top" });
  });
}

// ---------- 9. collection map ----------
{
  const s = slide("تخطيط ومراقبة الإيرادات");
  title(s, "54 جهة تحصيل على نظام واحد", "التحصيل موزّع، والرقابة موحّدة");
  const st = [["47", "بلدية فرعية", HEX.g], ["5", "قطاعات", HEX.teal], ["1", "الإدارة المالية", HEX.slate], ["1", "وكالة الاستثمار والتخصيص", "B7801A"]];
  st.forEach(([n, t, col], i) => {
    const { x, w } = colX(i, 4, 0.3);
    card(s, x, 1.8, w, 2.4);
    T(s, n, { x: x + 0.3, y: 1.95, w: w - 0.6, h: 1.2, fontSize: 60, bold: true, color: col });
    T(s, t, { x: x + 0.3, y: 3.2, w: w - 0.6, h: 0.75, fontSize: 17, bold: true, color: HEX.slate, valign: "top" });
  });
  const cols = [
    ["جهات التحصيل", ["تحصّل الإيرادات", "تطالب المتأخرين", "تنفّذ التوصيات"], HEX.slate, "FFFFFF"],
    ["إدارة تخطيط ومراقبة الإيرادات", ["تخطط وتوزّع المستهدفات", "تقيس وتقارن وترتّب الأداء", "تنبّه مبكراً وتقترح الحلول"], HEX.g, HEX.g],
  ];
  cols.forEach(([h, items, col, fill], i) => {
    const { x, w } = colX(i, 2, 0.4, L, CW);
    const dark = fill !== "FFFFFF";
    card(s, x, 4.5, w, 2.6, { fill });
    T(s, h, { x: x + 0.35, y: 4.65, w: w - 0.7, h: 0.55, fontSize: 19, bold: true, color: dark ? "FFFFFF" : col });
    T(s, bullets(items, { space: 4 }), { x: x + 0.35, y: 5.25, w: w - 0.7, h: 1.7, fontSize: 16, color: dark ? "FFFFFF" : HEX.dk, valign: "top" });
  });
  card(s, L, 7.35, CW, 0.95, { fill: "FBF3E2", flat: true });
  T(s, [{ text: "الغرامات: ", options: { bold: true, color: "B7801A" } }, { text: "هدفها الالتزام لا الإيراد؛ نقيس نسبة تحصيل المستحق منها، ولا نضع مستهدفاً لزيادتها" }], { x: L + 0.3, y: 7.35, w: CW - 0.6, h: 0.95, fontSize: 16 });
}

// ---------- 10. revenue sources & initiatives ----------
{
  const s = slide("تخطيط ومراقبة الإيرادات");
  title(s, "مصادر الإيرادات: الفرص والمخاطر", "مرتبة حسب حجمها");
  const src = [
    ["الإيرادات الاستثمارية", "التأجير", "الأكبر والأكثر استقراراً؛ نعالج المتأخرات ونراجع التسعير مع السوق", HEX.g],
    ["الرسوم والتراخيص", "التجارية والإنشائية", "الإنشائية تتبع دورة البناء؛ تنبؤ متحفظ لا يفترض دوام الطفرة", HEX.teal],
    ["الجزاءات والغرامات", "", "للالتزام لا للإيراد؛ نقيس نسبة التحصيل", HEX.slate],
    ["التبغ والإيواء", "", "الإيواء فرصة نمو مع إكسبو وكأس العالم؛ التبغ متوقع تراجعه", "B7801A"],
  ];
  src.forEach(([t, sub, note, col], i) => {
    const { x, w } = colX(i, 4, 0.3);
    card(s, x, 1.8, w, 3.6);
    badge(s, x + w - 0.95, 2.05, 0.65, String(i + 1), col);
    T(s, t, { x: x + 0.3, y: 2.85, w: w - 0.6, h: 0.55, fontSize: 18, bold: true, color: col });
    if (sub) T(s, sub, { x: x + 0.3, y: 3.35, w: w - 0.6, h: 0.4, fontSize: 13, color: HEX.grey });
    T(s, note, { x: x + 0.3, y: 3.8, w: w - 0.6, h: 1.45, fontSize: 14, color: HEX.dk, valign: "top" });
  });
  T(s, "أبرز المبادرات", { x: L, y: 5.7, w: CW, h: 0.5, fontSize: 19, bold: true });
  const ini = [["2026", "اللوحة الموحدة للجهات الـ54"], ["2027", "نموذج التنبؤ وترتيب البلديات"], ["2028", "برنامج كشف الفاقد واجتماع المراجعة الشهري"], ["2029–2030", "مصادر جديدة وخطة إيرادات الأحداث الكبرى"]];
  ini.forEach(([y0, t], i) => {
    const { x, w } = colX(i, 4, 0.3);
    card(s, x, 6.3, w, 2.0, { fill: HEX.lt, flat: true });
    T(s, y0, { x: x + 0.3, y: 6.45, w: w - 0.6, h: 0.5, fontSize: 18, bold: true, color: HEX.g });
    T(s, t, { x: x + 0.3, y: 7.0, w: w - 0.6, h: 1.15, fontSize: 15, valign: "top" });
  });
}

// ---------- 11. four tracks ----------
{
  const s = slide("الائتمان والتمويل");
  title(s, "أربعة مسارات لتمويل المشاريع من خارج الميزانية", "إدارة الائتمان والتمويل تدرس كل مشروع وتوصي بالمسار الأنسب");
  const tr = [
    ["المستثمر", "BOT، الامتياز، الخدمة مقابل الإعلان، التطوير مقابل الأرض، الرعايات", "مشاريع تولّد دخلاً", HEX.g, "بلا دين"],
    ["جهة حكومية", "تمويل كامل أو مشترك مع جهة لها مصلحة في المشروع", "أثر استراتيجي مشترك", HEX.teal, "بلا دين"],
    ["البنوك", "تمويل المشاريع، التمويل عبر المقاول، التسهيلات", "مشاريع كبيرة بمصدر سداد واضح", HEX.slate, "دين"],
    ["السندات والصكوك", "صكوك مرتبطة بإيرادات مشروع، صكوك خضراء", "محافظ كبيرة طويلة الأجل", "B7801A", "دين"],
  ];
  tr.forEach(([t, tools, fit, col, debt], i) => {
    const { x, w } = colX(i, 4, 0.3);
    card(s, x, 1.85, w, 4.6);
    s.addShape(pres.ShapeType.roundRect, { x, y: 1.85, w, h: 1.1, rectRadius: 0.16, fill: { color: col }, line: { type: "none" } });
    s.addShape(pres.ShapeType.rect, { x, y: 2.6, w, h: 0.35, fill: { color: col }, line: { type: "none" } });
    T(s, `${i + 1}. ${t}`, { x: x + 0.3, y: 1.85, w: w - 0.6, h: 1.1, fontSize: 20, bold: true, color: "FFFFFF" });
    T(s, tools, { x: x + 0.3, y: 3.15, w: w - 0.6, h: 1.6, fontSize: 14, valign: "top" });
    T(s, [{ text: "يناسب: ", options: { bold: true } }, { text: fit }], { x: x + 0.3, y: 4.8, w: w - 0.6, h: 0.9, fontSize: 14, color: HEX.slate, valign: "top" });
    s.addShape(pres.ShapeType.roundRect, { x: x + 0.3, y: 5.8, w: 1.5, h: 0.42, rectRadius: 0.21, fill: { color: debt === "دين" ? "FBE3E1" : HEX.mint }, line: { type: "none" } });
    T(s, debt, { x: x + 0.3, y: 5.8, w: 1.5, h: 0.42, align: "center", fontSize: 13, bold: true, color: debt === "دين" ? "B4483F" : HEX.g });
  });
  card(s, L, 6.75, CW, 0.7, { fill: HEX.g, flat: true });
  T(s, "الأولوية: المستثمر ← الجهة الحكومية ← البنوك ← السندات", { x: L + 0.3, y: 6.75, w: CW - 0.6, h: 0.7, fontSize: 18, bold: true, color: "FFFFFF" });
  card(s, L, 7.6, CW, 0.75, { fill: HEX.lt, flat: true });
  T(s, [{ text: "مصادر السداد: ", options: { bold: true, color: HEX.teal } }, { text: "إيرادات الأمانة، ثم إيرادات المشروع نفسه، ثم مخصصات الميزانية — وكل تمويل بنكي أو صكوك يمر على الاستدامة المالية" }], { x: L + 0.3, y: 7.6, w: CW - 0.6, h: 0.75, fontSize: 15 });
}

// ---------- 12. project path ----------
{
  const s = slide("الائتمان والتمويل");
  title(s, "مسار المشروع من الفكرة إلى التنفيذ", "سبع خطوات، ولكل خطوة مسؤول واضح");
  const steps = [
    ["تحديد الاحتياج", "إدارات المشاريع والقطاعات والبلديات"],
    ["الفرز الأولي", "الائتمان والتمويل"],
    ["الدراسة والهيكلة ومقارنة المسارات", "الائتمان والتمويل"],
    ["مراجعة الاستدامة", "الاستدامة المالية"],
    ["القرار", "اللجنة العليا للتمويل"],
    ["الطرح أو اتفاقية التمويل", "وكالة الاستثمار / الإدارة / الجهة الشريكة"],
    ["متابعة الأثر والوفر", "الائتمان والتمويل"],
  ];
  const cols = [HEX.slate, HEX.teal, HEX.teal, "B7801A", HEX.g, HEX.slate, HEX.teal];
  s.addShape(pres.ShapeType.line, { x: 1.2, y: 3.0, w: 10.9, h: 0, line: { color: HEX.line, width: 3 } });
  steps.forEach(([t, own], i) => {
    const { x, w } = colX(i, 7, 0.2);
    badge(s, x + w / 2 - 0.45, 2.55, 0.9, String(i + 1), cols[i], "FFFFFF", 22);
    card(s, x, 3.75, w, 3.9);
    T(s, t, { x: x + 0.15, y: 3.95, w: w - 0.3, h: 1.6, fontSize: 16, bold: true, align: "center", valign: "top" });
    s.addShape(pres.ShapeType.line, { x: x + 0.3, y: 5.65, w: w - 0.6, h: 0, line: { color: HEX.line, width: 1 } });
    T(s, own, { x: x + 0.15, y: 5.8, w: w - 0.3, h: 1.7, fontSize: 13, color: cols[i], bold: true, align: "center", valign: "top" });
  });
}

// ---------- 13. committee ----------
{
  const s = slide("الائتمان والتمويل");
  title(s, "أمين سر اللجنة العليا للتمويل وذراعها التمكيني", "اللجنة برئاسة مساعد الأمين لإدارة المدينة");
  const x = 8.6, w = R - x;
  card(s, x, 1.8, w, 6.55, { fill: HEX.teal });
  T(s, "أمانة السر", { x: x + 0.35, y: 2.0, w: w - 0.7, h: 0.55, fontSize: 20, bold: true, color: "FFFFFF" });
  [["قبل الاجتماع", "الدراسة، ومقارنة المسارات، ورأي الاستدامة، والملف والتوصية"], ["أثناء الاجتماع", "عرض الملفات وتدوين المحضر والقرارات"], ["بعد الاجتماع", "إبلاغ جهات التنفيذ ومتابعة القرارات والتقارير"]].forEach(([h, t], i) => {
    const y = 2.75 + i * 1.8;
    T(s, h, { x: x + 0.35, y, w: w - 0.7, h: 0.45, fontSize: 16, bold: true, color: HEX.gold });
    T(s, t, { x: x + 0.35, y: y + 0.5, w: w - 0.7, h: 1.1, fontSize: 15, color: "FFFFFF", valign: "top" });
  });
  const en = [
    ["فني", "دراسات الجدوى والنماذج المالية والتوصيات"],
    ["تنظيمي", "الميثاق ومعايير المسار والسياسات والأدلة"],
    ["العلاقات", "جاهزية البنوك والجهات الحكومية والمستثمرين"],
    ["المعلومات", "لوحة محفظة التمويل ورصد السوق"],
    ["نظامي", "ملفات موافقات المالية وإدارة الدين والسوق المالية والتخصيص"],
    ["القدرات", "تأهيل كوادر الأمانة في التمويل والشراكات"],
  ];
  T(s, "التمكين", { x: L, y: 1.8, w: 7.8, h: 0.5, fontSize: 20, bold: true, color: HEX.teal });
  en.forEach(([h, t], i) => {
    const r = Math.floor(i / 2), { x: ex, w: ew } = colX(i % 2, 2, 0.3, L, 7.8);
    const y = 2.45 + r * 2.0;
    card(s, ex, y, ew, 1.8);
    T(s, "التمكين " + h, { x: ex + 0.3, y: y + 0.15, w: ew - 0.6, h: 0.5, fontSize: 17, bold: true, color: HEX.teal });
    T(s, t, { x: ex + 0.3, y: y + 0.7, w: ew - 0.6, h: 0.95, fontSize: 14, valign: "top" });
  });
}

// ---------- 14. park model ----------
{
  const s = slide("الائتمان والتمويل");
  title(s, "نموذج تطبيقي: الحديقة", "مستثمر ينشئ ويصون ويشغّل ويستثمر 15 سنة ثم تعود الحديقة للأمانة");
  const x = 9.0, w = R - x;
  card(s, x, 1.8, w, 3.0, { fill: HEX.g });
  T(s, "≈ 50 مليون", { x: x + 0.3, y: 1.95, w: w - 0.6, h: 1.0, fontSize: 40, bold: true, color: "FFFFFF" });
  T(s, "ريال وفر توضيحي خلال 15 سنة (إنشاء 20 مليون + صيانة مليونان سنوياً) — أرقام افتراضية", { x: x + 0.3, y: 2.95, w: w - 0.6, h: 1.7, fontSize: 14, color: HEX.mint, valign: "top" });
  card(s, x, 5.0, w, 3.35);
  T(s, "لمن إيراد التسمية؟", { x: x + 0.3, y: 5.15, w: w - 0.6, h: 0.5, fontSize: 17, bold: true, color: HEX.teal });
  T(s, bullets(["حديقة جاذبة: تحتفظ به الأمانة", "حديقة ضعيفة الجاذبية: يُعطى للمستثمر ليضمن التنفيذ"], { space: 8 }), { x: x + 0.3, y: 5.7, w: w - 0.6, h: 2.5, fontSize: 15, valign: "top" });
  const two = [["الأمانة تحصل على", ["حديقة بلا تكلفة إنشاء", "صيانة وتشغيل 15 سنة", "أصل يعود إليها", "إيجار أو نسبة + إيراد التسمية"]], ["المستثمر يربح من", ["مطاعم ومقاهٍ في جزء محدد", "أنشطة وألعاب وفعاليات", "إعلانات ومواقف", "التسمية (حسب الاتفاق)"]]];
  two.forEach(([h, items], i) => {
    const { x: cx, w: cw } = colX(i, 2, 0.3, L, 8.2);
    card(s, cx, 1.8, cw, 3.0);
    T(s, h, { x: cx + 0.3, y: 1.95, w: cw - 0.6, h: 0.5, fontSize: 17, bold: true, color: i ? HEX.slate : HEX.g });
    T(s, bullets(items, { space: 3 }), { x: cx + 0.3, y: 2.5, w: cw - 0.6, h: 2.15, fontSize: 15, valign: "top" });
  });
  card(s, L, 5.0, 8.2, 3.35, { fill: HEX.lt, flat: true });
  T(s, "اشتراطات حماية المصلحة العامة", { x: L + 0.3, y: 5.15, w: 7.6, h: 0.5, fontSize: 17, bold: true, color: HEX.g });
  const cond = ["دخول مجاني أو رمزي", "سقف للمساحة التجارية", "مؤشرات صيانة وغرامات", "ساعات تخدم الجمهور", "تسليم بحالة تشغيلية جيدة", "ضمان بنكي للأداء", "حق التدخل عند التعثر", "موافقة الأمانة على اسم الراعي"];
  [cond.slice(0, 4), cond.slice(4)].forEach((items, i) => {
    const { x: cx, w: cw } = colX(i, 2, 0.3, L + 0.3, 7.6);
    T(s, bullets(items, { space: 5 }), { x: cx, y: 5.7, w: cw, h: 2.0, fontSize: 15, valign: "top" });
  });
  T(s, "الحدائق الصغيرة: تُجمَّع 5–10 منها مع حديقة كبيرة مربحة في عقد واحد", { x: L + 0.3, y: 7.75, w: 7.6, h: 0.5, fontSize: 14, bold: true, color: HEX.teal });
}

// ---------- 15. sustainability ----------
{
  const s = slide("الاستدامة المالية");
  title(s, "إدارة الاستدامة المالية", "تضمن ألا تصنع قرارات اليوم عبئاً على الأمانة غداً");
  const trio = [["الإيرادات", "هذا ما سنحصّله", HEX.g], ["التمويل", "هذا ما نحتاج تمويله", HEX.teal], ["الاستدامة", "هذا ما نتحمله بأمان", "B7801A"]];
  trio.forEach(([h, t, col], i) => {
    const { x, w } = colX(i, 3, 0.3);
    card(s, x, 1.8, w, 1.4, { fill: col });
    T(s, h, { x: x + 0.3, y: 1.9, w: w - 0.6, h: 0.5, fontSize: 16, bold: true, color: "FFFFFF" });
    T(s, `«${t}»`, { x: x + 0.3, y: 2.4, w: w - 0.6, h: 0.65, fontSize: 21, bold: true, color: "FFFFFF" });
  });
  T(s, "القواعد المالية المقترحة", { x: L, y: 3.45, w: CW, h: 0.5, fontSize: 19, bold: true });
  const rules = [
    ["سقف خدمة الدين", "الأقساط لا تتجاوز X% من الإيرادات المستقرة"], ["سقف الالتزامات", "دفعات الإتاحة والضمانات لا تتجاوز Y%"], ["لا تمويل بلا سداد", "مصدر سداد محدد بالأرقام"], ["القاعدة الذهبية", "الاقتراض للمشاريع الرأسمالية فقط"],
    ["تنويع الآجال", "لا تتركز الأقساط في سنة واحدة"], ["هامش أمان", "جزء من القدرة يبقى للطوارئ"], ["مراجعة إلزامية", "للصفقات فوق حد المدة أو القيمة"], ["تكلفة دورة الحياة", "التشغيل والصيانة طوال العمر"],
  ];
  rules.forEach(([h, t], i) => {
    const r = Math.floor(i / 4), { x, w } = colX(i % 4, 4, 0.3);
    const y = 4.05 + r * 2.15;
    card(s, x, y, w, 1.95);
    badge(s, x + w - 0.8, y + 0.25, 0.55, String(i + 1), "B7801A", "FFFFFF", 13);
    T(s, h, { x: x + 0.3, y: y + 0.25, w: w - 1.2, h: 0.55, fontSize: 16, bold: true });
    T(s, t, { x: x + 0.3, y: y + 0.95, w: w - 0.6, h: 0.85, fontSize: 14, color: HEX.slate, valign: "top" });
  });
  T(s, "تعتمدها اللجنة العليا للتمويل ثم سمو الأمين، وتُحدَّد النسب بعد خط الأساس", { x: L, y: 8.38, w: CW, h: 0.4, fontSize: 13, color: HEX.grey });
}

// ---------- 16. repayment suitability & stress ----------
{
  const s = slide("الاستدامة المالية");
  title(s, "على أي إيرادات نبني قدرة الاقتراض؟", "سقوف الاقتراض تُحسب على الإيرادات المستقرة فقط");
  const rows = [
    ["الإيرادات الاستثمارية", "مناسبة جداً — مستقرة وتعاقدية", "2E9E6B"], ["الرسوم والتراخيص", "بحذر — الإنشائية تتذبذب", HEX.gold], ["الإيواء", "بحذر — مرتبط بالسياحة والأحداث", HEX.gold],
    ["الجزاءات والغرامات", "غير مناسبة — هدفها الالتزام", HEX.pink], ["التبغ", "غير مناسب — متوقع تراجعه", HEX.pink],
  ];
  const x0 = 5.6, w0 = R - x0;
  rows.forEach(([h, t, col], i) => {
    const y = 1.8 + i * 1.3;
    card(s, x0, y, w0, 1.12);
    s.addShape(pres.ShapeType.ellipse, { x: x0 + w0 - 0.7, y: y + 0.36, w: 0.4, h: 0.4, fill: { color: col }, line: { type: "none" } });
    T(s, h, { x: x0 + 3.4, y, w: w0 - 4.2, h: 1.12, fontSize: 17, bold: true });
    T(s, t, { x: x0 + 0.3, y, w: 3.1, h: 1.12, fontSize: 14, color: HEX.slate });
  });
  const w = 4.8;
  card(s, L, 1.8, w, 3.9, { fill: HEX.slate });
  T(s, "اختبارات الضغط (2028)", { x: L + 0.3, y: 1.95, w: w - 0.6, h: 0.5, fontSize: 17, bold: true, color: HEX.gold });
  T(s, bullets(["انخفاض ميزانية الدولة", "تباطؤ سوق البناء", "ارتفاع أسعار الفائدة (سايبور)", "تعثّر مستثمر كبير"], { space: 6 }), { x: L + 0.3, y: 2.55, w: w - 0.6, h: 3.0, fontSize: 16, color: "FFFFFF", valign: "top" });
  card(s, L, 5.9, w, 2.45, { fill: "FBF3E2", flat: true });
  T(s, "كفاءة الإنفاق وإرث الأحداث", { x: L + 0.3, y: 6.05, w: w - 0.6, h: 0.5, fontSize: 16, bold: true, color: "B7801A" });
  T(s, "الإدارة تقود برنامج كفاءة الإنفاق في الأمانة، وتحسب تكلفة تشغيل أصول إكسبو وكأس العالم بعد انتهائهما", { x: L + 0.3, y: 6.6, w: w - 0.6, h: 1.6, fontSize: 14, valign: "top" });
}

// ---------- 17. roadmap ----------
{
  const s = slide("خارطة الطريق والمؤشرات");
  title(s, "خارطة الطريق 2026–2030");
  const ph = [["2026", "التأسيس"], ["2027", "الأدوات والمكاسب السريعة"], ["2028", "العمق والتوسع"], ["2029–2030", "التنويع والاستعداد"]];
  const rowsD = [
    ["الإيرادات", HEX.g, ["توحيد البيانات واللوحة الموحدة", "نموذج التنبؤ وتوزيع المستهدفات", "كشف الفاقد واجتماع المراجعة", "مصادر جديدة وإيرادات الأحداث"]],
    ["التمويل", HEX.teal, ["حصر المشاريع وميثاق اللجنة", "3–5 مشاريع تجريبية والرعايات", "مشاريع كبرى وأول تمويل بنكي", "محفظة إكسبو ودراسة الصكوك"]],
    ["الاستدامة", "B7801A", ["سجل الالتزامات والقواعد المالية", "إسقاطات 10 سنوات وأول تقرير", "اختبارات الضغط وتكلفة الخدمات", "الجاهزية للتصنيف وإرث الأحداث"]],
  ];
  const lw = 1.9, gx = L, gw = CW - lw - 0.2;
  ph.forEach(([y0, t], i) => {
    const { x, w } = colX(i, 4, 0.2, gx, gw);
    card(s, x, 1.55, w, 1.2, { fill: HEX.dk, flat: true });
    T(s, y0, { x: x + 0.2, y: 1.6, w: w - 0.4, h: 0.6, fontSize: 20, bold: true, color: HEX.gold, align: "center" });
    T(s, t, { x: x + 0.2, y: 2.15, w: w - 0.4, h: 0.5, fontSize: 14, color: "FFFFFF", align: "center" });
  });
  rowsD.forEach(([d, col, cells], r) => {
    const y = 3.0 + r * 1.85;
    card(s, R - lw, y, lw, 1.65, { fill: col, flat: true });
    T(s, d, { x: R - lw, y, w: lw, h: 1.65, fontSize: 18, bold: true, color: "FFFFFF", align: "center" });
    cells.forEach((t, i) => {
      const { x, w } = colX(i, 4, 0.2, gx, gw);
      card(s, x, y, w, 1.65, { line: HEX.line, flat: true });
      T(s, t, { x: x + 0.2, y, w: w - 0.4, h: 1.65, fontSize: 15, align: "center" });
    });
  });
}

// ---------- 18. KPIs ----------
{
  const s = slide("خارطة الطريق والمؤشرات");
  title(s, "مؤشرات الأداء الرئيسية", "بقية المستهدفات تُحدَّد بعد بناء خط الأساس في 2026");
  const k = [
    ["5,000", "مليون ريال", "إجمالي الإيرادات في 2030", HEX.g],
    ["100%", "فأكثر سنوياً", "تحقيق مستهدف الإيرادات", HEX.g],
    ["54 / 54", "من 2026", "جهات على اللوحة الموحدة", HEX.teal],
    ["100%", "قبل الطرح", "مشاريع بدراسة القيمة مقابل المال", HEX.teal],
    ["100%", "قبل الاعتماد", "صفقات تمر على مراجعة الاستدامة", "B7801A"],
    ["ضمن السقف", "ربع سنوي", "خدمة الدين إلى الإيرادات المستقرة", "B7801A"],
  ];
  k.forEach(([v, u, t, col], i) => {
    const r = Math.floor(i / 3), { x, w } = colX(i % 3, 3, 0.35);
    const y = 1.85 + r * 3.35;
    card(s, x, y, w, 3.05);
    T(s, v, { x: x + 0.3, y: y + 0.3, w: w - 0.6, h: 1.1, fontSize: 46, bold: true, color: col });
    T(s, u, { x: x + 0.3, y: y + 1.4, w: w - 0.6, h: 0.45, fontSize: 14, color: HEX.grey });
    T(s, t, { x: x + 0.3, y: y + 1.95, w: w - 0.6, h: 0.9, fontSize: 17, bold: true, valign: "top" });
  });
}

// ---------- 19. risks ----------
{
  const s = slide("خارطة الطريق والمؤشرات");
  title(s, "أبرز المخاطر والمعالجة");
  const rk = [
    ["تباطؤ سوق البناء", "تنبؤ متحفظ وتنويع المصادر"],
    ["تداخل الأدوار مع الوكالة والإدارة المالية", "مصفوفة مسؤوليات واتفاقيات مستوى خدمة"],
    ["ضعف إقبال المستثمرين", "استطلاع السوق والتجميع وحوافز التسمية"],
    ["التزامات مخفية في العقود", "مراجعة إلزامية وسجل الالتزامات"],
    ["ارتفاع أسعار الفائدة", "تنويع الآجال وهامش أمان"],
    ["نقص الكفاءات وجودة البيانات", "تأهيل وشهادات، ومصدر بيانات موحد"],
  ];
  rk.forEach(([r0, m], i) => {
    const row = Math.floor(i / 2), { x, w } = colX(i % 2, 2, 0.35);
    const y = 1.7 + row * 2.25;
    card(s, x, y, w, 2.0);
    s.addShape(pres.ShapeType.roundRect, { x: x + w - 1.35, y: y + 0.3, w: 1.05, h: 0.42, rectRadius: 0.21, fill: { color: "FBE3E1" }, line: { type: "none" } });
    T(s, "الخطر", { x: x + w - 1.35, y: y + 0.3, w: 1.05, h: 0.42, fontSize: 13, bold: true, color: "B4483F", align: "center" });
    T(s, r0, { x: x + 0.3, y: y + 0.25, w: w - 1.85, h: 0.55, fontSize: 18, bold: true });
    s.addShape(pres.ShapeType.roundRect, { x: x + w - 1.35, y: y + 1.15, w: 1.05, h: 0.42, rectRadius: 0.21, fill: { color: HEX.mint }, line: { type: "none" } });
    T(s, "المعالجة", { x: x + w - 1.35, y: y + 1.15, w: 1.05, h: 0.42, fontSize: 13, bold: true, color: HEX.g, align: "center" });
    T(s, m, { x: x + 0.3, y: y + 1.1, w: w - 1.85, h: 0.55, fontSize: 16, color: HEX.slate });
  });
}

// ---------- 20. decisions ----------
{
  const s = slide("القرارات المطلوبة", "GREEN");
  title(s, "القرارات المطلوبة", "لبدء التنفيذ في 2026", true);
  const d = [
    "اعتماد الاستراتيجية والهيكل التنظيمي للإدارة",
    "اعتماد ميثاق اللجنة العليا للتمويل وتسمية أعضائها",
    "اعتماد مستهدف 5 مليارات ريال بحلول 2030",
    "توجيه الجهات بتزويد البيانات لبناء خط الأساس",
    "اعتماد اتفاقية مستوى الخدمة مع وكالة الاستثمار والتخصيص",
    "البدء في حصر المشاريع غير الممولة",
  ];
  d.forEach((t, i) => {
    const row = Math.floor(i / 2), { x, w } = colX(i % 2, 2, 0.35);
    const y = 1.95 + row * 2.2;
    s.addShape(pres.ShapeType.roundRect, { x, y, w, h: 1.9, rectRadius: 0.16, fill: { color: "FFFFFF", transparency: 90 }, line: { color: "5FA98A", width: 1 } });
    badge(s, x + w - 1.05, y + 0.55, 0.8, String(i + 1), HEX.gold, HEX.dk, 20);
    T(s, t, { x: x + 0.35, y, w: w - 1.6, h: 1.9, fontSize: 19, bold: true, color: "FFFFFF" });
  });
}

pres.writeFile({ fileName: OUT }).then(async () => { await applyTheme(OUT, THEME); console.log("wrote", OUT); });
