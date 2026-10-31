// سلسلة العمل: Python للأرقام، Ollama للمسودات، Claude للمراجعة والتحرير، وFaisal يعتمد.
const fs = require('fs');
const path = require('path');

const MONTHS = ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو', 'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'];
const DEPTS = [
  { id: 'lic', name: 'الرسوم والتراخيص', tab: 'الرسوم والتراخيص' },
  { id: 'inv', name: 'الإيرادات الاستثمارية', tab: 'إيرادات إستثمارية' },
  { id: 'pen', name: 'الجزاءات والغرامات', tab: 'جزاءات وغرامات' },
  { id: 'tob', name: 'التبغ والإيواء', tab: 'التبغ والإيواء' },
  { id: 'oth', name: 'الإيرادات الأخرى', tab: 'أخرى' }
];
const norm = (s) => String(s || '').replace(/[أإآ]/g, 'ا').replace(/ى/g, 'ي').replace(/ة/g, 'ه').replace(/\s+/g, ' ').trim();

const WRITING_RULES = `قواعد الكتابة (دليل التحليل المعتمد):
- صيغة الجهة أو المبني للمجهول، لا المتكلم: «يقدر» لا «نقدر».
- «مليون» مع الأرقام العشرية، و«ريال» بعد المبالغ في النص.
- الهمزات: إصدار، إيرادات، إلغاء، إشغال، الأراضي، الإيواء، تأسيس، استثمار.
- بدون تشكيل، وتبقى كلمة «شهر» قبل أسماء الأشهر.
- «إضافة إلى» بدل «بالإضافة إلى».
- لا عبارات ظن مثل «يبدو أن»، ولا ألفاظ أجنبية.
- السبب بدون دليل مباشر يُكتب «مؤشر»، والرقم التقديري يُكتب «تقدير».
- ممنوع أي رقم مو موجود في البيانات المعطاة، وممنوع أسماء أشخاص أو أرقام فواتير.`;

// ------------------------------------------------------------ intents
function parseRequest(text, now = new Date()) {
  const t = norm(text);
  if (/فحص|افحص|تحقق من الملفات|العناوين/.test(t)) return { kind: 'inspect' };
  if (/متابعه|المتابعه|ملف المتابعه|حدث/.test(t)) {
    let month = null;
    MONTHS.forEach((m, i) => { if (t.includes(norm(m))) month = i + 1; });
    const ym = t.match(/(20\d\d)/);
    let year = ym ? +ym[1] : now.getFullYear();
    if (!month) { month = now.getMonth(); if (month === 0) { month = 12; year -= 1; } }
    return { kind: 'monthly', year, month };
  }
  return { kind: 'unknown' };
}

// ------------------------------------------------------------ helpers
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
function say(store, from, to, text, extra = {}) { return store.addChat({ from, to, text, kind: 'finding', ...extra }); }
function work(store, id, text, detail = '', screen) { store.setStatus(id, { kind: 'work', text, detail, ...(screen ? { screen } : {}) }); }
function done(store, id, text, detail = '') { store.setStatus(id, { kind: 'done', text, detail }); }
function alert(store, id, text, detail = '') { store.setStatus(id, { kind: 'alert', text, detail }); }

function parseJSON(text) {
  const m = String(text).match(/\{[\s\S]*\}/);
  if (!m) return null;
  try { return JSON.parse(m[0]); } catch (_) { return null; }
}

function deptData(summary, tab) {
  const k = norm(tab);
  return {
    period: summary.period,
    ytd: summary.ytd.filter((r) => norm(r.tab) === k),
    month: summary.month.filter((r) => norm(r.tab) === k)
  };
}

// ------------------------------------------------------------ inspect
async function runInspect(ctx) {
  const { store, engine, settings, workDir } = ctx;
  store.setRun({ active: true, stage: 1, label: 'فحص ملفات الدرايف' });
  work(store, 'recv', 'يفحص ملفات الدرايف', 'يقرأ العناوين فقط، بدون بيانات');
  const cfgPath = engine.writeConfig(settings.engineConfig(), workDir);
  const items = [];
  await engine.run(['inspect', '--config', cfgPath], (ev) => { if (ev.ev === 'inspect') items.push(ev.item); });
  const lines = items.map((it) => it.exists
    ? `✓ ${it.source}: ${it.size_mb} ميقا، ${(it.sheets || []).length} ورقة${it.error ? ' — خطأ: ' + it.error : ''}`
    : `✗ ${it.source}: ما لقيت الملف (${it.path})`);
  done(store, 'recv', 'خلّص الفحص', lines.join(' | '));
  store.setStatus('recv', { screen: lines.slice(0, 8) });
  say(store, 'recv', 'faisal', 'نتيجة فحص الدرايف: ' + lines.join('، '));
  store.setRun({ stage: 0, label: 'بانتظار طلبك' });
  const report = path.join(workDir, 'inspect.json');
  fs.writeFileSync(report, JSON.stringify(items, null, 1));
  return { items, report };
}

// ------------------------------------------------------------ monthly
async function runMonthly(ctx, { year, month }) {
  const { store, engine, llm, settings, workDir, renderPdf } = ctx;
  const S = settings.get();
  const label = `متابعة شهر ${MONTHS[month - 1]} ${year}`;
  const runDir = path.join(workDir, 'runs', `${year}-${String(month).padStart(2, '0')}-${Date.now()}`);
  fs.mkdirSync(runDir, { recursive: true });

  store.setRun({ active: true, stage: 0, label, awaiting: null, files: null });
  work(store, 'faisal', 'طلب ' + label);
  say(store, 'faisal', 'angel', `حدّث ملف ${label}.`, { kind: 'handoff', via: 'email' });
  await sleep(800);
  store.setStatus('faisal', { kind: 'wait', text: 'ينتظر التقرير' });

  // 1) الأرقام: Python
  store.setRun({ stage: 1 });
  const cfgPath = engine.writeConfig(settings.engineConfig(), runDir);
  let result;
  try {
    result = await engine.run(['monthly', '--config', cfgPath, '--year', String(year), '--month', String(month), '--out', runDir], (ev) => {
      if (ev.ev === 'status') store.setStatus(ev.id, { kind: ev.kind, text: ev.text, detail: ev.detail, screen: ev.screen, at: ev.at });
      else if (ev.ev === 'chat') store.addChat({ t: ev.at, from: ev.from, to: ev.to, kind: ev.kind, text: ev.text });
    });
  } catch (e) {
    alert(store, 'angel', 'توقف التشغيل', e.message);
    say(store, 'angel', 'faisal', 'وقف التشغيل في مرحلة الأرقام: ' + e.message, { kind: 'return' });
    store.setRun({ stage: 0, label: 'توقف: ' + label });
    throw e;
  }
  const summaryPath = result.summary;
  const summary = JSON.parse(fs.readFileSync(summaryPath, 'utf8'));
  const failed = summary.checks.filter((c) => !c.ok);
  if (failed.length) {
    alert(store, 'angel', 'يرفع ملاحظات التدقيق', failed.map((c) => c.check).join('، '));
    say(store, 'angel', 'faisal', `ملاحظات تحتاج قرارك (${failed.length}): ` + failed.map((c) => `${c.check} (${c.note})`).join('، ') + '. كمّلت المسودة وما راح تُنشر إلا باعتمادك.', { kind: 'finding', via: 'email' });
  } else {
    done(store, 'angel', 'اعتمد البيانات', 'كل فحوص الأرقام سليمة');
    say(store, 'angel', 'lic0', 'البيانات معتمدة، وزّعتها على الأقسام الخمسة.', { kind: 'handoff' });
  }

  // 2) الأقسام: Ollama يكتب، والقائد (Claude) يراجع
  store.setRun({ stage: 2 });
  const notes = [];
  for (const d of DEPTS) {
    const data = deptData(summary, d.tab);
    const writer = d.id + '1', lead = d.id + '0';
    if (!data.ytd.length) { done(store, writer, 'ما فيه بيانات للقسم'); continue; }
    const dataJson = JSON.stringify(data);
    work(store, writer, 'يكتب ملاحظة القسم', `مسودة ${d.name} من أرقام المحرك`, data.ytd.slice(0, 4).map((r) => `${r.org}: فعلي ${(r.actual / 1e6).toFixed(1)} مليون`));
    let draft;
    try {
      draft = await llm.draft([
        { role: 'system', content: `أنت محلل في قسم ${d.name} بفريق تخطيط ومراقبة الإيرادات في أمانة منطقة الرياض. تكتب ملاحظة قصيرة (3 إلى 5 جمل) عن أداء القسم من البيانات فقط.\n${WRITING_RULES}\nالأرقام بالريال؛ اكتبها بالمليون بخانة عشرية وحدة. النسب المئوية موجودة جاهزة في الحقول المنتهية بـ _pct. لا تحسب أرقام جديدة.` },
        { role: 'user', content: `بيانات ${label}:\n${dataJson}\n\nاكتب الملاحظة الحين، نص عادي بدون عناوين.` }
      ]);
    } catch (e) {
      alert(store, writer, 'تعذّر الكتابة', e.message);
      say(store, writer, lead, 'ما قدرت أكتب المسودة: ' + e.message, { kind: 'return' });
      draft = '';
    }
    let chk = draft ? await engine.numcheck(summaryPath, draft) : { ok: false, unknown: [] };
    done(store, writer, 'سلّم المسودة', chk.ok ? 'الأرقام كلها من البيانات' : 'أرقام تحتاج مراجعة: ' + chk.unknown.join('، '));
    say(store, writer, lead, `سلّمت مسودة ${d.name}.` + (chk.ok ? '' : ' فيها أرقام ما طابقت: ' + chk.unknown.join('، ')), { kind: 'handoff' });

    work(store, lead, 'يراجع المسودة', `قائد ${d.name}: يتأكد من الأرقام والاستنتاجات`);
    let final = draft, decision = 'approve', reason = '';
    try {
      const out = await llm.review(
        `أنت قائد قسم ${d.name}. راجع مسودة المحلل مقابل البيانات.\n<بيانات>${dataJson}</بيانات>\n<مسودة>${draft || '(فاضية)'}</مسودة>\n` +
        (chk.ok ? '' : `<أرقام_غير_مطابقة>${chk.unknown.join('، ')}</أرقام_غير_مطابقة>\n`) +
        `<قرار>إذا المسودة سليمة اعتمدها، وإذا فيها خطأ أو رقم مو من البيانات أو استنتاج غير مبرر صحّحها بنفسك.</قرار>\n` +
        `رد بـ JSON فقط: {"decision":"approve" أو "fix","note":"النص النهائي للملاحظة","reason":"سبب قصير"}`,
        WRITING_RULES);
      const j = parseJSON(out);
      if (j && j.note) { final = j.note; decision = j.decision || 'approve'; reason = j.reason || ''; }
      else if (out) { final = out; decision = 'fix'; }
    } catch (e) {
      alert(store, lead, 'تعذّرت المراجعة', e.message);
      reason = 'المراجعة ما تمت: ' + e.message;
      decision = 'pending';
    }
    chk = final ? await engine.numcheck(summaryPath, final) : { ok: false, unknown: [] };
    notes.push({ dept: d.name, id: d.id, text: final, decision, reason, numbers_ok: chk.ok, unknown: chk.unknown });
    if (decision === 'pending') { /* بقي التنبيه */ }
    else if (chk.ok) done(store, lead, decision === 'approve' ? 'اعتمد ملاحظة القسم' : 'صحّح واعتمد', reason);
    else alert(store, lead, 'أرقام غير مطابقة', chk.unknown.join('، '));
    say(store, lead, 'aziz', `${decision === 'approve' ? 'اعتمدت' : 'صحّحت'} ملاحظة ${d.name}.${reason ? ' ' + reason : ''}`, { kind: decision === 'approve' ? 'approve' : 'return' });
  }

  // 3) Salah سبق دقق الأرقام في المحرك؛ Aziz يدقق المنطق ويحرر اللغة
  store.setRun({ stage: 5 });
  await sleep(400);
  store.setRun({ stage: 6 });
  work(store, 'aziz', 'يدقق المنطق ويحرر اللغة', 'هل الاستنتاجات مبررة؟ وهل «المؤشر» مكتوب كمؤشر؟');
  const region = summary.ytd.filter((r) => r.org === 'منطقة الرياض');
  const memoIn = { period: summary.period, region, notes: notes.map((n) => ({ dept: n.dept, text: n.text })), checks_failed: failed.map((c) => c.check) };
  let memo = { headline: label, summary: '', notes: notes.map((n) => ({ dept: n.dept, text: n.text })) };
  try {
    const out = await llm.review(
      `أنت المدقق Aziz ومحرر التقرير. عندك ملاحظات الأقسام المعتمدة وأرقام المنطقة. سوِّ ثلاثة أشياء:\n` +
      `1) مانشيت قصير (سطر واحد).\n2) ملخص من 3 إلى 4 جمل للمنطقة كلها.\n3) حرّر لغة كل ملاحظة بدون تغيير أي رقم أو حكم.\n` +
      `<بيانات>${JSON.stringify(memoIn)}</بيانات>\n<قرار>\nرد بـ JSON فقط: {"headline":"...","summary":"...","notes":[{"dept":"...","text":"..."}]}`,
      WRITING_RULES);
    const j = parseJSON(out);
    if (j && j.headline && Array.isArray(j.notes)) memo = j;
  } catch (e) {
    alert(store, 'aziz', 'تعذّر التحرير', e.message);
  }
  const allText = [memo.headline, memo.summary, ...memo.notes.map((n) => n.text)].join('\n');
  const lock = await engine.numcheck(summaryPath, allText);
  if (!lock.ok) {
    // قفل الأرقام: إذا التحرير غيّر رقم، نرجع لنصوص القادة كما هي
    say(store, 'aziz', 'faisal', 'التحرير غيّر أرقام (' + lock.unknown.join('، ') + ')، فرجعت لنصوص القادة المعتمدة كما هي.', { kind: 'return' });
    memo = { headline: label, summary: '', notes: notes.map((n) => ({ dept: n.dept, text: n.text })) };
  }
  done(store, 'aziz', 'اعتمد المنطق واللغة', lock.ok ? 'كل الأرقام مطابقة للمصدر' : 'رجّعت للنصوص الأصلية');
  say(store, 'aziz', 'neves', 'المذكرة جاهزة للطباعة.', { kind: 'approve' });

  // 4) Neves: PDF + نسخ المخرجات لمجلد المسودات في الدرايف
  store.setRun({ stage: 7 });
  work(store, 'neves', 'يطبع المذكرة', 'PDF عربي من اليمين');
  const html = memoHtml({ memo, summary, notes, failed, label });
  const tag = `${year}-${String(month).padStart(2, '0')}`;
  const htmlPath = path.join(runDir, `مذكرة_المتابعة_${tag}.html`);
  fs.writeFileSync(htmlPath, html);
  const pdfPath = path.join(runDir, `مذكرة_المتابعة_${tag}.pdf`);
  if (renderPdf) await renderPdf(html, pdfPath);
  const draftDir = S.driveRoot ? path.join(S.driveRoot, S.draftDir, tag) : path.join(runDir, 'draft');
  fs.mkdirSync(draftDir, { recursive: true });
  const files = [];
  for (const f of [result.xlsx, renderPdf ? pdfPath : htmlPath]) {
    const dst = uniquePath(path.join(draftDir, path.basename(f)));
    fs.copyFileSync(f, dst);
    files.push(dst);
  }
  done(store, 'neves', 'سلّم المسودة لـ Faisal', files.map((f) => path.basename(f)).join('، '));
  say(store, 'neves', 'faisal', `${label} جاهز للاعتماد: ${files.map((f) => path.basename(f)).join(' و')}.`, { kind: 'handoff', via: 'email' });

  store.setRun({ stage: 8, awaiting: 'approval', files, tag, label, runDir });
  store.setStatus('faisal', { kind: 'wait', text: 'عنده مسودة للاعتماد' });
  fs.writeFileSync(path.join(runDir, 'notes.json'), JSON.stringify({ notes, memo }, null, 1));
  return { files, summary, notes, memo, runDir };
}

async function approve(ctx, comment) {
  const { store, settings } = ctx;
  const S = settings.get(), run = store.run;
  if (run.awaiting !== 'approval' || !run.files) throw new Error('ما فيه مسودة تنتظر الاعتماد');
  const dir = S.driveRoot ? path.join(S.driveRoot, S.publishDir, run.tag) : path.join(run.runDir, 'published');
  fs.mkdirSync(dir, { recursive: true });
  const out = run.files.map((f) => { const d = uniquePath(path.join(dir, path.basename(f))); fs.copyFileSync(f, d); return d; });
  say(store, 'faisal', 'neves', 'اعتمدت. انشر.' + (comment ? ' ' + comment : ''), { kind: 'approve' });
  done(store, 'faisal', 'اعتمد ' + (run.label || 'التقرير'));
  done(store, 'neves', 'نشر المعتمد', out.map((f) => path.basename(f)).join('، '));
  store.setRun({ awaiting: null, stage: 0, label: 'بانتظار طلبك', published: out });
  return out;
}

function sendBack(ctx, comment) {
  const { store } = ctx;
  say(store, 'faisal', 'angel', 'رجّعت المسودة: ' + (comment || 'بدون ملاحظة'), { kind: 'return', via: 'email' });
  store.setStatus('angel', { kind: 'alert', text: 'استلم ترجيع Faisal', detail: comment || '' });
  store.setRun({ awaiting: null, stage: 0, label: 'رجعت المسودة: عدّل الطلب وأرسله مرة ثانية' });
}

function uniquePath(p) {
  if (!fs.existsSync(p)) return p;
  const ext = path.extname(p), base = p.slice(0, -ext.length);
  let i = 2; while (fs.existsSync(`${base}_${i}${ext}`)) i++;
  return `${base}_${i}${ext}`;
}

// ------------------------------------------------------------ memo html (Neves ما يعدّل المحتوى)
function memoHtml({ memo, summary, notes, failed, label }) {
  const e = (s) => String(s == null ? '' : s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const m = (v) => v == null ? '—' : (v / 1e6).toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  const n = (v) => `<span dir="ltr">${v}</span>`;
  const p = (v) => v == null ? '—' : `${v > 0 ? '+' : v < 0 ? '−' : ''}${Math.abs(v).toFixed(1)}٪`;
  const rows = summary.ytd.filter((r) => r.org === 'منطقة الرياض').map((r) =>
    `<tr><td>${e(r.tab)}</td><td>${n(m(r.billed))}</td><td>${n(m(r.actual))}</td><td>${n(m(r.prev_years))}</td><td>${n(m(r.target))}</td><td class="${(r.dev_target_pct || 0) < 0 ? 'neg' : 'pos'}">${n(p(r.dev_target_pct))}</td><td>${n(r.collection_pct == null ? '—' : r.collection_pct.toFixed(1) + '٪')}</td></tr>`).join('');
  const deptNotes = (memo.notes || []).map((n) => {
    const src = notes.find((x) => x.dept === n.dept) || {};
    return `<section><h3>${e(n.dept)}</h3><p>${e(n.text)}</p>${src.numbers_ok === false ? '<p class="flag">تنبيه: فيه أرقام ما طابقت المصدر، راجعها قبل الاعتماد.</p>' : ''}</section>`;
  }).join('');
  const flags = failed.length ? `<h2>ملاحظات التدقيق</h2><ul>${failed.map((c) => `<li>${e(c.check)}: ${e(c.note)}</li>`).join('')}</ul>` : '';
  return `<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><title>${e(label)}</title>
<style>
@page{size:A4;margin:14mm 13mm}
body{font-family:"IBM Plex Sans Arabic","Segoe UI",Tahoma,sans-serif;color:#1c1c1c;font-size:10.5pt;line-height:1.7;margin:0}
.mast{border-bottom:3px double #1c1c1c;padding-bottom:6px;display:flex;justify-content:space-between;align-items:end}
.mast h1{font-size:22pt;margin:0}.mast small{color:#555}
.badge{display:inline-block;border:1px solid #b35;color:#b35;padding:0 8px;border-radius:4px;font-size:9pt;margin-top:6px}
h2{font-size:14pt;border-bottom:1px solid #1c1c1c;padding-bottom:2px;margin:16px 0 6px}
h3{font-size:11.5pt;margin:10px 0 2px}
.head{font-size:17pt;font-weight:700;margin:10px 0 6px;line-height:1.4}
table{width:100%;border-collapse:collapse;font-variant-numeric:tabular-nums;font-size:9.5pt}
th,td{border-bottom:1px solid #ddd;padding:4px 6px;text-align:right}th{background:#f3f1ec}
.neg{color:#b3261e}.pos{color:#1d6b3a}.flag{color:#b3261e;font-size:9pt}
footer{margin-top:18px;border-top:1px solid #ccc;padding-top:4px;font-size:8.5pt;color:#666}
</style></head><body>
<div class="mast"><div><h1>مذكرة المتابعة الشهرية</h1><small>إدارة تخطيط ومراقبة الإيرادات · أمانة منطقة الرياض</small></div><small>${e(label)}</small></div>
<span class="badge">مسودة: بانتظار اعتماد Faisal</span>
<div class="head">${e(memo.headline || label)}</div>
${memo.summary ? `<p>${e(memo.summary)}</p>` : ''}
<h2>منطقة الرياض حتى تاريخه (مليون ريال)</h2>
<table><thead><tr><th>التبويب</th><th>مفوتر</th><th>فعلي</th><th>محصل سنوات سابقة</th><th>مستهدف</th><th>الانحراف عن المستهدف</th><th>نسبة التحصيل</th></tr></thead><tbody>${rows}</tbody></table>
<h2>ملاحظات الأقسام</h2>${deptNotes}
${flags}
<footer>الأرقام محسوبة آلياً من ملفات الفواتير ومدققة بفحوص المطابقة، والنصوص كتبها الفريق وراجعها قادة الأقسام. المستهدف من ملف المتابعة السابق.</footer>
</body></html>`;
}

module.exports = { parseRequest, runInspect, runMonthly, approve, sendBack, DEPTS, MONTHS, memoHtml };
