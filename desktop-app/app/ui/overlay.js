// شريط الطلبات، الاعتماد، والإعدادات.
(function () {
  const R = window.revenue;
  if (!R) return;
  const MONTHS = ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو', 'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'];
  const el = (tag, attrs = {}, ...kids) => {
    const n = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) { if (k === 'on') Object.entries(v).forEach(([e, f]) => n.addEventListener(e, f)); else if (v !== undefined && v !== null) n.setAttribute(k, v); }
    kids.flat().forEach((k) => n.append(k instanceof Node ? k : document.createTextNode(String(k))));
    return n;
  };
  const prev = new Date(); prev.setDate(1); prev.setMonth(prev.getMonth() - 1);
  const prevLabel = `حدّث ملف متابعة ${MONTHS[prev.getMonth()]} ${prev.getFullYear()}`;

  document.body.classList.add('rv');
  const input = el('input', { id: 'rvIn', placeholder: 'اكتب طلبك للفريق… مثل: ' + prevLabel, autocomplete: 'off', 'aria-label': 'طلبك للفريق' });
  const send = el('button', { class: 'pri', type: 'submit' }, 'إرسال');
  const state = el('span', { id: 'rvState', role: 'status' }, '');
  const bar = el('div', { id: 'rvBar' },
    el('form', { on: { submit: (e) => { e.preventDefault(); submit(input.value); } } }, input, send),
    el('div', { id: 'rvChips' },
      el('button', { type: 'button', on: { click: () => submit(prevLabel) } }, 'متابعة ' + MONTHS[prev.getMonth()]),
      el('button', { type: 'button', on: { click: () => submit('افحص ملفات الدرايف') } }, 'فحص الدرايف')),
    state,
    el('button', { type: 'button', on: { click: openSettings } }, 'الإعدادات'));
  document.body.prepend(bar);

  async function submit(text) {
    text = String(text || '').trim();
    if (!text) return;
    send.disabled = true;
    const r = await R.request(text);
    send.disabled = false;
    if (r.ok) { input.value = ''; state.textContent = 'الفريق استلم الطلب'; }
    else if (r.error) { state.textContent = r.error; if (r.openSettings) openSettings(); }
  }

  // ---- الاعتماد
  let approveBox = null;
  function renderRun(run) {
    if (!run) return;
    state.textContent = run.label || '';
    if (run.awaiting === 'approval' && run.files) showApprove(run);
    else if (approveBox) { approveBox.remove(); approveBox = null; }
  }
  function showApprove(run) {
    if (approveBox) approveBox.remove();
    const note = el('input', { placeholder: 'ملاحظة (اختيارية)' });
    approveBox = el('section', { id: 'rvApprove', 'aria-label': 'اعتماد' },
      el('h3', {}, (run.label || 'المسودة') + ': جاهزة لاعتمادك'),
      el('ul', {}, run.files.map((f) => el('li', {}, el('button', { type: 'button', dir: 'ltr', on: { click: () => R.openPath(f) } }, f.split(/[\\/]/).pop())))),
      el('div', { class: 'row' }, note,
        el('button', { class: 'pri', type: 'button', on: { click: async () => { const r = await R.approve(note.value); state.textContent = r.ok ? 'اعتُمد ونُشر في الدرايف' : r.error; } } }, 'اعتماد ونشر'),
        el('button', { type: 'button', on: { click: () => R.sendBack(note.value) } }, 'ترجيع')));
    document.body.append(approveBox);
  }
  R.snapshot().then((s) => renderRun(s.run));
  window.addEventListener('revenue:change', (e) => { if (e.detail.type === 'run') renderRun(e.detail.data); });

  // ---- الإعدادات
  async function openSettings() {
    const s = await R.getSettings();
    const f = {};
    const inp = (key, val, cls) => (f[key] = el('input', { value: val || '', class: cls || '' }));
    const browse = (key, base) => el('button', { type: 'button', on: { click: async () => {
      const v = key === 'driveRoot' ? await R.pickFolder() : await R.pickFile(f.driveRoot.value || base);
      if (v) f[key].value = v; } } }, 'اختيار');
    const field = (label, key, val, pick, cls) => el('label', {}, label, el('div', { class: 'row' }, inp(key, val, cls), pick ? browse(key) : ''));
    const visit = el('select', {}, el('option', { value: 'ollama' }, 'Ollama (مجاني على جهازك)'), el('option', { value: 'claude' }, 'Claude'));
    visit.value = s.visitModel;
    const claudeOn = el('select', {}, el('option', { value: 'yes' }, 'مفعّل (باشتراكك عن طريق Claude Code)'), el('option', { value: 'no' }, 'متوقف (Ollama يسوي كل شي)'));
    claudeOn.value = s.claude.enabled ? 'yes' : 'no';
    const adv = el('textarea', { spellcheck: 'false' }); adv.value = JSON.stringify(s.engine, null, 2);
    const out = el('div', { id: 'rvCheck' });
    const M = el('div', { class: 'rvM', role: 'dialog', 'aria-modal': 'true', 'aria-label': 'الإعدادات' }, el('div', { class: 'box' },
      el('h2', {}, 'الإعدادات'),
      el('p', { class: 'hint' }, 'كل الإعدادات تنحفظ على جهازك فقط. البرنامج يقرأ من الدرايف وما يعدّل أي ملف أصلي؛ المسودات تروح لمجلد الطباعة، والمعتمد لمجلد النشر.'),
      el('fieldset', {}, el('legend', {}, 'Google Drive'),
        field('مجلد الدرايف المشترك', 'driveRoot', s.driveRoot, true),
        field('المفوتر (السنة الحالية)', 'amana_billed', s.sources.amana_billed, true),
        field('المسدد عن سنوات سابقة', 'amana_paid_prev', s.sources.amana_paid_prev, true),
        field('ملف الإعدادات (الحسابات)', 'settings', s.sources.settings, true),
        field('ملف المتابعة السابق (القالب)', 'template', s.sources.template, true),
        field('مجلد المسودات', 'draftDir', s.draftDir),
        field('مجلد المعتمد', 'publishDir', s.publishDir)),
      el('fieldset', {}, el('legend', {}, 'النماذج'),
        field('عنوان Ollama', 'ollamaUrl', s.ollama.url, false, 'ltr'),
        field('نموذج Ollama', 'ollamaModel', s.ollama.model, false, 'ltr'),
        el('label', {}, 'Claude', claudeOn),
        field('نموذج Claude', 'claudeModel', s.claude.model, false, 'ltr'),
        field('أمر Claude Code', 'claudeCli', s.claude.cliPath, false, 'ltr'),
        el('label', {}, 'رد الزيارة التفقدية', visit)),
      el('fieldset', {}, el('legend', {}, 'متقدم'),
        field('مسار Python (اختياري)', 'pythonPath', s.pythonPath, false, 'ltr'),
        el('label', {}, 'أعمدة الملفات والبلديات', adv)),
      out,
      el('div', { class: 'foot' },
        el('button', { class: 'pri', type: 'button', on: { click: save } }, 'حفظ'),
        el('button', { type: 'button', on: { click: async () => { await save(true); check(); } } }, 'حفظ وفحص الاتصال'),
        el('button', { type: 'button', on: { click: () => M.remove() } }, 'إغلاق'))));
    document.body.append(M);
    async function save(keepOpen) {
      let engine;
      try { engine = JSON.parse(adv.value); } catch (e) { out.textContent = 'خطأ في الإعدادات المتقدمة: ' + e.message; return; }
      await R.saveSettings({ ...s, driveRoot: f.driveRoot.value.trim(), draftDir: f.draftDir.value.trim(), publishDir: f.publishDir.value.trim(),
        sources: { amana_billed: f.amana_billed.value.trim(), amana_paid_prev: f.amana_paid_prev.value.trim(), settings: f.settings.value.trim(), template: f.template.value.trim() },
        ollama: { url: f.ollamaUrl.value.trim(), model: f.ollamaModel.value.trim() },
        claude: { enabled: claudeOn.value === 'yes', model: f.claudeModel.value.trim(), cliPath: f.claudeCli.value.trim() || 'claude' },
        visitModel: visit.value, pythonPath: f.pythonPath.value.trim(), engine });
      out.textContent = 'انحفظت.';
      if (keepOpen !== true) M.remove();
    }
    async function check() {
      out.textContent = 'أفحص Ollama وClaude والدرايف…';
      const r = await R.checkSettings();
      const L = [];
      L.push(r.ollama.ok ? `✓ Ollama شغّال. النماذج: ${r.ollama.models.join('، ') || 'ما فيه'}${r.ollama.hasModel ? '' : ' — النموذج المختار مو موجود، نزّله بـ ollama pull'}` : '✗ Ollama: ' + r.ollama.error);
      if (r.claude) L.push(r.claude.ok ? '✓ Claude رد: ' + r.claude.reply : '✗ Claude: ' + r.claude.error);
      if (r.drive) L.push(r.drive.ok ? '✓ ملفات الدرايف كلها موجودة' : '✗ الدرايف: ' + (r.drive.error || r.drive.items.filter((i) => !i.exists || i.error).map((i) => i.source + (i.error ? ' (' + i.error + ')' : ' ما انلقى')).join('، ')));
      out.textContent = L.join('\n');
    }
  }
  // زر «اطلب تقرير» في المبنى كان للمحاكاة؛ في البرنامج يوديك لشريط الطلبات
  document.addEventListener('click', (e) => {
    if (e.target.closest && e.target.closest('#btnReport')) { e.preventDefault(); e.stopPropagation(); input.focus(); }
  }, true);
  R.getSettings().then((s) => { if (!s.driveRoot) openSettings(); });
})();
