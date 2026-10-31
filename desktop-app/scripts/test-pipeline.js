// يشغّل السلسلة كاملة على البيانات الوهمية بنموذج وهمي (بدون Electron ولا Ollama ولا Claude).
process.env.REVENUE_MOCK_LLM = '1';
const fs = require('fs'), os = require('os'), path = require('path'), assert = require('assert');
const { Store } = require('../app/store');
const { Settings } = require('../app/settings');
const { Engine } = require('../app/engine');
const { makeLLM } = require('../app/llm');
const P = require('../app/pipeline');

(async () => {
  const root = path.resolve(__dirname, '..');
  const samples = path.join(root, 'samples', 'drive');
  if (!fs.existsSync(samples)) require('child_process').execFileSync('python3', [path.join(root, 'samples', 'make_sample_data.py')]);
  const work = fs.mkdtempSync(path.join(os.tmpdir(), 'rev-test-'));
  const drive = path.join(work, 'drive');
  fs.cpSync(samples, drive, { recursive: true });
  const settings = new Settings(path.join(work, 'settings.json'));
  const sample = JSON.parse(fs.readFileSync(path.join(root, 'samples', 'config.sample.json'), 'utf8'));
  settings.save({ ...settings.get(), driveRoot: drive, sources: sample.sources, engine: { ...settings.get().engine, municipality_groups: sample.municipality_groups } });
  const store = new Store(null);
  const ctx = { store, settings, engine: new Engine({ resourcesDir: root }), llm: makeLLM(settings.get()), workDir: work };

  assert.deepStrictEqual(P.parseRequest('حدّث ملف متابعة أكتوبر'), { kind: 'monthly', year: new Date().getFullYear(), month: 10 });
  assert.strictEqual(P.parseRequest('افحص ملفات الدرايف').kind, 'inspect');

  const before = fs.readdirSync(path.join(drive, 'الفواتير', 'الأمانة')).map((f) => fs.statSync(path.join(drive, 'الفواتير', 'الأمانة', f)).mtimeMs);
  const ins = await P.runInspect(ctx);
  assert.ok(ins.items.every((i) => i.exists), 'كل المصادر موجودة');
  const r = await P.runMonthly(ctx, { year: 2026, month: 10 });
  assert.strictEqual(r.files.length, 2);
  r.files.forEach((f) => assert.ok(fs.existsSync(f) && f.includes('06 طباعة'), f));
  assert.ok(r.notes.length === 5 && r.notes.every((n) => n.numbers_ok), 'ملاحظات الأقسام وأرقامها مطابقة');
  const pub = await P.approve(ctx);
  pub.forEach((f) => assert.ok(fs.existsSync(f) && f.includes('07 معتمد ومنشور'), f));
  const after = fs.readdirSync(path.join(drive, 'الفواتير', 'الأمانة')).map((f) => fs.statSync(path.join(drive, 'الفواتير', 'الأمانة', f)).mtimeMs);
  assert.deepStrictEqual(before, after, 'الملفات الأصلية ما تغيّرت');
  const ids = Object.keys(store.status);
  console.log('statuses:', ids.join(' '));
  console.log('chat:', store.chat.length, 'رسالة');
  console.log(store.chat.slice(-6).map((m) => `${m.from}→${m.to}: ${m.text}`).join('\n'));
  console.log('OK', work);
})().catch((e) => { console.error('FAIL', e); process.exit(1); });
