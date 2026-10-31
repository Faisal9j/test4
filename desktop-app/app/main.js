// البرنامج الرئيسي: نافذة المبنى + سلسلة العمل + الإعدادات.
const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Store } = require('./store');
const { Settings } = require('./settings');
const { Engine } = require('./engine');
const { makeLLM, checkAll } = require('./llm');
const P = require('./pipeline');

let win, store, settings, busy = false;
if (process.env.REVENUE_USER_DATA) app.setPath('userData', process.env.REVENUE_USER_DATA);
const resourcesDir = () => (app.isPackaged ? process.resourcesPath : path.join(__dirname, '..'));

function ctx() {
  const s = settings.get();
  return {
    store, settings, workDir: app.getPath('userData'),
    engine: new Engine({ resourcesDir: resourcesDir(), pythonPath: s.pythonPath }),
    llm: makeLLM(s), renderPdf
  };
}

async function renderPdf(html, outPath) {
  const tmp = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'rev-pdf-')), 'memo.html');
  fs.writeFileSync(tmp, html);
  const w = new BrowserWindow({ show: false, webPreferences: { sandbox: true, javascript: false } });
  try {
    await w.loadFile(tmp);
    const pdf = await w.webContents.printToPDF({ pageSize: 'A4', printBackground: true, preferCSSPageSize: true });
    fs.writeFileSync(outPath, pdf);
  } finally { w.destroy(); fs.rmSync(path.dirname(tmp), { recursive: true, force: true }); }
}

function allowedPath(p) {
  const s = settings.get();
  const roots = [app.getPath('userData'), s.driveRoot].filter(Boolean).map((r) => path.resolve(r) + path.sep);
  const abs = path.resolve(p);
  return roots.some((r) => abs.startsWith(r));
}

async function handleRequest(text) {
  const req = P.parseRequest(text);
  if (req.kind === 'unknown') {
    store.addChat({ from: 'faisal', to: 'angel', text, kind: 'handoff', via: 'email' });
    store.addChat({ from: 'angel', to: 'faisal', kind: 'return', text: 'هذا الطلب مو متاح في النسخة الأولى. المتاح الحين: «حدّث ملف متابعة <الشهر>» و«افحص ملفات الدرايف». باقي الطلبات (تقرير الرخص، الأثر والسيناريوهات) جاية في المراحل الجاية.' });
    return { ok: false, kind: 'unknown' };
  }
  if (busy) return { ok: false, error: 'الفريق مشغول بطلب ثاني، انتظره يخلص.' };
  const s = settings.get();
  if (!s.driveRoot || !fs.existsSync(s.driveRoot)) return { ok: false, error: 'حدد مجلد الدرايف من الإعدادات أول.', openSettings: true };
  busy = true;
  (async () => {
    try {
      if (req.kind === 'inspect') await P.runInspect(ctx());
      else await P.runMonthly(ctx(), req);
    } catch (e) {
      store.addChat({ from: 'angel', to: 'faisal', kind: 'return', text: 'وقف التشغيل: ' + e.message });
      store.setRun({ stage: 0, label: 'توقف التشغيل' });
    } finally { busy = false; }
  })();
  return { ok: true, kind: req.kind, ...req };
}

function createWindow() {
  win = new BrowserWindow({
    width: 1440, height: 900, minWidth: 980, minHeight: 640, backgroundColor: '#0a0f17', title: 'فريق الإيرادات',
    autoHideMenuBar: true,
    webPreferences: { preload: path.join(__dirname, 'preload.js'), contextIsolation: true, nodeIntegration: false, sandbox: true }
  });
  win.webContents.setWindowOpenHandler(({ url }) => { if (/^https:/.test(url)) shell.openExternal(url); return { action: 'deny' }; });
  win.webContents.on('will-navigate', (e) => e.preventDefault());
  win.loadFile(path.join(__dirname, 'renderer', 'building.html'));
}

app.whenReady().then(() => {
  const ud = app.getPath('userData');
  fs.mkdirSync(ud, { recursive: true });
  settings = new Settings(path.join(ud, 'settings.json'));
  store = new Store(path.join(ud, 'store.json'));
  ['angel', 'recv', 'cls', 'ent', 'vandaik', 'salah', 'aziz', 'neves'].forEach((id) => store.setStatus(id, { kind: 'idle', text: 'جاهز' }));
  store.setStatus('faisal', { kind: 'idle', text: 'في مكتبه' });
  store.on((c) => { if (win && !win.isDestroyed()) win.webContents.send('store:change', c); });

  ipcMain.handle('store:snapshot', () => store.snapshot());
  ipcMain.handle('store:addChat', (_e, m) => store.addChat(m));
  ipcMain.handle('request', (_e, text) => handleRequest(String(text || '').slice(0, 2000)));
  ipcMain.handle('approve', async (_e, c) => { try { return { ok: true, files: await P.approve(ctx(), c) }; } catch (e) { return { ok: false, error: e.message }; } });
  ipcMain.handle('sendBack', (_e, c) => { P.sendBack(ctx(), c); return { ok: true }; });
  ipcMain.handle('sample', async (_e, turns) => {
    try { return { text: await makeLLM(settings.get()).chat(turns) }; } catch (e) { return { error: e.message }; }
  });
  ipcMain.handle('settings:get', () => settings.get());
  ipcMain.handle('settings:save', (_e, s) => settings.save(s));
  ipcMain.handle('settings:check', async () => {
    const out = await checkAll(settings.get());
    try {
      const r = await P.runInspect(ctx());
      out.drive = { ok: r.items.every((i) => i.exists), items: r.items.map((i) => ({ source: i.source, exists: i.exists, error: i.error })) };
    } catch (e) { out.drive = { ok: false, error: e.message }; }
    return out;
  });
  ipcMain.handle('pick:folder', async () => {
    const r = await dialog.showOpenDialog(win, { properties: ['openDirectory'] });
    return r.canceled ? null : r.filePaths[0];
  });
  ipcMain.handle('pick:file', async (_e, base) => {
    const r = await dialog.showOpenDialog(win, { defaultPath: base || undefined, properties: ['openFile'], filters: [{ name: 'Excel / CSV', extensions: ['xlsx', 'xlsm', 'csv'] }] });
    if (r.canceled) return null;
    const f = r.filePaths[0];
    return base && path.resolve(f).startsWith(path.resolve(base) + path.sep) ? path.relative(base, f) : f;
  });
  ipcMain.handle('open:path', (_e, p) => (allowedPath(p) ? shell.openPath(p) : 'غير مسموح'));
  createWindow();
});
app.on('window-all-closed', () => app.quit());
