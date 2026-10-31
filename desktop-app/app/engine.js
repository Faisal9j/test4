// تشغيل محرك Python: يقرأ أحداثه سطر سطر ويمرّرها للمبنى.
const { spawn } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

function enginePaths(resourcesDir) {
  const exe = path.join(resourcesDir, 'engine', 'dist', process.platform === 'win32' ? 'revenue_engine.exe' : 'revenue_engine');
  const py = path.join(resourcesDir, 'engine', 'revenue_engine.py');
  return { exe, py };
}

class Engine {
  constructor({ resourcesDir, pythonPath }) {
    this.paths = enginePaths(resourcesDir);
    this.pythonPath = pythonPath;
  }
  _cmd(args) {
    if (fs.existsSync(this.paths.exe)) return [this.paths.exe, args];
    const py = this.pythonPath || (process.platform === 'win32' ? 'python' : 'python3');
    return [py, ['-X', 'utf8', this.paths.py, ...args]];
  }
  run(args, onEvent) {
    return new Promise((resolve, reject) => {
      const [cmd, argv] = this._cmd(args);
      const p = spawn(cmd, argv, { env: { ...process.env, PYTHONIOENCODING: 'utf-8' }, windowsHide: true });
      let buf = '', err = '', result = null, failure = null;
      p.stdout.on('data', (d) => {
        buf += d.toString('utf8');
        let i;
        while ((i = buf.indexOf('\n')) >= 0) {
          const line = buf.slice(0, i).trim(); buf = buf.slice(i + 1);
          if (!line) continue;
          let ev; try { ev = JSON.parse(line); } catch (_) { continue; }
          if (ev.ev === 'error') failure = ev.message;
          if (ev.ev === 'result' || ev.ev === 'numcheck') result = ev;
          if (onEvent) onEvent(ev);
        }
      });
      p.stderr.on('data', (d) => { err += d.toString('utf8'); });
      p.on('error', (e) => reject(new Error(`ما قدرت أشغّل Python (${cmd}): ${e.message}`)));
      p.on('close', (code) => {
        if (code === 0) resolve(result);
        else reject(new Error(failure || err.trim().split('\n').slice(-3).join(' ') || `المحرك وقف (${code})`));
      });
    });
  }
  writeConfig(cfg, dir) {
    const f = path.join(dir, 'engine-config.json');
    fs.writeFileSync(f, JSON.stringify(cfg, null, 1));
    return f;
  }
  async numcheck(summaryPath, text) {
    const f = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'rev-')), 'text.txt');
    fs.writeFileSync(f, text);
    const r = await this.run(['numcheck', '--summary', summaryPath, '--text', f]);
    fs.rmSync(path.dirname(f), { recursive: true, force: true });
    return r || { ok: true, unknown: [] };
  }
}
module.exports = { Engine };
