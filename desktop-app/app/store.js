// مخزن الحالة المحلي: نفس شكل قاعدة بيانات المبنى (status / chat / run)، ويُحفظ على الجهاز فقط.
const fs = require('fs');

class Store {
  constructor(file) {
    this.file = file;
    this.status = {};
    this.chat = [];
    this.run = { active: true, stage: 0, label: 'بانتظار طلبك' };
    this.listeners = new Set();
    try {
      const d = JSON.parse(fs.readFileSync(file, 'utf8'));
      this.chat = (d.chat || []).slice(-400);
    } catch (_) {}
    this._t = null;
  }
  on(fn) { this.listeners.add(fn); return () => this.listeners.delete(fn); }
  _emit(change) { for (const fn of this.listeners) { try { fn(change); } catch (_) {} } this._save(); }
  _save() {
    if (!this.file) return;
    clearTimeout(this._t);
    this._t = setTimeout(() => {
      try { fs.writeFileSync(this.file, JSON.stringify({ chat: this.chat.slice(-400) })); } catch (_) {}
    }, 500);
  }
  snapshot() { return { status: this.status, chat: this.chat.slice(-400), run: this.run }; }
  setStatus(id, d) {
    const prev = this.status[id] || {};
    this.status[id] = { ...prev, ...d, at: d.at || Date.now() };
    this._emit({ type: 'status', id, data: this.status[id] });
  }
  addChat(m) {
    const msg = { t: Date.now(), kind: 'finding', ...m };
    this.chat.push(msg);
    if (this.chat.length > 800) this.chat = this.chat.slice(-400);
    this._emit({ type: 'chat', data: msg });
    return msg;
  }
  setRun(d) {
    this.run = { ...this.run, ...d };
    this._emit({ type: 'run', data: this.run });
  }
}
module.exports = { Store };
