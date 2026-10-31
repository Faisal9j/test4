// الإعدادات: تنحفظ في مجلد البرنامج على الجهاز (مو في المستودع).
const fs = require('fs');
const path = require('path');

const DEFAULTS = {
  driveRoot: '',
  sources: {
    amana_billed: 'الفواتير/الأمانة/مفوتر 2026 بدون ملغي.xlsx',
    amana_paid_prev: 'الفواتير/الأمانة/مسدد 2026 عن سنوات سابقة.xlsx',
    settings: 'الإعدادات.xlsx',
    template: 'المتابعة/ملف المتابعة.xlsx'
  },
  draftDir: '06 طباعة',
  publishDir: '07 معتمد ومنشور',
  pythonPath: '',
  ollama: { url: 'http://127.0.0.1:11434', model: 'qwen3:14b' },
  claude: { enabled: true, cliPath: 'claude', model: 'opus' },
  visitModel: 'ollama',
  mock: false,
  engine: {
    columns: {
      account: ['رقم الحساب', 'رمز الحساب', 'الحساب'],
      amount: ['مبلغ الفاتورة', 'المبلغ', 'قيمة الفاتورة'],
      paid: ['المبلغ المسدد', 'المسدد'],
      invoice: ['رقم الفاتورة'],
      status: ['حالة الفاتورة', 'الحالة'],
      issue_date: ['تاريخ الفاتورة', 'تاريخ الإصدار'],
      pay_date: ['تاريخ السداد'],
      municipality: ['البلدية', 'الجهة'],
      sector: ['القطاع']
    },
    pii_keywords: ['اسم', 'هوية', 'سجل', 'جوال', 'هاتف', 'بريد', 'مستفيد', 'عنوان', 'آيبان', 'IBAN'],
    city_markers: ['امانه', 'قطاع'],
    municipality_groups: {},
    tab_aliases: {}
  }
};

function merge(a, b) {
  if (Array.isArray(a) || typeof a !== 'object' || a === null) return b === undefined ? a : b;
  const out = { ...a };
  for (const k of Object.keys(b || {})) out[k] = k in a ? merge(a[k], b[k]) : b[k];
  return out;
}

class Settings {
  constructor(file) { this.file = file; this.data = merge(DEFAULTS, this._read()); }
  _read() { try { return JSON.parse(fs.readFileSync(this.file, 'utf8')); } catch (_) { return {}; } }
  get() { return this.data; }
  save(d) {
    this.data = merge(DEFAULTS, d);
    fs.mkdirSync(path.dirname(this.file), { recursive: true });
    fs.writeFileSync(this.file, JSON.stringify(this.data, null, 2));
    return this.data;
  }
  engineConfig() {
    const d = this.data;
    return { drive_root: d.driveRoot, sources: d.sources, ...d.engine };
  }
}
module.exports = { Settings, DEFAULTS };
