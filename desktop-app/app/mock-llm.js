// نموذج وهمي للاختبار بدون Ollama ولا Claude: يكتب جمل من الأرقام اللي تجيه بس.
function firstRow(text) {
  const m = text.match(/\{[^{}]*"tab"[^{}]*\}/);
  try { return m ? JSON.parse(m[0]) : null; } catch (_) { return null; }
}
const M = (v) => (v / 1e6).toFixed(1);
module.exports = {
  async draft(messages) {
    const r = firstRow(messages.map((m) => m.content).join('\n'));
    if (!r) return 'ما وصلتني بيانات كافية للقسم.';
    return `بلغ المحصل الفعلي في ${r.org} لتبويب ${r.tab} ${M(r.actual)} مليون ريال، مقابل مفوتر ${M(r.billed)} مليون ريال، بنسبة تحصيل ${r.collection_pct}٪.`;
  },
  async review(prompt) {
    if (/مانشيت/.test(prompt)) {
      const b = prompt.match(/<بيانات>([\s\S]*?)<\/بيانات>/);
      const d = b ? JSON.parse(b[1]) : { notes: [] };
      return JSON.stringify({ headline: 'متابعة شهر ' + d.period.month_name + ' ' + d.period.year, summary: '', notes: d.notes });
    }
    const m = prompt.match(/<مسودة>([\s\S]*?)<\/مسودة>/);
    const note = m ? m[1].trim() : 'تمت المراجعة.';
    if (/<قرار>/.test(prompt)) return JSON.stringify({ decision: 'approve', note, reason: 'متسق مع البيانات' });
    return note;
  },
  async chat() { return 'تمام، استلمت توجيهك وبرفعه لـ Angel.'; }
};
