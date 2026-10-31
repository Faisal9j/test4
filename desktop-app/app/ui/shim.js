// يخلي المبنى يشتغل في البرنامج بدون تعديل: نفس واجهة window.claude.use('db' | 'sample') لكنها محلية.
(function () {
  const R = window.revenue;
  if (!R) return;
  let snap = null;
  const subs = { run: [], status: [], chat: [] };
  const ready = R.snapshot().then((s) => { snap = s; });
  const statusSnap = () => ({ docs: Object.entries(snap.status).map(([id, d]) => ({ id, data: () => d })) });
  R.onChange((c) => {
    if (!snap) return;
    if (c.type === 'status') { snap.status[c.id] = c.data; subs.status.forEach((f) => f(statusSnap())); }
    else if (c.type === 'chat') { snap.chat.push(c.data); subs.chat.forEach((f) => f({ docChanges: () => [{ type: 'added', doc: { data: () => c.data } }] })); }
    else if (c.type === 'run') { snap.run = c.data; subs.run.forEach((f) => f({ data: () => snap.run })); }
    window.dispatchEvent(new CustomEvent('revenue:change', { detail: c }));
  });
  const db = {
    doc() {
      return { onSnapshot(cb) { ready.then(() => { subs.run.push(cb); cb({ data: () => snap.run }); }); return () => {}; } };
    },
    collection(name) {
      const q = {
        orderBy() { return q; },
        limit() { return q; },
        onSnapshot(cb) {
          ready.then(() => {
            if (name === 'status') { subs.status.push(cb); cb(statusSnap()); }
            else if (name === 'chat') {
              subs.chat.push(cb);
              const init = snap.chat.slice(-400);
              cb({ docChanges: () => init.map((m) => ({ type: 'added', doc: { data: () => m } })) });
            }
          });
          return () => {};
        },
        doc() { return { set: (obj) => R.addChat(obj) }; }
      };
      return q;
    }
  };
  async function sample(turns, opts) {
    const r = await R.sample(turns);
    if (r.error) throw Object.assign(new Error(r.error), { code: 'local_error' });
    if (opts && opts.onText) opts.onText({ text: r.text });
    return { text: r.text };
  }
  window.claude = { use: async (n) => { await ready; return n === 'db' ? db : n === 'sample' ? sample : null; } };
})();
