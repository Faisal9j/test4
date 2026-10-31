// النماذج: Ollama محلياً للمسودات، وClaude (عن طريق Claude Code باشتراكك) للمراجعة والتحرير.
const { spawn } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const stripThink = (t) => String(t || '').replace(/<think>[\s\S]*?<\/think>/g, '').trim();

async function ollamaChat(cfg, messages, { timeoutMs = 600000 } = {}) {
  const ctl = new AbortController();
  const timer = setTimeout(() => ctl.abort(), timeoutMs);
  try {
    const r = await fetch(cfg.url.replace(/\/$/, '') + '/api/chat', {
      method: 'POST', signal: ctl.signal, headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ model: cfg.model, messages, stream: false, options: { temperature: 0.2, num_ctx: 16384 } })
    });
    if (!r.ok) throw new Error(`Ollama ${r.status}: ${(await r.text()).slice(0, 200)}`);
    const d = await r.json();
    return stripThink(d.message && d.message.content);
  } catch (e) {
    if (e.name === 'AbortError') throw new Error('Ollama ما رد في الوقت المحدد');
    if (/fetch failed|ECONNREFUSED/.test(String(e.message))) throw new Error('Ollama مو شغّال على الجهاز (افتح Ollama وجرّب مرة ثانية)');
    throw e;
  } finally { clearTimeout(timer); }
}

async function ollamaModels(cfg) {
  const r = await fetch(cfg.url.replace(/\/$/, '') + '/api/tags');
  const d = await r.json();
  return (d.models || []).map((m) => m.name);
}

// Claude Code بدون أدوات: نص يدخل ونص يطلع، وفي مجلد فاضي عشان ما يوصل لأي ملف.
function claudeAsk(cfg, prompt, { system, timeoutMs = 600000 } = {}) {
  return new Promise((resolve, reject) => {
    const cwd = fs.mkdtempSync(path.join(os.tmpdir(), 'rev-claude-'));
    const args = ['-p', '--output-format', 'json', '--model', cfg.model || 'opus', '--max-turns', '1',
      '--disallowedTools', 'Bash,Edit,Write,Read,Glob,Grep,WebFetch,WebSearch,NotebookEdit,Task'];
    const p = spawn(cfg.cliPath || 'claude', args, { cwd, shell: process.platform === 'win32', windowsHide: true });
    let out = '', err = '';
    const timer = setTimeout(() => { p.kill(); reject(new Error('Claude ما رد في الوقت المحدد')); }, timeoutMs);
    p.stdout.on('data', (d) => { out += d.toString('utf8'); });
    p.stderr.on('data', (d) => { err += d.toString('utf8'); });
    p.on('error', (e) => { clearTimeout(timer); reject(new Error('ما لقيت Claude Code على الجهاز: ' + e.message)); });
    p.on('close', (code) => {
      clearTimeout(timer);
      fs.rmSync(cwd, { recursive: true, force: true });
      try {
        const d = JSON.parse(out);
        if (d.is_error) return reject(new Error('Claude: ' + (d.result || d.subtype || 'خطأ')));
        resolve(String(d.result || '').trim());
      } catch (_) {
        reject(new Error(`Claude وقف (${code}): ${(err || out).slice(-300)}`));
      }
    });
    // التعليمات تنرسل مع النص عبر stdin (مو كوسيط في سطر الأوامر) عشان العربي والأسطر على Windows
    p.stdin.end(system ? `<تعليمات>\n${system}\n</تعليمات>\n\n${prompt}` : prompt);
  });
}

function makeLLM(settings) {
  const s = settings;
  if (s.mock || process.env.REVENUE_MOCK_LLM) return require('./mock-llm');
  return {
    async draft(messages) { return ollamaChat(s.ollama, messages); },
    async review(prompt, system) {
      if (!s.claude.enabled) return ollamaChat(s.ollama, [{ role: 'system', content: system || '' }, { role: 'user', content: prompt }]);
      return claudeAsk(s.claude, prompt, { system });
    },
    async chat(messages, which) {
      if ((which || s.visitModel) === 'claude' && s.claude.enabled) {
        const prompt = messages.map((m) => (m.role === 'user' ? 'Faisal: ' : 'أنت: ') + m.content).join('\n\n');
        return claudeAsk(s.claude, prompt);
      }
      return ollamaChat(s.ollama, messages);
    }
  };
}

async function checkAll(settings) {
  const out = {};
  try { out.ollama = { ok: true, models: await ollamaModels(settings.ollama) }; out.ollama.hasModel = out.ollama.models.includes(settings.ollama.model); }
  catch (e) { out.ollama = { ok: false, error: e.message }; }
  if (settings.claude.enabled) {
    try { out.claude = { ok: true, reply: await claudeAsk(settings.claude, 'رد بكلمة وحدة: جاهز', { timeoutMs: 120000 }) }; }
    catch (e) { out.claude = { ok: false, error: e.message }; }
  }
  return out;
}

module.exports = { makeLLM, checkAll, stripThink };
