// ينسخ المبنى من revenue-team/building.html للبرنامج: three.js محلي + الجسر + شريط الطلبات.
const fs = require('fs');
const path = require('path');
const root = path.resolve(__dirname, '..');
const src = path.resolve(root, '..', 'revenue-team', 'building.html');
const out = path.join(root, 'app', 'renderer');
fs.mkdirSync(path.join(out, 'vendor'), { recursive: true });
const three = path.join(root, 'node_modules', 'three');
fs.copyFileSync(path.join(three, 'build', 'three.min.js'), path.join(out, 'vendor', 'three.min.js'));
fs.copyFileSync(path.join(three, 'examples', 'js', 'controls', 'OrbitControls.js'), path.join(out, 'vendor', 'OrbitControls.js'));
for (const f of ['shim.js', 'overlay.js', 'overlay.css']) fs.copyFileSync(path.join(root, 'app', 'ui', f), path.join(out, f));
let html = fs.readFileSync(src, 'utf8');
const n0 = html.length;
html = html
  .replace('https://cdn.jsdelivr.net/npm/three@0.147.0/build/three.min.js', 'vendor/three.min.js')
  .replace('https://cdn.jsdelivr.net/npm/three@0.147.0/examples/js/controls/OrbitControls.js', 'vendor/OrbitControls.js')
  ;
// المبنى محفوظ كجزء صفحة (بدون html/head/body)، فنلفّه هنا
html = `<!doctype html>
<html lang="ar" dir="rtl"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<script src="shim.js"></script>
</head><body>
${html}
<link rel="stylesheet" href="overlay.css">
<script src="overlay.js"></script>
</body></html>
`;
if (html.includes('cdn.jsdelivr.net/npm/three')) throw new Error('ما انربط three.js محلياً');
if (!html.includes('shim.js') || !html.includes('overlay.js')) throw new Error('ما انحقن الجسر أو الشريط');
fs.writeFileSync(path.join(out, 'building.html'), html);
console.log('building synced', n0, '→', html.length);
