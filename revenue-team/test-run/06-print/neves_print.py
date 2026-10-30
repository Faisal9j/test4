"""Neves: يحوّل المسودة المعتمدة لتقرير PDF بشكل صحيفة. ما يعدّل المحتوى."""
import json,base64,os,html,datetime,subprocess,sys
R=os.path.dirname(os.path.abspath(__file__));T=os.path.dirname(R);ROOT=os.path.dirname(os.path.dirname(T))
d=json.load(open(f'{T}/04-review/draft.json'))
logo=base64.b64encode(open(f'{ROOT}/revenue-targets-dashboard/logo.png','rb').read()).decode()
e=html.escape
def m(v):return f'{v/1e6:,.1f}'
def pct(v):return f'{v:+.2f}٪' if v is not None else '—'
t=d['totals']
rk=d['ranking'];mx=max(abs(x['deviation']) for x in rk)
bars=''.join(f'''<div class="bar"><span class="lb">{e(x['dept'])}</span><span class="tr"><i class="{'neg' if x['deviation']<0 else 'pos'}" style="width:{abs(x['deviation'])/mx*100:.1f}%"></i></span><b>{'+' if x['deviation']>=0 else '−'}{m(abs(x['deviation']))}</b></div>''' for x in rk)
rows=''
for dp in d['departments']:
    for i,r in enumerate(dp['rows']):
        rows+=f'''<tr class="{'first' if i==0 else ''}"><td>{e(dp['dept']) if i==0 else ''}</td><td>{e(r['org'])}</td><td>{m(r['actual'])}</td><td>{m(r['target'])}</td><td class="{'neg' if r['deviation']<0 else 'pos'}">{m(r['deviation'])}</td><td>{pct(r['deviation_pct'])}</td><td>{r['collection_pct']:.1f}٪</td><td>{pct(r['growth_pct'])}</td></tr>'''
finds=''.join(f'<li>{e(f)}</li>' for f in d['findings'])
flags=''.join(f'<li>{e(f)}</li>' for f in d['data_flags'])
issue=sys.argv[1] if len(sys.argv)>1 else 'تجريبي'
today=datetime.date.today().isoformat()
doc=f'''<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8">
<link href="https://fonts.googleapis.com/css2?family=Amiri:wght@400;700&family=IBM+Plex+Sans+Arabic:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
@page{{size:A4;margin:14mm 13mm}}
body{{font-family:"IBM Plex Sans Arabic",sans-serif;color:#1c1c1c;font-size:10.5pt;line-height:1.6;margin:0}}
.mast{{display:grid;grid-template-columns:auto 1fr auto;align-items:center;gap:14px;border-bottom:3px double #1c1c1c;padding-bottom:8px}}
.mast img{{height:62px}}
.mast h1{{font-family:Amiri,serif;font-size:30pt;margin:0;text-align:center;letter-spacing:.5px}}
.mast .meta{{font-size:8.5pt;color:#555;text-align:left;line-height:1.5}}
.strap{{display:flex;justify-content:space-between;font-size:8.5pt;color:#555;border-bottom:1px solid #1c1c1c;padding:3px 0;margin-bottom:10px}}
.head{{font-family:Amiri,serif;font-size:21pt;font-weight:700;line-height:1.35;margin:6px 0 10px}}
.kpis{{display:grid;grid-template-columns:repeat(4,1fr);border-top:1px solid #1c1c1c;border-bottom:1px solid #1c1c1c;margin-bottom:12px}}
.kpis div{{padding:6px 10px;border-left:1px solid #ccc}} .kpis div:last-child{{border-left:0}}
.kpis small{{display:block;color:#666;font-size:8.5pt}} .kpis b{{font-size:15pt;font-variant-numeric:tabular-nums}}
.cols{{display:grid;grid-template-columns:1.25fr 1fr;gap:16px}}
h2{{font-family:Amiri,serif;font-size:14pt;margin:4px 0 6px;border-bottom:1px solid #1c1c1c}}
ol,ul{{margin:0;padding-right:18px}} li{{margin-bottom:5px}}
.bar{{display:grid;grid-template-columns:110px 1fr 70px;align-items:center;gap:6px;margin:5px 0;font-size:9pt}}
.bar .tr{{background:#eee;height:12px}} .bar i{{display:block;height:12px}} .bar i.pos{{background:#1f6f5c}} .bar i.neg{{background:#b23b3b}}
.bar b{{text-align:left;font-variant-numeric:tabular-nums}}
.note{{font-size:8pt;color:#666}}
.flags{{border:1.5px solid #b23b3b;padding:8px 10px;margin-top:12px;background:#fff7f6}}
.flags h2{{color:#b23b3b;border-color:#b23b3b}}
table{{width:100%;border-collapse:collapse;font-size:8.6pt;font-variant-numeric:tabular-nums;margin-top:6px}}
th{{background:#1c1c1c;color:#fff;font-weight:600;padding:4px 5px}} td{{padding:3px 5px;border-bottom:1px solid #ddd;text-align:center}}
td:first-child{{text-align:right;font-weight:600}} tr.first td{{border-top:1px solid #888}}
.pos{{color:#1f6f5c}} .neg{{color:#b23b3b}}
.chain{{margin-top:10px;font-size:8.5pt;color:#444;border-top:1px solid #1c1c1c;padding-top:6px}}
.pb{{page-break-before:always}}
</style></head><body>
<div class="mast"><img src="data:image/png;base64,{logo}"><h1>تقرير الإيرادات</h1><div class="meta">العدد: {e(issue)}<br>{today}<br>مسودة بانتظار اعتماد رئيس التحرير</div></div>
<div class="strap"><span>فريق تخطيط ومراقبة الإيرادات</span><span>المصدر: {e(d['source'])}</span><span>المبالغ بالمليون ريال</span></div>
<div class="head">{e(d['headline'])}</div>
<div class="kpis">
<div><small>الفعلي حتى تاريخه</small><b>{m(t['actual'])}</b></div>
<div><small>المستهدف</small><b>{m(t['target'])}</b></div>
<div><small>الانحراف</small><b class="{'pos' if t['deviation']>=0 else 'neg'}">{m(t['deviation'])} ({pct(t['deviation_pct'])})</b></div>
<div><small>التحصيل من المفوتر · النمو</small><b>{t['collection_pct']:.1f}٪ · {pct(t['growth_pct'])}</b></div></div>
<div class="cols">
<section><h2>أبرز النتائج</h2><ol>{finds}</ol></section>
<section><h2>الانحراف حسب نوع الإيراد</h2>{bars}<p class="note">الأخضر فوق المستهدف، والأحمر دونه.</p>
<div class="flags"><h2>قرارات مطلوبة منك</h2><ul>{flags}</ul></div></section></div>
<div class="pb"></div>
<h2>تفصيل الأقسام حسب الجهة</h2>
<table><thead><tr><th>القسم</th><th>الجهة</th><th>الفعلي</th><th>المستهدف</th><th>الانحراف</th><th>نسبته</th><th>التحصيل</th><th>النمو</th></tr></thead><tbody>{rows}</tbody></table>
<div class="chain"><b>سلسلة الاعتماد:</b> Angel (فحص الملف: 30 مطابقة من 30) ← قادة الأقسام ({e(d['approvals']['leaders'])}) ← المحللون ← Vandaik ({e(d['approvals']['vandaik'])}) ← Salah (اعتمد بعد ترجيع واحد) ← Aziz ← Neves ← بانتظار Faisal.<br>
تعريفات: الانحراف = الفعلي − المستهدف · نسبته = الانحراف ÷ المستهدف · جوهري إذا تجاوز ±5٪ · "مؤشر" = تفسير بدون دليل مباشر.</div>
</body></html>'''
open(f'{R}/report.html','w').write(doc)
js=f"""const {{chromium}}=require('playwright');(async()=>{{const b=await chromium.launch({{proxy:{{server:process.env.HTTPS_PROXY}}}});const p=await b.newPage({{ignoreHTTPSErrors:true}});
await p.goto('file://{R}/report.html');await p.waitForTimeout(2500);await p.pdf({{path:'{R}/تقرير_الإيرادات_تجريبي.pdf',format:'A4',printBackground:true}});await b.close();}})();"""
open(f'{R}/_pdf.js','w').write(js);subprocess.run(['node',f'{R}/_pdf.js'],check=True);os.remove(f'{R}/_pdf.js');print('pdf ok')
