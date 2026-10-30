"""قائد القسم: يقارن نتيجتي العضوين على الأرقام الرئيسية (فرق 0.5٪ أو أقل = تطابق)."""
import json,glob,os,sys
D=os.path.dirname(os.path.abspath(__file__))
KEYS=['actual','target','deviation','deviation_pct','growth_pct','collection_pct']
def close(a,b,key):
    if a is None or b is None: return False
    if key.endswith('_pct'): return abs(a-b)<=0.05          # نقاط مئوية
    return abs(a-b)<=max(1,abs(b)*0.005)
out={}
for dept in ['lic','inv','pen','tob','oth']:
    files=sorted(glob.glob(f'{D}/{dept}-*.json'))
    res=[json.load(open(f)) for f in files]
    if len(res)<2: out[dept]={'status':'ناقص','members':len(res)};continue
    a,b=res[0],res[1];mism=[]
    for ra in a['rows']:
        rb=next((r for r in b['rows'] if r['org']==ra['org']),None)
        for k in KEYS:
            if rb is None or not close(ra.get(k),rb.get(k),k): mism.append({'org':ra['org'],'key':k,'m1':ra.get(k),'m2':rb and rb.get(k)})
    out[dept]={'dept':a['dept'],'type':a['type'],'status':'متطابق' if not mism else 'مختلف','mismatches':mism,'approved':a if not mism else None,'notes':{'1':a.get('notes'),'2':b.get('notes')}}
json.dump(out,open(f'{D}/leaders.json','w'),ensure_ascii=False,indent=1)
for k,v in out.items(): print(k,v.get('status'),len(v.get('mismatches',[])))
