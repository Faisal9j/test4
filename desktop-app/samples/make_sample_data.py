"""يولّد بيانات وهمية بالكامل (للتجربة فقط): ما فيها أي بيانات حقيقية."""
import os, random, datetime as dt, shutil, json, openpyxl
R = os.path.dirname(os.path.abspath(__file__)); OUT = os.path.join(R, 'drive')
random.seed(7)
os.makedirs(f'{OUT}/الفواتير/الأمانة', exist_ok=True)
TABS = {'14155': 'إيرادات إستثمارية', '142162': 'إيرادات إستثمارية', '11452104': 'الرسوم والتراخيص', '11452103': 'الرسوم والتراخيص',
        '1431': 'جزاءات وغرامات', '1421911': 'التبغ والإيواء', '11442': 'التبغ والإيواء', '1499': 'أخرى'}
MUNIS = ['قطاع الامانة - الجنوب', 'قطاع الامانة - الشرق', 'قطاع الامانة - الشمال', 'قطاع الامانة - الغرب', 'قطاع الامانة - الوسط',
         'بلدية الخرج', 'بلدية الدرعية', 'بلدية المجمعة', 'بلدية الزلفي', 'بلدية عفيف']
wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'الإعدادات'; ws.append(['رقم الحساب', 'اسم الحساب', 'التبويب'])
for a, t in TABS.items(): ws.append([int(a), 'حساب ' + a, t])
wb.save(f'{OUT}/الإعدادات.xlsx')
H = ['رقم الفاتورة', 'اسم المستفيد', 'رقم الهوية', 'رقم الحساب', 'البلدية', 'تاريخ الفاتورة', 'مبلغ الفاتورة', 'المبلغ المسدد', 'تاريخ السداد', 'حالة الفاتورة']
wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Sheet1'; ws.append(H)
accs = list(TABS)
for i in range(6000):
    d = dt.date(2026, 1, 1) + dt.timedelta(days=random.randint(0, 300))
    amt = round(random.uniform(500, 200000), 2); paid = amt if random.random() < .7 else 0
    pd = d + dt.timedelta(days=random.randint(0, 40)) if paid else None
    st = 'ملغاة' if random.random() < .03 else 'نشطة'
    ws.append([f'INV{i:06d}', 'مستفيد وهمي', '1000000000', int(random.choice(accs) if random.random() < .995 else '9999'), random.choice(MUNIS), d, amt, paid, pd, st])
    if random.random() < .01: ws.append([f'INV{i:06d}x', 'مستفيد وهمي', '1000000000', int(accs[0]), MUNIS[0], d, amt, 0, None, 'نشطة'])
wb.save(f'{OUT}/الفواتير/الأمانة/مفوتر 2026 بدون ملغي.xlsx')
wb = openpyxl.Workbook(); ws = wb.active; ws.append(H)
for i in range(1500):
    d = dt.date(2026, 1, 1) + dt.timedelta(days=random.randint(0, 300)); amt = round(random.uniform(500, 90000), 2)
    ws.append([f'OLD{i:06d}', 'مستفيد وهمي', '1000000000', int(random.choice(list(TABS))), random.choice(MUNIS), dt.date(2024, 5, 1), amt, amt, d, 'مسددة'])
wb.save(f'{OUT}/الفواتير/الأمانة/مسدد 2026 عن سنوات سابقة.xlsx')
os.makedirs(f'{OUT}/المتابعة', exist_ok=True)
shutil.copy(os.path.join(R, '..', '..', 'data', 'revenue-source.xlsx'), f'{OUT}/المتابعة/ملف المتابعة السابق.xlsx')
print('ok', OUT)
