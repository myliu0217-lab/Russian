import json,re
from pathlib import Path
import openpyxl

root=Path(__file__).parents[1]
w=openpyxl.load_workbook(root/'大学俄语1_逐课生词分类与变位_修订版.xlsx',rich_text=True)
s=w['动词']
plain=lambda x: str(x or '').replace('\u0301','')
special={'есть':'特殊变位','быть':'特殊变位','сдать':'特殊变位',
         'хотеть':'混合变位','бежать':'混合变位'}
rows=[]; counts={}; unresolved=[]
for r in range(4,s.max_row+1):
    infinitive=plain(s.cell(r,1).value)
    if infinitive.startswith('Урок'):
        rows.append(None)
        continue
    word=infinitive.split()[0]
    ty=plain(s.cell(r,7).value).split()[0]
    oni=plain(s.cell(r,11).value).split()[0]
    if word in special:
        label=special[word]
    elif re.search(r'ишь(ся|сь)?$',ty) and re.search(r'(ат|ят)(ся|сь)?$',oni):
        label='第二变位法'
    elif re.search(r'(ешь|ёшь)(ся|сь)?$',ty) and re.search(r'(ут|ют)(ся|сь)?$',oni):
        label='第一变位法'
    else:
        label='待核查'
        unresolved.append((r,word,ty,oni))
    rows.append(label)
    counts[label]=counts.get(label,0)+1
(root/'.work/conjugation_labels.json').write_text(json.dumps(rows,ensure_ascii=False),encoding='utf8')
print('counts',counts)
print('unresolved',unresolved)
