import re
from collections import Counter
from openpyxl import load_workbook
w=load_workbook('tmp/artifact/revised.xlsx',read_only=True)
bad=[];mono=[]
v='аоуыэяёюиеАOУЫЭЯЁЮИЕ'
for sh in w:
 if sh.title=='说明':continue
 for row in sh.iter_rows(min_row=4):
  for cell in row:
   x=str(cell.value or '')
   if not x or 'Урок ' in x:continue
   for i,c in enumerate(x):
    if c!='\u0301':continue
    prev=x[i-1] if i else ''
    if prev not in v:bad.append((sh.title,cell.coordinate,x,prev))
print('bad_marks',len(bad))
for x in bad[:100]:print(x)
