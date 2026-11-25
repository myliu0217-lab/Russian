import re, unicodedata, zipfile
from collections import defaultdict
from openpyxl import load_workbook

path='tmp/artifact/revised.xlsx'
with zipfile.ZipFile(path) as z:
 assert z.testzip() is None
w=load_workbook(path,read_only=False)
assert len(w.sheetnames)==12

def norm(s):
 s=unicodedata.normalize('NFC',unicodedata.normalize('NFD',str(s or '').lower()).replace('\u0301','')).replace('ё','е')
 return re.sub('[^а-я]','',s)

by_sheet={}
for sh in w:
 if sh.title=='说明':continue
 lesson=None; groups=defaultdict(set)
 for row in sh.iter_rows(min_row=4):
  word=row[0].value
  if not word:continue
  m=re.search(r'Урок\s+(\d+)',str(word))
  if m:lesson=int(m.group(1));continue
  if lesson:groups[lesson].add(norm(word))
  for cell in row:
   text=str(cell.value or '')
   for i,ch in enumerate(text):
    if ch=='\u0301':assert i>0 and text[i-1] in 'аоуыэяёюиеАОУЫЭЯЁЮИЕ',(sh.title,cell.coordinate,text)
  if sh.title in ('名词','名词变格') and str(word).endswith('ь'):
   assert re.match(r'^\[(阳|阴)\]',str(row[1].value or '')),(sh.title,word)
 by_sheet[sh.title]=groups

assert all(x not in by_sheet['代词'].get(1,set()) for x in ('тот','та','то','те'))
assert 'мост' in {norm(c[0].value) for c in w['名词变格'].iter_rows(min_row=4) if c[0].value}
assert 'календарь' in by_sheet['名词'][8]
assert 'поэтому' in by_sheet['副词'][1]
assert 'сдать' in by_sheet['动词'][12]
assert 'лететь' in by_sheet['动词'][18]
assert 'молодой' in by_sheet['形容词'][10]
assert 'пойти' in by_sheet['动词'][12]
print('PASS',len(w.sheetnames),'sheets; textbook spot checks, accent placement and ь gender labels')
