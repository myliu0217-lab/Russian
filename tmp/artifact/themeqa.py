import re,unicodedata
from openpyxl import load_workbook
from pathlib import Path
def norm(s):
 s=unicodedata.normalize('NFD',s.lower()).replace('\u0301','').replace('ё','е')
 return re.sub('[^а-я]','',s)
w=load_workbook('tmp/artifact/revised.xlsx',read_only=True)
words=set()
for sh in w:
 if sh.title in ('说明','名词变格'): continue
 for row in sh.iter_rows(min_row=4,max_col=1,values_only=True):
  if row[0] and not str(row[0]).startswith('Урок '):words.add(norm(str(row[0])))
text=Path('俄语1review.md').read_text(encoding='utf8').split('## 主题生词分类',1)[1]
for m in re.finditer(r'### （(\d)）([^\n]+)\n(.*?)(?=### （|\Z)',text,re.S):
 title=m.group(2);body=m.group(3)
 print('\n',title)
 for line in body.splitlines():
  if line.startswith('|') and not line.startswith('|--') and '俄语' not in line:
   raw=line.split('|')[1]
  else:
   raw=re.search(r'<td>([^<>]+)</td><td>',line)
   if raw:raw=raw.group(1)
   else:continue
  for part in re.split(r'\s*/\s*',raw):
   part=part.replace(' †','')
   if re.search('[А-Яа-яЁё]',part) and norm(part) not in words:
    print(part)
