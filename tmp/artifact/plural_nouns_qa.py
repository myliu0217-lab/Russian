from openpyxl import load_workbook
from collections import Counter
import re, unicodedata, zipfile

old='大学俄语1_逐课生词分类与变位.xlsx'
new='tmp/artifact/plural_nouns_revised.xlsx'
with zipfile.ZipFile(new) as z:assert z.testzip() is None
a=load_workbook(old,read_only=False)
b=load_workbook(new,read_only=False)
assert a.sheetnames==b.sheetnames
for name in a.sheetnames:
 if name=='名词变格':continue
 x,y=a[name],b[name]
 assert x.max_row==y.max_row and x.max_column==y.max_column,name
 for row in x:
  for c in row:assert c.value==y[c.coordinate].value,(name,c.coordinate)
s=b['名词变格']
assert s.max_column==15
assert len(s.merged_cells.ranges)==len(a['名词变格'].merged_cells.ranges)
groups={};lesson=None
for row in s.iter_rows(min_row=4,values_only=True):
 if not row[0]:continue
 m=re.search('Урок (\\d+)',str(row[0]))
 if m:lesson=int(m.group(1));groups.setdefault(lesson,{});continue
 assert row[0] not in groups[lesson],(lesson,row[0])
 groups[lesson][row[0]]=row
 for value in row[3:]:
  for i,c in enumerate(str(value or '')):
   if c=='\u0301':assert i>0 and str(value)[i-1] in 'аоуыэяёюиеАОУЫЭЯЁЮИЕ',(row[0],value)
all_words={row[0] for rows in groups.values() for row in rows.values()}
assert 'и́мя' not in all_words and 'вре́мя' not in all_words
for l,word,singular,plural in [(5,'о́вощи','о́вощ','о́вощи'),(7,'ту́фли','ту́фля','ту́фли'),(7,'стихи́','стих','стихи́'),(10,'роди́тели','роди́тель','роди́тели'),(12,'фру́кты','фрукт','фру́кты')]:
 row=groups[l][word]
 assert len(row[:15])==15 and row[3]==singular and row[9]==plural,(l,word)
assert groups[5]['о́вощи'][10]=='овоще́й'
assert groups[7]['ту́фли'][10]=='ту́фель'
assert groups[7]['стихи́'][10]=='стихо́в'
assert groups[10]['роди́тели'][12]=='роди́телей'
print('PASS: two -мя rows removed; five plural-source paradigms added; other sheets preserved')
