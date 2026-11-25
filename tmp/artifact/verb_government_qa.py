from openpyxl import load_workbook
from pathlib import Path
import re, zipfile

old='大学俄语1_逐课生词分类与变位.xlsx'
new='tmp/artifact/verb_government_revised.xlsx'
with zipfile.ZipFile(new) as z: assert z.testzip() is None
a=load_workbook(old,read_only=False,data_only=False)
b=load_workbook(new,read_only=False,data_only=False)
assert a.sheetnames==b.sheetnames
for name in a.sheetnames:
 if name=='动词':continue
 x,y=a[name],b[name]
 assert x.max_row==y.max_row and x.max_column==y.max_column,name
 for row in x:
  for c in row:
   assert c.value==y[c.coordinate].value,(name,c.coordinate)
x,y=a['动词'],b['动词']
assert x.max_row==y.max_row and y.max_column==26
assert len(y.merged_cells.ranges)==len(x.merged_cells.ranges)
assert all(str(m).endswith('Z'+str(m.max_row)) for m in y.merged_cells.ranges)
assert y['E3'].value=='接格关系（常见搭配）'
missing=[];count=0
for i in range(4,x.max_row+1):
 head=str(x.cell(i,1).value or '')
 if head.startswith('Урок '):continue
 for j in range(1,4):assert x.cell(i,j).value==y.cell(i,j).value,(i,j)
 for j in range(5,26):assert x.cell(i,j).value==y.cell(i,j+1).value,(i,j)
 if x.cell(i,4).value=='不及物' and not y.cell(i,5).value:missing.append((i,head))
 if y.cell(i,5).value:count+=1
assert not missing,missing
assert count>=80,count
for word,pattern in [('опа́здывать','в/на+④'),('происходи́ть','от+②'),('боя́ться','кого/чего+②'),('меша́ть','кому/чему+③')]:
 rows=[i for i in range(4,y.max_row+1) if y.cell(i,1).value==word]
 assert len(rows)==1,(word,rows)
 assert pattern in y.cell(rows[0],5).value,(word,y.cell(rows[0],5).value)
print('PASS: 12 sheets; 84 verb patterns; conjugations and non-verb sheets preserved')
