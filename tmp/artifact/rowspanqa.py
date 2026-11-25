import re
from pathlib import Path
t=Path('俄语1review.md').read_text(encoding='utf8')
for section in re.findall(r'<table>.*?</table>',t,re.S):
 if 'rowspan=' not in section:continue
 rows=re.findall(r'<tr>.*?</tr>',section,re.S)
 last=None;count=0
 for row in rows:
  m=re.search(r'<td rowspan="(\d+)">([^<]+)</td>',row)
  if m:
   if last:print(last,count,'ok' if count==last[1] else 'MISMATCH')
   last=(m.group(2),int(m.group(1)));count=1
  elif last:count+=1
 if last:print(last,count,'ok' if count==last[1] else 'MISMATCH')
