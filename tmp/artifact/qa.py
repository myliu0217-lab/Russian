import re
import unicodedata
from collections import defaultdict, Counter
from openpyxl import load_workbook

p='tmp/artifact/revised.xlsx'
w=load_workbook(p,read_only=False,data_only=False)
def norm(s):
    s=unicodedata.normalize('NFD',str(s or '').lower()).replace('\u0301','').replace('ё','е')
    return re.sub('[^а-я]','',s)
for sh in w:
    if sh.title=='说明': continue
    lesson=None; groups=defaultdict(list)
    for row in sh.iter_rows(min_row=4):
        word=row[0].value
        if not word: continue
        m=re.search(r'Урок\s+(\d+)',str(word))
        if m: lesson=int(m.group(1)); continue
        if lesson: groups[lesson].append(str(word))
    print('\n',sh.title,' '.join(f'{k}:{len(v)}' for k,v in sorted(groups.items())))
    for l,v in sorted(groups.items()):
        c=Counter(norm(x) for x in v)
        d=[x for x,n in c.items() if n>1]
        if d: print('DUP',l,d)
    if sh.title in ('名词','动词','形容词','副词','数词','代词'):
        for l in (1,3,8,9,10,11,12,16,17,18):
            if l in groups: print('L',l,':',' | '.join(groups[l]))
    if sh.title=='名词':
        for row in sh.iter_rows(min_row=4):
            v=str(row[0].value or '')
            if v.endswith('ь') and not re.match(r'^\[(阳|阴)\]',str(row[1].value or '')):
                print('GENDER_MISSING',v,row[1].value)
    if sh.title in ('名词','动词','形容词','副词','代词','数词','连接词','语气词','感叹词','前置词'):
        for row in sh.iter_rows(min_row=4):
            v=str(row[0].value or '')
            if not v or 'Урок ' in v: continue
            for token in re.findall(r'[А-Яа-яЁё\u0301]+',v):
                vowel_count=sum(c.lower() in 'аоуыэяёюие' for c in token)
                if vowel_count==1 and '\u0301' in token: print('MONO_ACCENT',sh.title,v)
                if re.search(r'[^аоуыэяёюиеАOUЫЭЯЁЮИЕ]\u0301',token): print('BAD_ACCENT',sh.title,v)
