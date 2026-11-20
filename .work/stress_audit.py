import gzip, collections, re, sys
from pathlib import Path
import openpyxl

root=Path(__file__).parents[1]
d=collections.defaultdict(set)
with gzip.open(root/'.work/ru-stress-dict.gz','rt',encoding='utf8') as f:
    for line in f:
        try:
            word,pos=line.rstrip('\n').split('\t')
            d[word].add(int(pos)-1)
        except ValueError: pass
yo=collections.defaultdict(set)
with gzip.open(root/'.work/ru-yo-dict.gz','rt',encoding='utf8') as f:
    for line in f:
        try:
            word,pos=line.rstrip('\n').split('\t')
            yo[word].add(int(pos)-1)
        except ValueError: pass
w=openpyxl.load_workbook(root/'大学俄语1_逐课生词分类与变位.xlsx',read_only=True)
pattern=re.compile('[А-Яа-яЁё]+')
missing=collections.Counter(); ambiguous=collections.Counter(); total=0; covered=0
for sh in w:
    for row in sh.iter_rows(min_row=4):
        if isinstance(row[0].value,str) and row[0].value.startswith('Урок'): continue
        cols=[0]+(list(range(3,15)) if sh.title=='名词变格' else list(range(5,23)) if sh.title=='动词' else [])
        for col in cols:
            if col>=len(row) or not isinstance(row[col].value,str): continue
            for token in pattern.findall(row[col].value):
                total+=1
                low=token.lower().replace('ё','е')
                positions=d.get(low)
                if positions or 'ё' in token.lower() or len([c for c in low if c in 'аеёиоуыэюя'])==1:
                    covered+=1
                    if positions and len(positions)>1: ambiguous[token]+=1
                else: missing[token]+=1
print('total',total,'covered',covered,'missing',sum(missing.values()),'unique_missing',len(missing),'ambiguous',len(ambiguous))
print('MISSING',missing.most_common(100))
print('AMBIGUOUS',ambiguous.most_common(30))
