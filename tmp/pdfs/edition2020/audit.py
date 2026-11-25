from pathlib import Path
import sys
import re,unicodedata,difflib
import openpyxl

root=Path(__file__).parent
book=root.parents[2]/'大学俄语1_逐课生词分类与变位.xlsx'
wb=openpyxl.load_workbook(book,read_only=True,data_only=True)
bylesson={i:[] for i in range(1,19)}
for sheet in wb:
    if sheet.title in ['说明','名词变格']:
        continue
    lesson=None
    for row in sheet.values:
        first=str(row[0] or '')
        m=re.search(r'Урок\s+(\d+)',first,re.I)
        if m:
            lesson=int(m.group(1))
        elif lesson and first and first not in ['俄语生词','俄语词汇']:
            bylesson[lesson].append((sheet.title,first))

pages={1:[19],2:[28],3:[37],4:[46],5:[56,57],6:[67],7:[77,78],8:[90],9:[105,106],10:[120],11:[136],12:[151,152],13:[168,169],14:[181,182],15:[197,198],16:[213,214],17:[229,230],18:[244,245]}
if len(sys.argv)>1:
    for n in map(int,sys.argv[1:]):
        print('===',n)
        print(' / '.join(f'{c}:{w}' for c,w in bylesson[n]))
    raise SystemExit
def norm(s):
    s=unicodedata.normalize('NFD',s.lower()).replace('\u0301','').replace('ё','е')
    return ''.join(c for c in s if 'а'<=c<='я')

for lesson, entries in bylesson.items():
    text=' '.join((root/'vocab'/f'new-{p:03}.txt').read_text(encoding='utf8') for p in pages[lesson])
    tokens=[norm(t) for t in re.findall(r'[А-Яа-яЁё́]+',text)]
    tokens=[t for t in tokens if len(t)>1]
    oldtext=' '.join((root/'vocab'/f'old-{p:03}.txt').read_text(encoding='utf8') for p in pages[lesson])
    oldtokens=[norm(t) for t in re.findall(r'[А-Яа-яЁё́]+',oldtext)]
    def found(variants,tokens):
        return any(v in tokens or (len(v)>4 and difflib.get_close_matches(v,tokens,n=1,cutoff=.78)) for v in variants)
    missing=[]
    oldonly=[]
    for category,word in entries:
        variants=[norm(v) for v in re.findall(r'[А-Яа-яЁё́]+',word)]
        if not variants: continue
        if not found(variants,tokens):
            missing.append((category,word))
            if found(variants,oldtokens): oldonly.append((category,word))
    print(f'=== LESSON {lesson}: {len(entries)} entries, {len(missing)} OCR-missing')
    print(' / '.join(f'{cat}:{w}' for cat,w in missing))
    print('OLD-ONLY:', ' / '.join(f'{cat}:{w}' for cat,w in oldonly))
