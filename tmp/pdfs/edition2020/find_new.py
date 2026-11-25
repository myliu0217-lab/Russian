from pathlib import Path
import re,unicodedata,difflib
import openpyxl

root=Path(__file__).parent
wb=openpyxl.load_workbook(root.parents[2]/'大学俄语1_逐课生词分类与变位.xlsx',read_only=True,data_only=True)
entries=[]
for s in wb:
    if s.title in ['说明','名词变格']: continue
    for row in s.values:
        if row[0] and not re.search('Урок|俄语生词|按课次|来源',str(row[0])):
            entries.append(str(row[0]))
def norm(s):
    s=unicodedata.normalize('NFD',s.lower()).replace('\u0301','').replace('ё','е')
    return ''.join(c for c in s if 'а'<=c<='я')
lex=[norm(v) for e in entries for v in re.findall(r'[А-Яа-яЁё́]+',e)]
for path in sorted((root/'russian_crops').glob('new-*-*.tsv')):
    output=[]
    for line in path.read_text(encoding='utf8').splitlines():
        parts=line.split('\t')
        if len(parts)<3: continue
        y=int(float(parts[1])); raw=parts[2]
        russian=re.findall(r'[А-Яа-яЁё́]{2,}',raw)
        if not russian: continue
        first=norm(russian[0])
        if len(first)<3 or first in ['новые','слова','урок','упражнение','развитие','речи','письмо','русский','текст','диалог']:
            continue
        best=difflib.get_close_matches(first,lex,n=1,cutoff=.73)
        if not best: output.append(f'{y:4} {raw}')
    print('###',path.stem)
    print('\n'.join(output))
