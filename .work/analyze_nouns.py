import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent / 'pydeps'))
import openpyxl
import pymorphy3

w = openpyxl.load_workbook(Path(__file__).parents[1] / '大学俄语1_逐课生词分类与变位.xlsx', read_only=True)
existing = {r[0] for r in list(w['名词变格'].values)[3:] if isinstance(r[0], str)}
m = pymorphy3.MorphAnalyzer()
cases = ['nomn','gent','datv','accs','ablt','loct']
for row in list(w['名词'].values)[3:]:
    noun, meaning = row[:2]
    if not isinstance(noun, str) or noun.startswith('Урок') or noun in existing:
        continue
    parses = [p for p in m.parse(noun) if 'NOUN' in p.tag and p.normal_form == noun]
    if not parses:
        continue
    p = parses[0]
    forms = []
    for number in ['sing','plur']:
        for case in cases:
            f = p.inflect({number,case})
            forms.append(f.word if f else '—')
    irregular_flags = []
    if 'Fixd' in p.tag: irregular_flags.append('indeclinable')
    if 'Pltm' in p.tag: irregular_flags.append('plural-only')
    if 'Sgtm' in p.tag: irregular_flags.append('singular-only')
    if forms[6] not in ['—',noun+'ы',noun+'и',noun[:-1]+'ы',noun[:-1]+'и',noun[:-1]+'а',noun[:-1]+'я',noun[:-1]+'е',noun[:-1]+'ья']:
        irregular_flags.append('plural-stem/ending')
    if irregular_flags:
        print(f'{noun}\t{meaning}\t{",".join(irregular_flags)}\t'+','.join(forms))
