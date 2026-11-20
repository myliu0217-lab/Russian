import json, re, sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).parent/'pydeps'))
import openpyxl, pymorphy3

root=Path(__file__).parents[1]
w=openpyxl.load_workbook(root/'大学俄语1_逐课生词分类与变位.xlsx',read_only=True)
lesson=0; noun_info={}
for row in w['名词'].values:
    v=row[0]
    if isinstance(v,str) and v.startswith('Урок'):
        lesson=int(re.search(r'\d+',v).group())
    elif isinstance(v,str) and lesson and v not in noun_info:
        noun_info[v]=(lesson,row[1],len(noun_info))
old={}; lesson=0
for row in list(w['名词变格'].values)[3:]:
    v=row[0]
    if isinstance(v,str) and v.startswith('Урок'):
        lesson=int(re.search(r'\d+',v).group())
    elif isinstance(v,str): old[v]=(lesson,list(row))

reasons={
'сон':'间接格脱落 о','дом':'复数为 дома','фото':'不变格','кафе':'不变格',
'повар':'复数为 повара','рынок':'间接格脱落 о','сумка':'复数第二格插入 о',
'окно':'复数第二格为 окон','утка':'复数第二格插入 о','яблоко':'复数第二格为 яблок',
'лодка':'复数第二格插入 о','полка':'复数第二格插入 о','ёлка':'复数第二格插入 о',
'шапка':'复数第二格插入 о','яйцо':'复数第二格为 яиц','поезд':'复数为 поезда',
'муж':'复数为 мужья','жена':'复数词干变为 жён-','метро':'不变格',
'кухня':'复数第二格为 кухонь','снег':'第六格另有 в снегу','тётя':'复数第二格为 тёть',
'радио':'不变格','дочка':'复数第二格为 дочек','адрес':'复数为 адреса',
'весна':'复数词干变为 вёсн-','такси':'不变格','сестра':'复数词干变为 сёстр-；第二格 сестёр',
'озеро':'复数词干变为 озёр-','письмо':'复数第二格为 писем','кино':'不变格',
'пальто':'不变格','платье':'复数第二格为 платьев','ребята':'只有复数',
'вечер':'复数为 вечера','китаец':'间接格词干变为 китайц-','иностранец':'间接格脱落 е',
'городок':'间接格脱落 о','столовая':'名词化形容词，按形容词词尾变格',
'директор':'复数为 директора','животное':'名词化形容词，按形容词词尾变格',
'солнце':'复数第二格为 солнц','остров':'复数为 острова','учитель':'复数为 учителя',
'стул':'复数为 стулья','дерево':'复数为 деревья','цех':'复数为 цеха',
}
m=pymorphy3.MorphAnalyzer(); cases=['nomn','gent','datv','accs','ablt','loct']
overrides={
 'стул':['стул','стула','стулу','стул','стулом','стуле','стулья','стульев','стульям','стулья','стульями','стульях'],
 'ребята':['—']*6+['ребята','ребят','ребятам','ребят','ребятами','ребятах'],
}
combined={}
for noun,(les,row) in old.items(): combined[noun]=row
for noun,reason in reasons.items():
    if noun in old: continue
    if noun not in noun_info:
        print('SKIP NOT IN TEXTBOOK',noun)
        continue
    parses=[p for p in m.parse(noun) if 'NOUN' in p.tag and p.normal_form==noun]
    if not parses: print('NO PARSE',noun); continue
    p=parses[0]
    forms=overrides.get(noun)
    if forms is None:
        forms=[]
        for number in ('sing','plur'):
            for case in cases:
                f=p.inflect({number,case})
                forms.append(f.word if f else '—')
    combined[noun]=[noun,noun_info[noun][1],reason]+forms
rows=[list(list(w['名词变格'].values)[i]) for i in range(3)]
for les in sorted({old[n][0] for n in old}|{noun_info[n][0] for n in reasons if n in noun_info}):
    nouns=[n for n in combined if (noun_info[n][0] if n in noun_info else old[n][0])==les]
    if not nouns: continue
    rows.append([f'Урок {les}｜第 {les} 课']+[None]*14)
    for n in sorted(nouns,key=lambda n:noun_info[n][2] if n in noun_info else 99999):
        rows.append(combined[n])
out={'rows':rows,'original_count':len(old),'added_count':len(combined)-len(old),'added':[n for n in combined if n not in old]}
(root/'.work/noun_rows.json').write_text(json.dumps(out,ensure_ascii=False),encoding='utf8')
print('old',out['original_count'],'added',out['added_count'],'rows',len(rows))
print('added',', '.join(out['added']))
