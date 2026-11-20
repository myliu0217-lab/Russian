import collections, gzip, re, zipfile
from pathlib import Path
from lxml import etree

root=Path(__file__).parents[1]
src=root/'.work/styled_base.xlsx'
dst=root/'.work/stressed_final.xlsx'
ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
q=lambda s:f'{{{ns}}}{s}'
word_re=re.compile(r'[А-Яа-яЁё]+')
vowels=set('аеёиоуыэюя')
stress=collections.defaultdict(set)
with gzip.open(root/'.work/ru-stress-dict.gz','rt',encoding='utf8') as f:
    for line in f:
        try:
            word,pos=line.rstrip('\n').split('\t')
            stress[word].add(int(pos)-1)
        except ValueError: pass
class_a={}
for line in (root/'.work/ru-class-a.tsv').read_text(encoding='utf8').splitlines():
    if line.startswith('#') or not line: continue
    word,sg,pl=line.split('\t')
    class_a[word]=(int(sg)-1,int(pl)-1)
manual={'едемте':0,'летимте':3,'видьте':1,'китайски':3,'китайц':3,'русско':1}
names=['名词','名词变格','动词','副词','形容词','代词','前置词','数词','连接词','语气词','感叹词','说明']
counts=collections.Counter(); missing=collections.Counter(); bad=collections.Counter()

def stressed_indices(s,sheet,col):
    out=set()
    for match in word_re.finditer(s):
        word=match.group(); key=word.lower().replace('ё','е')
        if not any(c in vowels for c in word.lower()): continue
        if 'ё' in word.lower():
            positions=[i for i,c in enumerate(word.lower()) if c=='ё']
        elif key in manual: positions=[manual[key]]
        elif sheet=='名词变格' and key in class_a and col>=4:
            positions=[class_a[key][0 if col<=9 else 1]]
        elif key in stress: positions=[min(stress[key])]
        else:
            vv=[i for i,c in enumerate(word.lower()) if c in vowels]
            if len(vv)==1: positions=vv
            else:
                missing[word]+=1; continue
        for p in positions:
            if p<len(word) and word[p].lower() in vowels:
                out.add(match.start()+p)
            else: bad[word]+=1
    return out

def append_text_run(is_el,chunk,red):
    if not chunk: return
    r=etree.SubElement(is_el,q('r'))
    if red:
        rp=etree.SubElement(r,q('rPr'))
        etree.SubElement(rp,q('rFont'),val='Microsoft YaHei')
        etree.SubElement(rp,q('sz'),val='10')
        etree.SubElement(rp,q('color'),rgb='FFFF0000')
    t=etree.SubElement(r,q('t'))
    if chunk[0].isspace() or chunk[-1].isspace():
        t.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
    t.text=chunk

def rich_cell(c,positions):
    v=c.find(q('v'))
    if v is None or v.text is None: return
    s=v.text
    c.remove(v)
    c.set('t','inlineStr')
    is_el=etree.SubElement(c,q('is'))
    start=0
    for p in sorted(positions):
        append_text_run(is_el,s[start:p],False)
        ch=s[p]
        append_text_run(is_el,ch if ch.lower()=='ё' else ch+'\u0301',True)
        start=p+1
    append_text_run(is_el,s[start:],False)

with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst,'w') as zout:
    for entry in zin.infolist():
        data=zin.read(entry.filename)
        m=re.fullmatch(r'xl/worksheets/sheet(\d+)\.xml',entry.filename)
        if m:
            idx=int(m.group(1))-1
            if idx<len(names)-1:
                name=names[idx]
                tree=etree.fromstring(data)
                for row in tree.findall('.//'+q('sheetData')+'/'+q('row')):
                    rnum=int(row.get('r'))
                    if rnum<4: continue
                    first=row.find(q('c'))
                    if first is not None:
                        v=first.find(q('v'))
                        if v is not None and (v.text or '').startswith('Урок'): continue
                    for c in row.findall(q('c')):
                        ref=c.get('r')
                        col=0
                        for ch in re.match('[A-Z]+',ref).group(): col=col*26+ord(ch)-64
                        target=(col==1 or (name=='名词变格' and 3<=col<=15) or (name=='动词' and 6<=col<=24))
                        if not target: continue
                        v=c.find(q('v'))
                        if v is None or not v.text or not word_re.search(v.text): continue
                        indices=stressed_indices(v.text,name,col)
                        if indices:
                            rich_cell(c,indices)
                            counts[name]+=1
                data=etree.tostring(tree,encoding='utf-8',xml_declaration=True)
        zout.writestr(entry,data)
print('colored cells',dict(counts))
print('missing',dict(missing))
print('invalid positions',dict(bad))
print('output',dst)
