import collections,copy,re,zipfile
from pathlib import Path
from lxml import etree

root=Path(__file__).parents[1]
src=root/'.work/base_after_cleanup.xlsx'
dst=root/'.work/word_endings_final.xlsx'
ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
q=lambda x:f'{{{ns}}}{x}'
word_re=re.compile(r'[А-Яа-яЁё]+')
accent='\u0301'
inf_re=re.compile(r'(ться|тись|чься|ть|ти|чь)$')
finite_re={
 1:re.compile(r'(юсь|усь|ю|у|м)$'),
 2:re.compile(r'(ешься|ёшься|ишься|шься|ешь|ёшь|ишь|шь)$'),
 3:re.compile(r'(ется|ётся|ится|тся|ет|ёт|ит|ст|т)$'),
 4:re.compile(r'(емся|ёмся|имся|мся|ем|ём|им|м)$'),
 5:re.compile(r'(етесь|ётесь|итесь|тесь|ете|ёте|ите|те)$'),
 6:re.compile(r'(ются|утся|ятся|атся|ют|ут|ят|ат)$'),
}
past_re={
 1:re.compile(r'(лся|лсь|л)$'),
 2:re.compile(r'(лась|ла)$'),
 3:re.compile(r'(лось|ло)$'),
 4:re.compile(r'(лись|ли)$'),
}
imperative_re={
 1:re.compile(r'(йся|йсь|ись|ися|ься|мся|й|и|ь|м)$'),
 2:re.compile(r'(йтесь|итесь|ьтесь|мтесь|йте|ите|ьте|мте|тесь|те)$'),
}
stats=collections.Counter(); unmatched=collections.defaultdict(list)

def colnum(ref):
    n=0
    for c in re.match(r'[A-Z]+',ref).group(): n=n*26+ord(c)-64
    return n

def text_runs(cell):
    is_el=cell.find(q('is'))
    if is_el is None: return None,None
    chars=[]
    for r in is_el:
        if r.tag==q('r'):
            rp=r.find(q('rPr'))
            t=r.find(q('t'))
            if t is not None and t.text:
                for ch in t.text: chars.append((ch,copy.deepcopy(rp) if rp is not None else None))
        elif r.tag==q('t') and r.text:
            for ch in r.text: chars.append((ch,None))
    return is_el,chars

def suffix_ranges(clean,col):
    words=list(word_re.finditer(clean))
    if not words: return []
    spans=[]
    if col==1:
        targets=[(words[0],inf_re)]
    elif 7<=col<=12:
        targets=[(words[0],finite_re[col-6])]
    elif 13<=col<=18:
        targets=[(words[0],finite_re[col-12])]
        if len(words)>1 and inf_re.search(words[1].group()):
            targets.append((words[1],inf_re))
    elif 19<=col<=22:
        targets=[(words[0],past_re[col-18])]
    elif 23<=col<=24:
        targets=[(words[0],imperative_re[col-22])]
    else: return []
    for word,pattern in targets:
        match=pattern.search(word.group().lower())
        if match: spans.append((word.start()+match.start(),word.start()+match.end()))
        else: unmatched[col].append(word.group())
    return spans

def bold_cell(cell,col):
    is_el,chars=text_runs(cell)
    if is_el is None or not chars: return
    s=''.join(ch for ch,_ in chars)
    clean=[]; mapping=[]
    for i,ch in enumerate(s):
        if ch!=accent:
            clean.append(ch); mapping.append(i)
    spans=suffix_ranges(''.join(clean),col)
    if not spans:return
    bold=set()
    for start,end in spans:
        for i in range(start,end):
            orig=mapping[i]
            bold.add(orig)
            if orig+1<len(s) and s[orig+1]==accent:bold.add(orig+1)
    if not bold:return
    new=[]
    for i,(ch,rp) in enumerate(chars):
        is_bold=i in bold
        style=etree.tostring(rp) if rp is not None else None
        key=(style,is_bold)
        if new and new[-1][0]==key:
            new[-1][1]+=ch
        else:
            new.append([key,ch,rp])
    for child in list(is_el): is_el.remove(child)
    for (style,is_bold),chunk,rp in new:
        r=etree.SubElement(is_el,q('r'))
        if is_bold:
            props=copy.deepcopy(rp) if rp is not None else etree.Element(q('rPr'))
            if props.find(q('b')) is None:
                b=etree.Element(q('b'))
                props.insert(1 if props.find(q('rFont')) is not None else 0,b)
            r.append(props)
        elif rp is not None:r.append(copy.deepcopy(rp))
        t=etree.SubElement(r,q('t'))
        if chunk[0].isspace() or chunk[-1].isspace():
            t.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
        t.text=chunk
    stats[col]+=1

with zipfile.ZipFile(src) as zin,zipfile.ZipFile(dst,'w') as zout:
    for entry in zin.infolist():
        data=zin.read(entry.filename)
        if entry.filename=='xl/worksheets/sheet3.xml':
            tree=etree.fromstring(data)
            for row in tree.findall('.//'+q('sheetData')+'/'+q('row')):
                if int(row.get('r'))<4:continue
                first=row.find(q('c'))
                if first is not None:
                    is_el=first.find(q('is'))
                    first_text=''.join(is_el.itertext()) if is_el is not None else ''
                    if first_text.startswith('Урок'):continue
                for cell in row.findall(q('c')):
                    col=colnum(cell.get('r'))
                    if col==1 or 7<=col<=24:bold_cell(cell,col)
            data=etree.tostring(tree,encoding='utf-8',xml_declaration=True)
        zout.writestr(entry,data)
print('bold cells',dict(sorted(stats.items())))
print('unmatched',{k:sorted(set(v)) for k,v in sorted(unmatched.items())})
print(dst)
