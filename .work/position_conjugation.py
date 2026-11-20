import copy,re,zipfile
from pathlib import Path
from lxml import etree

root=Path(__file__).parents[1]
src=root/'.work/conjugation_base.xlsx'
dst=root/'大学俄语1_逐课生词分类与变位_变位法.xlsx'
ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
q=lambda x:f'{{{ns}}}{x}'

def colnum(s):
    n=0
    for c in s: n=n*26+ord(c)-64
    return n
def colname(n):
    s=''
    while n:
        n,rem=divmod(n-1,26)
        s=chr(65+rem)+s
    return s
def newcol(n): return 6 if n==25 else n+1 if n>=6 else n
def newref(ref):
    m=re.fullmatch(r'(\$?)([A-Z]+)(\$?)(\d+)',ref)
    return m.group(1)+colname(newcol(colnum(m.group(2))))+m.group(3)+m.group(4)

with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst,'w') as zout:
    for entry in zin.infolist():
        content=zin.read(entry.filename)
        if entry.filename=='xl/worksheets/sheet3.xml':
            tree=etree.fromstring(content)
            cols=tree.find(q('cols'))
            if cols is not None:
                shifted=[]
                for el in cols.findall(q('col')):
                    for n in range(int(el.get('min')),int(el.get('max'))+1):
                        clone=copy.deepcopy(el)
                        clone.set('min',str(newcol(n)))
                        clone.set('max',str(newcol(n)))
                        shifted.append(clone)
                cols.clear()
                for el in sorted(shifted,key=lambda x:int(x.get('min'))): cols.append(el)
            for row in tree.findall('.//'+q('sheetData')+'/'+q('row')):
                cells=list(row.findall(q('c')))
                for c in cells: c.set('r',newref(c.get('r')))
                for c in cells: row.remove(c)
                for c in sorted(cells,key=lambda x:colnum(re.match(r'[A-Z]+',x.get('r')).group())): row.append(c)
            for merge in tree.findall('.//'+q('mergeCells')+'/'+q('mergeCell')):
                merge.set('ref',':'.join(newref(r) for r in merge.get('ref').split(':')))
            dim=tree.find(q('dimension'))
            if dim is not None and dim.get('ref'):
                dim.set('ref',':'.join(newref(r) for r in dim.get('ref').split(':')))
            content=etree.tostring(tree,encoding='utf-8',xml_declaration=True)
        zout.writestr(entry,content)
print(dst)
