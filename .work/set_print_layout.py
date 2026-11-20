import copy,zipfile
from pathlib import Path
from lxml import etree

root=Path(__file__).parents[1]
src=root/'.work/print_base.xlsx'
dst=root/'.work/print_ready.xlsx'
ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
q=lambda s:f'{{{ns}}}{s}'
specs=[
 ('名词',501,'B','A4','portrait',1,None),
 ('名词变格',92,'O','A3','landscape',2,'C'),
 ('动词',195,'Y','A3','landscape',3,'B'),
 ('副词',107,'B','A4','portrait',1,None),
 ('形容词',106,'B','A4','portrait',1,None),
 ('代词',58,'B','A4','portrait',1,None),
 ('前置词',22,'B','A4','portrait',1,None),
 ('数词',43,'B','A4','portrait',1,None),
 ('连接词',16,'B','A4','portrait',1,None),
 ('语气词',14,'B','A4','portrait',1,None),
 ('感叹词',18,'B','A4','portrait',1,None),
 ('说明',5,'B','A4','portrait',1,None),
]
paper={'A4':'9','A3':'8'}

def worksheet_layout(data,spec):
    name,rows,last,paper_name,orientation,width,repeat=spec
    tree=etree.fromstring(data)
    pr=tree.find(q('sheetPr'))
    if pr is None:
        pr=etree.Element(q('sheetPr'));tree.insert(0,pr)
    fit=pr.find(q('pageSetUpPr'))
    if fit is None:fit=etree.SubElement(pr,q('pageSetUpPr'))
    fit.set('fitToPage','1')
    opts=tree.find(q('printOptions'))
    if opts is None:
        opts=etree.Element(q('printOptions'))
        margins=tree.find(q('pageMargins'))
        tree.insert(tree.index(margins) if margins is not None else len(tree),opts)
    opts.set('horizontalCentered','1')
    margins=tree.find(q('pageMargins'))
    if margins is None:
        margins=etree.Element(q('pageMargins'))
        tree.insert(tree.index(opts)+1,margins)
    for key,value in {'left':'0.28','right':'0.28','top':'0.42','bottom':'0.42','header':'0.18','footer':'0.18'}.items():
        margins.set(key,value)
    setup=tree.find(q('pageSetup'))
    if setup is None:
        setup=etree.Element(q('pageSetup'))
        tree.insert(tree.index(margins)+1,setup)
    for key,value in {'paperSize':paper[paper_name],'orientation':orientation,'fitToWidth':str(width),'fitToHeight':'0','pageOrder':'overThenDown'}.items():
        setup.set(key,value)
    setup.attrib.pop('scale',None)
    return etree.tostring(tree,encoding='utf-8',xml_declaration=True)

def workbook_names(data):
    tree=etree.fromstring(data)
    names=tree.find(q('definedNames'))
    if names is None:
        names=etree.Element(q('definedNames'))
        sheets=tree.find(q('sheets'))
        tree.insert(tree.index(sheets)+1,names)
    for item in list(names):
        if item.get('name') in ('_xlnm.Print_Area','_xlnm.Print_Titles'):
            names.remove(item)
    for idx,(name,rows,last,paper_name,orientation,width,repeat) in enumerate(specs):
        quoted="'"+name.replace("'","''")+"'"
        area=etree.SubElement(names,q('definedName'),name='_xlnm.Print_Area',localSheetId=str(idx))
        area.text=f'{quoted}!$A$1:${last}${rows}'
        if name!='说明':
            title=etree.SubElement(names,q('definedName'),name='_xlnm.Print_Titles',localSheetId=str(idx))
            title.text=f'{quoted}!$1:$3'+(f',{quoted}!$A:${repeat}' if repeat else '')
    return etree.tostring(tree,encoding='utf-8',xml_declaration=True)

with zipfile.ZipFile(src) as zin,zipfile.ZipFile(dst,'w') as zout:
    for entry in zin.infolist():
        data=zin.read(entry.filename)
        if entry.filename=='xl/workbook.xml':
            data=workbook_names(data)
        elif entry.filename.startswith('xl/worksheets/sheet') and entry.filename.endswith('.xml'):
            number=int(entry.filename.split('sheet')[-1].split('.xml')[0])
            if 1<=number<=len(specs):data=worksheet_layout(data,specs[number-1])
        zout.writestr(entry,data)
print(dst)
