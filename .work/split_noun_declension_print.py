import zipfile
from pathlib import Path
from lxml import etree

root=Path(__file__).parents[1]
src=root/'大学俄语1_逐课生词分类与变位.xlsx'
dst=root/'.work/noun_print_split.xlsx'
ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
q=lambda s:f'{{{ns}}}{s}'

with zipfile.ZipFile(src) as zin,zipfile.ZipFile(dst,'w') as zout:
    for entry in zin.infolist():
        data=zin.read(entry.filename)
        if entry.filename=='xl/worksheets/sheet2.xml':
            tree=etree.fromstring(data)
            pr=tree.find(q('sheetPr'))
            fit=pr.find(q('pageSetUpPr'))
            fit.set('fitToPage','0')
            setup=tree.find(q('pageSetup'))
            setup.attrib.pop('fitToWidth',None)
            setup.attrib.pop('fitToHeight',None)
            setup.set('scale','95')
            old=tree.find(q('colBreaks'))
            if old is not None:tree.remove(old)
            breaks=etree.Element(q('colBreaks'),count='1',manualBreakCount='1')
            etree.SubElement(breaks,q('brk'),id='9',min='0',max='1048575',man='1')
            tree.insert(tree.index(setup)+1,breaks)
            data=etree.tostring(tree,encoding='utf-8',xml_declaration=True)
        zout.writestr(entry,data)
print(dst)
