import zipfile
from pathlib import Path
import openpyxl

root=Path(__file__).parents[1]
source=root/'大学俄语1_逐课生词分类与变位.xlsx'
target=root/'.work/noun_print_split.xlsx'
a=openpyxl.load_workbook(source,rich_text=True)
b=openpyxl.load_workbook(target,rich_text=True)
s=b['名词变格']
assert s.sheet_properties.pageSetUpPr.fitToPage is False
assert s.page_setup.scale==95 and s.page_setup.fitToWidth is None
assert [(x.id,x.man) for x in s.col_breaks.brk]==[(9,True)]
assert s.print_title_cols=='$A:$C'
assert str(s.page_setup.paperSize)=='8' and s.page_setup.orientation=='landscape'
assert a.sheetnames==b.sheetnames
for sh in a:
    for row in sh:
        for c in row:
            assert str(c.value)==str(b[sh.title][c.coordinate].value),(sh.title,c.coordinate)
with zipfile.ZipFile(source) as za,zipfile.ZipFile(target) as zb:
    assert za.namelist()==zb.namelist()
    assert all(za.read(n)==zb.read(n) for n in za.namelist() if n!='xl/worksheets/sheet2.xml')
    assert zb.testzip() is None
print('PASS: break after I; A:C repeated; A3 landscape at 95%; all cells and other sheets unchanged')
