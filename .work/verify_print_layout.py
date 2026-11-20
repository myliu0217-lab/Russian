import zipfile
from pathlib import Path
import openpyxl
from openpyxl.cell.rich_text import CellRichText

root=Path(__file__).parents[1]
old=openpyxl.load_workbook(root/'大学俄语1_逐课生词分类与变位.xlsx',rich_text=True)
new=openpyxl.load_workbook(root/'.work/print_ready.xlsx',rich_text=True)
assert old.sheetnames==new.sheetnames
for a,b in zip(old,new):
    assert a.max_row==b.max_row and a.max_column==b.max_column,(a.title,a.max_row,b.max_row)
    for row in a:
        for c in row:
            assert str(c.value)==str(b[c.coordinate].value),(a.title,c.coordinate)
    assert b.sheet_properties.pageSetUpPr.fitToPage
    assert b.page_setup.fitToHeight==0
    assert b.print_area
    if b.max_column==2:
        assert b.page_setup.fitToWidth==1
        assert b.page_setup.orientation=='portrait'
    else:
        assert b.page_setup.orientation=='landscape'
        assert b.print_title_cols in ('$A:$B','$A:$C')
    if b.title!='说明': assert b.print_title_rows=='$1:$3'
oldrich=sum(isinstance(c.value,CellRichText) for s in old for row in s for c in row)
newrich=sum(isinstance(c.value,CellRichText) for s in new for row in s for c in row)
assert oldrich==newrich,(oldrich,newrich)
with zipfile.ZipFile(root/'.work/print_ready.xlsx') as z: assert z.testzip() is None
print('PASS',len(new.sheetnames),'sheets, rich cells',newrich)
for s in new:
    print(s.title,s.page_setup.paperSize,s.page_setup.orientation,s.page_setup.fitToWidth,s.print_area,s.print_title_rows,s.print_title_cols)
