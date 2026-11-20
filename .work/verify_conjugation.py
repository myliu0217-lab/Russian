import zipfile
from pathlib import Path
import openpyxl
from openpyxl.cell.rich_text import CellRichText

root=Path(__file__).parents[1]
old=openpyxl.load_workbook(root/'大学俄语1_逐课生词分类与变位_修订版.xlsx',rich_text=True)
new=openpyxl.load_workbook(root/'大学俄语1_逐课生词分类与变位_变位法.xlsx',rich_text=True)
assert old.sheetnames==new.sheetnames
def normalized(v): return str(v) if v is not None else None
for a,b in zip(old,new):
    for row in a:
        for c in row:
            nc=b.cell(c.row,c.column+1 if a.title=='动词' and c.column>=6 else c.column)
            assert normalized(c.value)==normalized(nc.value),(a.title,c.coordinate,nc.coordinate,c.value,nc.value)
s=new['动词']
assert s.max_column==25 and s.max_row==195
assert s['F3'].value=='变位法'
assert [s[f'{x}3'].value for x in 'GHIJKL']==['я','ты','он/она','мы','вы','они']
assert s['F5'].value=='第一变位法' and s['F6'].value=='第二变位法'
assert s['F8'].value=='特殊变位'
assert s['F5'].border.bottom.color.rgb[-6:]=='BFDCEB'
assert s.column_dimensions['F'].width==18
assert any(str(r)=='A4:Y4' for r in s.merged_cells.ranges)
oldrich=sum(isinstance(c.value,CellRichText) for sh in old for row in sh for c in row)
newrich=sum(isinstance(c.value,CellRichText) for sh in new for row in sh for c in row)
assert oldrich==newrich,(oldrich,newrich)
with zipfile.ZipFile(root/'大学俄语1_逐课生词分类与变位_变位法.xlsx') as z:
    assert z.testzip() is None
print('PASS',len(new.sheetnames),'sheets, 176 conjugation labels, rich text cells',newrich)
