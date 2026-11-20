import re, sys, zipfile
from pathlib import Path
import openpyxl
from openpyxl.cell.rich_text import CellRichText

root=Path(__file__).parents[1]
old=openpyxl.load_workbook(root/'大学俄语1_逐课生词分类与变位.xlsx',rich_text=True)
new=openpyxl.load_workbook(root/'.work/stressed_final.xlsx',rich_text=True)
plain=lambda x: str(x).replace('\u0301','') if x is not None else None
assert new.sheetnames==old.sheetnames+['说明']
for oldsh in old:
    newsh=new[oldsh.title]
    if oldsh.title!='名词变格':
        assert newsh.max_row==oldsh.max_row,(oldsh.title,newsh.max_row,oldsh.max_row)
        for row in oldsh:
            for c in row:
                if c.coordinate=='A2': continue
                assert plain(newsh[c.coordinate].value)==plain(c.value),(oldsh.title,c.coordinate,c.value,newsh[c.coordinate].value)
    for r in range(4,newsh.max_row+1):
        assert newsh.row_dimensions[r].height==20,(oldsh.title,r,newsh.row_dimensions[r].height)
    for cell in [newsh['A3'],newsh[f'A{newsh.max_row}']]:
        assert cell.border.bottom.color.rgb[-6:]=='BFDCEB',(oldsh.title,cell.coordinate,cell.border.bottom.color)
        assert cell.border.right.color.rgb[-6:]=='BFDCEB',(oldsh.title,cell.coordinate,cell.border.right.color)
old_rows={str(r[0].value):[plain(c.value) for c in r] for r in old['名词变格'].iter_rows(min_row=4) if r[0].value and not str(r[0].value).startswith('Урок')}
new_rows={plain(r[0].value):[plain(c.value) for c in r] for r in new['名词变格'].iter_rows(min_row=4) if r[0].value and not plain(r[0].value).startswith('Урок')}
assert len(old_rows)==28 and len(new_rows)==74
for noun,oldrow in old_rows.items(): assert new_rows[noun]==oldrow,(noun,oldrow,new_rows[noun])
for coord in ['A5','A6']:
    cell=new['名词'][coord]
    assert isinstance(cell.value,CellRichText),cell.value
    assert '\u0301' in str(cell.value),cell.value
    assert any(getattr(run,'font',None) and run.font.color and run.font.color.rgb=='FFFF0000' for run in cell.value),cell.value
for p in ['.work/stressed_final.xlsx']:
    with zipfile.ZipFile(root/p) as z: assert z.testzip() is None
print('PASS sheets',len(new.sheetnames),'nouns',len(new_rows),'original noun rows preserved',len(old_rows))
print('sample noun',str(new['名词']['A5'].value))
print('sample declension',str(new['名词变格']['A5'].value))
print('sample verb',str(new['动词']['A5'].value))
