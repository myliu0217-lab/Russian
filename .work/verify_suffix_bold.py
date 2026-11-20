import collections,zipfile
from pathlib import Path
import openpyxl
from openpyxl.cell.rich_text import CellRichText

root=Path(__file__).parents[1]
old=openpyxl.load_workbook(root/'大学俄语1_逐课生词分类与变位.xlsx',rich_text=True)
new=openpyxl.load_workbook(root/'.work/word_endings_final.xlsx',rich_text=True)
assert old.sheetnames==new.sheetnames
for a,b in zip(old,new):
    for row in a:
        for c in row:
            if a.title=='动词' and c.column==6 and c.row>=4:continue
            assert str(c.value)==str(b[c.coordinate].value),(a.title,c.coordinate)
labels=collections.Counter(new['动词'][f'F{r}'].value for r in range(4,196))
assert labels['一']==115 and labels['二']==56 and labels[None]==21,labels
noun=new['名词变格']
assert all(not c.font.bold for row in noun.iter_rows(min_row=4) for c in row if c.value is not None and not str(noun.cell(c.row,1).value).startswith('Урок'))
def runs(c):
    return [(str(x),bool(getattr(x,'font',None) and x.font.b),
             x.font.color.rgb if getattr(x,'font',None) and x.font.color and x.font.color.type=='rgb' else None)
            for x in c.value] if isinstance(c.value,CellRichText) else []
assert any(b for _,b,_ in runs(new['动词']['G11']))
assert any(b and color=='FFFF0000' for _,b,color in runs(new['动词']['G5']))
oldred=sum(sum(1 for _,_,color in runs(c) if color=='FFFF0000') for s in old for row in s for c in row)
newred=sum(sum(1 for _,_,color in runs(c) if color=='FFFF0000') for s in new for row in s for c in row)
assert oldred==newred,(oldred,newred)
with zipfile.ZipFile(root/'.work/word_endings_final.xlsx') as z: assert z.testzip() is None
print('PASS labels',dict(labels),'red runs',newred)
for address in ['A11','G11','H23','M11','S11','W11','G5']:
    c=new['动词'][address]
    print(address,str(c.value),runs(c))
