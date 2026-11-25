from pathlib import Path
import pypdfium2 as pdfium

root=Path(__file__).parent
workspace=root.parents[2]
new=pdfium.PdfDocument(str(next(workspace.glob('*2020*.pdf'))))
old=pdfium.PdfDocument(str(workspace/'大学俄语东方（新版）学生用书1.pdf'))
pages=[19,28,37,46,56,57,67,77,78,90,105,106,120,136,151,152,168,169,181,182,197,198,213,214,229,230,244,245]
out=root/'vocab'
out.mkdir(exist_ok=True)
for name,doc,offset in [('new',new,0),('old',old,1)]:
    for p in pages:
        path=out/f'{name}-{p:03}.png'
        if not path.exists():
            image=doc[p-1+offset].render(scale=2).to_pil().convert('RGB')
            image.save(path)
print(len(pages))
