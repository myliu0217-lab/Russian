from pathlib import Path
import pypdfium2 as pdfium

root = Path(__file__).parent
book = next(root.parents[2].glob('*2020*.pdf'))
pdf = pdfium.PdfDocument(str(book))
out = root / 'all'
out.mkdir(exist_ok=True)
for i in range(len(pdf)):
    path = out / f'{i+1:03}.jpg'
    if not path.exists():
        page = pdf[i]
        image = page.render(scale=0.65).to_pil().convert('RGB')
        image.save(path, quality=70)
print(len(pdf))
