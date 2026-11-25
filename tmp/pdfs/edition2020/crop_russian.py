from pathlib import Path
from PIL import Image, ImageOps, ImageEnhance

root=Path(__file__).parent
out=root/'russian_crops'
out.mkdir(exist_ok=True)
for path in sorted((root/'vocab').glob('new-*.png')):
    image=Image.open(path).convert('RGB')
    w,h=image.size
    for name,(x0,x1) in {'left':(0.075,0.35),'right':(0.48,0.74)}.items():
        if path.stem=='new-019' and name=='right': x0,x1=0.36,0.58
        crop=image.crop((round(w*x0),0,round(w*x1),h))
        gray=ImageOps.grayscale(crop)
        enhanced=ImageEnhance.Contrast(gray).enhance(1.8)
        enhanced.save(out/f'{path.stem}-{name}.png')
