from pathlib import Path
from PIL import Image, ImageDraw

root = Path(__file__).parent
files = sorted((root / 'all').glob('*.jpg'))
for start in range(0, len(files), 30):
    batch = files[start:start+30]
    w,h,cols = 240,370,5
    canvas = Image.new('RGB',(w*cols,h*((len(batch)+cols-1)//cols)),'white')
    d=ImageDraw.Draw(canvas)
    for j,p in enumerate(batch):
        im=Image.open(p)
        im.thumbnail((w-5,h-20))
        x,y=(j%cols)*w,(j//cols)*h
        canvas.paste(im,(x,y+20))
        d.text((x+3,y+2),p.stem,fill='black')
    canvas.save(root/f'contact-{start+1:03}-{start+len(batch):03}.jpg',quality=80)
