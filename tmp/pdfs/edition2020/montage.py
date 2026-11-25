from pathlib import Path
from PIL import Image,ImageDraw

root=Path(__file__).parent
paths=sorted(root.glob('p-*.png'))
canvas=Image.new('RGB',(360*4,530*((len(paths)+3)//4)),'white')
draw=ImageDraw.Draw(canvas)
for i,path in enumerate(paths):
    im=Image.open(path).convert('RGB')
    im.thumbnail((360,500))
    x,y=(i%4)*360,(i//4)*530
    canvas.paste(im,(x,y+20))
    draw.text((x+5,y+2),path.stem,fill='black')
canvas.save(root/'contact.png')
