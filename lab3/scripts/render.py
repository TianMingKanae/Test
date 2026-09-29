import json, textwrap
from PIL import Image, ImageDraw, ImageFont
S=2
MONO=ImageFont.truetype('/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf',13*S)
CJK=ImageFont.truetype('/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',13*S)
BG=(48,10,36); FG=(235,235,235); GREEN=(138,226,52); BAR=(60,60,60); BARTXT=(220,220,220)
CW=MONO.getlength('M'); LH=int(17*S)
def wide(c): return ord(c)>0x2e80
def term(title, text, cols=100, min_rows=0):
    lines=[]
    for ln in text.split('\n'):
        ln=ln.replace('\t','    ')
        if not ln: lines.append(''); continue
        while ln:
            # wrap by display width
            w=0;i=0
            while i<len(ln) and w+(2 if wide(ln[i]) else 1)<=cols: w+=2 if wide(ln[i]) else 1; i+=1
            lines.append(ln[:i]); ln=ln[i:]
    rows=max(len(lines),min_rows)
    pad=10*S; bar=26*S
    W=int(cols*CW+2*pad); H=bar+rows*LH+2*pad
    im=Image.new('RGB',(W,H),BG); d=ImageDraw.Draw(im)
    d.rectangle([0,0,W,bar],fill=BAR)
    for k,col in enumerate([(237,106,94),(245,191,79),(98,197,84)]):
        cx=12*S+k*18*S; d.ellipse([cx-5*S,bar/2-5*S,cx+5*S,bar/2+5*S],fill=col)
    tw=d.textlength(title,font=CJK); d.text(((W-tw)/2,bar/2),title,font=CJK,fill=BARTXT,anchor='lm')
    y=bar+pad
    for ln in lines:
        x=pad
        if ln.startswith('$ '):
            d.text((x,y),'$',font=MONO,fill=GREEN); x+=2*CW; ln=ln[2:]
        for c in ln:
            f=CJK if wide(c) else MONO
            d.text((x,y),c,font=f,fill=FG); x+=(2 if wide(c) else 1)*CW
        y+=LH
    return im
def hstack(ims,gap=12*S):
    H=max(i.height for i in ims); W=sum(i.width for i in ims)+gap*(len(ims)-1)
    out=Image.new('RGB',(W,H),'white'); x=0
    for i in ims: out.paste(i,(x,0)); x+=i.width+gap
    return out
def vstack(ims,gap=12*S):
    W=max(i.width for i in ims); H=sum(i.height for i in ims)+gap*(len(ims)-1)
    out=Image.new('RGB',(W,H),'white'); y=0
    for i in ims: out.paste(i,(0,y)); y+=i.height+gap
    return out
def save(im,name): im.save(name, optimize=True); print(name, im.size)

d={}
for f in ['q12.json','q3.json','q4.json']: d.update(json.load(open(f)))
T=lambda key,i: d[key][i]
# Q1
(t1,x1),(t2,x2)=d['q1']
save(hstack([term(t1,x1,52,5),term(t2,x2,52,5)]),'img_q1.png')
# Q2
subs=[term(t,x,62,13) for t,x in d['q2'][:4]]
pt,px=d['q2'][4]
save(vstack([hstack(subs[:2]),hstack(subs[2:]),term(pt,px,129)]),'img_q2.png')
# Q3
(at,ax),(st,sx),(pt,px)=d['q3']
save(term(at,ax,108),'img_q3a.png')
save(vstack([term(st,sx,108),term(pt,px,108)]),'img_q3b.png')
# Q4
(ct,cx),(ast,asx),(apt,apx),(bst,bsx),(bpt,bpx)=d['q4']
save(term(ct,cx,129),'img_q4a.png')
save(vstack([hstack([term(ast,asx,62,8),term(bst,bsx,62,8)]),hstack([term(apt,apx,62,6),term(bpt,bpx,62,6)])]),'img_q4b.png')
