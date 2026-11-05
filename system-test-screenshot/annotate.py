#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""截圖標註：在動作元素畫加粗紅框＋一支紅箭頭指向框。供測試報告截圖共用。"""
import math
from PIL import Image, ImageDraw

RED=(230,0,0)
BOX_W=7      # 紅框粗細(原本4→加粗)
ARROW_W=8    # 箭頭粗細
ARROW_LEN=120  # 箭尾離框距離
HEAD=26      # 箭頭頭部長度
GAP=16       # 箭頭尖端與框的間距(不要貼到框)

def _arrow(d, tail, head, color=RED, width=ARROW_W):
    d.line([tail, head], fill=color, width=width)
    ang=math.atan2(head[1]-tail[1], head[0]-tail[0])
    for a in (ang+math.radians(148), ang-math.radians(148)):
        d.line([head, (head[0]+HEAD*math.cos(a), head[1]+HEAD*math.sin(a))], fill=color, width=width)

def annotate(img_path, box, out_path=None, pad=6):
    """box=dict(x,y,width,height)(viewport座標)。畫加粗紅框＋箭頭(自動挑空間最大那側)。"""
    im=Image.open(img_path).convert("RGB"); d=ImageDraw.Draw(im); W,H=im.size
    x0,y0=box["x"]-pad, box["y"]-pad
    x1,y1=box["x"]+box["width"]+pad, box["y"]+box["height"]+pad
    for w in range(BOX_W):
        d.rectangle([x0-w,y0-w,x1+w,y1+w], outline=RED)
    cx,cy=(x0+x1)/2,(y0+y1)/2
    g=BOX_W+GAP   # 從框外緣再退開一點，箭頭尖端不貼框
    spaces={"top":y0,"bottom":H-y1,"left":x0,"right":W-x1}
    side=max(spaces,key=spaces.get)
    if side=="top":    head=(cx, y0-g); tail=(cx, max(8, y0-g-ARROW_LEN))
    elif side=="bottom": head=(cx, y1+g); tail=(cx, min(H-8, y1+g+ARROW_LEN))
    elif side=="left":  head=(x0-g, cy); tail=(max(8, x0-g-ARROW_LEN), cy)
    else:               head=(x1+g, cy); tail=(min(W-8, x1+g+ARROW_LEN), cy)
    _arrow(d, tail, head)
    im.save(out_path or img_path)

def annotate_page(page, box, out_path, pad=6):
    """Playwright: 截當前viewport→標註。"""
    tmp=str(out_path)+".raw.png"; page.screenshot(path=tmp)
    annotate(tmp, box, out_path, pad)
    import os; os.remove(tmp)

def annotate_multi(img_path, boxes, out_path, pad=6):
    """多個動作點：只畫紅框＋紅底白字編號(1,2,3…)，不畫箭頭（多箭頭易壓到相鄰文字）。
    boxes 只有一個時改走 annotate()（框＋箭頭）。圖說用①②③對應編號。"""
    from PIL import ImageFont
    boxes=[b for b in boxes if b]
    if len(boxes)==1: return annotate(img_path, boxes[0], out_path, pad)
    font=ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial Bold.ttf", 26)
    im=Image.open(img_path).convert("RGB"); d=ImageDraw.Draw(im)
    for i,b in enumerate(boxes,1):
        x0,y0,x1,y1=b["x"]-pad,b["y"]-pad,b["x"]+b["width"]+pad,b["y"]+b["height"]+pad
        d.rectangle([x0,y0,x1,y1], outline=RED, width=5)
        r=17; cx,cy=max(r,x0-r+4),max(r,y0-r+4)
        d.ellipse([cx-r,cy-r,cx+r,cy+r], fill=RED)
        d.text((cx,cy), str(i), fill="white", font=font, anchor="mm")
    im.save(out_path)

def annotate_manual(img_path, boxes, out_path, pad=5, box_w=3, badge_d=26, number=True):
    """【手冊模式，2026-11-01】操作手冊用：一律紅框＋紅底白字編號圓圈、不畫箭頭，
    單一動作點也編號①（圖說用①②③對應）。框線預設 3px、圓圈直徑 26px（比測試報告
    的 7px 框細，手冊縮進 Word 16cm 寬仍可辨識）。boxes=[dict(x,y,width,height)…]。
    number=False 時只畫框不編號。"""
    from PIL import ImageFont
    boxes=[b for b in boxes if b]
    font=ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial Bold.ttf", int(badge_d*0.72))
    im=Image.open(img_path).convert("RGB"); d=ImageDraw.Draw(im)
    for i,b in enumerate(boxes,1):
        x0,y0,x1,y1=b["x"]-pad,b["y"]-pad,b["x"]+b["width"]+pad,b["y"]+b["height"]+pad
        d.rectangle([x0,y0,x1,y1], outline=RED, width=box_w)
        if number:
            r=badge_d//2; cx,cy=max(r,x0-r+4),max(r,y0-r+4)
            d.ellipse([cx-r,cy-r,cx+r,cy+r], fill=RED, outline="white", width=2)
            d.text((cx,cy), str(i), fill="white", font=font, anchor="mm")
    im.save(out_path)

def annotate_manual_page(page, boxes, out_path, **kw):
    """Playwright：截當前viewport→手冊模式標註。"""
    tmp=str(out_path)+".raw.png"; page.screenshot(path=tmp)
    annotate_manual(tmp, boxes, out_path, **kw)
    import os; os.remove(tmp)

if __name__=="__main__":
    # 產示意圖：假畫面+假按鈕，展示加粗框+箭頭
    im=Image.new("RGB",(900,500),(245,247,250)); d=ImageDraw.Draw(im)
    d.rectangle([560,300,700,360], fill=(60,130,200))  # 假按鈕
    d.text((585,322),"匯出Excel",fill=(255,255,255))
    im.save("/tmp/_demo.png")
    annotate("/tmp/_demo.png", {"x":560,"y":300,"width":140,"height":60},
             "/tmp/_demo_annotated.png")
    print("示意圖已產出")
