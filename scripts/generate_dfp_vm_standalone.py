#!/usr/bin/env python3
"""Generate DFP cell BMPs on Ubuntu (needs cairosvg + pillow)."""
import json, math, os, io
import cairosvg
from PIL import Image

ROOT = "/tmp/dfp_gen"
NODES_PATH = os.path.join(ROOT, "lib/cde/assets/dfp_icon_nodes.json")
OUT_DIR = os.path.join(ROOT, "lib/cde/assets/dfp_cells")
CELL_PX, ICON_PX, STROKE = 20, 14, "#ffffff"
CATEGORY_HUE = {1:4,2:28,3:268,4:210,5:150,6:320,7:190}
CODE_CATEGORY = {
    "1.01":1,"1.02":1,"1.03":1,"1.04":1,"1.05":1,"1.07":1,"1.08":1,"1.09":1,"1.10":1,"1.11":1,
    "2.1":2,"2.2":2,"2.3":2,"2.4":2,"2.6":2,"2.7":2,"2.8":2,"2.9":2,
    "3.1":3,"3.2":3,"3.3":3,"4.1":4,"4.2":4,"4.3":4,"4.7":4,
    "5.1":5,"5.2":5,"5.3":5,"5.4":5,"6.1":6,"6.2":6,"6.3":6,"6.4":6,"6.5":6,
    "7.1":7,"7.2":7,"7.3":7,"7.4":7,
}

def hsl_to_rgb(h,s,l):
    s,l=s/100,l/100;c=(1-abs(2*l-1))*s;hp=(h%360)/60;x=c*(1-abs(hp%2-1))
    if hp<1:r1,g1,b1=c,x,0
    elif hp<2:r1,g1,b1=x,c,0
    elif hp<3:r1,g1,b1=0,c,x
    elif hp<4:r1,g1,b1=0,x,c
    elif hp<5:r1,g1,b1=x,0,c
    else:r1,g1,b1=c,0,x
    m=l-c/2
    return int((r1+m)*255),int((g1+m)*255),int((b1+m)*255)

def attrs(a):
    skip = frozenset(["key", "fill"])
    return " ".join('{}="{}"'.format(k, v) for k, v in sorted(a.items()) if k not in skip)

def svg_elems(node):
    p=[]
    for item in node:
        tag = item[0]
        attr = item[1] if len(item) > 1 else {}
        a = attrs(attr)
        if tag=="path": p.append('<path {} fill="none" stroke="{}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>'.format(a,STROKE))
        elif tag=="circle": p.append('<circle {} fill="none" stroke="{}" stroke-width="2"/>'.format(a,STROKE))
        elif tag=="line": p.append('<line {} stroke="{}" stroke-width="2" stroke-linecap="round"/>'.format(a,STROKE))
        elif tag=="rect": p.append('<rect {} fill="none" stroke="{}" stroke-width="2"/>'.format(a,STROKE))
        elif tag in ("polyline","polygon"): p.append('<{} {} fill="none" stroke="{}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>'.format(tag,a,STROKE))
    return "\n".join(p)

def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    data=json.load(open(NODES_PATH))
    for code,entry in sorted(data.items()):
        if code=="fallback": continue
        cat=CODE_CATEGORY.get(code,1); rgb=hsl_to_rgb(CATEGORY_HUE.get(cat,220),64,45)
        svg='<?xml version="1.0"?><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="{}" height="{}">{}</svg>'.format(ICON_PX,ICON_PX,svg_elems(entry["node"]))
        icon=Image.open(io.BytesIO(cairosvg.svg2png(bytestring=svg.encode()))).convert("RGBA")
        cell=Image.new("RGB",(CELL_PX,CELL_PX),rgb)
        o=(CELL_PX-ICON_PX)//2; cell.paste(icon,(o,o),icon)
        cell.save(os.path.join(OUT_DIR,"func_{}.bmp".format(code.replace(".","_"))),"BMP")
        print("ok",code)
    print("done",OUT_DIR)

if __name__=="__main__": main()
