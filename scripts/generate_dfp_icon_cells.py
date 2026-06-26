# -*- coding: utf-8 -*-
"""Generate per-function DFP cell BMPs from Lucide icon nodes (Hub parity).

Run once after updating ``lib/cde/assets/dfp_icon_nodes.json``::

    python scripts/generate_dfp_icon_cells.py
"""
from __future__ import print_function

import json
import math
import os
import xml.etree.ElementTree as ET

import cairosvg
from PIL import Image

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
NODES_PATH = os.path.join(ROOT, "lib", "cde", "assets", "dfp_icon_nodes.json")
OUT_DIR = os.path.join(ROOT, "lib", "cde", "assets", "dfp_cells")

CELL_PX = 20
ICON_PX = 14
STROKE = "#ffffff"

# Hub DoorMarkersOverlay CATEGORY_HUE (degrees)
CATEGORY_HUE = {
    1: 4,
    2: 28,
    3: 268,
    4: 210,
    5: 150,
    6: 320,
    7: 190,
}

# code -> category (mirrors nobel-project-hub dfp-catalog.ts)
CODE_CATEGORY = {
    "1.01": 1, "1.02": 1, "1.03": 1, "1.04": 1, "1.05": 1,
    "1.07": 1, "1.08": 1, "1.09": 1, "1.10": 1, "1.11": 1,
    "2.1": 2, "2.2": 2, "2.3": 2, "2.4": 2, "2.6": 2, "2.7": 2,
    "2.8": 2, "2.9": 2,
    "3.1": 3, "3.2": 3, "3.3": 3,
    "4.1": 4, "4.2": 4, "4.3": 4, "4.7": 4,
    "5.1": 5, "5.2": 5, "5.3": 5, "5.4": 5,
    "6.1": 6, "6.2": 6, "6.3": 6, "6.4": 6, "6.5": 6,
    "7.1": 7, "7.2": 7, "7.3": 7, "7.4": 7,
}


def _hsl_to_rgb(h, s_pct, l_pct):
    s = s_pct / 100.0
    l = l_pct / 100.0
    c = (1.0 - abs(2.0 * l - 1.0)) * s
    hp = (h % 360) / 60.0
    x = c * (1.0 - abs(hp % 2.0 - 1.0))
    if hp < 1:
        r1, g1, b1 = c, x, 0
    elif hp < 2:
        r1, g1, b1 = x, c, 0
    elif hp < 3:
        r1, g1, b1 = 0, c, x
    elif hp < 4:
        r1, g1, b1 = 0, x, c
    elif hp < 5:
        r1, g1, b1 = x, 0, c
    else:
        r1, g1, b1 = c, 0, x
    m = l - c / 2.0
    return (
        int(round((r1 + m) * 255)),
        int(round((g1 + m) * 255)),
        int(round((b1 + m) * 255)),
    )


def _category_rgb(code):
    cat = CODE_CATEGORY.get(code, 1)
    hue = CATEGORY_HUE.get(cat, 220)
    return _hsl_to_rgb(hue, 64, 45)


def _attrs_to_str(attrs):
    parts = []
    for key, val in sorted(attrs.items()):
        if key == "key":
            continue
        parts.append('{}="{}"'.format(key, val))
    return " ".join(parts)


def _node_to_svg_elements(node):
    parts = []
    for item in node:
        tag = item[0]
        attrs = item[1] if len(item) > 1 else {}
        if tag == "path":
            parts.append('<path {} fill="none" stroke="{}" stroke-width="2" '
                         'stroke-linecap="round" stroke-linejoin="round"/>'.format(
                             _attrs_to_str(attrs), STROKE))
        elif tag == "circle":
            parts.append('<circle {} fill="none" stroke="{}" stroke-width="2"/>'.format(
                _attrs_to_str(attrs), STROKE))
        elif tag == "line":
            parts.append('<line {} stroke="{}" stroke-width="2" '
                         'stroke-linecap="round"/>'.format(
                             _attrs_to_str(attrs), STROKE))
        elif tag == "rect":
            parts.append('<rect {} fill="none" stroke="{}" stroke-width="2"/>'.format(
                _attrs_to_str(attrs), STROKE))
        elif tag == "polyline":
            parts.append('<polyline {} fill="none" stroke="{}" stroke-width="2" '
                         'stroke-linecap="round" stroke-linejoin="round"/>'.format(
                             _attrs_to_str(attrs), STROKE))
        elif tag == "polygon":
            parts.append('<polygon {} fill="none" stroke="{}" stroke-width="2" '
                         'stroke-linecap="round" stroke-linejoin="round"/>'.format(
                             _attrs_to_str(attrs), STROKE))
    return "\n    ".join(parts)


def _render_icon_png(node):
    inner = _node_to_svg_elements(node)
    svg = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" '
        'width="{size}" height="{size}">'
        '  {inner}'
        '</svg>'
    ).format(size=ICON_PX, inner=inner)
    return cairosvg.svg2png(bytestring=svg.encode("utf-8"))


def _code_filename(code):
    safe = code.replace(".", "_")
    return "func_{}.bmp".format(safe)


def main():
    with open(NODES_PATH, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    if not os.path.isdir(OUT_DIR):
        os.makedirs(OUT_DIR)

    chroma = (0, 128, 128)
    count = 0
    for code, entry in sorted(data.items()):
        if code == "fallback":
            continue
        rgb = _category_rgb(code)
        png_bytes = _render_icon_png(entry["node"])
        icon = Image.open(__import__("io").BytesIO(png_bytes)).convert("RGBA")

        cell = Image.new("RGB", (CELL_PX, CELL_PX), rgb)
        ox = (CELL_PX - ICON_PX) // 2
        oy = (CELL_PX - ICON_PX) // 2
        cell.paste(icon, (ox, oy), icon)

        out_path = os.path.join(OUT_DIR, _code_filename(code))
        cell.save(out_path, format="BMP")
        count += 1
        print("Wrote {}".format(out_path))

    # Fallback icon (neutral gray)
    fb = data.get("fallback")
    if fb:
        png_bytes = _render_icon_png(fb["node"])
        icon = Image.open(__import__("io").BytesIO(png_bytes)).convert("RGBA")
        cell = Image.new("RGB", (CELL_PX, CELL_PX), (100, 100, 110))
        ox = (CELL_PX - ICON_PX) // 2
        oy = (CELL_PX - ICON_PX) // 2
        cell.paste(icon, (ox, oy), icon)
        cell.save(os.path.join(OUT_DIR, "func_fallback.bmp"), format="BMP")
        count += 1

    # Unset-door placeholder (collapsed summary when no DFP keys authored)
    from PIL import ImageDraw
    clear_cell = Image.new("RGB", (CELL_PX, CELL_PX), (72, 76, 88))
    draw = ImageDraw.Draw(clear_cell)
    draw.ellipse([3, 3, CELL_PX - 4, CELL_PX - 4], outline=(255, 255, 255), width=1)
    clear_path = os.path.join(OUT_DIR, "func_clear.bmp")
    clear_cell.save(clear_path, format="BMP")
    count += 1
    print("Wrote {}".format(clear_path))

    print("Generated {} cell bitmap(s) in {}".format(count, OUT_DIR))


if __name__ == "__main__":
    main()
