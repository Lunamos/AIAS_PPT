"""Recover the imitation-learning curves of paper Fig. 2 from its vector PDF.

Usage: python extract_fig2_curves.py source/overleaf/figures/bias.pdf build/data/fig2_curves.json
Axis calibration (tick positions in PDF points) was read from the figure's own tick marks.
"""
import json, math, sys
import pymupdf as fitz

src, out = sys.argv[1:3]
names = {(0.835, 0.369, 0.0): "FlyGM", (0.0, 0.447, 0.698): "MLP", (0.0, 0.62, 0.451): "Rewired", (0.4, 0.4, 0.4): "ER-Random"}
res = {"a": {}, "b": {}}
for d in fitz.open(src)[0].get_drawings():
    if d["type"] != "s" or len(d["items"]) < 100:
        continue
    nm = names[tuple(round(v, 3) for v in d["color"])]
    pts = []
    for it in d["items"]:
        if not pts:
            pts.append((it[1].x, it[1].y))
        pts.append((it[2].x, it[2].y) if it[0] == "l" else (it[4].x, it[4].y))
    if pts[0][0] < 260:   # panel a: action-mean MSE, x 0->42.84pt, 20k->215.84pt; y 0.03->134.78pt, 3->47.43pt
        xs = [(px - 42.84) / (215.84 - 42.84) * 20 for px, _ in pts]
        ys = [10 ** (math.log10(0.03) + (134.78 - py) / (134.78 - 47.43) * 2) for _, py in pts]
        res["a"][nm] = {"x": xs, "y": ys}
    else:                 # panel b: training objective, y 2->130.27pt, 100->42.87pt
        xs = [(px - 297.49) / (470.49 - 297.49) * 20 for px, _ in pts]
        ys = [10 ** (math.log10(2) + (130.27 - py) / (130.27 - 42.87) * math.log10(50)) for _, py in pts]
        res["b"][nm] = {"x": xs, "y": ys}
json.dump(res, open(out, "w"))
