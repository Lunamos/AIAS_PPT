"""Generate image assets for the FlyGM oral deck.

Usage: python make_assets.py V1_MEDIA_DIR PAPER_FIG_DIR OUT_DIR
"""
import sys, os
import numpy as np
from PIL import Image
import qrcode
import networkx as nx
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

media, figs, out = sys.argv[1:4]
os.makedirs(out, exist_ok=True)

INK = "#1F1A2E"
COL = {"flygm": "#D55E00", "rewired": "#009E73", "random": "#666666", "mlp": "#0072B2"}

# 1. QR code for the demo page
qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=20, border=2)
qr.add_data("https://lnsgroup.cc/research/FlyGM/")
qr.make(fit=True)
qr.make_image(fill_color=INK, back_color="white").convert("RGB").save(os.path.join(out, "qr_flygm.png"))

# 2. Force-directed connectome layout on transparent background (for the dark title slide)
im = np.asarray(Image.open(os.path.join(media, "image7.png")).convert("RGB")).astype(np.float32) / 255
white = im.min(axis=2)                       # 1 for pure white background
alpha = np.clip((1 - white) * 1.6, 0, 1)
rgb = np.clip((im - white[..., None]) / np.maximum(1 - white[..., None], 1e-3), 0, 1)  # un-premultiply
rgb = np.clip(rgb * 1.15 + 0.08, 0, 1)
Image.fromarray((np.dstack([rgb, alpha]) * 255).astype(np.uint8), "RGBA").save(os.path.join(out, "connectome_layout_alpha.png"))
# same layout for light backgrounds: transparent, colors un-premultiplied but not brightened
rgb_l = np.clip((im - white[..., None]) / np.maximum(1 - white[..., None], 1e-3), 0, 1)
alpha_l = np.clip((1 - white) * 1.25, 0, 1)
Image.fromarray((np.dstack([rgb_l, alpha_l]) * 255).astype(np.uint8), "RGBA").save(os.path.join(out, "connectome_layout_alpha_light.png"))

# 3. Panel (a) of the connectome structure figure
net = Image.open(os.path.join(figs, "network.png")).convert("RGB")
w, h = net.size
panel = net.crop((int(w * 0.005), 0, int(w * 0.565), h))
# blank out the panel letter "a" in the top-left corner
from PIL import ImageDraw
ImageDraw.Draw(panel).rectangle([0, 0, int(w * 0.045), int(h * 0.075)], fill="white")
panel.save(os.path.join(out, "network_panel_a.png"))

# 4. Controller glyphs: same 14 nodes, different wiring
rng = np.random.default_rng(3)
tiers = {"a": [0, 1, 2, 3], "i": [4, 5, 6, 7, 8, 9], "e": [10, 11, 12, 13]}
pos = {}
for k, (y, ids) in enumerate([(1.0, tiers["a"]), (0.5, tiers["i"]), (0.0, tiers["e"])]):
    xs = np.linspace(0.08, 0.92, len(ids))
    for x, n in zip(xs, ids):
        pos[n] = (x, y + (0.07 if (n % 2 and ids is tiers["i"]) else 0))
bio = nx.DiGraph()
bio.add_nodes_from(range(14))
edges = [(0, 4), (0, 5), (1, 5), (1, 4), (2, 8), (3, 9), (2, 9), (3, 8),
         (4, 5), (5, 6), (6, 4), (8, 9), (9, 7), (7, 8), (5, 7), (6, 10), (6, 11),
         (7, 12), (7, 13), (4, 10), (9, 13), (11, 6), (12, 7), (5, 11)]
bio.add_edges_from(edges)
rew = bio.copy()
nx.directed_edge_swap(rew, nswap=40, max_tries=5000, seed=7)
er = nx.gnm_random_graph(14, bio.number_of_edges(), seed=11, directed=True)


def draw(G, name, color, layered=False):
    fig = plt.figure(figsize=(2, 1.6), dpi=200)
    ax = fig.add_axes([0.02, 0.02, 0.96, 0.96]); ax.axis("off")
    P = pos
    if layered:
        layers = [4, 5, 5, 3]
        P, G = {}, nx.DiGraph()
        ids = []
        n = 0
        for li, k in enumerate(layers):
            ys = np.linspace(0.05, 0.95, k)
            cur = []
            for y in ys:
                P[n] = (0.06 + li * 0.88 / (len(layers) - 1), y); cur.append(n); n += 1
            if ids:
                G.add_edges_from((a, b) for a in ids[-1] for b in cur)
            ids.append(cur)
    for a, b in G.edges():
        (x1, y1), (x2, y2) = P[a], P[b]
        rad = 0.0 if layered else 0.18
        ax.annotate("", xy=(x2, y2), xytext=(x1, y1),
                    arrowprops=dict(arrowstyle="-", color=color, alpha=0.45 if not layered else 0.3,
                                    lw=0.9, shrinkA=3, shrinkB=3, connectionstyle=f"arc3,rad={rad}"))
    xs = [P[n][0] for n in G.nodes()]; ys = [P[n][1] for n in G.nodes()]
    ax.scatter(xs, ys, s=46, c=color, edgecolors="white", linewidths=1.0, zorder=3)
    ax.set_xlim(-0.04, 1.04); ax.set_ylim(-0.1, 1.14)
    fig.savefig(os.path.join(out, f"glyph_{name}.png"), transparent=True)
    plt.close(fig)


draw(bio, "flygm", COL["flygm"])
draw(rew, "rewired", COL["rewired"])
draw(er, "random", COL["random"])
draw(None, "mlp", COL["mlp"], layered=True)
print("assets written to", out)
