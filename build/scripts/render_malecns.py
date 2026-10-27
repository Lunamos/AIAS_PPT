"""Render the male CNS (Berg et al., Cell 2026) from the public FlyEM v1.0 release.

Inputs (downloaded from gs://flyem-male-cns via storage.googleapis.com):
  shell_brain.ngmesh   rois/brain-shell-v2.2/mesh/brain-shell.ngmesh
  shell_vnc.ngmesh     rois/vnc-shell-v2/mesh/vnc-shell.ngmesh
  mesh_CV.ngmesh       rois/malecns-major-compartments-v2/mesh/CV.ngmesh (neck connective)
  sel.tsv              sampled bodyIds + superclass (v1.0 body annotations)
  swc/<bodyId>.swc     v1.0 skeletons (8 nm units)
Usage: render_malecns.py SRC_DIR OUT.png [N_PER_CLASS]
Writes OUT.png (tissue + neurons, transparent background).
"""
import sys, os
import numpy as np
import pandas as pd
from scipy import ndimage
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.collections import LineCollection
from PIL import Image

src, out = sys.argv[1], sys.argv[2]
nper = int(sys.argv[3]) if len(sys.argv) > 3 else 70

# data axes: x left-right, y dorsal->ventral, z anterior->posterior (nm)
# dorsal view, anterior to the left: screen u = z, v = -x, depth w = -y (toward viewer = dorsal)
R = np.array([[0, 0, 1], [-1, 0, 0], [0, -1, 0]], float)
center = np.array([384000, 300000, 578000.0])


def load_mesh(fn):
    b = open(fn, "rb").read()
    n = np.frombuffer(b[:4], np.uint32)[0]
    v = np.frombuffer(b[4:4 + 12 * n], np.float32).reshape(-1, 3).astype(np.float64)
    f = np.frombuffer(b[4 + 12 * n:], np.uint32).reshape(-1, 3)
    return v, f


def proj(p):
    return (p - center) @ R.T


meshes = [load_mesh(os.path.join(src, fn)) for fn in ["shell_brain.ngmesh", "shell_vnc.ngmesh", "mesh_CV.ngmesh"]]
allq = np.concatenate([proj(v) for v, _ in meshes])
lo, hi = allq[:, :2].min(0), allq[:, :2].max(0)
pad = 0.02 * (hi - lo).max()
lo -= pad; hi += pad
WPX = 3200
scale = WPX / (hi[0] - lo[0])
HPX = int(np.ceil((hi[1] - lo[1]) * scale))


def to_px(q):
    return (q[:, 0] - lo[0]) * scale, (hi[1] - q[:, 1]) * scale


# ---- software z-buffer of the tissue shells
zbuf = np.full((HPX, WPX), -np.inf)
nbuf = np.zeros((HPX, WPX, 3))
for v, f in meshes:
    q = proj(v)
    a, b, c = q[f[:, 0]], q[f[:, 1]], q[f[:, 2]]
    n = np.cross(b - a, c - a)
    n /= np.linalg.norm(n, axis=1, keepdims=True) + 1e-12
    n[n[:, 2] < 0] *= -1  # orientation-agnostic: face the viewer
    # barycentric samples per triangle
    W = np.array([[1, 0, 0], [0, 1, 0], [0, 0, 1], [1/3, 1/3, 1/3], [.5, .5, 0], [0, .5, .5], [.5, 0, .5],
                  [2/3, 1/6, 1/6], [1/6, 2/3, 1/6], [1/6, 1/6, 2/3]])
    for w in W:
        p = w[0] * a + w[1] * b + w[2] * c
        px, py = to_px(p)
        ix = np.clip(np.round(px).astype(int), 0, WPX - 1)
        iy = np.clip(np.round(py).astype(int), 0, HPX - 1)
        lin = iy * WPX + ix
        z = p[:, 2]
        order = np.argsort(z)  # later (closer) writes win
        flat_z = zbuf.reshape(-1); flat_n = nbuf.reshape(-1, 3)
        lo_ = lin[order]; zz = z[order]; nn = n[order]
        closer = zz > flat_z[lo_]
        # assign in sorted order; duplicates resolved by last (largest z)
        flat_z[lo_[closer]] = zz[closer]
        flat_n[lo_[closer]] = nn[closer]

mask = np.isfinite(zbuf)
mask_f = ndimage.binary_closing(mask, iterations=3)
mask_f = ndimage.binary_fill_holes(mask_f)
# fill missing normals from neighbours
for k in range(3):
    ch = nbuf[..., k]
    filled = ndimage.grey_dilation(np.where(mask, ch, -9), size=3)
    nbuf[..., k] = np.where(mask, ch, np.where(filled > -9, filled, 0))
nrm = nbuf / (np.linalg.norm(nbuf, axis=2, keepdims=True) + 1e-9)
nrm = ndimage.gaussian_filter(nrm, sigma=(2.2, 2.2, 0))
nrm /= np.linalg.norm(nrm, axis=2, keepdims=True) + 1e-9
light = np.array([-0.45, 0.55, 0.70]); light /= np.linalg.norm(light)
diff = np.clip(nrm @ light, 0, 1)
rim = 1 - np.clip(nrm[..., 2], 0, 1)          # silhouette darkening
lum = 0.70 + 0.30 * diff - 0.12 * rim
base = np.array([0.86, 0.84, 0.91])
rgb = np.clip(lum[..., None] * base[None, None, :], 0, 1)
alpha = ndimage.gaussian_filter(mask_f.astype(float), 0.8)
tissue = np.dstack([rgb, alpha * 0.80])
Image.fromarray((tissue * 255).astype(np.uint8), "RGBA").save(out.replace(".png", "_tissue.png"))

# ---- neurons
sel = pd.read_csv(os.path.join(src, "sel.tsv"), sep="\t", header=None, names=["bodyId", "sc"])
style = {
    "cb_intrinsic": ("#8E8A99", 0.25, 0.30),
    "vnc_intrinsic": ("#8E8A99", 0.25, 0.30),
    "vnc_motor": ("#7B3FA0", 0.40, 0.80),
    "ascending_neuron": ("#2B7BBF", 0.40, 0.80),
    "descending_neuron": ("#E06A1B", 0.40, 0.85),
}
counts = {"cb_intrinsic": nper, "vnc_intrinsic": nper, "vnc_motor": nper // 2,
          "ascending_neuron": int(nper * 0.9), "descending_neuron": int(nper * 0.6)}
dpi = 200
fig = plt.figure(figsize=(WPX / dpi, HPX / dpi), dpi=dpi)
ax = fig.add_axes([0, 0, 1, 1]); ax.set_xlim(0, WPX); ax.set_ylim(HPX, 0); ax.axis("off")
for sc in ["cb_intrinsic", "vnc_intrinsic", "vnc_motor", "ascending_neuron", "descending_neuron"]:
    segs = []
    for bid in sel[sel.sc == sc].bodyId.head(counts[sc]):
        try:
            s = np.loadtxt(os.path.join(src, "swc", f"{bid}.swc"), ndmin=2)
        except Exception:
            continue
        if len(s) < 2:
            continue
        idx = {int(r[0]): i for i, r in enumerate(s)}
        px, py = to_px(proj(s[:, 2:5] * 8.0))
        for i, p in enumerate(s[:, 6].astype(int)):
            j = idx.get(p)
            if j is not None:
                segs.append(((px[i], py[i]), (px[j], py[j])))
    c, lw, al = style[sc]
    ax.add_collection(LineCollection(segs, colors=c, linewidths=lw * 1.7, alpha=al))
fig.savefig(out.replace(".png", "_neurons.png"), transparent=True, dpi=dpi)
plt.close(fig)

t = Image.open(out.replace(".png", "_tissue.png")).convert("RGBA")
ne = Image.open(out.replace(".png", "_neurons.png")).convert("RGBA").resize(t.size, Image.LANCZOS)
t.alpha_composite(ne)
bbox = t.getbbox()
t = t.crop(bbox)
t.save(out)
print("saved", out, t.size)
