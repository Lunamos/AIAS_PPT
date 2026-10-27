#!/usr/bin/env bash
# Download the public FlyEM male CNS v1.0 data used for the outlook-slide rendering
# (Berg et al., Cell 2026; gs://flyem-male-cns, served over HTTPS by storage.googleapis.com).
# Usage: build/scripts/fetch_malecns.sh OUT_DIR   (~380 MB)
# Then:  python3 build/scripts/render_malecns.py OUT_DIR build/assets/malecns_dorsal.png 70
set -euo pipefail
OUT="$1"; mkdir -p "$OUT/swc"
B=https://storage.googleapis.com/flyem-male-cns
curl -sS -o "$OUT/shell_brain.ngmesh" "$B/rois/brain-shell-v2.2/mesh/brain-shell.ngmesh"
curl -sS -o "$OUT/shell_vnc.ngmesh"   "$B/rois/vnc-shell-v2/mesh/vnc-shell.ngmesh"
curl -sS -o "$OUT/mesh_CV.ngmesh"     "$B/rois/malecns-major-compartments-v2/mesh/CV.ngmesh"
curl -sS -o "$OUT/ann.feather"        "$B/v1.0/connectome-data/flat-connectome/body-annotations-male-cns-v1.0-minconf-0.5.feather"
# neuron sample used in the deck (random_state=0 per superclass; see build/data/malecns_sampled_bodyids.tsv)
cp "$(dirname "$0")/../data/malecns_sampled_bodyids.tsv" "$OUT/sel.tsv"
cut -f1 "$OUT/sel.tsv" | xargs -P 16 -I{} curl -sS -o "$OUT/swc/{}.swc" "$B/v1.0/segmentation/skeletons-malecns/skeletons-swc/{}.swc"
