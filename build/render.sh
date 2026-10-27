#!/usr/bin/env bash
# Rebuild image assets + the V2 deck, then render a PDF and per-slide previews.
# Usage: build/render.sh            (run from anywhere)
# Needs: python3 with python-pptx, matplotlib, networkx, qrcode, Pillow; LibreOffice + pdftoppm for previews.
set -euo pipefail
cd "$(dirname "$0")/.."
CACHE=build/.cache
MEDIA="$CACHE/v1media"
if [ ! -d "$MEDIA" ]; then
  mkdir -p "$CACHE/v1"
  python3 -c "import zipfile,sys; zipfile.ZipFile(sys.argv[1]).extractall(sys.argv[2])" \
    source/FlyGM_AIAS2026_Oral_V1.pptx "$CACHE/v1"
  mv "$CACHE/v1/ppt/media" "$MEDIA"
fi
python3 -I build/scripts/make_assets.py "$MEDIA" source/overleaf/figures build/assets
python3 -I build/scripts/build_deck.py source/FlyGM_AIAS2026_Oral_V1.pptx build/assets "$MEDIA" \
  source/overleaf/figures build/data output/FlyGM_AIAS2026_Oral_V2.pptx

# previews (LibreOffice rendering: fonts/charts are close to, not identical with, PowerPoint)
SOFFICE=${SOFFICE:-soffice}
PREV=output/preview
rm -rf "$PREV" && mkdir -p "$PREV"
cp output/FlyGM_AIAS2026_Oral_V2.pptx "$CACHE/deck.pptx"
if [ -n "${PPTX_SKILL:-}" ]; then
  (cd "$CACHE" && timeout 300 python "$PPTX_SKILL/scripts/office/soffice.py" --headless --convert-to pdf deck.pptx >/dev/null 2>&1)
else
  (cd "$CACHE" && timeout 300 "$SOFFICE" --headless --convert-to pdf deck.pptx >/dev/null 2>&1)
fi
cp "$CACHE/deck.pdf" output/FlyGM_AIAS2026_Oral_V2_preview.pdf
pdftoppm -jpeg -r 110 output/FlyGM_AIAS2026_Oral_V2_preview.pdf "$PREV/slide"
ls "$PREV" | wc -l
