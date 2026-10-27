"""Build the FlyGM AIAS+ 2026 oral deck (V2) on top of the V1 deck.

V1 is kept as the package base so that its videos (walking, flight, synced rollout)
and the polished rollout / latent-response slides survive unchanged; every other
slide is rebuilt from scratch. All slides use a white background.

Every text box is checked against Arial-metric fonts (Liberation Sans) and the
build prints a warning when text would overflow its box.

Usage:
  python build_deck.py V1.pptx ASSET_DIR V1_MEDIA_DIR PAPER_FIG_DIR DATA_DIR OUT.pptx
"""
import json
import math
import os
import re
import sys

from lxml import etree
from PIL import ImageFont
from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_MARKER_STYLE, XL_TICK_MARK
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

V1, ASSETS, MEDIA, FIGS, DATA, OUT = sys.argv[1:7]

# ---------------------------------------------------------------- palette
INK = "1F1A2E"        # titles, primary text
BODY = "3A3447"       # body text
MUTED = "6E6880"      # captions, footers
PURPLE = "660874"     # Tsinghua purple: kickers, structure, key boxes
VERM = "D55E00"       # FlyGM (Okabe-Ito vermillion, as in the paper's Fig. 2)
GREEN = "009E73"      # degree-preserving rewiring
GREY = "666666"       # Erdos-Renyi random graph
BLUE = "0072B2"       # MLP
TINT = "F4F1F8"       # card fill
TINT_V = "FBEDE3"     # FlyGM card fill
LINE = "E2DCEA"       # hairlines, gridlines
PEACH = "F6C7A5"      # light accent text on purple
AFF, INT_, EFF = "2E75B6", "4E9F5A", "E07B39"   # flow types in the paper's figures
FONT = "Arial"

TITLE_PT = 28
FOOT_PT = 10

LIB = "/usr/share/fonts/truetype/liberation/"
_fonts = {}


def _font(size, bold):
    key = (round(size * 4), bold)
    if key not in _fonts:
        _fonts[key] = ImageFont.truetype(LIB + ("LiberationSans-Bold.ttf" if bold else "LiberationSans-Regular.ttf"),
                                         max(1, int(round(size * 4))))
    return _fonts[key]


def text_w(text, size, bold=False):
    """Width in inches of a single line of Arial text."""
    return _font(size, bold).getlength(text) / 4 / 72


WARNINGS = []


def rgb(h):
    return RGBColor.from_string(h)


def A(name):
    return os.path.join(ASSETS, name)


def M(name):
    return os.path.join(MEDIA, name)


# ---------------------------------------------------------------- text helpers
def _apply_run(run, size=None, color=None, bold=None, italic=None, spc=None, baseline=None):
    f = run.font
    if size is not None:
        f.size = Pt(size)
    if color is not None:
        f.color.rgb = rgb(color)
    if bold is not None:
        f.bold = bold
    if italic is not None:
        f.italic = italic
    f.name = FONT
    rPr = run._r.get_or_add_rPr()
    if spc is not None:
        rPr.set("spc", str(int(spc)))
    if baseline is not None:
        rPr.set("baseline", str(int(baseline)))


def _bullet(p, color, indent=0.2):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(Inches(indent)))
    pPr.set("indent", str(-Inches(indent)))
    bc = etree.SubElement(pPr, qn("a:buClr"))
    etree.SubElement(bc, qn("a:srgbClr")).set("val", color)
    etree.SubElement(pPr, qn("a:buFont")).set("typeface", FONT)
    etree.SubElement(pPr, qn("a:buChar")).set("char", "•")


def _fit_check(label, w, h, paras, line):
    """Greedy word wrap with Arial metrics; record a warning if the text overflows."""
    total = 0.0
    for runs, popts in paras:
        size_max = max(o["size"] for _, o in runs) if runs else 12
        avail = w - (popts.get("indent", 0.2) if popts.get("bullet") else 0)
        lines, cur = 1, 0.0
        for text, o in runs:
            for k, word in enumerate(text.split(" ")):
                piece = (" " if (k > 0 or cur > 0) and word else "") + word
                ww = text_w(piece, o["size"], o["bold"]) * (1 + (o.get("spc") or 0) / 1000)
                if cur + ww > avail + 1e-6 and cur > 0:
                    lines += 1
                    cur = text_w(word, o["size"], o["bold"])
                else:
                    cur += ww
        total += lines * size_max * 1.17 * popts.get("line", line) / 72
        total += (popts.get("after", 0) + popts.get("before", 0)) / 72
    if total > h + 0.03:
        WARNINGS.append(f"OVERFLOW {label}: needs {total:.2f}in, box {h:.2f}in")


def _nbsp_cites(text):
    text = re.sub(r" \[(?=[0-9])", "\u00a0[", text)
    return re.sub(r"\[[0-9,–\s]+\]", lambda m: m.group(0).replace(" ", "\u00a0"), text)


def tb(slide, x, y, w, h, paras, size=14, color=BODY, bold=False, italic=False, align="l", anchor="t",
       line=1.05, after=0, name=None, bullet=None, check=True):
    """Text box. `paras`: str | list of paragraphs; a paragraph is str | list of (text, opts) |
    dict(runs=..., bullet=..., after=..., line=..., align=...)."""
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        shape.name = name
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    if isinstance(paras, str):
        paras = [paras]
    norm = []
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        popts = {}
        if isinstance(para, dict):
            popts = dict(para)
            para = popts.pop("runs")
        if isinstance(para, str):
            para = [(para, {})]
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[popts.get("align", align)]
        p.line_spacing = popts.get("line", line)
        if popts.get("after", after):
            p.space_after = Pt(popts.get("after", after))
        if popts.get("before"):
            p.space_before = Pt(popts["before"])
        b = popts.get("bullet", bullet)
        if b:
            _bullet(p, b, popts.get("indent", 0.2))
            popts["bullet"] = b
        popts.setdefault("after", after)
        runs = []
        for text, o in para:
            text = _nbsp_cites(text)
            r = p.add_run()
            r.text = text
            o2 = {"size": o.get("size", popts.get("size", size)), "bold": o.get("bold", popts.get("bold", bold)),
                  "spc": o.get("spc")}
            _apply_run(r, size=o2["size"], color=o.get("color", popts.get("color", color)), bold=o2["bold"],
                       italic=o.get("italic", italic), spc=o.get("spc"), baseline=o.get("baseline"))
            if o.get("link"):
                r.hyperlink.address = o["link"]
            runs.append((text, o2))
        norm.append((runs, popts))
    if check:
        first = norm[0][0][0][0][:30] if norm and norm[0][0] else ""
        _fit_check(f"slide {len(prs.slides._sldIdLst)} '{first}'", w, h, norm, line)
    return shape


def _drop_style(shape):
    st = shape._element.find(qn("p:style"))
    if st is not None:
        shape._element.remove(st)


def box(slide, x, y, w, h, fill=TINT, radius=0.08, line=None, line_w=1.0, shape=MSO_SHAPE.ROUNDED_RECTANGLE, name=None):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        s.name = name
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = rgb(fill)
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = rgb(line)
        s.line.width = Pt(line_w)
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = radius
    _drop_style(s)
    return s


def icon_circle(slide, x, y, d, color, icon, pad=0.2):
    box(slide, x, y, d, d, fill=color, shape=MSO_SHAPE.OVAL, name="Icon circle")
    slide.shapes.add_picture(A(os.path.join("icons", icon)), Inches(x + d * pad), Inches(y + d * pad),
                             Inches(d * (1 - 2 * pad)), Inches(d * (1 - 2 * pad))).name = "Icon"


def num_circle(slide, x, y, d, color, n, size=14):
    box(slide, x, y, d, d, fill=color, shape=MSO_SHAPE.OVAL)
    tb(slide, x, y, d, d, [[(str(n), {"bold": True})]], size=size, color="FFFFFF", align="c", anchor="m", check=False)


def chip(slide, x, y, parts, size=12, fill=TINT, line=None, h=0.38, pad=0.32):
    """Rounded pill sized to its text. parts: list of (text, color, bold). Returns its width."""
    w = sum(text_w(t, size, b) for t, _, b in parts) + pad
    box(slide, x, y, w, h, fill=fill, radius=0.5, line=line, line_w=1.25)
    tb(slide, x, y, w, h, [[(t, {"color": c, "bold": b}) for t, c, b in parts]], size=size, align="c", anchor="m",
       check=False)
    return w


def arrow(slide, x1, y1, x2, y2, color=MUTED, width=1.5, head=True):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = rgb(color)
    c.line.width = Pt(width)
    _drop_style(c)
    if head:
        te = etree.SubElement(c.line._get_or_add_ln(), qn("a:tailEnd"))
        te.set("type", "triangle")
        te.set("w", "med")
        te.set("len", "med")
    return c


def pic(slide, path, x, y, w=None, h=None, name=None, link=None):
    p = slide.shapes.add_picture(path, Inches(x), Inches(y), Inches(w) if w else None, Inches(h) if h else None)
    if name:
        p.name = name
    if link:
        p.click_action.hyperlink.address = link
    return p


def remove(shape):
    el = shape._element
    el.getparent().remove(el)


# ---------------------------------------------------------------- slide frame
def kicker(slide, text):
    box(slide, 0.5, 0.3, 0.1, 0.1, fill=VERM, shape=MSO_SHAPE.OVAL, name="Kicker dot")
    tb(slide, 0.68, 0.24, 7.0, 0.22, [[(text.upper(), {"spc": 120, "bold": True})]], size=10.5, color=PURPLE,
       anchor="m", name="Kicker")


def set_title(slide, layout, text, size=TITLE_PT, color=INK, x=0.5, y=0.48, w=9.0, h=0.55):
    t = slide.shapes.title
    if t is None:
        lt = [ph for ph in layout.placeholders if ph.placeholder_format.type == PP_PLACEHOLDER.TITLE][0]
        slide.shapes.clone_placeholder(lt)
        t = slide.shapes.title
    t.left, t.top, t.width, t.height = Inches(x), Inches(y), Inches(w), Inches(h)
    tf = t.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.line_spacing = 1.0
    r = p.add_run()
    r.text = text
    _apply_run(r, size=size, color=color, bold=True)
    if text_w(text, size, True) > w:
        WARNINGS.append(f"TITLE WRAPS: {text!r} ({text_w(text, size, True):.2f}in > {w}in)")
    return t


def footer(slide, text):
    tb(slide, 0.5, 5.22, 8.3, 0.25, text, size=FOOT_PT, color=MUTED, anchor="m", name="Source")


def slide_number(slide):
    s = slide.shapes.add_textbox(Inches(9.0), Inches(5.22), Inches(0.5), Inches(0.25))
    s.name = "Slide number"
    tf = s.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    fld = etree.SubElement(p._p, qn("a:fld"))
    fld.set("id", "{5C1D9F7E-2B47-4C0B-9E55-6B0D1A3E1F01}")
    fld.set("type", "slidenum")
    rPr = etree.SubElement(fld, qn("a:rPr"))
    rPr.set("lang", "en-US")
    rPr.set("sz", str(FOOT_PT * 100))
    etree.SubElement(etree.SubElement(rPr, qn("a:solidFill")), qn("a:srgbClr")).set("val", MUTED)
    etree.SubElement(rPr, qn("a:latin")).set("typeface", FONT)
    etree.SubElement(fld, qn("a:t")).text = "#"


def strip_placeholders(slide, keep_title=True):
    for ph in list(slide.placeholders):
        if keep_title and ph.placeholder_format.type == PP_PLACEHOLDER.TITLE:
            continue
        remove(ph)


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


# ---------------------------------------------------------------- chart helpers
def chart_frame_clean(chart, size=11):
    cs = chart._chartSpace
    chart.font.name = FONT
    chart.font.size = Pt(size)
    chart.font.color.rgb = rgb(MUTED)
    chart.has_title = False
    chart.has_legend = False
    for parent in (cs, cs.find(".//" + qn("c:plotArea"))):
        old = parent.find(qn("c:spPr"))
        if old is not None:
            parent.remove(old)
        sp = etree.Element(qn("c:spPr"))
        etree.SubElement(sp, qn("a:noFill"))
        etree.SubElement(etree.SubElement(sp, qn("a:ln")), qn("a:noFill"))
        if parent is cs:
            cs.find(qn("c:chart")).addnext(sp)
        else:
            parent.append(sp)


def plot_layout(chart, x, y, w, h):
    pa = chart._chartSpace.find(".//" + qn("c:plotArea"))
    lay = pa.find(qn("c:layout"))
    if lay is None:
        lay = etree.Element(qn("c:layout"))
        pa.insert(0, lay)
    for c in list(lay):
        lay.remove(c)
    ml = etree.SubElement(lay, qn("c:manualLayout"))
    for tag, val in (("layoutTarget", "inner"), ("xMode", "edge"), ("yMode", "edge"),
                     ("x", x), ("y", y), ("w", w), ("h", h)):
        etree.SubElement(ml, qn("c:" + tag)).set("val", str(val))


def style_axis(ax, grid=True, line=True, fmt=None, size=11, color=MUTED):
    ax.has_major_gridlines = grid
    if grid:
        ax.major_gridlines.format.line.color.rgb = rgb(LINE)
        ax.major_gridlines.format.line.width = Pt(0.75)
    ax.has_minor_gridlines = False
    ax.major_tick_mark = XL_TICK_MARK.NONE
    ax.minor_tick_mark = XL_TICK_MARK.NONE
    if line:
        ax.format.line.color.rgb = rgb("B9B1C6")
        ax.format.line.width = Pt(0.75)
    else:
        ax.format.line.fill.background()
    ax.tick_labels.font.size = Pt(size)
    ax.tick_labels.font.color.rgb = rgb(color)
    ax.tick_labels.font.name = FONT
    if fmt:
        ax.tick_labels.number_format = fmt
        ax.tick_labels.number_format_is_linked = False


def axis_xml(ax, crosses_at=None, low_labels=False, log=False):
    el = ax._element
    if log:
        lb = etree.Element(qn("c:logBase"))
        lb.set("val", "10")
        el.find(qn("c:scaling")).insert(0, lb)
    if low_labels:
        tl = el.find(qn("c:tickLblPos"))
        if tl is None:
            tl = etree.Element(qn("c:tickLblPos"))
            el.find(qn("c:crossAx")).addprevious(tl)
        tl.set("val", "low")
    if crosses_at is not None:
        cr = el.find(qn("c:crosses"))
        ca = etree.Element(qn("c:crossesAt"))
        ca.set("val", str(crosses_at))
        if cr is not None:
            cr.addprevious(ca)
            el.remove(cr)
        else:
            el.find(qn("c:crossAx")).addnext(ca)


def point_label(pt, text, color, size=10.5, bold=True, pos=XL_LABEL_POSITION.OUTSIDE_END):
    dl = pt.data_label
    dl.has_text_frame = True
    dl.text_frame.text = text
    _apply_run(dl.text_frame.paragraphs[0].runs[0], size=size, color=color, bold=bold)
    dl.position = pos


# ================================================================ build
prs = Presentation(V1)
old = list(prs.slides)
LAY = prs.slide_layouts.get_by_name("1_Title and Content")
DEMO = "https://lnsgroup.cc/research/FlyGM/"


def new_slide():
    s = prs.slides.add_slide(LAY)
    strip_placeholders(s)
    return s


order = []

# ---------------------------------------------------------------- 1. title
s = new_slide()
pic(s, M("image7.png"), 5.25, 0.38, 4.5, 4.5, name="FlyGM connectome graph layout")
tb(s, 0.6, 0.72, 4.8, 0.26, [[("AIAS+ 2026  ·  ORAL PRESENTATION", {"spc": 150, "bold": True})]], size=11, color=VERM)
set_title(s, LAY, "FlyGM", size=66, color=PURPLE, x=0.6, y=1.0, w=4.6, h=1.05)
tb(s, 0.6, 2.1, 4.7, 1.02, "Whole-Brain Connectomic Graph Model for Whole-Body Locomotion Control",
   size=20, color=INK, line=1.02)
tb(s, 0.6, 3.22, 4.7, 0.3, "Zehao Jin, Yaoye Zhu, Chen Zhang, Yanan Sui", size=15, color=INK, bold=True)
tb(s, 0.6, 3.56, 4.7, 0.28, "Tsinghua University  ·  LNS Group", size=13, color=MUTED)
pic(s, M("image1.png"), 0.6, 4.08, 0.86, 0.86, name="Tsinghua University logo")
pic(s, M("image2.png"), 1.66, 4.3, 1.68, 0.42, name="LNS Group logo")
pic(s, A("qr_flygm.png"), 3.78, 4.02, 0.92, 0.92, name="QR code: lnsgroup.cc/research/FlyGM", link=DEMO)
tb(s, 3.62, 4.96, 1.24, 0.22, "Demo page", size=10, color=MUTED, align="c")
notes(s, """Target time: 30 seconds.

Good morning. Our work asks a simple question: can the wiring diagram of an animal's brain be used to build a controller for its body? We study the fruit fly, because it is the one animal for which we now have both a complete brain connectome and a detailed whole-body physics simulator. FlyGM turns that connectome into the architecture of a learned controller. I will explain why this question is timely, how FlyGM works, what the results say about the value of the real wiring, and where this could go next.""")
order.append(s)

# ---------------------------------------------------------------- 2. why now
s = new_slide()
kicker(s, "Why now")
set_title(s, LAY, "Fly brain and body are now both digital")
pic(s, M("image4.png"), 0.4, 1.15, 4.3, 4.3 * 600 / 768, name="FlyWire whole-brain rendering")
for i, (v, l) in enumerate([("139,255", "neurons"), ("54.5 M", "synapses"), ("8,453", "cell types")]):
    x = 0.5 + i * 1.45
    tb(s, x, 4.38, 1.42, 0.45, [[(v, {"bold": True})]], size=26, color=PURPLE)
    tb(s, x, 4.83, 1.42, 0.27, l, size=13, color=MUTED)
tl = [
    ("2024", PURPLE, "FlyWire whole-brain map", "Complete adult brain wiring [2, 3]"),
    ("2024", PURPLE, "Connectome-based models", "Wiring predicts neural activity [4, 5]"),
    ("2025", VERM, "flybody simulator", "A whole fly body in physics [6, 7]"),
    ("2026", VERM, "Male CNS connectome", "Brain + nerve cord, 166,700 neurons [8]"),
]
x0, y0, dy = 5.05, 1.15, 0.98
arrow(s, x0 + 0.36, y0 + 0.2, x0 + 0.36, y0 + 3 * dy + 0.2, color=LINE, width=2.5, head=False)
for i, (yr, col, head, desc) in enumerate(tl):
    y = y0 + i * dy
    box(s, x0, y + 0.04, 0.72, 0.32, fill=col, radius=0.5)
    tb(s, x0, y + 0.04, 0.72, 0.32, [[(yr, {"bold": True})]], size=12, color="FFFFFF", align="c", anchor="m")
    tb(s, x0 + 0.92, y, 3.55, 0.36, [[(head, {"bold": True})]], size=15.5, color=INK, anchor="m")
    tb(s, x0 + 0.92, y + 0.4, 3.55, 0.3, desc, size=13, color=BODY)
footer(s, "Rendering: FlyWire [2]. Counts: [2, 3, 8].")
slide_number(s)
notes(s, """Target time: 60 seconds.

Why now? In 2024 the FlyWire consortium published the first complete wiring diagram of an adult brain: an adult female fruit fly, 139,255 neurons, about 54.5 million synapses, and more than eight thousand annotated cell types. Within months, connectome-constrained models showed that this wiring can predict neural activity in the visual system and sensorimotor responses across the brain.

At the same time, the body became digital too. flybody, published in Nature in 2025, provides an anatomically detailed fly in the MuJoCo physics engine, with learned controllers that walk and fly; NeuroMechFly v2 is a complementary platform. And this year the male central nervous system connectome added the ventral nerve cord, the fly's equivalent of the spinal cord.

So, for the first time, one animal has both a complete brain map and a whole-body simulator. The obvious next question is how to connect them.""")
order.append(s)

# ---------------------------------------------------------------- 3. timeline and debate
s = new_slide()
kicker(s, "2026: the question goes public")
set_title(s, LAY, "Can a wiring diagram drive a body?")
cards = [
    (TINT_V, VERM, "SEP 2025  ·  OURS", "FlyGM paper drafted", VERM, "Connectome drives a whole-body fly [1]"),
    (TINT, PURPLE, "MAR 2026", "“We’ve uploaded a fruit fly”", INK, "Eon Systems demo goes viral [9]"),
    (TINT, PURPLE, "MAR 2026 →", "“No, we haven’t uploaded a fly yet”", INK, "Neuroscientists push back [10, 11]"),
]
cw, gap = 2.8, 0.3
for i, (fill, ccol, date, head, hcol, desc) in enumerate(cards):
    x = 0.5 + i * (cw + gap)
    box(s, x, 1.12, cw, 1.62, fill=fill, radius=0.07, line=(VERM if i == 0 else None), line_w=1.5)
    chip(s, x + 0.18, 1.26, [(date, "FFFFFF", True)], size=10.5, fill=ccol, h=0.3, pad=0.3)
    tb(s, x + 0.18, 1.66, cw - 0.32, 0.56, [[(head, {"bold": True})]], size=14.5, color=hcol, line=1.0)
    tb(s, x + 0.18, 2.24, cw - 0.32, 0.42, desc, size=12.5, color=BODY, line=1.0)
    if i < 2:
        arrow(s, x + cw + 0.05, 1.93, x + cw + gap - 0.05, 1.93, color="B9B1C6", width=2)
tb(s, 0.5, 2.95, 9.0, 0.25, [[("WHAT THE CRITICS POINT OUT", {"bold": True, "spc": 120})]], size=10.5, color=MUTED)
gaps = [
    ("tune_w.png", "Synapse strength", "Paths, not influence [12]"),
    ("chart_w.png", "Neuron dynamics", "Simplified neuron models [4]"),
    ("cable_w.png", "Body interface", "Hand-picked control neurons [9, 11]"),
]
for i, (ic, head, d) in enumerate(gaps):
    x = 0.5 + i * (cw + gap)
    icon_circle(s, x, 3.3, 0.56, PURPLE, ic)
    tb(s, x + 0.7, 3.27, cw - 0.72, 0.3, [[(head, {"bold": True})]], size=14.5, color=INK)
    tb(s, x + 0.7, 3.58, cw - 0.72, 0.5, d, size=12.5, color=BODY, line=1.0)
box(s, 0.5, 4.35, 9.0, 0.66, fill=PURPLE, radius=0.14)
tb(s, 0.75, 4.35, 8.5, 0.66, [[("Open question:  ", {"bold": True, "color": PEACH}),
                              ("how should a connectome control a whole body?", {"bold": True})]],
   size=17, color="FFFFFF", anchor="m")
footer(s, "FlyGM first draft completed 18 Sep 2025 [1]. Public discussion: [9–11].")
slide_number(s)
notes(s, """Target time: 60 seconds.

This question has recently become very public, and we were early. The first complete draft of the FlyGM paper, a whole-brain connectome driving a whole-body simulated fly, was finished in September 2025, half a year before the debate started.

In March 2026, the startup Eon Systems announced "We've uploaded a fruit fly": a connectome-based spiking brain model coupled to the NeuroMechFly body. The video of a virtual fly walking, grooming and feeding went viral, and neuroscientists pushed back quickly, for example in the essays "No, we haven't uploaded a fly yet" and "Deconstructing viral fly sims".

The critics point to three things the wiring diagram leaves out. Synapse counts tell us which paths exist, not how strongly neurons influence each other in the living animal. The neuron models are simplified. And the interface to the body: in that demo, a few hand-picked neurons trigger the body model's built-in walking and feeding controllers.

So the open question is not whether we have the wiring. It is how a connectome should control a whole body, and that is exactly the question FlyGM was built to answer.""")
order.append(s)

# ---------------------------------------------------------------- 4. our approach
s = new_slide()
kicker(s, "Our approach")
set_title(s, LAY, "FlyGM: wiring as the controller's architecture")
cols = [
    ("Emulate the wiring", "biotech_w.png", "8E86A0", TINT, M("image4.png"),
     [("+", "Real wiring, no training"), ("−", "Body moved by built-in controllers")], "e.g. [4, 9]"),
    ("Learn a generic net", "robot_w.png", "8E86A0", TINT, M("image5.png"),
     [("+", "Realistic walking and flight"), ("−", "Network unrelated to the brain")], "e.g. flybody [7]"),
    ("FlyGM (ours)", "hub_w.png", VERM, TINT_V, A("connectome_layout_alpha_light.png"),
     [("+", "The real wiring is the network"), ("+", "Learns whole-body control")], ""),
]
cw, gap = 2.86, 0.21
for i, (head, ic, col, fill, img, items, ref) in enumerate(cols):
    x = 0.5 + i * (cw + gap)
    box(s, x, 1.1, cw, 3.0, fill=fill, radius=0.06, line=(VERM if i == 2 else None), line_w=1.5)
    if i == 0:
        pic(s, img, x + 0.6, 1.2, 1.66, 1.3)
    elif i == 1:
        pic(s, img, x + 0.47, 1.2, 1.93, 1.3)
    else:
        pic(s, img, x + 0.78, 1.17, 1.32, 1.32)
    icon_circle(s, x + 0.16, 2.62, 0.42, col, ic)
    tb(s, x + 0.68, 2.62, cw - 0.75, 0.42, [[(head, {"bold": True})]], size=15, color=(VERM if i == 2 else INK),
       anchor="m")
    paras = []
    for sym, t in items:
        sc = {"+": "2E8B57", "−": "B03A2E"}[sym]
        paras.append({"runs": [(sym + "  ", {"bold": True, "color": sc}), (t, {})], "after": 4})
    tb(s, x + 0.2, 3.12, cw - 0.32, 0.76, paras, size=13, color=BODY, line=1.0)
    if ref:
        tb(s, x + 0.2, 3.86, cw - 0.32, 0.22, ref, size=10.5, color=MUTED, italic=True)
box(s, 0.5, 4.28, 9.0, 0.78, fill=PURPLE, radius=0.12)
tb(s, 0.75, 4.28, 8.5, 0.78, [[("RESEARCH QUESTION", {"bold": True, "color": PEACH, "size": 10, "spc": 120})],
                              [("Can whole-brain wiring be a useful prior for whole-body control?", {"bold": True})]],
   size=17, color="FFFFFF", anchor="m", line=1.05)
footer(s, "Images: FlyWire brain [2]; flybody fly [7]; FlyGM connectome layout [1].")
slide_number(s)
notes(s, """Target time: 60 seconds.

There are two common ways to connect a brain map to behavior. The first is to emulate the wiring: put a simple spiking neuron at every node and run it. This keeps the biology and needs no training, but the body is usually moved by hand-designed controllers. The second is to learn a generic controller with deep reinforcement learning. flybody does this very well, producing realistic walking and flight, but the network is a multilayer perceptron with no relation to the brain.

FlyGM sits in between. The real wiring is the network: every motor command flows through the 139-thousand-neuron connectome. Learning supplies what the connectome does not specify, and we train the whole system end to end on whole-body tasks. This gives a clean scientific question: can whole-brain wiring be a useful prior for whole-body control? And because the architecture is explicit, we can test it against controls of the same size with different wiring.""")
order.append(s)

# ---------------------------------------------------------------- 5. sensorimotor loop
s = new_slide()
kicker(s, "Method")
set_title(s, LAY, "FlyGM closes the sensorimotor loop")
pic(s, os.path.join(FIGS, "mainfig.png"), 0.5, 1.12, 9.0, 9.0 * 513 / 2004, name="FlyGM framework (paper Fig. 1)")
steps = [
    (AFF, "Observe", "Body senses → encoder → afferent neurons"),
    (INT_, "Propagate", "Signals flow along real synapses: M = W H"),
    (EFF, "Act", "Efferent neurons → decoder → 59 actuators"),
]
cw, gap = 2.86, 0.21
for i, (col, head, d) in enumerate(steps):
    x = 0.5 + i * (cw + gap)
    box(s, x, 3.6, cw, 1.42, fill=TINT, radius=0.07)
    num_circle(s, x + 0.18, 3.74, 0.44, col, i + 1, size=15)
    tb(s, x + 0.76, 3.74, cw - 0.85, 0.44, [[(head, {"bold": True})]], size=16, color=INK, anchor="m")
    tb(s, x + 0.2, 4.28, cw - 0.36, 0.66, d, size=13.5, color=BODY, line=1.02)
footer(s, "[1] Fig. 1, §2.1. Walking: 1,253-D input, 59 actions; flight: 104-D input, 12 actions.")
slide_number(s)
notes(s, """Target time: 60 seconds.

Here is the full loop. At every control step, the simulated fly produces an observation: proprioception, touch and, for walking, binocular vision; 1,253 numbers in total. A light encoder compresses this observation and injects it into the afferent, or input-facing, neurons of the connectome.

Information then propagates along the real synapses. Each neuron aggregates the states of its presynaptic partners through the connectome weight matrix, M equals W times H, and a shared update network, conditioned on a learned descriptor for that neuron, computes its new state. This is repeated for four message-passing layers.

Finally, the efferent, or output-facing, neurons are decoded into 59 actuator commands for the legs, adhesion pads, head and abdomen. The body moves in MuJoCo, and the next observation closes the loop. The connectome determines which neurons can communicate; the encoder, decoder and update rule are learned.""")
order.append(s)

# ---------------------------------------------------------------- 6. fixed vs learned
s = new_slide()
kicker(s, "Method")
set_title(s, LAY, "Wiring is fixed; computation is learned")
pic(s, A("network_panel_a.png"), 0.45, 1.12, h=3.42, name="Aggregated synapse graph (paper Fig. 7a)")
tb(s, 0.5, 4.62, 4.4, 0.3, "FlyWire neurons grouped by class [1, 2]", size=11, color=MUTED, italic=True)
box(s, 5.2, 1.1, 4.3, 1.62, fill=TINT, radius=0.06)
tb(s, 5.42, 1.2, 3.9, 0.32, [[("FIXED  ", {"bold": True, "color": PURPLE, "spc": 100, "size": 11}),
                            ("from the connectome", {"bold": True})]], size=15, color=INK)
tb(s, 5.42, 1.58, 3.95, 1.1, [
    {"runs": [("Who connects to whom", {})], "bullet": PURPLE, "after": 3},
    {"runs": [("Signed weights from synapses [13]", {})], "bullet": PURPLE, "after": 3},
    {"runs": [("Afferent → intrinsic → efferent", {})], "bullet": PURPLE},
], size=13.5, color=BODY, line=1.02)
box(s, 5.2, 2.86, 4.3, 1.42, fill=TINT_V, radius=0.06)
tb(s, 5.42, 2.96, 3.9, 0.32, [[("LEARNED  ", {"bold": True, "color": VERM, "spc": 100, "size": 11}),
                            ("by training", {"bold": True})]], size=15, color=INK)
tb(s, 5.42, 3.34, 3.95, 0.92, [
    {"runs": [("Encoder and decoder", {})], "bullet": VERM, "after": 3},
    {"runs": [("A descriptor per neuron, η", {}), ("v", {"baseline": -25000})], "bullet": VERM, "after": 3},
    {"runs": [("One shared update network, f", {}), ("ψ", {"baseline": -25000})], "bullet": VERM},
], size=13.5, color=BODY, line=1.02)
box(s, 5.2, 4.42, 4.3, 0.6, fill=None, radius=0.2, line=LINE, line_w=1)
tb(s, 5.25, 4.42, 4.2, 0.6, [[("M", {"italic": True}), ("t", {"baseline": -25000}), (" = W H", {"italic": True}),
                              ("t", {"baseline": -25000}), ("      H", {"italic": True}), ("t+1", {"baseline": -25000}),
                              (" [v] = f", {"italic": True}), ("ψ", {"baseline": -25000}),
                              ("( M", {"italic": True}), ("t", {"baseline": -25000}), (" [v] ‖ η", {"italic": True}),
                              ("v", {"baseline": -25000}), (" )", {"italic": True})]],
   size=16, color=INK, align="c", anchor="m", check=False)
footer(s, "[1] §2.1. States are learned vectors, not firing rates; signs follow [12, 13].")
slide_number(s)
notes(s, """Target time: 50 seconds.

It helps to separate what is fixed from what is learned. Fixed, from the connectome: who connects to whom; an approximate signed weight for every connection, computed as excitatory minus inhibitory synapse counts using neurotransmitter predictions from electron microscopy; and FlyWire's partition into afferent, intrinsic and efferent neurons. The figure on the left shows this structure aggregated by class.

Learned, by training: the observation encoder and action decoder, a descriptor for every neuron that can capture properties such as excitability or gain that the connectome does not record, and one shared update network.

Two caveats keep the claim precise. The latent states are learned vectors, not measured firing rates. And the signed weights are a modeling convention, not a receptor-level account of physiology. FlyGM is a connectome-structured policy; its biological fidelity is a separate question.""")
order.append(s)

# ---------------------------------------------------------------- 7. training
s = new_slide()
kicker(s, "Training")
set_title(s, LAY, "Imitation gives a head start, PPO refines")
flow_y, ih = 1.2, 1.6
nodes = [
    ("img", M("image5.png"), "Expert", "flybody MLP [7]"),
    ("step", 1, "Imitation", "Match the expert's actions"),
    ("img", A("connectome_layout_alpha_light.png"), "FlyGM", "Initialized policy"),
    ("step", 2, "PPO", "Maximize task reward [15]"),
    ("img", M("image8.jpeg"), "Tasks", "Walk, turn, fly"),
]
nw, ngap = 1.56, 0.30
for i, (kind, a, head, d) in enumerate(nodes):
    x = 0.5 + i * (nw + ngap)
    if kind == "img":
        if "image8" in a:
            pic(s, a, x + (nw - ih * 0.94) / 2, flow_y, ih * 0.94, ih)
        elif "connectome" in a:
            pic(s, a, x + (nw - ih) / 2, flow_y, ih, ih)
        else:
            pic(s, a, x - 0.05, flow_y + 0.24, nw + 0.1, (nw + 0.1) * 252 / 376)
        hc = INK
    else:
        col = PURPLE if a == 1 else VERM
        box(s, x, flow_y + 0.12, nw, ih - 0.24, fill=(TINT if a == 1 else TINT_V), radius=0.1)
        num_circle(s, x + nw / 2 - 0.3, flow_y + ih / 2 - 0.3, 0.6, col, a, size=22)
        hc = col
    tb(s, x - 0.12, flow_y + ih + 0.14, nw + 0.24, 0.32, [[(head, {"bold": True})]], size=16, color=hc, align="c")
    tb(s, x - 0.12, flow_y + ih + 0.48, nw + 0.24, 0.5, d, size=12.5, color=BODY, align="c", line=1.0)
    if i < len(nodes) - 1:
        arrow(s, x + nw + 0.06, flow_y + ih / 2, x + nw + ngap - 0.06, flow_y + ih / 2, color="B9B1C6", width=2)
cx = 0.5
for t, c in [("Same pipeline for every controller compared", PURPLE), ("Walking and flight: separate policies", MUTED)]:
    cx += chip(s, cx, 4.45, [(t, c, True)], size=13, fill="FFFFFF", line=c, h=0.44, pad=0.4) + 0.25
footer(s, "[1] §2.2, Appendix B. Expert policies and environments: flybody [7].")
slide_number(s)
notes(s, """Target time: 50 seconds.

Training has two stages. First, we roll out the existing flybody expert, an MLP policy, and keep successful episodes longer than one hundred steps. FlyGM is trained to match the expert's action distribution with a KL objective plus an annealed mean-squared-error term. This gives the connectome-structured policy stable walking and flight quickly.

Second, we fine-tune with proximal policy optimization on the task reward in MuJoCo, using an MLP value network and parallel rollouts, so the policy can improve beyond the demonstrations.

Two points matter for interpreting the results. Every controller we compare goes through exactly the same imitation-plus-PPO pipeline, so the comparisons isolate the architecture. And walking and flight are separately trained policies; we do not claim one policy switches between them.""")
order.append(s)

# ---------------------------------------------------------------- 8. behaviors (V1 slide 7, videos kept)
s = old[6]
for sh in list(s.shapes):
    if sh.shape_id not in {3, 4, 900, 901}:
        remove(sh)
kicker(s, "Behavior")
set_title(s, LAY, "Coordinated walking, turning and flight")
vw, vh = 4.38, 4.38 * 1080 / 1920
geo = {3: (0.5, 1.12), 900: (0.5, 1.12), 4: (5.12, 1.12), 901: (5.12, 1.12)}
for sh in s.shapes:
    if sh.shape_id in geo:
        x, y = geo[sh.shape_id]
        sh.left, sh.top, sh.width, sh.height = Inches(x), Inches(y), Inches(vw), Inches(vh)
lab_y = 1.12 + vh + 0.12
tb(s, 0.5, lab_y, 4.38, 0.32, [[("Walking policy", {"bold": True})]], size=16, color=INK)
tb(s, 5.12, lab_y, 4.38, 0.32, [[("Flight policy", {"bold": True}), ("  separately trained", {"color": MUTED, "size": 13})]],
   size=16, color=INK)
rows = [(0.5, [("Onset", "≈80 ms"), ("Tripod gait", "3 cm/s"), ("Turns", "10 rad/s")]),
        (5.12, [("Forward flight", "20 cm/s"), ("Wing-beat control", "")])]
size = 12.5
while max(sum(text_w(a + ("  " + b if b else ""), size, True) + 0.26 for a, b in items) + 0.1 * (len(items) - 1)
          for _, items in rows) > vw:
    size -= 0.25
for x0, items in rows:
    cx = x0
    for a, b in items:
        parts = [(a, INK, True)] + ([("  " + b, VERM, True)] if b else [])
        cx += chip(s, cx, lab_y + 0.42, parts, size=size, h=0.38, pad=0.26) + 0.1
tb(s, 0.5, lab_y + 0.92, 9.0, 0.26, "Click a video to play. Playback is slowed; read the simulation timestamps.",
   size=11, color=MUTED, italic=True)
footer(s, "[1] §3.2, Figs. 3–4. Left panels: connectome layout with afferent, intrinsic and efferent activity.")
slide_number(s)
notes(s, """Target time: 60 seconds.

Before the numbers, here is what the policies do. On the left is the walking policy. From rest, it initiates a gait within roughly the first 80 milliseconds, irregular steps quickly giving way to rhythmic stepping. During straight walking at 3 centimeters per second it produces the classic alternating tripod gait, and it turns at 10 radians per second by shortening the strides on the inner side, without any turn-specific rules.

On the right is a separately trained flight policy. It modulates the wing-beat pattern generator to keep stable forward flight at 20 centimeters per second while holding the body aligned with the target direction.

In both videos, the left panel shows the connectome in a force-directed layout, colored by afferent, intrinsic and efferent activity. The layout reflects connectivity, not the physical anatomy. The videos are slowed down, so the simulation timestamps are the right timescale. The next question is whether the specific wiring is what makes this work.""")
order.append(s)

# ---------------------------------------------------------------- 9. does the wiring matter? (merged V1 slides 8 + 9)
s = new_slide()
kicker(s, "Results · Does the wiring matter?")
set_title(s, LAY, "Real wiring learns faster, steers better")
ctrl = [
    ("FlyGM", VERM, "glyph_flygm.png", "Real wiring", TINT_V),
    ("DP-Rewired", GREEN, "glyph_rewired.png", "Same degrees, shuffled", TINT),
    ("ER-Random", GREY, "glyph_random.png", "Same size, random", TINT),
    ("MLP", BLUE, "glyph_mlp.png", "No graph, more params", TINT),
]
cw, gap = 2.13, 0.16
for i, (nm, col, g, d, fill) in enumerate(ctrl):
    x = 0.5 + i * (cw + gap)
    box(s, x, 1.1, cw, 0.74, fill=fill, radius=0.12)
    pic(s, A(g), x + 0.06, 1.15, 0.8, 0.64)
    tb(s, x + 0.9, 1.15, cw - 0.94, 0.3, [[(nm, {"bold": True})]], size=14, color=col)
    tb(s, x + 0.9, 1.46, cw - 0.94, 0.34, d, size=11, color=BODY, line=1.0)

# -- left: imitation convergence (Fig. 2a), log scale
curves = json.load(open(os.path.join(DATA, "fig2_curves.json")))["a"]
cd = XyChartData()
for nm in ["ER-Random", "Rewired", "MLP", "FlyGM"]:
    ser = cd.add_series(nm)
    xs, ys = curves[nm]["x"], curves[nm]["y"]
    for k in range(0, len(xs), 5):
        ser.add_data_point(round(xs[k], 3), round(ys[k], 5))
    ser.add_data_point(round(xs[-1], 3), round(ys[-1], 5))
reach_x, mlp_final = 8.91, curves["MLP"]["y"][-1]
ser = cd.add_series("MLP final level")
ser.add_data_point(reach_x, mlp_final)
ser.add_data_point(23.0, mlp_final)
ser = cd.add_series("FlyGM reaches it")
ser.add_data_point(reach_x, mlp_final)
tb(s, 0.5, 2.0, 4.5, 0.3, [[("Imitation: ", {"bold": True, "color": PURPLE}),
                           ("error to the expert (log)", {"bold": True})]], size=14, color=INK)
CX, CY, CW, CH = 0.4, 2.36, 4.6, 2.5
PL = (0.11, 0.03, 0.86, 0.80)
gf = s.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS, Inches(CX), Inches(CY), Inches(CW), Inches(CH), cd)
gf.name = "Chart: imitation convergence"
ch = gf.chart
chart_frame_clean(ch)
plot_layout(ch, *PL)
colors = {"ER-Random": GREY, "Rewired": GREEN, "MLP": BLUE, "FlyGM": VERM}
for ser in ch.plots[0].series:
    ser.smooth = False
    if ser.name in colors:
        ser.format.line.color.rgb = rgb(colors[ser.name])
        ser.format.line.width = Pt(2.75 if ser.name == "FlyGM" else 1.75)
        ser.marker.style = XL_MARKER_STYLE.NONE
    elif ser.name == "MLP final level":
        ser.format.line.color.rgb = rgb(BLUE)
        ser.format.line.width = Pt(1.25)
        ser.format.line.dash_style = MSO_LINE.DASH
        ser.marker.style = XL_MARKER_STYLE.NONE
    else:
        ser.format.line.fill.background()
        ser.marker.style = XL_MARKER_STYLE.CIRCLE
        ser.marker.size = 9
        ser.marker.format.fill.solid()
        ser.marker.format.fill.fore_color.rgb = rgb("FFFFFF")
        ser.marker.format.line.color.rgb = rgb(VERM)
        ser.marker.format.line.width = Pt(2)
xa, ya = ch.category_axis, ch.value_axis
xa.minimum_scale, xa.maximum_scale, xa.major_unit = 0, 23, 5
style_axis(xa, grid=False, fmt="0")
ya.minimum_scale, ya.maximum_scale = 0.04, 4
style_axis(ya, grid=True, line=False, fmt="0.0#")
axis_xml(ya, log=True, low_labels=True)
axis_xml(xa, crosses_at=0.04, low_labels=True)   # keep the x axis at the bottom of the log scale


def xy_to_slide(xv, yv, xmin=0.0, xmax=23.0, ymin=0.04, ymax=4.0):
    fx = (xv - xmin) / (xmax - xmin)
    fy = (math.log10(ymax) - math.log10(yv)) / (math.log10(ymax) - math.log10(ymin))
    return CX + CW * (PL[0] + PL[2] * fx), CY + CH * (PL[1] + PL[3] * fy)


plot_bottom = CY + CH * (PL[1] + PL[3])
tb(s, CX + CW * PL[0], plot_bottom + 0.26, CW * PL[2], 0.24, "Gradient updates (×10³)", size=11, color=MUTED, align="c")
ax_, ay_ = xy_to_slide(reach_x, mlp_final)
tb(s, ax_ - 0.3, ay_ - 1.13, 2.35, 0.62, [[("2.6× fewer updates", {"bold": True, "color": VERM, "size": 14})],
                                          [("to reach the MLP's final error", {})]], size=11.5, color=BODY, line=1.0)
arrow(s, ax_ + 0.02, ay_ - 0.5, ax_ + 0.0, ay_ - 0.1, color=VERM, width=1.5)

# -- right: angle error after IL + PPO (Table 1)
cats = ["2 cm/s\nstraight", "3 cm/s\nstraight", "3 cm/s\nturn 4", "3 cm/s\nturn 7"]
vals = {
    "FlyGM": [4.96, 5.57, 6.36, 8.29],
    "DP-Rewired": [7.84, 7.77, 9.72, 13.55],
    "ER-Random": [12.11, 11.33, 17.45, 125.36],
    "MLP": [6.76, 7.18, 8.85, 13.90],
}
cd = CategoryChartData()
cd.categories = cats
for k, v in vals.items():
    cd.add_series(k, v)
tb(s, 5.2, 2.0, 4.3, 0.3, [[("After IL + PPO: ", {"bold": True, "color": PURPLE}),
                           ("angle error (°)", {"bold": True})]], size=14, color=INK)
BX, BY, BW, BH = 5.1, 2.36, 4.45, 2.8
PB = (0.09, 0.1, 0.89, 0.7)
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(BX), Inches(BY), Inches(BW), Inches(BH), cd)
gf.name = "Chart: angle error"
ch = gf.chart
chart_frame_clean(ch)
plot_layout(ch, *PB)
plot = ch.plots[0]
plot.gap_width = 55
plot.overlap = 0
for ser, col in zip(plot.series, [VERM, GREEN, GREY, BLUE]):
    ser.format.fill.solid()
    ser.format.fill.fore_color.rgb = rgb(col)
    ser.format.line.fill.background()
    ser.invert_if_negative = False
for i, v in enumerate(vals["FlyGM"]):
    point_label(plot.series[0].points[i], f"{v:.1f}", VERM, size=10.5)
ca, va = ch.category_axis, ch.value_axis
style_axis(ca, grid=False, size=10.5)
va.minimum_scale, va.maximum_scale, va.major_unit = 0, 20, 5
style_axis(va, grid=True, line=False, fmt="0")
cat_w = BW * PB[2] / 4
bar_w = cat_w / (4 + 0.55)
er_x = BX + BW * PB[0] + 3 * cat_w + 0.55 * bar_w / 2 + 2.5 * bar_w
tb(s, er_x - 0.4, BY + BH * PB[1] - 0.26, 0.8, 0.24, [[("125 ↑", {"bold": True})]], size=10.5, color=GREY, align="c")
box(s, 5.6, 2.66, 2.3, 0.56, fill="FFFFFF", radius=0.1, line=LINE, line_w=0.75)
tb(s, 5.68, 2.66, 2.16, 0.56, [[("Lowest in all 4 tasks", {"bold": True, "color": VERM})],
                               [("22–39% below best baseline", {})]], size=11, color=BODY, anchor="m", line=1.0)
footer(s, "[1] Fig. 2a (one trace per model) and Table 1 (10 rollouts per condition; turn in rad/s).")
slide_number(s)
notes(s, """Target time: 90 seconds.

This is the central result. To test whether the specific wiring matters, we compare FlyGM with three controls, color-coded at the top. A degree-preserving rewiring keeps every neuron's number of inputs and outputs but shuffles its partners. An Erdos-Renyi random graph keeps only the number of neurons and edges. And an MLP has no graph at all, with more parameters than FlyGM. All four go through the identical imitation-plus-PPO pipeline.

On the left is imitation learning: the error between each model's actions and the expert's, on a log scale. FlyGM drops fastest and ends lowest. A simple way to read it: FlyGM reaches the MLP's final error after about nine thousand of the twenty-three thousand updates, roughly 2.6 times fewer. Both rewired and random graphs lag throughout.

On the right is control after the full pipeline: heading-angle error for straight walking at two speeds and for turning at 4 and 7 radians per second. FlyGM has the lowest error in every condition, 22 to 39 percent below the best baseline. The gap grows with turning rate: at 7 radians per second the rewired graph has 13.6 degrees and the random graph fails with 125 degrees, against 8.3 for FlyGM.

Two honest caveats: the training curves are one trace per model, and the error bars come from ten evaluation rollouts of one checkpoint, not from independently trained seeds.""")
order.append(s)

# ---------------------------------------------------------------- 10. what drives the gain
s = new_slide()
kicker(s, "Results · What drives the gain?")
set_title(s, LAY, "The gain survives ablations and GNNs")
gnn = [("GCN", 43.13), ("EdgeCNN", 90.48), ("GraphSAGE", 125.55), ("GAT", 132.65), ("PNA", 145.33), ("FlyGM", 334.34)]
tb(s, 0.5, 1.1, 4.4, 0.6, [[("Generic GNN layers fall short", {"bold": True})],
                           [("Same FlyWire neurons; reward after IL + PPO", {"color": MUTED, "size": 12, "bold": False})]],
   size=14.5, color=INK, line=1.05)
cd = CategoryChartData()
cd.categories = [g[0] for g in gnn]
cd.add_series("Reward", [g[1] for g in gnn])
gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(0.4), Inches(1.75), Inches(4.55), Inches(2.6), cd)
gf.name = "Chart: GNN baselines"
ch = gf.chart
chart_frame_clean(ch)
plot_layout(ch, 0.22, 0.02, 0.64, 0.86)
plot = ch.plots[0]
plot.gap_width = 40
ser = plot.series[0]
ser.format.line.fill.background()
for i, (nm, v) in enumerate(gnn):
    pt = ser.points[i]
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = rgb(VERM if nm == "FlyGM" else "B9B1C6")
    point_label(pt, f"{v:.0f}", VERM if nm == "FlyGM" else MUTED, size=11, bold=(nm == "FlyGM"))
ca, va = ch.category_axis, ch.value_axis
style_axis(ca, grid=False, size=11.5, color=INK)
va.minimum_scale, va.maximum_scale, va.major_unit = 0, 350, 100
style_axis(va, grid=True, line=False, fmt="0")
box(s, 0.5, 4.45, 4.4, 0.6, fill=TINT, radius=0.12)
tb(s, 0.68, 4.45, 4.1, 0.6, [[("2.3× ", {"bold": True, "color": VERM, "size": 16}),
                             ("the reward of the best GNN (PNA)", {})]], size=13, color=BODY, anchor="m")
abl = [
    ("MLP", 13.90, BLUE), ("DP-Rewired", 13.55, GREEN),
    ("Shared desc. + unweighted", 13.43, "F2B48A"), ("No desc. + unweighted", 13.23, "F2B48A"),
    ("Unweighted", 11.00, "F2B48A"),
    ("No desc.", 9.17, "E58A4E"), ("Shared desc.", 9.10, "E58A4E"),
    ("FlyGM (full)", 8.29, VERM),
]
tb(s, 5.2, 1.1, 4.3, 0.6, [[("Every connectome variant wins", {"bold": True})],
                           [("Angle error (°) at 3 cm/s, turn 7 rad/s", {"color": MUTED, "size": 12, "bold": False})]],
   size=14.5, color=INK, line=1.05)
cd = CategoryChartData()
cd.categories = [a[0] for a in abl]
cd.add_series("Angle error", [a[1] for a in abl])
gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(5.05), Inches(1.75), Inches(4.55), Inches(2.6), cd)
gf.name = "Chart: FlyGM ablations"
ch = gf.chart
chart_frame_clean(ch)
plot_layout(ch, 0.47, 0.01, 0.42, 0.88)
plot = ch.plots[0]
plot.gap_width = 35
ser = plot.series[0]
ser.format.line.fill.background()
for i, (nm, v, col) in enumerate(abl):
    pt = ser.points[i]
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = rgb(col)
    lc = {"MLP": BLUE, "DP-Rewired": GREEN}.get(nm, VERM)
    point_label(pt, f"{v:.1f}", lc, size=10.5, bold=(nm in ("FlyGM (full)", "MLP", "DP-Rewired")))
ca, va = ch.category_axis, ch.value_axis
style_axis(ca, grid=False, size=10.5, color=INK)
va.minimum_scale, va.maximum_scale, va.major_unit = 0, 16, 4
style_axis(va, grid=True, line=False, fmt="0")
box(s, 5.2, 4.45, 4.3, 0.6, fill=TINT_V, radius=0.12)
tb(s, 5.38, 4.45, 4.0, 0.6, [[("Signed weights help most: ", {"bold": True, "color": VERM}),
                             ("8.3° vs 11.0°", {})]], size=13, color=BODY, anchor="m")
footer(s, "[1] Tables 2–4. GNNs [16] sized to fit one A100 80 GB; desc. = per-neuron descriptor.")
slide_number(s)
notes(s, """Target time: 50 seconds. Optional slide: skip if short on time.

Two follow-up questions. First, could any graph neural network do this? On the left, we trained five standard GNN architectures, GCN, EdgeCNN, GraphSAGE, GAT and PNA, on the same FlyWire neurons and the same pipeline. None comes close: FlyGM earns 2.3 times the reward of the best one. A spiking network with the MLP layout never learns a stable gait. A caveat: the GNNs were sized to the largest configuration that fits one A100, so some are narrower than FlyGM.

Second, which part of FlyGM matters? On the right are ablations at the hardest condition, fast turning. Removing signed weights or per-neuron descriptors costs accuracy, and signed weights help most here: 8.3 versus 11.0 degrees. But every connectome variant, even unweighted and without descriptors, stays below the rewired graph and the MLP. That points to the topology itself as the main source of the prior.""")
order.append(s)

# ---------------------------------------------------------------- 11. rollout (V1 slide 10, polished; frame restyled only)
s = old[9]
for sh in list(s.shapes):
    if sh.shape_id in (5, 2, 4):
        remove(sh)
kicker(s, "Inside the model")
set_title(s, LAY, "One rollout links walking, pausing and turning")
vw = 9.0
for sh in s.shapes:
    if sh.shape_id in (3, 900):
        sh.left, sh.top, sh.width, sh.height = Inches(0.5), Inches(1.12), Inches(vw), Inches(vw * 850 / 1920)
footer(s, "Rollout 0–768 ms; click to play at 10× slow motion. Activity panel: [1] Fig. 5.")
slide_number(s)
order.append(s)

# ---------------------------------------------------------------- 12. latent responses (V1 slide 11, polished; frame restyled only)
s = old[10]
for sh in list(s.shapes):
    if sh.shape_id in (28, 2, 27):
        remove(sh)
for sh in s.shapes:
    sh.top = sh.top + Inches(0.15)
kicker(s, "Inside the model")
set_title(s, LAY, "Latent responses differ across neuron classes")
footer(s, "[1] Fig. 5, Appendix F; random-graph control: Fig. 8. Cf. [17, 18].")
slide_number(s)
order.append(s)

# ---------------------------------------------------------------- 13. outlook: male CNS
s = new_slide()
kicker(s, "Outlook")
set_title(s, LAY, "Next: from brain to brain + nerve cord")
iw = 5.3
ih = iw * 2265 / 3081
pic(s, A("malecns_dorsal.png"), 0.4, 1.08, iw, ih, name="Male CNS rendering (FlyEM v1.0)")
tb(s, 0.55, 1.08, 1.2, 0.26, [[("BRAIN", {"bold": True, "spc": 100})]], size=11, color=PURPLE)
tb(s, 1.8, 2.52, 1.0, 0.26, [[("NECK", {"bold": True, "spc": 100})]], size=11, color=PURPLE, align="c")
tb(s, 3.2, 1.62, 2.5, 0.26, [[("NERVE CORD", {"bold": True, "spc": 100})]], size=11, color=PURPLE, align="r")
lx = 0.5
for c, t in [("E06A1B", "Descending"), ("2B7BBF", "Ascending"), ("7B3FA0", "Motor neurons")]:
    box(s, lx, 5.0, 0.15, 0.15, fill=c, shape=MSO_SHAPE.OVAL)
    tb(s, lx + 0.22, 4.94, 1.4, 0.27, t, size=11.5, color=BODY, anchor="m")
    lx += 0.32 + text_w(t, 11.5) + 0.25
tb(s, 6.0, 1.08, 3.5, 0.62, [[("166,700", {"bold": True})]], size=34, color=PURPLE)
tb(s, 6.0, 1.7, 3.5, 0.3, "neurons: brain + nerve cord [8]", size=13.5, color=BODY)
tb(s, 6.0, 2.02, 3.5, 0.3, [[("1,314", {"bold": True, "color": "E06A1B"}), (" down  ·  ", {}),
                            ("1,846", {"bold": True, "color": "2B7BBF"}), (" up the neck", {})]], size=13, color=BODY)
tb(s, 6.0, 2.55, 3.5, 0.26, [[("WHAT IT ADDS TO FLYGM", {"bold": True, "spc": 100})]], size=10.5, color=MUTED)
tb(s, 6.0, 2.88, 3.5, 1.3, [
    {"runs": [("Real motor circuits below the brain", {})], "bullet": VERM, "after": 5},
    {"runs": [("Body feedback via real pathways", {})], "bullet": VERM, "after": 5},
    {"runs": [("Updated synapses and cell types [19, 20]", {})], "bullet": VERM},
], size=13.5, color=BODY, line=1.02)
box(s, 6.0, 4.4, 3.5, 0.62, fill=TINT, radius=0.14)
tb(s, 6.16, 4.4, 3.25, 0.62, [[("Next tests: ", {"bold": True, "color": PURPLE}),
                              ("longer tasks, recovery, navigation", {})]], size=12.5, color=BODY, anchor="m", line=1.0)
footer(s, "Outlook, not a result. Rendered from FlyEM male CNS v1.0: shells + ~280 sampled neurons [8].")
slide_number(s)
notes(s, """Target time: 50 seconds.

Where next? FlyGM currently uses the female brain connectome. The brain does not control the legs directly: descending neurons send commands through the neck to the ventral nerve cord, where premotor and motor circuits generate the actual leg and wing movements. In FlyGM, that last step is played by a learned decoder.

This image is rendered from the public male CNS release published this year: the brain on the left, the ventral nerve cord on the right, joined at the neck. In orange are a sample of the 1,314 descending neurons, in blue ascending neurons that carry feedback from the body back to the brain, and in purple motor neurons. The full dataset has 166,700 neurons.

For FlyGM, this would replace part of the learned decoder with measured motor wiring, and route body feedback through real ascending pathways. Together with updated synapse detections and sex-specific annotations, it would let us test connectome-based control on longer tasks, recovery from perturbations, and navigation. These are directions, not results.""")
order.append(s)

# ---------------------------------------------------------------- 14. takeaways
s = new_slide()
kicker(s, "Takeaways")
set_title(s, LAY, "A useful prior for whole-body control")
cards = [
    ("tree_w.png", "Wiring as architecture", "The fly connectome becomes a whole-body controller."),
    ("speed_w.png", "Real wiring matters", "Learns faster and steers better than rewired, random and MLP controls."),
    ("insights_w.png", "Structured dynamics", "Neuron classes show distinct activity patterns (qualitative)."),
]
cw, gap = 2.86, 0.21
for i, (ic, head, d) in enumerate(cards):
    x = 0.5 + i * (cw + gap)
    box(s, x, 1.12, cw, 2.42, fill=(TINT_V if i == 1 else TINT), radius=0.06)
    icon_circle(s, x + 0.22, 1.3, 0.62, (VERM if i == 1 else PURPLE), ic)
    tb(s, x + cw - 0.8, 1.26, 0.6, 0.6, [[(str(i + 1), {"bold": True})]], size=28, color="D9D2E3", align="r")
    tb(s, x + 0.22, 2.08, cw - 0.4, 0.34, [[(head, {"bold": True})]], size=16, color=INK)
    tb(s, x + 0.22, 2.48, cw - 0.4, 1.0, d, size=14, color=BODY, line=1.05)
box(s, 0.5, 3.74, 9.0, 0.7, fill=PURPLE, radius=0.12)
tb(s, 0.75, 3.74, 8.5, 0.7, [[("Not an upload: ", {"bold": True, "color": PEACH}),
                              ("a testable, connectome-structured controller.", {"bold": True})]],
   size=18, color="FFFFFF", anchor="m")
tb(s, 0.5, 4.6, 9.0, 0.4, [[("Open:  ", {"bold": True, "color": PURPLE}),
                            ("biological fidelity  ·  learning without demonstrations  ·  other bodies", {})]],
   size=13.5, color=BODY, anchor="m")
footer(s, "[1]. All comparisons use the shared IL + PPO pipeline; the representation analysis is qualitative.")
slide_number(s)
notes(s, """Target time: 45 seconds.

To summarize. First, a whole-brain connectome can serve as the architecture of a controller: FlyGM turns the 139-thousand-neuron FlyWire graph into a trainable policy that walks, turns and flies a detailed simulated body. Second, the real wiring matters: under the same training pipeline it learns faster and steers more accurately than rewired, random and MLP controls, and earns more than twice the reward of standard GNNs. Third, its internal dynamics are structured by neuron class, which a random graph does not show, though this analysis is qualitative.

To come back to the debate: this is not an upload of a fly. It is something more modest but testable, a connectome-structured controller whose assumptions are explicit. What remains open is the biological fidelity of the learned dynamics, learning without expert demonstrations, and transfer to other bodies and simulators.""")
order.append(s)

# ---------------------------------------------------------------- 15. thank you / demo page
s = new_slide()
set_title(s, LAY, "Thank you", size=46, color=PURPLE, x=0.6, y=0.75, w=4.2, h=0.85)
tb(s, 0.6, 1.62, 4.2, 0.36, "Questions are welcome.", size=18, color=BODY)
tb(s, 0.6, 2.3, 4.2, 0.26, [[("SEE THE DEMOS", {"bold": True, "spc": 150})]], size=11, color=VERM)
tb(s, 0.6, 2.6, 4.3, 0.4, [[("lnsgroup.cc/research/FlyGM", {"bold": True})]], size=20, color=INK)
pic(s, A("qr_flygm.png"), 0.6, 3.25, 1.5, 1.5, name="QR code: lnsgroup.cc/research/FlyGM", link=DEMO)
tb(s, 2.3, 3.4, 2.6, 1.2, [[("Zehao Jin, Yaoye Zhu,", {"bold": True})], [("Chen Zhang, Yanan Sui", {"bold": True})],
                           [("Tsinghua University · LNS Group", {"color": MUTED, "bold": False, "size": 12})]],
   size=13, color=INK, anchor="m", line=1.15)
FX, FY, FW, FH = 5.0, 0.75, 4.5, 4.3
shot = A("site_screenshot.png")
if os.path.exists(shot):
    from PIL import Image
    box(s, FX, FY, FW, FH, fill="FFFFFF", radius=0.03, line=LINE, line_w=1, name="Browser frame")
    box(s, FX, FY, FW, 0.32, fill="E9E5EF", radius=0.2, name="Browser bar")
    box(s, FX, FY + 0.16, FW, 0.16, fill="E9E5EF", radius=0.0, shape=MSO_SHAPE.RECTANGLE)
    for k, c in enumerate(["F26B5B", "F5BF4F", "5FC25D"]):
        box(s, FX + 0.14 + k * 0.17, FY + 0.105, 0.11, 0.11, fill=c, shape=MSO_SHAPE.OVAL)
    box(s, FX + 0.75, FY + 0.06, FW - 0.95, 0.2, fill="FFFFFF", radius=0.5)
    tb(s, FX + 0.85, FY + 0.06, FW - 1.1, 0.2, "lnsgroup.cc/research/FlyGM/", size=9, color=MUTED, anchor="m", check=False)
    iw_, ih_ = Image.open(shot).size
    vw_, vh_ = FW - 0.04, FH - 0.36
    p = pic(s, shot, FX + 0.02, FY + 0.34, vw_, vw_ * ih_ / iw_, name="Demo page screenshot", link=DEMO)
    if vw_ * ih_ / iw_ > vh_:
        p.crop_bottom = 1 - vh_ / (vw_ * ih_ / iw_)
        p.height = Inches(vh_)
else:
    # no screenshot of the page available at build time: stills from the demo videos
    box(s, FX, FY, FW, FH, fill=TINT, radius=0.04, name="Demo stills panel")
    tb(s, FX + 0.25, FY + 0.18, FW - 0.5, 0.26, [[("FROM THE DEMO VIDEOS", {"bold": True, "spc": 120})]],
       size=10.5, color=PURPLE)
    sw = (FW - 0.6) / 2
    for k, (img, cap) in enumerate([(M("image9.png"), "Walking"), (M("image10.png"), "Flight")]):
        x = FX + 0.25 + k * (sw + 0.1)
        pic(s, img, x, FY + 0.55, sw, sw * 1080 / 1920, link=DEMO)
        tb(s, x, FY + 0.6 + sw * 1080 / 1920, sw, 0.25, cap, size=11, color=BODY)
    bw_ = FW - 0.5
    pic(s, M("image12.png"), FX + 0.25, FY + 1.98, bw_, bw_ * 850 / 1920, link=DEMO)
    tb(s, FX + 0.25, FY + 2.03 + bw_ * 850 / 1920, bw_, 0.25, "Walk → pause → turn, with model activity", size=11,
       color=BODY)
notes(s, """Target time: 15 seconds.

Thank you. The demo page, at lnsgroup.cc slash research slash FlyGM, has the walking, turning and flight videos with live connectome activity; the QR code links there. I am happy to take questions.""")
order.append(s)

# ---------------------------------------------------------------- 16-17. references and resources
refs = [
    "Z. Jin, Y. Zhu, C. Zhang, Y. Sui. Whole-brain connectomic graph model enables whole-body locomotion control in fruit fly. AIAS+ 2026; arXiv:2602.17997.",
    "S. Dorkenwald et al. Neuronal wiring diagram of an adult brain. Nature 634, 124–138 (2024).",
    "P. Schlegel et al. Whole-brain annotation and multi-connectome cell typing of Drosophila. Nature 634, 139–152 (2024).",
    "P. K. Shiu et al. A Drosophila computational brain model reveals sensorimotor processing. Nature 634, 210–219 (2024).",
    "J. K. Lappalainen et al. Connectome-constrained networks predict neural activity across the fly visual system. Nature 634, 1132–1140 (2024).",
    "S. Wang-Chen et al. NeuroMechFly v2: simulating embodied sensorimotor control in adult Drosophila. Nat. Methods 21, 2353–2362 (2024).",
    "R. Vaxenburg et al. Whole-body physics simulation of fruit fly locomotion. Nature 643, 1312–1320 (2025).",
    "S. Berg et al. Sexual dimorphism in the complete connectome of the Drosophila male central nervous system. Cell 189, 5504–5526 (2026).",
    "Eon Systems. “We’ve uploaded a fruit fly”; “How the Eon team produced a virtual embodied fly”. eon.systems/updates (Mar 2026).",
    "A. Zeleznikow-Johnston. No, we haven’t uploaded a fly yet. LessWrong (Mar 2026).",
    "P. Mineault. Deconstructing viral fly sims. NeuroAI, neuroai.science (2026).",
    "D. A. Pospisil et al. The fly connectome reveals a path to the effectome. Nature 634, 201–209 (2024).",
    "N. Eckstein et al. Neurotransmitter classification from electron microscopy images at synaptic sites in Drosophila melanogaster. Cell 187, 2574–2594 (2024).",
    "E. Todorov, T. Erez, Y. Tassa. MuJoCo: a physics engine for model-based control. IROS, 5026–5033 (2012).",
    "J. Schulman et al. Proximal policy optimization algorithms. arXiv:1707.06347 (2017).",
    "GNN baselines: GCN, Kipf & Welling, ICLR (2017); EdgeConv, Wang et al., ACM TOG 38 (2019); GraphSAGE, Hamilton et al., NeurIPS (2017); GAT, Veličković et al., ICLR (2018); PNA, Corso et al., NeurIPS (2020).",
    "B. E. Brezovec et al. Mapping the neural dynamics of locomotion across the Drosophila brain. Curr. Biol. 34, 710–726 (2024).",
    "E. S. Schaffer et al. The spatial and temporal structure of neural activity across the fly brain. Nat. Commun. 14, 5572 (2023).",
    "S.-C. Yu et al. New synapse detection in the whole-brain connectome of Drosophila. bioRxiv (2025).",
    "D. Deutsch et al. Sexually-dimorphic neurons in the Drosophila whole-brain connectome. Research Square (2025).",
]
res = [
    ("Demo page", "lnsgroup.cc/research/FlyGM", DEMO),
    ("Paper", "arXiv:2602.17997", "https://arxiv.org/abs/2602.17997"),
    ("Connectome data", "codex.flywire.ai · male-cns.janelia.org", "https://codex.flywire.ai"),
    ("Body simulator", "github.com/TuragaLab/flybody", "https://github.com/TuragaLab/flybody"),
]
for part, (lo, hi) in enumerate([(0, 9), (9, 20)]):
    s = new_slide()
    kicker(s, "Backup")
    set_title(s, LAY, "References and resources" if part == 0 else "References (continued)")
    y = 1.12
    if part == 0:
        box(s, 0.5, 1.08, 9.0, 0.68, fill=TINT, radius=0.12)
        fs = 11
        widths = lambda f: [max(text_w(lab.upper(), 9, True) * 1.1, text_w(txt, f, True)) + 0.06 for lab, txt, _ in res]
        while sum(widths(fs)) > 8.6 - 0.15 * (len(res) - 1):
            fs -= 0.5
        ws = widths(fs)
        spare = (8.6 - sum(ws)) / (len(res) - 1)
        rx = 0.7
        for (lab, txt, url), w_ in zip(res, ws):
            tb(s, rx, 1.14, w_, 0.56, [[(lab.upper(), {"bold": True, "spc": 80, "size": 9, "color": PURPLE})],
                                      [(txt, {"bold": True, "link": url})]], size=fs, color=INK, line=1.1, check=False)
            rx += w_ + spare
        y = 1.95
    paras = [{"runs": [(f"[{k + 1}]  ", {"bold": True, "color": PURPLE}), (refs[k], {})], "after": 5}
             for k in range(lo, hi)]
    tb(s, 0.5, y, 9.0, 5.12 - y, paras, size=11, color=BODY, line=1.0)
    slide_number(s)
    notes(s, "Backup slide. Numbered references cited on the slides; [1] is our paper. "
             "The male CNS rendering on the outlook slide was produced from the public FlyEM v1.0 release "
             "(gs://flyem-male-cns).")
    order.append(s)

# ---------------------------------------------------------------- reorder and drop unused V1 slides
sldIdLst = prs.slides._sldIdLst
by_part = {}
for sld in list(sldIdLst):
    by_part[id(prs.part.related_part(sld.get(qn("r:id"))))] = sld
keep = [by_part[id(sl.part)] for sl in order]
for sld in list(sldIdLst):
    if sld not in keep:
        prs.part.drop_rel(sld.get(qn("r:id")))
    sldIdLst.remove(sld)
for sld in keep:
    sldIdLst.append(sld)

prs.core_properties.title = "FlyGM: Whole-Brain Connectomic Graph Model for Whole-Body Locomotion Control"
prs.core_properties.subject = "AIAS+ 2026 oral presentation"
prs.save(OUT)
for w in WARNINGS:
    print("WARNING", w)
print("saved", OUT, "slides:", len(prs.slides), "warnings:", len(WARNINGS))
