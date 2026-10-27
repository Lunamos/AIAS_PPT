"""Build the FlyGM AIAS+ 2026 oral deck (V2) on top of the V1 deck.

V1 is kept as the package base so that its videos (walking, flight, synced rollout)
and the polished rollout / latent-response slides survive unchanged; every other
slide is rebuilt from scratch.

Usage:
  python build_deck.py V1.pptx ASSET_DIR V1_MEDIA_DIR PAPER_FIG_DIR DATA_DIR OUT.pptx
"""
import json
import os
import sys

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_MARKER_STYLE, XL_TICK_MARK, XL_TICK_LABEL_POSITION
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

V1, ASSETS, MEDIA, FIGS, DATA, OUT = sys.argv[1:7]

# ---------------------------------------------------------------- palette
INK = "1F1A2E"        # titles, primary text
BODY = "3A3447"       # body text
MUTED = "7A7387"      # captions, footers
PURPLE = "660874"     # Tsinghua purple: kickers, structure
VERM = "D55E00"       # FlyGM (Okabe-Ito vermillion, as in the paper's Fig. 2)
GREEN = "009E73"      # degree-preserving rewiring
GREY = "666666"       # Erdos-Renyi random graph
BLUE = "0072B2"       # MLP
TINT = "F4F1F8"       # card fill
TINT_V = "FBEDE3"     # FlyGM card fill
LINE = "E2DCEA"       # hairlines, gridlines
DARK = "17111F"       # dark slides
DARK_CARD = "251C31"
AFF, INT_, EFF = "2E75B6", "4E9F5A", "E07B39"   # flow types in the paper's figures
FONT = "Arial"

W, H = 10.0, 5.625


def rgb(h):
    return RGBColor.from_string(h)


def A(name):
    return os.path.join(ASSETS, name)


def M(name):
    return os.path.join(MEDIA, name)


# ---------------------------------------------------------------- text helpers
def _apply_run(run, size=None, color=None, bold=None, italic=None, font=FONT, spc=None, baseline=None):
    f = run.font
    if size is not None:
        f.size = Pt(size)
    if color is not None:
        f.color.rgb = rgb(color)
    if bold is not None:
        f.bold = bold
    if italic is not None:
        f.italic = italic
    if font:
        f.name = font
    rPr = run._r.get_or_add_rPr()
    if spc is not None:
        rPr.set("spc", str(int(spc)))
    if baseline is not None:
        rPr.set("baseline", str(int(baseline)))


def _bullet(p, color, indent=0.16, char="•"):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(Inches(indent)))
    pPr.set("indent", str(-Inches(indent)))
    bc = etree.SubElement(pPr, qn("a:buClr"))
    c = etree.SubElement(bc, qn("a:srgbClr"))
    c.set("val", color)
    bf = etree.SubElement(pPr, qn("a:buFont"))
    bf.set("typeface", FONT)
    bch = etree.SubElement(pPr, qn("a:buChar"))
    bch.set("char", char)


def tb(slide, x, y, w, h, paras, size=14, color=BODY, bold=False, italic=False, align="l", anchor="t",
       line=1.08, after=0, name=None, wrap=True, bullet=None, margin=0.0):
    """Add a text box. `paras` is a str, or a list of paragraphs; each paragraph is a str or a
    list of runs (text, opts-dict). Paragraph-level overrides via a dict with key 'p'."""
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        shape.name = name
    tf = shape.text_frame
    tf.word_wrap = wrap
    tf.auto_size = MSO_AUTO_SIZE.NONE
    for side in ("left", "right", "top", "bottom"):
        setattr(tf, f"margin_{side}", Inches(margin))
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    if isinstance(paras, str):
        paras = [paras]
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        popts = {}
        if isinstance(para, dict):
            popts = para
            para = para["runs"]
        if isinstance(para, str):
            para = [(para, {})]
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[popts.get("align", align)]
        p.line_spacing = popts.get("line", line)
        sa = popts.get("after", after)
        if sa:
            p.space_after = Pt(sa)
        sb = popts.get("before", 0)
        if sb:
            p.space_before = Pt(sb)
        b = popts.get("bullet", bullet)
        if b:
            _bullet(p, b, popts.get("indent", 0.16))
        for text, o in para:
            r = p.add_run()
            r.text = text
            _apply_run(r, size=o.get("size", popts.get("size", size)), color=o.get("color", popts.get("color", color)),
                       bold=o.get("bold", popts.get("bold", bold)), italic=o.get("italic", italic),
                       spc=o.get("spc"), baseline=o.get("baseline"), font=o.get("font", FONT))
            if o.get("link"):
                r.hyperlink.address = o["link"]
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
    c = box(slide, x, y, d, d, fill=color, shape=MSO_SHAPE.OVAL)
    c.name = "Icon circle"
    p = slide.shapes.add_picture(A(os.path.join("icons", icon)), Inches(x + d * pad), Inches(y + d * pad),
                                 Inches(d * (1 - 2 * pad)), Inches(d * (1 - 2 * pad)))
    p.name = "Icon"
    return c


def num_circle(slide, x, y, d, color, n, size=13):
    box(slide, x, y, d, d, fill=color, shape=MSO_SHAPE.OVAL)
    tb(slide, x, y, d, d, [[(str(n), {"bold": True})]], size=size, color="FFFFFF", align="c", anchor="m")


def arrow(slide, x1, y1, x2, y2, color=MUTED, width=1.5, head=True, dash=False):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = rgb(color)
    c.line.width = Pt(width)
    _drop_style(c)
    if dash:
        c.line.dash_style = MSO_LINE.DASH
    if head:
        ln = c.line._get_or_add_ln()
        te = etree.SubElement(ln, qn("a:tailEnd"))
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
def kicker(slide, text, dark=False):
    box(slide, 0.5, 0.305, 0.085, 0.085, fill=VERM, shape=MSO_SHAPE.OVAL, name="Kicker dot")
    tb(slide, 0.65, 0.255, 6.5, 0.2, [[(text.upper(), {"spc": 120, "bold": True})]], size=9.5,
       color=("C9B8E8" if dark else PURPLE), anchor="m", name="Kicker")


def set_title(slide, layout, text, size=24, color=INK, x=0.5, y=0.47, w=9.0, h=0.5):
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
    return t


def footer(slide, text, dark=False):
    tb(slide, 0.5, 5.24, 8.3, 0.22, text, size=8.5, color=("9C93AC" if dark else MUTED), anchor="m", name="Source")


def slide_number(slide, dark=False):
    s = slide.shapes.add_textbox(Inches(9.0), Inches(5.24), Inches(0.5), Inches(0.22))
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
    rPr.set("sz", "900")
    sf = etree.SubElement(rPr, qn("a:solidFill"))
    c = etree.SubElement(sf, qn("a:srgbClr"))
    c.set("val", "9C93AC" if dark else MUTED)
    lat = etree.SubElement(rPr, qn("a:latin"))
    lat.set("typeface", FONT)
    t = etree.SubElement(fld, qn("a:t"))
    t.text = "#"


def strip_placeholders(slide, keep_title=True):
    for ph in list(slide.placeholders):
        if keep_title and ph.placeholder_format.type == PP_PLACEHOLDER.TITLE:
            continue
        remove(ph)


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def dark_bg(slide):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(DARK)


# ---------------------------------------------------------------- chart helpers
def chart_frame_clean(chart, size=9):
    cs = chart._chartSpace
    chart.font.name = FONT
    chart.font.size = Pt(size)
    chart.font.color.rgb = rgb(MUTED)
    chart.has_title = False
    chart.has_legend = False
    # no fill / border on chart space and plot area
    for parent in (cs, cs.find(".//" + qn("c:plotArea"))):
        old = parent.find(qn("c:spPr"))
        if old is not None:
            parent.remove(old)
        sp = etree.Element(qn("c:spPr"))
        etree.SubElement(sp, qn("a:noFill"))
        ln = etree.SubElement(sp, qn("a:ln"))
        etree.SubElement(ln, qn("a:noFill"))
        if parent is cs:
            cs.find(qn("c:chart")).addnext(sp)
        else:
            # plotArea: spPr goes after the axes, before extLst
            ext = parent.find(qn("c:extLst"))
            if ext is not None:
                ext.addprevious(sp)
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
        e = etree.SubElement(ml, qn("c:" + tag))
        e.set("val", str(val))


def style_axis(ax, grid=True, line=True, fmt=None, size=9):
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
    ax.tick_labels.font.color.rgb = rgb(MUTED)
    ax.tick_labels.font.name = FONT
    if fmt:
        ax.tick_labels.number_format = fmt
        ax.tick_labels.number_format_is_linked = False


def log_axis(ax):
    sc = ax._element.find(qn("c:scaling"))
    lb = etree.Element(qn("c:logBase"))
    lb.set("val", "10")
    sc.insert(0, lb)


def axis_title(ax, text, size=9):
    ax.has_title = True
    tf = ax.axis_title.text_frame
    tf.text = text
    r = tf.paragraphs[0].runs[0]
    _apply_run(r, size=size, color=MUTED, bold=False)


def point_label(pt, text, color, size=9, bold=True, pos=XL_LABEL_POSITION.OUTSIDE_END):
    dl = pt.data_label
    dl.has_text_frame = True
    dl.text_frame.text = text
    r = dl.text_frame.paragraphs[0].runs[0]
    _apply_run(r, size=size, color=color, bold=bold)
    dl.position = pos


# ================================================================ build
prs = Presentation(V1)
old = list(prs.slides)
LAY = prs.slide_layouts.get_by_name("1_Title and Content")


def new_slide():
    s = prs.slides.add_slide(LAY)
    strip_placeholders(s)
    return s


order = []

# ---------------------------------------------------------------- 1. title
s = new_slide()
dark_bg(s)
pic(s, A("connectome_layout_alpha.png"), 5.05, -0.35, 5.6, 5.6, name="Connectome graph layout")
tb(s, 0.6, 0.62, 5.0, 0.25, [[("AIAS+ 2026  ·  ORAL PRESENTATION", {"spc": 150, "bold": True})]], size=10,
   color="F0A46B")
set_title(s, LAY, "FlyGM", size=60, color="FFFFFF", x=0.6, y=0.98, w=5.0, h=1.0)
tb(s, 0.6, 2.02, 5.3, 0.75, "Whole-Brain Connectomic Graph Model for Whole-Body Locomotion Control",
   size=19, color="E6DEF3", line=1.05)
tb(s, 0.6, 3.12, 4.9, 0.3, "Zehao Jin, Yaoye Zhu, Chen Zhang, Yanan Sui", size=13, color="FFFFFF", bold=True)
tb(s, 0.6, 3.42, 4.9, 0.3, "Tsinghua University  ·  LNS Group", size=11, color="B9AECB")
# logo + QR card
box(s, 0.6, 4.18, 3.95, 0.92, fill="FFFFFF", radius=0.14, name="Logo card")
pic(s, M("image1.png"), 0.74, 4.27, 0.74, 0.74, name="Tsinghua University logo")
pic(s, M("image2.png"), 1.62, 4.46, 1.45, 0.36, name="LNS Group logo")
pic(s, A("qr_flygm.png"), 3.62, 4.25, 0.78, 0.78, name="QR code: lnsgroup.cc/research/FlyGM",
    link="https://lnsgroup.cc/research/FlyGM/")
tb(s, 3.03, 4.36, 0.55, 0.6, [[("Demo", {"bold": True})], "page"], size=8.5, color=BODY, align="r", anchor="m")
notes(s, """Target time: 30 seconds.

Good morning. Our work asks a simple question: can the wiring diagram of an animal's brain be used to build a controller for its body? We study the fruit fly, because it is the one animal for which we now have both a complete brain connectome and a detailed whole-body physics simulator. FlyGM turns that connectome into the architecture of a learned controller. I will explain why this question is timely, how FlyGM works, what the results say about the value of the real wiring, and where this could go next.""")
order.append(s)

# ---------------------------------------------------------------- 2. why now
s = new_slide()
kicker(s, "Why now")
set_title(s, LAY, "The fly brain and the fly body are now both digital")
pic(s, M("image4.png"), 0.45, 1.12, 4.1, 3.2, name="FlyWire whole-brain rendering")
stats = [("139,255", "neurons"), ("54.5 M", "synapses"), ("8,453", "cell types")]
for i, (v, l) in enumerate(stats):
    x = 0.5 + i * 1.38
    tb(s, x, 4.3, 1.35, 0.42, [[(v, {"bold": True})]], size=22, color=PURPLE)
    tb(s, x, 4.7, 1.35, 0.25, l, size=11, color=MUTED)
tl = [
    ("2024", PURPLE, "FlyWire: a complete adult brain",
     "Synapse-resolution wiring diagram of one female fly, proofread by a global community [2, 3]"),
    ("2024", PURPLE, "Connectome-constrained models",
     "Predict visual-system activity [5] and whole-brain sensorimotor responses [4]"),
    ("2025", VERM, "flybody: a whole-body fly in physics",
     "Anatomically detailed MuJoCo body; learned controllers walk and fly [7]; see also NeuroMechFly v2 [6]"),
    ("2026", VERM, "Male CNS: brain + nerve cord",
     "166,700 neurons, brain and ventral nerve cord joined at the neck [8]"),
]
x0, y0, dy = 5.05, 1.14, 0.92
arrow(s, x0 + 0.33, y0 + 0.2, x0 + 0.33, y0 + 3 * dy + 0.2, color=LINE, width=2, head=False)
for i, (yr, col, head, desc) in enumerate(tl):
    y = y0 + i * dy
    box(s, x0, y + 0.06, 0.66, 0.28, fill=col, radius=0.5)
    tb(s, x0, y + 0.06, 0.66, 0.28, [[(yr, {"bold": True})]], size=10.5, color="FFFFFF", align="c", anchor="m")
    tb(s, x0 + 0.85, y, 3.65, 0.3, [[(head, {"bold": True})]], size=13, color=INK)
    tb(s, x0 + 0.85, y + 0.3, 3.65, 0.55, desc, size=10.5, color=BODY, line=1.05)
footer(s, "Brain rendering: FlyWire [2]. Statistics: FlyWire v783 [2, 3]; male CNS v1.0 release [8].")
slide_number(s)
notes(s, """Target time: 60 seconds.

Why now? In 2024 the FlyWire consortium published the first complete wiring diagram of an adult brain: an adult female fruit fly, 139,255 neurons, about 54.5 million synapses, and more than eight thousand annotated cell types. Within months, connectome-constrained models showed that this wiring can predict neural activity in the visual system and sensorimotor responses across the brain.

At the same time, the body became digital too. flybody, published in Nature in 2025, provides an anatomically detailed fly in the MuJoCo physics engine, with learned controllers that walk and fly; NeuroMechFly v2 is a complementary platform. And this year the male central nervous system connectome added the ventral nerve cord, the fly's equivalent of the spinal cord.

So, for the first time, one animal has both a complete brain map and a whole-body simulator. The obvious next question is how to connect them.""")
order.append(s)

# ---------------------------------------------------------------- 3. the debate
s = new_slide()
kicker(s, "The 2026 debate")
set_title(s, LAY, "Can a wiring diagram, by itself, drive a body?")
tb(s, 0.5, 1.12, 4.4, 0.25, [[("WHAT WENT VIRAL", {"spc": 100, "bold": True})]], size=9, color=MUTED)
quotes = [
    (TINT_V, VERM, "“We’ve uploaded a fruit fly.”", "Eon Systems, March 2026 [9]",
     "A connectome-based LIF brain model [4] coupled to a NeuroMechFly body [6]: it walks, grooms and feeds."),
    (TINT, PURPLE, "“No, we haven’t uploaded a fly yet.”", "A. Zeleznikow-Johnston, LessWrong, March 2026 [10]",
     "Asks for evidence beyond a video: data, a write-up, known neural signatures."),
    (TINT, PURPLE, "“Deconstructing viral fly sims”", "P. Mineault, NeuroAI, 2026 [11]",
     "Most of the behavior comes from the body model's built-in controllers, not the connectome."),
]
for i, (fill, col, q, who, d) in enumerate(quotes):
    y = 1.42 + i * 1.2
    box(s, 0.5, y, 4.4, 1.08, fill=fill, radius=0.08)
    tb(s, 0.68, y + 0.1, 4.1, 0.3, [[(q, {"bold": True})]], size=14, color=INK)
    tb(s, 0.68, y + 0.4, 4.1, 0.22, who, size=9.5, color=col, bold=True)
    tb(s, 0.68, y + 0.63, 4.1, 0.42, d, size=10.5, color=BODY, line=1.02)
tb(s, 5.3, 1.12, 4.2, 0.25, [[("WHAT THE WIRING DIAGRAM LEAVES OUT", {"spc": 100, "bold": True})]], size=9, color=MUTED)
gaps = [
    ("tune_w.png", "Synaptic strength and sign in vivo",
     "The connectome gives possible paths, not how strongly neurons actually influence each other [12]."),
    ("chart_w.png", "Neuron dynamics",
     "Emulations assume simplified, largely uniform neuron models for very different cell types [4]."),
    ("cable_w.png", "The interface to the body",
     "In the demo, a few hand-picked neurons drive the body model's built-in locomotion and feeding controllers [9, 11]."),
]
for i, (ic, head, d) in enumerate(gaps):
    y = 1.45 + i * 1.08
    icon_circle(s, 5.3, y + 0.02, 0.5, PURPLE, ic)
    tb(s, 5.98, y, 3.55, 0.28, [[(head, {"bold": True})]], size=13, color=INK)
    tb(s, 5.98, y + 0.29, 3.55, 0.7, d, size=10.5, color=BODY, line=1.03)
box(s, 5.3, 4.56, 4.2, 0.58, fill=INK, radius=0.16)
tb(s, 5.46, 4.56, 3.95, 0.58, [[("OPEN QUESTION", {"bold": True, "color": "F0A46B", "size": 8.5, "spc": 100})],
                               [("How should a connectome control a whole body?", {"bold": True})]],
   size=12, color="FFFFFF", anchor="m", line=1.05)
footer(s, "Public discussion summarized from the sources cited; quotes are headlines of [9–11].")
slide_number(s)
notes(s, """Target time: 60 seconds.

This combination has already caught the public imagination. In March 2026, the startup Eon Systems announced "We've uploaded a fruit fly": a connectome-based leaky integrate-and-fire brain model, from Shiu and colleagues, coupled to the NeuroMechFly body. The video of a virtual fly walking, grooming and feeding went viral.

The response from neuroscientists was quick. Essays such as "No, we haven't uploaded a fly yet" and Patrick Mineault's "Deconstructing viral fly sims" pointed out what the wiring diagram leaves out. First, synapse counts tell us which paths exist, not how strongly neurons influence each other in the living animal; that is the effectome argument of Pospisil and colleagues. Second, the neuron models are simplified. Third, and most relevant for us, the interface to the body: according to Eon's own write-up, a few hand-picked descending neurons switch pre-built walking, grooming and feeding programs in the body model.

So the open question is not whether we have the wiring. It is how a connectome should be used to control a whole body.""")
order.append(s)

# ---------------------------------------------------------------- 4. our approach
s = new_slide()
kicker(s, "Our approach")
set_title(s, LAY, "FlyGM: use the wiring as the controller's architecture")
cols = [
    ("Emulate the wiring", "biotech_w.png", "8E86A0", TINT, M("image4.png"),
     [("+", "Biological connectome, no training"), ("−", "Generic neuron models"),
      ("−", "Body control handed to built-in controllers")], "e.g. [4, 9]"),
    ("Learn a generic network", "robot_w.png", "8E86A0", TINT, M("image5.png"),
     [("+", "Realistic whole-body walking and flight"), ("−", "Architecture unrelated to the brain"),
      ("−", "Hard to relate to biological circuits")], "e.g. flybody [7]"),
    ("FlyGM (ours)", "hub_w.png", VERM, TINT_V, A("connectome_layout_alpha_light.png"),
     [("+", "Motor commands pass through the 139k-neuron connectome"),
      ("+", "Trained end-to-end on whole-body tasks"),
      ("→", "Tests whether real wiring helps")], ""),
]
cw, gap = 2.86, 0.21
for i, (head, ic, col, fill, img, items, ref) in enumerate(cols):
    x = 0.5 + i * (cw + gap)
    box(s, x, 1.1, cw, 3.12, fill=fill, radius=0.06, line=(VERM if i == 2 else None), line_w=1.5)
    if i == 0:
        pic(s, img, x + 0.6, 1.2, 1.66, 1.3)
    elif i == 1:
        pic(s, img, x + 0.47, 1.2, 1.93, 1.3)
    else:
        pic(s, img, x + 0.78, 1.17, 1.32, 1.32)
    icon_circle(s, x + 0.16, 2.6, 0.36, col, ic)
    tb(s, x + 0.6, 2.6, cw - 0.7, 0.36, [[(head, {"bold": True})]], size=13.5,
       color=(VERM if i == 2 else INK), anchor="m")
    paras = []
    for sym, t in items:
        sc = {"+": "2E8B57", "−": "B03A2E", "→": VERM}[sym]
        paras.append({"runs": [(sym + "  ", {"bold": True, "color": sc}), (t, {})], "after": 3})
    tb(s, x + 0.18, 3.06, cw - 0.32, 0.9, paras, size=10.5, color=BODY, line=1.02)
    if ref:
        tb(s, x + 0.18, 3.96, cw - 0.32, 0.2, ref, size=9, color=MUTED, italic=True)
box(s, 0.5, 4.36, 9.0, 0.72, fill=INK, radius=0.12)
tb(s, 0.72, 4.36, 8.6, 0.72, [[("RESEARCH QUESTION", {"bold": True, "color": "F0A46B", "size": 8.5, "spc": 120})],
                              [("Can whole-brain wiring serve as a useful structural prior for whole-body control?",
                                {"bold": True})]], size=15, color="FFFFFF", anchor="m", line=1.05)
footer(s, "Images: FlyWire brain [2]; flybody fly model [7]; FlyGM force-directed connectome layout [1].")
slide_number(s)
notes(s, """Target time: 60 seconds.

There are two common ways to connect a brain map to behavior. The first is to emulate the wiring: put a simple spiking neuron at every node and run it. This keeps the biology and needs no training, but the neuron models are generic, and the body is usually driven by hand-designed controllers. The second is to learn a generic controller with deep reinforcement learning. flybody does this very well, producing realistic walking and flight, but the network is a multilayer perceptron with no relation to the brain.

FlyGM sits in between. We keep the wiring as the architecture: every motor command flows through the 139-thousand-neuron connectome. But we let learning supply what the connectome does not specify, and we train the whole system end to end on whole-body tasks. This framing gives us a clean scientific question: can whole-brain connectivity serve as a useful structural prior for whole-body control? And because the architecture is explicit, we can test it against controls with the same size but different wiring.""")
order.append(s)

# ---------------------------------------------------------------- 5. sensorimotor loop
s = new_slide()
kicker(s, "Method")
set_title(s, LAY, "FlyGM closes the sensorimotor loop")
pic(s, os.path.join(FIGS, "mainfig.png"), 0.5, 1.1, 9.0, 9.0 * 513 / 2004, name="FlyGM framework (paper Fig. 1)")
steps = [
    (AFF, "Observe", "Proprioception, touch and binocular vision (1,253-D) are encoded and injected into afferent neurons."),
    (INT_, "Propagate and update", "Messages flow along the synapses of ~139k neurons, M = WH; a shared MLP updates each neuron with its own descriptor."),
    (EFF, "Act", "Efferent states are decoded into 59 actuator commands for legs, adhesion, head and abdomen in MuJoCo."),
]
cw, gap = 2.86, 0.21
for i, (col, head, d) in enumerate(steps):
    x = 0.5 + i * (cw + gap)
    box(s, x, 3.6, cw, 1.42, fill=TINT, radius=0.07)
    num_circle(s, x + 0.16, 3.74, 0.38, col, i + 1)
    tb(s, x + 0.66, 3.74, cw - 0.75, 0.38, [[(head, {"bold": True})]], size=13.5, color=INK, anchor="m")
    tb(s, x + 0.18, 4.2, cw - 0.32, 0.8, d, size=10.5, color=BODY, line=1.04)
footer(s, "FlyGM [1]: Fig. 1 and §2.1. Walking configuration shown; flight uses a 104-D input and 12 actions. Simulator: flybody [7] in MuJoCo [14].")
slide_number(s)
notes(s, """Target time: 60 seconds.

Here is the full loop. At every control step, the simulated fly produces an observation: proprioception, touch and, for walking, binocular vision; 1,253 numbers in total. A light encoder compresses this observation and injects it into the afferent, or input-facing, neurons of the connectome.

Information then propagates along the real synapses. Each neuron aggregates the states of its presynaptic partners through the connectome weight matrix, M equals W times H, and a shared update network, conditioned on a learned descriptor for that neuron, computes its new state. This is repeated for four message-passing layers.

Finally, the efferent, or output-facing, neurons are decoded into 59 actuator commands for the legs, adhesion pads, head and abdomen. The body moves in MuJoCo, and the next observation closes the loop. The important design choice is that the connectome determines which neurons can communicate; the encoder, decoder and update rule are learned.""")
order.append(s)

# ---------------------------------------------------------------- 6. fixed vs learned
s = new_slide()
kicker(s, "Method")
set_title(s, LAY, "The wiring is fixed; the computation is learned")
pic(s, A("network_panel_a.png"), 0.5, 1.12, h=3.3, name="Aggregated synapse graph (paper Fig. 7a)")
tb(s, 0.5, 4.5, 4.4, 0.45, "FlyWire neurons grouped by superclass and flow type; edges show aggregated synapse counts.",
   size=9.5, color=MUTED, italic=True)
# fixed card
box(s, 5.25, 1.12, 4.25, 1.78, fill=TINT, radius=0.06)
tb(s, 5.45, 1.22, 3.9, 0.3, [[("FIXED  ", {"bold": True, "color": PURPLE, "spc": 100, "size": 10}),
                             ("from the connectome", {"bold": True})]], size=13, color=INK)
tb(s, 5.45, 1.57, 3.95, 1.3, [
    {"runs": [("Directed topology: ", {"bold": True}), ("who can talk to whom", {})], "bullet": PURPLE, "after": 3},
    {"runs": [("Signed weights: ", {"bold": True}), ("W", {}), ("vu", {"baseline": -25000}), (" = N", {}),
              ("exc", {"baseline": -25000}), (" − N", {}), ("inh", {"baseline": -25000}),
              ("  (ACh, DA +;  GABA, Glu, 5-HT, OA −) [12, 13]", {})], "bullet": PURPLE, "after": 3},
    {"runs": [("Flow types: ", {"bold": True}), ("afferent → intrinsic → efferent", {})], "bullet": PURPLE},
], size=10.5, color=BODY, line=1.04)
# learned card
box(s, 5.25, 3.02, 4.25, 1.42, fill=TINT_V, radius=0.06)
tb(s, 5.45, 3.12, 3.9, 0.3, [[("LEARNED  ", {"bold": True, "color": VERM, "spc": 100, "size": 10}),
                             ("by training", {"bold": True})]], size=13, color=INK)
tb(s, 5.45, 3.47, 3.95, 0.95, [
    {"runs": [("Observation encoder and action decoder", {})], "bullet": VERM, "after": 3},
    {"runs": [("Per-neuron descriptor η", {}), ("v", {"baseline": -25000}), (" (e.g. excitability, gain)", {})],
     "bullet": VERM, "after": 3},
    {"runs": [("Shared state-update network f", {}), ("ψ", {"baseline": -25000})], "bullet": VERM},
], size=10.5, color=BODY, line=1.04)
# equation
box(s, 5.25, 4.56, 4.25, 0.5, fill=None, radius=0.2, line=LINE, line_w=1)
tb(s, 5.3, 4.56, 4.15, 0.5, [[("M", {"italic": True}), ("t", {"baseline": -25000}), (" = W H", {"italic": True}),
                              ("t", {"baseline": -25000}), ("        H", {"italic": True}), ("t+1", {"baseline": -25000}),
                              (" [v] = f", {"italic": True}), ("ψ", {"baseline": -25000}),
                              ("( M", {"italic": True}), ("t", {"baseline": -25000}), (" [v] ‖ η", {"italic": True}),
                              ("v", {"baseline": -25000}), (" )", {"italic": True})]],
   size=14, color=INK, align="c", anchor="m")
footer(s, "FlyGM [1]: §2.1 and Appendix E. Latent states are learned vectors, not firing rates; signed weights are a modeling convention.")
slide_number(s)
notes(s, """Target time: 50 seconds.

It helps to separate what is fixed from what is learned. Fixed, from the connectome: the directed topology, meaning who can talk to whom; an approximate signed weight for every connection, computed as excitatory minus inhibitory synapse counts using neurotransmitter predictions from electron microscopy; and FlyWire's partition into afferent, intrinsic and efferent neurons. The figure on the left shows this structure aggregated by superclass.

Learned, by training: the observation encoder and action decoder, a descriptor for every neuron that can capture properties such as excitability or gain that the connectome does not record, and one shared update network.

Two caveats keep the claim precise. The latent states are learned vectors, not measured firing rates. And the signed weights are a modeling convention, not a receptor-level account of physiology. FlyGM is a connectome-structured policy; its biological fidelity is a separate question.""")
order.append(s)

# ---------------------------------------------------------------- 7. training
s = new_slide()
kicker(s, "Training")
set_title(s, LAY, "Imitation gives a head start; PPO refines it")
flow_y, ih = 1.18, 1.62
nodes = [
    ("img", M("image5.png"), "Expert behavior", "flybody MLP policy [7]; rollouts kept if > 100 steps"),
    ("step", 1, "Imitation learning", "Match the expert's action distribution: KL + annealed MSE"),
    ("img", A("connectome_layout_alpha_light.png"), "FlyGM policy", "Connectome-structured controller, initialized"),
    ("step", 2, "PPO refinement", "Optimize task reward in MuJoCo [15]; MLP value network, parallel rollouts"),
    ("img", M("image8.jpeg"), "Task interaction", "Walking, turning and flight tasks"),
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
            pic(s, a, x - 0.05, flow_y + 0.2, nw + 0.1, (nw + 0.1) * 252 / 376)
        hc = INK
    else:
        col = PURPLE if a == 1 else VERM
        box(s, x, flow_y + 0.12, nw, ih - 0.24, fill=(TINT if a == 1 else TINT_V), radius=0.1)
        num_circle(s, x + nw / 2 - 0.28, flow_y + ih / 2 - 0.28, 0.56, col, a, size=20)
        hc = col
    tb(s, x - 0.08, flow_y + ih + 0.12, nw + 0.16, 0.3, [[(head, {"bold": True})]], size=12.5, color=hc, align="c")
    tb(s, x - 0.08, flow_y + ih + 0.42, nw + 0.16, 0.75, d, size=10, color=BODY, align="c", line=1.03)
    if i < len(nodes) - 1:
        arrow(s, x + nw + 0.06, flow_y + ih / 2, x + nw + ngap - 0.06, flow_y + ih / 2, color="B9B1C6", width=2)
chips = [("Same IL + PPO pipeline for every controller we compare", PURPLE),
         ("Walking and flight use separately trained policies", MUTED)]
cx = 0.5
for t, c in chips:
    wch = 0.35 + 0.075 * len(t)
    box(s, cx, 4.42, wch, 0.42, fill=None, radius=0.5, line=c, line_w=1.25)
    tb(s, cx, 4.42, wch, 0.42, t, size=11, color=c, bold=True, align="c", anchor="m")
    cx += wch + 0.25
footer(s, "FlyGM [1]: §2.2 and Appendix B. Expert policies and environments from flybody [7].")
slide_number(s)
notes(s, """Target time: 50 seconds.

Training has two stages. First, we roll out the existing flybody expert, an MLP policy, and keep successful episodes longer than one hundred steps. FlyGM is trained to match the expert's action distribution with a KL objective plus an annealed mean-squared-error term. This gives the connectome-structured policy stable walking and flight quickly.

Second, we fine-tune with proximal policy optimization on the task reward in MuJoCo, using an MLP value network and parallel rollouts, so the policy can improve beyond the demonstrations.

Two points matter for interpreting the results. Every controller we compare goes through exactly the same imitation-plus-PPO pipeline, so the comparisons isolate the architecture. And walking and flight are separately trained policies; we do not claim one policy switches between them.""")
order.append(s)

# ---------------------------------------------------------------- 8. behaviors (V1 slide 7, videos kept)
s = old[6]
keep_ids = {3, 4, 900, 901}
for sh in list(s.shapes):
    if sh.shape_id not in keep_ids:
        remove(sh)
kicker(s, "Behavior")
set_title(s, LAY, "FlyGM drives coordinated walking, turning and flight")
vw, vh = 4.38, 4.38 * 1080 / 1920
geo = {3: (0.5, 1.12), 900: (0.5, 1.12), 4: (5.12, 1.12), 901: (5.12, 1.12)}
for sh in s.shapes:
    if sh.shape_id in geo:
        x, y = geo[sh.shape_id]
        sh.left, sh.top, sh.width, sh.height = Inches(x), Inches(y), Inches(vw), Inches(vh)
lab_y = 1.12 + vh + 0.14
tb(s, 0.5, lab_y, 4.38, 0.28, [[("Walking policy", {"bold": True})]], size=13, color=INK)
tb(s, 5.12, lab_y, 4.38, 0.28, [[("Flight policy", {"bold": True}), ("  (trained separately)", {"color": MUTED})]],
   size=13, color=INK)
chip_rows = [
    (0.5, [("Gait onset", "≈ 80 ms"), ("Tripod gait", "3 cm/s"), ("Turns", "10 rad/s")]),
    (5.12, [("Forward flight", "20 cm/s"), ("Wing-beat generator modulation", "")]),
]
for x0, chips in chip_rows:
    cx = x0
    for a, b in chips:
        txt = a + ("  " + b if b else "")
        wch = 0.26 + 0.07 * len(txt)
        box(s, cx, lab_y + 0.36, wch, 0.36, fill=TINT, radius=0.5)
        tb(s, cx, lab_y + 0.36, wch, 0.36, [[(a, {"bold": True}), (("  " + b) if b else "", {"color": VERM, "bold": True})]],
           size=10, color=INK, align="c", anchor="m")
        cx += wch + 0.1
tb(s, 0.5, lab_y + 0.84, 9.0, 0.24, "Click a video to play. Playback is slowed for viewing; read the simulation timestamps.", size=9.5,
   color=MUTED, italic=True)
footer(s, "FlyGM [1]: §3.2, Figs. 3–4. Left panels: force-directed connectome layout with afferent, intrinsic and efferent activity.")
slide_number(s)
notes(s, """Target time: 60 seconds.

Before the numbers, here is what the policies do. On the left is the walking policy. From rest, it initiates a gait within roughly the first 80 milliseconds, irregular steps quickly giving way to rhythmic stepping. During straight walking at 3 centimeters per second it produces the classic alternating tripod gait, and it turns at 10 radians per second by shortening the strides on the inner side, without any turn-specific rules.

On the right is a separately trained flight policy. It modulates the wing-beat pattern generator to keep stable forward flight at 20 centimeters per second while holding the body aligned with the target direction.

In both videos, the left panel shows the connectome in a force-directed layout, colored by afferent, intrinsic and efferent activity. The layout reflects connectivity, not the physical anatomy. The videos are slowed down, so the simulation timestamps are the right timescale. The next question is whether the specific wiring is what makes this work.""")
order.append(s)

# ---------------------------------------------------------------- 9. topology matters (merged V1 slides 8 + 9)
s = new_slide()
kicker(s, "Results · Does the wiring matter?")
set_title(s, LAY, "The real wiring learns faster and steers more accurately")
ctrl = [
    ("FlyGM", VERM, "glyph_flygm.png", "Real FlyWire wiring with signed weights", TINT_V),
    ("DP-Rewired", GREEN, "glyph_rewired.png", "Same in/out-degree per neuron, shuffled partners", TINT),
    ("ER-Random", GREY, "glyph_random.png", "Same numbers of neurons and edges", TINT),
    ("MLP", BLUE, "glyph_mlp.png", "4 × 512 layers; more parameters than FlyGM", TINT),
]
cw, gap = 2.13, 0.16
for i, (nm, col, g, d, fill) in enumerate(ctrl):
    x = 0.5 + i * (cw + gap)
    box(s, x, 1.1, cw, 0.84, fill=fill, radius=0.1)
    pic(s, A(g), x + 0.06, 1.17, 0.88, 0.7)
    tb(s, x + 0.98, 1.16, cw - 1.04, 0.26, [[(nm, {"bold": True})]], size=12.5, color=col)
    tb(s, x + 0.98, 1.43, cw - 1.06, 0.48, d, size=9, color=BODY, line=1.0)

# -- left chart: imitation convergence (Fig. 2a), log scale
curves = json.load(open(os.path.join(DATA, "fig2_curves.json")))["a"]
cd = XyChartData()
for nm in ["ER-Random", "Rewired", "MLP", "FlyGM"]:
    ser = cd.add_series(nm)
    xs, ys = curves[nm]["x"], curves[nm]["y"]
    step = 5
    for k in range(0, len(xs), step):
        ser.add_data_point(round(xs[k], 3), round(ys[k], 5))
    ser.add_data_point(round(xs[-1], 3), round(ys[-1], 5))
reach_x, mlp_final = 8.91, curves["MLP"]["y"][-1]
ser = cd.add_series("MLP final level")
ser.add_data_point(reach_x, mlp_final)
ser.add_data_point(23.0, mlp_final)
ser = cd.add_series("FlyGM reaches it")
ser.add_data_point(reach_x, mlp_final)
CX, CY, CW, CH = 0.42, 2.38, 4.55, 2.68
tb(s, 0.5, 2.06, 4.4, 0.3, [[("Imitation: ", {"bold": True, "color": PURPLE}),
                             ("error to the expert's actions", {"bold": True})]], size=12, color=INK)
gf = s.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS, Inches(CX), Inches(CY), Inches(CW), Inches(CH), cd)
gf.name = "Chart: imitation convergence"
ch = gf.chart
chart_frame_clean(ch)
PL = (0.13, 0.04, 0.83, 0.76)
plot_layout(ch, *PL)
colors = {"ER-Random": GREY, "Rewired": GREEN, "MLP": BLUE, "FlyGM": VERM}
for ser in ch.plots[0].series:
    ser.smooth = False
    if ser.name in colors:
        ser.format.line.color.rgb = rgb(colors[ser.name])
        ser.format.line.width = Pt(2.5 if ser.name == "FlyGM" else 1.6)
        ser.marker.style = XL_MARKER_STYLE.NONE
    elif ser.name == "MLP final level":
        ser.format.line.color.rgb = rgb(BLUE)
        ser.format.line.width = Pt(1.25)
        ser.format.line.dash_style = MSO_LINE.DASH
        ser.marker.style = XL_MARKER_STYLE.NONE
    else:
        ser.format.line.fill.background()
        ser.marker.style = XL_MARKER_STYLE.CIRCLE
        ser.marker.size = 8
        ser.marker.format.fill.solid()
        ser.marker.format.fill.fore_color.rgb = rgb("FFFFFF")
        ser.marker.format.line.color.rgb = rgb(VERM)
        ser.marker.format.line.width = Pt(2)
xa, ya = ch.category_axis, ch.value_axis
xa.minimum_scale, xa.maximum_scale, xa.major_unit = 0, 23, 5
style_axis(xa, grid=False, fmt="0")
axis_title(xa, "Gradient updates (×10³)")
ya.minimum_scale, ya.maximum_scale = 0.04, 4
log_axis(ya)
style_axis(ya, grid=True, line=False, fmt="0.0#")
axis_title(ya, "Action-mean MSE (log)")


def xy_to_slide(xv, yv, xmin=0, xmax=23, ymin=0.04, ymax=4):
    import math
    fx = (xv - xmin) / (xmax - xmin)
    fy = (math.log10(ymax) - math.log10(yv)) / (math.log10(ymax) - math.log10(ymin))
    return CX + CW * (PL[0] + PL[2] * fx), CY + CH * (PL[1] + PL[3] * fy)


ax_, ay_ = xy_to_slide(reach_x, mlp_final)
tb(s, ax_ - 0.2, ay_ - 1.12, 2.45, 0.62, [[("2.6× fewer updates", {"bold": True, "color": VERM, "size": 12})],
                                          [("FlyGM reaches the MLP's final error after ~9k of 23k steps", {})]],
   size=9.5, color=BODY, line=1.0)
arrow(s, ax_ + 0.05, ay_ - 0.48, ax_ + 0.01, ay_ - 0.1, color=VERM, width=1.25)

# -- right chart: angle error after IL + PPO (Table 1)
cats = ["v=2, ψ=0", "v=3, ψ=0", "v=3, ψ=4", "v=3, ψ=7"]
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
BX, BY, BW, BH = 5.08, 2.38, 4.5, 2.68
tb(s, 5.2, 2.06, 4.3, 0.3, [[("After IL + PPO: ", {"bold": True, "color": PURPLE}),
                             ("heading-angle error (°)", {"bold": True})]], size=12, color=INK)
gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(BX), Inches(BY), Inches(BW), Inches(BH), cd)
gf.name = "Chart: angle error"
ch = gf.chart
chart_frame_clean(ch)
PB = (0.09, 0.1, 0.89, 0.74)
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
    point_label(plot.series[0].points[i], f"{v:.1f}", VERM, size=8.5)
ca, va = ch.category_axis, ch.value_axis
style_axis(ca, grid=False, size=9)
va.minimum_scale, va.maximum_scale, va.major_unit = 0, 20, 5
style_axis(va, grid=True, line=False, fmt="0")
# ER-Random high-yaw bar is off-scale (125.4 deg)
cat_w = BW * PB[2] / 4
bar_w = cat_w / (4 + 0.55)
er_x = BX + BW * PB[0] + 3 * cat_w + 0.55 * bar_w / 2 + 2.5 * bar_w
tb(s, er_x - 0.4, BY + BH * PB[1] - 0.24, 0.8, 0.22, [[("125 ↑", {"bold": True})]], size=8.5, color=GREY, align="c")
box(s, 5.62, 2.62, 2.25, 0.5, fill="FFFFFF", radius=0.1, line=LINE, line_w=0.75)
tb(s, 5.7, 2.62, 2.12, 0.5, [[("Lowest error in all 4 conditions", {"bold": True, "color": VERM})],
                             [("22–39% below the best baseline", {})]], size=9, color=BODY, anchor="m", line=1.0)
footer(s, "FlyGM [1]: Fig. 2a (one training trace per model, EMA) and Table 1 (10 evaluation rollouts of a fixed checkpoint; v in cm/s, ψ in rad/s).")
slide_number(s)
notes(s, """Target time: 90 seconds.

This is the central result. To test whether the specific wiring matters, we compare FlyGM with three controls, color-coded at the top. A degree-preserving rewiring keeps every neuron's number of inputs and outputs but shuffles its partners. An Erdos-Renyi random graph keeps only the number of neurons and edges. And an MLP has no graph at all, with more parameters than FlyGM. All four go through the identical imitation-plus-PPO pipeline.

On the left is imitation learning: the error between each model's actions and the expert's, on a log scale. FlyGM drops fastest and ends lowest. A simple way to read it: FlyGM reaches the MLP's final error after about nine thousand of the twenty-three thousand updates, roughly 2.6 times fewer. Both rewired and random graphs lag throughout.

On the right is control after the full pipeline: heading-angle error across speeds and turning rates. FlyGM has the lowest error in every condition, 22 to 39 percent below the best baseline. The gap grows with turning rate: at 7 radians per second the rewired graph has 13.6 degrees and the random graph fails with 125 degrees, against 8.3 for FlyGM.

Two honest caveats: the training curves are one trace per model, and the error bars come from ten evaluation rollouts of one checkpoint, not from independently trained seeds.""")
order.append(s)

# ---------------------------------------------------------------- 10. what drives the gain
s = new_slide()
kicker(s, "Results · What drives the gain?")
set_title(s, LAY, "The gain holds up under ablations and GNN baselines")
# left: standard GNNs on the same neurons
gnn = [("GCN", 43.13), ("EdgeCNN", 90.48), ("GraphSAGE", 125.55), ("GAT", 132.65), ("PNA", 145.33), ("FlyGM", 334.34)]
tb(s, 0.5, 1.1, 4.4, 0.5, [[("Generic GNN layers fall short", {"bold": True})],
                           [("Five standard GNNs on the same FlyWire neurons: average reward after IL + PPO", {"color": MUTED, "size": 10})]],
   size=12.5, color=INK, line=1.02)
cd = CategoryChartData()
cd.categories = [g[0] for g in gnn]
cd.add_series("Reward", [g[1] for g in gnn])
GX, GY, GW, GH = 0.4, 1.68, 4.55, 2.62
gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(GX), Inches(GY), Inches(GW), Inches(GH), cd)
gf.name = "Chart: GNN baselines"
ch = gf.chart
chart_frame_clean(ch)
plot_layout(ch, 0.2, 0.03, 0.66, 0.84)
plot = ch.plots[0]
plot.gap_width = 45
ser = plot.series[0]
ser.format.line.fill.background()
for i, (nm, v) in enumerate(gnn):
    pt = ser.points[i]
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = rgb(VERM if nm == "FlyGM" else "B9B1C6")
    point_label(pt, f"{v:.0f}", VERM if nm == "FlyGM" else MUTED, size=9, bold=(nm == "FlyGM"))
ca, va = ch.category_axis, ch.value_axis
style_axis(ca, grid=False, size=10)
ca.tick_labels.font.color.rgb = rgb(INK)
va.minimum_scale, va.maximum_scale, va.major_unit = 0, 350, 100
style_axis(va, grid=True, line=False, fmt="0")
box(s, 0.5, 4.38, 4.4, 0.66, fill=TINT, radius=0.1)
tb(s, 0.65, 4.38, 4.15, 0.66, [[("2.3× ", {"bold": True, "color": VERM, "size": 13}),
                                ("the reward of the best standard GNN (PNA). A spiking network with the MLP layout never learns a stable gait.", {})]],
   size=10, color=BODY, anchor="m", line=1.02)
# right: ablations at the hardest condition
abl = [
    ("MLP", 13.90, BLUE), ("DP-Rewired", 13.55, GREEN),
    ("Shared desc., unweighted", 13.43, "F2B48A"), ("No desc., unweighted", 13.23, "F2B48A"),
    ("Unweighted", 11.00, "F2B48A"),
    ("No desc., weighted", 9.17, "E58A4E"), ("Shared desc., weighted", 9.10, "E58A4E"),
    ("FlyGM (full)", 8.29, VERM),
]
tb(s, 5.2, 1.1, 4.3, 0.5, [[("Every connectome variant beats the controls", {"bold": True})],
                           [("Ablations at the hardest condition (v = 3 cm/s, ψ = 7 rad/s): angle error (°)", {"color": MUTED, "size": 10})]],
   size=12.5, color=INK, line=1.02)
cd = CategoryChartData()
cd.categories = [a[0] for a in abl]
cd.add_series("Angle error", [a[1] for a in abl])
gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(5.1), Inches(1.68), Inches(4.5), Inches(2.62), cd)
gf.name = "Chart: FlyGM ablations"
ch = gf.chart
chart_frame_clean(ch)
plot_layout(ch, 0.42, 0.02, 0.46, 0.86)
plot = ch.plots[0]
plot.gap_width = 40
ser = plot.series[0]
ser.format.line.fill.background()
for i, (nm, v, col) in enumerate(abl):
    pt = ser.points[i]
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = rgb(col)
    lc = {"MLP": BLUE, "DP-Rewired": GREEN}.get(nm, VERM)
    point_label(pt, f"{v:.2f}", lc, size=8.5, bold=(nm in ("FlyGM (full)", "MLP", "DP-Rewired")))
ca, va = ch.category_axis, ch.value_axis
style_axis(ca, grid=False, size=9)
ca.tick_labels.font.color.rgb = rgb(INK)
va.minimum_scale, va.maximum_scale, va.major_unit = 0, 16, 4
style_axis(va, grid=True, line=False, fmt="0")
box(s, 5.2, 4.38, 4.3, 0.66, fill=TINT_V, radius=0.1)
tb(s, 5.35, 4.38, 4.05, 0.66, [[("Topology carries most of the prior; ", {"bold": True, "color": VERM}),
                                ("signed weights help most where control is hardest (8.3° vs 11.0° unweighted).", {})]],
   size=10, color=BODY, anchor="m", line=1.02)
footer(s, "FlyGM [1]: Tables 2–4 (Appendix). GNNs [16–20] use the largest width and depth that fit one A100 80 GB; reward table repetitions not verified.")
slide_number(s)
notes(s, """Target time: 50 seconds. Optional slide: skip if short on time.

Two follow-up questions. First, is the gain simply from using a graph neural network? On the left, we put five standard GNN layers, GCN, EdgeCNN, GraphSAGE, GAT and PNA, on the same FlyWire neurons and the same pipeline. None comes close: FlyGM earns 2.3 times the reward of the best one. A spiking network with the MLP layout never learns a stable gait. A caveat: the GNNs were sized to the largest configuration that fits one A100, so some are narrower than FlyGM.

Second, which part of FlyGM matters? On the right are ablations at the hardest condition, fast turning. Removing signed weights or per-neuron descriptors costs accuracy, and signed weights help most here: 8.3 versus 11.0 degrees. But every connectome variant, even unweighted and without descriptors, stays below the rewired graph and the MLP. That points to the topology itself as the main source of the prior.""")
order.append(s)

# ---------------------------------------------------------------- 11. rollout (V1 slide 10, polished; restyled frame only)
s = old[9]
for sh in list(s.shapes):
    if sh.shape_id in (5, 2, 4):
        remove(sh)
kicker(s, "Inside the model")
set_title(s, LAY, "One rollout links walking, pausing and turning")
vw = 9.0
vh = vw * 850 / 1920
for sh in s.shapes:
    if sh.shape_id in (3, 900):
        sh.left, sh.top, sh.width, sh.height = Inches(0.5), Inches(1.08), Inches(vw), Inches(vh)
footer(s, "Original rollout: 0–768 ms · click to play at 10× slow motion · illustrative archived recording; activity panel from FlyGM [1] Fig. 5.")
slide_number(s)
order.append(s)

# ---------------------------------------------------------------- 12. latent responses (V1 slide 11, polished; restyled frame only)
s = old[10]
for sh in list(s.shapes):
    if sh.shape_id in (28, 2, 27):
        remove(sh)
dy = Inches(0.13)
for sh in s.shapes:
    sh.top = sh.top + dy
kicker(s, "Inside the model")
set_title(s, LAY, "Latent responses differ across neuron classes")
footer(s, "FlyGM [1]: Fig. 5 and Appendix F (PCA, class grouping, within-class ordering); random-graph control in Fig. 8. Cf. [21, 22].")
slide_number(s)
order.append(s)

# ---------------------------------------------------------------- 13. outlook: male CNS
s = new_slide()
kicker(s, "Outlook")
set_title(s, LAY, "Next: from a brain-only graph to brain + nerve cord")
iw = 5.35
ih = iw * 2265 / 3081
pic(s, A("malecns_dorsal.png"), 0.4, 1.0, iw, ih, name="Male CNS rendering (FlyEM v1.0)")
tb(s, 0.55, 1.0, 1.2, 0.25, [[("BRAIN", {"bold": True, "spc": 100})]], size=9.5, color=PURPLE)
tb(s, 1.85, 2.5, 1.0, 0.25, [[("NECK", {"bold": True, "spc": 100})]], size=9.5, color=PURPLE, align="c")
tb(s, 3.55, 1.56, 2.2, 0.25, [[("VENTRAL NERVE CORD", {"bold": True, "spc": 100})]], size=9.5, color=PURPLE, align="r")
leg = [("E06A1B", "Descending (brain → VNC)"), ("2B7BBF", "Ascending (VNC → brain)"), ("7B3FA0", "VNC motor neurons")]
lx = 0.5
for c, t in leg:
    box(s, lx, 4.98, 0.13, 0.13, fill=c, shape=MSO_SHAPE.OVAL)
    tb(s, lx + 0.2, 4.92, 1.75, 0.25, t, size=9, color=BODY, anchor="m")
    lx += 1.78
tb(s, 6.05, 1.08, 3.45, 0.5, [[("166,700", {"bold": True})]], size=30, color=PURPLE)
tb(s, 6.05, 1.62, 3.45, 0.45, "neurons in the first complete male central nervous system: brain and ventral nerve cord [8]",
   size=10.5, color=BODY, line=1.03)
tb(s, 6.05, 2.16, 3.45, 0.28, [[("1,314", {"bold": True, "color": "E06A1B"}), (" descending  ·  ", {}),
                               ("1,846", {"bold": True, "color": "2B7BBF"}), (" ascending neurons", {})]],
   size=10.5, color=BODY)
tb(s, 6.05, 2.62, 3.45, 0.25, [[("WHAT IT WOULD ADD TO FLYGM", {"bold": True, "spc": 100})]], size=9, color=MUTED)
tb(s, 6.05, 2.9, 3.45, 1.45, [
    {"runs": [("Real premotor and motor circuits ", {"bold": True}), ("between descending neurons and the legs", {})],
     "bullet": VERM, "after": 4},
    {"runs": [("Ascending feedback ", {"bold": True}), ("from legs and wings through measured pathways", {})],
     "bullet": VERM, "after": 4},
    {"runs": [("Richer annotations: ", {"bold": True}), ("updated synapses [23], sex-specific cell types [24]", {})],
     "bullet": VERM},
], size=10.5, color=BODY, line=1.04)
box(s, 6.05, 4.42, 3.45, 0.62, fill=TINT, radius=0.12)
tb(s, 6.18, 4.42, 3.25, 0.62, [[("Next tests  ", {"bold": True, "color": PURPLE}),
                                ("longer sequences · perturbation recovery · navigation", {})]],
   size=10, color=BODY, anchor="m", line=1.02)
footer(s, "Outlook, not a FlyGM result. Rendered from the public FlyEM male CNS v1.0 release: brain/VNC shells and ~280 sampled neurons [8].")
slide_number(s)
notes(s, """Target time: 50 seconds.

Where next? FlyGM currently uses the female brain connectome. The brain does not control the legs directly: descending neurons send commands through the neck to the ventral nerve cord, where premotor and motor circuits generate the actual leg and wing movements. In FlyGM, that last step is played by a learned decoder.

This image is rendered from the public male CNS release published this year: the brain on the left, the ventral nerve cord on the right, joined at the neck. In orange are a sample of the 1,314 descending neurons, in blue ascending neurons that carry feedback from the body back to the brain, and in purple motor neurons. The full dataset has 166,700 neurons.

For FlyGM, this would replace part of the learned decoder with measured premotor and motor wiring, and route body feedback through real ascending pathways. Together with updated synapse detections and sex-specific annotations, it would let us test connectome-based control on longer sequences, recovery from perturbations, and navigation. These are directions, not results.""")
order.append(s)

# ---------------------------------------------------------------- 14. takeaways
s = new_slide()
kicker(s, "Takeaways")
set_title(s, LAY, "Brain wiring is a useful prior for whole-body control")
cards = [
    ("tree_w.png", "Wiring as architecture",
     "The 139k-neuron FlyWire graph becomes a trainable whole-body controller for walking, turning and flight."),
    ("speed_w.png", "The real wiring matters",
     "Faster imitation and lower steering error than rewired, random and MLP controls; 2.3× the reward of standard GNNs."),
    ("insights_w.png", "Structured dynamics",
     "Neuron classes show distinct, behavior-locked latent patterns that a random graph lacks (qualitative)."),
]
cw, gap = 2.86, 0.21
for i, (ic, head, d) in enumerate(cards):
    x = 0.5 + i * (cw + gap)
    box(s, x, 1.15, cw, 2.55, fill=(TINT_V if i == 1 else TINT), radius=0.06)
    icon_circle(s, x + 0.22, 1.35, 0.62, (VERM if i == 1 else PURPLE), ic)
    tb(s, x + cw - 0.75, 1.3, 0.55, 0.55, [[(str(i + 1), {"bold": True})]], size=26, color="D9D2E3", align="r")
    tb(s, x + 0.22, 2.1, cw - 0.4, 0.32, [[(head, {"bold": True})]], size=14, color=INK)
    tb(s, x + 0.22, 2.48, cw - 0.4, 1.15, d, size=11.5, color=BODY, line=1.06)
box(s, 0.5, 3.92, 9.0, 0.62, fill=INK, radius=0.12)
tb(s, 0.72, 3.92, 8.6, 0.62, [[("Not an upload, ", {"bold": True, "color": "F0A46B"}),
                               ("but a testable, connectome-structured controller.", {"bold": True})]],
   size=15, color="FFFFFF", anchor="m")
tb(s, 0.5, 4.66, 9.0, 0.4, [[("Open:  ", {"bold": True, "color": PURPLE}),
                             ("biological fidelity of the learned dynamics  ·  learning without demonstrations  ·  other bodies and simulators", {})]],
   size=11, color=BODY, anchor="m")
footer(s, "FlyGM [1]. Comparisons use the shared IL + PPO pipeline; representation analysis is qualitative (Appendix H).")
slide_number(s)
notes(s, """Target time: 45 seconds.

To summarize. First, a whole-brain connectome can serve as the architecture of a controller: FlyGM turns the 139-thousand-neuron FlyWire graph into a trainable policy that walks, turns and flies a detailed simulated body. Second, the real wiring matters: under the same training pipeline it learns faster and steers more accurately than rewired, random and MLP controls, and than five standard GNNs. Third, its internal dynamics are structured by neuron class, which a random graph does not show, though this analysis is qualitative.

To come back to the debate: this is not an upload of a fly. It is something more modest but testable, a connectome-structured controller whose assumptions are explicit. What remains open is the biological fidelity of the learned dynamics, learning without expert demonstrations, and transfer to other bodies and simulators.""")
order.append(s)

# ---------------------------------------------------------------- 15. thank you / demo page
s = new_slide()
dark_bg(s)
shot = A("site_screenshot.png")
set_title(s, LAY, "Thank you", size=40, color="FFFFFF", x=0.6, y=0.8, w=4.0, h=0.8)
tb(s, 0.6, 1.62, 4.0, 0.35, "Questions are welcome.", size=16, color="E6DEF3")
tb(s, 0.6, 2.3, 4.0, 0.3, [[("SEE THE DEMOS", {"bold": True, "spc": 150})]], size=10, color="F0A46B")
tb(s, 0.6, 2.6, 4.2, 0.35, [[("lnsgroup.cc/research/FlyGM", {"bold": True})]], size=17, color="FFFFFF")
tb(s, 0.6, 2.98, 4.0, 0.5, "Walking, turning and flight videos with live connectome activity.", size=11.5, color="B9AECB", line=1.05)
box(s, 0.6, 3.72, 1.35, 1.35, fill="FFFFFF", radius=0.1)
pic(s, A("qr_flygm.png"), 0.68, 3.8, 1.19, 1.19, name="QR code: lnsgroup.cc/research/FlyGM",
    link="https://lnsgroup.cc/research/FlyGM/")
tb(s, 2.12, 3.72, 2.6, 1.35, [[("Zehao Jin, Yaoye Zhu,", {"bold": True})], [("Chen Zhang, Yanan Sui", {"bold": True})],
                              [("Tsinghua University · LNS Group", {"color": "B9AECB", "bold": False})]],
   size=11, color="FFFFFF", anchor="m", line=1.15)
FX, FY, FW = 5.0, 0.85, 4.5
FH = 4.2
if os.path.exists(shot):
    # browser frame around a real screenshot of the demo page
    box(s, FX, FY, FW, FH, fill="FFFFFF", radius=0.03, name="Browser frame")
    box(s, FX, FY, FW, 0.32, fill="E9E5EF", radius=0.2, name="Browser bar")
    box(s, FX, FY + 0.16, FW, 0.16, fill="E9E5EF", radius=0.0, shape=MSO_SHAPE.RECTANGLE)
    for k, c in enumerate(["F26B5B", "F5BF4F", "5FC25D"]):
        box(s, FX + 0.14 + k * 0.17, FY + 0.105, 0.11, 0.11, fill=c, shape=MSO_SHAPE.OVAL)
    box(s, FX + 0.75, FY + 0.06, FW - 0.95, 0.2, fill="FFFFFF", radius=0.5)
    tb(s, FX + 0.85, FY + 0.06, FW - 1.1, 0.2, "lnsgroup.cc/research/FlyGM/", size=8, color=MUTED, anchor="m")
    from PIL import Image
    iw_, ih_ = Image.open(shot).size
    vw_ = FW - 0.04
    vh_ = FH - 0.36
    p = pic(s, shot, FX + 0.02, FY + 0.34, vw_, vw_ * ih_ / iw_, name="Demo page screenshot",
            link="https://lnsgroup.cc/research/FlyGM/")
    # crop to frame height
    if vw_ * ih_ / iw_ > vh_:
        frac = 1 - vh_ / (vw_ * ih_ / iw_)
        p.crop_bottom = frac
        p.height = Inches(vh_)
else:
    # no screenshot of the page available at build time: show stills from the demo videos instead
    stills = [(M("image9.png"), "Gait initiation and walking"), (M("image10.png"), "Flight"),
              (M("image12.png"), "Walk → pause → turn, with model activity")]
    box(s, FX, FY, FW, FH, fill=DARK_CARD, radius=0.04, name="Demo stills panel")
    tb(s, FX + 0.25, FY + 0.2, FW - 0.5, 0.25, [[("STILLS FROM THE DEMO VIDEOS", {"bold": True, "spc": 120})]],
       size=9, color="C9B8E8")
    sw = (FW - 0.6) / 2
    for k, (img, cap) in enumerate(stills[:2]):
        x = FX + 0.25 + k * (sw + 0.1)
        pic(s, img, x, FY + 0.55, sw, sw * 1080 / 1920, link="https://lnsgroup.cc/research/FlyGM/")
        tb(s, x, FY + 0.6 + sw * 1080 / 1920, sw, 0.22, cap, size=8.5, color="B9AECB")
    bw_ = FW - 0.5
    pic(s, stills[2][0], FX + 0.25, FY + 1.95, bw_, bw_ * 850 / 1920, link="https://lnsgroup.cc/research/FlyGM/")
    tb(s, FX + 0.25, FY + 2.0 + bw_ * 850 / 1920, bw_, 0.22, stills[2][1], size=8.5, color="B9AECB")
notes(s, """Target time: 15 seconds.

Thank you. The demo page, at lnsgroup.cc slash research slash FlyGM, has the walking, turning and flight videos with live connectome activity; the QR code links there. I am happy to take questions.""")
order.append(s)

# ---------------------------------------------------------------- 16. references and resources
s = new_slide()
kicker(s, "Backup")
set_title(s, LAY, "References and resources")
res = [
    ("Demo page", "lnsgroup.cc/research/FlyGM", "https://lnsgroup.cc/research/FlyGM/"),
    ("Paper (preprint)", "arXiv:2602.17997", "https://arxiv.org/abs/2602.17997"),
    ("Connectome data", "codex.flywire.ai · male-cns.janelia.org", "https://codex.flywire.ai"),
    ("Body simulator", "github.com/TuragaLab/flybody", "https://github.com/TuragaLab/flybody"),
]
box(s, 0.5, 1.08, 9.0, 0.58, fill=TINT, radius=0.12)
rx = 0.68
for lab, txt, url in res:
    w_ = 0.3 + 0.062 * max(len(lab), len(txt))
    tb(s, rx, 1.13, w_, 0.48, [[(lab.upper(), {"bold": True, "spc": 80, "size": 8, "color": PURPLE})],
                               [(txt, {"bold": True, "link": url})]], size=10, color=INK, line=1.05)
    rx += w_ + 0.1
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
    "T. N. Kipf, M. Welling. Semi-supervised classification with graph convolutional networks. ICLR (2017).",
    "Y. Wang et al. Dynamic graph CNN for learning on point clouds. ACM Trans. Graph. 38, 1–12 (2019).",
    "W. Hamilton, Z. Ying, J. Leskovec. Inductive representation learning on large graphs. NeurIPS 30 (2017).",
    "P. Veličković et al. Graph attention networks. ICLR (2018).",
    "G. Corso et al. Principal neighbourhood aggregation for graph nets. NeurIPS 33, 13260–13271 (2020).",
    "B. E. Brezovec et al. Mapping the neural dynamics of locomotion across the Drosophila brain. Curr. Biol. 34, 710–726 (2024).",
    "E. S. Schaffer et al. The spatial and temporal structure of neural activity across the fly brain. Nat. Commun. 14, 5572 (2023).",
    "S.-C. Yu et al. New synapse detection in the whole-brain connectome of Drosophila. bioRxiv (2025).",
    "D. Deutsch et al. Sexually-dimorphic neurons in the Drosophila whole-brain connectome. Research Square (2025).",
]
half = 12
for col in range(2):
    x = 0.5 + col * 4.6
    paras = []
    for k in range(col * half, min(len(refs), (col + 1) * half)):
        paras.append({"runs": [(f"[{k + 1}]  ", {"bold": True, "color": PURPLE}), (refs[k], {})], "after": 2.2})
    tb(s, x, 1.8, 4.4, 3.35, paras, size=8, color=BODY, line=1.0)
slide_number(s)
notes(s, """Backup slide.

Numbered references cited on the slides. [1] is our paper; the demo page is lnsgroup.cc/research/FlyGM. [9]-[11] are the 2026 public discussion of the Eon Systems fly demo. GNN baselines [16]-[20] correspond to GCN, EdgeCNN (DGCNN), GraphSAGE, GAT and PNA. The male CNS rendering on the outlook slide was produced from the public FlyEM v1.0 release (gs://flyem-male-cns).""")
order.append(s)

# ---------------------------------------------------------------- reorder and drop unused V1 slides
sldIdLst = prs.slides._sldIdLst
ids = {sld.get("id"): sld for sld in list(sldIdLst)}
by_slide = {}
for sld in list(sldIdLst):
    rId = sld.get(qn("r:id"))
    by_slide[id(prs.part.related_part(rId))] = sld
keep = [by_slide[id(sl.part)] for sl in order]
for sld in list(sldIdLst):
    if sld not in keep:
        prs.part.drop_rel(sld.get(qn("r:id")))
    sldIdLst.remove(sld)
for sld in keep:
    sldIdLst.append(sld)

prs.core_properties.title = "FlyGM: Whole-Brain Connectomic Graph Model for Whole-Body Locomotion Control"
prs.core_properties.subject = "AIAS+ 2026 oral presentation"
prs.save(OUT)
print("saved", OUT, "slides:", len(prs.slides))
