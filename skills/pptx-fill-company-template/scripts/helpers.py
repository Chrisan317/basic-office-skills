"""Reusable helpers for rebuilding a deck on a company template with python-pptx.
usage:  import helpers as h; from helpers import *; src, prs = init(SRC, TPL) ... save(OUT)

顏色：init() 會從範本的佈景主題（theme）自動讀主色，更新 BLUE/NAVY/CYAN/CLOUD/ICE/ZEBRA。
      因為 `from helpers import *` 拿到的是 import 當下的值，init 之後請用 h.BLUE、h.CLOUD 這種寫法取顏色。
版面：new_slide() 可以給範本的版面名稱，也可以給種類 "cover" / "content" / "section" / "closing"，
      會自動在範本裡找最適合的版面（找不到 closing/section 時回傳 None）。
區域：content_area(slide) 回傳 (x, y, w, h)，依投影片大小與標題位置計算，16:9、4:3 都適用。"""
import re, math, copy, sys
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.oxml.ns import qn
from lxml import etree

BLUE, CYAN, ICE, CLOUD, NAVY, GRAY, BLACK, WHITE = "0055CC", "00B0F0", "9DC3E6", "DAE3F3", "002145", "767171", "000000", "FFFFFF"
ZEBRA = "F2F6FC"
RED = "E03C31"

def rgb(h): return RGBColor.from_string(h)


def _mix(hex_color, white_ratio):
    """把顏色和白色混合，white_ratio=0.85 代表 85% 白（做淺色卡片底色）。"""
    r, g, b = (int(hex_color[i:i + 2], 16) for i in (0, 2, 4))
    return "%02X%02X%02X" % tuple(round(c + (255 - c) * white_ratio) for c in (r, g, b))


def theme_colors(pres):
    """讀範本佈景主題的顏色，回傳 {'accent1': 'RRGGBB', 'dk2': ..., ...}。讀不到回傳 {}。"""
    try:
        from pptx.opc.constants import RELATIONSHIP_TYPE as RT
        theme = pres.slide_masters[0].part.part_related_by(RT.THEME)
        root = etree.fromstring(theme.blob)
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        out = {}
        for el in root.find(".//a:clrScheme", ns):
            name = el.tag.split("}")[1]
            c = el.find("a:srgbClr", ns)
            if c is not None:
                out[name] = c.get("val").upper()
            else:
                s = el.find("a:sysClr", ns)
                if s is not None and s.get("lastClr"):
                    out[name] = s.get("lastClr").upper()
        return out
    except Exception:
        return {}


def apply_theme_palette(pres):
    """用範本的主題色更新調色盤：accent1 = 主色，dk2 = 深色標題，accent2 = 輔色。"""
    global BLUE, CYAN, ICE, CLOUD, NAVY, ZEBRA
    th = theme_colors(pres)
    if th.get("accent1"):
        BLUE = th["accent1"]
        CLOUD, ICE, ZEBRA = _mix(BLUE, 0.85), _mix(BLUE, 0.6), _mix(BLUE, 0.94)
    if th.get("dk2"):
        NAVY = th["dk2"]
    if th.get("accent2"):
        CYAN = th["accent2"]
    return th

# ---------- state (set by init) ----------
src = None          # source deck (content is read from here)
prs = None          # output deck, opened from the template
L = {}              # template layout name -> layout
orig_ids = []       # template sample slides, removed by save()
warn = []           # overflow warnings collected by txt()

def init(src_path, tpl_path):
    """Open source + template. Returns (src, prs). Call once before anything else.
    src_path may be None when the source is not a pptx (use extract_source.py for the content)."""
    global src, prs, L, orig_ids
    src = Presentation(src_path) if src_path else None
    prs = Presentation(tpl_path)
    orig_ids = list(prs.slides._sldIdLst)
    L = {l.name: l for l in prs.slide_layouts}
    warn.clear()
    apply_theme_palette(prs)
    return src, prs

def save(out_path):
    """Drop the template's sample slides, save, print warnings."""
    for sid in orig_ids:
        prs.part.drop_rel(sid.rId)
        prs.slides._sldIdLst.remove(sid)
    prs.save(out_path)
    print("saved", out_path, "slides:", len(prs.slides))
    for w_ in warn: print("WARN", w_)

# ---------- helpers ----------
def is_cjk(ch): return ord(ch) > 0x2E80
def text_w(s, size):  # points
    return sum(size if is_cjk(c) else size * 0.55 for c in s)
def est_h(paras, w_in):
    """paras: list of (text,size,space_after_pt)"""
    h = 0
    for t, sz, sa in paras:
        lines = max(1, math.ceil(text_w(t, sz) / max(w_in * 72, 1)))
        h += lines * sz * 1.25 + sa
    return h / 72

KIND_NAMES = {
    "cover":   ["Cover", "Title Slide", "標題投影片", "封面"],
    "content": ["Title Only", "僅標題", "Infographic / Table", "Title and Content", "標題及內容"],
    "section": ["Section Header", "區段標題", "章節標題"],
    "closing": ["Closing", "Closing Slide", "Closing Slide_Dark", "End", "Thank You", "結尾", "謝謝"],
}
_KIND_WORDS = {"cover": ("cover", "title slide", "封面"), "section": ("section", "區段", "章節"),
               "closing": ("closing", "end", "thank", "結尾", "謝謝")}


def _ph_types(layout):
    return [ph.placeholder_format.type for ph in layout.placeholders]


def find_layout(kind):
    """依種類在範本中找版面：先比對常見名稱，再看名稱關鍵字，最後看預留位置的種類。"""
    for n in KIND_NAMES.get(kind, []):
        if n in L:
            return L[n]
    for name, lay in L.items():
        if any(w in name.lower() for w in _KIND_WORDS.get(kind, ())):
            return lay
    from pptx.enum.shapes import PP_PLACEHOLDER as P
    deco = {P.DATE, P.FOOTER, P.SLIDE_NUMBER}
    lays = list(prs.slide_layouts)
    if kind == "cover":
        return next((l for l in lays if P.CENTER_TITLE in _ph_types(l)), lays[0])
    if kind == "content":
        title_only = [l for l in lays if P.TITLE in _ph_types(l)
                      and not (set(_ph_types(l)) - deco - {P.TITLE})]
        if title_only:
            return title_only[0]
        return next((l for l in lays if P.TITLE in _ph_types(l)), lays[min(1, len(lays) - 1)])
    return None  # section / closing：範本沒有就不做


def layout_of(layout):
    if layout in L:
        return L[layout]
    if layout in KIND_NAMES:
        return find_layout(layout)
    raise KeyError(f"範本沒有版面「{layout}」。可用的版面：{list(L)}")


def new_slide(layout, title=None):
    """layout 可以是版面名稱，或種類 cover/content/section/closing。種類找不到時回傳 None。"""
    lay = layout_of(layout)
    if lay is None:
        return None
    s = prs.slides.add_slide(lay)
    if title is not None and s.shapes.title is not None:
        s.shapes.title.text = title
    return s


def set_subtitle(s, text):
    """把文字放進第一個副標題/內文預留位置（不同範本的 idx 不同，不要寫死 idx）。"""
    from pptx.enum.shapes import PP_PLACEHOLDER as P
    for ph in s.placeholders:
        if ph.placeholder_format.type in (P.SUBTITLE, P.BODY) and ph.has_text_frame:
            ph.text = text
            return ph
    return None


def drop_empty_placeholders(s):
    for ph in list(s.placeholders):
        if not ph.has_text_frame or not ph.text_frame.text.strip():
            ph._element.getparent().remove(ph._element)


def slide_size():
    return prs.slide_width / 914400, prs.slide_height / 914400


def content_area(s=None, margin=0.5, bottom=0.45):
    """回傳可放內容的 (x, y, w, h)（英吋）：左右留 margin，從標題下方開始。"""
    W, H = slide_size()
    top = 1.2
    if s is not None and s.shapes.title is not None and s.shapes.title.top is not None:
        t = s.shapes.title
        top = max(top * H / 7.5, (t.top + t.height) / 914400 + 0.15)
    return margin, top, W - 2 * margin, H - top - bottom


def box(s, x, y, w, h, fill=None, radius=0.05, line=None):
    fill = fill or CLOUD
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.adjustments[0] = radius
    sh.fill.solid(); sh.fill.fore_color.rgb = rgb(fill)
    if line: sh.line.color.rgb = rgb(line); sh.line.width = Pt(0.75)
    else: sh.line.fill.background()
    sh.shadow.inherit = False
    return sh

def txt(s, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT, margin=0.05, name=None):
    """paras: list of dict(text|runs, size, bold, color, after, align)"""
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(margin); tf.margin_top = tf.margin_bottom = Inches(0.03)
    tf.vertical_anchor = anchor
    est = []
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = p.get("align", align)
        para.space_after = Pt(p.get("after", 0))
        runs = p.get("runs") or [dict(text=p["text"])]
        full = ""
        for r in runs:
            run = para.add_run(); run.text = r["text"]; full += r["text"]
            f = run.font; f.size = Pt(r.get("size", p.get("size", 14)))
            f.bold = r.get("bold", p.get("bold", False))
            f.color.rgb = rgb(r.get("color", p.get("color", BLACK)))
        est.append((full, p.get("size", 14), p.get("after", 0)))
    need = est_h(est, w - 2 * margin) + 0.06
    if need > h + 0.02:
        warn.append(f"overflow? '{est[0][0][:20]}' need {need:.2f} > {h:.2f}")
    if name: tb.name = name
    return tb

def set_border(cell, side, color, w=6350):
    tcPr = cell._tc.get_or_add_tcPr()
    tag = {"B": "a:lnB", "T": "a:lnT", "L": "a:lnL", "R": "a:lnR"}[side]
    for e in tcPr.findall(qn(tag)): tcPr.remove(e)
    ln = etree.SubElement(tcPr, qn(tag), w=str(w), cap="flat", cmpd="sng", algn="ctr")
    if color is None:
        etree.SubElement(ln, qn("a:noFill"))
    else:
        sf = etree.SubElement(ln, qn("a:solidFill")); etree.SubElement(sf, qn("a:srgbClr"), val=color)
    # schema order: lnL lnR lnT lnB must precede fill -> move fills to end
    for f in tcPr.findall(qn("a:solidFill")) + tcPr.findall(qn("a:noFill")):
        tcPr.remove(f); tcPr.append(f)

def cell_set(cell, text, size, bold=False, color=BLACK, fill=None, align=PP_ALIGN.LEFT, ml=0.06, mt=0.02):
    tf = cell.text_frame; tf.word_wrap = True
    lines = text if isinstance(text, list) else [text]
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run(); r.text = ln
        r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = rgb(color)
        if not ln:  # empty run: size lives on endParaRPr, otherwise the row grows to 18pt
            p._p.remove(r._r)
        for e in p._p.findall(qn("a:endParaRPr")): p._p.remove(e)
        etree.SubElement(p._p, qn("a:endParaRPr"), lang="zh-TW", sz=str(int(size * 100)))
    cell.margin_left = cell.margin_right = Inches(ml)
    cell.margin_top = cell.margin_bottom = Inches(mt)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    if fill: cell.fill.solid(); cell.fill.fore_color.rgb = rgb(fill)
    else: cell.fill.background()

def table(s, x, y, widths, nrows, row_hs, name=None):
    tot = sum(widths)
    gf = s.shapes.add_table(nrows, len(widths), Inches(x), Inches(y), Inches(tot), Inches(sum(row_hs)))
    t = gf.table
    tblPr = t._tbl.tblPr
    for a in ("firstRow", "bandRow"): tblPr.attrib.pop(a, None)
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is not None: sid.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"  # No Style, No Grid
    for i, w in enumerate(widths): t.columns[i].width = Inches(w)
    for i, h in enumerate(row_hs): t.rows[i].height = Inches(h)
    if name: gf.name = name
    return t

def fit(widths, total):
    k = total / sum(widths); return [w * k for w in widths]

def norm(s): return re.sub(r"\s+", " ", s.replace("\x0b", " ")).strip()

def src_table(slide_no, nth=0):
    """nth table on source slide `slide_no` (1-based)."""
    ts = [sh for sh in src.slides[slide_no - 1].shapes if sh.has_table]
    return ts[nth].table
def cell_lines(c):
    return [norm(p.text) for p in c.text_frame.paragraphs if norm(p.text)]

def style_header_row(t, r, labels, size, aligns=None, fill=None):
    for ci, lb in enumerate(labels):
        c = t.cell(r, ci)
        cell_set(c, lb, size, True, WHITE, fill or BLUE, (aligns or {}).get(ci, PP_ALIGN.CENTER))

def zebra_body(t, r0, rows, size, aligns=None, bold_last=False, fills=True):
    n = len(rows)
    for ri, row in enumerate(rows):
        last = bold_last and ri == n - 1
        f = CLOUD if last else (ZEBRA if ri % 2 == 1 else None)
        for ci, v in enumerate(row):
            c = t.cell(r0 + ri, ci)
            cell_set(c, v, size, last, BLACK, f, (aligns or {}).get(ci, PP_ALIGN.LEFT))
            set_border(c, "B", CLOUD, 6350)

C = PP_ALIGN.CENTER


def walk_shapes(shs):
    """Yield every non-group shape, descending into groups."""
    for sh_ in shs:
        if sh_.shape_type == 6: yield from walk_shapes(sh_.shapes)
        else: yield sh_

def bar_line_combo(chart, line_series_index, color=BLUE, labels=True):
    """Turn one series of a clustered bar/column chart into a line on the same axes (combo chart)."""
    C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
    bar = chart._chartSpace.find(".//{%s}barChart" % C_NS)
    tot = bar.findall("{%s}ser" % C_NS)[line_series_index]; bar.remove(tot)
    for tag in ("invertIfNegative", "spPr"):
        for e in tot.findall("{%s}%s" % (C_NS, tag)): tot.remove(e)
    spPr = etree.fromstring('<c:spPr xmlns:c="%s" xmlns:a="%s"><a:ln w="31750" cap="rnd"><a:solidFill><a:srgbClr val="%s"/></a:solidFill><a:round/></a:ln></c:spPr>' % (C_NS, A_NS, color))
    marker = etree.fromstring('<c:marker xmlns:c="%s" xmlns:a="%s"><c:symbol val="circle"/><c:size val="7"/><c:spPr><a:solidFill><a:srgbClr val="%s"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr></c:marker>' % (C_NS, A_NS, color))
    tot.find("{%s}tx" % C_NS).addnext(spPr); spPr.addnext(marker)
    if labels:
        dl_ = etree.fromstring('<c:dLbls xmlns:c="%s" xmlns:a="%s"><c:numFmt formatCode="0" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="1"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="zh-TW"/></a:p></c:txPr><c:dLblPos val="t"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>' % (C_NS, A_NS, color))
        marker.addnext(dl_)
    etree.SubElement(tot, "{%s}smooth" % C_NS).set("val", "0")
    line = etree.Element("{%s}lineChart" % C_NS)
    etree.SubElement(line, "{%s}grouping" % C_NS).set("val", "standard")
    etree.SubElement(line, "{%s}varyColors" % C_NS).set("val", "0")
    line.append(tot)
    etree.SubElement(line, "{%s}marker" % C_NS).set("val", "1")
    for ax in bar.findall("{%s}axId" % C_NS):
        etree.SubElement(line, "{%s}axId" % C_NS).set("val", ax.get("val"))
    bar.addnext(line)

