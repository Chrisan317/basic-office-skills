"""把來源簡報（pptx）自動套到公司範本：分析每頁 → 寫出頁面計畫 plan.json → 依計畫產生 → 自動檢查。

用法：
  python auto_build.py 來源.pptx 範本.pptx 輸出.pptx
      第一次執行：分析來源，寫出 _work/plan.json，產生輸出並檢查。
  python auto_build.py 來源.pptx 範本.pptx 輸出.pptx --plan _work/plan.json
      照修改過的計畫重新產生（只改 plan.json，不要改這支程式）。
  其他選項：--work 資料夾（預設是輸出檔旁邊的 _work）、--no-png（不匯出圖片）

文字一律在產生時從來源讀，計畫裡只記「哪一頁、用什麼版型」，所以不會漏字、也不會出現自己編的字。
"""
import argparse
import hashlib
import io
import json
import math
import os
import re
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import helpers as H  # noqa: E402
from pptx.util import Inches, Pt  # noqa: E402
from pptx.enum.text import PP_ALIGN  # noqa: E402

AGENDA_WORDS = ("agenda", "目錄", "大綱", "outline", "contents", "議程")
KINDS = ("cover", "agenda", "table", "chart", "text", "graphic", "closing", "section", "custom", "skip")


# ---------------------------------------------------------------- 讀來源
def _all_shapes(shapes, tf=(0.0, 0.0, 1.0, 1.0)):
    """所有形狀（含群組內），回傳 (shape, 左, 上, 寬, 高) 吋。"""
    E = H.EMU
    ox, oy, sx, sy = tf
    for sh in shapes:
        l, t = (sh.left or 0) / E, (sh.top or 0) / E
        w, h = (sh.width or 0) / E, (sh.height or 0) / E
        L_, T_, W_, H_ = ox + l * sx, oy + t * sy, w * sx, h * sy
        try:
            grp = sh.shape_type == 6
        except NotImplementedError:
            grp = False
        if grp:
            x = sh._element.find(".//" + H.qn("a:chOff")); e = sh._element.find(".//" + H.qn("a:chExt"))
            cx, cy = (int(x.get("x")) / E, int(x.get("y")) / E) if x is not None else (0, 0)
            cw, ch = (int(e.get("cx")) / E, int(e.get("cy")) / E) if e is not None else (w, h)
            nsx, nsy = (W_ / cw if cw else sx), (H_ / ch if ch else sy)
            yield from _all_shapes(sh.shapes, (L_ - cx * nsx, T_ - cy * nsy, nsx, nsy))
        else:
            yield sh, L_, T_, W_, H_


def _max_size(sh):
    ss = [r.font.size.pt for p in sh.text_frame.paragraphs for r in p.runs if r.font.size]
    return max(ss) if ss else 0


def _clean(t):
    """去掉前後的裝飾符號（例如「- 名字-」）。"""
    return re.sub(r"^[\s\-–—_~·•]+|[\s\-–—_~·•]+$", "", t)


DECOR = set()  # 很多頁都出現的圖（logo、裝飾），不搬


def _is_pic(sh):
    try:
        return sh.shape_type == 13 or (sh.is_placeholder and hasattr(sh, "image") and sh.image is not None)
    except Exception:
        return False


def _hash(sh):
    try:
        return hashlib.md5(sh.image.blob).hexdigest()
    except Exception:
        return None


def find_decor(src):
    """出現在 3 頁以上、且超過 4 成頁面的同一張圖，當成裝飾（logo、底圖）。"""
    from collections import Counter
    cnt = Counter()
    for sl in src.slides:
        cnt.update({_hash(sh) for sh, *_ in _all_shapes(sl.shapes) if _is_pic(sh)} - {None})
    n = len(src.slides)
    DECOR.clear()
    DECOR.update(k for k, v in cnt.items() if v >= 3 and v >= 0.4 * n)


def pictures(n):
    """來源第 n 頁要搬的圖片：夠大（0.6 平方吋以上）、不是整頁底圖、不是每頁都有的裝飾圖。依位置排序。"""
    sl = H.src.slides[n - 1]
    W, Hh = H.src.slide_width / H.EMU, H.src.slide_height / H.EMU
    out = []
    for sh, l, t, w, h in _all_shapes(sl.shapes):
        if not _is_pic(sh) or w * h < 0.6 or w * h > 0.85 * W * Hh or _hash(sh) in DECOR:
            continue
        if l < -0.2 or t < -0.2 or l + w > W + 0.2 or t + h > Hh + 0.2:  # 超出投影片邊緣的多半是裝飾圖
            continue
        edges = sum([l < 0.05, t < 0.05, l + w > W - 0.05, t + h > Hh - 0.05])
        if edges >= 2:  # 貼著投影片角落的圖多半是裝飾
            continue
        out.append(dict(sh=sh, l=l, t=t, w=w, h=h))
    out.sort(key=lambda p: (round(p["t"], 1), p["l"]))
    return out


def place_pictures(s, pics, x, y, w, h, gap=0.15):
    """把圖片排進 (x, y, w, h)：試不同欄數，挑圖片總面積最大的排法；保持比例、保留裁切。"""
    n = len(pics)
    best = None
    for cols in range(1, n + 1):
        rows = math.ceil(n / cols)
        cw, ch = (w - gap * (cols - 1)) / cols, (h - gap * (rows - 1)) / rows
        area = sum(min(cw / p["w"], ch / p["h"]) ** 2 * p["w"] * p["h"] for p in pics)
        if best is None or area > best[0]:
            best = (area, cols, rows, cw, ch)
    _, cols, rows, cw, ch = best
    bottom = y
    for i, p in enumerate(pics):
        k = min(cw / p["w"], ch / p["h"])
        pw, ph = p["w"] * k, p["h"] * k
        cx = x + (i % cols) * (cw + gap) + (cw - pw) / 2
        cy = y + (i // cols) * (ch + gap)
        sh = p["sh"]
        pic = s.shapes.add_picture(io.BytesIO(sh.image.blob), Inches(cx), Inches(cy), Inches(pw), Inches(ph))
        for side in ("crop_left", "crop_right", "crop_top", "crop_bottom"):
            try:
                setattr(pic, side, getattr(sh, side))
            except Exception:
                pass
        bottom = max(bottom, cy + ph)
    return bottom


def is_graphic(sl):
    """圖解頁：有連接線、文字放在圖形裡的流程圖、或用很多文字預留位置排的版面（時間軸等）。
    這種頁面重排會失去位置關係，所以整組圖形搬過去。"""
    from pptx.enum.shapes import PP_PLACEHOLDER as P
    conn = text_shapes = text_ph = 0
    for sh, *_ in _all_shapes(sl.shapes):
        tag = sh.element.tag.rsplit("}", 1)[-1]
        if tag == "cxnSp":
            conn += 1
        elif sh.has_text_frame and sh.text_frame.text.strip():
            if sh.is_placeholder and sh.placeholder_format.type not in (P.TITLE, P.CENTER_TITLE) + H._SKIP_PH:
                text_ph += 1
            elif not sh.is_placeholder and sh.shape_type == 1:  # 文字直接寫在圖形裡
                text_shapes += 1
    return conn >= 1 or text_shapes >= 3 or text_ph >= 4


def layout_ph_size(lay, ph, style="titleStyle"):
    """版面預留位置的字級（pt）：預留位置本身 → 母片的標題/內文樣式。找不到回傳 None。"""
    try:
        for tag in ("a:defRPr", "a:rPr", "a:endParaRPr"):
            for r in ph._element.iter(H.qn(tag)):
                if r.get("sz"):
                    return int(r.get("sz")) / 100
        for st in lay.slide_master._element.iter(H.qn("p:" + style)):
            for r in st.iter(H.qn("a:defRPr")):
                if r.get("sz"):
                    return int(r.get("sz")) / 100
    except Exception:
        pass
    return None


def inherited_size(s, ph):
    """預留位置繼承的字級（pt）：先看版面的預留位置，再看母片的標題樣式。找不到回傳 None。"""
    try:
        idx = ph.placeholder_format.idx
        lay_ph = next((x for x in s.slide_layout.placeholders if x.placeholder_format.idx == idx), None)
        for el in ([lay_ph._element] if lay_ph is not None else []):
            for tag in ("a:defRPr", "a:rPr", "a:endParaRPr"):
                for r in el.iter(H.qn(tag)):
                    if r.get("sz"):
                        return int(r.get("sz")) / 100
        master = s.slide_layout.slide_master._element
        for st in master.iter(H.qn("p:titleStyle")):
            for r in st.iter(H.qn("a:defRPr")):
                if r.get("sz"):
                    return int(r.get("sz")) / 100
    except Exception:
        pass
    return None


def fit_text(ph, text, size0):
    """預留位置放不下時把字縮小（最小 18pt）：用寬度算行數、用高度檢查。放得下就不動，保留範本原本的字級。"""
    w, h = ph.width / H.EMU - 0.2, ph.height / H.EMU
    size = size0
    while size > 18:
        lines = max(1, math.ceil(H.text_w(text, size) * 1.08 / max(w * 72, 1)))
        if lines * size * 1.2 / 72 <= h + 0.05:
            break
        size -= 2
    if size < size0:
        for p_ in ph.text_frame.paragraphs:
            for r_ in p_.runs:
                r_.font.size = Pt(size)


# ---------------------------------------------------------------- 用範本自己的版面
def all_layouts():
    return [l for m in H.prs.slide_masters for l in m.slide_layouts]


def _ptype(ph):
    return ph.placeholder_format.type


def _top_title(lay):
    from pptx.enum.shapes import PP_PLACEHOLDER as P
    Hs = H.prs.slide_height
    return any(_ptype(ph) in (P.TITLE, P.CENTER_TITLE) and ph.top is not None and ph.top < Hs * 0.25
               for ph in lay.placeholders)


def pair_slots(lay):
    """版面裡「小標＋內文」成對的預留位置。回傳 (pairs[(小標idx, 內文idx)], 沒配對的文字位置數, 有沒有圖/表/圖表位置)。"""
    from pptx.enum.shapes import PP_PLACEHOLDER as P
    Hs = H.prs.slide_height
    phs = [ph for ph in lay.placeholders if ph.top is not None and ph.width]
    media = [ph for ph in phs if _ptype(ph) in (P.PICTURE, P.CHART, P.TABLE)]
    texts = sorted([ph for ph in phs if _ptype(ph) in (P.BODY, P.OBJECT)], key=lambda p: (p.top, p.left))
    used, pairs = set(), []
    for b in texts:
        if b.placeholder_format.idx in used:
            continue
        for c in texts:
            ci = c.placeholder_format.idx
            if ci in used or c is b:
                continue
            ov = min(b.left + b.width, c.left + c.width) - max(b.left, c.left)
            if c.top >= b.top + b.height * 0.5 and c.top - (b.top + b.height) < Hs * 0.15 \
                    and ov > 0.6 * min(b.width, c.width):
                pairs.append((b.placeholder_format.idx, ci))
                used |= {b.placeholder_format.idx, ci}
                break
    extra = [p for p in texts if p.placeholder_format.idx not in used]
    return pairs, extra, media


def fit_lines(lay, ph, lines, style="bodyStyle", min_size=10):
    """文字放進版面預留位置要用多大的字：放得下回傳 (字級, 要不要改字級)；最小字級也放不下回傳 None。"""
    size0 = layout_ph_size(lay, ph, style) or 18
    w, h = ph.width / H.EMU - 0.2, ph.height / H.EMU
    size = size0
    while True:
        n = sum(max(1, math.ceil(H.text_w(t, size) * 1.1 / max(w * 72, 1))) for t in lines) if lines else 1
        if n * size * 1.25 / 72 + 0.1 <= h:
            return size, size < size0
        if size <= min_size:
            return None
        size = max(min_size, size - 1)


def fill_ph(ph, lines, size=None):
    ph.text_frame.text = "\n".join(lines)
    if size:
        for p in ph.text_frame.paragraphs:
            for r in p.runs:
                r.font.size = Pt(size)


def set_title(s, text):
    """填標題；標題太長放不下時縮小字（短標題保留範本字級）。"""
    t = s.shapes.title
    if t is None:
        return
    t.text = text
    if text and t.width and t.height:
        fit_text(t, text, inherited_size(s, t) or 40)


def add_slide_number(s):
    """範本版面有頁碼標籤（例如「SLIDE ‹#›」）時加到投影片上（python-pptx 新增投影片時不會帶）。"""
    import copy as _copy
    from pptx.enum.shapes import PP_PLACEHOLDER as P
    if any(_ptype(ph) == P.SLIDE_NUMBER for ph in s.placeholders):
        return
    for ph in s.slide_layout.placeholders:
        if _ptype(ph) == P.SLIDE_NUMBER and ph.has_text_frame and ph.text_frame.text.strip():
            s.shapes._spTree.append(_copy.deepcopy(ph._element))
            return


def drop_empty(s):
    """刪掉沒填的預留位置（標題、頁碼、頁尾除外），編輯時才不會看到「按一下以新增文字」。"""
    from pptx.enum.shapes import PP_PLACEHOLDER as P
    keep = (P.TITLE, P.CENTER_TITLE, P.SLIDE_NUMBER, P.FOOTER, P.DATE)
    for ph in list(s.placeholders):
        if _ptype(ph) in keep:
            continue
        has_text = ph.has_text_frame and ph.text_frame.text.strip()
        if not has_text and ph._element.tag.endswith("}sp"):
            ph._element.getparent().remove(ph._element)


def try_pair_layout(e, n):
    """文字卡片頁：找範本裡「小標＋內文」組數剛好的版面，把文字填進範本的預留位置（字型、字級、位置照範本）。"""
    if e.get("template_layout") == "none" or e.get("subtitle"):
        return False
    if pictures(n) and e.get("pictures") != "none":
        return False
    blocks, _, _ = blocks_for(n, e, side_ok=False)
    if any(b["type"] != "text" for b in blocks):
        return False
    groups = [(b.get("head") or "", b["lines"]) for b in blocks]
    cards = [g for g in groups if g[1]]
    if not cards or len(cards) > 6 or any(not g[1] for g in groups):
        return False
    cands = []
    for lay in all_layouts():
        if not _top_title(lay):
            continue
        pairs, extra, media = pair_slots(lay)
        if pairs and not extra and not media and len(pairs) >= len(cards):
            cands.append((len(pairs) - len(cards), lay, pairs))
    for _, lay, pairs in sorted(cands, key=lambda c: c[0]):
        lph = {ph.placeholder_format.idx: ph for ph in lay.placeholders}
        plan = []
        for (hi, bi), (hd, items) in zip(pairs, cards):
            fh = fit_lines(lay, lph[hi], [hd], min_size=12) if hd else (None, False)
            fb = fit_lines(lay, lph[bi], items)
            if fb is None or (hd and fh is None):
                plan = None
                break
            plan.append((hi, bi, hd, items, fh, fb))
        if not plan:
            continue
        s = H.prs.slides.add_slide(lay)
        set_title(s, e.get("title", ""))
        add_slide_number(s)
        sph = {ph.placeholder_format.idx: ph for ph in s.placeholders}
        for hi, bi, hd, items, fh, fb in plan:
            if hd:
                fill_ph(sph[hi], [hd], fh[0] if fh[1] else None)
            fill_ph(sph[bi], items, fb[0] if fb[1] else None)
        drop_empty(s)
        return True
    return False


def try_chart_layout(e, n):
    """圖表頁：用範本有圖表位置的版面，圖表放在那個位置，副標放進旁邊的小標位置。"""
    from pptx.enum.shapes import PP_PLACEHOLDER as P
    if e.get("template_layout") == "none":
        return False
    best = None
    for lay in all_layouts():
        if not _top_title(lay):
            continue
        charts = [ph for ph in lay.placeholders if _ptype(ph) == P.CHART]
        pairs, extra, media = pair_slots(lay)
        if len(charts) != 1 or extra or len(media) != 1 or len(pairs) > 1:
            continue
        area = charts[0].width * charts[0].height
        if best is None or area > best[0]:
            best = (area, lay, charts[0], pairs)
    if not best:
        return False
    _, lay, cph, pairs = best
    blocks, _, _ = blocks_for(n, e, side_ok=False)
    lines = [t for b in blocks if b["type"] == "text" for t in ([b.get("head")] if b.get("head") else []) + b["lines"]]
    s = H.prs.slides.add_slide(lay)
    if s.shapes.title is not None:
        s.shapes.title.text = e.get("title", "")
    add_slide_number(s)
    sph = {ph.placeholder_format.idx: ph for ph in s.placeholders}
    c = sph[cph.placeholder_format.idx]
    area = (c.left / H.EMU, c.top / H.EMU, c.width / H.EMU, c.height / H.EMU)
    c._element.getparent().remove(c._element)
    H.column_chart(s, area=area, **H.src_chart(n))
    if pairs:
        hi, bi = pairs[0]
        lph = {ph.placeholder_format.idx: ph for ph in lay.placeholders}
        if e.get("subtitle"):
            f = fit_lines(lay, lph[hi], [e["subtitle"]], min_size=12)
            fill_ph(sph[hi], [e["subtitle"]], f[0] if f and f[1] else None)
        if lines:
            f = fit_lines(lay, lph[bi], lines) or (10, True)
            fill_ph(sph[bi], lines, f[0] if f[1] else None)
    elif e.get("subtitle"):
        H.set_subtitle(s, e["subtitle"])
    drop_empty(s)
    return True


def try_agenda_layout(title, items):
    """目錄：用範本「編號＋項目」一列一列的版面（項目數剛好的），填進範本的位置。"""
    from pptx.enum.shapes import PP_PLACEHOLDER as P
    W = H.prs.slide_width
    for lay in all_layouts():
        if not _top_title(lay):
            continue
        texts = [ph for ph in lay.placeholders if _ptype(ph) == P.BODY and ph.width]
        small = [p for p in texts if p.width < W * 0.1]
        rows = []
        for sm in sorted(small, key=lambda p: p.top):
            wide = [p for p in texts if p.width > W * 0.3 and abs(p.top - sm.top) < H.prs.slide_height * 0.03
                    and p.left > sm.left]
            if wide:
                rows.append((sm, min(wide, key=lambda p: p.left)))
        others = [p for p in texts if all(p is not a and p is not b for a, b in rows)]
        if len(rows) != len(items) or len(others) > 1 or any(_ptype(p) in (P.PICTURE, P.CHART, P.TABLE)
                                                              for p in lay.placeholders):
            continue
        if any(fit_lines(lay, b, [it], min_size=12) is None for (_, b), it in zip(rows, items)):
            continue
        s = H.prs.slides.add_slide(lay)
        set_title(s, title)
        add_slide_number(s)
        sph = {ph.placeholder_format.idx: ph for ph in s.placeholders}
        for i, ((a, b), it) in enumerate(zip(rows, items), 1):
            fill_ph(sph[a.placeholder_format.idx], [f"{i:02d}"])
            f = fit_lines(lay, b, [it], min_size=12)
            fill_ph(sph[b.placeholder_format.idx], [it], f[0] if f[1] else None)
        drop_empty(s)
        return True
    return False


def _set_xfrm(el, x, y, w, h):
    """直接改 XML 的位置大小（適用一般圖形、群組、圖片、表格框）。沒有 xfrm（原本繼承版面）就補一個。"""
    from lxml import etree
    xf = None
    for tag in ("p:spPr", "p:grpSpPr"):
        pr = el.find(H.qn(tag))
        if pr is not None:
            xf = pr.find(H.qn("a:xfrm"))
            if xf is None:
                xf = etree.SubElement(pr, H.qn("a:xfrm"))
                pr.insert(0, xf)
            break
    if xf is None:
        xf = el.find(H.qn("p:xfrm"))
    if xf is None:
        return
    off = xf.find(H.qn("a:off")); ext = xf.find(H.qn("a:ext"))
    if off is None:
        off = etree.SubElement(xf, H.qn("a:off")); xf.insert(0, off)
    if ext is None:
        ext = etree.SubElement(xf, H.qn("a:ext"))
    off.set("x", str(x)); off.set("y", str(y)); ext.set("cx", str(w)); ext.set("cy", str(h))


def copy_graphic(s, sl, area):
    """把來源頁的圖形（標題、頁碼、頁尾、裝飾圖以外）照相對位置複製到 area，等比例縮放；圖片一起帶過去。"""
    import copy as _copy
    from pptx.enum.shapes import PP_PLACEHOLDER as P
    E = H.EMU
    picks = []
    titles = set(H.src_title(list(H.src.slides).index(sl) + 1))
    for sh in sl.shapes:
        if sh.has_text_frame and H.norm(sh.text_frame.text) in {H.norm(t) for t in titles}:
            continue  # 標題已經放在新投影片的標題
        if sh.is_placeholder and sh.placeholder_format.type in (P.TITLE, P.CENTER_TITLE) + H._SKIP_PH:
            continue
        if _is_pic(sh) and _hash(sh) in DECOR:
            continue
        if sh.has_text_frame and all(H._is_noise(x) or x.strip() == "" for x in sh.text_frame.text.split("\n")) \
                and not sh.text_frame.text.strip() == "" and sh.shape_type != 1:
            continue
        if sh.left is None or sh.width is None:
            continue
        if sh.has_text_frame and not sh.text_frame.text.strip() and (sh.is_placeholder or sh.shape_type == 17):
            continue  # 空的預留位置、空文字框（會把範圍撐大）
        picks.append(sh)
    if not picks:
        return
    l0 = min(sh.left for sh in picks); t0 = min(sh.top for sh in picks)
    r0 = max(sh.left + sh.width for sh in picks); b0 = max(sh.top + sh.height for sh in picks)
    x, y, w, h = area
    k = min(w * E / max(r0 - l0, 1), h * E / max(b0 - t0, 1), 1.5)
    ox = x * E + (w * E - (r0 - l0) * k) / 2
    for sh in picks:
        el = _copy.deepcopy(sh.element)
        # 預留位置改成一般圖形（不再繼承版面）
        for ph in el.iter(H.qn("p:ph")):
            ph.getparent().remove(ph)
            pr = el.find(H.qn("p:spPr"))
            if pr is not None and pr.find(H.qn("a:prstGeom")) is None and pr.find(H.qn("a:custGeom")) is None:
                from lxml import etree
                g = etree.Element(H.qn("a:prstGeom"), prst="rect")
                etree.SubElement(g, H.qn("a:avLst"))
                first = pr[0] if len(pr) else None  # 順序：xfrm 之後、填色之前
                pr.insert(1 if first is not None and first.tag == H.qn("a:xfrm") else 0, g)
            break
        s.shapes._spTree.append(el)
        _set_xfrm(el, int(ox + (sh.left - l0) * k), int(y * E + (sh.top - t0) * k), int(sh.width * k), int(sh.height * k))
        # 圖片的關聯要重新建立
        for blip in el.iter(H.qn("a:blip")):
            rid = blip.get(H.qn("r:embed"))
            if rid and rid in sl.part.rels:
                blob = sl.part.related_part(rid).blob
                _, new_rid = s.part.get_or_add_image_part(io.BytesIO(blob))
                blip.set(H.qn("r:embed"), new_rid)
        # 縮小時字也跟著縮；沒設字級的（原本繼承版面）給 14pt 起算
        if k < 0.98:
            for r in el.iter(H.qn("a:rPr")):
                sz = r.get("sz")
                r.set("sz", str(max(700, int((int(sz) if sz else 1400) * k))))


PAIR = re.compile(r"([^:：\s][^:：]*?)\s*[:：]\s*([-+]?\d[\d,]*(?:\.\d+)?)\s*([%％]|[^\s\d:：]{1,2})?(?=\s|$)")


def _pairs(line):
    """「標籤：數字單位」的組合。整行都是這種組合才算（中間只能有空白）。"""
    out, rest = [], line
    for m in PAIR.finditer(line):
        out.append((m.group(1).strip(), m.group(2), (m.group(3) or "").strip()))
    leftover = PAIR.sub("", line).strip()
    return out if out and not leftover else []


def analyze(src, n, total, agenda_items):
    """分析來源第 n 頁，回傳計畫的一筆（只記結構，不記內文）。"""
    sl = src.slides[n - 1]
    items = list(_all_shapes(sl.shapes))
    tables = [(sh, l, t, w, h) for sh, l, t, w, h in items if getattr(sh, "has_table", False) and sh.has_table]
    charts = [(sh, l, t, w, h) for sh, l, t, w, h in items if getattr(sh, "has_chart", False) and sh.has_chart]
    titles = H.src_title(n)
    cards = H.src_cards(n, pos=True)
    chars = sum(len(x) for c in cards for x in [c["head"]] + c["items"]) + sum(len(t) for t in titles)
    e = dict(src=n)
    low = " ".join(titles).lower()
    if n == 1 and not tables and not charts:
        e["kind"] = "cover"
    elif n == total and n > 1 and not tables and not charts and chars <= 30:
        e["kind"] = "closing"
    elif titles and any(w in titles[0].lower() for w in AGENDA_WORDS) and not tables:
        e["kind"] = "agenda"
    elif charts:
        e["kind"] = "chart"
    elif tables:
        e["kind"] = "table"
    elif is_graphic(sl):
        e["kind"] = "graphic"
    elif titles and not cards and not pictures(n):
        e["kind"] = "section"
    else:
        e["kind"] = "text"
    if e["kind"] in ("table", "chart", "text", "graphic", "section"):
        t0, sub = _titles(titles, agenda_items)
        e["title"] = t0
        if sub:
            e["subtitle"] = sub
    if e["kind"] == "text":
        heads = [c for c in cards if c["items"]]
        e["cols"] = len(heads) if 2 <= len(heads) <= 3 else 2
    return e


def _titles(titles, agenda_items):
    """決定標題和副標：章節名（和目錄同名）有其他標題時讓位；太長的標題在冒號處拆成標題＋副標。"""
    ts = [t for t in titles if t]
    if len(ts) > 1 and ts[0] in agenda_items:
        ts = ts[1:] + [ts[0]]
        ts = ts[:-1]  # 章節名在目錄頁已經有了
    if not ts:
        return "", ""
    t0, rest = ts[0], ts[1:]
    if len(t0) > 22 and re.search(r"[:：]", t0):
        a, b = re.split(r"[:：]", t0, maxsplit=1)
        if a.strip() and b.strip():
            t0, rest = a.strip(), [b.strip()] + rest
    rest = [r for r in rest if r not in t0]
    return t0, "　".join(rest)


def make_plan(src):
    total = len(src.slides)
    agenda_items = []
    for n in range(1, total + 1):  # 先找目錄頁，章節名稱給其他頁判斷標題用
        ts = H.src_title(n)
        if ts and any(w in ts[0].lower() for w in AGENDA_WORDS):
            agenda_items = _agenda_items(n)
            break
    return dict(version=1, pages=[analyze(src, n, total, agenda_items) for n in range(1, total + 1)])


def _agenda_items(n):
    out = []
    for c in H.src_cards(n):
        for x in ([c[0]] if c[0] else []) + c[1]:
            x = re.sub(r"^\s*(\d+|[IVX]+)[.、)]\s*", "", x).strip()
            if x:
                out.append(x)
    return out


# ---------------------------------------------------------------- 產生
def table_rows(tb):
    return [[H.cell_lines(c) for c in r.cells] for r in tb.rows]


def header_rows(tb):
    """表頭有合併儲存格（上下或左右合併）時，表頭算兩列。"""
    try:
        r0 = list(tb.rows[0].cells)
        if len(tb.rows) > 2 and any(c.is_merge_origin and (c.span_height > 1 or c.span_width > 1) for c in r0):
            return 2
    except Exception:
        pass
    return 1


def est_table_h(rows, size, w, compact):
    widths = H.col_widths(rows, size, w)
    pad, ln = (0.02, 1.2) if compact else (0.04, H.LINE)
    min_h = size * H.SCALE * ln / 72 + (0.02 if compact else 0.05)
    return sum(H.row_height(r, widths, size, min_h, pad, ln) for r in rows), widths


def place_table(s, src_tb, x, y, w, size, compact, bold_last=False):
    rows = table_rows(src_tb)
    ncol = max(len(r) for r in rows)
    rows = [(list(r) + [""] * ncol)[:ncol] for r in rows]
    total_h, widths = est_table_h(rows, size, w, compact)
    pad, ln = (0.02, 1.2) if compact else (0.04, H.LINE)
    min_h = size * H.SCALE * ln / 72 + (0.02 if compact else 0.05)
    hs = [H.row_height(r, widths, size, min_h, pad, ln) for r in rows]
    nh = header_rows(src_tb)
    body = rows[nh:]
    aligns = {c: H.C for c in range(ncol) if c > 0 and body and all(H.is_num(r[c]) for r in body if r[c])}
    t = H.table(s, x, y, widths, len(rows), hs)
    for r in range(nh):
        H.style_header_row(t, r, rows[r], size)
    last = bold_last or (body and re.search(r"^(總計|合計|小計|total|sum)", " ".join(body[-1][0]).strip(), re.I))
    H.zebra_body(t, nh, body, size, aligns, bold_last=bool(last))
    if compact:
        for r_ in t.rows:
            for c_ in r_.cells:
                c_.margin_top = c_.margin_bottom = Inches(0.01)
    _copy_merges(src_tb, t, nh)
    _copy_colors(src_tb, t)
    return y + sum(hs), size


def _copy_merges(src_tb, t, nh):
    """表頭的合併儲存格照來源合併。"""
    try:
        for r in range(nh):
            for c, cell in enumerate(src_tb.rows[r].cells):
                if cell.is_merge_origin and (cell.span_height > 1 or cell.span_width > 1):
                    r2 = min(r + cell.span_height - 1, len(t.rows) - 1)
                    c2 = min(c + cell.span_width - 1, len(t.columns) - 1)
                    t.cell(r, c).merge(t.cell(r2, c2))
    except Exception:
        pass


def _copy_colors(src_tb, t):
    """來源用顏色表示狀態（例如紅色 ●）時保留：有明確設定、不是黑白的字色照抄。"""
    for r, row in enumerate(src_tb.rows):
        if r >= len(t.rows):
            break
        for c, cell in enumerate(row.cells):
            if c >= len(t.columns):
                break
            cols = [H.run_color(run) for p in cell.text_frame.paragraphs for run in p.runs if run.text.strip()]
            cols = [x for x in cols if x and re.fullmatch(r"[0-9A-F]{6}", x) and x not in ("000000", "FFFFFF")]
            if not cols:
                continue
            for p in t.cell(r, c).text_frame.paragraphs:
                for run in p.runs:
                    run.font.color.rgb = H.rgb(cols[0])


def text_paras(lines, size, head=None):
    ps = []
    if head:
        ps.append(dict(text=head, size=size + 1, bold=True, color=H.NAVY, after=2))
    ps += [dict(text=x, size=size, color=H.BLACK, after=2) for x in lines]
    return ps


def build_flow(s, e, blocks, area):
    """依來源由上而下的順序排：數字卡片、文字段落、表格、成對數字小表。表格自動挑字級放進剩下的高度。"""
    x, y, w, h = area
    gap = 0.18
    fixed, tables = 0.0, []
    for b in blocks:
        if b["type"] == "stats":
            b["h"] = 1.05 * H.SCALE
        elif b["type"] == "text":
            b["size"] = 12
            b["h"] = H.est_h([(p["text"], p["size"], p["after"]) for p in text_paras(b["lines"], 12, b.get("head"))], w - 0.1) + 0.1
        elif b["type"] == "pairs":
            b["h"] = (0.3 * (len(b["rows"]) + 1) + (0.35 if b.get("caption") else 0)) * H.SCALE
        if b["type"] != "table":
            fixed += b["h"] + gap
        else:
            tables.append(b)
    budget = h - fixed - gap * len(tables)
    # 表格字級：所有表格用同一個字級，從 12pt 往下試，放不下再用緊湊列高，最小 8pt
    choice = None
    for size, compact in [(12, False), (11, False), (10, False), (10, True), (9, True), (8.5, True), (8, True),
                          (7.5, True), (7, True)]:
        need = sum(est_table_h(table_rows(b["tb"]), size, w, compact)[0] for b in tables)
        choice = (size, compact)
        if need <= budget:
            break
    for b in blocks:
        if b["type"] == "stats":
            H.stat_cards(s, b["items"], x=x, y=y, w=w, h=b["h"], value_size=28)
            y += b["h"] + gap
        elif b["type"] == "text":
            H.txt(s, x, y, w, b["h"], text_paras(b["lines"], b["size"], b.get("head")))
            y += b["h"] + gap
        elif b["type"] == "pairs":
            if b.get("caption"):
                H.txt(s, x, y, w, 0.34 * H.SCALE, [dict(text=b["caption"], size=13, bold=True, color=H.NAVY)])
                y += 0.37 * H.SCALE
            rows = [b["header"]] + b["rows"]
            _, y2 = H.put_table(s, rows, x, y, min(w, 1.6 * len(b["header"]) + 0.5), size=11, name="數字表")
            y = y2 + gap
        elif b["type"] == "table":
            y, _ = place_table(s, b["tb"], x, y, w, choice[0], choice[1])
            y += gap
    if tables and choice and choice[0] < 10:
        e.setdefault("_notes", []).append(f"輸出第 {len(H.prs.slides) - len(H.orig_ids)} 頁（來源第 {e.get('src')} 頁）的表格用 {choice[0]}pt "
                                          "才放得進一頁，請使用者確認字會不會太小")


def side_panel(s, cards, x, y, w, h):
    H.box(s, x, y, w, h, H.CLOUD)
    paras = []
    for c in cards:
        if c["head"]:
            paras.append(dict(text=c["head"], size=14, bold=True, color=H.BLUE, after=4))
        paras += [dict(text=t, size=11, color=H.BLACK, after=4) for t in c["items"]]
    size_scale = 1.0
    while size_scale > 0.7 and H.est_h([(p["text"], p["size"] * size_scale, p["after"]) for p in paras], w - 0.4) > h - 0.3:
        size_scale -= 0.05
    for p in paras:
        p["size"] = round(p["size"] * size_scale, 1)
    H.txt(s, x + 0.15, y + 0.12, w - 0.3, h - 0.2, paras)


def blocks_for(n, e, side_ok=True):
    """把來源頁的內容分成排版區塊：表格、數字卡片、成對數字表、文字；表格旁邊的文字當側欄。"""
    sl = H.src.slides[n - 1]
    items = list(_all_shapes(sl.shapes))
    tables = sorted([(sh, l, t, w, h) for sh, l, t, w, h in items if getattr(sh, "has_table", False) and sh.has_table],
                    key=lambda x: x[2])
    cards = H.src_cards(n, pos=True)
    side = []
    if tables and side_ok and e.get("side", "auto") != "none":
        big = max(tables, key=lambda x: x[3] * x[4])
        _, bl, bt, bw, bh = big
        for c in cards:
            beside = c["l"] >= bl + bw * 0.75 and c["t"] < bt + bh and c["b"] > bt + 0.3
            if beside and (len(c["items"]) >= 1):
                side.append(c)
    rest = [c for c in cards if c not in side]
    blocks = []
    # 成對數字：「A：23 B：17 C：73%」很多行、標籤一樣 → 小表格；只有一兩個 → 數字卡片
    stat_items, pair_rows, pair_hdr, texts = [], [], None, []
    pair_caption, row_t = None, None
    for c in rest:
        # 同一列（上緣差 0.5 吋內）的卡片用同一個排序值，保留 src_cards 的左→右順序
        row_t = c["t"] if row_t is None or abs(c["t"] - row_t) >= 0.5 else row_t
        lines = ([c["head"]] if c["head"] else []) + c["items"]
        # 整張卡片都是「標籤：數字」（每行 1–2 組）才做成數字卡片
        if lines and all(1 <= len(_pairs(ln)) <= 2 for ln in lines) and len(lines) <= 3:
            for ln in lines:
                stat_items += _pairs(ln)
            continue
        kept = []
        for li, ln in enumerate(lines):
            pr = _pairs(ln)
            if li == 0 and 1 <= len(pr) <= 2:  # 卡片第一行就是「標籤：數字」→ 數字卡片
                stat_items += pr
                continue
            if len(pr) >= 3:
                pre = re.match(r"^([A-Za-z]?\d{1,2}|[一二三四]季|Q\d)", pr[0][0])
                key = pre.group(1) if pre else ""
                labels = [pr[0][0][len(key):].strip()] + [p[0] for p in pr[1:]]
                hdr = [""] + labels if key else labels
                if pair_hdr is None or pair_hdr == hdr:
                    if pair_hdr is None and kept and not _pairs(kept[-1]) and len(kept[-1]) <= 20:
                        pair_caption = kept.pop()  # 數字表上方的小標（例如「每季○○」）
                    pair_hdr = hdr
                    pair_rows.append(([key] if key else []) + [p[1] + p[2] for p in pr])
                    continue
            kept.append(ln)
        if kept:
            head = c["head"] if c["head"] and c["head"] in kept[:1] else None
            body = kept[1:] if head else kept
            texts.append(dict(type="text", head=head, lines=body, t=row_t))
    if len(stat_items) > 6:  # 太多就不做卡片，當文字
        texts.append(dict(type="text", head=None, lines=[f"{a}：{b}{u}" for a, b, u in stat_items], t=0))
        stat_items = []
    if stat_items:
        blocks.append(dict(type="stats", items=stat_items, t=-1))
    for tb in tables:
        blocks.append(dict(type="table", tb=tb[0].table, t=tb[2]))
    if pair_rows and len(pair_rows) >= 2:
        blocks.append(dict(type="pairs", header=pair_hdr, rows=pair_rows, caption=pair_caption,
                           t=min([c["t"] for c in rest] or [0])))
    elif pair_rows:  # 只有一行就當文字
        texts.append(dict(type="text", head=pair_caption, lines=[" ".join(r) for r in pair_rows], t=0))
    blocks += texts
    blocks.sort(key=lambda b: b["t"])  # sort 是穩定的：同一列維持左→右
    return blocks, side, cards


def render_page(e, notes):
    kind = e.get("kind", "text")
    n = e.get("src")
    if kind == "skip":
        notes.append(f"來源第 {n} 頁：計畫設成 skip，沒有產生（{e.get('reason', '沒有寫原因')}）")
        return
    if kind == "custom":
        s = H.new_slide(e.get("layout", "content"), e.get("title", ""))
        path = e.get("script")
        code = open(path, encoding="utf-8").read()
        env = dict(vars(H)); env.update(s=s, src_no=n, H=H)
        exec(compile(code, path, "exec"), env)
        return
    if kind == "cover":
        boxes = []
        for sh, l, t, w, h in H._abs_boxes(H.src.slides[n - 1].shapes):
            for i, ln in enumerate(H._lines(sh)):
                boxes.append((_max_size(sh) if i == 0 else _max_size(sh) * 0.5, t + i * 0.01, _clean(ln)))
        boxes = [b for b in boxes if b[2]]
        boxes.sort(key=lambda b: (-b[0], b[1]))
        title = e.get("title") or (boxes[0][2] if boxes else "")
        rest = [b[2] for b in sorted(boxes[1:], key=lambda b: b[1])] if not e.get("title") else [b[2] for b in boxes if b[2] != title]
        date_re = re.compile(r"\d{4}|\d{1,2}/\d{1,2}|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|月", re.I)
        dates = [r for r in rest if date_re.search(r)]
        others = [r for r in rest if r not in dates]
        s = H.new_slide("cover", title)
        t_ = s.shapes.title
        if t_ is not None and t_.width and t_.height:
            fit_text(t_, title, inherited_size(s, t_) or 44)
        from pptx.enum.shapes import PP_PLACEHOLDER as P
        nsub = len([ph for ph in s.placeholders if ph.placeholder_format.type in (P.SUBTITLE, P.BODY)])
        sub, date = e.get("subtitle", "　".join(others)), e.get("date", "　".join(dates))
        phs = []
        if nsub >= 2:
            if sub:
                phs.append(H.set_subtitle(s, sub))
            if date:
                phs.append(H.set_subtitle(s, date, nth=1))
        elif sub or date:  # 範本封面只有一個副標位置：分兩行放
            phs.append(H.set_subtitle(s, "\n".join(x for x in (sub, date) if x)))
        for ph in phs:  # 副標字太多：縮小字到放得下（最小 12pt）
            if ph is None or not ph.width or not ph.height:
                continue
            text = ph.text_frame.text
            size = 24
            while size > 12 and H.est_h([(t, size, 0) for t in text.split("\n")], ph.width / H.EMU - 0.2) > ph.height / H.EMU:
                size -= 2
            if size < 24:
                for p_ in ph.text_frame.paragraphs:
                    for r_ in p_.runs:
                        r_.font.size = Pt(size * H.SCALE)
        return
    if kind == "closing":
        lines = H.src_title(n) + [x for c in H.src_cards(n) for x in ([c[0]] if c[0] else []) + c[1]]
        s = H.new_slide("closing")
        if s is None:  # 範本沒有結尾版面：用來源的結尾字當標題
            s = H.new_slide("section", lines[0] if lines else "") or H.new_slide("content", lines[0] if lines else "")
            notes.append("範本沒有結尾版面，用" + ("章節頁" if s.slide_layout.name else "內容頁") + "放來源的結尾文字")
        elif s.shapes.title is not None and not s.shapes.title.text_frame.text.strip():
            # 結尾版面的標題是空的：用範本範例頁（同一個版面）上的結尾字，例如「Thank You!」
            sample = next((sl.shapes.title.text_frame.text for sl in H.prs.slides
                           if sl is not s and sl.slide_layout == s.slide_layout and sl.shapes.title is not None
                           and sl.shapes.title.text_frame.text.strip()), "")
            sample = re.sub(r"[\ue000-\uf8ff]", "", sample).strip()  # 去掉圖示字型的私用字元
            s.shapes.title.text = sample or (lines[0] if lines else "Thank You")
            if lines and sample:
                notes.append(f"結尾頁：用範本的「{sample}」，來源的「{'、'.join(lines)[:30]}」沒有放")
        elif lines:
            notes.append(f"結尾頁：來源的「{'、'.join(lines)[:30]}」沒有放（範本的結尾頁已經有結尾文字）")
        return
    if kind == "agenda":
        title = e.get("title") or (H.src_title(n) or ["Agenda"])[0]
        items = e.get("items") or _agenda_items(n)
        if e.get("template_layout") != "none" and try_agenda_layout(title, items):
            return
        s = H.new_slide("content", title)
        add_slide_number(s)
        H.agenda(s, items)
        drop_empty(s)
        return
    if kind == "section":
        s = H.new_slide("section", e.get("title", ""))
        if s is None:
            s = H.new_slide("content", e.get("title", ""))
        if e.get("subtitle"):
            H.set_subtitle(s, e["subtitle"])
        return
    if kind == "text" and try_pair_layout(e, n):
        return
    if kind == "chart" and try_chart_layout(e, n):
        return
    s = H.new_slide(e.get("layout", "content"), e.get("title", ""))
    set_title(s, e.get("title", ""))
    add_slide_number(s)
    if e.get("subtitle"):
        H.set_subtitle(s, e["subtitle"])
    drop_empty(s)
    x, y, w, h = H.content_area(s)
    if kind == "chart":
        blocks, side, cards = blocks_for(n, e, side_ok=False)
        txts = [b for b in blocks if b["type"] == "text"]
        for b in txts:
            hh = H.est_h([(p["text"], p["size"], p["after"]) for p in text_paras(b["lines"], 12, b.get("head"))], w) + 0.1
            H.txt(s, x, y, w, hh, text_paras(b["lines"], 12, b.get("head")))
            y += hh + 0.1
        H.column_chart(s, area=(x, y, w, h - (y - H.content_area(s)[1])), **H.src_chart(n))
        return
    if kind == "graphic":
        copy_graphic(s, H.src.slides[n - 1], (x, y, w, h))
        return
    pics = [] if e.get("pictures") == "none" else pictures(n)
    if kind == "table":
        blocks, side, _ = blocks_for(n, e)
        if pics:  # 表格頁有圖：圖放右邊一欄（有側欄時放在側欄下方）
            pw = float(e.get("picture_width", w * 0.35))
            px = x + w - pw
            py = y
            if side:
                sh_ = min(h * 0.45, 0.6 + 0.4 * sum(1 + len(c["items"]) for c in side))
                side_panel(s, side, px, y, pw, sh_)
                py = y + sh_ + 0.2
            place_pictures(s, pics, px, py, pw, y + h - py)
            build_flow(s, e, blocks, (x, y, w - pw - 0.25, h))
            notes += e.pop("_notes", [])
            return
        if side:
            sw = float(e.get("side_width", 3.4))
            build_flow(s, e, blocks, (x, y, w - sw - 0.25, h))
            side_panel(s, side, x + w - sw, y, sw, min(h, 0.6 + 0.4 * sum(1 + len(c["items"]) for c in side) + 0.6))
        else:
            build_flow(s, e, blocks, (x, y, w, h))
        notes += e.pop("_notes", [])
        return
    # text
    blocks, _, cards = blocks_for(n, e, side_ok=False)
    other = [b for b in blocks if b["type"] != "text"]
    txt_blocks = [b for b in blocks if b["type"] == "text"]
    if pics:
        if not blocks:  # 只有圖：整個內容區放圖
            place_pictures(s, pics, x, y, w, h)
            return
        # 文字在左、圖在右；文字少時文字欄窄一點
        chars = sum(len(t) for b in txt_blocks for t in b["lines"] + [b.get("head") or ""])
        tw = float(e.get("text_width", w * (0.32 if chars < 120 else 0.42)))
        place_pictures(s, pics, x + tw + 0.3, y, w - tw - 0.3, h)
        w = tw
        e = dict(e, cols=1)
    if other:  # 先放數字卡片、數字表
        x0, y0, w0, h0 = x, y, w, h
        build_flow(s, e, other, (x0, y0, w0, h0))
        used = sum((b.get("h") or 0) + 0.18 for b in other)
        y += used
        h -= used
    groups = [(b.get("head") or "", b["lines"]) for b in txt_blocks]
    with_items = [g for g in groups if g[1]]
    heads_only = [g[0] for g in groups if not g[1] and g[0]]
    if len(with_items) == 1 and len(with_items[0][1]) >= 6:
        hd, its = with_items[0]
        need = H.est_h([(t, 12, 6) for t in its], w - 0.6) + 0.9
        if need > h and not e.get("cols"):  # 一張卡片放不下：分成左右兩欄
            half = (len(its) + 1) // 2
            with_items = [(hd, its[:half]), ("", its[half:])]
            e = dict(e, cols=2)
    if with_items:
        cols = int(e.get("cols", 2))
        if len(with_items) == 1:
            cols = 1
        H.card_grid(s, with_items, cols=cols, area=(x, y, w, h - (0.4 if heads_only else 0)), min_size=9)
        y_end = max((sh.top + sh.height) / H.EMU for sh in s.shapes if sh.top is not None)
    else:
        y_end = y
    if heads_only:  # 只有標題沒有內容的區塊：只放標題，不補「無」
        ps = [dict(text=t, size=12, bold=True, color=H.NAVY, after=2) for t in heads_only]
        hh = H.est_h([(p["text"], p["size"], p["after"]) for p in ps], w - 0.1) + 0.1
        H.txt(s, x, min(y_end + 0.15, H.slide_size()[1] - 0.3 - hh), w, hh, ps)


# ---------------------------------------------------------------- 主程式
def report(src, out_path, notes, png, work):
    H.prs = None
    from pptx import Presentation
    out = Presentation(out_path)
    issues = []
    owner = getattr(H, "OWNER_OVERRIDE", None) or []
    kind_of = {e["src"] - 1: e.get("kind") for e in PLAN["pages"]}
    out_kind = {i + 1: kind_of.get(o) for i, o in enumerate(owner)}

    def keep(m):
        mm = re.match(r"第 (\d+) 頁：", m)
        if not mm:
            return True
        k = out_kind.get(int(mm.group(1)))
        if k == "graphic" and ("重疊" in m or "超出" in m):
            return False  # 圖解頁照來源的相對位置搬，本來就疊在一起（例如圖上的標註）
        if k in ("cover", "graphic") and "一整串清單" in m:
            return False
        return True
    issues += [m for m in H.overlap_check(out) if keep(m)]
    issues += H.chart_check(src, out)
    issues += H.moved_text(src, out)
    issues += [m for m in H.structure_check(src, out) if keep(m)]
    miss = H.missing_text(src, out)
    closing_src = [p["src"] for p in PLAN["pages"] if p.get("kind") == "closing"]
    for no, toks in miss.items():
        if no in closing_src:
            continue
        issues.append(f"MISSING 來源第 {no} 頁：" + "、".join(toks[:15]) + (f" …等 {len(toks)} 處" if len(toks) > 15 else ""))
    # 圖片：來源每頁要搬的圖，輸出對應的頁面要有一樣多
    for e in PLAN["pages"]:
        if e.get("kind") in ("cover", "closing", "skip", "agenda", "section", "custom"):
            continue
        k = len(pictures(e["src"]))
        if not k:
            continue
        got = sum(1 for i, o in enumerate(owner) if o == e["src"] - 1 and i < len(out.slides)
                  for sh in out.slides[i].shapes if _is_pic(sh))
        if got < k:
            if e.get("pictures") == "none":
                notes.append(f"來源第 {e['src']} 頁的 {k} 張圖依計畫沒有放（pictures=none）")
            else:
                issues.append(f"來源第 {e['src']} 頁有 {k} 張圖，輸出只有 {got} 張：圖片不可以省略")
    issues += [w for w in H.warn if not w.startswith("overflow?")]
    issues += [w for w in H.warn if w.startswith("overflow?")]
    png_issues = []
    if png:
        try:
            sys.path.insert(0, os.path.join(HERE, "..", "..", "powerpoint-pptx", "scripts"))
            import office_convert
            rep = []
            folder = os.path.join(work, "out_png")
            files = office_convert.export_slides(out_path, folder, report=rep)
            print(f"已匯出 {len(files)} 張圖到 {folder}（{rep[-1] if rep else ''}）")
            try:
                import pptx_tool
                sheets = pptx_tool.overview_sheets(files)
                for s_ in sheets:
                    print("  總覽圖：" + s_)
            except Exception as ex:
                print(f"（總覽圖沒做成：{ex}）")
            png_issues = list(getattr(office_convert, "layout_issues", []))
        except Exception as ex:
            print(f"（沒有匯出圖片：{ex}。改用 check_bounds.py 估算）")
    # 範本標題框本身的小溢出（範本範例頁也一樣）只當說明，不算問題
    titles_out = {i + 1: (sl.shapes.title.text_frame.text.strip()[:16] if sl.shapes.title is not None else None)
                  for i, sl in enumerate(out.slides)}
    tpl_small = []
    for m in png_issues:
        mm = re.match(r"第 (\d+) 頁：「(.+?)」 的文字超出自己的框 ([\d.]+) 吋", m)
        if mm and float(mm.group(3)) <= 0.3 and titles_out.get(int(mm.group(1))) and                 titles_out[int(mm.group(1))].startswith(mm.group(2)[:8]):
            tpl_small.append(int(mm.group(1)))
        else:
            issues.append(f"PowerPoint 實際排版：{m}")
    if tpl_small:
        notes.append(f"標題比範本的標題框高一點點（{len(tpl_small)} 頁，0.3 吋以內；範本自己的範例頁也一樣），看起來正常")
    print()
    for m in notes:
        print("說明 " + m)
    for m in issues:
        print("問題 " + m)
    print()
    if issues:
        print(f"結果：還有 {len(issues)} 個問題。改 plan.json（版型、欄數、側欄、標題）後重跑；"
              "程式判斷錯的頁面可以用 kind=custom 自己寫那一頁。不要帶著問題交付。")
    else:
        print("結果：自動檢查全部通過。看總覽圖確認版面後就可以交付；回報時附上上面的「說明」。")
    return issues


PLAN = {}


def main():
    global PLAN
    ap = argparse.ArgumentParser(description="來源簡報自動套範本")
    ap.add_argument("src"); ap.add_argument("tpl"); ap.add_argument("out")
    ap.add_argument("--plan", help="照這個計畫產生（預設：重新分析）")
    ap.add_argument("--work", help="工作資料夾（預設：輸出檔旁的 _work）")
    ap.add_argument("--no-png", action="store_true", help="不匯出圖片")
    a = ap.parse_args()
    work = a.work or os.path.join(os.path.dirname(os.path.abspath(a.out)), "_work")
    os.makedirs(work, exist_ok=True)
    for p in (a.src, a.tpl):
        if not os.path.exists(p):
            sys.exit(f"找不到檔案：{p}")
    if os.path.abspath(a.out) in (os.path.abspath(a.src), os.path.abspath(a.tpl)):
        sys.exit("輸出檔不可以和來源或範本同名（不要覆蓋原始檔）")
    src, prs = H.init(a.src, a.tpl)
    find_decor(src)
    if a.plan:
        PLAN = json.load(open(a.plan, encoding="utf-8"))
        print(f"使用計畫：{a.plan}")
    else:
        PLAN = make_plan(src)
        pp = os.path.join(work, "plan.json")
        with open(pp, "w", encoding="utf-8") as f:
            json.dump(PLAN, f, ensure_ascii=False, indent=1)
        print(f"已寫出頁面計畫：{pp}")
    print("頁面計畫：")
    for e in PLAN["pages"]:
        extra = "".join(f" {k}={e[k]}" for k in ("title", "subtitle", "cols", "side") if e.get(k))
        print(f"  來源第 {e['src']:>2} 頁 → {e['kind']}{extra}")
    notes = []
    owner = []  # 每個輸出頁對應的來源頁（0 起算），給檢查用
    for e in PLAN["pages"]:
        before = len(H.prs.slides)
        if e.get("kind") not in KINDS:
            sys.exit(f"plan.json 第 {e.get('src')} 頁的 kind「{e.get('kind')}」不認得，可用：{', '.join(KINDS)}")
        try:
            render_page(e, notes)
        except Exception as ex:
            import traceback
            traceback.print_exc()
            sys.exit(f"產生來源第 {e.get('src')} 頁時出錯：{ex}。在 plan.json 把這頁改成別的 kind（或 custom）再試")
        owner += [e.get("src", 1) - 1] * (len(H.prs.slides) - before)
    H.OWNER_OVERRIDE = owner
    for sid in H.orig_ids:
        H.prs.part.drop_rel(sid.rId)
        H.prs.slides._sldIdLst.remove(sid)
    H.prs.save(a.out)
    print(f"已輸出：{os.path.abspath(a.out)}（{len(H.prs.slides)} 頁）")
    report(src, a.out, notes, not a.no_png, work)


if __name__ == "__main__":
    main()
