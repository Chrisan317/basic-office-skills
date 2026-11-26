"""Reusable helpers for rebuilding a deck on a company template with python-pptx.
usage:  import helpers as H; from helpers import *; src, prs = init(SRC, TPL) ... save(OUT)

顏色：init() 會從範本的佈景主題（theme）自動挑顏色，更新 BLUE/NAVY/CYAN/CLOUD/ICE/ZEBRA/GRAY，並印出挑選結果。
      主色 BLUE 不一定是 accent1：有些範本的 accent1 是很淺的底色，所以會挑「在白底上看得清楚、有彩度」的 accent。
      init() 也會更新呼叫端 `from helpers import *` 拿到的同名變數，所以直接寫 BLUE、CLOUD 也可以；
      用 H.BLUE 也行（別把模組取名成 h，h 常被拿來當「高度」變數而蓋掉模組）。
版面：new_slide() 可以給範本的版面名稱，也可以給種類 "cover" / "content" / "section" / "closing"，
      會自動在範本裡找最適合的版面（找不到 closing/section 時回傳 None）。
區域：content_area(slide) 回傳 (x, y, w, h)，依投影片大小、標題與已填的副標位置計算，16:9、4:3 都適用。"""
import re, math, copy, sys, os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.oxml.ns import qn
from lxml import etree

try:  # Windows 中文環境：輸出用 UTF-8，檔名、警告才不會變亂碼
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

# 預設調色盤（範本讀不到主題色時使用）
BLUE, CYAN, ICE, CLOUD, NAVY, GRAY, BLACK, WHITE = "0055CC", "00B0F0", "9DC3E6", "DAE3F3", "002145", "767171", "000000", "FFFFFF"
ZEBRA = "F2F6FC"
RED = "E03C31"
PALETTE_NAMES = ("BLUE", "CYAN", "ICE", "CLOUD", "NAVY", "GRAY", "ZEBRA")

def rgb(h): return RGBColor.from_string(h)


def _mix(hex_color, white_ratio, to="FFFFFF"):
    """把顏色和白色（或 to）混合，white_ratio=0.85 代表 85% 白（做淺色卡片底色）。"""
    c1 = [int(hex_color[i:i + 2], 16) for i in (0, 2, 4)]
    c2 = [int(to[i:i + 2], 16) for i in (0, 2, 4)]
    return "%02X%02X%02X" % tuple(round(a + (b - a) * white_ratio) for a, b in zip(c1, c2))


def _lum(hex_color):
    """WCAG 相對亮度 0（黑）～1（白）。"""
    def ch(v):
        v /= 255
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    r, g, b = (int(hex_color[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b)


def _contrast(a, b="FFFFFF"):
    la, lb = sorted((_lum(a), _lum(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


def _hsv(hex_color):
    import colorsys
    r, g, b = (int(hex_color[i:i + 2], 16) / 255 for i in (0, 2, 4))
    return colorsys.rgb_to_hsv(r, g, b)


def _hue_near(a, b, deg=35):
    d = abs(_hsv(a)[0] - _hsv(b)[0]) * 360
    return min(d, 360 - d) <= deg


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


def pick_palette(th):
    """從主題色挑調色盤。回傳 {名稱: (顏色, 來源)}。
    BLUE  主色：依 accent1→6 順序，第一個「白底上對比 ≥ 3、有彩度」的（很淺的 accent 會跳過）
    CLOUD 卡片底：和主色同色系、很淺的 accent；沒有就用主色混 85% 白
    ICE   淺色數列：和主色同色系、中等亮度的 accent；沒有就混 60% 白
    CYAN  輔色：主色以外、彩度高、在白底看得到的 accent
    GRAY  次要文字：低彩度、夠深的 accent；沒有就保留預設
    NAVY  深色標題：dk2（夠深的話），否則把主色加深
    ZEBRA 表格斑馬紋：主色混 95% 白"""
    acc = [(f"accent{i}", th[f"accent{i}"]) for i in range(1, 7) if th.get(f"accent{i}")]
    out = {}
    if not acc:
        return out
    sat = lambda c: _hsv(c)[1]
    # 依 accent1→6 的順序，取第一個「白底上對比 ≥ 3（大字/白字表頭看得清楚）且有彩度」的顏色
    strong = [(n, c) for n, c in acc if _contrast(c) >= 3 and sat(c) >= 0.25]
    if not strong:
        return out  # 主題色都太淺或沒有彩度：保留預設調色盤
    bn, blue = strong[0]
    out["BLUE"] = (blue, bn)
    light = [(n, c) for n, c in acc if _lum(c) >= 0.7 and 0.04 <= sat(c) and _hue_near(c, blue)]
    out["CLOUD"] = light[0][::-1] if light else (_mix(blue, 0.85), f"{bn} 混 85% 白")
    mid = [(n, c) for n, c in acc if 0.3 <= _lum(c) < 0.7 and sat(c) >= 0.15 and _hue_near(c, blue)
           and c != blue]
    out["ICE"] = mid[0][::-1] if mid else (_mix(blue, 0.6), f"{bn} 混 60% 白")
    cyan = [(n, c) for n, c in acc if c not in (blue, out["ICE"][0]) and sat(c) >= 0.5 and _contrast(c) >= 1.8]
    if cyan:
        out["CYAN"] = cyan[0][::-1]
    gray = [(n, c) for n, c in acc if sat(c) < 0.12 and _contrast(c) >= 3.5]
    if gray:
        out["GRAY"] = gray[0][::-1]
    dk2 = th.get("dk2")
    out["NAVY"] = (dk2, "dk2") if dk2 and _contrast(dk2) >= 7 else (_mix(blue, 0.55, "000000"), f"{bn} 加深")
    out["ZEBRA"] = (_mix(blue, 0.95), f"{bn} 混 95% 白")
    return out


def apply_theme_palette(pres, verbose=True):
    """用範本的主題色更新調色盤（見 pick_palette）。回傳主題色 dict。"""
    th = theme_colors(pres)
    pal = pick_palette(th)
    g = globals()
    for name, (col, why) in pal.items():
        g[name] = col
    if verbose:
        if pal:
            print("調色盤（取自範本主題）：" + "、".join(f"{n}={c}（{w}）" for n, (c, w) in pal.items()))
        else:
            print("範本主題色讀不到或都太淺，使用預設調色盤")
    return th

# ---------- state (set by init) ----------
src = None          # source deck (content is read from here)
prs = None          # output deck, opened from the template
L = {}              # template layout name -> layout
orig_ids = []       # template sample slides, removed by save()
warn = []           # overflow warnings collected by txt()
# powerpoint-pptx 技能的工具（讀簡報、匯出圖片）：和本技能資料夾在同一層
PPT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                    "..", "..", "powerpoint-pptx", "scripts", "pptx_tool.py"))

def init(src_path, tpl_path, theme=True):
    """Open source + template. Returns (src, prs). Call once before anything else.
    src_path may be None when the source is not a pptx (use extract_source.py for the content).
    theme=False：不讀範本主題色，保留目前的調色盤（例如你已經手動指定顏色）。"""
    global src, prs
    src = Presentation(src_path) if src_path else None
    prs = Presentation(tpl_path)
    # 就地更新（不要重新指定）：`from helpers import *` 拿到的 L、orig_ids 才不會是空的舊物件
    orig_ids[:] = list(prs.slides._sldIdLst)
    L.clear(); L.update({l.name: l for l in prs.slide_layouts})
    warn.clear()
    print(f"範本版面 {len(L)} 個：" + "、".join(L))
    if os.path.exists(PPT):
        print(f'匯出圖片：python "{PPT}" slides 輸出.pptx -o _work/out_png')
    else:
        print(f"找不到 powerpoint-pptx 技能的 pptx_tool.py（預期在 {PPT}）")
    old = {n: globals()[n] for n in PALETTE_NAMES}
    if theme:
        apply_theme_palette(prs)
    # 呼叫端用 `from helpers import *` 拿到的是舊值：同名且還是舊值的就一起更新
    try:
        caller = sys._getframe(1).f_globals
        if caller is not globals():
            for n in PALETTE_NAMES:
                if caller.get(n) == old[n]:
                    caller[n] = globals()[n]
    except Exception:
        pass
    return src, prs

def save(out_path, check=True):
    """Drop the template's sample slides, save, print warnings.
    有來源簡報時（check=True）會比對文字，印出「來源有、輸出沒有」的字（MISSING），一定要逐條處理。"""
    for sid in orig_ids:
        prs.part.drop_rel(sid.rId)
        prs.slides._sldIdLst.remove(sid)
    prs.save(out_path)
    print("saved", out_path, "slides:", len(prs.slides))
    msgs = list(warn)
    if check:
        msgs += overlap_check(prs)
        if src is not None:
            msgs += chart_check(src, prs)
            msgs += moved_text(src, prs)
            msgs += structure_check(src, prs)
    for w_ in msgs: print("WARN", w_)
    if check and src is not None:
        miss = missing_text(src, prs)
        if miss:
            n = sum(len(v) for v in miss.values())
            print(f"MISSING 來源有、輸出沒有的文字共 {n} 處（漏掉的表格、條列、數字）。逐條補上。\n"
                  "  例外：來源的錯字、裝飾字、範本已經有的字（例如結尾頁的 Thank you、公司名）不要硬加回去，"
                  "在回報裡說明即可。**不要為了消掉 MISSING 在投影片上疊文字。**")
            for no, toks in miss.items():
                more = f" …等 {len(toks)} 處" if len(toks) > 15 else ""
                print(f"  來源第 {no} 頁：" + "、".join(toks[:15]) + more)
        else:
            print("文字比對：來源的文字都有出現在輸出裡")
    if msgs:
        print(f"共 {len(msgs)} 則 WARN：重疊、超出、拆頁、圖表數列不符都要修好再存；看總覽圖看不出來時打開那一頁的單頁圖確認。")


# ---------- 存檔時的自動檢查 ----------
EMU = 914400


def _text_rect(sh_):
    """估計文字實際佔的範圍 (左, 上, 右, 下)（吋）：文字比框小時用文字大小，文字溢出時往下延伸。"""
    l, t = (sh_.left or 0) / EMU, (sh_.top or 0) / EMU
    w, h = (sh_.width or 0) / EMU, (sh_.height or 0) / EMU
    tf = sh_.text_frame
    ml = (tf.margin_left if tf.margin_left is not None else 91440) / EMU
    mr = (tf.margin_right if tf.margin_right is not None else 91440) / EMU
    inner = max(w - ml - mr, 0.2)
    maxw, est, align = 0.0, 0.0, None
    for p in tf.paragraphs:
        # 每個 run 字級可能不同（大數字＋小單位）：逐 run 算寬度，遇到手動換行另起一行
        segs, cur, big = [], 0.0, 0
        for r in p.runs:
            sz = r.font.size.pt if r.font.size else 18
            big = max(big, sz)
            parts = r.text.replace("\x0b", "\n").split("\n")
            for k, part in enumerate(parts):
                if k:
                    segs.append(cur); cur = 0.0
                cur += text_w(part, sz) / 72
        segs.append(cur)
        big = big or 18
        for sw in segs:
            maxw = max(maxw, min(sw, inner))
            est += max(1, math.ceil(sw * 1.08 / inner)) * big * 1.25 / 72
        est += p.space_after.pt / 72 if p.space_after else 0
        align = align or p.alignment
    est += 0.1
    tw = min(w, maxw + ml + mr)
    if align == PP_ALIGN.CENTER:
        l += (w - tw) / 2
    elif align == PP_ALIGN.RIGHT:
        l += w - tw
    if sh_.is_placeholder:  # 預留位置的垂直對齊是繼承的，用框的高度
        return (l, t, l + tw, t + h)
    anchor = tf.vertical_anchor
    if anchor == MSO_ANCHOR.MIDDLE and est < h:
        t += (h - est) / 2
    elif anchor == MSO_ANCHOR.BOTTOM and est < h:
        t += h - est
    return (l, t, l + tw, t + max(est, 0.2) if est > h or anchor in (None, MSO_ANCHOR.TOP) else t + est)


def _items(shapes, from_layout=False, slide_area=None):
    out = []
    for sh_ in walk_shapes(shapes):
        if from_layout and sh_.is_placeholder:
            continue
        if sh_.is_placeholder and sh_.placeholder_format.type in _SKIP_PH:
            continue
        box_ = ((sh_.left or 0) / EMU, (sh_.top or 0) / EMU,
                ((sh_.left or 0) + (sh_.width or 0)) / EMU, ((sh_.top or 0) + (sh_.height or 0)) / EMU)
        if slide_area and (box_[2] - box_[0]) * (box_[3] - box_[1]) > slide_area * 0.4:
            continue  # 背景圖
        if getattr(sh_, "has_table", False) and sh_.has_table:
            out.append(("表格", box_, from_layout))
        elif getattr(sh_, "has_chart", False) and sh_.has_chart:
            out.append(("圖表", box_, from_layout))
        elif sh_.shape_type == 13:  # picture
            out.append(("範本圖片" if from_layout else "圖片", box_, from_layout))
        elif sh_.has_text_frame and sh_.text_frame.text.strip():
            label = norm(sh_.text_frame.text)[:14]
            out.append((("範本文字「%s」" if from_layout else "「%s」") % label, _text_rect(sh_), from_layout))
    return out


def overlap_check(pres):
    """找出文字和表格／圖表／其他文字／範本圖案重疊、或超出投影片的地方，回傳警告 list。
    （卡片底色這類沒有文字的圖形不算。）"""
    W, H = pres.slide_width / EMU, pres.slide_height / EMU
    msgs = []
    for no, sl in enumerate(pres.slides, 1):
        its = _items(sl.shapes)
        its += _items(sl.slide_layout.shapes, True, W * H)
        for name, (l, t, r, b), tpl in its:
            if not tpl and (r > W + 0.05 or b > H + 0.05 or l < -0.05 or t < -0.05):
                msgs.append(f"第 {no} 頁：{name} 超出投影片（右 {r:.2f}/{W:.2f}、下 {b:.2f}/{H:.2f}）")
        for i in range(len(its)):
            for j in range(i + 1, len(its)):
                (n1, a, t1), (n2, b_, t2) = its[i], its[j]
                if t1 and t2:
                    continue
                ow = min(a[2], b_[2]) - max(a[0], b_[0])
                oh = min(a[3], b_[3]) - max(a[1], b_[1])
                if ow > 0.12 and oh > 0.1:
                    msgs.append(f"第 {no} 頁：{n1} 和 {n2} 重疊")
    return msgs


def _charts(pres):
    out = []
    for no, sl in enumerate(pres.slides, 1):
        for sh_ in walk_shapes(sl.shapes):
            if getattr(sh_, "has_chart", False) and sh_.has_chart:
                ser = {}
                for pl in sh_.chart.plots:
                    kind = "折線" if "Line" in type(pl).__name__ else "長條"
                    for se in pl.series:
                        vals = list(se.values)
                        if vals and any(v is not None for v in vals):
                            ser[re.sub(r"\s+", "", se.name or "")] = (kind, [round(v or 0, 6) for v in vals])
                out.append((no, ser))
    return out


def chart_check(src_pres, out_pres):
    """比對來源圖表和輸出圖表：數列有沒有少、折線有沒有變成長條、數值是否一樣。"""
    outs = _charts(out_pres)
    msgs = []
    for no, ser in _charts(src_pres):
        if not ser:
            continue
        best = max(outs, key=lambda o: len(set(o[1]) & set(ser)), default=None)
        if best is None or not set(best[1]) & set(ser):
            msgs.append(f"來源第 {no} 頁的圖表沒有重建（輸出找不到同名數列）：用 column_chart(s, **src_chart({no}))")
            continue
        for nm, (kind, vals) in ser.items():
            got = best[1].get(nm)
            if got is None:
                msgs.append(f"來源第 {no} 頁圖表的數列「{nm}」不見了（輸出第 {best[0]} 頁）")
            elif got[0] != kind:
                msgs.append(f"來源第 {no} 頁圖表的「{nm}」原本是{kind}，輸出變成{got[0]}：用 src_chart() 的 line")
            elif got[1] != vals:
                msgs.append(f"來源第 {no} 頁圖表的「{nm}」數值和來源不同：數值要用 src_chart() 讀，不要手打")
    return msgs


# ---------- 讀來源（數值、文字一律用程式讀，不要手打） ----------
_SKIP_PH = (PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.DATE)


def _texts_of(shapes, tables=True):
    """依閱讀順序（上→下、左→右）回傳 (top, left, 文字) list；群組會進入，換行拆成多段。"""
    out = []
    for sh_ in walk_shapes(shapes):
        if sh_.is_placeholder and sh_.placeholder_format.type in _SKIP_PH:
            continue
        top, left = (sh_.top or 0) / 914400, (sh_.left or 0) / 914400
        if sh_.has_text_frame:
            for p in sh_.text_frame.paragraphs:
                for part in p.text.replace("\x0b", "\n").split("\n"):
                    if part.strip():
                        out.append((round(top, 2), round(left, 2), part.strip()))
        elif tables and getattr(sh_, "has_table", False) and sh_.has_table:
            for r in sh_.table.rows:
                for c in r.cells:
                    for part in c.text.replace("\x0b", "\n").split("\n"):
                        if part.strip():
                            out.append((round(top, 2), round(left, 2), part.strip()))
    out.sort(key=lambda t: (t[0], t[1]))
    return out


def src_paras(slide_no, pos=False):
    """來源第 slide_no 頁（從 1 起算）所有文字框的文字，一行一筆，依位置由上而下排。
    同一個文字框裡的多行（例如「有效: 120筆」「到期: 35筆」）會拆開。
    表格不含在內（用 src_tables）。pos=True 時回傳 (top吋, left吋, 文字)，可用位置判斷哪個標籤配哪個值。"""
    rows = _texts_of(src.slides[slide_no - 1].shapes, tables=False)
    return rows if pos else [t for _, _, t in rows]


def _abs_boxes(shapes, tf=(0.0, 0.0, 1.0, 1.0)):
    """所有有文字的形狀（含群組裡的），回傳 (shape, 左, 上, 寬, 高) 吋，群組內的座標換算成投影片座標。"""
    ox, oy, sx, sy = tf
    for sh_ in shapes:
        l, t = (sh_.left or 0) / EMU, (sh_.top or 0) / EMU
        w, h = (sh_.width or 0) / EMU, (sh_.height or 0) / EMU
        L_, T_, W_, H_ = ox + l * sx, oy + t * sy, w * sx, h * sy
        if sh_.shape_type == 6:
            x = sh_._element.find(".//" + qn("a:chOff")); e = sh_._element.find(".//" + qn("a:chExt"))
            cx, cy = (int(x.get("x")) / EMU, int(x.get("y")) / EMU) if x is not None else (0, 0)
            cw, ch = (int(e.get("cx")) / EMU, int(e.get("cy")) / EMU) if e is not None else (w, h)
            nsx, nsy = (W_ / cw if cw else sx), (H_ / ch if ch else sy)
            yield from _abs_boxes(sh_.shapes, (L_ - cx * nsx, T_ - cy * nsy, nsx, nsy))
        elif sh_.has_text_frame:
            yield sh_, L_, T_, W_, H_


def _lines(sh_):
    return [norm(x) for p in sh_.text_frame.paragraphs for x in p.text.replace("\x0b", "\n").split("\n") if x.strip()]


def src_title(slide_no):
    """來源頁的標題文字 list（標題預留位置，和最上方 0.6 吋內的單行字，例如章節名、頁標題）。
    這些字放進 new_slide 的標題就好，不要再放進內文。"""
    boxes = [(sh_, l, t, w, h, _lines(sh_)) for sh_, l, t, w, h in _abs_boxes(src.slides[slide_no - 1].shapes)]
    boxes = [b for b in boxes if b[5]]
    def size(sh_):
        ss = [r.font.size.pt for p in sh_.text_frame.paragraphs for r in p.runs if r.font.size]
        return max(ss) if ss else 0
    cand = [b for b in boxes if b[2] < 0.6 and len(b[5]) == 1]
    out = []
    for sh_, l, t, w, h, ln in boxes:
        ph = sh_.is_placeholder and sh_.placeholder_format.type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)
        if ph:
            out.append((999, " ".join(ln)))
        elif t < 0.6 and len(ln) == 1:
            # 下方 1.2 吋內、左右重疊的地方有內文的，是小標（例如「第一期」），不是頁標題
            below = any(b[0] is not sh_ and b not in cand and 0 < b[2] - t <= 1.2
                        and min(l + w, b[1] + b[3]) - max(l, b[1]) > 0.1 for b in boxes)
            if not below:
                out.append((size(sh_), ln[0]))
    out.sort(key=lambda x: -x[0])  # 字最大的排第一個
    return [t for _, t in out]


def src_cards(slide_no):
    """把來源頁散放的文字框，依位置分成 [(小標, [內容行...]), ...]，可以直接 card_grid(s, src_cards(N))。
    小標＝短的單行（或兩行）文字框，內容＝它下方、左右重疊的文字框。依「上→下、同一列左→右」排序。
    沒有小標的內容，小標是 ""；只有小標沒有內容的，內容是 []。標題（src_title）不包含在內。
    適合：並排的卡片、左右兩欄、「日期＋說明」這類來源頁。不要把整頁 src_paras() 倒成一串清單。"""
    titles = set(src_title(slide_no))
    heads, bodies = [], []
    for sh_, l, t, w, h in _abs_boxes(src.slides[slide_no - 1].shapes):
        if getattr(sh_, "has_table", False):
            continue
        ln = _lines(sh_)
        if not ln or (len(ln) == 1 and ln[0] in titles and t < 0.6) or (
                sh_.is_placeholder and sh_.placeholder_format.type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)):
            continue
        rec = dict(l=l, t=t, r=l + w, b=t + h, lines=ln)
        (heads if len(ln) <= 2 and sum(len(x) for x in ln) <= 20 and h <= 0.9 else bodies).append(rec)
    # 只有一行的短字，如果正上方 1.2 吋內有另一個小標，是那個小標的內容（例如小標下的「無」）
    heads.sort(key=lambda r: r["t"])
    for hd in list(heads):
        above = [o for o in heads if o is not hd and 0 < hd["t"] - o["t"] <= 1.2 and len(hd["lines"]) == 1
                 and min(o["r"], hd["r"]) - max(o["l"], hd["l"]) > 0.1]
        if above and not any(b["t"] > max(a["t"] for a in above) and b["t"] < hd["t"] for b in bodies):
            heads.remove(hd); bodies.append(hd)
    cards = [dict(head=hd, items=[]) for hd in heads]
    orphans = []
    for bd in sorted(bodies, key=lambda r: r["t"]):
        best = None
        for c in cards:
            hd = c["head"]
            if hd["t"] <= bd["t"] + 0.3 and min(hd["r"], bd["r"]) - max(hd["l"], bd["l"]) > -0.2:
                d = bd["t"] - hd["t"]
                if best is None or d < best[0]:
                    best = (d, c)
        if best:
            best[1]["items"] += bd["lines"]
        else:
            orphans.append(dict(head=dict(l=bd["l"], t=bd["t"], lines=[]), items=bd["lines"]))
    allc = cards + orphans
    allc.sort(key=lambda c: c["head"]["t"])
    rows = []
    for c in allc:  # 上緣差 0.5 吋內算同一列
        if rows and abs(c["head"]["t"] - rows[-1][0]["head"]["t"]) < 0.5:
            rows[-1].append(c)
        else:
            rows.append([c])
    return [(" ".join(c["head"]["lines"]), c["items"]) for r in rows for c in sorted(r, key=lambda c: c["head"]["l"])]


def src_tables(slide_no):
    """來源第 slide_no 頁的**所有**表格（含群組裡的），依位置由上而下、由左而右。一頁常常不只一個表格。"""
    ts = [sh_ for sh_ in walk_shapes(src.slides[slide_no - 1].shapes)
          if getattr(sh_, "has_table", False) and sh_.has_table]
    ts.sort(key=lambda t: ((t.top or 0), (t.left or 0)))
    return [t.table for t in ts]


def src_chart(slide_no, nth=0):
    """讀來源圖表的**全部**數列（含組合圖的折線，跳過沒有值的隱藏數列），
    回傳 dict(categories, bars, line)，可以直接 column_chart(s, **src_chart(12))。"""
    chs = [sh_ for sh_ in walk_shapes(src.slides[slide_no - 1].shapes)
           if getattr(sh_, "has_chart", False) and sh_.has_chart]
    ch = chs[nth].chart
    cats = [norm(str(c)) for c in ch.plots[0].categories]
    bars, lines = [], []
    for pl in ch.plots:
        kind = type(pl).__name__
        for se in pl.series:
            vals = list(se.values)
            if not vals or all(v is None for v in vals):
                continue
            item = (norm(se.name or ""), [v or 0 for v in vals])
            (lines if "Line" in kind else bars).append(item)
    if not bars and lines:  # 只有折線：第一條當長條
        bars.append(lines.pop(0))
    return dict(categories=cats, bars=bars, line=lines or None)


_SEP = re.compile(r"[\s:：、,，;；/()（）\[\]【】「」\"'*~\-–—|•·●○■□◆◇★☆※✓✔→]+")
_RUN = re.compile(r"[A-Za-z0-9.%&+#]+|[^\x00-\x7f]+")  # 中文和英數分開比，改寫成表格欄位也算有


def _blob(pres):
    return re.sub(r"\s+", "", "".join(t for sl in pres.slides for _, _, t in _texts_of(sl.shapes))).lower()


def missing_text(src_pres, out_pres):
    """回傳 {來源頁碼: [輸出裡找不到的字]}。比對時忽略空白和標點，所以拆成標籤＋數字也算有。"""
    blob = _blob(out_pres)
    miss = {}
    for no, sl in enumerate(src_pres.slides, 1):
        seen = []
        for _, _, t in _texts_of(sl.shapes):
            for tok in (x for part in _SEP.split(t) for x in _RUN.findall(part)):
                tok = tok.strip(".。").lower()
                if len(tok) < 2 and not tok.isdigit():
                    continue
                if tok not in blob and tok not in seen:
                    seen.append(tok)
        if seen:
            miss[no] = seen
    return miss


def _toks(sl):
    out = []
    for _, _, t in _texts_of(sl.shapes):
        for tok in (x for part in _SEP.split(t) for x in _RUN.findall(part)):
            tok = tok.strip(".。").lower()
            if len(tok) >= 2 and not tok.isdigit() and tok not in out:
                out.append(tok)
    return out


def moved_text(src_pres, out_pres):
    """找出「來源某頁的字，只出現在對應別的來源頁的輸出頁」：通常是為了消 MISSING 把內容搬到別頁。
    每個輸出頁先找最像的來源頁（共同字最多），再看來源頁獨有的字有沒有跑到別頁。回傳訊息 list。"""
    src_t, out_b, owner = _owners(src_pres, out_pres)
    msgs = []
    for i, ts in enumerate(src_t):
        mine = [k for k, o in enumerate(owner) if o == i]
        if not mine:
            continue
        others = set().union(*[set(t) for j, t in enumerate(src_t) if j != i]) if len(src_t) > 1 else set()
        where = {}
        for tok in ts:
            if (tok.isascii() and len(tok) < 4) or tok in others or any(tok in out_b[k] for k in mine):
                continue
            for k, b in enumerate(out_b):
                if tok in b:
                    where.setdefault(k, []).append(tok)
                    break
        for k, toks in where.items():
            o = owner[k]
            pages = "、".join(str(m + 1) for m in mine)
            msgs.append(f"來源第 {i + 1} 頁的字放到了輸出第 {k + 1} 頁（那頁對應來源第 {o + 1 if o is not None else '?'} 頁）："
                        + "、".join(toks[:8]) + (f" …等 {len(toks)} 處" if len(toks) > 8 else "")
                        + f"。要放回輸出第 {pages} 頁（版面不夠就縮小表格、用 put_table 讓出側欄），不要搬到別頁")
    return msgs


def structure_check(src_pres, out_pres):
    """兩種「內容都在、但版面亂掉」的情況：
    1. 來源是好幾個分開的區塊（小標＋內容），輸出卻變成一個文字框的一串清單 → 用 card_grid(s, src_cards(N))
    2. 內文裡又出現一次標題（把 src_paras() 整頁倒進去時常見）"""
    global src
    msgs = []
    _, _, owner = _owners(src_pres, out_pres)
    keep, src = src, src_pres
    try:
        for k, sl in enumerate(out_pres.slides):
            ttl = sl.shapes.title.text_frame.text.strip() if sl.shapes.title is not None else ""
            boxes = [sh_ for sh_ in walk_shapes(sl.shapes) if sh_.has_text_frame and sh_.text_frame.text.strip()
                     and not (sh_.is_placeholder and sh_.placeholder_format.type in
                              (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE, *_SKIP_PH))]
            if ttl:
                for sh_ in boxes:
                    if sh_.name == "Subtitle" or (sh_.is_placeholder and sh_.placeholder_format.type == PP_PLACEHOLDER.SUBTITLE):
                        continue
                    if any(norm(x) == norm(ttl) for x in _lines(sh_)):
                        msgs.append(f"第 {k + 1} 頁：內文又出現一次標題「{ttl}」：來源的標題（src_title）放標題就好，從內文拿掉")
                        break
            o = owner[k]
            other = [sh_ for sh_ in walk_shapes(sl.shapes)
                     if getattr(sh_, "has_table", False) and sh_.has_table or getattr(sh_, "has_chart", False) and sh_.has_chart]
            if o is None or other or len(boxes) > 1:
                continue
            cards = [c for c in src_cards(o + 1) if c[0] and c[1]]
            if len(cards) >= 2:
                msgs.append(f"第 {k + 1} 頁：來源第 {o + 1} 頁是 {len(cards)} 個分開的區塊（"
                            + "、".join(c[0] for c in cards[:4]) + "），輸出變成一整串清單："
                            "用 card_grid(s, src_cards(" + str(o + 1) + "), cols=...) 保留分區")
    finally:
        src = keep
    return msgs


def _owners(src_pres, out_pres):
    """回傳 (每個來源頁的字, 每個輸出頁的文字, 每個輸出頁對應的來源頁 index 或 None)。"""
    src_t = [_toks(sl) for sl in src_pres.slides]
    out_b = [re.sub(r"\s+", "", "".join(t for _, _, t in _texts_of(sl.shapes))).lower() for sl in out_pres.slides]
    # 輸出頁照來源順序排：用動態規劃找「來源頁碼不倒退」且共同字總數最多的對應
    scores = [[sum(1 for t in ts if t in b) for ts in src_t] for b in out_b]
    ks = [k for k, sc in enumerate(scores) if sc and max(sc) >= 2]
    owner = [None] * len(out_b)
    n = len(src_t)
    if ks and n:
        best = [scores[ks[0]][:]]
        back = [[None] * n]
        for k in ks[1:]:
            row, br, run, arg = [], [], -1, 0
            for i in range(n):
                if best[-1][i] > run:
                    run, arg = best[-1][i], i
                row.append(run + scores[k][i]); br.append(arg)
            best.append(row); back.append(br)
        i = max(range(n), key=lambda j: best[-1][j])
        for step in range(len(ks) - 1, -1, -1):
            owner[ks[step]] = i
            if step:
                i = back[step][i]
    return src_t, out_b, owner

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


def set_subtitle(s, text, nth=0, align=None, pos=None):
    """把文字放進第 nth 個（從 0 起算）副標題/內文預留位置；不同範本的 idx 不同，不要寫死 idx。
    封面常有兩個：nth=0 放講者/部門、nth=1 放日期。
    內容頁（有一般標題、副標預留位置又小又在上方）沒給 pos 時，會自動把副標移到標題正下方、
    寬度拉滿、靠左，之後 content_area() 會從副標下方開始。
    align：PP_ALIGN.LEFT 等（預留位置常繼承置中）。
    pos：(x, y, w, h) 英吋，要自己指定位置時用。
    版面沒有副標預留位置時：內容頁會在標題下方加一個文字框代替（印出提示）；其他頁回傳 None 並警告。"""
    from pptx.enum.shapes import PP_PLACEHOLDER as P
    phs = sorted([ph for ph in s.placeholders
                  if ph.placeholder_format.type in (P.SUBTITLE, P.BODY) and ph.has_text_frame],
                 key=lambda ph: ph.placeholder_format.idx)
    if nth >= len(phs):
        t = s.shapes.title
        if nth == 0 and t is not None and t.top is not None:  # 內容頁：在標題下方補一個文字框
            W, _ = slide_size()
            x0, y0 = t.left / 914400, (t.top + t.height) / 914400
            x0, y0, w0, h0 = pos or (x0, y0, W - x0 - 0.5, 0.45)
            print(f"set_subtitle：版面「{s.slide_layout.name}」沒有副標預留位置，改在標題下方加文字框")
            return txt(s, x0, y0, w0, h0, [dict(text=text, size=16, bold=True, color=BLUE)],
                       align=align or PP_ALIGN.LEFT, name="Subtitle")
        warn.append(f"set_subtitle：版面「{s.slide_layout.name}」沒有第 {nth} 個副標預留位置，「{text[:12]}」沒有放上去")
        return None
    ph = phs[nth]
    W, H = slide_size()
    t = s.shapes.title
    is_content = t is not None and t.placeholder_format.type == P.TITLE
    if pos is None and is_content and ph.top is not None and ph.top / 914400 < H * 0.3             and ph.height / 914400 < 1.0 and t.top is not None:
        tb = (t.top + t.height) / 914400
        pos = (t.left / 914400, tb, W - t.left / 914400 - 0.5, 0.45)
        if align is None:
            align = PP_ALIGN.LEFT
    if pos:
        ph.left, ph.top, ph.width, ph.height = (Inches(v) for v in pos)
    ph.text = text
    if align is not None:
        for p in ph.text_frame.paragraphs:
            p.alignment = align
    return ph


def drop_empty_placeholders(s):
    for ph in list(s.placeholders):
        if not ph.has_text_frame or not ph.text_frame.text.strip():
            ph._element.getparent().remove(ph._element)


def slide_size():
    return prs.slide_width / 914400, prs.slide_height / 914400


def content_area(s=None, margin=0.5, bottom=0.45):
    """回傳可放內容的 (x, y, w, h)（英吋）：左右留 margin，從標題（和已填的副標）下方開始。"""
    W, H = slide_size()
    top = 1.2 * H / 7.5
    if s is not None:
        t = s.shapes.title
        if t is not None and t.top is not None:
            top = max(top, (t.top + t.height) / 914400 + 0.15)
        for ph in s.placeholders:  # 已填字、位在上方的小預留位置（副標）
            if ph is t or not ph.has_text_frame or not ph.text_frame.text.strip() or ph.top is None:
                continue
            if ph.top / 914400 < H * 0.3 and ph.height / 914400 < 1.0:
                top = max(top, (ph.top + ph.height) / 914400 + 0.15)
        for sh_ in s.shapes:  # set_subtitle 補的副標文字框
            if sh_.name == "Subtitle" and not sh_.is_placeholder:
                top = max(top, (sh_.top + sh_.height) / 914400 + 0.15)
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
    lines = (text if isinstance(text, list) else [text]) or [""]  # 空 list（cell_lines 讀到空白格）也要設字級
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
    """儲存格的文字，一行一筆：段落和段落內的手動換行（Shift+Enter）都會拆開。"""
    out = []
    for p in c.text_frame.paragraphs:
        for part in p.text.replace("\x0b", "\n").split("\n"):
            if norm(part):
                out.append(norm(part))
    return out

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

def bar_line_combo(chart, line_series_index, color=None, labels=True):
    """Turn one series of a clustered bar/column chart into a line on the same axes (combo chart).
    color 預設用 init() 之後的主色 BLUE。"""
    color = color or BLUE
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
    nxt = bar.getnext()
    if nxt is not None and nxt.tag == "{%s}lineChart" % C_NS:  # 已經有折線群組：加進去
        nxt.find("{%s}marker" % C_NS).addprevious(tot)
        return
    line = etree.Element("{%s}lineChart" % C_NS)
    etree.SubElement(line, "{%s}grouping" % C_NS).set("val", "standard")
    etree.SubElement(line, "{%s}varyColors" % C_NS).set("val", "0")
    line.append(tot)
    etree.SubElement(line, "{%s}marker" % C_NS).set("val", "1")
    for ax in bar.findall("{%s}axId" % C_NS):
        etree.SubElement(line, "{%s}axId" % C_NS).set("val", ax.get("val"))
    bar.addnext(line)


# =====================================================================
# 常用版型：不含任何內容，給對應腳本直接呼叫。位置預設用 content_area(s)。
# =====================================================================
LINE = 1.3  # 中文字型實際行高約字級的 1.2–1.35 倍，估高度時用 1.3


def lines_needed(text, size, w_in):
    """一段文字在寬 w_in 吋、字級 size 時大約幾行（會算換行字元，並預留斷字的浪費）。"""
    n = 0
    for part in str(text).split("\n"):
        n += max(1, math.ceil(text_w(part, size) * 1.08 / max(w_in * 72, 1)))
    return n


def row_height(cells, widths, size, min_h=0.3, pad=0.1, line=LINE):
    """表格一列需要的高度（吋）。cells 的每格可以是字串或字串 list（多個段落）。"""
    n = 1
    for v, w in zip(cells, widths):
        parts = v if isinstance(v, list) else [v]
        n = max(n, sum(lines_needed(p, size, w - 0.14) for p in parts) if parts else 1)
    return max(min_h, n * size * line / 72 + pad)


def col_widths(rows, size, total, min_w=0.6, max_w=5.0):
    """依內容長度分配欄寬，總寬 total 吋。"""
    def flat(v): return " ".join(v) if isinstance(v, list) else str(v)
    ncol = max(len(r) for r in rows)
    w = [max(min_w, min(max(text_w(flat(r[c]), size) for r in rows if c < len(r)) / 72 + 0.3, max_w))
         for c in range(ncol)]
    return fit(w, total)


def paginate(heights, caps):
    """把列高分頁。caps：每頁可用高度（list，最後一個值用在之後所有頁）。
    頁數最少，且盡量讓各頁一樣高（不會最後一頁只剩一兩列）。回傳 [[列索引...], ...]。"""
    def greedy(limit):
        pages, cur, used = [], [], 0.0
        for i, hh in enumerate(heights):
            cap = min(limit, caps[min(len(pages), len(caps) - 1)])
            if cur and used + hh > cap:
                pages.append(cur); cur, used = [], 0.0
            cur.append(i); used += hh
        pages.append(cur)
        return pages
    if not heights:
        return [[]]
    best = greedy(max(caps))
    n = len(best)
    if n == 1:
        return best
    lo, hi = max(max(heights), sum(heights) / n) * 0.999, max(caps)
    for _ in range(30):
        mid = (lo + hi) / 2
        if len(greedy(mid)) <= n:
            hi = mid
        else:
            lo = mid
    return greedy(hi)


def is_num(t):
    t = " ".join(t) if isinstance(t, list) else str(t)
    return bool(re.fullmatch(r"[-+]?[\d,.]+\s*%?", t.strip()))


def table_pages(title, rows, first=None, top=None, sizes=(16, 14, 12, 11, 10), widths=None,
                bold_last=False, aligns=None, layout="content", name="資料表", fill_ratio=0.55,
                max_pages=None, area=None):
    """把表格（rows[0] 是表頭）放進一頁或多頁內容頁，回傳用到的投影片 list。
    - 自動挑字級：在「頁數最少」的前提下用最大的字級（小表格會放大，不會縮在左上角）。
    - 依每列文字的換行數估列高，分頁時讓各頁一樣高；多頁時標題加 (1/n)。
    - max_pages=1：**來源只有一頁的表格用這個**。放不下時會再縮小字級（最小 8pt）、
      把列高壓緊；還是放不下才分頁並發出警告。沒給 max_pages 而表格被拆頁時也會警告。
    - first：已建立的投影片當第一頁（上方已放說明文字時，用 top 指定表格起點 y）。
    - area=(x, y, w, h)：只用第一頁的這個範圍（例如右邊要留側欄），不會自動分頁到別頁以外的區域。
    - aligns：{欄: PP_ALIGN}；沒給時數字欄置中。bold_last：最後一列（合計）粗體＋淺底。"""
    ncol = max(len(r) for r in rows)
    rows = [(list(r) + [""] * ncol)[:ncol] for r in rows]
    head, body = rows[0], rows[1:]
    s0 = first or new_slide(layout, title)
    x, y, cw, ch = area or content_area(s0)
    y1 = top if top is not None else y
    avail1, avail = ch - (y1 - y), ch

    def plan(size, compact):
        w = widths or col_widths(rows, size, cw)
        # compact：儲存格上下邊界 0.01 吋、行高用 1.2 倍（PowerPoint 小字的實際行高）
        pad, ln = (0.02, 1.2) if compact else (0.1, LINE)
        min_h = size * ln / 72 + (0.02 if compact else 0.12)
        hh = row_height(head, w, size, min_h, pad, ln) + (0 if compact else 0.04)
        bh = [row_height(r, w, size, min_h, pad, ln) for r in body]
        pages = paginate(bh, [avail1 - hh, avail - hh])
        return (len(pages), size, w, hh, bh, pages, compact)

    tries = [(sz, False) for sz in sorted(sizes, reverse=True)]
    if max_pages:
        small = sorted({sz for sz in (10, 9, 8.5, 8) if sz <= min(sizes)} | set(sizes), reverse=True)
        tries += [(sz, True) for sz in small]
    best = None
    for size, compact in tries:
        cand = plan(size, compact)
        if best is None or cand[0] < best[0]:
            best = cand
        if max_pages and cand[0] <= max_pages:
            best = cand
            break
    if not max_pages:
        npg = best[0]
        for size in sorted(sizes, reverse=True):  # 同樣頁數下用最大字級
            cand = plan(size, False)
            if cand[0] == npg:
                best = cand
                break
    _, size, w, hh, bh, pages, compact = best
    if len(pages) > (max_pages or 1):
        warn.append(f"table '{title}' 拆成 {len(pages)} 頁（{size}pt）。來源只有一頁的表格要用 max_pages=1；"
                    "8pt 還放不下時先問使用者要不要拆頁")
    elif max_pages and size < 10:
        print(f"table_pages '{title}'：為了放進 {max_pages} 頁用 {size}pt，請在回報裡請使用者確認字是否太小")
    if aligns is None:
        aligns = {c: C for c in range(ncol) if c > 0 and all(is_num(r[c]) for r in body if r[c])}
    slides = []
    for pi, idx in enumerate(pages):
        t_title = title + (f" ({pi + 1}/{len(pages)})" if len(pages) > 1 else "")
        if pi == 0:
            s = s0
            if s.shapes.title is not None:
                s.shapes.title.text = t_title
            ty, cap = y1, avail1
        else:
            s = new_slide(layout, t_title)
            ty, cap = y, avail
        hs = [bh[i] for i in idx]
        if len(pages) == 1 and hh + sum(hs) < cap * fill_ratio:  # 少列的表：列高放大一點比較好讀
            k = min(1.5, cap * fill_ratio / (hh + sum(hs)))
            hs = [v * k for v in hs]
        if ty + hh + sum(hs) > ty + cap + 0.05:
            warn.append(f"table '{t_title}' 估計高度 {hh + sum(hs):.2f} 超過可用 {cap:.2f}")
        t = table(s, x, ty, w, len(idx) + 1, [hh] + hs, name=name)
        style_header_row(t, 0, head, size)
        part = [body[i] for i in idx]
        zebra_body(t, 1, part, size, aligns, bold_last=bold_last and pi == len(pages) - 1)
        if compact:
            for r_ in t.rows:
                for c_ in r_.cells:
                    c_.margin_top = c_.margin_bottom = Inches(0.01)
        slides.append(s)
    return slides


def put_table(s, rows, x, y, w, size=11, widths=None, aligns=None, bold_last=False, header=True,
              compact=True, name="資料表"):
    """把一個表格放在指定位置（不分頁）：一頁放兩個表格、表格旁邊要留側欄時用。
    rows[0] 是表頭（header=False 時全部當內容）；格子可以是字串或字串 list（多段落）。
    回傳 (table, 下緣 y)，下一個元件從下緣往下排。超出投影片底部會在 save() 時警告。"""
    ncol = max(len(r) for r in rows)
    rows = [(list(r) + [""] * ncol)[:ncol] for r in rows]
    w_ = widths and fit(widths, w) or col_widths(rows, size, w)
    pad = 0.04 if compact else 0.1
    min_h = size * LINE / 72 + (0.05 if compact else 0.12)
    hs = [row_height(r, w_, size, min_h, pad) for r in rows]
    body = rows[1:] if header else rows
    if aligns is None:
        aligns = {c: C for c in range(ncol) if c > 0 and body and all(is_num(r[c]) for r in body if r[c])}
    t = table(s, x, y, w_, len(rows), hs, name=name)
    if header:
        style_header_row(t, 0, rows[0], size)
    zebra_body(t, 1 if header else 0, body, size, aligns, bold_last=bold_last)
    bottom = y + sum(hs)
    if bottom > slide_size()[1] - 0.2:
        warn.append(f"put_table '{name}' 下緣 {bottom:.2f} 超出投影片（高 {slide_size()[1]:.2f}）")
    return t, bottom


def run_color(run):
    """讀文字顏色（用顏色表示狀態時用）：回傳 'FF0000' 這種 hex、'accent1'/'tx1'/'bg1' 這種主題色名稱
    （有調亮暗時加上 '+lumMod75000' 之類），沒設定顏色時回傳 None。不會因為主題色而出錯。"""
    rPr = run._r.find(qn("a:rPr"))
    if rPr is None:
        return None
    fill = rPr.find(qn("a:solidFill"))
    if fill is None or len(fill) == 0:
        return None
    clr = fill[0]
    val = clr.get("val")
    mods = "".join(f"+{etree.QName(m).localname}{m.get('val')}" for m in clr)
    return (val.upper() if etree.QName(clr).localname == "srgbClr" else val) + mods


def stat_cards(s, items, x=None, y=None, w=None, h=1.4, gap=0.25, value_size=36):
    """一排數字卡片：小標籤（灰）＋大數字（主色）＋單位（灰）。
    items：[(標籤, 數值), (標籤, 數值, 單位), ...]。回傳卡片下緣的 y。"""
    ax, ay, aw, _ = content_area(s)
    x = ax if x is None else x
    y = ay if y is None else y
    w = aw if w is None else w
    n = len(items)
    cw = (w - gap * (n - 1)) / n
    for i, it in enumerate(items):
        lab, val, unit = (list(it) + ["", ""])[:3]
        val, unit = str(val).strip(), str(unit or "").strip()
        if unit and val.endswith(unit):  # 「120筆」＋單位「筆」→ 不要變成「120筆 筆」
            val = val[: -len(unit)].strip()
        if re.search(r"[\n\x0b]", val) or len(val) > 10 or not re.search(r"\d", val):
            warn.append(f"stat_cards 數值「{norm(val)[:20]}」不像單一數字（有換行、太長或沒有數字），"
                        "可能把兩行解析成一個值：用 src_paras() 一行一行讀")
        cx = x + i * (cw + gap)
        box(s, cx, y, cw, h, CLOUD)
        txt(s, cx + 0.15, y + 0.08, cw - 0.3, 0.4, [dict(text=str(lab), size=14, color=GRAY)])
        vs = value_size
        while vs > 18 and text_w(str(val), vs) + text_w(" " + str(unit), 16) > (cw - 0.4) * 72:
            vs -= 2
        runs = [dict(text=str(val), size=vs, bold=True, color=BLUE)]
        if unit:
            runs.append(dict(text=" " + str(unit), size=16, color=GRAY))
        txt(s, cx + 0.15, y + 0.45, cw - 0.3, h - 0.5, [dict(runs=runs, size=vs)])
    return y + h


def agenda(s, items, gap=0.25):
    """目錄頁：編號＋項目的卡片。5 項以內排一列，更多排兩列。"""
    x, y, w, hh = content_area(s)
    n = len(items)
    cols = n if n <= 5 else math.ceil(n / 2)
    rws = math.ceil(n / cols)
    cw = (w - gap * (cols - 1)) / cols
    ch = min(3.4, (hh - 0.3 - gap * (rws - 1)) / rws)
    y0 = y + max(0, (hh - (ch * rws + gap * (rws - 1))) / 2)
    for i, it in enumerate(items):
        cx, cy = x + (i % cols) * (cw + gap), y0 + (i // cols) * (ch + gap)
        box(s, cx, cy, cw, ch, CLOUD)
        ns = 54 if ch >= 2.5 else 36
        txt(s, cx + 0.2, cy + 0.25, cw - 0.4, ns / 72 * 1.3 + 0.1, [dict(text=f"{i + 1:02d}", size=ns, bold=True, color=BLUE)])
        size = 28 if cw >= 2.5 else 20
        while size > 14 and lines_needed(it, size, cw - 0.5) > 2:
            size -= 2
        txt(s, cx + 0.2, cy + ch * 0.56, cw - 0.4, ch * 0.4, [dict(text=it, size=size, bold=True, color=NAVY)])


def card_grid(s, cards, cols=2, gap=0.23, title_size=18, body_size=14, area=None):
    """多張卡片排成格狀（例如 2×2）。cards：[(卡片標題, 項目 list), ...]；
    項目可以是字串，或 (主文字, 灰色補充說明)。回傳卡片的 (x, y, w, h) list。"""
    x, y, w, hh = area or content_area(s)
    n = len(cards)
    rws = math.ceil(n / cols)
    cw, ch = (w - gap * (cols - 1)) / cols, (hh - gap * (rws - 1)) / rws
    out = []
    def build(items, bs):
        paras = []
        for it in items:
            if isinstance(it, (tuple, list)):
                main, sub = (list(it) + [""])[:2]
                paras.append(dict(text=str(main), size=bs, bold=True, color=BLACK, after=0 if sub else 4))
                if sub:
                    paras.append(dict(text=str(sub), size=bs - 2, color=GRAY, after=6))
            else:
                paras.append(dict(text=str(it), size=bs, color=BLACK, after=6))
        return paras
    # 字級：所有卡片用同一個字級，放不下就縮小（最小 10pt），還是放不下 txt 會發出 overflow 警告
    bs = body_size
    while bs > 10 and any(
            est_h([(p["text"], p["size"], p["after"]) for p in build(items, bs)], cw - 0.5) + 0.06 > ch - 0.75
            for _, items in cards):
        bs -= 1
    # 內容少時卡片不要撐滿整頁：高度改成內容需要的高度（最少 1.6 吋）
    need = max((est_h([(p["text"], p["size"], p["after"]) for p in build(items, bs)], cw - 0.5) for _, items in cards),
               default=0) + 0.95
    ch = min(ch, max(need, 1.6))
    for i, (ttl, items) in enumerate(cards):
        cx, cy = x + (i % cols) * (cw + gap), y + (i // cols) * (ch + gap)
        box(s, cx, cy, cw, ch, CLOUD)
        txt(s, cx + 0.2, cy + 0.12, cw - 0.4, 0.45, [dict(text=ttl, size=title_size, bold=True, color=BLUE)])
        paras = build(items, bs)
        if paras:
            txt(s, cx + 0.2, cy + 0.65, cw - 0.4, ch - 0.75, paras)
        else:
            warn.append(f"card_grid 卡片「{ttl}」沒有內容：檢查解析程式")
        out.append((cx, cy, cw, ch))
    if bs < body_size:
        print(f"card_grid：內容較多，內文縮到 {bs}pt")
    return out


def progress_bar(s, x, y, w, pct, h=0.14, track=None, color=None):
    """進度條（例如達成率）。pct 是 0–100。"""
    for ww, col in ((w, track or WHITE), (max(w * min(float(pct), 100) / 100, 0.05), color or BLUE)):
        r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(ww), Inches(h))
        r.fill.solid(); r.fill.fore_color.rgb = rgb(col); r.line.fill.background(); r.shadow.inherit = False


def column_chart(s, categories, bars, line=None, area=None, colors=None, labels=True, name="圖表"):
    """原生直條圖（可加一條折線，例如合計）。
    bars：[(數列名稱, 數值 list), ...]；line：(名稱, 數值 list)、多條時用 list，或 None。回傳 chart。
    來源是圖表時直接 column_chart(s, **src_chart(頁碼))，不會漏掉組合圖的折線。"""
    if line and isinstance(line[0], str):
        line = [line]
    line = list(line or [])
    x, y, w, hh = area or content_area(s)
    cd = CategoryChartData()
    cd.categories = [str(c) for c in categories]
    allv = []
    for nm, vals in bars:
        v = [float(a or 0) for a in vals]; allv += v
        cd.add_series(nm, v)
    for nm, vals in line:
        cd.add_series(nm, [float(a or 0) for a in vals])
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(x), Inches(y), Inches(w), Inches(hh), cd)
    gf.name = name
    ch = gf.chart
    nser = len(bars) + len(line)
    ch.has_legend = nser > 1
    if ch.has_legend:
        ch.legend.position = XL_LEGEND_POSITION.TOP
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(12)
    pal = colors or ([CYAN, ICE, NAVY, GRAY] if line else [BLUE, ICE, CYAN, NAVY, GRAY])
    for se, col in zip(list(ch.plots[0].series)[:len(bars)], pal * 3):
        se.format.fill.solid(); se.format.fill.fore_color.rgb = rgb(col)
    for k in range(len(line)):  # 每次移走一條，下一條的索引仍是 len(bars)
        bar_line_combo(ch, len(bars), [BLUE, NAVY, GRAY, CYAN][k % 4], labels)
    pl = ch.plots[0]
    pl.gap_width = 60; pl.overlap = 0
    if labels:
        pl.has_data_labels = True
        dl = pl.data_labels
        dl.font.size = Pt(9)
        whole = all(abs(v - round(v)) < 1e-9 for v in allv)
        dl.number_format = "#,##0;-#,##0;;" if whole else "#,##0.0;-#,##0.0;;"
        dl.number_format_is_linked = False
        dl.position = XL_LABEL_POSITION.OUTSIDE_END
    ca = ch.category_axis
    ca.tick_labels.font.size = Pt(10)
    ca.format.line.color.rgb = rgb(GRAY)
    if len(categories) > 8 or max((text_w(str(c), 10) for c in categories), default=0) > 60:
        txPr = ca._element.find(qn("c:txPr"))
        if txPr is not None:
            txPr.find(qn("a:bodyPr")).set("rot", "-2700000")
    va = ch.value_axis
    va.tick_labels.font.size = Pt(10)
    va.major_gridlines.format.line.color.rgb = rgb(CLOUD)
    va.format.line.fill.background()
    return ch

