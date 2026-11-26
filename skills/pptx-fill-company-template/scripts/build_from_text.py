"""把非 pptx 的來源（md/txt/csv/xlsx/docx）套到任何公司範本上。這是起點，不是成品：
真正交付時複製一份，依內容調整每一節要怎麼呈現。

用法：python build_from_text.py SOURCE TEMPLATE.pptx OUT.pptx
規則：
  封面（文件標題＋短副標）→ 目錄（3–8 個主章節時）→ 每一節：
    - 2–4 個「短、含數字」的條列（例如「新客戶 12 家」）→ 數字卡片
    - 段落/條列 → 淺色卡片，字級依數量調整；太多時分兩欄、再多就分頁 (1/n)
    - 表格 → 自動挑字級、依文字換行估列高、分頁時各頁平均；表格前的短說明放在表格上方
    - 小型數值表（2–12 列）→ 另加一頁原生直條圖（合計欄畫成折線）
  → 結尾頁（範本有的話）
版面自動偵測（cover / content / closing），顏色取自範本的佈景主題，尺寸依投影片大小計算（16:9、4:3 都可）。
不摘要、不編造：所有文字都來自來源；只有「目錄」「概要」「圖表」這類頁面標籤是腳本加的。
"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import helpers as H  # noqa: E402
from helpers import (LINE, agenda, box, column_chart, content_area, drop_empty_placeholders,  # noqa: E402
                     init, is_cjk, is_num, lines_needed, new_slide, paginate, save, set_subtitle,
                     stat_cards, table_pages, txt)
from extract_source import extract  # noqa: E402

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

SRC, TPL, OUT = sys.argv[1:4]
doc = extract(SRC)
init(None, TPL)
CJK = any(is_cjk(c) for c in doc["title"] + "".join(s["heading"] for s in doc["sections"]))
LBL = dict(agenda="目錄", overview="概要", chart="圖表") if CJK else dict(agenda="Agenda", overview="Overview", chart="Chart")
TOTAL = re.compile(r"^(合計|總計|小計|總和|total|sum|subtotal)$", re.I)

# ---------------- 封面：文件標題 + 標題區第一段短文字當副標 ----------------
s = new_slide("cover")
if s.shapes.title is not None:
    s.shapes.title.text = doc["title"]
used_sub = None
for sec in doc["sections"]:
    for b in sec["blocks"]:
        if b["type"] == "paragraph" and len(b["text"]) <= 40:
            used_sub = b
            break
    if used_sub or sec["blocks"]:
        break
if used_sub:
    set_subtitle(s, used_sub["text"])
drop_empty_placeholders(s)

# ---------------- 整理章節 ----------------
secs = []
for sec in doc["sections"]:
    blocks = [b for b in sec["blocks"] if b is not used_sub]
    if not blocks:
        continue
    head = sec["heading"] or doc["title"]
    if head == doc["title"]:
        head = LBL["overview"]
    secs.append((head, sec["level"], blocks))

top_level = min((lv for _, lv, _ in secs), default=1)
tops = [hd for hd, lv, _ in secs if lv == top_level and hd != LBL["overview"]]
if 3 <= len(tops) <= 8:
    s = new_slide("content", LBL["agenda"])
    agenda(s, [re.sub(r"^\d+(\.\d+)*[.、)）]?\s*", "", t) for t in tops])


# ---------------- 各種頁面 ----------------
KPI = re.compile(r"^(?P<label>[^\d]*?)\s*(?P<val>[-+]?\d[\d,.]*)\s*(?P<unit>%|[^\d\s，,。.]{0,2})(?P<rest>.*)$")


def parse_kpi(t):
    """「新客戶 12 家」→ ("新客戶", "12", "家")。太長或沒有數字回傳 None。"""
    if len(t) > 22:
        return None
    m = KPI.match(t.strip())
    if not m or not m.group("label").strip():
        return None
    label = (m.group("label") + " " + m.group("rest")).strip()
    return (label, m.group("val"), m.group("unit"))


def card_h(paras, w):
    return sum(lines_needed(p["text"], p["size"], w) * p["size"] * LINE / 72 + p["after"] / 72 for p in paras)


def text_slides(head, items, first=None, top=None):
    """items：[(文字, 是否條列)]。卡片高度跟著內容；放不下就分兩欄，再放不下就分頁。"""
    s = first or new_slide("content", head)
    x, y, w, hh = content_area(s)
    y1 = top if top is not None else y
    caps = [hh - (y1 - y) - 0.35, hh - 0.35]
    gap = 0.23

    def paras_of(its, size):
        return [dict(text=("• " if bul else "") + t, size=size, color=H.BLACK, after=10 if size >= 16 else 8)
                for t, bul in its]
    choice = None
    for size, cols in ((20, 1), (18, 1), (16, 1), (16, 2), (14, 2)):
        if cols == 2 and len(items) < 4:
            continue
        cw = w if cols == 1 else (w - gap) / 2
        hs = [card_h(paras_of([it], size), cw - 0.6) for it in items]
        pages = paginate(hs, [c * cols for c in caps])
        if len(pages) == 1 and len(paginate(hs, caps[:1])) <= cols:
            choice = (size, cols, cw, hs, pages)
            break
    if choice is None:                              # 一頁放不下：16pt 兩欄分頁
        size, cols = 16, 2
        cw = (w - gap) / 2
        hs = [card_h(paras_of([it], size), cw - 0.6) for it in items]
        choice = (size, cols, cw, hs, paginate(hs, [c * cols for c in caps]))
    size, cols, cw, hs, pages = choice
    for pi, idx in enumerate(pages):
        title = head + (f" ({pi + 1}/{len(pages)})" if len(pages) > 1 else "")
        if pi == 0:
            if s.shapes.title is not None:
                s.shapes.title.text = title
            sy, cap = y1, caps[0]
        else:
            s = new_slide("content", title)
            sy, cap = y, caps[1]
        if cols == 1:
            columns = [idx]
        else:                                       # 兩欄平均分：找最接近一半高度的切點
            tot, acc, k = sum(hs[i] for i in idx), 0.0, 0
            best = (tot, 0)
            for j, i in enumerate(idx):
                acc += hs[i]
                best = min(best, (abs(tot - 2 * acc), j + 1))
            k = best[1]
            columns = [c for c in (idx[:k], idx[k:]) if c]
        ch = min(cap + 0.35, max(1.2, max(sum(hs[i] for i in c) for c in columns) + 0.4))
        for ci, col in enumerate(columns):
            cx = x + ci * (cw + gap)
            box(s, cx, sy, cw, ch, H.CLOUD)
            txt(s, cx + 0.3, sy + 0.15, cw - 0.6, ch - 0.3, paras_of([items[i] for i in col], size),
                name=f"內容{ci + 1}")


def chart_slide(head, rows):
    """小型數值表加一頁直條圖：第一欄當類別，數字欄當數列，合計欄畫成折線，合計列略過。"""
    hdr, body = rows[0], [r for r in rows[1:] if not TOTAL.match(str(r[0]).strip())]
    ncol = len(hdr)
    num_cols = [c for c in range(1, ncol) if all(is_num(r[c]) for r in body if c < len(r) and r[c])
                and any(c < len(r) and r[c] for r in body)]
    if not (2 <= len(body) <= 12 and 1 <= len(num_cols) <= 6):
        return

    def val(v):
        try:
            return float(str(v).replace(",", "").replace("%", "").strip() or 0)
        except ValueError:
            return 0.0
    tot = [c for c in num_cols if TOTAL.match(str(hdr[c]).strip())]
    bars = [(str(hdr[c]), [val(r[c]) for r in body]) for c in num_cols if c not in tot]
    line = (str(hdr[tot[0]]), [val(r[tot[0]]) for r in body]) if len(tot) == 1 and bars else None
    if not bars:
        bars, line = [(str(hdr[tot[0]]), [val(r[tot[0]]) for r in body])], None
    s = new_slide("content", f"{head}｜{LBL['chart']}" if CJK else f"{head} | {LBL['chart']}")
    column_chart(s, [r[0] for r in body], bars, line)


for head, level, blocks in secs:
    items, kpis, tables = [], None, []
    for b in blocks:
        if b["type"] == "paragraph":
            items.append((b["text"], False))
        elif b["type"] == "bullets":
            k = [parse_kpi(t) for t in b["items"]]
            if kpis is None and 2 <= len(k) <= 4 and all(k):
                kpis = k
            else:
                items += [(t, True) for t in b["items"]]
        elif b["type"] == "table" and len(b["rows"]) > 1:
            tables.append(b["rows"])
    note = tables and not kpis and items and len(items) <= 2 and sum(len(t) for t, _ in items) <= 160
    s, top = None, None
    if kpis:
        s = new_slide("content", head)
        top = stat_cards(s, kpis) + 0.3
    if items and not note:
        text_slides(head, items, first=s, top=top)
        s, top = None, None
    elif note:                                      # 表格前的短說明：放在表格上方
        s = new_slide("content", head)
        x, y, w, _ = content_area(s)
        paras = [dict(text=t, size=14, color=H.GRAY, after=4) for t, _ in items]
        nh = card_h(paras, w - 0.1) + 0.1
        txt(s, x, y, w, nh, paras, name="說明")
        top = y + nh + 0.12
    for ti, rows in enumerate(tables):
        if ti > 0:
            s, top = None, None
        last = str(rows[-1][0]).strip() if rows[-1] else ""
        table_pages(head, rows, first=s, top=top, bold_last=bool(TOTAL.match(last)))
        chart_slide(head, rows)
        s, top = None, None

end = new_slide("closing")                          # 範本沒有結尾版面就略過
if end is not None:
    drop_empty_placeholders(end)
save(OUT)
