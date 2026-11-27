"""產生測試用的來源簡報和通用範本（內容全部是虛構的，不含任何真實資料）。

用法：python tests/make_samples.py [輸出資料夾]   （預設 tests/samples）
產生：
  src_a.pptx   16:9，常見月報結構：封面、目錄、表格＋說明、表格＋側欄、圖表、卡片、數字、條列、結尾
  src_b.pptx   4:3，比較難的結構：大表格、一頁兩個表格、左右兩欄、群組裡的文字、成對數字
  tpl_plain.pptx  python-pptx 內建範本（4:3，沒有結尾版面），測試換範本也能用
"""
import os
import sys

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches, Pt

sys.stdout.reconfigure(encoding="utf-8")


def tb(s, x, y, w, h, lines, size=14, bold=False):
    box = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size)
        r.font.bold = bold
    return box


def table(s, x, y, w, h, rows, red=()):
    t = s.shapes.add_table(len(rows), len(rows[0]), Inches(x), Inches(y), Inches(w), Inches(h)).table
    for r, row in enumerate(rows):
        for c, v in enumerate(row):
            t.cell(r, c).text = str(v)
            for p in t.cell(r, c).text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(10)
                    if (r, c) in red:
                        run.font.color.rgb = RGBColor(0xFF, 0, 0)
    return t


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def make_a(path):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    # 1 封面：大字標題＋報告人＋日期
    s = blank(prs)
    tb(s, 1, 2, 11, 1.5, ["Quarterly Review"], size=60, bold=True)
    tb(s, 1, 3.8, 6, 0.6, ["— 營運企劃部 —"], size=24)
    tb(s, 1, 5.5, 6, 0.5, ["2026 Oct 15"], size=18)
    # 2 目錄
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "目錄"
    tb(s, 1, 1.8, 8, 3, ["1. 銷售概況", "2. 專案進度", "3. 人力資源", "4. 費用"], size=24)
    # 3 表格＋上方說明＋下方註記
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "銷售概況"
    tb(s, 0.6, 1.3, 6, 0.5, ["本季新增客戶 12 家，流失 2 家"], size=14)
    rows = [["區域", "一月", "二月", "三月", "合計"], ["北區", 120, 135, 128, 383], ["中區", 98, 102, 110, 310],
            ["南區", 87, 91, 95, 273], ["合計", 305, 328, 333, 966]]
    table(s, 0.6, 2.0, 9, 2.5, rows)
    tb(s, 0.6, 4.8, 9, 0.5, ["註：金額單位為萬元，三月數字為預估值"], size=12)
    # 4 表格＋右側說明欄（側欄）
    s = blank(prs)
    tb(s, 0.5, 0.2, 8, 0.7, ["年度專案排程"], size=28, bold=True)
    tb(s, 3.5, 0.75, 6, 0.4, ["●進行中 ●延後 ●完成"], size=12)
    rows = [["專案", "負責單位", "Q1", "Q2", "Q3", "Q4", "狀態"]]
    for i in range(1, 19):
        rows.append([f"專案 {chr(64 + i)}", ["研發部", "製造部", "品保部"][i % 3], "●" if i % 2 else "", "●" if i % 3 else "",
                     "●" if i % 4 else "", "", ["進行中", "延後", "完成"][i % 3]])
    table(s, 0.3, 1.3, 8.6, 5.9, rows, red={(3, 3), (7, 4), (11, 2)})
    tb(s, 9.3, 1.4, 3.6, 0.5, ["本月重點"], size=16, bold=True)
    tb(s, 9.3, 2.0, 3.6, 2.5, ["10/05 專案 C 完成驗收", "10/18 專案 F 延後至 Q4", "10/27 新增專案 S 評估"], size=12)
    # 5 圖表
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "月營收趨勢"
    cd = CategoryChartData()
    cd.categories = ["一月", "二月", "三月", "四月", "五月", "六月"]
    cd.add_series("產品線甲", (32, 35, 38, 36, 41, 44))
    cd.add_series("產品線乙", (21, 22, 25, 27, 26, 30))
    s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1.5), Inches(11), Inches(5.5), cd)
    # 6 三張卡片（小標＋內容），其中一張放在群組裡
    s = blank(prs)
    tb(s, 0.5, 0.2, 8, 0.7, ["本月活動"], size=28, bold=True)
    tb(s, 0.8, 1.6, 3.5, 0.4, ["10/08"], size=16, bold=True)
    tb(s, 0.8, 2.1, 3.5, 1.5, ["新進人員訓練", "參加人數：8人"], size=13)
    tb(s, 4.9, 1.6, 3.5, 0.4, ["10/16"], size=16, bold=True)
    tb(s, 4.9, 2.1, 3.5, 1.5, ["消防演練", "全廠區"], size=13)
    grp = s.shapes.add_group_shape()
    tb(grp, 9.0, 1.6, 3.5, 0.4, ["10/25"], size=16, bold=True)
    tb(grp, 9.0, 2.1, 3.5, 1.5, ["供應商稽核", "地點：協力廠"], size=13)
    # 7 數字：統計卡片＋每季成對數字
    s = blank(prs)
    tb(s, 0.5, 0.2, 8, 0.7, ["人力統計"], size=28, bold=True)
    tb(s, 1, 1.4, 6, 0.5, ["在職人數：248人    離職率：3.2%"], size=16)
    tb(s, 1, 2.4, 4, 0.4, ["各季招募"], size=16, bold=True)
    tb(s, 1, 2.9, 10, 2, ["Q1預計招募：20  實際招募：18  達成率：90%", "Q2預計招募：15  實際招募：15  達成率：100%",
                          "Q3預計招募：25  實際招募：19  達成率：76%"], size=14)
    tb(s, 1, 5.3, 4, 0.4, ["未達成原因"], size=16, bold=True)
    # 8 一整段條列
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "費用說明"
    tb(s, 0.8, 1.5, 11, 5, ["差旅費較上季增加 8%，主因海外展覽", "設備維修費持平", "訓練費用依計畫執行",
                            "辦公用品改為集中採購，預計每季節省 5%", "水電費因夏季用電增加 12%"], size=16)
    # 9 結尾
    s = blank(prs)
    tb(s, 4, 3, 5, 1, ["謝謝聆聽"], size=40, bold=True)
    prs.save(path)


def make_b(path):
    prs = Presentation()  # 4:3
    # 1 封面用版面的標題＋副標
    s = prs.slides.add_slide(prs.slide_layouts[0])
    s.shapes.title.text = "品質月報"
    s.placeholders[1].text = "品保處　2026/09"
    # 2 大表格（30 列）
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "不良品統計"
    rows = [["料號", "名稱", "數量", "不良", "不良率"]]
    for i in range(1, 31):
        rows.append([f"P-{1000 + i}", f"零件 {i}", 500 + i * 7, i % 6, f"{(i % 6) / (500 + i * 7) * 100:.2f}%"])
    rows.append(["合計", "", sum(r[2] for r in rows[1:]), sum(r[3] for r in rows[1:]), ""])
    table(s, 0.3, 1.2, 9.4, 6.1, rows)
    # 3 一頁兩個表格＋小標
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "客訴處理"
    tb(s, 0.5, 1.3, 4, 0.4, ["本月客訴"], size=14, bold=True)
    table(s, 0.5, 1.8, 9, 1.6, [["客戶", "問題", "狀態"], ["客戶甲", "外觀刮傷", "已結案"], ["客戶乙", "包裝破損", "處理中"]])
    tb(s, 0.5, 3.8, 4, 0.4, ["待追蹤事項"], size=14, bold=True)
    table(s, 0.5, 4.3, 9, 1.2, [["事項", "負責人", "期限"], ["改善包材", "品保課", "10/31"]])
    # 4 左右兩欄
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "稽核結果"
    tb(s, 0.5, 1.5, 4.2, 0.4, ["內部稽核"], size=16, bold=True)
    tb(s, 0.5, 2.0, 4.2, 3, ["缺失 3 項", "已改善 2 項", "預計 10/20 全部完成"], size=13)
    tb(s, 5.3, 1.5, 4.2, 0.4, ["外部稽核"], size=16, bold=True)
    tb(s, 5.3, 2.0, 4.2, 3, ["客戶稽核 1 次", "無重大缺失"], size=13)
    # 5 數字卡片
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "良率"
    tb(s, 0.8, 1.6, 4, 0.5, ["本月良率：98.6%"], size=18)
    tb(s, 5, 1.6, 4, 0.5, ["目標：98%"], size=18)
    tb(s, 0.8, 2.6, 8, 2, ["連續三個月達標", "主要改善：焊接參數調整"], size=14)
    # 6 結尾（用版面的標題）
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "Q&A"
    prs.save(path)


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "samples")
    os.makedirs(out, exist_ok=True)
    make_a(os.path.join(out, "src_a.pptx"))
    make_b(os.path.join(out, "src_b.pptx"))
    Presentation().save(os.path.join(out, "tpl_plain.pptx"))
    print("已產生：", ", ".join(sorted(os.listdir(out))))


if __name__ == "__main__":
    main()
