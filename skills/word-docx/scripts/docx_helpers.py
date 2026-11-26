"""建立 Word 文件時可 import 的輔助函式（解決中文字型問題）。

用法（在自己寫的腳本裡）：
    import sys; sys.path.insert(0, r"<本技能資料夾>/scripts")
    from docx_helpers import new_document, set_cell_bg, add_table
"""
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt


def set_run_font(run, name="微軟正黑體", size=None, bold=None):
    """設定 run 字型，包含中文（eastAsia）。只設 run.font.name 中文不會變。"""
    run.font.name = name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), name)
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold


def set_style_font(style, name="微軟正黑體", size=None):
    style.font.name = name
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), name)
    if size:
        style.font.size = Pt(size)


def new_document(font="微軟正黑體", size=12, template=None):
    """建立新文件，內文與標題都套用中文字型。template 可給公司範本 .docx。"""
    doc = Document(template) if template else Document()
    set_style_font(doc.styles["Normal"], font, size)
    for i in range(1, 4):
        try:
            set_style_font(doc.styles[f"Heading {i}"], font)
        except KeyError:
            pass
    try:
        set_style_font(doc.styles["Title"], font)
    except KeyError:
        pass
    return doc


def set_cell_bg(cell, hex_color="D9E2F3"):
    tcpr = cell._element.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcpr.append(shd)


def add_table(doc, header, rows, style="Table Grid", header_bg="D9E2F3"):
    """header: ['欄1','欄2']；rows: [[...],[...]]。標題列粗體加底色。"""
    t = doc.add_table(rows=1, cols=len(header))
    t.style = style
    for i, h in enumerate(header):
        cell = t.rows[0].cells[i]
        cell.text = ""
        set_run_font(cell.paragraphs[0].add_run(str(h)), bold=True)
        set_cell_bg(cell, header_bg)
    for r in rows:
        cells = t.add_row().cells
        for i, v in enumerate(r):
            cells[i].text = "" if v is None else str(v)
    return t
