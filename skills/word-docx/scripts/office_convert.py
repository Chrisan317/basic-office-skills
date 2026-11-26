"""找 Office / LibreOffice、用它們轉檔。處理 32 位元與 64 位元 Office 的差異。

不要猜 Office 的安裝路徑：
  64 位元 Office → C:\\Program Files\\Microsoft Office\\root\\Office16\\
  32 位元 Office → C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\
本模組會從登錄檔（App Paths，64/32 位元兩邊都查）和兩個 Program Files 資料夾找。

用法（命令列）：
  python office_convert.py find                       列出找到的 Word / Excel / PowerPoint / LibreOffice
  python office_convert.py topdf 檔案.pptx|docx|xlsx [-o 輸出.pdf]
  python office_convert.py slides 簡報.pptx [-o 圖片資料夾]   每頁輸出成 PNG
在程式裡：
  from office_convert import to_pdf, export_slides, find_app
"""
import argparse
import glob
import os
import shutil
import subprocess
import sys
import tempfile

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

APPS = {"word": "WINWORD.EXE", "excel": "EXCEL.EXE", "powerpoint": "POWERPNT.EXE", "libreoffice": "soffice.exe"}
PROGIDS = {"word": "Word.Application", "excel": "Excel.Application", "powerpoint": "PowerPoint.Application"}


def find_app(exe):
    """回傳執行檔完整路徑，找不到回傳 None。exe 例如 'POWERPNT.EXE'、'soffice.exe'。"""
    cands = []
    if sys.platform == "win32":
        import winreg
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
                try:
                    k = winreg.OpenKey(hive, rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{exe}",
                                       0, winreg.KEY_READ | view)
                    v, _ = winreg.QueryValueEx(k, None)
                    cands.append(os.path.expandvars(v.strip('"')))
                except OSError:
                    pass
        bases = {os.environ.get(n) for n in ("ProgramW6432", "ProgramFiles", "ProgramFiles(x86)")}
        bases |= {r"C:\Program Files", r"C:\Program Files (x86)"}
        for b in filter(None, bases):
            cands += glob.glob(os.path.join(b, "Microsoft Office", "root", "Office*", exe))
            cands += glob.glob(os.path.join(b, "Microsoft Office", "Office*", exe))
            cands += glob.glob(os.path.join(b, "LibreOffice*", "program", exe))
    w = shutil.which(exe) or shutil.which(os.path.splitext(exe)[0])
    if w:
        cands.append(w)
    for c in cands:
        if c and os.path.isfile(c):
            return c
    return None


def describe():
    lines = []
    for name, exe in APPS.items():
        p = find_app(exe)
        bit = ""
        if p:
            bit = "（32 位元）" if "(x86)" in p else "（64 位元）" if sys.platform == "win32" and "Program Files" in p else ""
        lines.append(f"{name:12s}: {p or '找不到'} {bit}")
    return "\n".join(lines)


def _office_is_32bit():
    p = find_app("POWERPNT.EXE") or find_app("WINWORD.EXE") or find_app("EXCEL.EXE")
    return bool(p and "(x86)" in p)


def _powershells():
    """回傳要試的 PowerShell（64 位元、32 位元）。
    Office 是 32 位元時，64 位元 PowerShell 的 New-Object -ComObject 常常拿到不能用的物件
    （例如 $a.Presentations 是 null），所以先試 32 位元 PowerShell。"""
    root = os.environ.get("SystemRoot", r"C:\Windows")
    out = []
    for sub in ("Sysnative", "System32", "SysWOW64"):  # Sysnative：32 位元 Python 呼叫 64 位元 PowerShell
        p = os.path.join(root, sub, "WindowsPowerShell", "v1.0", "powershell.exe")
        if os.path.isfile(p) and p not in out:
            out.append(p)
    if _office_is_32bit():
        out.sort(key=lambda p: 0 if "SysWOW64" in p else 1)
    return out


def _ps_bits(ps):
    return "32" if "SysWOW64" in ps else "64"


def _ps_quote(s):
    return "'" + s.replace("'", "''") + "'"


def _run_ps(script):
    """依序用各個 PowerShell 執行，回傳成功的 powershell.exe 路徑；全部失敗時丟出各層的錯誤。"""
    errors = []
    # 錯誤訊息改成 UTF-8 輸出，中文 Windows 才不會變亂碼
    script = "[Console]::OutputEncoding=[Text.Encoding]::UTF8;" + script
    for ps in _powershells():
        r = subprocess.run([ps, "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
                           capture_output=True, text=True, encoding="utf-8", errors="replace")
        if r.returncode == 0:
            return ps
        msg = (r.stderr or r.stdout).strip().splitlines()
        errors.append(f"{_ps_bits(ps)} 位元 PowerShell: {msg[0][:200] if msg else r.returncode}")
    raise RuntimeError("\n".join(errors) or "找不到 PowerShell")


def _kind(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".ppt", ".pptx", ".pptm", ".potx"):
        return "powerpoint"
    if ext in (".doc", ".docx", ".docm", ".rtf", ".dotx"):
        return "word"
    if ext in (".xls", ".xlsx", ".xlsm"):
        return "excel"
    raise ValueError(f"不支援的檔案類型: {ext}")


def _com_pdf_pywin32(kind, src, dst):
    import win32com.client  # noqa: 需要 pywin32
    app = win32com.client.DispatchEx(PROGIDS[kind])
    try:
        if kind == "word":
            app.Visible = False
            d = app.Documents.Open(src, ReadOnly=True)
            d.SaveAs2(dst, FileFormat=17)
            d.Close(False)
        elif kind == "powerpoint":
            p = app.Presentations.Open(src, ReadOnly=True, WithWindow=False)
            p.SaveAs(dst, 32)
            p.Close()
        else:
            app.Visible = False
            app.DisplayAlerts = False
            wb = app.Workbooks.Open(src, ReadOnly=True)
            wb.ExportAsFixedFormat(0, dst)
            wb.Close(False)
    finally:
        app.Quit()


def _com_pdf_ps(kind, src, dst):
    s, d = _ps_quote(src), _ps_quote(dst)
    if kind == "word":
        body = (f"$a=New-Object -ComObject Word.Application;$a.Visible=$false;"
                f"$x=$a.Documents.Open({s},$false,$true);$x.SaveAs2({d},17);$x.Close($false);$a.Quit()")
    elif kind == "powerpoint":
        body = (f"$a=New-Object -ComObject PowerPoint.Application;"
                f"$x=$a.Presentations.Open({s},-1,0,0);$x.SaveAs({d},32);$x.Close();$a.Quit()")
    else:
        body = (f"$a=New-Object -ComObject Excel.Application;$a.Visible=$false;$a.DisplayAlerts=$false;"
                f"$x=$a.Workbooks.Open({s},0,$true);$x.ExportAsFixedFormat(0,{d});$x.Close($false);$a.Quit()")
    return _run_ps("$ErrorActionPreference='Stop';" + body)


def _soffice_pdf(src, dst):
    so = find_app("soffice.exe")
    if not so:
        raise RuntimeError("找不到 LibreOffice")
    outdir = tempfile.mkdtemp(prefix="lo_")
    subprocess.run([so, "--headless", "--convert-to", "pdf", "--outdir", outdir, src],
                   check=True, capture_output=True)
    produced = os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + ".pdf")
    shutil.move(produced, dst)


def to_pdf(src, dst=None):
    """轉 PDF。依序試：pywin32 COM → PowerShell COM（64 位元，再 32 位元）→ LibreOffice。回傳使用的方法。"""
    src = os.path.abspath(src)
    dst = os.path.abspath(dst or os.path.splitext(src)[0] + ".pdf")
    kind = _kind(src)
    errors = []
    if sys.platform == "win32":
        try:
            _com_pdf_pywin32(kind, src, dst)
            return "Office（pywin32）"
        except Exception as e:
            errors.append(f"pywin32: {e}")
        try:
            ps = _com_pdf_ps(kind, src, dst)
            return f"Office（{_ps_bits(ps)} 位元 PowerShell）"
        except Exception as e:
            errors.append(f"PowerShell: {e}")
    try:
        _soffice_pdf(src, dst)
        return "LibreOffice"
    except Exception as e:
        errors.append(f"LibreOffice: {e}")
    raise RuntimeError("轉 PDF 失敗：\n" + "\n".join(errors) +
                       "\n需要 Microsoft Office 或 LibreOffice（winget install TheDocumentFoundation.LibreOffice）")


def _slide_ratio(pptx):
    """投影片高/寬比例（16:9 = 0.5625、4:3 = 0.75）。"""
    try:
        from pptx import Presentation
        p = Presentation(pptx)
        return p.slide_height / p.slide_width
    except Exception:
        return 9 / 16


layout_issues = []  # 最近一次 pywin32 匯出時，PowerPoint 實際排版量到的問題


def _layout_check(p):
    """用 PowerPoint 實際排版後的大小檢查：超出投影片、文字超出自己的框。
    （python-pptx 只知道設定的列高；PowerPoint 會把放不下的表格列撐高，只有這裡量得到。）"""
    out = []
    W, H = p.PageSetup.SlideWidth, p.PageSetup.SlideHeight
    for i in range(1, p.Slides.Count + 1):
        for sh in p.Slides.Item(i).Shapes:
            try:
                name = sh.Name
                text = ""
                if sh.HasTextFrame and sh.TextFrame.HasText:
                    text = sh.TextFrame.TextRange.Text.replace("\r", " ").replace("\x0b", " ").strip()[:16]
                    name = f"「{text}」"
                elif sh.HasTable:
                    name = "表格"
                bottom, right = sh.Top + sh.Height, sh.Left + sh.Width
                if bottom > H + 3 or right > W + 3:
                    out.append(f"第 {i} 頁：{name} 超出投影片（下緣 {bottom / 72:.2f} 吋／投影片高 {H / 72:.2f} 吋）")
                elif text:
                    tr = sh.TextFrame.TextRange
                    tb = tr.BoundTop + tr.BoundHeight
                    if tb > H + 3:
                        out.append(f"第 {i} 頁：{name} 的文字超出投影片底部")
                    elif tb > bottom + 6 and sh.TextFrame.AutoSize == 0:
                        out.append(f"第 {i} 頁：{name} 的文字超出自己的框 {(tb - bottom) / 72:.2f} 吋（可能蓋到下面的東西）")
            except Exception:
                continue
    return out


def _slides_pywin32(pptx, outdir, width, height):
    import win32com.client  # noqa: 需要 pywin32
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        p = app.Presentations.Open(pptx, ReadOnly=True, WithWindow=False)
        for i in range(1, p.Slides.Count + 1):
            p.Slides.Item(i).Export(os.path.join(outdir, f"slide_{i:03d}.png"), "PNG", width, height)
        try:
            layout_issues[:] = _layout_check(p)
        except Exception:
            layout_issues[:] = []
        p.Close()
    finally:
        app.Quit()


def export_slides(pptx, outdir=None, width=1600, report=None):
    """每頁簡報輸出成 PNG，回傳圖片路徑清單。
    依序試：pywin32 COM → PowerShell COM（依 Office 位元數決定先試 32 或 64 位元）→ 轉 PDF 再用 PyMuPDF 轉圖。
    report 給一個 list 時，會把使用的方法和失敗的層記進去。"""
    pptx = os.path.abspath(pptx)
    outdir = os.path.abspath(outdir or os.path.splitext(pptx)[0] + "_頁面圖")
    os.makedirs(outdir, exist_ok=True)
    report = report if report is not None else []

    def done():
        return sorted(glob.glob(os.path.join(outdir, "slide_*.png")))

    def clear():
        for f in done():
            os.remove(f)
    clear()
    layout_issues[:] = []
    height = int(round(width * _slide_ratio(pptx)))
    if sys.platform == "win32":
        try:
            _slides_pywin32(pptx, outdir, width, height)
            if done():
                report.append("PowerPoint（pywin32）")
                return done()
        except Exception as e:
            report.append(f"失敗 pywin32: {str(e)[:200]}")
            clear()
        script = ("$ErrorActionPreference='Stop';$a=New-Object -ComObject PowerPoint.Application;"
                  f"$x=$a.Presentations.Open({_ps_quote(pptx)},-1,0,0);"
                  "for($i=1;$i -le $x.Slides.Count;$i++){"
                  f"$x.Slides.Item($i).Export((Join-Path {_ps_quote(outdir)} ('slide_{{0:D3}}.png' -f $i)),'PNG',{width},{height})}};"
                  "$x.Close();$a.Quit()")
        try:
            ps = _run_ps(script)
            if done():
                report.append(f"PowerPoint（{_ps_bits(ps)} 位元 PowerShell）")
                return done()
        except Exception as e:
            report.append(f"失敗 PowerShell: {e}")
            clear()
    pdf = os.path.join(tempfile.mkdtemp(prefix="slides_"), "deck.pdf")
    how = to_pdf(pptx, pdf)
    try:
        import pymupdf
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "pymupdf"])
        import pymupdf
    files = []
    with pymupdf.open(pdf) as doc:
        for i, page in enumerate(doc, 1):
            p = os.path.join(outdir, f"slide_{i:03d}.png")
            page.get_pixmap(dpi=110).save(p)
            files.append(p)
    report.append(f"{how} 轉 PDF → PyMuPDF")
    return files


def main():
    ap = argparse.ArgumentParser(description="Office 尋找與轉檔")
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("find")
    s = sub.add_parser("topdf"); s.add_argument("file"); s.add_argument("-o", "--output")
    s = sub.add_parser("slides"); s.add_argument("file"); s.add_argument("-o", "--output")
    a = ap.parse_args()
    if a.cmd == "find":
        print(describe())
    elif a.cmd == "topdf":
        dst = a.output or os.path.splitext(a.file)[0] + ".pdf"
        how = to_pdf(a.file, dst)
        print(f"已用 {how} 轉成: {os.path.abspath(dst)}")
    else:
        rep = []
        files = export_slides(a.file, a.output, report=rep)
        print(f"已用 {rep[-1]} 輸出 {len(files)} 張: {os.path.dirname(files[0]) if files else ''}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        sys.exit(str(e))
