"""auto_build 的回歸測試：每個「來源 × 範本」組合跑一次，統計自動檢查的問題數。

用法：
  python tests/run_tests.py                       測試組（tests/samples，沒有就先產生）
  python tests/run_tests.py 來源.pptx 範本.pptx    另外加測自己的檔案（可以多組，成對給）
  加 --png 會匯出圖片並做 PowerPoint 實際排版檢查（比較慢）
每次修改 auto_build.py / helpers.py 都要整組重跑，問題數不能變多。
"""
import os
import subprocess
import sys

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
AUTO = os.path.join(HERE, "..", "skills", "pptx-fill-company-template", "scripts", "auto_build.py")


def run(src, tpl, out_dir, png):
    name = os.path.splitext(os.path.basename(src))[0] + "__" + os.path.splitext(os.path.basename(tpl))[0]
    work = os.path.join(out_dir, name)
    os.makedirs(work, exist_ok=True)
    out = os.path.join(work, "out.pptx")
    cmd = [sys.executable, AUTO, src, tpl, out, "--work", work] + ([] if png else ["--no-png"])
    p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    text = p.stdout + p.stderr
    issues = [l[3:] for l in text.splitlines() if l.startswith("問題 ")]
    ok = p.returncode == 0 and "結果：" in text
    return name, ok, issues, text


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    png = "--png" in sys.argv
    samples = os.path.join(HERE, "samples")
    if not os.path.exists(os.path.join(samples, "src_a.pptx")):
        subprocess.run([sys.executable, os.path.join(HERE, "make_samples.py"), samples], check=True)
    cases = [(os.path.join(samples, s), os.path.join(samples, "tpl_plain.pptx")) for s in ("src_a.pptx", "src_b.pptx")]
    cases += list(zip(args[0::2], args[1::2]))
    extra_tpls = [t for _, t in zip(args[0::2], args[1::2])]
    for t in extra_tpls:  # 自己的範本也跑一次測試組
        cases += [(os.path.join(samples, s), t) for s in ("src_a.pptx", "src_b.pptx")]
    out_dir = os.path.join(samples, "_out")
    total = 0
    for src, tpl in cases:
        name, ok, issues, text = run(src, tpl, out_dir, png)
        total += len(issues) + (0 if ok else 1)
        print(f"{'OK ' if ok and not issues else 'NG '} {name}：{'執行失敗' if not ok else f'{len(issues)} 個問題'}")
        for i in issues:
            print("     - " + i)
        if not ok:
            print("\n".join("     " + l for l in text.splitlines()[-15:]))
    print(f"\n合計 {total} 個問題")
    sys.exit(1 if total else 0)


if __name__ == "__main__":
    main()
