---
name: make-skill
description: 把一個做法存成新技能，或修改現有技能，讓以後不用重教。Use when the user says 做成技能、存成技能、記住這個做法、以後都這樣做, or asks to create/update a skill.
---

# 建立 / 更新技能

## 技能放哪裡
放在**和本技能（make-skill）同一個 skills 資料夾**：`<skills 資料夾>/<技能名稱>/SKILL.md`
（OpenCode 是 `~/.config/opencode/skills`，Claude Code 是 `~/.claude/skills`，通用位置是 `~/.agents/skills`；Windows 的 `~` 是 `C:\Users\<使用者>`）

- 技能名稱：小寫英文 + 連字號，例如 `invoice-check`。資料夾名稱 = name。
- 檔名必須是全大寫 `SKILL.md`。
- 建立前先列出這個 skills 資料夾，已有類似技能就更新它，不要重複建立。
- **只能在這個 skills 資料夾建立或修改技能。**
  其他 agent 資料夾裡的技能，以及唯讀或有 `.git` 的技能，一律不可修改；
  想改進它們時，把建議告訴使用者，由使用者決定。
- 新技能的名稱不要和已存在的技能（包含上面那些資料夾）重複。

## SKILL.md 格式
```markdown
---
name: invoice-check
description: 一句話說明做什麼 + 什麼時候用（中英文關鍵字都放，例如：發票、invoice、核對）。
---

# 標題

## 步驟
1. ……（寫具體指令，例如 `python "<skills 資料夾>/invoice-check/scripts/check.py" "檔案"`）
2. ……

## 注意事項
- 之前踩過的坑、使用者特別要求的規則
```
description 最重要：模型是靠它決定要不要載入技能，要寫清楚「什麼情況下用」。

## 有程式就存成腳本
- 用到 Python 的做法，把**測試成功的程式**存到 `<技能資料夾>/scripts/xxx.py`，SKILL.md 裡寫怎麼執行。以後直接跑，不用每次重寫。
- 腳本要求：
  - 開頭 `import sys; sys.stdout.reconfigure(encoding="utf-8")`
  - 用 `argparse` 接收檔案路徑等參數，不要寫死路徑
  - 缺套件時自動 `pip install`
  - 長輸出寫到檔案，不要全部印出來

## 建立後
1. 用一個實際檔案測試腳本，確認能跑。
2. 告訴使用者技能名稱和存放位置，並提醒：**新技能通常要重新啟動 agent（OpenCode、Claude Code 等）才會出現**。
