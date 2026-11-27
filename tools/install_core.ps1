<#
basic-office-skills 安裝的主要程式（由 install.ps1 以 UTF-8 載入執行，不要直接執行這個檔案）。
參數和說明見 install.ps1。
#>
param(
    [ValidateSet("opencode", "claude", "agents")]
    [string]$Agent = "opencode",
    [string]$ConfigDir = "",
    [switch]$NoPip,
    [string]$RepoRoot = ""
)

$ErrorActionPreference = "Stop"
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}
$Repo = "Chrisan317/basic-office-skills"
$Utf8 = New-Object System.Text.UTF8Encoding($false)

function Say($msg) { Write-Host $msg }

# ---------- 安裝位置 ----------
$HomeDir = [Environment]::GetFolderPath("UserProfile")
if (-not $ConfigDir) {
    switch ($Agent) {
        "opencode" { $ConfigDir = Join-Path $HomeDir ".config\opencode" }
        "claude"   { $ConfigDir = Join-Path $HomeDir ".claude" }
        "agents"   { $ConfigDir = Join-Path $HomeDir ".agents" }
    }
}
$RuleName = if ($Agent -eq "claude") { "CLAUDE.md" } else { "AGENTS.md" }
$SkillsDst = Join-Path $ConfigDir "skills"
$RuleDst = Join-Path $ConfigDir $RuleName

# ---------- 來源：在 repo 資料夾裡執行就用本機檔案，否則從 GitHub 下載 ----------
$SrcRoot = $null
$TempDir = $null
if ($RepoRoot -and (Test-Path (Join-Path $RepoRoot "skills")) -and (Test-Path (Join-Path $RepoRoot "AGENTS.md"))) { $SrcRoot = $RepoRoot }
if (-not $SrcRoot) {
    $TempDir = Join-Path ([IO.Path]::GetTempPath()) ("basic-office-skills-" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Force $TempDir | Out-Null
    $zip = Join-Path $TempDir "repo.zip"
    Say "下載 https://github.com/$Repo ..."
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -UseBasicParsing "https://github.com/$Repo/archive/refs/heads/main.zip" -OutFile $zip
    Expand-Archive -Path $zip -DestinationPath $TempDir -Force
    $SrcRoot = (Get-ChildItem $TempDir -Directory | Where-Object { Test-Path (Join-Path $_.FullName "skills") } | Select-Object -First 1).FullName
    if (-not $SrcRoot) { throw "下載的檔案裡找不到 skills 資料夾" }
}

try {
    # ---------- 1. 技能 ----------
    New-Item -ItemType Directory -Force $SkillsDst | Out-Null
    $installed = @()
    foreach ($d in Get-ChildItem (Join-Path $SrcRoot "skills") -Directory) {
        $target = Join-Path $SkillsDst $d.Name
        if (Test-Path (Join-Path $target ".git")) {
            Say "  略過 $($d.Name)：目標資料夾是 git repo（使用者自己維護），不覆蓋"
            continue
        }
        New-Item -ItemType Directory -Force $target | Out-Null
        Copy-Item (Join-Path $d.FullName "*") $target -Recurse -Force
        $installed += $d.Name
    }
    Say ("已安裝技能到 {0}：{1}" -f $SkillsDst, ($installed -join "、"))
    $others = Get-ChildItem $SkillsDst -Directory | Where-Object { $installed -notcontains $_.Name } | ForEach-Object { $_.Name }
    if ($others) { Say ("  原本就有的其他技能（保留不動）：{0}" -f ($others -join "、")) }

    # ---------- 2. 全域規則：用標記區塊合併 ----------
    $begin = "<!-- basic-office-skills:begin（這個區塊由 install.ps1 管理，重新安裝會更新；自己的規則請寫在區塊外） -->"
    $end = "<!-- basic-office-skills:end -->"
    $rules = [IO.File]::ReadAllText((Join-Path $SrcRoot "AGENTS.md"), $Utf8).Trim()
    $block = "$begin`r`n$rules`r`n$end"
    if (-not (Test-Path $RuleDst)) {
        [IO.File]::WriteAllText($RuleDst, $block + "`r`n", $Utf8)
        Say "已建立全域規則：$RuleDst"
    } else {
        $old = [IO.File]::ReadAllText($RuleDst, $Utf8)
        $pattern = "(?s)<!-- basic-office-skills:begin.*?<!-- basic-office-skills:end -->"
        if ($old -match $pattern) {
            $new = [regex]::Replace($old, $pattern, [Text.RegularExpressions.MatchEvaluator] { param($m) $block })
            $how = if ($new -eq $old) { "basic-office-skills 區塊已經是最新的" } else { "更新了既有的 basic-office-skills 區塊" }
        } elseif ($old.Contains($rules)) {
            $new = $old
            $how = "內容已經包含這份規則，沒有改動"
        } else {
            $new = $old.TrimEnd() + "`r`n`r`n" + $block + "`r`n"
            $how = "原本的內容保留，規則加在檔案最後"
        }
        if ($new -ne $old) {
            Copy-Item $RuleDst ($RuleDst + ".bak") -Force
            [IO.File]::WriteAllText($RuleDst, $new, $Utf8)
            Say "全域規則：$RuleDst（$how；原檔備份為 $RuleName.bak）"
        } else {
            Say "全域規則：$RuleDst（$how）"
        }
    }

    # ---------- 3. Python 套件 ----------
    if (-not $NoPip) {
        $py = $null
        foreach ($c in @("python", "py")) { if (Get-Command $c -ErrorAction SilentlyContinue) { $py = $c; break } }
        if ($py) {
            Say "安裝 Python 套件（已經裝過的會略過）..."
            & $py -m pip install -q pymupdf python-docx openpyxl python-pptx pandas tabulate xlrd pillow pywin32
            if ($LASTEXITCODE -ne 0) { Say "  有套件沒裝成功；不影響使用，腳本執行時缺套件會自動安裝" }
        } else {
            Say "找不到 Python：請先安裝 Python 3.10 以上（https://www.python.org/），技能腳本需要它"
        }
    }

    Say ""
    Say "完成。請重新啟動 $Agent，新技能才會出現在技能清單。"
}
finally {
    if ($TempDir -and (Test-Path $TempDir)) { Remove-Item $TempDir -Recurse -Force -ErrorAction SilentlyContinue }
}
