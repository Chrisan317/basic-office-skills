<#
basic-office-skills installer (Windows PowerShell 5.1+ / PowerShell 7)

One-line install from GitHub (OpenCode):
  irm https://raw.githubusercontent.com/Chrisan317/basic-office-skills/main/install.ps1 | iex

With options (e.g. Claude Code):
  & ([scriptblock]::Create((irm https://raw.githubusercontent.com/Chrisan317/basic-office-skills/main/install.ps1))) -Agent claude

From a downloaded copy of the repo (uses the local files):
  powershell -ExecutionPolicy Bypass -File install.ps1

Options:
  -Agent opencode|claude|agents   target agent (default: opencode)
  -ConfigDir <folder>             custom target (skills go to <folder>\skills, rules file to <folder>)
  -NoPip                          skip installing Python packages

This file is ASCII only on purpose, so it works both with "irm | iex" and "-File".
The installer itself (Chinese messages) is tools/install_core.ps1, loaded as UTF-8.
#>
param(
    [ValidateSet("opencode", "claude", "agents")]
    [string]$Agent = "opencode",
    [string]$ConfigDir = "",
    [switch]$NoPip
)

$ErrorActionPreference = "Stop"
$utf8 = New-Object System.Text.UTF8Encoding($false)
$repoRoot = ""
$core = $null
if ($PSCommandPath) {
    $here = Split-Path -Parent $PSCommandPath
    $local = Join-Path $here "tools\install_core.ps1"
    if (Test-Path $local) {
        $core = [IO.File]::ReadAllText($local, $utf8)
        $repoRoot = $here
    }
}
if (-not $core) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $url = "https://raw.githubusercontent.com/Chrisan317/basic-office-skills/main/tools/install_core.ps1"
    $core = $utf8.GetString((New-Object Net.WebClient).DownloadData($url))
}
$core = $core.TrimStart([char]0xFEFF)
& ([scriptblock]::Create($core)) -Agent $Agent -ConfigDir $ConfigDir -NoPip:$NoPip -RepoRoot $repoRoot
