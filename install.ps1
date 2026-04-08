# Claude Code チーム設定インストーラー（Windows）
# 使い方: .\install.ps1
# オプション: .\install.ps1 -Force  （確認なしで上書き）

param(
    [switch]$Force
)

$RepoRoot  = $PSScriptRoot
$ClaudeDir = "$env:USERPROFILE\.claude"

function Write-Step { param($msg) Write-Host "`n>> $msg" -ForegroundColor Cyan }
function Write-Ok   { param($msg) Write-Host "   OK: $msg" -ForegroundColor Green }
function Write-Warn { param($msg) Write-Host "   WARN: $msg" -ForegroundColor Yellow }
function Write-Fail { param($msg) Write-Host "   ERROR: $msg" -ForegroundColor Red }

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " Claude Code チーム設定インストーラー" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# --- rules ---
Write-Step "rules をインストール中..."
$rulesDir = "$ClaudeDir\rules"
New-Item -ItemType Directory -Force -Path $rulesDir | Out-Null
Get-ChildItem "$RepoRoot\rules\*.md" | ForEach-Object {
    Copy-Item $_.FullName -Destination $rulesDir -Force
    Write-Ok $_.Name
}

# --- agents ---
Write-Step "agents をインストール中..."
$agentsDir = "$ClaudeDir\agents"
New-Item -ItemType Directory -Force -Path $agentsDir | Out-Null
Get-ChildItem "$RepoRoot\agents\*.md" | ForEach-Object {
    Copy-Item $_.FullName -Destination $agentsDir -Force
    Write-Ok $_.Name
}

# --- plugins ---
Write-Step "plugins をインストール中..."
$pluginsDir = "$ClaudeDir\plugins"
New-Item -ItemType Directory -Force -Path $pluginsDir | Out-Null
Get-ChildItem "$RepoRoot\plugins" -Directory | ForEach-Object {
    $dest = "$pluginsDir\$($_.Name)"
    Copy-Item $_.FullName -Destination $dest -Recurse -Force
    Write-Ok $_.Name
}

# --- CLAUDE.md ---
Write-Step "CLAUDE.md を確認中..."
$claudeMd = "$ClaudeDir\CLAUDE.md"
if (Test-Path $claudeMd) {
    if ($Force) {
        Copy-Item "$RepoRoot\CLAUDE.md" -Destination $claudeMd -Force
        Write-Ok "CLAUDE.md を上書きしました（-Force）"
    } else {
        Write-Warn "CLAUDE.md はすでに存在します。個人設定を保護するためスキップします。"
        Write-Warn "上書きする場合は: .\install.ps1 -Force"
    }
} else {
    Copy-Item "$RepoRoot\CLAUDE.md" -Destination $claudeMd
    Write-Ok "CLAUDE.md を新規作成しました"
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host " インストール完了！" -ForegroundColor Green
Write-Host " Claude Code を再起動して設定を反映してください。" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
