#!/usr/bin/env bash
# Claude Code チーム設定インストーラー（Mac/Linux）
# 使い方: ./install.sh
# オプション: ./install.sh --force  （確認なしで上書き）

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "$0")" && pwd)"
CLAUDE_DIR="$HOME/.claude"
FORCE=false

for arg in "$@"; do
  [[ "$arg" == "--force" ]] && FORCE=true
done

step() { echo -e "\n\033[36m>> $1\033[0m"; }
ok()   { echo -e "   \033[32mOK: $1\033[0m"; }
warn() { echo -e "   \033[33mWARN: $1\033[0m"; }

echo "========================================"
echo " Claude Code チーム設定インストーラー"
echo "========================================"

# --- rules ---
step "rules をインストール中..."
mkdir -p "$CLAUDE_DIR/rules"
for f in "$REPO_ROOT/rules/"*.md; do
  cp "$f" "$CLAUDE_DIR/rules/"
  ok "$(basename "$f")"
done

# --- agents ---
step "agents をインストール中..."
mkdir -p "$CLAUDE_DIR/agents"
for f in "$REPO_ROOT/agents/"*.md; do
  cp "$f" "$CLAUDE_DIR/agents/"
  ok "$(basename "$f")"
done

# --- plugins ---
step "plugins をインストール中..."
mkdir -p "$CLAUDE_DIR/plugins"
for d in "$REPO_ROOT/plugins/"/*/; do
  name="$(basename "$d")"
  cp -r "$d" "$CLAUDE_DIR/plugins/$name"
  ok "$name"
done

# --- CLAUDE.md ---
step "CLAUDE.md を確認中..."
if [[ -f "$CLAUDE_DIR/CLAUDE.md" ]]; then
  if [[ "$FORCE" == true ]]; then
    cp "$REPO_ROOT/CLAUDE.md" "$CLAUDE_DIR/CLAUDE.md"
    ok "CLAUDE.md を上書きしました（--force）"
  else
    warn "CLAUDE.md はすでに存在します。個人設定を保護するためスキップします。"
    warn "上書きする場合は: ./install.sh --force"
  fi
else
  cp "$REPO_ROOT/CLAUDE.md" "$CLAUDE_DIR/CLAUDE.md"
  ok "CLAUDE.md を新規作成しました"
fi

echo ""
echo "========================================"
echo " インストール完了！"
echo " Claude Code を再起動して設定を反映してください。"
echo "========================================"
