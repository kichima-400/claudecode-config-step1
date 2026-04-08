# Claude Code チーム設定

チーム共通の Claude Code 設定（ルール・エージェント・スキル）を管理するリポジトリです。

## 含まれる設定

### ルール（常時適用）
| ファイル | 内容 |
|---|---|
| `rules/security.md` | セキュリティチェックリスト（OWASP準拠） |
| `rules/git-operations.md` | Git操作前の確認ルール |
| `rules/mcp.md` | MCP管理ルール（10個以下制限） |
| `rules/vba.md` | VBA固有の注意事項（AscWバグ等） |

### エージェント（呼び出し時に起動）
| ファイル | 内容 |
|---|---|
| `agents/python-reviewer.md` | Pythonコードレビュー専門 |
| `agents/vba-testing-guide.md` | VBAテスト戦略専門（Opusモデル） |

### プラグイン（文脈に応じて自動参照）
| プラグイン | 含まれるスキル |
|---|---|
| `plugins/ecc-python/` | python-patterns / python-testing |
| `plugins/vba-skills/` | vba-core-architecture / vba-error-events / vba-naming-and-data / vba-object-model-and-performance / vba-procedures-and-encapsulation |

### スキル（文脈に応じて自動参照）
| スキル | 内容 | 取得元 |
|---|---|---|
| `skills/powershell-expert/` | PowerShell スクリプト・GUI開発 | [hmohamed01/powershell-expert](https://github.com/hmohamed01/powershell-expert) |
| `skills/windows-shell/` | Windows パス処理・シェルコマンドガイドライン | [nicoforclaude/claude-windows-shell](https://github.com/nicoforclaude/claude-windows-shell) |
| `skills/xlsx/` | Excel/スプレッドシート操作（.xlsx/.xlsm/.csv/.tsv） | [anthropics/skills](https://github.com/anthropics/skills) |

## インストール方法

### Windows

```powershell
git clone https://github.com/yourorg/claude-team-config.git
cd claude-team-config
.\install.ps1
```

### Mac / Linux

```bash
git clone https://github.com/yourorg/claude-team-config.git
cd claude-team-config
chmod +x install.sh
./install.sh
```

インストール後、**Claude Code を再起動**してください。

## 設定の更新

チーム設定が更新された場合：

```powershell
# Windows
git pull
.\install.ps1

# Mac/Linux
git pull
./install.sh
```

## 個人設定について

- `CLAUDE.md` はすでに存在する場合スキップされます（個人設定を保護）
- 新規PCの初回セットアップ時のみ自動作成されます
- 個人的な追記は `~/.claude/CLAUDE.md` に直接行ってください
- 上書きしたい場合は `.\install.ps1 -Force`（Windows）または `./install.sh --force`（Mac/Linux）

## 設定を追加・変更する場合

1. このリポジトリを編集してPR・コミット
2. 各メンバーが `git pull` + インストーラー再実行

## ディレクトリ構成

```
claude-team-config/
├── README.md
├── install.ps1          # Windows用インストーラー
├── install.sh           # Mac/Linux用インストーラー
├── .gitignore
├── CLAUDE.md            # チーム共通の基本指示
├── rules/               # 常時適用ルール
├── agents/              # エージェント定義
├── plugins/             # プラグイン（スキルセット）
│   ├── ecc-python/
│   └── vba-skills/
└── skills/              # スキル（単体）
    ├── powershell-expert/
    ├── windows-shell/
    └── xlsx/
```
