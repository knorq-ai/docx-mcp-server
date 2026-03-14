# docx-mcp-server

[![CI](https://github.com/knorq-ai/docx-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/knorq-ai/docx-mcp-server/actions/workflows/ci.yml)

ローカル [MCP](https://modelcontextprotocol.io/) サーバ。Word (.docx) ファイルの読み取り・編集を行う。Claude Code、Cursor、その他の MCP 対応クライアントで動作する。

**32 ツール** — ドキュメント内容、書式設定、コメント、ページレイアウト、変更履歴を、ファイルアップロード不要の stdio トランスポートで処理する。

## 機能一覧

| カテゴリ | ツール |
|---------|--------|
| **読み取り** | `read_document`, `get_document_info`, `search_text`, `list_images` |
| **編集** | `replace_text`, `edit_paragraph`, `edit_paragraphs`, `insert_paragraph`, `insert_paragraphs`, `delete_paragraph`, `delete_paragraphs` |
| **書式** | `format_text`, `set_paragraph_format`, `set_paragraph_formats`, `highlight_text`, `set_heading`, `set_headings` |
| **構造** | `insert_table`, `create_document` |
| **レビュー** | `add_comment`, `add_comments`, `read_comments`, `reply_to_comment`, `delete_comment` |
| **変更履歴** | `accept_all_changes`, `reject_all_changes` |
| **ページレイアウト** | `get_page_layout`, `set_page_layout` |
| **ヘッダ/フッタ** | `read_header_footer` |
| **テーブル** | `edit_table_cell`, `edit_table_cells` |
| **脚注** | `read_footnotes` |

### 変更履歴 (Track Changes)

編集ツール (`replace_text`, `edit_paragraph`, `insert_paragraph`, `delete_paragraph`) は **変更履歴** に対応している。編集は Word のリビジョン (`w:ins`/`w:del`) として著者名・タイムスタンプ付きで記録され、Word 上で承認・却下ができる。

変更履歴は **デフォルトで有効** である。直接編集したい場合は `track_changes: false` を指定する。

`read_document` に `show_revisions: true` を渡すと、変更履歴が `[-削除-]` と `[+挿入+]` のアノテーション付きで表示される。デフォルトでは承認済みテキストのみ表示される。

`accept_all_changes` / `reject_all_changes` で全リビジョンを一括承認・却下できる。

### ページレイアウト

`get_page_layout` / `set_page_layout` は以下をサポートする:

- **用紙サイズプリセット**: A3, A4, A5, B4, B5, Letter, Legal
- **余白プリセット**: Normal, Narrow, Wide, JP Court 25mm, JP Court 30/20mm
- **カスタム値** (ミリメートル指定)
- **用紙の向き** (portrait / landscape)

## クイックスタート

### 方法 1: npm からインストール

```bash
npm install -g docx-mcp-server
```

MCP 設定に追加する（下記 [設定](#設定) を参照）。

### 方法 2: npx を使用（インストール不要）

設定を追加するだけで `npx` が自動的にダウンロード・実行する:

```json
{
  "mcpServers": {
    "docx-editor": {
      "command": "npx",
      "args": ["-y", "docx-mcp-server"]
    }
  }
}
```

### 方法 3: ソースからビルド

```bash
git clone https://github.com/knorq-ai/docx-mcp-server.git
cd docx-mcp-server
npm install
npm run build
npm link        # docx-mcp-server コマンドをグローバルに登録
```

## 設定

### Claude Code

プロジェクトの `.mcp.json` (プロジェクト単位) または `~/.claude/settings.json` (グローバル) に追加する:

```json
{
  "mcpServers": {
    "docx-editor": {
      "command": "npx",
      "args": ["-y", "docx-mcp-server"]
    }
  }
}
```

### Cursor

Cursor の MCP サーバ設定に追加する:

```json
{
  "mcpServers": {
    "docx-editor": {
      "command": "npx",
      "args": ["-y", "docx-mcp-server"]
    }
  }
}
```

### ローカルビルドを使用する場合

ソースからビルドして `npm link` を実行済みの場合:

```json
{
  "mcpServers": {
    "docx-editor": {
      "command": "docx-mcp-server"
    }
  }
}
```

## ツールリファレンス

### 読み取り

**`read_document`** — ブロックインデックス、スタイル、書式ヒント付きでドキュメント内容を読み取る。`show_revisions` で変更履歴を表示。
```
file_path, start_paragraph?, end_paragraph?, show_revisions?
```

**`get_document_info`** — 段落数、見出しアウトライン、テーブル数、コメント状態を取得。
```
file_path
```

**`search_text`** — コンテキストスニペット付きでテキスト検索。
```
file_path, query, case_sensitive?
```

**`list_images`** — 埋め込み画像の一覧（ファイル名、サイズ、代替テキスト、ブロックインデックス）。
```
file_path
```

### 編集

すべての編集ツールは `track_changes` (デフォルト `true`) と `author` (デフォルト `"Claude"`) を受け付ける。

**`replace_text`** — ドキュメント全体でテキストの検索・置換。複数ランにまたがるテキストにも対応。
```
file_path, search, replace, case_sensitive?, track_changes?, author?, include_headers_footers?
```

**`edit_paragraph`** — インデックス指定で段落テキストを置換。
```
file_path, paragraph_index, new_text, track_changes?, author?
```

**`edit_paragraphs`** — 複数段落を一括置換。1 回のファイル読み書きで処理。
```
file_path, edits (array of {paragraph_index, new_text}), track_changes?, author?
```

**`insert_paragraph`** — 指定位置に新しい段落を挿入。
```
file_path, text, position, style?, track_changes?, author?
```

**`insert_paragraphs`** — 複数段落を一括挿入。インデックスシフトを内部で処理。
```
file_path, items (array of {text, position, style?}), track_changes?, author?
```

**`delete_paragraph`** — インデックス指定で段落を削除。
```
file_path, paragraph_index, track_changes?, author?
```

**`delete_paragraphs`** — 複数段落を一括削除。インデックス再順序を内部で処理。
```
file_path, paragraph_indices, track_changes?, author?
```

### 書式設定

**`format_text`** — 太字、斜体、下線、フォント、サイズ、色、ハイライトをマッチするテキストに適用。
```
file_path, search, bold?, italic?, underline?, strikethrough?, highlight_color?, font_name?, font_size?, font_color?, case_sensitive?
```

**`set_paragraph_format`** — 段落の配置、間隔、インデントを設定。
```
file_path, paragraph_index, alignment?, space_before?, space_after?, line_spacing?, indent_left?, indent_right?, first_line_indent?, hanging_indent?
```

**`set_paragraph_formats`** — 複数段落の書式を一括設定。
```
file_path, groups (array of {indices, alignment?, space_before?, ...})
```

**`highlight_text`** — マッチするテキストにハイライトカラーを適用。
```
file_path, search, color?, case_sensitive?
```

**`set_heading`** — 段落を見出しに変換 (レベル 1-9)。
```
file_path, paragraph_index, level
```

**`set_headings`** — 複数段落を一括で見出しに変換。
```
file_path, items (array of {paragraph_index, level})
```

### 構造

**`insert_table`** — テーブルを挿入（オプションでセルデータ指定可）。
```
file_path, position, rows, cols, data?
```

**`create_document`** — 新しい .docx ファイルを作成（タイトル・本文はオプション）。
```
file_path, title?, content?
```

### レビュー

**`add_comment`** — 特定テキストにコメントをアンカー。
```
file_path, anchor_text, comment_text, author?
```

**`add_comments`** — 複数コメントを一括追加。部分的成功をサポート。
```
file_path, comments (array of {anchor_text, comment_text, author?}), default_author?
```

**`read_comments`** — 全コメントの一覧（ID、著者、テキスト、スレッド返信）。
```
file_path
```

**`reply_to_comment`** — 既存コメントに返信（スレッド会話を作成）。
```
file_path, parent_comment_id, comment_text, author?
```

**`delete_comment`** — ID 指定でコメントを削除。
```
file_path, comment_id
```

### 変更履歴

**`accept_all_changes`** — 全変更履歴を承認。挿入は確定、削除は除去。
```
file_path
```

**`reject_all_changes`** — 全変更履歴を却下。挿入は除去、削除テキストは復元。
```
file_path
```

### ページレイアウト

**`get_page_layout`** — 用紙サイズ、余白、向きを読み取り。
```
file_path
```

**`set_page_layout`** — プリセットまたはカスタム mm 値で用紙サイズ、余白、向きを設定。
```
file_path, page_size_preset?, orientation?, width_mm?, height_mm?, margin_preset?, top_mm?, right_mm?, bottom_mm?, left_mm?, header_mm?, footer_mm?, gutter_mm?
```

### ヘッダ/フッタ

**`read_header_footer`** — 全ヘッダ・フッタのテキスト内容を読み取り。
```
file_path
```

### テーブル

**`edit_table_cell`** — ブロック・行・列インデックス指定でテーブルセルのテキストを置換。
```
file_path, block_index, row_index, col_index, new_text, track_changes?, author?
```

**`edit_table_cells`** — 複数テーブルセルを一括編集。異なるテーブルにまたがることも可能。
```
file_path, edits (array of {block_index, row_index, col_index, new_text}), track_changes?, author?
```

### 脚注

**`read_footnotes`** — 全脚注の ID とテキスト内容を読み取り。
```
file_path
```

## 動作要件

- Node.js 20+

## ライセンス

MIT
