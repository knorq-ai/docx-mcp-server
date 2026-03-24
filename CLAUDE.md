# CLAUDE.md — docx-mcp-server

ローカル MCP サーバ。DOCX ファイルの読み取り・編集・書式設定・コメント・画像一覧を提供する。

## ファイル構成

```
src/
  index.ts            … MCP サーバ本体（ツール登録・stdio transport）
  docx-engine.ts      … バレルモジュール（engine/* を再エクスポート + 公開 API 関数）
  engine/
    xml-helpers.ts    … XNode 型定義、fast-xml-parser の parser/builder、DOM ヘルパー
    docx-io.ts        … DOCX の ZIP 読み書き (JSZip)、ErrorCode/EngineError
    text.ts           … テキスト抽出、ブロック列挙、cross-run 置換、track changes
    formatting.ts     … 文字書式 (bold/italic 等)、段落書式 (alignment/spacing 等)
    comments.ts       … コメント XML 解析、アンカーマッチング、マーカー挿入
    layout.ts         … ページサイズ / マージンのプリセットと変換
    images.ts         … 画像一覧（リレーションシップ解析、w:drawing 走査）
    file-lock.ts      … ファイル単位の Promise チェーン書き込みロック
  __tests__/
    helpers.ts        … テストユーティリティ（tmp ファイル管理、フィクスチャ生成）
    docx-reading.test.ts
    docx-editing.test.ts
    docx-formatting.test.ts
    docx-comments.test.ts
    docx-structure-layout.test.ts
    docx-advanced-features.test.ts
    docx-bulk-operations.test.ts
    file-lock.test.ts
```

### モジュール依存グラフ（非循環）

```
xml-helpers ← docx-io ← text ← formatting
                    ↑
                    ├── comments
                    └── images
              layout (独立)
              file-lock (独立)
```

## ビルド・テスト

```bash
npm run build     # TypeScript → dist/
npx vitest run    # 全テスト実行
```

## ツール使用ワークフロー（推奨）

1. `get_document_info` でドキュメントの構造を把握する
2. `read_document` で対象範囲を読む（start_paragraph / end_paragraph で範囲指定可能）
3. `search_text` で編集対象のブロックを特定する
4. 編集系ツール（`edit_paragraph`, `replace_text` 等）で変更を行う

## デフォルト動作

| パラメータ | デフォルト値 | 備考 |
|---|---|---|
| `track_changes` | `true` | 変更履歴を w:del/w:ins として記録する |
| `author` | `"Claude"` | 変更履歴・コメントの著者名 |
| `case_sensitive` | `false` | 検索・置換時の大文字小文字区別 |

## パラメータ規約

- **ファイルパス**: すべて絶対パスで指定する
- **ブロックインデックス**: `read_document` / `get_document_info` の出力に表示されるゼロベースのインデックス
- **単位系**:
  - フォントサイズ: ポイント（pt）
  - インデント: twips（1440 twips = 1 inch）
  - ページサイズ / マージン: ミリメートル（mm）で指定、内部で twips に変換する
  - 画像サイズ: EMU（914400 EMU = 1 inch）

## 構造化レスポンス

`get_document_info`, `search_text`, `read_comments`, `list_images` はテキストの後に `<json>...</json>` ブロックで構造化データを返す。LLM はテキスト部分で自然言語応答を構成し、プログラムは JSON 部分をパースして利用できる。

## 書き込みロック

書き込み関数は `withFileLock` でラップされており、同一ファイルへの並行書き込みを自動直列化する。読み取り関数はロック不要。

## 番号付き段落の挿入

Word のリスト定義による自動番号付き見出し（例: 第1条、第2条…）を `insert_paragraph` / `insert_paragraphs` で再現するには 2 つの方法がある。

### 方法 A: `num_id` + `num_level` を明示指定

```
insert_paragraph(text="遡及適用", position=104, num_id=14, num_level=0)
```

`w:pPr` に `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="14"/></w:numPr>` が挿入される。`num_id` の値は既存段落の `read_document` 出力や document.xml から確認できる。`style` と併用可能。

### 方法 B: `copy_format_from` で既存段落の書式をコピー

```
insert_paragraph(text="遡及適用", position=104, copy_format_from=103)
```

指定ブロックインデックスの `w:pPr` を丸ごと deep-copy する。番号定義・インデント・行間・罫線等すべてが引き継がれる。`copy_format_from` 指定時は `style` / `num_id` / `num_level` は無視される。

## アンチパターン

- `read_document` で全体を読んでから書き換える → ブロックインデックスのずれが発生する。代わりに `search_text` で対象を特定してから最小範囲の編集を行う
- `track_changes: false` でサイレント編集 → 変更が追跡されず、レビューが困難になる。明示的な理由がない限りデフォルト（true）を使う
- 大量の段落書式を個別に `set_paragraph_format` で設定 → `set_paragraph_formats` でまとめて適用する
- 複数の段落を個別に `edit_paragraph` で編集 → `edit_paragraphs` でまとめて適用する（1 回のファイル読み書きで済む）
- 複数の段落を個別に `insert_paragraph` で挿入 → `insert_paragraphs` でまとめて挿入する（インデックスシフトも内部で処理される）
- 複数のセルを個別に `edit_table_cell` で編集 → `edit_table_cells` でまとめて適用する
- 複数の見出しを個別に `set_heading` で設定 → `set_headings` でまとめて設定する
