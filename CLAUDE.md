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
    anchors.ts        … 安定アンカー (w14:paraId) の生成・収集・名前空間付与・解決
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
                    ├── images
                    └── anchors
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
3. `search_text` で編集対象のブロックを特定する（テーブル内のマッチは row/col 付きで返る）
4. ピンポイント確認には軽量ツールを使う: `read_table_structure` / `read_table_cell`（全体を読まずにテーブルを調べる）、`get_paragraph_format`（段落書式を調べる。`copy_format_from` の参照元探しに有用）
5. 編集系ツール（`edit_paragraphs`, `replace_texts` 等）で変更を行う

複数ステップで挿入・削除を伴う編集では、まず `ensure_anchors` を 1 回呼んでアンカー（`w14:paraId`）を割り当て、以降は `paragraph_index` の代わりに `anchor` で対象を指定する。アンカーはインデックスのずれに影響されないので、編集のたびに読み直す必要がなくなる（後述「安定アンカー」）。

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

`get_document_info`, `search_text`, `read_comments`, `list_images`, `read_table_structure`, `read_table_cell`, `get_paragraph_format` はテキストの後に `<json>...</json>` ブロックで構造化データを返す。LLM はテキスト部分で自然言語応答を構成し、プログラムは JSON 部分をパースして利用できる。

## 書き込みロック

書き込み関数は `withFileLock` でラップされており、同一ファイルへの並行書き込みを自動直列化する。読み取り関数はロック不要。

## 改行（`\n`）の扱い

`edit_paragraphs` / `edit_table_cells` / `insert_paragraphs` のテキスト中の `\n` は**段落区切り**として扱う。

- **変更履歴オフ（`track_changes: false`）**: 行ごとに別々の `<w:p>` へ分割する。各段落は元段落の `pPr`（番号付け・ぶら下げインデント・配置）と先頭 run の `rPr` を継承するため、手書きの番号付きリストや複数行セルが正しく描画される。
  - `edit_table_cells` はセル**全体**を置換する（先頭段落だけでなく全 `<w:p>` を入れ替える）。再編集しても古い行は残らない。`w:tcPr`・ネストテーブルは保持する。
  - 複数行編集は段落数を増やすため、以降のブロックインデックスがずれる。編集後は `search_text` で対象を取り直す。
- **変更履歴オン（デフォルト）**: 単一段落のまま `\n` をソフトブレーク（`<w:br/>`）として描画する。追跡された段落マーク挿入は accept/reject で正しく往復しないため、段落分割は行わない。

## 番号付き段落の挿入

Word のリスト定義による自動番号付き見出し（例: 第1条、第2条…）を `insert_paragraphs` で再現するには 2 つの方法がある。

### 方法 A: `num_id` + `num_level` を明示指定

```
insert_paragraphs(paragraphs=[{text: "遡及適用", position: 104, num_id: 14, num_level: 0}])
```

`w:pPr` に `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="14"/></w:numPr>` が挿入される。`num_id` の値は既存段落の `read_document` 出力や document.xml から確認できる。`style` と併用可能。

### 方法 B: `copy_format_from` で既存段落の書式をコピー

```
insert_paragraphs(paragraphs=[{text: "遡及適用", position: 104, copy_format_from: 103}])
```

指定ブロックインデックスの `w:pPr` を丸ごと deep-copy する。番号定義・インデント・行間・罫線等すべてが引き継がれる。`copy_format_from` 指定時は `style` / `num_id` / `num_level` は無視される。

<<<<<<< HEAD
## テーブルセル内の段落編集

`edit_table_cells` はセル全体を対象にする。複数段落セル（番号付きリスト等）の特定の 1 段落だけを操作するには、セル内ローカルの段落インデックス（セル内の `w:p` のみを 0 始まりで数える）で指定する専用ツールを使う。

- `edit_table_paragraphs(block, row, col, paragraph_index, new_text)` … セル内 1 段落だけ置換
- `delete_table_paragraphs(block, row, col, paragraph_index)` … セル内 1 段落だけ削除（最後の 1 段落を消すと空段落を残す。`w:numPr` の自動番号は Word が振り直す）
- `insert_table_paragraphs(block, row, col, position, text, copy_format_from?)` … セル内ローカル位置に挿入（`position: -1` で末尾追加。`copy_format_from` は同一セル内の段落インデックス）

row/col は物理的な `w:tc` 位置（`edit_table_cells` と同じ規約）。セル内の段落構成は `read_document` の `[TABLE]` 出力で確認できる。
=======
## 安定アンカー

段落はブロックインデックスで指定するが、挿入・削除のたびに以降の番号がずれる。**アンカー**（Word の `w14:paraId`）は段落に貼り付いたまま動く安定 ID で、ずれの影響を受けない。

- `ensure_anchors(file_path)` … 全トップレベル段落にアンカーを割り当て、インデックス→アンカー対応表を返す。冪等。欠落 / 不正値 / 重複の ID は再採番し、`w14`/`mc` 名前空間を宣言する。
- `read_document(show_anchors=true)` でインライン表示、`search_text` は各マッチの `anchor` を返す。
- 編集系ツールは `paragraph_index` の代わりに `anchor`（`set_paragraph_formats` / `delete_paragraphs` は `anchors`）で指定できる。`insert_paragraphs` は `anchor` + `placement: before|after` で配置し、新規段落のアンカーを返す。
- 編集・挿入で触れた段落には自動でアンカーが付与される（触れていない段落には付与しない）。
- v1 のスコープは body 直下の段落のみ。テーブルセル内・`w:sdt` 内の段落は対象外（テーブルブロックの `anchor` は `null`）。
>>>>>>> c58cdbb (Add stable paragraph anchors (w14:paraId) for index-independent editing (#7))

## アンチパターン

- `read_document` で全体を読んでから書き換える → ブロックインデックスのずれが発生する。代わりに `search_text` で対象を特定してから最小範囲の編集を行う（連続編集では `ensure_anchors` + `anchor` 指定が確実）
- `track_changes: false` でサイレント編集 → 変更が追跡されず、レビューが困難になる。明示的な理由がない限りデフォルト（true）を使う
- バルクツールを 1 件ずつ繰り返し呼び出す → 1 回のコールに集約する。バルクツール (`replace_texts`, `edit_paragraphs`, `insert_paragraphs`, `delete_paragraphs`, `set_paragraph_formats`, `set_headings`, `edit_table_cells`) は単一アイテム配列でも動作し、1 回のファイル読み書きで複数件を処理できる
