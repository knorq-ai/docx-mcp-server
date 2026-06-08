# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [3.2.0] — 2026-06-08

Consolidates four feature tracks (issues #4, #5, #6, #7) plus a multi-agent QA
hardening pass into a single release. Adds 7 tools (**34 total**). All changes
are additive or bug fixes; documents without top-level content controls (`w:sdt`)
are unaffected by the block-index change noted under Changed.

### Added
- **Per-paragraph table-cell editing (#5):** `edit_table_paragraphs`, `delete_table_paragraphs`, `insert_table_paragraphs` — address an individual `<w:p>` inside a `<w:tc>` by a cell-local paragraph index (row/column are physical `w:tc` positions). Deleting a cell's last paragraph keeps a blank one so the cell stays valid; `insert_table_paragraphs` supports `num_id`/`num_level` and same-cell `copy_format_from`. All three honour tracked changes (default) + the `allow_untracked_edit` flag and refuse (tracked) paragraphs with pending revisions.
- **Table / paragraph introspection (#6):** `read_table_structure` (row/col dimensions, per-cell preview, `gridSpan`/`vMerge`), `read_table_cell` (one cell's paragraphs + merge info), and `get_paragraph_format` (style, heading level, alignment, numbering, indentation in twips, spacing in points — same units `set_paragraph_formats` accepts).
- **Stable paragraph anchors (#7):** `ensure_anchors` (`w14:paraId`) assigns a valid, unique anchor to every top-level paragraph lacking one (repairing missing/invalid/duplicate ids, declaring the `w14`/`mc` namespaces) and returns the index→anchor map (idempotent). `read_document` gains `show_anchors`; `search_text` reports each match's `anchor`; the paragraph edit tools accept an `anchor` (or `anchors`) as an alternative to the integer index, and `insert_paragraphs` accepts `anchor` + `placement: before|after` / `copy_format_from_anchor` and returns the new anchors. Editing auto-seeds anchors on touched paragraphs. v1 scope: direct-body paragraphs only (table / `w:sdt` paragraphs report a `null` anchor).
- New error codes `AMBIGUOUS_ANCHOR` and `INVALID_LOCATOR`.

### Changed
- **`\n` is a paragraph break (#4)** across `edit_paragraphs`, `edit_table_cells`, `insert_paragraphs`: **untracked** edits/inserts split on `\n` into separate `<w:p>` (each inheriting the source `pPr` — numbering, hanging indent, alignment — and first-run `rPr`); **tracked** edits/inserts render `\n` as a `<w:br/>` soft break (a tracked paragraph-mark insertion does not round-trip through `accept_all_changes` / `reject_all_changes`). `replace_texts` and untracked `edit_table_paragraphs` likewise render `\n` as a soft break, and CRLF / lone CR is normalized to `\n`. An untracked multi-line edit grows the block count — re-read or use anchors / `search_text` to re-locate targets.
- **`edit_table_cells` (untracked) replaces the whole cell** rather than only its first paragraph, preserving `w:tcPr` and nested block content; re-editing a multi-paragraph cell leaves no residue.
- **Unified block-index space (#6/#7):** `read_document`, `search_text`, `get_document_info`, and every index-consuming edit/table/format/anchor tool now share ONE block numbering (descending into a top-level `w:sdt`). For a document **with** a top-level content control this changes the indices those tools accept so they finally agree with what `read_document`/`search_text` report — previously they could silently target the wrong block. Documents without content controls are unaffected. Inserting/deleting a paragraph whose index falls inside a `w:sdt` is refused with `INVALID_LOCATOR` (edit it in place by index/anchor).

### Fixed
- **Illegal XML control characters** (U+0000–08, 0B, 0C, 0E–1F) in text are stripped instead of silently producing a corrupt `.docx` that Word / strict readers reject.
- **Empty search queries** no longer hang the (single-threaded) server or spin to a `RangeError` — `search_text` / `format_text` / `highlight_text` reject empty input.
- **`insert_table`** now emits the required `<w:tblGrid>`, and a table cell never ends in a nested `<w:tbl>` (a trailing paragraph is kept) — both previously failed strict OOXML validators such as python-docx.
- **`reject_all_changes`** removes a rejected tracked-inserted paragraph entirely (no residual empty paragraph), keeping a blank one only where a cell/body would otherwise be left with none.
- **Structured `<json>` responses** are delimiter-safe (document text containing `</json>` no longer breaks extraction); zero-result reads still emit a `<json>` block; `get_paragraph_format` JSON includes explicit `alignment` / `style` / `headingLevel` defaults.
- **Index validation:** non-integer / wrong-type indices are rejected with `INDEX_OUT_OF_RANGE` everywhere (no raw `TypeError`; no silent string-coercion that mutated the wrong cell).
- **`copy_format_from`** (in `insert_paragraphs` and the table-paragraph tools) no longer carries the source paragraph's stale tracked-change metadata (`w:pPrChange` / `pPr > rPr > w:rPrChange`) onto the new paragraph.
- `ensure_anchors` block indices now match `read_document` / `search_text`.

## [3.1.0] — 2026-05-03

### Added
- New `PENDING_REVISIONS` error code. Tracked-mode editing tools (`replace_texts`, `edit_paragraphs`, `delete_paragraphs`, `edit_table_cells`) now refuse to operate on a paragraph or table cell that already contains tracked-change markup. The guard detects:
  - run-level `w:ins` / `w:del`
  - move-tracking `w:moveFrom` / `w:moveTo`
  - paragraph-mark revisions under `pPr > rPr`
  - revisions nested inside inline `w:sdt > w:sdtContent` (Google Docs export pattern)
  - existing revisions in header/footer paragraphs when `include_headers_footers: true`
  Previously the matcher walked into existing tracked wrappers as if they were normal text, producing nested or overlapping revision markup that did not round-trip through `accept_all_changes` / `reject_all_changes`. Resolution: call `accept_all_changes` or `reject_all_changes` first, or pass `track_changes: false` (with `allow_untracked_edit: true`).

### Fixed
- Tracked-change `w:id` allocation now scans every DOCX part (`word/document.xml` plus all `header*.xml`, `footer*.xml`, `footnotes.xml`, `endnotes.xml`) before seeding new revision IDs. The scan accepts both single-quoted and double-quoted attribute values and tolerates whitespace around `=`. Previously only the body was scanned with a strict double-quoted regex, so existing revisions in header/footer parts (or any external tool that emits single-quoted XML) could collide with newly minted revision IDs. Affects `replace_texts`, `edit_paragraphs`, `insert_paragraphs`, `delete_paragraphs`, and `edit_table_cells`.

## [3.0.0] — 2026-05-03

### Changed (BREAKING)
- Removed the following single-item MCP tools and their underlying engine functions; use the bulk equivalent in every case:
  - `replace_text` → `replace_texts({items: [{search, replace}]})`
  - `edit_paragraph` → `edit_paragraphs({edits: [{paragraph_index, new_text}]})`
  - `insert_paragraph` → `insert_paragraphs({paragraphs: [{text, position, ...}]})`
  - `delete_paragraph` → `delete_paragraphs({paragraph_indices: [idx]})`
  - `set_heading` → `set_headings({headings: [{paragraph_index, level}]})`
  - `set_paragraph_format` → `set_paragraph_formats({groups: [{indices: [idx], alignment?, space_before?, ...}]})`
  - `edit_table_cell` → `edit_table_cells({edits: [{block_index, row_index, col_index, new_text}]})`
- Rationale: every MCP tool's schema is loaded into the LLM context window on every turn. The duplicate single+bulk tools doubled the schema-token cost with no gain in capability — bulk tools handle the single-item case identically. `add_comment` / `add_comments` was deliberately kept as a pair because the singular form throws on missing anchors while the bulk form returns per-item failures, which is a meaningful behavioral difference.

### Added
- `replace_texts` tool — apply one or more find/replace operations in a single open/save cycle. Per-item `case_sensitive` flag.
  - Under `track_changes: false`, items are applied sequentially: a later item can match text produced by an earlier item (e.g. `alpha→beta` then `beta→gamma` yields `gamma`).
  - Under `track_changes: true` (default), the engine rejects overlapping items where item N's `search` shares text with any earlier item M's `replace` (in either direction). Reason: tracked sequential replacement cannot safely chain overlapping items — the resulting nested `w:ins`/`w:del` markup does not round-trip through `reject_all_changes`. Workaround: issue separate `replace_texts` calls (one per item) or use `track_changes: false` with `allow_untracked_edit: true`.
- Engine-level guard rejecting empty `search` strings (would otherwise loop forever on the existing `replaceInParagraph` matcher).

## [2.0.0] — 2026-04-17

### Changed (BREAKING)
- Renamed package from `docx-mcp-server` to `@knorq/docx-mcp-server`. Update your `.mcp.json` / install commands to the scoped name.
- Pinned `engines.node` to `>=18.0.0`.
- `track_changes` now defaults to `true` at the schema level on every editing tool (`replace_text`, `edit_paragraph`, `edit_paragraphs`, `insert_paragraph`, `insert_paragraphs`, `delete_paragraph`, `delete_paragraphs`, `edit_table_cell`, `edit_table_cells`). Previously the default lived only in the engine; the schema treated the field as optional, so an LLM passing `false` could silently slip through.
- Setting `track_changes: false` now requires also passing `allow_untracked_edit: true`. Without the second flag the call fails with `UNTRACKED_EDIT_NOT_ALLOWED`. This is a safety guard for regulated-industry use: prompt injection or long-context drift cannot ship silent edits unless two independent flags are set.

### Added
- `allow_untracked_edit` capability flag (default `false`) on all editing tools.
- GitHub Actions workflow that publishes to npm with `--provenance --access public` on tag push, signed via OIDC.

## [1.4.3] and earlier

See git history.
