# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
