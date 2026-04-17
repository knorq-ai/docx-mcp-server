# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.0.0] — 2026-04-17

### Changed (BREAKING)
- Renamed package from `docx-mcp-server` to `@llamadrive/docx-mcp-server`. Update your `.mcp.json` / install commands to the scoped name.
- Pinned `engines.node` to `>=18.0.0`.
- `track_changes` now defaults to `true` at the schema level on every editing tool (`replace_text`, `edit_paragraph`, `edit_paragraphs`, `insert_paragraph`, `insert_paragraphs`, `delete_paragraph`, `delete_paragraphs`, `edit_table_cell`, `edit_table_cells`). Previously the default lived only in the engine; the schema treated the field as optional, so an LLM passing `false` could silently slip through.
- Setting `track_changes: false` now requires also passing `allow_untracked_edit: true`. Without the second flag the call fails with `UNTRACKED_EDIT_NOT_ALLOWED`. This is a safety guard for regulated-industry use: prompt injection or long-context drift cannot ship silent edits unless two independent flags are set.

### Added
- `allow_untracked_edit` capability flag (default `false`) on all editing tools.
- GitHub Actions workflow that publishes to npm with `--provenance --access public` on tag push, signed via OIDC.

## [1.4.3] and earlier

See git history.
