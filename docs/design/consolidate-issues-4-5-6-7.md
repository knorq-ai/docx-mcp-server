# Consolidation + QA hardening: issues #4 / #5 / #6 / #7

**Status:** in progress · **Target version:** 3.2.0 · **Date:** 2026-06-08

Single PR that resolves the four open issues, consolidated from the existing
stacked PR chain and hardened by a multi-agent QA pass (dogfooding + monkey
testing) before review.

## Background

Four open issues each already have a dedicated PR, built as a **linear stack**:

| Issue | PR  | Branch                           | Stacked on |
|-------|-----|----------------------------------|------------|
| #4    | #8  | `fix/4-newline-paragraph-breaks` | `main`     |
| #6    | #9  | `feat/6-read-introspection-tools`| #8         |
| #5    | #10 | `feat/5-table-paragraph-editing` | #9         |
| #7    | #11 | `feat/7-stable-anchors`          | #10        |

`feat/7-stable-anchors` (the stack top) therefore already contains all four
commits conflict-free. Consolidation is a fast-forward of that branch; the real
work is **QA**: find and fix bugs / UX gaps the authors' own tests would not
catch, then reconcile version + docs into one coherent release.

## Scope

In scope (full): **#4**, **#5**, **#7**, and #6's high-priority items
(6.1, 6.2, 6.3, 6.5).

Deferred from #6 (spun out as follow-up issues, per maintainer ym259's
prioritization — 6.4 medium, 6.7 low):
- **6.4** session/file-level "untracked edit" mode toggle — needs design for a
  stateless server.
- **6.7** plural-form tool-naming polish.

## Acceptance criteria (what "resolved" means)

### #4 — `\n` as paragraph break
- `\n` is a paragraph break across `edit_paragraphs`, `edit_table_cells`,
  `insert_paragraphs` (and any tool taking multi-line text).
- **Untracked:** split into N `<w:p>` sharing the source `pPr` + first-run
  `rPr` (numbering, hanging indent, alignment inherited).
- **Tracked:** single paragraph, `\n` → `<w:br/>` soft break (no tracked
  paragraph-mark insertion — accept/reject never merges paragraphs here).
- Read/write symmetry: text shown by `read_document` as `a\nb` round-trips.
- Accept/reject of tracked multi-line edits restores correctly (deleted text
  spanning a soft break emitted as `<w:br/>` inside `<w:del>`).
- Untracked multi-line edit grows block count → index shift documented.

### #5 — table cell editing
- **Issue 1:** `edit_table_cells` (untracked) fully replaces the cell —
  clears existing `<w:p>` while preserving `<w:tcPr>` and non-paragraph block
  content (nested `<w:tbl>`, `<w:sdt>`, bookmarks). *QA must confirm tracked
  mode does not silently append.*
- **Issue 2:** per-paragraph cell ops — `edit_table_paragraphs`,
  `delete_table_paragraphs`, `insert_table_paragraphs`; cell-local 0-based
  paragraph index; physical `w:tc` row/col.
- Deleting the last paragraph keeps a blank `<w:p>` (a `<w:tc>` needs ≥1 block
  child).
- Merged-cell awareness: `gridSpan`/`vMerge` surfaced; indexing = physical grid
  position (documented).
- `copy_format_from` within the same cell; `num_id`/`num_level` supported.
- Honors `track_changes` (default true) + `allow_untracked_edit`; refuses
  tracked edits on paragraphs with pending revisions; validates indices
  (rejects non-integers).

### #6 — introspection (high-priority subset)
- 6.1 `insert_paragraphs` same-`position` reverse-order semantics documented.
- 6.2 `read_table_structure`, `read_table_cell`.
- 6.3 `search_text` reports `rowIndex`/`colIndex` for table matches.
- 6.5 `get_paragraph_format` (style, alignment, numbering, indent twips,
  spacing points — same units `set_paragraph_formats` accepts).

### #7 — stable anchors
- `ensure_anchors`: assign valid unique `w14:paraId` to every top-level
  paragraph lacking one; repair missing/invalid/duplicate; declare `w14`/`mc`
  namespaces; return index→anchor map; idempotent.
- `read_document` gains `show_anchors`; `search_text` reports `anchor`.
- Edit tools accept `anchor` as an alternative to index (`edit_paragraphs`,
  `set_headings`, `set_paragraph_formats`, `delete_paragraphs`,
  `insert_paragraphs` with `placement: before|after` + `copy_format_from_anchor`).
- Auto-seed anchors on touched/created paragraphs only.
- `w14:paraId` in `00000001`–`7FFFFFFF`, part-wide unique.
- Error paths: `AMBIGUOUS_ANCHOR`, `INVALID_LOCATOR`, `ANCHOR_NOT_FOUND`;
  refuse (`INVALID_DOCX`) a doc binding `w14`/`mc` to a different URI.
- v1 scope: direct-body paragraphs only; table/`sdt` paragraphs report `null`.

## Consolidation mechanics

- Branch `feat/consolidate-issues-4-5-6-7` off `origin/feat/7-stable-anchors`
  in an isolated worktree (main stays clean).
- Preserve the 4 logical commits; add QA-fix commits; finish with one reconcile
  commit: single version bump `3.1.0 → 3.2.0` (all-additive + bugfixes, no
  breaking changes), one CHANGELOG entry, README / README.ja / CLAUDE.md, tool
  count (33 + 3 + 3 + 1 = **40**).
  - **Correction (M8):** the planning arithmetic above was off — the real
    pre-consolidation base is **27** tools (not 33), so the shipped total is
    **34** (= 27 + the consolidation's net additions), matching
    README / CHANGELOG / the actual tool registration.
- Baseline gate: `npm ci && npm run build && npx vitest run` green before any QA.
  Actual suite size is **349 tests** (the PRs' "305" was machine/point-in-time
  dependent). One pre-existing test (`docx-newline-handling.test.ts`,
  "re-editing a multi-line cell") false-failed on this machine because it
  asserted `not.toContain("x1"/"x3")` against `readDocument` output, which
  embeds the absolute temp Path (`/var/folders/.../h73hykpx0x1bd.../` contains
  "x1"). Fixed to assert on raw cell XML; engine behaviour was already correct.

## QA plan (multi-agent)

Read-only "dogfood persona" agents run in parallel against the built server +
real fixture DOCX, each returning a structured findings report
(severity / repro / expected vs actual / evidence):

1. #4 newline correctness & round-trips
2. #5 table editing (incl. tracked-mode replace question, merged cells,
   block-content preservation)
3. #6 introspection accuracy vs raw XML + table-aware search
4. #7 anchors (seed/repair/namespace, survival across mutations, error paths)
5. Cross-feature integration — realistic 業務委託契約書 multi-pass session
   chaining all features; output validated with `python-docx` + `xmllint`
6. Adversarial / monkey — unicode/emoji/RTL/XML-special chars, fractional/
   negative/NaN/out-of-range indices, empty docs, pre-existing revisions,
   duplicate/odd anchors, concurrent writes (file lock), large docs
7. UX/DX + schema — actual MCP layer (`node dist/index.js` JSON-RPC): 40 zod
   schemas vs descriptions, error-message quality, `<json>` response shape,
   README/CLAUDE.md accuracy

**Validation tooling:** `xmllint` (OOXML well-formedness on every part),
`python-docx` (independent open/re-save), `unzip`, engine round-trip (vitest).

**Fix loop:** triage + dedup → sequential TDD fixes (shared `docx-engine.ts`,
so no parallel-worktree fixes) → re-run affected personas + fresh integration
pass → loop until a round finds nothing new.

## Known risks / overlaps to verify

- **#5 tracked-mode replace:** append-vs-replace fixed only for *untracked*
  cell edits (#8). Confirm tracked behavior; treat as in-scope if it appends.
- **`copy_format_from` overlap:** #10 and #11 both strip stale
  `w:pPrChange`/`w:rPrChange`. #11 sits on top — confirm the final code path is
  coherent (not double-applied / dead).
- **Feature interaction:** multi-line split (#4) seeding anchors (#7) on all new
  paragraphs; anchor stability across table-paragraph ops (#5).

## Review gates

1. Spec-compliance (this doc's acceptance criteria)
2. Code quality
3. Security (opus, `~/.claude/skills/security-review/security-reviewer-prompt.md`)

All three LGTM required. Then optional final Codex gate (rate-limit checked
with user first). All-Claude reviewers must approve before Codex.

## Ship

Reconcile commit → push → one PR (`Fixes #4`, `Fixes #5`, `Fixes #7`,
`Closes #6` + follow-up issues for 6.4/6.7) → close the 4 stacked sub-PRs
pointing to it → leave merge to the maintainer.
