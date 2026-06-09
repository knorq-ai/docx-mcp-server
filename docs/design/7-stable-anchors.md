# Design: Stable paragraph anchors (Issue #7, Option B)

Status: **implemented** · Target version: 3.5.0

## Problem

Paragraphs are addressed only by integer block index. Every `insert`/`delete`
shifts the index of every later block, so a multi-step editing session must
re-read or re-`search_text` after almost every mutation. The agent's mental
model is "edit the clause that says X", not "edit block 67".

## Chosen approach

Option B from the issue: **persistent anchors backed by Word's built-in
`w14:paraId`**. An anchor is an opaque, stable per-paragraph id that survives
index shifts. Preferred over content locators (Option A) because it adds no new
addressing tools (aligned with the v3.0.0 "minimize schema-token surface"
decision), round-trips through Word (paraId is Word-native), and has no
false-match risk.

## Anchor identity

- The anchor string **is** the paragraph's `w14:paraId` — 8 uppercase hex
  chars, e.g. `3F2A11B0`.
- **MS-DOCX validity (revised):** `w14:paraId` must be a 32-bit value that is
  **nonzero and < 0x80000000** (i.e. `00000001`–`7FFFFFFF`) and **unique within
  the part**. The existing `generateParaId()` (`engine/comments.ts`) returns
  arbitrary 4 bytes, so ~half its outputs set the high bit (invalid) and
  `00000000` is possible — **it must not be used for document anchors as-is.**
  Add a dedicated generator:

  ```
  generateDocParaId(used: Set<string>): string
    // random uint32 masked to 0x7FFFFFFF, reroll on 0 or collision with `used`,
    // formatted as 8-char uppercase hex.
  ```

  **Part-wide uniqueness (revised):** the `used` set must be built from the
  `w14:paraId` of **every `<w:p>` in `word/document.xml`** — including table-cell
  and SDT paragraphs — not just direct-body paragraphs. v1 only *seeds/exposes*
  direct-body anchors, but a newly generated id must not collide with any
  existing paraId anywhere in the part (Word requires part-wide uniqueness, and
  collisions would corrupt comment threading / future table anchors). A small
  `collectAllParaIds(body)` walk (recursing into tables and SDT) feeds both the
  generator's `used` set and duplicate detection.

  (The comment-threading paraIds are a separate, pre-existing concern; this
  generator simply must avoid colliding with them.)
- Opaque to clients — they pass it back verbatim and must not construct it.

## Scope (v1)

**In scope:** **direct-body** top-level paragraphs — `<w:p>` that are direct
children of `<w:body>`. This is exactly the set `blockBodyIndices()` enumerates,
which is the set every mutation tool already resolves against
(`editParagraphs`, `deleteParagraphs`, `setParagraphFormats`, `setHeadings`,
`insertParagraphs`).

**Explicitly excluded (v1):**
- **Paragraphs inside a top-level `<w:sdt><w:sdtContent>`.** `enumerateBlocks()`
  (used by `read_document`) counts these, but `blockBodyIndices()` (used by all
  edits) does **not** — the two indexings already diverge for SDT docs. Anchors
  follow the *edit* model (`blockBodyIndices`), so SDT-contained paragraphs get
  **no anchor** in v1, and the read surfaces report `null` for them. Unifying
  the two indexings is a separate change.
- Table blocks (`<w:tbl>` has no paragraph mark → no paraId) and paragraphs
  inside table cells (the #5 tools stay index-based). Anchor reported `null`.
- Header/footer/footnote paragraphs.

**Internal resolver is designed for extension:** the resolver returns a
`ParagraphLocation` (below) carrying the parent container, so table-cell anchors
can be added later **without redesign** — only the seeding/enumeration set
widens.

## Resolved-location model

Anchor-based delete and insert-before/after need the *containing array* and
position, not just the element. The resolver returns:

```
interface ParagraphLocation {
  element: XNode;     // the <w:p> wrapper node
  parent: XNode[];    // the array that holds it (the document body in v1)
  bodyIndex: number;  // element's index within `parent`
  blockIndex: number; // its 0-based blockBodyIndices position (for messages)
}
```

`resolveAnchors(body, anchors)` builds a `paraId → ParagraphLocation[]` map over
current direct-body paragraphs and returns one location per requested anchor.

## OOXML namespace seeding (critical, revised)

Writing `w14:paraId` requires the `<w:document>` root to declare the `w14`
namespace. `setAttr()` works on the root node because preserveOrder attributes
live at `node[":@"]["@_name"]` (`xml-helpers.ts`), and the builder re-emits
them. `createDocument` currently declares only `xmlns:w`/`xmlns:r`. The seeding
step (`ensureW14Namespace(root)`, idempotent) operates on the root's
`:@` attribute bag, adding/normalizing only what is missing:

1. **`xmlns:w14`** — expected URI
   `http://schemas.microsoft.com/office/word/2010/wordml`.
   - absent → add it.
   - present and **equal** to the expected URI → leave as-is.
   - present but bound to a **different** URI → **refuse** with `INVALID_DOCX`
     (the `w14:` prefix already means something else in this part; writing
     `w14:paraId` would be semantically wrong, so we do not silently corrupt it).
2. **`xmlns:mc`** — same three-way rule with expected URI
   `http://schemas.openxmlformats.org/markup-compatibility/2006` (mismatch →
   `INVALID_DOCX`).
3. **`mc:Ignorable`** — token-set update, not string append. Split the existing
   value on whitespace into a token set (empty/absent → empty set), add `"w14"`
   only if absent, and re-join with single spaces. This handles a missing
   attribute, an empty attribute, an already-present `w14`, and unrelated
   existing tokens (e.g. `w15`) without clobbering or duplicating.

(Implementation note: the `w14` / `mc` prefixes are fixed because `w14:paraId`
and `mc:Ignorable` literally require those prefix strings as attribute names. The
mismatch check above guards the rare case where a doc has already bound one of
those prefixes to a *different* URI; rather than silently emit a wrong-namespace
attribute, seeding refuses with `INVALID_DOCX`.)

## Error codes

- **Reuse** `ANCHOR_NOT_FOUND` (already in `ErrorCode`) — anchor resolves to no
  current paragraph.
- **New** `AMBIGUOUS_ANCHOR` — anchor matches more than one paragraph
  (malformed/duplicate-id doc). Anchor **writes fail** rather than guessing.
- **New** `INVALID_LOCATOR` — an item supplied both `paragraph_index` and
  `anchor`, or neither, where exactly one is required.

## API surface

### New tool: `ensure_anchors(file_path)` (write, locked)

Idempotently seeds a valid unique `w14:paraId` on every direct-body paragraph
missing one, repairs **duplicate** paraIds (keep the first occurrence, reseed
later duplicates), and ensures the namespace declarations. Returns the full
map so a caller adopts anchors in one call:

```
ensure_anchors(file_path) -> {
  file, seeded, repaired,
  blocks: [{ index, type: "paragraph"|"table", anchor: string|null, textPreview }]
}
```

Reads stay **pure (no lock/save)**. Most Word-authored docs already carry
paraIds, so this mainly matters for server-created/stripped docs — but it is
always safe to call.

### Read tools — expose anchors (Q1 resolved: low-noise)

- `read_document`: **unchanged by default.** New opt-in `show_anchors: boolean`
  (default `false`) appends a compact `@<anchor>` token to paragraph lines that
  have one. Keeps the default read output clean for LLM consumption.
- `searchTextStructured` / `search_text`: `SearchMatch` gains `anchor?: string`
  (present only for direct-body paragraph matches; table-cell/SDT matches omit
  it). The text line is unchanged; the value rides in the `<json>` block.
- `get_document_info`: **unchanged.** Its structured output
  (`getDocumentInfoStructured`) is aggregate counts + a heading `outline`, with
  no per-block list, so there is nowhere to hang per-paragraph anchors without
  inventing a new `blocks` array — which `ensure_anchors` already provides. The
  full index→anchor map is therefore obtained from `ensure_anchors`; targeted
  anchors from `search_text`; inline anchors from `read_document` with
  `show_anchors: true`.

Reads never seed, so anchors are reported only when already present (after
`ensure_anchors` or any auto-seeding write).

### Edit tools — accept anchor as an alternative locator

Each item accepts **either** `paragraph_index` **or** `anchor` (exactly one;
else `INVALID_LOCATOR`). Resolution happens before any mutation;
`ANCHOR_NOT_FOUND` / `AMBIGUOUS_ANCHOR` on failure. Index behavior unchanged.

- `edit_paragraphs`: `edits[]` → `{ paragraph_index? | anchor?, new_text }`.
- `set_headings`: `headings[]` → `{ paragraph_index? | anchor?, level }`.
- `set_paragraph_formats`: each group gains `anchors?: string[]` beside
  `indices?: number[]` (a group may use either; at least one non-empty).
- `delete_paragraphs` (Q2 resolved): add a new optional
  `targets: [{ paragraph_index? | anchor? }]` param; keep `paragraph_indices:
  number[]` as a **deprecated** compat input (at least one of the two required).
  **Asymmetry documented:** `paragraph_index` can target a paragraph *or a
  table block*; an `anchor` only ever targets a paragraph (tables have no
  paraId). This is called out in the tool description.

### Insert — anchor-relative placement + returns new anchors

`insert_paragraphs` item placement gains an alternative to `position`:
`{ anchor, placement: "before" | "after" }` (exactly one of `position` /
`anchor` per item). `copy_format_from` (Q3 resolved) also accepts
`copy_format_from_anchor` as an alternative to the integer `copy_format_from`.
New paragraphs are seeded with fresh paraIds; the tool result reports each new
paragraph's `anchor` so a pipeline keeps editing them.

### Auto-seeding on write (Q5 resolved: touched-only)

A mutation tool seeds a paraId only on the paragraphs it **touches or creates**
(and ensures the namespace once). It does **not** seed untouched paragraphs —
that would turn a one-line edit into a whole-document diff. Whole-document
seeding stays `ensure_anchors`' explicit job. Newly inserted paragraphs always
get an anchor (returned in the result).

## Resolution & ambiguity

- Duplicate paraIds → anchor **writes fail** with `AMBIGUOUS_ANCHOR`;
  `ensure_anchors` repairs them.
- Anchor of an untracked-deleted paragraph → `ANCHOR_NOT_FOUND`.
- A **tracked**-deleted paragraph still exists in the tree, so its anchor still
  resolves (consistent with tracked-delete keeping the node).

## Backward compatibility

Purely additive. Every existing `paragraph_index` / `paragraph_indices` /
`position` / `copy_format_from` parameter keeps working unchanged. Default
`read_document` output is byte-for-byte unchanged (anchors are opt-in there).

## Testing plan

- `generateDocParaId`: always in `00000001`–`7FFFFFFF`; never collides with a
  supplied used-set; 8-char uppercase.
- `ensure_anchors`: seeds missing ids; declares `xmlns:w14`, `xmlns:mc`, and
  `mc:Ignorable=…w14`; idempotent (second call seeds 0); preserves existing ids;
  repairs duplicate ids; returns correct map; leaves a Word-style doc (ids +
  namespaces already present) byte-stable.
- Round-trip: read anchor → delete an earlier paragraph (indices shift) → edit
  by the original anchor → correct paragraph changed.
- `edit_paragraphs` / `set_headings` / `set_paragraph_formats` / `delete_paragraphs`
  by anchor.
- `insert_paragraphs` before/after anchor and `copy_format_from_anchor`; returned
  new anchors resolve on a follow-up edit.
- Locator validation: both/neither → `INVALID_LOCATOR`; stale → `ANCHOR_NOT_FOUND`;
  duplicate id → `AMBIGUOUS_ANCHOR`.
- Namespace: a `createDocument` doc gains the three declarations exactly once;
  re-running seeding adds nothing.
- SDT exclusion: a doc with a top-level `w:sdt` paragraph reports `anchor: null`
  for it and `ensure_anchors` does not seed it.
- Tracked-delete then resolve-by-anchor succeeds; untracked-delete then resolve
  → `ANCHOR_NOT_FOUND`.

## Decisions on the original open questions

- **Q1:** Default `read_document` unchanged; anchors via `search_text` JSON +
  `ensure_anchors` (full map), plus opt-in `show_anchors` on `read_document`.
  `get_document_info` is left unchanged (no per-block list to extend).
- **Q2:** `delete_paragraphs` gains `targets: [{index?|anchor?}]`;
  `paragraph_indices` kept as deprecated compat.
- **Q3:** `copy_format_from_anchor` supported in v1 (insert-by-anchor is in).
- **Q4:** Direct-body paragraphs only; **SDT paragraphs excluded** (they aren't
  in the edit-index set); table-cell anchors deferred but the resolver carries
  parent/container metadata so they slot in later without redesign.
- **Q5:** Touched-only auto-seed; `ensure_anchors` for whole-document bootstrap.
