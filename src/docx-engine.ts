/**
 * DOCX Engine — Direct OOXML manipulation for the MCP server.
 *
 * This is a barrel module that re-exports helpers from engine/ sub-modules
 * and defines the public API functions consumed by index.ts.
 */

import * as path from "path";
import * as fs from "fs/promises";
import JSZip from "jszip";

import { withFileLock } from "./engine/file-lock.js";

// Re-export types and helpers needed by consumers (index.ts, tests)
export { ErrorCode, EngineError, escapeXml } from "./engine/docx-io.js";
export type { ErrorCodeType } from "./engine/docx-io.js";
export type { BlockInfo } from "./engine/text.js";
export type { TextFormatting, ParagraphFormat } from "./engine/formatting.js";
export type { CommentInfo, BatchCommentInput } from "./engine/comments.js";
export type { PageLayoutOptions } from "./engine/layout.js";
export type { ImageInfo, ListImagesResult } from "./engine/images.js";

// Internal imports
import {
  type XNode,
  parser,
  builder,
  attr,
  setAttr,
  findAll,
  findOne,
  el,
  textNode,
  cloneNode,
  normalizeNewlines,
} from "./engine/xml-helpers.js";
import {
  ErrorCode,
  EngineError,
  type DocxHandle,
  openDocx,
  saveDocx,
  parseDocXml,
  serializeDocXml,
  getBody,
  forEachParagraphInTable,
  getHeaderFooterFiles,
  escapeXml,
} from "./engine/docx-io.js";
import {
  type RevisionContext,
  extractRunText,
  extractParagraphText,
  extractTableText,
  extractCellText,
  getParagraphStyle,
  getHeadingLevel,
  enumerateBlocks,
  type BlockRef,
  enumerateBlockRefs,
  replaceInParagraph,
  replaceInParagraphTracked,
  paragraphHasRevisions,
  scanMaxId,
  newRevisionContext,
  allocRevId,
  getRunRPr,
  makeTextRun,
  makeTextRuns,
  makeDelTextRun,
  makeDelTextRuns,
  wrapInDel,
  wrapInIns,
  computeMinimalDiff,
  collectRunsWithIndices,
  acceptChangesInNodes,
  rejectChangesInNodes,
  extractTextFromHdrFtr,
} from "./engine/text.js";
import {
  type TextFormatting,
  type ParagraphFormat,
  setRunFormatting,
  applyParagraphFormat,
  formatInParagraph,
} from "./engine/formatting.js";
import {
  type CommentInfo,
  type BatchCommentInput,
  parseCommentsXml,
  getCommentsArray,
  getNextCommentId,
  ensureCommentsInfrastructure,
  parseCommentsExtendedXml,
  getCommentsExtendedArray,
  ensureCommentsExtendedInfrastructure,
  generateParaId,
  findAnchorParagraph,
  insertCommentRangeMarkers,
} from "./engine/comments.js";
import {
  type PageLayoutOptions,
  TWIPS_PER_MM,
  TWIPS_PER_INCH,
  PAGE_SIZE_PRESETS,
  MARGIN_PRESETS,
  twipsToMm,
  mmToTwips,
  detectPageSizePreset,
  getSectPr,
} from "./engine/layout.js";
import {
  type ImageInfo,
  type ListImagesResult,
  scanImages,
} from "./engine/images.js";
import {
  type ParagraphLocation,
  collectAllParaIds,
  getDocumentRoot,
  ensureW14Namespace,
  directBodyParagraphLocations,
  buildAnchorIndex,
  resolveAnchor,
  seedParagraphAnchor,
  isValidParaId,
} from "./engine/anchors.js";

// ===========================================================================================
// PUBLIC API
// ===========================================================================================

/**
 * Serialize `value` into a delimiter-safe `<json>…</json>` block (M4).
 *
 * `JSON.stringify` does not escape the angle-bracket characters, so document
 * text that literally contains the `</json>` sequence would appear verbatim
 * inside the payload and yield a second, ambiguous `</json>` that breaks
 * extraction. We unicode-escape every "<" to "\\u003c" and ">" to "\\u003e" in
 * the stringified output; this is still valid JSON that `JSON.parse` decodes
 * back to the original characters, so consumers get correct data AND a literal
 * `</json>` sequence can never occur inside the payload.
 *
 * Consumers should extract the LAST `<json>…</json>` pair in a tool's output
 * (the structured block is always emitted last, after the human-readable prose).
 */
function jsonBlock(value: unknown): string {
  const safe = JSON.stringify(value).replace(/</g, "\\u003c").replace(/>/g, "\\u003e");
  return "<json>\n" + safe + "\n</json>";
}

/**
 * Throw `PENDING_REVISIONS` with a consistent message pointing the caller
 * at `accept_all_changes` / `reject_all_changes`. Used by tracked-mode
 * editing operations that refuse to operate on paragraphs already
 * containing tracked-change markup, because the existing matcher treats
 * runs inside an existing w:ins as if they were normal text and produces
 * nested w:ins/w:del that does not round-trip through accept/reject.
 */
function throwPendingRevisions(where: string): never {
  throw new EngineError(
    ErrorCode.PENDING_REVISIONS,
    `${where} contains pre-existing tracked changes (w:ins/w:del). ` +
      `Tracked editing operations cannot safely modify content that already ` +
      `has revision markup — the resulting nested markup would not round-trip ` +
      `through accept_all_changes / reject_all_changes. ` +
      `Resolve existing revisions first via accept_all_changes or ` +
      `reject_all_changes, or set track_changes: false (with allow_untracked_edit: true) ` +
      `to apply the edit without generating new revision markup.`,
  );
}

/**
 * Walk the body and report the first paragraph that has pending revisions
 * AND would be affected by the operation. The `predicate` decides which
 * paragraphs are considered "affected" — for replaceTexts that's any
 * paragraph containing a match for any item; for editParagraphs/deleteParagraphs/etc.
 * that's only the paragraphs at the targeted indices. Returns null if
 * no conflicting paragraph is found.
 */
function findRevisionConflict(
  body: XNode[],
  isAffected: (pChildren: XNode[]) => boolean,
): { kind: "body" | "table"; description: string } | null {
  for (let i = 0; i < body.length; i++) {
    const child = body[i];
    if (child["w:p"]) {
      const pChildren = child["w:p"] as XNode[];
      if (isAffected(pChildren) && paragraphHasRevisions(pChildren)) {
        return { kind: "body", description: `Body paragraph at body-index ${i}` };
      }
    } else if (child["w:tbl"]) {
      let conflict: { row: number; col: number } | null = null;
      forEachParagraphInTable(child["w:tbl"], (pChildren) => {
        if (conflict) return;
        if (isAffected(pChildren) && paragraphHasRevisions(pChildren)) {
          conflict = { row: -1, col: -1 };
        }
      });
      if (conflict) {
        return { kind: "table", description: `A cell in the table at body-index ${i}` };
      }
    }
  }
  return null;
}

/**
 * Compute the maximum revision-style w:id across every part of the DOCX
 * (document.xml plus all header/footer/footnote/endnote XML), so newly
 * minted w:ins/w:del IDs are guaranteed unique throughout the file.
 *
 * scanMaxId on the body alone misses pre-existing revision markup in
 * header/footer parts; seeding from the body's max can collide with
 * existing revision IDs in those parts when Word later round-trips
 * the document. Reading the other parts as raw XML and matching
 * `w:id="N"` is much cheaper than parsing each part.
 */
async function scanMaxIdAcrossParts(
  handle: DocxHandle,
  bodyParsed: XNode[],
): Promise<number> {
  let max = scanMaxId(bodyParsed);
  for (const name of Object.keys(handle.zip.files)) {
    if (name === "word/document.xml") continue;
    if (!name.startsWith("word/") || !name.endsWith(".xml")) continue;
    const xml = await handle.zip.file(name)?.async("string");
    if (!xml) continue;
    // XML allows single OR double quotes and whitespace around '='. Match both.
    for (const m of xml.matchAll(/\bw:id\s*=\s*["'](\d+)["']/g)) {
      const v = parseInt(m[1], 10);
      if (!isNaN(v) && v > max) max = v;
    }
  }
  return max;
}

// ---------------------------------------------------------------------------
// Stable paragraph anchors (issue #7) — locator resolution + auto-seeding
// ---------------------------------------------------------------------------

/**
 * Resolve a block index against the canonical block-index space (the SAME
 * numbering read_document / get_document_info / search_text expose). Throws
 * INDEX_OUT_OF_RANGE for a non-integer / out-of-bounds index. The returned
 * `BlockRef` carries the element plus its container + position so callers can
 * read, edit in place, or splice at exactly the right spot — including for a
 * paragraph that lives inside a top-level `w:sdt`.
 */
function resolveBlockRef(refs: BlockRef[], blockIndex: number): BlockRef {
  if (!Number.isInteger(blockIndex) || blockIndex < 0 || blockIndex >= refs.length) {
    throw new EngineError(
      ErrorCode.INDEX_OUT_OF_RANGE,
      `Block index ${blockIndex} out of range (0–${refs.length - 1}).`,
    );
  }
  return refs[blockIndex];
}

/** A paragraph locator: exactly one of paragraphIndex / anchor. */
export interface ParagraphLocator {
  paragraphIndex?: number;
  anchor?: string;
}

/**
 * Resolve a locator (exactly one of paragraph_index / anchor) to a paragraph
 * location. `anchorIndex` is built once per call via buildAnchorIndex(body).
 */
function locatorToLocation(
  refs: BlockRef[],
  anchorIndex: Map<string, ParagraphLocation[]>,
  loc: ParagraphLocator,
): ParagraphLocation {
  const hasIndex = loc.paragraphIndex !== undefined;
  const hasAnchor = loc.anchor !== undefined && loc.anchor !== "";
  if (hasIndex === hasAnchor) {
    throw new EngineError(
      ErrorCode.INVALID_LOCATOR,
      `Provide exactly one of paragraph_index or anchor.`,
    );
  }
  if (hasAnchor) {
    return resolveAnchor(anchorIndex, loc.anchor as string);
  }
  const idx = loc.paragraphIndex as number;
  const ref = resolveBlockRef(refs, idx);
  if (!ref.element["w:p"]) {
    throw new EngineError(ErrorCode.NOT_A_PARAGRAPH, `Block ${idx} is not a paragraph (it may be a table).`);
  }
  return { element: ref.element, parent: ref.container, bodyIndex: ref.indexInContainer };
}

/**
 * Lazy auto-seeder for the paragraphs a write touches or creates. Declares the
 * w14 namespace on first use and assigns paraIds from a part-wide `used` set so
 * new ids stay unique. Touched-only: untouched paragraphs are never seeded.
 */
class AnchorSeeder {
  private used: Set<string> | null = null;
  private nsReady = false;
  constructor(private root: XNode, private body: XNode[]) {}

  /** Ensure `element` has an anchor, returning it. */
  seed(element: XNode): string {
    if (!this.used) {
      this.used = collectAllParaIds(this.body).all;
    }
    const existing = attr(element, "w14:paraId");
    if (existing) return existing;
    if (!this.nsReady) {
      ensureW14Namespace(this.root);
      this.nsReady = true;
    }
    return seedParagraphAnchor(element, this.used);
  }
}

// ---------------------------------------------------------------------------
// ensure_anchors
// ---------------------------------------------------------------------------

export interface AnchorBlockInfo {
  index: number;
  type: "paragraph" | "table";
  anchor: string | null;
  textPreview: string;
}

export interface EnsureAnchorsResult {
  file: string;
  seeded: number;
  repaired: number;
  blocks: AnchorBlockInfo[];
}

const ANCHOR_PREVIEW_LEN = 50;

function anchorPreview(text: string): string {
  const oneLine = text.replace(/\s+/g, " ").trim();
  return oneLine.length > ANCHOR_PREVIEW_LEN ? oneLine.slice(0, ANCHOR_PREVIEW_LEN) + "…" : oneLine;
}

export async function ensureAnchorsStructured(filePath: string): Promise<EnsureAnchorsResult> {
  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);
    const parsed = await parseDocXml(handle);
    const root = getDocumentRoot(parsed);
    const body = getBody(parsed);

    const { all: used, duplicates } = collectAllParaIds(body);
    let seeded = 0;
    let repaired = 0;
    let nsReady = false;
    const ensureNs = (): void => {
      if (!nsReady) {
        ensureW14Namespace(root);
        nsReady = true;
      }
    };

    // Guarantee every direct-body paragraph ends with a valid, unique paraId not
    // shared (case-insensitively) with any other <w:p> in the part. Reseed when
    // the id is missing, invalid, part-wide-duplicated, or case-collides with an
    // earlier body paragraph.
    const assigned = new Set<string>(); // canonical (uppercase) ids
    for (const loc of directBodyParagraphLocations(body)) {
      const id = attr(loc.element, "w14:paraId");
      const canonical = id?.toUpperCase();
      const needsReseed =
        !id ||
        !isValidParaId(id) ||
        duplicates.has(id) ||
        (canonical !== undefined && assigned.has(canonical));
      if (!needsReseed) {
        assigned.add(canonical as string);
        continue;
      }
      ensureNs();
      const fresh = seedFreshAnchor(loc.element, used);
      assigned.add(fresh);
      if (id) repaired++;
      else seeded++;
    }

    if (seeded > 0 || repaired > 0) {
      serializeDocXml(handle, parsed);
      await saveDocx(handle);
    }

    // Build the full block map (paragraphs + tables) for the result, numbered in
    // the SAME unified block-index space read_document / get_document_info /
    // search_text use (M5). SDT-contained paragraphs are listed (and advance the
    // index) but carry no anchor — anchor scope is still direct-body paragraphs
    // only, so only the printed index changes for SDT docs, never the anchors.
    const blocks: AnchorBlockInfo[] = [];
    enumerateBlockRefs(body).forEach((ref, index) => {
      if (ref.type === "paragraph") {
        const isDirectBody = ref.container === body;
        blocks.push({
          index,
          type: "paragraph",
          anchor: isDirectBody ? attr(ref.element, "w14:paraId") ?? null : null,
          textPreview: anchorPreview(extractParagraphText(ref.element["w:p"] as XNode[], false)),
        });
      } else {
        blocks.push({
          index,
          type: "table",
          anchor: null,
          textPreview: anchorPreview(extractTableText(ref.element["w:tbl"] as XNode[], false)),
        });
      }
    });

    return { file: path.basename(filePath), seeded, repaired, blocks };
  });
}

/** Assign a brand-new anchor to a paragraph, overwriting any existing id (duplicate repair). */
function seedFreshAnchor(element: XNode, used: Set<string>): string {
  // Remove the stale id from the element so seedParagraphAnchor assigns a new one.
  if (element[":@"]) delete element[":@"]["@_w14:paraId"];
  return seedParagraphAnchor(element, used);
}

export async function ensureAnchors(filePath: string): Promise<string> {
  const r = await ensureAnchorsStructured(filePath);
  let out = `Anchors in ${r.file}: seeded ${r.seeded}, repaired ${r.repaired}, ${r.blocks.length} block(s).\n\n`;
  for (const b of r.blocks) {
    const a = b.anchor ? `@${b.anchor}` : "(no anchor)";
    out += `[${b.index}] ${a}${b.type === "table" ? " [table]" : ""} ${b.textPreview}\n`;
  }
  out += "\n" + jsonBlock(r);
  return out;
}

// ---------------------------------------------------------------------------
// 1. read_document
// ---------------------------------------------------------------------------

export async function readDocument(
  filePath: string,
  startParagraph?: number,
  endParagraph?: number,
  showRevisions: boolean = false,
  showAnchors: boolean = false,
): Promise<string> {
  const handle = await openDocx(filePath);
  const parsed = await parseDocXml(handle);
  const body = getBody(parsed);
  const blocks = enumerateBlocks(body, showRevisions);
  const anchorByBlock = showAnchors ? buildAnchorByBlockIndex(body) : undefined;

  const start = startParagraph ?? 0;
  const end = endParagraph ?? blocks.length;
  const filtered = blocks.slice(start, end);

  let output = `Document: ${path.basename(filePath)}\n`;
  output += `Path: ${filePath}\n`;
  output += `Total blocks: ${blocks.length}\n`;
  if (start > 0 || end < blocks.length) {
    output += `Showing blocks ${start}–${Math.min(end, blocks.length) - 1}\n`;
  }
  output += "---\n\n";

  for (const b of filtered) {
    let prefix = `[${b.index}]`;
    if (showAnchors) {
      const a = anchorByBlock?.get(b.index);
      if (a) prefix += ` @${a}`;
    }
    if (b.headingLevel) {
      prefix += ` (H${b.headingLevel})`;
    } else if (b.style && b.style !== "Normal" && b.style !== "Table") {
      prefix += ` (${b.style})`;
    }
    if (b.alignment && b.alignment !== "left") {
      prefix += ` [align:${b.alignment}]`;
    }
    if (b.type === "table") {
      prefix += " [table]";
    }
    output += `${prefix} ${b.text}\n`;
  }

  return output;
}

// ---------------------------------------------------------------------------
// 2. get_document_info
// ---------------------------------------------------------------------------

export async function getDocumentInfo(filePath: string): Promise<string> {
  const info = await getDocumentInfoStructured(filePath);

  let output = `Document: ${info.file}\n`;
  output += `Path: ${info.path}\n`;
  output += `Total blocks: ${info.totalBlocks}\n`;
  output += `  Headings: ${info.headings}\n`;
  output += `  Paragraphs: ${info.paragraphs}\n`;
  output += `  Tables: ${info.tables}\n`;
  output += `  Has comments: ${info.hasComments}\n`;

  if (info.outline.length > 0) {
    output += "\nDocument outline:\n";
    for (const h of info.outline) {
      const indent = "  ".repeat(h.level - 1);
      output += `${indent}H${h.level}: ${h.text} [block ${h.blockIndex}]\n`;
    }
  }

  output += "\n\n" + jsonBlock(info);
  return output;
}

// ---------------------------------------------------------------------------
// 2b. get_document_info (structured)
// ---------------------------------------------------------------------------

/** getDocumentInfoStructured の戻り値型 */
export interface DocumentInfoResult {
  file: string;
  path: string;
  totalBlocks: number;
  headings: number;
  paragraphs: number;
  tables: number;
  hasComments: boolean;
  outline: { level: number; text: string; blockIndex: number }[];
}

/**
 * getDocumentInfo と同じロジックだが、テキストではなく構造化オブジェクトを返す。
 */
export async function getDocumentInfoStructured(
  filePath: string,
): Promise<DocumentInfoResult> {
  const handle = await openDocx(filePath);
  const parsed = await parseDocXml(handle);
  const body = getBody(parsed);
  const blocks = enumerateBlocks(body);

  const headingBlocks = blocks.filter((b) => b.headingLevel);
  const tables = blocks.filter((b) => b.type === "table");
  const paragraphs = blocks.filter(
    (b) => b.type === "paragraph" && !b.headingLevel,
  );

  // コメントの有無を確認
  const commentsXml = await handle.zip
    .file("word/comments.xml")
    ?.async("string");
  const hasComments = !!commentsXml && commentsXml.includes("w:comment");

  const outline = headingBlocks.map((h) => ({
    level: h.headingLevel!,
    text: h.text.substring(0, 100) + (h.text.length > 100 ? "..." : ""),
    blockIndex: h.index,
  }));

  return {
    file: path.basename(filePath),
    path: filePath,
    totalBlocks: blocks.length,
    headings: headingBlocks.length,
    paragraphs: paragraphs.length,
    tables: tables.length,
    hasComments,
    outline,
  };
}

// ---------------------------------------------------------------------------
// 3. search_text
// ---------------------------------------------------------------------------

export async function searchText(
  filePath: string,
  query: string,
  caseSensitive: boolean = false,
): Promise<string> {
  const result = await searchTextStructured(filePath, query, caseSensitive);

  if (result.totalMatches === 0) {
    // Still emit the structured block (zero totals / empty array) so a program
    // that always parses the <json> block has a stable contract (L2).
    return `No matches found for "${query}" in ${result.file}.\n\n` + jsonBlock(result);
  }

  let output = `Found ${result.totalMatches} match(es) for "${query}":\n\n`;
  for (const m of result.matches) {
    const loc =
      m.rowIndex !== undefined
        ? `[Block ${m.blockIndex} · cell ${m.rowIndex},${m.colIndex}]`
        : `[Block ${m.blockIndex}]`;
    const a = m.anchor ? ` @${m.anchor}` : "";
    output += `${loc}${a} ${m.context}\n`;
  }

  output += "\n\n" + jsonBlock(result);
  return output;
}

// ---------------------------------------------------------------------------
// 3b. search_text (structured)
// ---------------------------------------------------------------------------

/** 検索結果の個別マッチ */
export interface SearchMatch {
  blockIndex: number;
  occurrences: number;
  context: string;
  fullText: string;
  /** When the match is inside a table cell, the cell's zero-based row index. */
  rowIndex?: number;
  /** When the match is inside a table cell, the cell's zero-based column index. */
  colIndex?: number;
  /** Stable anchor of the matched paragraph, when it is a direct-body paragraph
   * that already carries a w14:paraId (table/SDT matches omit it). */
  anchor?: string;
}

/**
 * Map enumerateBlocks() block index → anchor, mirroring its walk (paragraph and
 * table each advance the index; SDT-contained paragraphs advance it but are not
 * anchorable in v1). Only direct-body paragraphs that already have a paraId are
 * recorded, so a match's anchor is correct or absent — never mismatched.
 */
function buildAnchorByBlockIndex(body: XNode[]): Map<number, string> {
  const map = new Map<number, string>();
  let idx = 0;
  for (const node of body) {
    if (node["w:p"]) {
      const a = attr(node, "w14:paraId");
      if (a) map.set(idx, a);
      idx++;
    } else if (node["w:tbl"]) {
      idx++;
    } else if (node["w:sdt"]) {
      const content = findOne(node["w:sdt"] as XNode[], "w:sdtContent");
      if (content) {
        for (const cc of content["w:sdtContent"] as XNode[]) {
          if (cc["w:p"]) idx++;
        }
      }
    }
  }
  return map;
}

/** searchTextStructured の戻り値型 */
export interface SearchTextResult {
  file: string;
  query: string;
  totalMatches: number;
  matches: SearchMatch[];
}

/** Count occurrences of searchStr in compare and build a context snippet from fullText. */
function buildTextMatch(
  fullText: string,
  compare: string,
  searchStr: string,
  queryLen: number,
): { occurrences: number; context: string } | null {
  // An empty needle never advances `searchFrom` (indexOf returns the same index
  // forever) → infinite synchronous loop that bricks the single-threaded server.
  // Report zero matches instead (the MCP layer also rejects empty queries).
  if (searchStr.length === 0) return null;

  const first = compare.indexOf(searchStr);
  if (first === -1) return null;

  let occurrences = 0;
  let searchFrom = 0;
  while (true) {
    const idx = compare.indexOf(searchStr, searchFrom);
    if (idx === -1) break;
    occurrences++;
    searchFrom = idx + searchStr.length;
  }

  const ctxStart = Math.max(0, first - 40);
  const ctxEnd = Math.min(fullText.length, first + queryLen + 40);
  const context =
    (ctxStart > 0 ? "..." : "") +
    fullText.substring(ctxStart, ctxEnd) +
    (ctxEnd < fullText.length ? "..." : "");
  return { occurrences, context };
}

/**
 * searchText と同じロジックだが、テキストではなく構造化オブジェクトを返す。
 *
 * Block indices mirror read_document. A match inside a table descends into the
 * table and is reported per cell with rowIndex / colIndex (instead of a single
 * block-level match against the flattened table text), so callers can drive
 * edit_table_cells directly.
 */
export async function searchTextStructured(
  filePath: string,
  query: string,
  caseSensitive: boolean = false,
): Promise<SearchTextResult> {
  const handle = await openDocx(filePath);
  const parsed = await parseDocXml(handle);
  const body = getBody(parsed);

  const searchStr = caseSensitive ? query : query.toLowerCase();
  const matches: SearchMatch[] = [];
  let totalMatches = 0;

  // An empty query matches everything / nothing meaningfully and would otherwise
  // spin in buildTextMatch's loop — return zero matches immediately.
  if (searchStr.length === 0) {
    return { file: path.basename(filePath), query, totalMatches: 0, matches };
  }

  // A direct-body paragraph match carries its anchor (when present); SDT/table
  // matches omit it (those paragraphs aren't anchored in v1).
  const tryMatchParagraph = (blockIndex: number, pChildren: XNode[], anchor?: string): void => {
    const text = extractParagraphText(pChildren, false);
    const compare = caseSensitive ? text : text.toLowerCase();
    const m = buildTextMatch(text, compare, searchStr, query.length);
    if (m) {
      totalMatches += m.occurrences;
      matches.push({
        blockIndex,
        occurrences: m.occurrences,
        context: m.context,
        fullText: text,
        ...(anchor ? { anchor } : {}),
      });
    }
  };

  // Walk the body in the same order/indexing as enumerateBlocks (table counts
  // as one block; w:sdt content controls expand to their inner paragraphs).
  let idx = 0;
  for (const child of body) {
    if (child["w:p"]) {
      tryMatchParagraph(idx, child["w:p"] as XNode[], attr(child, "w14:paraId"));
      idx++;
    } else if (child["w:tbl"]) {
      const rows = findAll(child["w:tbl"] as XNode[], "w:tr");
      rows.forEach((row, rowIndex) => {
        const cells = findAll(row["w:tr"] as XNode[], "w:tc");
        cells.forEach((tc, colIndex) => {
          const text = tableCellText(tc["w:tc"] as XNode[], false);
          const compare = caseSensitive ? text : text.toLowerCase();
          const m = buildTextMatch(text, compare, searchStr, query.length);
          if (m) {
            totalMatches += m.occurrences;
            matches.push({
              blockIndex: idx,
              occurrences: m.occurrences,
              context: m.context,
              fullText: text,
              rowIndex,
              colIndex,
            });
          }
        });
      });
      idx++;
    } else if (child["w:sdt"]) {
      const sdtContent = findOne(child["w:sdt"] as XNode[], "w:sdtContent");
      if (sdtContent) {
        for (const contentChild of sdtContent["w:sdtContent"] as XNode[]) {
          if (contentChild["w:p"]) {
            tryMatchParagraph(idx, contentChild["w:p"] as XNode[]);
            idx++;
          }
        }
      }
    }
  }

  return {
    file: path.basename(filePath),
    query,
    totalMatches,
    matches,
  };
}

// ---------------------------------------------------------------------------
// Table reading helpers (read-only navigation + introspection)
// ---------------------------------------------------------------------------

/** Resolve a table block to its <w:tbl> children, validating index and type. */
function getTableChildren(refs: BlockRef[], blockIndex: number): XNode[] {
  const ref = resolveBlockRef(refs, blockIndex);
  if (!ref.element["w:tbl"]) {
    throw new EngineError(ErrorCode.NOT_A_TABLE, `Block ${blockIndex} is not a table.`);
  }
  return ref.element["w:tbl"] as XNode[];
}

/** Horizontal merge span of a cell (w:gridSpan), default 1. */
function cellGridSpan(tcChildren: XNode[]): number {
  const tcPr = findOne(tcChildren, "w:tcPr");
  if (!tcPr) return 1;
  const gs = findOne(tcPr["w:tcPr"] as XNode[], "w:gridSpan");
  if (!gs) return 1;
  const v = parseInt(attr(gs, "w:val") ?? "1", 10);
  return Number.isFinite(v) && v > 0 ? v : 1;
}

/**
 * Vertical merge state of a cell: "restart" (top of a merged span),
 * "continue" (covered by the cell above), or null (not vertically merged).
 * A bare <w:vMerge/> with no w:val means "continue".
 */
function cellVMerge(tcChildren: XNode[]): string | null {
  const tcPr = findOne(tcChildren, "w:tcPr");
  if (!tcPr) return null;
  const vm = findOne(tcPr["w:tcPr"] as XNode[], "w:vMerge");
  if (!vm) return null;
  return attr(vm, "w:val") ?? "continue";
}

/** Text of one cell: its paragraphs joined by real "\n" (nested tables flattened). */
function tableCellText(tcChildren: XNode[], showRevisions: boolean): string {
  const parts: string[] = [];
  for (const child of tcChildren) {
    if (child["w:p"]) parts.push(extractParagraphText(child["w:p"], showRevisions));
    else if (child["w:tbl"]) parts.push(extractTableText(child["w:tbl"], showRevisions));
  }
  return parts.join("\n");
}

// ---------------------------------------------------------------------------
// 3c. read_table_structure (structured)
// ---------------------------------------------------------------------------

export interface TableStructureCell {
  rowIndex: number;
  colIndex: number;
  /** Horizontal span (w:gridSpan); >1 means this cell covers several grid columns. */
  gridSpan: number;
  /** Vertical merge: "restart" | "continue" | null. */
  vMerge: string | null;
  /** Truncated cell text for a quick overview. */
  preview: string;
}

export interface TableStructureResult {
  file: string;
  blockIndex: number;
  rows: number;
  /** Maximum number of physical <w:tc> cells in any row. */
  cols: number;
  cells: TableStructureCell[];
}

const PREVIEW_LEN = 40;

function previewText(text: string): string {
  const oneLine = text.replace(/\s+/g, " ").trim();
  return oneLine.length > PREVIEW_LEN ? oneLine.slice(0, PREVIEW_LEN) + "…" : oneLine;
}

export async function readTableStructureStructured(
  filePath: string,
  blockIndex: number,
): Promise<TableStructureResult> {
  const handle = await openDocx(filePath);
  const parsed = await parseDocXml(handle);
  const body = getBody(parsed);
  const refs = enumerateBlockRefs(body);
  const tblChildren = getTableChildren(refs, blockIndex);

  const rows = findAll(tblChildren, "w:tr");
  const cells: TableStructureCell[] = [];
  let maxCols = 0;
  rows.forEach((row, rowIndex) => {
    const tcs = findAll(row["w:tr"] as XNode[], "w:tc");
    if (tcs.length > maxCols) maxCols = tcs.length;
    tcs.forEach((tc, colIndex) => {
      const tcChildren = tc["w:tc"] as XNode[];
      cells.push({
        rowIndex,
        colIndex,
        gridSpan: cellGridSpan(tcChildren),
        vMerge: cellVMerge(tcChildren),
        preview: previewText(tableCellText(tcChildren, false)),
      });
    });
  });

  return {
    file: path.basename(filePath),
    blockIndex,
    rows: rows.length,
    cols: maxCols,
    cells,
  };
}

export async function readTableStructure(filePath: string, blockIndex: number): Promise<string> {
  const r = await readTableStructureStructured(filePath, blockIndex);
  let out = `Table at block ${r.blockIndex} in ${r.file}: ${r.rows} row(s) × up to ${r.cols} column(s).\n\n`;
  for (let row = 0; row < r.rows; row++) {
    const rowCells = r.cells.filter((c) => c.rowIndex === row);
    const parts = rowCells.map((c) => {
      let label = `[${c.rowIndex},${c.colIndex}]`;
      if (c.gridSpan > 1) label += `(span ${c.gridSpan})`;
      if (c.vMerge) label += `(vMerge:${c.vMerge})`;
      return `${label} ${c.preview}`;
    });
    out += "| " + parts.join(" | ") + " |\n";
  }
  out += "\n" + jsonBlock(r);
  return out;
}

// ---------------------------------------------------------------------------
// 3d. read_table_cell (structured)
// ---------------------------------------------------------------------------

export interface TableCellParagraphInfo {
  paragraphIndex: number;
  text: string;
  style?: string;
  alignment?: string;
  numId?: number;
  numLevel?: number;
}

export interface TableCellReadResult {
  file: string;
  blockIndex: number;
  rowIndex: number;
  colIndex: number;
  gridSpan: number;
  vMerge: string | null;
  paragraphs: TableCellParagraphInfo[];
}

/** Navigate to a cell's <w:tc> children (read-only), validating each index. */
function getTableCellChildren(
  refs: BlockRef[],
  blockIndex: number,
  rowIndex: number,
  colIndex: number,
): XNode[] {
  const tblChildren = getTableChildren(refs, blockIndex);
  const rows = findAll(tblChildren, "w:tr");
  if (!Number.isInteger(rowIndex) || rowIndex < 0 || rowIndex >= rows.length) {
    throw new EngineError(ErrorCode.INDEX_OUT_OF_RANGE, `Row index ${rowIndex} out of range (0–${rows.length - 1}).`);
  }
  const cells = findAll(rows[rowIndex]["w:tr"] as XNode[], "w:tc");
  if (!Number.isInteger(colIndex) || colIndex < 0 || colIndex >= cells.length) {
    throw new EngineError(ErrorCode.INDEX_OUT_OF_RANGE, `Column index ${colIndex} out of range (0–${cells.length - 1}).`);
  }
  return cells[colIndex]["w:tc"] as XNode[];
}

export async function readTableCellStructured(
  filePath: string,
  blockIndex: number,
  rowIndex: number,
  colIndex: number,
): Promise<TableCellReadResult> {
  const handle = await openDocx(filePath);
  const parsed = await parseDocXml(handle);
  const body = getBody(parsed);
  const refs = enumerateBlockRefs(body);
  const tcChildren = getTableCellChildren(refs, blockIndex, rowIndex, colIndex);

  const paragraphs: TableCellParagraphInfo[] = [];
  let pIdx = 0;
  for (const child of tcChildren) {
    if (child["w:p"]) {
      const pChildren = child["w:p"] as XNode[];
      const numPr = readNumbering(pChildren);
      paragraphs.push({
        paragraphIndex: pIdx,
        text: extractParagraphText(pChildren, false),
        style: getParagraphStyle(pChildren),
        alignment: readAlignment(pChildren),
        numId: numPr?.numId,
        numLevel: numPr?.numLevel,
      });
      pIdx++;
    }
  }

  return {
    file: path.basename(filePath),
    blockIndex,
    rowIndex,
    colIndex,
    gridSpan: cellGridSpan(tcChildren),
    vMerge: cellVMerge(tcChildren),
    paragraphs,
  };
}

export async function readTableCell(
  filePath: string,
  blockIndex: number,
  rowIndex: number,
  colIndex: number,
): Promise<string> {
  const r = await readTableCellStructured(filePath, blockIndex, rowIndex, colIndex);
  let out = `Cell [${r.rowIndex},${r.colIndex}] of the table at block ${r.blockIndex} in ${r.file}`;
  if (r.gridSpan > 1) out += ` (gridSpan ${r.gridSpan})`;
  if (r.vMerge) out += ` (vMerge: ${r.vMerge})`;
  out += `: ${r.paragraphs.length} paragraph(s).\n\n`;
  for (const p of r.paragraphs) {
    let prefix = `(p${p.paragraphIndex})`;
    if (p.style && p.style !== "Normal") prefix += ` [${p.style}]`;
    if (p.numId !== undefined) prefix += ` [num ${p.numId}/${p.numLevel ?? 0}]`;
    out += `${prefix} ${p.text}\n`;
  }
  out += "\n" + jsonBlock(r);
  return out;
}

// ---------------------------------------------------------------------------
// 3e. get_paragraph_format (structured)
// ---------------------------------------------------------------------------

/** Read the alignment of a paragraph, normalized to left/center/right/justify. */
function readAlignment(pChildren: XNode[]): string | undefined {
  const pPr = findOne(pChildren, "w:pPr");
  if (!pPr) return undefined;
  const jc = findOne(pPr["w:pPr"] as XNode[], "w:jc");
  if (!jc) return undefined;
  const val = attr(jc, "w:val");
  if (!val) return undefined;
  return val === "both" ? "justify" : val;
}

/** Read numbering (w:numPr) of a paragraph. */
function readNumbering(pChildren: XNode[]): { numId: number; numLevel: number } | undefined {
  const pPr = findOne(pChildren, "w:pPr");
  if (!pPr) return undefined;
  const numPr = findOne(pPr["w:pPr"] as XNode[], "w:numPr");
  if (!numPr) return undefined;
  const numPrChildren = numPr["w:numPr"] as XNode[];
  const numIdEl = findOne(numPrChildren, "w:numId");
  if (!numIdEl) return undefined;
  const numId = parseInt(attr(numIdEl, "w:val") ?? "", 10);
  if (!Number.isFinite(numId)) return undefined;
  const ilvlEl = findOne(numPrChildren, "w:ilvl");
  const numLevel = ilvlEl ? parseInt(attr(ilvlEl, "w:val") ?? "0", 10) : 0;
  return { numId, numLevel: Number.isFinite(numLevel) ? numLevel : 0 };
}

function intAttr(node: XNode | undefined, name: string): number | undefined {
  if (!node) return undefined;
  const raw = attr(node, name);
  if (raw === undefined) return undefined;
  const n = parseInt(raw, 10);
  return Number.isFinite(n) ? n : undefined;
}

export interface ParagraphFormatResult {
  file: string;
  paragraphIndex: number;
  /** Style id, or null for an unstyled (default) paragraph — always emitted (L3). */
  style: string | null;
  /** Heading level, or null when the paragraph is not a heading — always emitted (L3). */
  headingLevel: number | null;
  /** Normalized: left | center | right | justify. Defaults to "left", always emitted (L3). */
  alignment: string;
  numId?: number;
  numLevel?: number;
  /** Indentation in twips (1440 = 1 inch), as accepted by set_paragraph_formats. */
  indentLeft?: number;
  indentRight?: number;
  firstLineIndent?: number;
  hangingIndent?: number;
  /** Spacing before/after in points (as accepted by set_paragraph_formats). */
  spaceBefore?: number;
  spaceAfter?: number;
  /** Line spacing: raw w:line value plus rule. spaceLine is points when the rule is "exact". */
  lineRule?: string;
  lineValue?: number;
  spaceLine?: number;
}

export async function getParagraphFormatStructured(
  filePath: string,
  paragraphIndex: number,
): Promise<ParagraphFormatResult> {
  const handle = await openDocx(filePath);
  const parsed = await parseDocXml(handle);
  const body = getBody(parsed);
  const refs = enumerateBlockRefs(body);

  const ref = resolveBlockRef(refs, paragraphIndex);
  if (!ref.element["w:p"]) {
    throw new EngineError(ErrorCode.NOT_A_PARAGRAPH, `Block ${paragraphIndex} is not a paragraph (it may be a table).`);
  }

  const pChildren = ref.element["w:p"] as XNode[];
  const style = getParagraphStyle(pChildren);
  const result: ParagraphFormatResult = {
    file: path.basename(filePath),
    paragraphIndex,
    // Resolve explicit defaults so the serialized JSON matches the human text
    // ("style default / alignment left (default)") instead of dropping the keys
    // when the readers return undefined for an unstyled paragraph (L3).
    style: style ?? null,
    headingLevel: getHeadingLevel(style) ?? null,
    alignment: readAlignment(pChildren) ?? "left",
  };

  const num = readNumbering(pChildren);
  if (num) {
    result.numId = num.numId;
    result.numLevel = num.numLevel;
  }

  const pPr = findOne(pChildren, "w:pPr");
  if (pPr) {
    const props = pPr["w:pPr"] as XNode[];
    const ind = findOne(props, "w:ind");
    result.indentLeft = intAttr(ind, "w:left");
    result.indentRight = intAttr(ind, "w:right");
    result.firstLineIndent = intAttr(ind, "w:firstLine");
    result.hangingIndent = intAttr(ind, "w:hanging");

    const spacing = findOne(props, "w:spacing");
    if (spacing) {
      const before = intAttr(spacing, "w:before");
      const after = intAttr(spacing, "w:after");
      if (before !== undefined) result.spaceBefore = before / 20;
      if (after !== undefined) result.spaceAfter = after / 20;
      const line = intAttr(spacing, "w:line");
      const rule = attr(spacing, "w:lineRule");
      if (line !== undefined) {
        result.lineValue = line;
        if (rule) result.lineRule = rule;
        if (rule === "exact" || rule === "atLeast") result.spaceLine = line / 20;
      }
    }
  }

  return result;
}

export async function getParagraphFormat(filePath: string, paragraphIndex: number): Promise<string> {
  const r = await getParagraphFormatStructured(filePath, paragraphIndex);
  const lines: string[] = [`Format of paragraph ${r.paragraphIndex} in ${r.file}:`];
  lines.push(`  style: ${r.style ?? "(default)"}${r.headingLevel ? ` (Heading ${r.headingLevel})` : ""}`);
  lines.push(`  alignment: ${r.alignment ?? "left (default)"}`);
  if (r.numId !== undefined) lines.push(`  numbering: numId ${r.numId}, level ${r.numLevel ?? 0}`);
  const indParts: string[] = [];
  if (r.indentLeft !== undefined) indParts.push(`left ${r.indentLeft}`);
  if (r.indentRight !== undefined) indParts.push(`right ${r.indentRight}`);
  if (r.firstLineIndent !== undefined) indParts.push(`firstLine ${r.firstLineIndent}`);
  if (r.hangingIndent !== undefined) indParts.push(`hanging ${r.hangingIndent}`);
  if (indParts.length) lines.push(`  indent (twips): ${indParts.join(", ")}`);
  const spParts: string[] = [];
  if (r.spaceBefore !== undefined) spParts.push(`before ${r.spaceBefore}pt`);
  if (r.spaceAfter !== undefined) spParts.push(`after ${r.spaceAfter}pt`);
  if (r.lineValue !== undefined) spParts.push(`line ${r.lineValue}${r.lineRule ? ` (${r.lineRule})` : ""}`);
  if (spParts.length) lines.push(`  spacing: ${spParts.join(", ")}`);
  return lines.join("\n") + "\n\n" + jsonBlock(r);
}

// ---------------------------------------------------------------------------
// 4. replace_texts (bulk find/replace)
// ---------------------------------------------------------------------------

export interface ReplaceTextItem {
  search: string;
  replace: string;
  caseSensitive?: boolean;
}

/**
 * Detect whether two non-empty strings share any positional overlap that
 * could cause one to land on or across the other in a document. Used by
 * the tracked-mode overlap guard in `replaceTexts`. Returns true if:
 *   - either string fully contains the other, OR
 *   - a non-empty prefix of `needle` equals a suffix of `haystack`, OR
 *   - a non-empty suffix of `needle` equals a prefix of `haystack`.
 * Caller must ensure both inputs are non-empty.
 */
function textsOverlap(needle: string, haystack: string): boolean {
  if (haystack.includes(needle)) return true;
  if (needle.includes(haystack)) return true;
  const maxLen = Math.min(needle.length, haystack.length) - 1;
  for (let len = 1; len <= maxLen; len++) {
    if (haystack.endsWith(needle.slice(0, len))) return true;
    if (haystack.startsWith(needle.slice(needle.length - len))) return true;
  }
  return false;
}

/**
 * Apply multiple find/replace operations in a single open/parse/save cycle.
 *
 * Items are applied sequentially in array order. With trackChanges=true a
 * single revision context is shared across all items so revision IDs stay
 * monotonic.
 */
export async function replaceTexts(
  filePath: string,
  items: ReplaceTextItem[],
  trackChanges: boolean = true,
  author: string = "Claude",
  includeHeadersFooters: boolean = false,
): Promise<string> {
  if (items.length === 0) {
    return `No replacements to apply.`;
  }

  for (let i = 0; i < items.length; i++) {
    if (items[i].search.length === 0) {
      throw new EngineError(
        ErrorCode.INVALID_PARAMETER,
        `items[${i}].search is empty; replacement would not terminate.`,
      );
    }
  }

  // Conservative tracked-mode overlap guard: if item N's search and an
  // earlier item M's replace share any non-empty substring (full
  // containment OR boundary overlap), item N's matcher could land on or
  // across M's w:ins, producing nested w:ins/w:del markup that does not
  // round-trip through rejectAllChanges. Untracked mode is unaffected
  // because replacements happen in plain text. This is a static check
  // on the items themselves, not the document — it errs on the side of
  // refusing risky combinations even when the document content might
  // not actually trigger the failure.
  if (trackChanges) {
    for (let i = 1; i < items.length; i++) {
      const cs = items[i].caseSensitive ?? false;
      const needle = cs ? items[i].search : items[i].search.toLowerCase();
      for (let j = 0; j < i; j++) {
        // Skip items with empty replace (delete operations) — they produce
        // no inserted text, so a later search cannot land on them.
        if (items[j].replace.length === 0) continue;
        const earlierReplace = cs ? items[j].replace : items[j].replace.toLowerCase();
        if (textsOverlap(needle, earlierReplace)) {
          throw new EngineError(
            ErrorCode.INVALID_PARAMETER,
            `items[${i}].search ("${items[i].search}") could match text produced by items[${j}].replace ("${items[j].replace}"). ` +
              `replace_texts conservatively refuses this under track_changes=true even when item ${j} produces no matches in practice, ` +
              `because tracked sequential replacement cannot safely chain overlapping items: nested w:ins/w:del does not round-trip through rejectAllChanges. ` +
              `Issue separate replace_texts calls (one per item) instead, or set track_changes: false.`,
          );
        }
      }
    }
  }

  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);
    const parsed = await parseDocXml(handle);
    const body = getBody(parsed);

    // Predicate: does any item match this paragraph's text?
    const itemsMatchParagraph = (pChildren: XNode[]): boolean => {
      const text = extractParagraphText(pChildren);
      return items.some((it) => {
        const cs = it.caseSensitive ?? false;
        const haystack = cs ? text : text.toLowerCase();
        const needle = cs ? it.search : it.search.toLowerCase();
        return haystack.includes(needle);
      });
    };

    // Parse header/footer files upfront so the pending-revisions guard
    // and the later write step share one parse pass per file.
    interface HfPart { name: string; parsed: XNode[]; children: XNode[]; }
    const hfParts: HfPart[] = [];
    if (includeHeadersFooters) {
      for (const hfFile of getHeaderFooterFiles(handle)) {
        const hfXml = await handle.zip.file(hfFile)?.async("string");
        if (!hfXml) continue;
        const hfParsed: XNode[] = parser.parse(hfXml);
        const rootEl = hfParsed.find((n: XNode) => n["w:hdr"] || n["w:ftr"]);
        if (!rootEl) continue;
        const hfChildren = rootEl["w:hdr"] ?? rootEl["w:ftr"];
        hfParts.push({ name: hfFile, parsed: hfParsed, children: hfChildren });
      }
    }

    if (trackChanges) {
      const conflict = findRevisionConflict(body, (pChildren) => {
        return paragraphHasRevisions(pChildren) && itemsMatchParagraph(pChildren);
      });
      if (conflict) throwPendingRevisions(conflict.description);
      // Same check for header/footer paragraphs.
      for (const hf of hfParts) {
        const hfConflict = findRevisionConflict(hf.children, (pChildren) => {
          return paragraphHasRevisions(pChildren) && itemsMatchParagraph(pChildren);
        });
        if (hfConflict) throwPendingRevisions(`A paragraph in ${hf.name}`);
      }
    }

    const counts: number[] = items.map(() => 0);
    const ctx = trackChanges
      ? newRevisionContext((await scanMaxIdAcrossParts(handle, parsed)) + 1, author)
      : null;

    const applyToChildren = (
      children: XNode[],
      revCtx: RevisionContext | null,
    ): void => {
      for (let i = 0; i < items.length; i++) {
        const it = items[i];
        const cs = it.caseSensitive ?? false;
        for (const child of children) {
          if (child["w:p"]) {
            counts[i] += revCtx
              ? replaceInParagraphTracked(child["w:p"], it.search, it.replace, cs, revCtx)
              : replaceInParagraph(child["w:p"], it.search, it.replace, cs);
          } else if (child["w:tbl"]) {
            forEachParagraphInTable(child["w:tbl"], (pChildren) => {
              counts[i] += revCtx
                ? replaceInParagraphTracked(pChildren, it.search, it.replace, cs, revCtx)
                : replaceInParagraph(pChildren, it.search, it.replace, cs);
            });
          }
        }
      }
    };

    applyToChildren(body, ctx);

    if (includeHeadersFooters) {
      let hfMaxId = ctx ? ctx.nextId - 1 : 0;
      for (const hf of hfParts) {
        const ctx2 = trackChanges ? newRevisionContext(hfMaxId + 1, author) : null;
        applyToChildren(hf.children, ctx2);
        if (ctx2) hfMaxId = ctx2.nextId - 1;
        handle.zip.file(hf.name, builder.build(hf.parsed));
      }
    }

    const total = counts.reduce((s, n) => s + n, 0);
    if (total === 0) {
      return `No occurrences found for any of the ${items.length} search term(s) in ${path.basename(filePath)}.`;
    }

    serializeDocXml(handle, parsed);
    await saveDocx(handle);

    const mode = trackChanges ? " (tracked)" : "";
    const breakdown = items
      .map((it, i) => `  ${i + 1}. "${it.search}" → "${it.replace}": ${counts[i]}`)
      .join("\n");
    return `Replaced ${total} occurrence(s) across ${items.length} item(s) in ${path.basename(filePath)}${mode}.\n${breakdown}`;
  });
}

// ---------------------------------------------------------------------------
// Shared helper: replace paragraph text content in-place
// ---------------------------------------------------------------------------

/**
 * True if a run (the children array of a <w:r>) carries inline content that a
 * text replacement must PRESERVE rather than rewrite: a drawing (image), a
 * footnote/endnote reference, or a field code (begin/separate/end fldChar or
 * the field instruction text). Such runs are kept verbatim and the new text is
 * placed around them; dropping them would orphan footnote/endnote bodies in
 * footnotes.xml/endnotes.xml and break cross-references / page-number fields
 * (N4). Their own text (extractRunText) is therefore EXCLUDED from the
 * old-text diff baseline, exactly as drawing runs already were.
 */
function runIsStructural(runC: XNode[]): boolean {
  return runC.some(
    (rc) =>
      rc["w:drawing"] !== undefined ||
      rc["w:footnoteReference"] !== undefined ||
      rc["w:endnoteReference"] !== undefined ||
      rc["w:fldChar"] !== undefined ||
      rc["w:instrText"] !== undefined,
  );
}

/**
 * True if a paragraph child must be preserved across a text replacement:
 * bookmark/comment range markers, a <w:fldSimple> field, or a <w:r> for which
 * runIsStructural() holds. Mirrors the inline preservation logic so all
 * replacement paths agree (N4). NOTE: a <w:hyperlink> is deliberately NOT
 * structural — it is a TEXT container (its visible text is part of the
 * paragraph's editable text), so preserving it opaquely while inserting the
 * replacement would duplicate the old link text. Editing a paragraph replaces
 * a hyperlink's text like any other run (the long-standing behaviour).
 */
function isStructuralParagraphChild(child: XNode): boolean {
  if (
    child["w:bookmarkStart"] !== undefined ||
    child["w:bookmarkEnd"] !== undefined ||
    child["w:commentRangeStart"] !== undefined ||
    child["w:commentRangeEnd"] !== undefined ||
    child["w:commentReference"] !== undefined ||
    child["w:fldSimple"] !== undefined
  ) {
    return true;
  }
  if (child["w:r"]) {
    return runIsStructural(child["w:r"] as XNode[]);
  }
  return false;
}

/**
 * Yield a paragraph's inline children, descending TRANSPARENTLY through
 * <w:hyperlink>/<w:smartTag> wrappers. These are TEXT containers, not opaque
 * structural elements: their inner runs' text is part of the paragraph's
 * editable text (so it enters the tracked diff baseline and gets replaced),
 * and any structural run inside them (footnote/endnote/field) must still be
 * collected. The wrapper itself is dropped on a text replacement.
 */
function* iterInlineChildren(children: XNode[]): Generator<XNode> {
  for (const child of children) {
    if (child["w:hyperlink"] !== undefined || child["w:smartTag"] !== undefined) {
      const tag = child["w:hyperlink"] !== undefined ? "w:hyperlink" : "w:smartTag";
      yield* iterInlineChildren(child[tag] as XNode[]);
    } else {
      yield child;
    }
  }
}

/** Replace the text content of a w:p element, preserving pPr, structural elements, and optionally tracking changes. */
function replaceParagraphText(
  element: XNode,
  newText: string,
  ctx: RevisionContext | null,
): void {
  // Normalize CRLF / lone CR up-front so no stray "\r" reaches a run and every
  // "\n" is handled consistently (soft break) by both the tracked diff path and
  // the untracked single-paragraph rebuild below.
  newText = normalizeNewlines(newText);
  const pChildren = element["w:p"] as XNode[];
  const pPr = findOne(pChildren, "w:pPr");

  // Collect structural elements that must be preserved (bookmarks, comment
  // ranges, drawings, footnote/endnote references, field codes, hyperlinks).
  const structuralElements: XNode[] = [];
  for (const child of iterInlineChildren(pChildren)) {
    if (isStructuralParagraphChild(child)) structuralElements.push(child);
  }

  if (ctx) {
    // Extract current full text and first rPr from existing runs
    let oldText = "";
    let firstRPr: XNode | null = null;
    for (const child of iterInlineChildren(pChildren)) {
      if (child["w:r"]) {
        const runC = child["w:r"] as XNode[];
        // Structural runs (drawings, footnote/endnote refs, field codes) are
        // preserved verbatim, so their text must NOT enter the diff baseline.
        if (runIsStructural(runC)) continue;
        const rPr = getRunRPr(runC);
        if (!firstRPr && rPr) firstRPr = rPr;
        oldText += extractRunText(runC);
      } else if (child["w:ins"]) {
        for (const insChild of child["w:ins"]) {
          if (insChild["w:r"]) {
            const runC = insChild["w:r"] as XNode[];
            if (runIsStructural(runC)) continue;
            const rPr = getRunRPr(runC);
            if (!firstRPr && rPr) firstRPr = rPr;
            oldText += extractRunText(runC);
          }
        }
      }
    }

    const diff = computeMinimalDiff(oldText, newText);
    const newChildren: XNode[] = [];
    if (pPr) newChildren.push(pPr);
    newChildren.push(...structuralElements);
    // Newlines in the inserted text become soft breaks (w:br) inside the
    // tracked run. Tracked mode keeps a single paragraph because inserting a
    // paragraph mark as a tracked change does not round-trip cleanly through
    // accept/reject in this engine. Untracked edits split into real paragraphs
    // (see replaceParagraphTextMultiline).
    if (diff.prefix) newChildren.push(...makeTextRuns(diff.prefix, firstRPr));
    if (diff.oldMiddle) newChildren.push(wrapInDel(makeDelTextRuns(diff.oldMiddle, firstRPr), ctx));
    if (diff.newMiddle) newChildren.push(wrapInIns(makeTextRuns(diff.newMiddle, firstRPr), ctx));
    if (diff.suffix) newChildren.push(...makeTextRuns(diff.suffix, firstRPr));
    element["w:p"] = newChildren;
  } else {
    // Untracked, single-paragraph rebuild (used by editTableParagraphs and the
    // tracked-cell fallthrough). A "\n" becomes a <w:br/> soft break — NOT a
    // paragraph split — because this editor addresses one <w:p> within a cell
    // by a cell-local index; splitting would change the cell's paragraph count.
    // This mirrors the tracked branch's makeTextRuns behavior so the two modes
    // agree (was: the whole newText wrapped verbatim in one <w:t>, leaving a
    // literal LF that Word renders as whitespace).
    const firstRPr = firstRunRPr(pChildren);
    const newChildren: XNode[] = [];
    if (pPr) newChildren.push(pPr);
    newChildren.push(...structuralElements);
    newChildren.push(...makeTextRuns(newText, firstRPr));
    element["w:p"] = newChildren;
  }
}

/**
 * Collect the inline structural elements (bookmarks, comment range/reference
 * markers, drawing-bearing runs) of a paragraph that must survive a text
 * replacement. Mirrors the preservation logic in replaceParagraphText.
 */
function collectStructuralElements(pChildren: XNode[]): XNode[] {
  const out: XNode[] = [];
  for (const child of iterInlineChildren(pChildren)) {
    if (isStructuralParagraphChild(child)) out.push(child);
  }
  return out;
}

/** Return the rPr of the first non-structural run in a paragraph (skipping
 * drawing/footnote/endnote/field runs, whose rPr is not body-text styling). */
function firstRunRPr(pChildren: XNode[]): XNode | null {
  for (const child of iterInlineChildren(pChildren)) {
    if (child["w:r"]) {
      const runC = child["w:r"] as XNode[];
      if (runIsStructural(runC)) continue;
      const rPr = getRunRPr(runC);
      if (rPr) return rPr;
    }
  }
  return null;
}

/**
 * Replace a paragraph's text, treating "\n" as a paragraph break.
 *
 *  - Untracked: split into multiple sibling <w:p>, each sharing the source
 *    pPr (deep-cloned for lines 2..N) and the first run's rPr, so list
 *    numbering, hanging indents and alignment apply to every line — matching
 *    what a human types pressing Enter. Empty lines become empty paragraphs.
 *  - Tracked, or text without "\n": delegate to replaceParagraphText, which
 *    keeps a single paragraph and (in tracked mode) renders "\n" as a soft
 *    break. Tracked paragraph-mark insertion is avoided because it does not
 *    round-trip cleanly through accept/reject.
 *
 * `parent` is the array holding `element` (document body or a table cell's
 * children); on a multi-paragraph split it is spliced in place.
 */
function replaceParagraphTextMultiline(
  parent: XNode[],
  element: XNode,
  newText: string,
  ctx: RevisionContext | null,
): void {
  // Normalize CRLF / lone CR first so the "\n" split decision below sees a
  // bare "\r" (no "\n") as a line break too, not a stray carriage return.
  newText = normalizeNewlines(newText);
  if (ctx || !newText.includes("\n")) {
    replaceParagraphText(element, newText, ctx);
    return;
  }

  const idx = parent.indexOf(element);
  if (idx === -1) {
    // Element is not where we expect — fall back to in-place replacement.
    replaceParagraphText(element, newText, ctx);
    return;
  }

  const pChildren = element["w:p"] as XNode[];
  const pPr = findOne(pChildren, "w:pPr");
  const firstRPr = firstRunRPr(pChildren);
  const structuralElements = collectStructuralElements(pChildren);

  const lines = newText.split("\n");
  const newParas: XNode[] = [];
  for (let i = 0; i < lines.length; i++) {
    const kids: XNode[] = [];
    if (pPr) kids.push(i === 0 ? pPr : cloneNode(pPr));
    if (i === 0) kids.push(...structuralElements);
    if (lines[i]) kids.push(makeTextRun(lines[i], firstRPr));
    newParas.push(el("w:p", kids));
  }

  // Carry the original paragraph's w14:paraId onto the FIRST produced line so a
  // captured/seeded anchor keeps resolving after the split (M6). The other lines
  // stay un-anchored, preserving paraId uniqueness (ensure_anchors seeds them
  // later if needed). paraId lives on the <w:p> element, not inside pPr.
  const originalParaId = attr(element, "w14:paraId");
  if (originalParaId) setAttr(newParas[0], "w14:paraId", originalParaId);

  parent.splice(idx, 1, ...newParas);
}

/**
 * ECMA-376 (CT_Tc) requires a table cell's content to END with a block-level
 * <w:p>; a nested <w:tbl> may not be the last block child (Word always keeps a
 * terminating paragraph after a nested table). After any rebuild that may drop
 * the trailing paragraph, call this to restore the invariant. The "cell has ≥1
 * w:p" check elsewhere is weaker — it passes for [w:p, w:tbl], which is invalid.
 */
function ensureCellEndsWithParagraph(cellChildren: XNode[]): void {
  let lastBlock: XNode | undefined;
  for (const child of cellChildren) {
    if (child["w:p"] || child["w:tbl"]) lastBlock = child;
  }
  if (lastBlock && lastBlock["w:tbl"]) {
    cellChildren.push(el("w:p", []));
  }
}

/**
 * Replace ALL block-level paragraphs of a table cell with paragraphs built from
 * newText (split on "\n" so each line is its own <w:p>). The new paragraphs
 * inherit the cell's first paragraph's pPr (deep-cloned for lines 2..N) and its
 * first run's rPr, and the first line keeps any inline structural elements
 * (bookmarks, comment markers, drawings). Non-paragraph children (w:tcPr, a
 * nested w:tbl) are preserved in place.
 *
 * This makes an untracked cell edit a full replacement rather than a first-
 * paragraph-only edit, so re-editing a cell that already holds several
 * paragraphs (e.g. a numbered list produced by an earlier split) does not leave
 * stale lines behind. A cell always keeps at least one paragraph (an empty
 * newText yields a single empty <w:p>), as OOXML requires.
 */
function replaceCellContentUntracked(cellChildren: XNode[], newText: string): void {
  // Normalize CRLF / lone CR so each line splits into a real <w:p> with no
  // stray "\r" left inside a run.
  newText = normalizeNewlines(newText);
  const firstPara = cellChildren.find((c: XNode) => c["w:p"]);
  const firstPChildren = firstPara ? (firstPara["w:p"] as XNode[]) : null;
  const pPr = firstPChildren ? findOne(firstPChildren, "w:pPr") : null;
  const firstRPr = firstPChildren ? firstRunRPr(firstPChildren) : null;
  const structuralElements = firstPChildren ? collectStructuralElements(firstPChildren) : [];

  const lines = newText.split("\n");
  const newParas: XNode[] = lines.map((line, i) => {
    const kids: XNode[] = [];
    if (pPr) kids.push(i === 0 ? pPr : cloneNode(pPr));
    if (i === 0) kids.push(...structuralElements);
    if (line) kids.push(makeTextRun(line, firstRPr));
    return el("w:p", kids);
  });

  const newChildren: XNode[] = [];
  let inserted = false;
  for (const child of cellChildren) {
    if (child["w:p"]) {
      if (!inserted) {
        newChildren.push(...newParas);
        inserted = true;
      }
    } else {
      newChildren.push(child);
    }
  }
  if (!inserted) newChildren.push(...newParas);
  cellChildren.length = 0;
  cellChildren.push(...newChildren);

  // The new paragraphs replaced the FIRST w:p block and any other w:p children
  // were dropped — including a paragraph that terminated the cell after a nested
  // table. Restore the "cell ends with w:p" invariant if it now ends in a table.
  ensureCellEndsWithParagraph(cellChildren);
}

// ---------------------------------------------------------------------------
// 5. edit_paragraphs
// ---------------------------------------------------------------------------

export interface EditParagraphItem {
  /** Exactly one of paragraphIndex / anchor must be set. */
  paragraphIndex?: number;
  anchor?: string;
  newText: string;
}

/** Human-readable description of a locator for error messages. */
function describeLocator(loc: ParagraphLocator): string {
  return loc.anchor !== undefined && loc.anchor !== ""
    ? `anchor "${loc.anchor}"`
    : `at index ${loc.paragraphIndex}`;
}

export async function editParagraphs(
  filePath: string,
  edits: EditParagraphItem[],
  trackChanges: boolean = true,
  author: string = "Claude",
): Promise<string> {
  if (edits.length === 0) {
    return `No edits to apply.`;
  }

  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);
    const parsed = await parseDocXml(handle);
    const root = getDocumentRoot(parsed);
    const body = getBody(parsed);
    const refs = enumerateBlockRefs(body);
    const anchorIndex = buildAnchorIndex(body);

    // Resolve and validate every locator (index or anchor) upfront. Element
    // references (not indices) are held because an untracked multi-line edit
    // splices new sibling paragraphs into the body, shifting later indices.
    // `directBody` records whether each target is a DIRECT-body paragraph; only
    // those are in anchor scope (v1), so only those auto-seed (M-rev). The
    // resolved container is also kept per target so the untracked multi-line
    // SPLIT path can refuse a paragraph that does not live directly in the body
    // (an SDT-inner paragraph cannot be split into sibling <w:p> — see F2 below).
    const targets: XNode[] = [];
    const directBody = new Set<XNode>();
    const containerOf = new Map<XNode, XNode[]>();
    for (const edit of edits) {
      const loc = locatorToLocation(refs, anchorIndex, edit);
      if (trackChanges && paragraphHasRevisions(loc.element["w:p"] as XNode[])) {
        throwPendingRevisions(`Paragraph ${describeLocator(edit)}`);
      }
      targets.push(loc.element);
      containerOf.set(loc.element, loc.parent);
      if (loc.parent === body) directBody.add(loc.element);
    }

    const ctx = trackChanges
      ? newRevisionContext((await scanMaxIdAcrossParts(handle, parsed)) + 1, author)
      : null;

    // Collapse duplicate edits to the same paragraph (last-write-wins). An
    // untracked multi-line edit splices the original element out of the body,
    // so applying once with the final text avoids mutating a detached node.
    // Seed the touched paragraph's anchor before the (possibly splitting) edit.
    const seeder = new AnchorSeeder(root, body);
    const finalText = new Map<XNode, string>();
    for (let i = 0; i < edits.length; i++) {
      finalText.set(targets[i], edits[i].newText);
    }
    for (const [element, newText] of finalText) {
      // F2: an UNTRACKED MULTI-LINE edit splits the paragraph into sibling <w:p>
      // (replaceParagraphTextMultiline splices into the container). That is only
      // valid when the paragraph lives directly in the body — for a paragraph
      // inside a content control (w:sdt) the container is the SDT's sdtContent,
      // and a structural paragraph add/remove there is refused by design (C1).
      // Without this guard the split would silently degrade to a single <w:p>
      // with <w:br/> soft breaks (wrong container, inconsistent + silent). Be
      // explicit: refuse, and tell the caller how to proceed. A single-line
      // untracked edit (no "\n") still edits in place; tracked edits use soft
      // breaks (no split) and are unaffected; a direct-body multi-line edit
      // still splits.
      const isMultiline = !ctx && normalizeNewlines(newText).includes("\n");
      if (isMultiline && containerOf.get(element) !== body) {
        throw new EngineError(
          ErrorCode.INVALID_LOCATOR,
          "cannot split a paragraph inside a content control (w:sdt); edit it as a single line, or by index in place",
        );
      }
      // Touched-paragraph auto-seed (kept by in-place single-line edits) — only
      // for direct-body paragraphs. A paragraph inside a top-level w:sdt is out
      // of anchor scope, so seeding it would write an unreachable, scope-
      // violating w14:paraId onto the SDT-contained <w:p>.
      if (directBody.has(element)) seeder.seed(element);
      replaceParagraphTextMultiline(body, element, newText, ctx);
    }

    serializeDocXml(handle, parsed);
    await saveDocx(handle);

    const mode = trackChanges ? " (tracked)" : "";
    return `Updated ${finalText.size} paragraph(s) in ${path.basename(filePath)}${mode}.`;
  });
}

// ---------------------------------------------------------------------------
// Shared helper: build a new paragraph element
// ---------------------------------------------------------------------------

/** Options for building a new paragraph beyond text and style. */
interface BuildParagraphOptions {
  numId?: number;
  numLevel?: number;
  /** Deep-cloned pPr from a source paragraph (copy_format_from). Overrides style/numId/numLevel. */
  sourcePPr?: XNode;
}

/**
 * Build one or more w:p elements with optional style, numbering, copied format,
 * and tracked insertion context.
 *
 * Untracked multi-line input becomes one real paragraph per line, each with its
 * own pPr (so numbering / hanging indents apply to every line — matching the
 * edit semantics in replaceParagraphTextMultiline). Tracked mode, and single-
 * line input, produce a single paragraph; under tracking "\n" renders as a soft
 * break because tracked paragraph-mark insertion does not round-trip cleanly
 * through accept/reject.
 */
function buildNewParagraph(
  text: string,
  style: string | undefined,
  ctx: RevisionContext | null,
  opts?: BuildParagraphOptions,
): XNode[] {
  // Normalize CRLF / lone CR so the "\n" handling below (untracked split /
  // tracked w:br) sees clean newlines and no stray "\r" reaches a run.
  text = normalizeNewlines(text);
  if (!ctx && text.includes("\n")) {
    return text.split("\n").map((line) => {
      const pChildren: XNode[] = [];
      const pPr = buildPPrForNewParagraph(style, ctx, opts);
      if (pPr) pChildren.push(pPr);
      if (line) {
        pChildren.push(
          el("w:r", [el("w:t", [textNode(line)], { "xml:space": "preserve" })]),
        );
      }
      return el("w:p", pChildren);
    });
  }

  const pChildren: XNode[] = [];

  // Build the pPr element
  const pPr = buildPPrForNewParagraph(style, ctx, opts);
  if (pPr) pChildren.push(pPr);

  // Build text runs
  const lines = text.split("\n");
  if (ctx) {
    const insRuns: XNode[] = [];
    for (let i = 0; i < lines.length; i++) {
      if (i > 0) insRuns.push(el("w:r", [el("w:br")]));
      if (lines[i]) {
        insRuns.push(
          el("w:r", [
            el("w:t", [textNode(lines[i])], { "xml:space": "preserve" }),
          ]),
        );
      }
    }
    pChildren.push(wrapInIns(insRuns, ctx));
  } else if (text) {
    pChildren.push(
      el("w:r", [el("w:t", [textNode(text)], { "xml:space": "preserve" })]),
    );
  }

  return [el("w:p", pChildren)];
}

/**
 * Build w:pPr for a new paragraph.
 * If sourcePPr is provided (copy_format_from), deep-clone it and inject tracked change marker.
 * Otherwise build from style + numId/numLevel.
 */
function buildPPrForNewParagraph(
  style: string | undefined,
  ctx: RevisionContext | null,
  opts?: BuildParagraphOptions,
): XNode | null {
  if (opts?.sourcePPr) {
    // Deep clone the source pPr
    const pPr = cloneNode(opts.sourcePPr);
    const pPrChildren = pPr["w:pPr"] as XNode[];

    // Strip stale tracked-change metadata that belongs to the SOURCE paragraph,
    // not the new one. A leftover w:pPrChange / w:rPrChange would otherwise be
    // read by accept/reject as a live revision on the freshly inserted node.
    for (let i = pPrChildren.length - 1; i >= 0; i--) {
      if (pPrChildren[i]["w:pPrChange"] !== undefined) pPrChildren.splice(i, 1);
    }

    // Always strip stale revision markers from the cloned rPr —
    // they belong to the source paragraph, not the new one.
    const existingRPr = findOne(pPrChildren, "w:rPr");
    if (existingRPr) {
      const rPrChildren = existingRPr["w:rPr"] as XNode[];
      for (let i = rPrChildren.length - 1; i >= 0; i--) {
        if (
          rPrChildren[i]["w:ins"] !== undefined ||
          rPrChildren[i]["w:del"] !== undefined ||
          rPrChildren[i]["w:rPrChange"] !== undefined
        ) {
          rPrChildren.splice(i, 1);
        }
      }
      // Clean up empty rPr
      if (rPrChildren.length === 0) {
        const rPrIdx = pPrChildren.indexOf(existingRPr);
        if (rPrIdx !== -1) pPrChildren.splice(rPrIdx, 1);
      }
    }

    if (ctx) {
      // Add our own w:ins marker for the new paragraph
      let rPr = findOne(pPrChildren, "w:rPr");
      if (!rPr) {
        rPr = el("w:rPr");
        pPrChildren.push(rPr);
      }
      (rPr["w:rPr"] as XNode[]).push(
        el("w:ins", [], {
          "w:id": String(allocRevId(ctx)),
          "w:author": ctx.author,
          "w:date": ctx.date,
        }),
      );
    }
    return pPr;
  }

  // Build pPr from scratch
  const pPrChildren: XNode[] = [];

  if (style) {
    pPrChildren.push(el("w:pStyle", [], { "w:val": style }));
  }

  if (opts?.numId !== undefined) {
    pPrChildren.push(
      el("w:numPr", [
        el("w:ilvl", [], { "w:val": String(opts.numLevel ?? 0) }),
        el("w:numId", [], { "w:val": String(opts.numId) }),
      ]),
    );
  }

  if (ctx) {
    const rPrIns = el("w:ins", [], {
      "w:id": String(allocRevId(ctx)),
      "w:author": ctx.author,
      "w:date": ctx.date,
    });
    pPrChildren.push(el("w:rPr", [rPrIns]));
  }

  if (pPrChildren.length === 0) return null;
  return el("w:pPr", pPrChildren);
}

/**
 * Insert paragraph element(s) into the body before the block at `position`
 * (unified block-index space), handling sectPr and append.
 *
 * An out-of-range position (including -1) appends before sectPr — the documented
 * "append" behaviour. An in-range position that resolves to a paragraph INSIDE a
 * top-level `w:sdt` is refused with a clear error rather than silently splicing
 * into the body at the wrong place: content controls must be edited in place
 * (by index/anchor) or have their inner paragraphs targeted directly.
 */
function spliceNewParagraph(
  body: XNode[],
  refs: BlockRef[],
  position: number,
  newParas: XNode[],
): void {
  if (!Number.isInteger(position)) {
    throw new EngineError(
      ErrorCode.INDEX_OUT_OF_RANGE,
      `Insert position ${position} must be an integer block index (use -1 to append).`,
    );
  }
  if (position < 0 || position >= refs.length) {
    const sectPrIdx = body.findIndex((n: XNode) => n["w:sectPr"]);
    if (sectPrIdx !== -1) {
      body.splice(sectPrIdx, 0, ...newParas);
    } else {
      body.push(...newParas);
    }
    return;
  }
  const ref = refs[position];
  if (ref.container !== body) {
    throw new EngineError(
      ErrorCode.INVALID_LOCATOR,
      `Cannot insert a paragraph relative to block ${position}: it is inside a content control (w:sdt). Edit the content control in place by index/anchor instead.`,
    );
  }
  body.splice(ref.indexInContainer, 0, ...newParas);
}

// ---------------------------------------------------------------------------
// 6. insert_paragraphs
// ---------------------------------------------------------------------------

/** Resolve BuildParagraphOptions from user-facing parameters. */
function resolveInsertOpts(
  refs: BlockRef[],
  numId?: number,
  numLevel?: number,
  copyFormatFrom?: number,
): BuildParagraphOptions | undefined {
  if (copyFormatFrom !== undefined) {
    const ref = resolveBlockRef(refs, copyFormatFrom);
    if (!ref.element["w:p"]) {
      throw new EngineError(
        ErrorCode.NOT_A_PARAGRAPH,
        `Block ${copyFormatFrom} is not a paragraph (copy_format_from must reference a paragraph).`,
      );
    }
    const pPr = findOne(ref.element["w:p"] as XNode[], "w:pPr");
    if (pPr) {
      return { sourcePPr: pPr };
    }
    // Source paragraph has no pPr — nothing to copy
    return undefined;
  }
  if (numId !== undefined) {
    return { numId, numLevel };
  }
  return undefined;
}

export interface InsertParagraphItem {
  text: string;
  /** Index placement: insert before this block index (-1/out-of-range appends). */
  position?: number;
  /** Anchor placement: insert before/after the paragraph with this anchor. */
  anchor?: string;
  placement?: "before" | "after";
  style?: string;
  numId?: number;
  numLevel?: number;
  copyFormatFrom?: number;
  /** Alternative to copyFormatFrom: copy pPr from the paragraph with this anchor. */
  copyFormatFromAnchor?: string;
}

/** Resolve build options for one insert item (index or anchor copy_format source). */
function resolveInsertOptsForItem(
  refs: BlockRef[],
  anchorIndex: Map<string, ParagraphLocation[]>,
  item: InsertParagraphItem,
): BuildParagraphOptions | undefined {
  if (item.copyFormatFromAnchor !== undefined && item.copyFormatFromAnchor !== "") {
    const srcEl = resolveAnchor(anchorIndex, item.copyFormatFromAnchor).element;
    const pPr = findOne(srcEl["w:p"] as XNode[], "w:pPr");
    return pPr ? { sourcePPr: pPr } : undefined;
  }
  return resolveInsertOpts(refs, item.numId, item.numLevel, item.copyFormatFrom);
}

export interface InsertParagraphsResult {
  file: string;
  inserted: number;
  /** Each inserted paragraph paired with its freshly assigned anchor, in input order. */
  newParagraphs: { text: string; anchor: string }[];
}

export async function insertParagraphsStructured(
  filePath: string,
  items: InsertParagraphItem[],
  trackChanges: boolean = true,
  author: string = "Claude",
): Promise<InsertParagraphsResult> {
  if (items.length === 0) {
    return { file: path.basename(filePath), inserted: 0, newParagraphs: [] };
  }

  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);
    const parsed = await parseDocXml(handle);
    const root = getDocumentRoot(parsed);
    const body = getBody(parsed);
    const refsForResolve = enumerateBlockRefs(body);
    const anchorIndex = buildAnchorIndex(body);

    // Validate the locator on every item and resolve its anchor target / opts
    // up-front (against the original doc, before any splice).
    const anchorTarget = new Map<InsertParagraphItem, XNode>();
    const resolvedOpts = new Map<InsertParagraphItem, BuildParagraphOptions | undefined>();
    for (const item of items) {
      const hasPos = item.position !== undefined;
      const hasAnchor = item.anchor !== undefined && item.anchor !== "";
      if (hasPos === hasAnchor) {
        throw new EngineError(
          ErrorCode.INVALID_LOCATOR,
          `Each insert must specify exactly one of position or anchor.`,
        );
      }
      if (hasAnchor) {
        if (item.placement !== "before" && item.placement !== "after") {
          throw new EngineError(ErrorCode.INVALID_PARAMETER, `Anchor placement must be "before" or "after".`);
        }
        anchorTarget.set(item, resolveAnchor(anchorIndex, item.anchor as string).element);
      }
      resolvedOpts.set(item, resolveInsertOptsForItem(refsForResolve, anchorIndex, item));
    }

    const ctx = trackChanges
      ? newRevisionContext((await scanMaxIdAcrossParts(handle, parsed)) + 1, author)
      : null;

    const seeder = new AnchorSeeder(root, body);
    const newParagraphs: { item: InsertParagraphItem; element: XNode }[] = [];

    // Position-based items keep the existing descending-by-position behaviour
    // (so same-position items follow the documented reverse-of-array order).
    const positionItems = items.filter((i) => i.position !== undefined);
    const sorted = [...positionItems].sort((a, b) => {
      const posA = a.position! < 0 || a.position! >= refsForResolve.length ? Infinity : a.position!;
      const posB = b.position! < 0 || b.position! >= refsForResolve.length ? Infinity : b.position!;
      return posB - posA;
    });
    // buildNewParagraph returns one or more <w:p> (untracked multi-line input
    // splits on "\n"); the first is the representative element for the returned
    // anchor.
    for (const item of sorted) {
      const newParas = buildNewParagraph(item.text, item.style, ctx, resolvedOpts.get(item));
      const currentRefs = enumerateBlockRefs(body);
      spliceNewParagraph(body, currentRefs, item.position as number, newParas);
      newParagraphs.push({ item, element: newParas[0] });
    }

    // Anchor-based items splice relative to their (stable) target element, in
    // input order. For repeated "after" inserts against the same anchor, an
    // after-cursor advances past each newly inserted paragraph so they keep
    // array order rather than reversing.
    const afterCursor = new Map<XNode, XNode>();
    for (const item of items) {
      if (item.position !== undefined) continue;
      const target = anchorTarget.get(item) as XNode;
      const newParas = buildNewParagraph(item.text, item.style, ctx, resolvedOpts.get(item));
      let insertAt: number;
      if (item.placement === "after") {
        const ref = afterCursor.get(target) ?? target;
        insertAt = body.indexOf(ref) + 1;
        afterCursor.set(target, newParas[newParas.length - 1]);
      } else {
        insertAt = body.indexOf(target);
      }
      body.splice(insertAt, 0, ...newParas);
      newParagraphs.push({ item, element: newParas[0] });
    }

    // Seed each new paragraph so the caller gets its anchor back.
    const reported: { text: string; anchor: string }[] = [];
    for (const item of items) {
      const entry = newParagraphs.find((np) => np.item === item);
      if (entry) reported.push({ text: item.text, anchor: seeder.seed(entry.element) });
    }

    serializeDocXml(handle, parsed);
    await saveDocx(handle);

    return { file: path.basename(filePath), inserted: items.length, newParagraphs: reported };
  });
}

export async function insertParagraphs(
  filePath: string,
  items: InsertParagraphItem[],
  trackChanges: boolean = true,
  author: string = "Claude",
): Promise<string> {
  const r = await insertParagraphsStructured(filePath, items, trackChanges, author);
  if (r.inserted === 0) return `No paragraphs to insert.`;
  const mode = trackChanges ? " (tracked)" : "";
  let out = `Inserted ${r.inserted} paragraph(s) in ${r.file}${mode}.`;
  if (r.newParagraphs.length > 0) {
    out += "\n\n" + jsonBlock({ newParagraphs: r.newParagraphs });
  }
  return out;
}

// ---------------------------------------------------------------------------
// Shared helper: mark a block element (w:p or w:tbl) as tracked deletion
// ---------------------------------------------------------------------------

/** Convert all runs in a paragraph to tracked deletions. */
function markParagraphRunsAsDeleted(pChildren: XNode[], ctx: RevisionContext): void {
  // Add deletion marker on the paragraph break via pPr > rPr > w:del
  let pPr = findOne(pChildren, "w:pPr");
  if (!pPr) {
    pPr = el("w:pPr");
    pChildren.unshift(pPr);
  }
  const pPrChildren = pPr["w:pPr"] as XNode[];
  let rPrInPPr = findOne(pPrChildren, "w:rPr");
  if (!rPrInPPr) {
    rPrInPPr = el("w:rPr");
    pPrChildren.push(rPrInPPr);
  }
  (rPrInPPr["w:rPr"] as XNode[]).push(
    el("w:del", [], {
      "w:id": String(allocRevId(ctx)),
      "w:author": ctx.author,
      "w:date": ctx.date,
    }),
  );

  markChildRunsAsDeleted(pChildren, ctx);
}

/**
 * Convert the direct `w:r` runs of a children array into a single trailing
 * `<w:del>` of `<w:delText>` runs, and recurse into transparent inline wrappers
 * (`w:hyperlink` / `w:smartTag`) so the runs inside them are marked deleted too
 * — KEEPING the wrapper element itself. Without the wrapper recursion the link
 * text would survive a tracked paragraph/table delete (accepting the delete
 * would then leave orphaned hyperlink text behind).
 */
function markChildRunsAsDeleted(children: XNode[], ctx: RevisionContext): void {
  const delRuns: XNode[] = [];
  const indicesToRemove: number[] = [];
  for (let i = 0; i < children.length; i++) {
    const child = children[i];
    if (child["w:r"]) {
      const runC = child["w:r"] as XNode[];
      const rPr = getRunRPr(runC);
      const runText = extractRunText(runC);
      if (runText) {
        delRuns.push(makeDelTextRun(runText, rPr));
      }
      indicesToRemove.push(i);
    } else if (child["w:hyperlink"] || child["w:smartTag"]) {
      // Recurse INTO the wrapper, converting its runs in place; the wrapper
      // node is left in `children` so the link/smartTag survives the delete
      // with its text now marked as a tracked deletion inside it.
      const wtag = child["w:hyperlink"] !== undefined ? "w:hyperlink" : "w:smartTag";
      markChildRunsAsDeleted(child[wtag] as XNode[], ctx);
    }
  }

  // Remove original runs (reverse order to preserve indices)
  for (let i = indicesToRemove.length - 1; i >= 0; i--) {
    children.splice(indicesToRemove[i], 1);
  }

  // Append the w:del element with all deleted runs
  if (delRuns.length > 0) {
    children.push(wrapInDel(delRuns, ctx));
  }
}

/**
 * Add a row-level tracked-deletion marker (`<w:trPr><w:del/>`) to a `<w:tr>`.
 * This is what Word writes when whole rows are deleted with track-changes on:
 * accept then removes the row entirely; reject strips the marker and keeps it.
 */
function markRowAsDeleted(trChildren: XNode[], ctx: RevisionContext): void {
  let trPr = findOne(trChildren, "w:trPr");
  if (!trPr) {
    trPr = el("w:trPr");
    // CT_Row child order is <w:tblPrEx>? then <w:trPr> then cells, so insert the
    // new trPr AFTER an existing tblPrEx (findIndex === -1 → splice at 0 = unshift).
    const tblPrExIdx = trChildren.findIndex((c: XNode) => c["w:tblPrEx"]);
    trChildren.splice(tblPrExIdx + 1, 0, trPr);
  }
  const trPrChildren = trPr["w:trPr"] as XNode[];
  // Avoid a duplicate marker if one is somehow already present.
  if (!findOne(trPrChildren, "w:del")) {
    trPrChildren.push(
      el("w:del", [], {
        "w:id": String(allocRevId(ctx)),
        "w:author": ctx.author,
        "w:date": ctx.date,
      }),
    );
  }
}

/**
 * Mark every row of a table (recursing into nested tables) as a tracked
 * deletion: row-level `<w:trPr><w:del/>` so accept removes the rows/table and
 * reject restores them, AND cell-paragraph runs converted to `<w:del>` so the
 * content shows struck-through (Word's faithful representation). N8.
 */
function markTableAsDeleted(tblChildren: XNode[], ctx: RevisionContext): void {
  const rows = findAll(tblChildren, "w:tr");
  for (const row of rows) {
    const trChildren = row["w:tr"] as XNode[];
    markRowAsDeleted(trChildren, ctx);
    const cells = findAll(trChildren, "w:tc");
    for (const cell of cells) {
      const tcChildren = cell["w:tc"] as XNode[];
      for (const child of tcChildren) {
        if (child["w:p"]) {
          markParagraphRunsAsDeleted(child["w:p"] as XNode[], ctx);
        } else if (child["w:tbl"]) {
          // Recurse so a NESTED table is also fully marked (was skipped before).
          markTableAsDeleted(child["w:tbl"] as XNode[], ctx);
        }
      }
    }
  }
}

/** Mark a body-level block element (paragraph or table) as tracked deletion. */
function markBlockAsDeleted(element: XNode, ctx: RevisionContext): void {
  if (element["w:p"]) {
    markParagraphRunsAsDeleted(element["w:p"] as XNode[], ctx);
  } else if (element["w:tbl"]) {
    markTableAsDeleted(element["w:tbl"] as XNode[], ctx);
  }
}

// ---------------------------------------------------------------------------
// 7. delete_paragraphs
// ---------------------------------------------------------------------------

export async function deleteParagraphs(
  filePath: string,
  paragraphIndices: number[],
  trackChanges: boolean = true,
  author: string = "Claude",
  anchors: string[] = [],
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);
    const parsed = await parseDocXml(handle);
    const body = getBody(parsed);
    const refs = enumerateBlockRefs(body);
    const anchorIndex = buildAnchorIndex(body);

    // Resolve a deduplicated set of target elements. An index can target a
    // paragraph OR a table block; an anchor only ever targets a paragraph
    // (tables have no paraId). Resolution by reference makes deletion
    // order-independent. A block inside a top-level w:sdt is refused (deleting
    // it would mean splicing the content control's internals — edit it in place
    // instead) rather than silently no-op'ing or deleting a body block.
    const targets = new Set<XNode>();
    for (const idx of paragraphIndices) {
      const ref = resolveBlockRef(refs, idx);
      if (ref.container !== body) {
        throw new EngineError(
          ErrorCode.INVALID_LOCATOR,
          `Cannot delete block ${idx}: it is inside a content control (w:sdt). Edit the content control in place by index/anchor instead.`,
        );
      }
      targets.add(ref.element);
    }
    for (const anchor of anchors) {
      targets.add(resolveAnchor(anchorIndex, anchor).element);
    }

    const elements = [...targets];
    if (elements.length === 0) {
      return `No paragraphs to delete.`;
    }

    if (trackChanges) {
      // Reject if any targeted block already has tracked-change markup.
      for (const element of elements) {
        if (element["w:p"] && paragraphHasRevisions(element["w:p"] as XNode[])) {
          throwPendingRevisions(`Targeted paragraph`);
        } else if (element["w:tbl"]) {
          let conflict = false;
          forEachParagraphInTable(element["w:tbl"], (pChildren) => {
            if (paragraphHasRevisions(pChildren)) conflict = true;
          });
          if (conflict) throwPendingRevisions(`Targeted table`);
        }
      }

      const maxId = await scanMaxIdAcrossParts(handle, parsed);
      const ctx = newRevisionContext(maxId + 1, author);

      for (const element of elements) {
        markBlockAsDeleted(element, ctx);
      }
    } else {
      // Hard delete by reference (order-independent).
      for (const element of elements) {
        const i = body.indexOf(element);
        if (i !== -1) body.splice(i, 1);
      }
    }

    serializeDocXml(handle, parsed);
    await saveDocx(handle);

    const mode = trackChanges ? " (tracked)" : "";
    return `Deleted ${elements.length} block(s) from ${path.basename(filePath)}${mode}.`;
  });
}

// ---------------------------------------------------------------------------
// 8. format_text
// ---------------------------------------------------------------------------

export async function formatText(
  filePath: string,
  search: string,
  formatting: TextFormatting,
  caseSensitive: boolean = false,
): Promise<string> {
  // Validate font size BEFORE the lock so a negative/absurd value never reaches
  // <w:sz>/<w:szCs> (ST_HpsMeasure is unsigned half-points; Word caps at
  // 1638 pt). 0 is treated as "unset" elsewhere (the `if (fmt.fontSize)` gate),
  // so only sign/range is the gap here (N11).
  if (formatting.fontSize !== undefined && formatting.fontSize !== 0) {
    if (!Number.isFinite(formatting.fontSize) || formatting.fontSize < 0 || formatting.fontSize > 1638) {
      throw new EngineError(
        ErrorCode.INVALID_PARAMETER,
        `Font size must be between 0 and 1638 pt. Got: ${formatting.fontSize}.`,
      );
    }
  }
  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);
    const parsed = await parseDocXml(handle);
    const body = getBody(parsed);

    let totalFormatted = 0;

    for (const child of body) {
      if (child["w:p"]) {
        totalFormatted += formatInParagraph(
          child["w:p"],
          search,
          formatting,
          caseSensitive,
        );
      } else if (child["w:tbl"]) {
        forEachParagraphInTable(child["w:tbl"], (pChildren) => {
          totalFormatted += formatInParagraph(pChildren, search, formatting, caseSensitive);
        });
      }
    }

    if (totalFormatted === 0) {
      return `No occurrences of "${search}" found in ${path.basename(filePath)}.`;
    }

    serializeDocXml(handle, parsed);
    await saveDocx(handle);

    const fmtDesc = Object.entries(formatting)
      .filter(([, v]) => v !== undefined)
      .map(([k, v]) => `${k}=${v}`)
      .join(", ");
    return `Applied formatting (${fmtDesc}) to ${totalFormatted} occurrence(s) matching "${search}" in ${path.basename(filePath)}.`;
  });
}

// ---------------------------------------------------------------------------
// 9. set_paragraph_formats
// ---------------------------------------------------------------------------

export async function setParagraphFormats(
  filePath: string,
  groups: Array<{ indices?: number[]; anchors?: string[]; format: ParagraphFormat }>,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);
    const parsed = await parseDocXml(handle);
    const root = getDocumentRoot(parsed);
    const body = getBody(parsed);
    const refs = enumerateBlockRefs(body);
    const anchorIndex = buildAnchorIndex(body);

    // Resolve each group's targets (indices and/or anchors) upfront.
    // `directBody` records which targets are DIRECT-body paragraphs (anchor
    // scope, v1); only those auto-seed. Anchor-resolved targets are always
    // direct-body by construction; an index may resolve into a top-level w:sdt.
    const resolved: { elements: XNode[]; format: ParagraphFormat }[] = [];
    const directBody = new Set<XNode>();
    let totalCount = 0;
    for (const group of groups) {
      const elements: XNode[] = [];
      for (const idx of group.indices ?? []) {
        const loc = locatorToLocation(refs, anchorIndex, { paragraphIndex: idx });
        elements.push(loc.element);
        if (loc.parent === body) directBody.add(loc.element);
      }
      for (const anchor of group.anchors ?? []) {
        const el = resolveAnchor(anchorIndex, anchor).element;
        elements.push(el);
        directBody.add(el);
      }
      if (elements.length === 0) {
        throw new EngineError(ErrorCode.INVALID_LOCATOR, `Each group must specify at least one of indices or anchors.`);
      }
      resolved.push({ elements, format: group.format });
      totalCount += elements.length;
    }

    // Apply formatting and auto-seed the touched paragraphs (direct-body only).
    const seeder = new AnchorSeeder(root, body);
    for (const { elements, format } of resolved) {
      for (const element of elements) {
        applyParagraphFormat(element["w:p"], format);
        if (directBody.has(element)) seeder.seed(element);
      }
    }

    serializeDocXml(handle, parsed);
    await saveDocx(handle);

    return `Applied formatting to ${totalCount} paragraph(s) in ${path.basename(filePath)}.`;
  });
}

// ---------------------------------------------------------------------------
// 10. add_comment
// ---------------------------------------------------------------------------

export async function addComment(
  filePath: string,
  anchorText: string,
  commentText: string,
  author: string = "Claude",
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);
    const parsed = await parseDocXml(handle);
    const body = getBody(parsed);

    // Find the anchor text using exact-then-fuzzy strategy
    const match = findAnchorParagraph(body, anchorText);
    if (!match) {
      throw new EngineError(ErrorCode.ANCHOR_NOT_FOUND, `Could not find anchor text "${anchorText}" in ${path.basename(filePath)}.`);
    }

    // Determine comment ID
    let commentsParsed = await parseCommentsXml(handle);
    let commentsChildren: XNode[];

    if (commentsParsed.length === 0) {
      commentsParsed = [
        el("w:comments", [], {
          "xmlns:w":
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        }),
      ];
      commentsChildren = commentsParsed[0]["w:comments"];
    } else {
      commentsChildren = getCommentsArray(commentsParsed);
    }

    const commentId = getNextCommentId(commentsChildren);
    const now = new Date().toISOString();

    // Add comment to comments.xml. Normalize CRLF / lone CR before splitting so
    // no stray "\r" survives inside a comment <w:t> (F3).
    const lines = normalizeNewlines(commentText).split("\n");
    const commentParas = lines.map((line) =>
      el("w:p", [
        el("w:r", [
          el("w:t", [textNode(line)], { "xml:space": "preserve" }),
        ]),
      ]),
    );

    const commentEl = el(
      "w:comment",
      commentParas,
      {
        "w:id": String(commentId),
        "w:author": author,
        "w:date": now,
      },
    );
    commentsChildren.push(commentEl);

    // Write comments.xml
    const commentsXml = builder.build(commentsParsed);
    handle.zip.file("word/comments.xml", commentsXml);

    // Insert comment range markers
    insertCommentRangeMarkers(match.pChildren, commentId, anchorText);

    // Ensure infrastructure
    await ensureCommentsInfrastructure(handle);

    serializeDocXml(handle, parsed);
    await saveDocx(handle);

    const fuzzyNote = match.strategy === "fuzzy" ? " (fuzzy-matched)" : "";
    return `Added comment (ID: ${commentId}) by ${author} on "${anchorText}"${fuzzyNote} in ${path.basename(filePath)}: "${commentText}"`;
  });
}

// ---------------------------------------------------------------------------
// 10b. add_comments
// ---------------------------------------------------------------------------

export async function addComments(
  filePath: string,
  comments: BatchCommentInput[],
  defaultAuthor: string = "Claude",
): Promise<string> {
  if (comments.length === 0) {
    throw new EngineError(ErrorCode.INVALID_PARAMETER, "Comments array must not be empty.");
  }

  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);
    const parsed = await parseDocXml(handle);
    const body = getBody(parsed);

    // Parse or create comments.xml
    let commentsParsed = await parseCommentsXml(handle);
    let commentsChildren: XNode[];

    if (commentsParsed.length === 0) {
      commentsParsed = [
        el("w:comments", [], {
          "xmlns:w":
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        }),
      ];
      commentsChildren = commentsParsed[0]["w:comments"];
    } else {
      commentsChildren = getCommentsArray(commentsParsed);
    }

    interface BatchResult {
      anchor_text: string;
      status: "added" | "failed";
      comment_id?: number;
      error?: string;
      fuzzy_matched?: boolean;
    }

    const results: BatchResult[] = [];
    let added = 0;
    let failed = 0;
    let nextId = getNextCommentId(commentsChildren);
    const now = new Date().toISOString();

    for (const c of comments) {
      const author = c.author ?? defaultAuthor;
      const match = findAnchorParagraph(body, c.anchor_text);

      if (!match) {
        results.push({
          anchor_text: c.anchor_text,
          status: "failed",
          error: `Anchor text not found: "${c.anchor_text}"`,
        });
        failed++;
        continue;
      }

      const commentId = nextId++;

      // Normalize CRLF / lone CR before splitting so no stray "\r" survives
      // inside a comment <w:t> (F3).
      const lines = normalizeNewlines(c.comment_text).split("\n");
      const commentParas = lines.map((line) =>
        el("w:p", [
          el("w:r", [
            el("w:t", [textNode(line)], { "xml:space": "preserve" }),
          ]),
        ]),
      );

      const commentEl = el("w:comment", commentParas, {
        "w:id": String(commentId),
        "w:author": author,
        "w:date": now,
      });
      commentsChildren.push(commentEl);

      insertCommentRangeMarkers(match.pChildren, commentId, c.anchor_text);

      results.push({
        anchor_text: c.anchor_text,
        status: "added",
        comment_id: commentId,
        fuzzy_matched: match.strategy === "fuzzy",
      });
      added++;
    }

    // Write comments.xml once
    const commentsXml = builder.build(commentsParsed);
    handle.zip.file("word/comments.xml", commentsXml);

    await ensureCommentsInfrastructure(handle);
    serializeDocXml(handle, parsed);
    await saveDocx(handle);

    // Build summary
    let summary = `Batch comments on ${path.basename(filePath)}: ${added} added, ${failed} failed.\n`;
    for (const r of results) {
      if (r.status === "added") {
        const fuzzy = r.fuzzy_matched ? " (fuzzy-matched)" : "";
        summary += `  [OK] ID ${r.comment_id}: "${r.anchor_text}"${fuzzy}\n`;
      } else {
        summary += `  [FAIL] "${r.anchor_text}": ${r.error}\n`;
      }
    }
    return summary;
  });
}

// ---------------------------------------------------------------------------
// 11. read_comments
// ---------------------------------------------------------------------------

export async function readComments(filePath: string): Promise<string> {
  const result = await readCommentsStructured(filePath);

  if (result.totalComments === 0) {
    // Emit the structured block (zero total / empty array) so the contract holds
    // even when there are no comments (L2).
    return `No comments found in ${result.file}.\n\n` + jsonBlock(result);
  }

  let output = `Comments in ${result.file} (${result.totalComments}):\n\n`;

  function formatComment(c: CommentData, indent: string): void {
    output += `${indent}[Comment ${c.id}] by ${c.author} (${c.date}):\n${indent}  ${c.text}\n\n`;
    for (const reply of c.replies) {
      formatComment(reply, indent + "  ");
    }
  }

  for (const c of result.comments) {
    formatComment(c, "");
  }

  output += "\n" + jsonBlock(result);
  return output;
}

// ---------------------------------------------------------------------------
// 11a-b. read_comments (structured)
// ---------------------------------------------------------------------------

/** コメントデータの型（返信をネストで保持） */
export interface CommentData {
  id: string;
  author: string;
  date: string;
  text: string;
  replies: CommentData[];
}

/** readCommentsStructured の戻り値型 */
export interface ReadCommentsResult {
  file: string;
  totalComments: number;
  comments: CommentData[];
}

/**
 * readComments と同じロジックだが、テキストではなく構造化オブジェクトを返す。
 */
export async function readCommentsStructured(
  filePath: string,
): Promise<ReadCommentsResult> {
  const handle = await openDocx(filePath);
  const commentsParsed = await parseCommentsXml(handle);

  if (commentsParsed.length === 0) {
    return { file: path.basename(filePath), totalComments: 0, comments: [] };
  }

  const commentsChildren = getCommentsArray(commentsParsed);
  const rawComments = findAll(commentsChildren, "w:comment");

  if (rawComments.length === 0) {
    return { file: path.basename(filePath), totalComments: 0, comments: [] };
  }

  // コメントデータのマップを構築
  const commentMap = new Map<
    string,
    { author: string; date: string; text: string }
  >();
  for (const c of rawComments) {
    const id = attr(c, "w:id") ?? "?";
    const author = attr(c, "w:author") ?? "Unknown";
    const date = attr(c, "w:date") ?? "";
    const cChildren = c["w:comment"] ?? [];
    const paras = findAll(cChildren, "w:p");
    const text = paras
      .map((p) => extractParagraphText(p["w:p"]))
      .join("\n");
    commentMap.set(id, { author, date, text });
  }

  // commentsExtended.xml を解析してスレッド構造を取得
  const extParsed = await parseCommentsExtendedXml(handle);
  const extChildren = getCommentsExtendedArray(extParsed);

  const paraIdToCommentId = new Map<string, string>();
  const childrenOf = new Map<string, string[]>();
  const isReply = new Set<string>();

  if (extChildren.length > 0) {
    for (const c of rawComments) {
      const id = attr(c, "w:id") ?? "?";
      const cChildren = c["w:comment"] ?? [];
      // Map ALL paragraph paraIds (commentsExtended may reference any paragraph, typically the last)
      for (const p of findAll(cChildren, "w:p")) {
        const pId = attr(p, "w14:paraId");
        if (pId) {
          paraIdToCommentId.set(pId, id);
        }
      }
    }

    const commentExEntries = findAll(extChildren, "w15:commentEx");
    for (const ce of commentExEntries) {
      const ceParaId = attr(ce, "w15:paraId") ?? "";
      const parentParaId = attr(ce, "w15:paraIdParent") ?? "";
      if (parentParaId && ceParaId) {
        const childId = paraIdToCommentId.get(ceParaId);
        const parentId = paraIdToCommentId.get(parentParaId);
        if (childId && parentId) {
          isReply.add(childId);
          const existing = childrenOf.get(parentId) ?? [];
          existing.push(childId);
          childrenOf.set(parentId, existing);
        }
      }
    }
  }

  // 再帰的にコメントツリーを構築
  function buildComment(id: string, depth: number = 0): CommentData | null {
    if (depth > 100) return null;
    const c = commentMap.get(id);
    if (!c) return null;
    const replies: CommentData[] = [];
    for (const replyId of childrenOf.get(id) ?? []) {
      const reply = buildComment(replyId, depth + 1);
      if (reply) replies.push(reply);
    }
    return {
      id,
      author: c.author,
      date: c.date,
      text: c.text,
      replies,
    };
  }

  const topLevel: CommentData[] = [];
  for (const c of rawComments) {
    const id = attr(c, "w:id") ?? "?";
    if (!isReply.has(id)) {
      const built = buildComment(id);
      if (built) topLevel.push(built);
    }
  }

  return {
    file: path.basename(filePath),
    totalComments: rawComments.length,
    comments: topLevel,
  };
}

// ---------------------------------------------------------------------------
// 11b. reply_to_comment
// ---------------------------------------------------------------------------

export async function replyToComment(
  filePath: string,
  parentCommentId: number,
  commentText: string,
  author: string = "Claude",
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);

    // Verify parent comment exists
    const commentsParsed = await parseCommentsXml(handle);
    if (commentsParsed.length === 0) {
      throw new EngineError(ErrorCode.INVALID_PARAMETER, `No comments found in ${path.basename(filePath)}.`);
    }

    const commentsChildren = getCommentsArray(commentsParsed);
    const parentComment = findAll(commentsChildren, "w:comment").find(
      (c) => attr(c, "w:id") === String(parentCommentId),
    );

    if (!parentComment) {
      throw new EngineError(ErrorCode.INVALID_PARAMETER, `Comment ID ${parentCommentId} not found.`);
    }

    // Use last <w:p>'s w14:paraId (commentsExtended convention uses the last paragraph)
    const parentChildren = parentComment["w:comment"] ?? [];
    const parentParas = findAll(parentChildren, "w:p");
    if (parentParas.length === 0) {
      throw new EngineError(ErrorCode.INVALID_PARAMETER, `Parent comment ${parentCommentId} has no paragraphs.`);
    }
    const parentLastPara = parentParas[parentParas.length - 1];
    let parentParaId = attr(parentLastPara, "w14:paraId") ?? "";
    if (!parentParaId) {
      parentParaId = generateParaId();
      setAttr(parentLastPara, "w14:paraId", parentParaId);
    }

    // Create reply comment
    const replyId = getNextCommentId(commentsChildren);
    const now = new Date().toISOString();
    const replyParaId = generateParaId();

    // Normalize CRLF / lone CR before splitting so no stray "\r" survives inside
    // a reply <w:t> (N3 — parity with addComment / addComments).
    const lines = normalizeNewlines(commentText).split("\n");
    const replyParas = lines.map((line, idx) => {
      const para = el("w:p", [
        el("w:r", [
          el("w:t", [textNode(line)], { "xml:space": "preserve" }),
        ]),
      ]);
      // Set paraId on last paragraph for threading (commentsExtended convention)
      if (idx === lines.length - 1) {
        setAttr(para, "w14:paraId", replyParaId);
      }
      return para;
    });

    const replyEl = el("w:comment", replyParas, {
      "w:id": String(replyId),
      "w:author": author,
      "w:date": now,
    });
    commentsChildren.push(replyEl);

    // N1: a w14:paraId is now written into comment paragraphs (the parent's last
    // <w:p> above and the reply's last <w:p>). The <w:comments> root must declare
    // the w14 namespace, or the part is namespace-malformed ("w14 … not defined").
    // Mirror the document.xml handling — ensure the binding on the comments root.
    const commentsRoot = commentsParsed.find((n: XNode) => n["w:comments"]);
    if (commentsRoot) ensureW14Namespace(commentsRoot);

    // Write comments.xml
    const commentsXml = builder.build(commentsParsed);
    handle.zip.file("word/comments.xml", commentsXml);

    // Create or update commentsExtended.xml
    let extParsed = await parseCommentsExtendedXml(handle);
    let extChildren: XNode[];

    if (extParsed.length === 0) {
      extParsed = [
        el("w15:commentsEx", [], {
          "xmlns:w15": "http://schemas.microsoft.com/office/word/2012/wordml",
          "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
          "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
          "mc:Ignorable": "w14 w15",
        }),
      ];
      extChildren = extParsed[0]["w15:commentsEx"];
    } else {
      extChildren = getCommentsExtendedArray(extParsed);
    }

    // Ensure parent has a commentEx entry
    const parentExExists = findAll(extChildren, "w15:commentEx").some(
      (ce) => attr(ce, "w15:paraId") === parentParaId,
    );
    if (!parentExExists) {
      extChildren.push(
        el("w15:commentEx", [], {
          "w15:paraId": parentParaId,
          "w15:done": "0",
        }),
      );
    }

    // Add reply commentEx entry with paraIdParent pointing to parent
    extChildren.push(
      el("w15:commentEx", [], {
        "w15:paraId": replyParaId,
        "w15:paraIdParent": parentParaId,
        "w15:done": "0",
      }),
    );

    const extXml = builder.build(extParsed);
    handle.zip.file("word/commentsExtended.xml", extXml);

    // Ensure infrastructure
    await ensureCommentsInfrastructure(handle);
    await ensureCommentsExtendedInfrastructure(handle);

    await saveDocx(handle);

    return `Added reply (ID: ${replyId}) to comment ${parentCommentId} by ${author} in ${path.basename(filePath)}: "${commentText}"`;
  });
}

// ---------------------------------------------------------------------------
// 12. delete_comment
// ---------------------------------------------------------------------------

export async function deleteComment(
  filePath: string,
  commentId: number,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);

    // Remove from comments.xml — capture ALL paraIds before removing
    const commentsParsed = await parseCommentsXml(handle);
    const deletedParaIds = new Set<string>();
    if (commentsParsed.length > 0) {
      const commentsChildren = getCommentsArray(commentsParsed);
      const idx = commentsChildren.findIndex(
        (c: XNode) =>
          c["w:comment"] !== undefined && attr(c, "w:id") === String(commentId),
      );
      if (idx !== -1) {
        // Extract all paraIds from all paragraphs before removing
        const cChildren = commentsChildren[idx]["w:comment"] ?? [];
        for (const p of findAll(cChildren, "w:p")) {
          const pId = attr(p, "w14:paraId");
          if (pId) deletedParaIds.add(pId);
        }
        commentsChildren.splice(idx, 1);
        const commentsXml = builder.build(commentsParsed);
        handle.zip.file("word/comments.xml", commentsXml);
      }
    }

    // Remove corresponding entries from commentsExtended.xml
    if (deletedParaIds.size > 0) {
      const extParsed = await parseCommentsExtendedXml(handle);
      const extChildren = getCommentsExtendedArray(extParsed);
      if (extChildren.length > 0) {
        for (let i = extChildren.length - 1; i >= 0; i--) {
          const ce = extChildren[i];
          if (ce["w15:commentEx"] !== undefined) {
            const ceParaId = attr(ce, "w15:paraId") ?? "";
            const ceParentId = attr(ce, "w15:paraIdParent") ?? "";
            if (deletedParaIds.has(ceParaId) || deletedParaIds.has(ceParentId)) {
              extChildren.splice(i, 1);
            }
          }
        }
        handle.zip.file("word/commentsExtended.xml", builder.build(extParsed));
      }
    }

    // Remove comment range markers and references from document.xml
    const parsed = await parseDocXml(handle);
    const body = getBody(parsed);
    const idStr = String(commentId);

    function removeCommentMarkers(nodes: XNode[]): void {
      for (let i = nodes.length - 1; i >= 0; i--) {
        const node = nodes[i];
        if (
          node["w:commentRangeStart"] !== undefined &&
          attr(node, "w:id") === idStr
        ) {
          nodes.splice(i, 1);
        } else if (
          node["w:commentRangeEnd"] !== undefined &&
          attr(node, "w:id") === idStr
        ) {
          nodes.splice(i, 1);
        } else if (node["w:r"]) {
          // Check if run contains only a commentReference for this id
          const runC = node["w:r"] as XNode[];
          const ref = findOne(runC, "w:commentReference");
          if (ref && attr(ref, "w:id") === idStr) {
            nodes.splice(i, 1);
          }
        }
      }
    }

    for (const child of body) {
      if (child["w:p"]) {
        removeCommentMarkers(child["w:p"]);
      } else if (child["w:tbl"]) {
        const rows = findAll(child["w:tbl"], "w:tr");
        for (const row of rows) {
          const cells = findAll(row["w:tr"], "w:tc");
          for (const cell of cells) {
            const paras = findAll(cell["w:tc"], "w:p");
            for (const p of paras) {
              removeCommentMarkers(p["w:p"]);
            }
          }
        }
      }
    }

    serializeDocXml(handle, parsed);
    await saveDocx(handle);

    return `Deleted comment ${commentId} from ${path.basename(filePath)}.`;
  });
}

// ---------------------------------------------------------------------------
// 13. create_document
// ---------------------------------------------------------------------------

export type CreateDocumentPreset = "ja-business";

function buildStylesXml(preset?: CreateDocumentPreset): string {
  if (preset === "ja-business") {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Yu Gothic" w:hAnsi="Yu Gothic" w:eastAsia="Yu Gothic"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="480" w:after="280" w:line="440" w:lineRule="auto"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="Heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="360" w:after="200" w:line="400" w:lineRule="auto"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="Heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="280" w:after="160" w:line="380" w:lineRule="auto"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>
</w:styles>`;
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="480" w:after="0"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="Heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="200" w:after="0"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="Heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="200" w:after="0"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>
</w:styles>`;
}

function assertCreateDocumentPreset(
  preset?: CreateDocumentPreset,
): void {
  if (preset && preset !== "ja-business") {
    throw new EngineError(
      ErrorCode.INVALID_PARAMETER,
      `Unsupported create_document preset: ${preset}`,
    );
  }
}

async function readStylesXml(handle: DocxHandle): Promise<XNode[]> {
  const stylesXml = await handle.zip.file("word/styles.xml")?.async("string");
  if (!stylesXml) {
    throw new EngineError(
      ErrorCode.INVALID_DOCX,
      "word/styles.xml not found in DOCX",
    );
  }
  return parser.parse(stylesXml);
}

function serializeStylesXml(handle: DocxHandle, parsed: XNode[]): void {
  handle.zip.file("word/styles.xml", builder.build(parsed));
}

function getStylesRoot(parsed: XNode[]): XNode {
  const root = parsed.find((n: XNode) => n["w:styles"]);
  if (!root) {
    throw new EngineError(
      ErrorCode.INVALID_DOCX,
      "w:styles root not found in word/styles.xml",
    );
  }
  return root;
}

function getPresetStylesNodes(
  preset: CreateDocumentPreset,
): {
  docDefaults?: XNode;
  headingStyles: Map<string, XNode>;
  normalStyle?: XNode;
} {
  const parsed = parser.parse(buildStylesXml(preset));
  const root = getStylesRoot(parsed);
  const children = root["w:styles"] as XNode[];

  const headingStyles = new Map<string, XNode>();
  let normalStyle: XNode | undefined;

  for (const node of children) {
    if (!node["w:style"]) continue;
    const styleId = attr(node, "w:styleId");
    if (!styleId) continue;
    if (styleId === "Normal") {
      normalStyle = cloneNode(node);
    } else if (styleId === "Heading1" || styleId === "Heading2" || styleId === "Heading3") {
      headingStyles.set(styleId, cloneNode(node));
    }
  }

  const docDefaultsNode = findOne(children, "w:docDefaults");

  return {
    docDefaults: docDefaultsNode ? cloneNode(docDefaultsNode) : undefined,
    headingStyles,
    normalStyle,
  };
}

function upsertStyleNode(
  children: XNode[],
  styleId: string,
  node: XNode,
): void {
  const idx = children.findIndex(
    (child) => child["w:style"] !== undefined && attr(child, "w:styleId") === styleId,
  );
  if (idx !== -1) {
    children[idx] = cloneNode(node);
  } else {
    children.push(cloneNode(node));
  }
}

export async function createDocument(
  filePath: string,
  title?: string,
  content?: string,
  preset?: CreateDocumentPreset,
): Promise<string> {
  // Validate file extension to prevent writing to arbitrary paths
  if (!/\.docx$/i.test(filePath)) {
    throw new EngineError(
      ErrorCode.INVALID_PARAMETER,
      `File path must end with .docx: ${filePath}`,
    );
  }

  assertCreateDocumentPreset(preset);

  return withFileLock(filePath, async () => {
    // Check parent directory exists
    const dir = path.dirname(filePath);
    await fs.mkdir(dir, { recursive: true });

    // Build paragraphs from content
    let bodyXml = "";
    if (title) {
      // Normalize CRLF / lone CR (F3) so a "\r" can't survive inside the title
      // <w:t> the way it does for content; render any "\n" as a soft break.
      const titleRuns = normalizeNewlines(title)
        .split("\n")
        .map((seg) => escapeXml(seg))
        .join('</w:t><w:br/><w:t xml:space="preserve">');
      bodyXml += `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">${titleRuns}</w:t></w:r></w:p>\n`;
    }

    if (content) {
      // Normalize CRLF / lone CR to "\n" BEFORE splitting so a preceding "\r"
      // does not survive inside a <w:t> (a conformant reader would otherwise
      // turn it into a literal newline char in the run) — F3.
      const lines = normalizeNewlines(content).split("\n");
      for (const line of lines) {
        bodyXml += `<w:p><w:r><w:t xml:space="preserve">${escapeXml(line)}</w:t></w:r></w:p>\n`;
      }
    } else if (!title) {
      // At least one empty paragraph
      bodyXml = "<w:p/>\n";
    }

    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
${bodyXml}
<w:sectPr>
<w:pgSz w:w="11906" w:h="16838"/>
<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
</w:sectPr>
</w:body>
</w:document>`;

    const stylesXml = buildStylesXml(preset);

    const numberingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:abstractNum w:abstractNumId="0">
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/></w:lvl>
</w:abstractNum>
<w:abstractNum w:abstractNumId="1">
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl>
</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;

    const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>`;

    const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

    const docRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>`;

    const zip = new JSZip();
    zip.file("[Content_Types].xml", contentTypesXml);
    zip.file("_rels/.rels", relsXml);
    zip.file("word/document.xml", documentXml);
    zip.file("word/styles.xml", stylesXml);
    zip.file("word/numbering.xml", numberingXml);
    zip.file("word/_rels/document.xml.rels", docRelsXml);

    const buf = await zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
    });
    await fs.writeFile(filePath, buf);

    return `Created document: ${filePath}`;
  });
}

export async function applyDocumentPreset(
  filePath: string,
  preset: CreateDocumentPreset,
): Promise<string> {
  assertCreateDocumentPreset(preset);

  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);
    const stylesParsed = await readStylesXml(handle);
    const root = getStylesRoot(stylesParsed);
    const children = root["w:styles"] as XNode[];
    const presetNodes = getPresetStylesNodes(preset);

    const defaultsIdx = children.findIndex((child) => child["w:docDefaults"] !== undefined);
    if (presetNodes.docDefaults) {
      if (defaultsIdx !== -1) {
        children[defaultsIdx] = cloneNode(presetNodes.docDefaults);
      } else {
        children.unshift(cloneNode(presetNodes.docDefaults));
      }
    }

    // Preserve any existing Normal style — users often customize it. Only seed
    // it when the document has none. Heading1–3 are part of the preset's
    // intent, so they're upserted unconditionally.
    if (presetNodes.normalStyle) {
      const hasNormal = children.some(
        (child) => child["w:style"] !== undefined && attr(child, "w:styleId") === "Normal",
      );
      if (!hasNormal) {
        children.push(cloneNode(presetNodes.normalStyle));
      }
    }

    for (const [styleId, node] of presetNodes.headingStyles.entries()) {
      upsertStyleNode(children, styleId, node);
    }

    serializeStylesXml(handle, stylesParsed);
    await saveDocx(handle);

    return `Applied document preset "${preset}" to ${path.basename(filePath)}.`;
  });
}

// ---------------------------------------------------------------------------
// 14. highlight_text
// ---------------------------------------------------------------------------

export async function highlightText(
  filePath: string,
  search: string,
  color: string = "yellow",
  caseSensitive: boolean = false,
): Promise<string> {
  return formatText(filePath, search, { highlightColor: color }, caseSensitive);
}

// ---------------------------------------------------------------------------
// 15. insert_table
// ---------------------------------------------------------------------------

/** Bounds for insert_table dimensions (defense-in-depth; also enforced by the
 * MCP zod schema). cols ≤ 63 matches Word's column limit; rows ≤ 1000 is a sane
 * document bound; the rows*cols ≤ 20000 product cap is the OOM guard — a single
 * bad call must not exhaust the long-lived server's heap. */
const MAX_TABLE_COLS = 63;
const MAX_TABLE_ROWS = 1000;
const MAX_TABLE_CELLS = 20000;

export async function insertTable(
  filePath: string,
  position: number,
  rows: number,
  cols: number,
  data?: string[][],
): Promise<string> {
  // Validate dimensions BEFORE acquiring the file lock or building any nodes,
  // so a hostile/huge count fails fast and can never OOM the process (N2/N9).
  if (!Number.isInteger(rows) || !Number.isInteger(cols)) {
    throw new EngineError(
      ErrorCode.INVALID_PARAMETER,
      `Table rows and cols must be integers. Got rows=${rows}, cols=${cols}.`,
    );
  }
  if (rows < 1 || cols < 1) {
    throw new EngineError(
      ErrorCode.INVALID_PARAMETER,
      `Table rows and cols must be at least 1. Got rows=${rows}, cols=${cols}.`,
    );
  }
  if (cols > MAX_TABLE_COLS) {
    throw new EngineError(
      ErrorCode.INVALID_PARAMETER,
      `Table cols must be at most ${MAX_TABLE_COLS} (Word's column limit). Got ${cols}.`,
    );
  }
  if (rows > MAX_TABLE_ROWS) {
    throw new EngineError(
      ErrorCode.INVALID_PARAMETER,
      `Table rows must be at most ${MAX_TABLE_ROWS}. Got ${rows}.`,
    );
  }
  if (rows * cols > MAX_TABLE_CELLS) {
    throw new EngineError(
      ErrorCode.INVALID_PARAMETER,
      `Table is too large: rows*cols must be at most ${MAX_TABLE_CELLS}. Got ${rows}x${cols}=${rows * cols}.`,
    );
  }

  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);
    const parsed = await parseDocXml(handle);
    const body = getBody(parsed);
    const refs = enumerateBlockRefs(body);

    // Build table XML nodes
    const tblChildren: XNode[] = [];

    // Table properties
    tblChildren.push(
      el("w:tblPr", [
        el("w:tblW", [], { "w:w": "0", "w:type": "auto" }),
        el("w:tblBorders", [
          el("w:top", [], { "w:val": "single", "w:sz": "4", "w:space": "0", "w:color": "auto" }),
          el("w:left", [], { "w:val": "single", "w:sz": "4", "w:space": "0", "w:color": "auto" }),
          el("w:bottom", [], { "w:val": "single", "w:sz": "4", "w:space": "0", "w:color": "auto" }),
          el("w:right", [], { "w:val": "single", "w:sz": "4", "w:space": "0", "w:color": "auto" }),
          el("w:insideH", [], { "w:val": "single", "w:sz": "4", "w:space": "0", "w:color": "auto" }),
          el("w:insideV", [], { "w:val": "single", "w:sz": "4", "w:space": "0", "w:color": "auto" }),
        ]),
      ]),
    );

    // Table grid — REQUIRED by ECMA-376 CT_Tbl, immediately after w:tblPr. Must
    // carry exactly `cols` w:gridCol children or strict validators (python-docx)
    // reject the file with InvalidXmlError.
    const gridCols: XNode[] = [];
    for (let c = 0; c < cols; c++) {
      gridCols.push(el("w:gridCol", [], { "w:w": "0" }));
    }
    tblChildren.push(el("w:tblGrid", gridCols));

    // Table rows. The grid is always rows×cols: `data` values beyond rows×cols
    // are ignored (extra rows/columns dropped), and missing cells (short rows or
    // omitted data) are blank-padded via the `?? ""` below. Documented in the
    // insert_table `data` description.
    for (let r = 0; r < rows; r++) {
      const cellNodes: XNode[] = [];
      for (let c = 0; c < cols; c++) {
        const cellText = data?.[r]?.[c] ?? "";
        // Render the cell text via makeTextRuns so an embedded "\n" (after CRLF
        // normalization) becomes a <w:br/> SOFT break inside the cell run rather
        // than a literal newline left inside a single <w:t> — matching the
        // tracked-cell / untracked editTableParagraphs convention (F3). An empty
        // cell yields a paragraph with no run (valid OOXML, reads as "").
        cellNodes.push(
          el("w:tc", [
            el("w:tcPr", [el("w:tcW", [], { "w:w": "0", "w:type": "auto" })]),
            el("w:p", makeTextRuns(cellText, null)),
          ]),
        );
      }
      tblChildren.push(el("w:tr", cellNodes));
    }

    const tblNode = el("w:tbl", tblChildren);

    // Insert at position (unified block-index space; SDT-inner positions are
    // refused, out-of-range appends before sectPr).
    spliceNewParagraph(body, refs, position, [tblNode]);

    serializeDocXml(handle, parsed);
    await saveDocx(handle);

    return `Inserted ${rows}x${cols} table at position ${position} in ${path.basename(filePath)}.`;
  });
}

// ---------------------------------------------------------------------------
// Shared helper: apply heading level to a paragraph element
// ---------------------------------------------------------------------------

/** Set pStyle to Heading{level} and outlineLvl to level-1 on a w:p element. */
function applyHeadingLevel(pChildren: XNode[], level: number): void {
  // Defensive: callers validate, but never emit a fractional style id /
  // outlineLvl (both must be integers per OOXML; see setHeadings, N5).
  if (!Number.isInteger(level) || level < 1 || level > 9) {
    throw new EngineError(ErrorCode.INVALID_PARAMETER, `Heading level must be an integer between 1 and 9. Got: ${level}.`);
  }
  let pPr = findOne(pChildren, "w:pPr");
  if (!pPr) {
    pPr = el("w:pPr");
    pChildren.unshift(pPr);
  }

  const props = pPr["w:pPr"] as XNode[];
  const styleId = `Heading${level}`;
  const pStyleIdx = props.findIndex((n: XNode) => n["w:pStyle"] !== undefined);
  const pStyleEl = el("w:pStyle", [], { "w:val": styleId });
  if (pStyleIdx !== -1) {
    props[pStyleIdx] = pStyleEl;
  } else {
    props.unshift(pStyleEl);
  }

  const olvlIdx = props.findIndex((n: XNode) => n["w:outlineLvl"] !== undefined);
  const olvlEl = el("w:outlineLvl", [], { "w:val": String(level - 1) });
  if (olvlIdx !== -1) {
    props[olvlIdx] = olvlEl;
  } else {
    props.push(olvlEl);
  }
}

// ---------------------------------------------------------------------------
// 16. set_headings
// ---------------------------------------------------------------------------

export interface SetHeadingItem {
  /** Exactly one of paragraphIndex / anchor must be set. */
  paragraphIndex?: number;
  anchor?: string;
  level: number;
}

export async function setHeadings(
  filePath: string,
  items: SetHeadingItem[],
): Promise<string> {
  if (items.length === 0) {
    return `No headings to set.`;
  }

  // Validate levels upfront (before acquiring file lock). Must be an INTEGER in
  // 1..9 — a fractional level (e.g. 1.5) would emit a nonexistent "Heading1.5"
  // style and a schema-invalid outlineLvl "0.5" (ST_DecimalNumber is integer),
  // which python-docx then chokes on (N5).
  for (const item of items) {
    if (!Number.isInteger(item.level) || item.level < 1 || item.level > 9) {
      throw new EngineError(ErrorCode.INVALID_PARAMETER, `Heading level must be an integer between 1 and 9. Got: ${item.level}.`);
    }
  }

  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);
    const parsed = await parseDocXml(handle);
    const root = getDocumentRoot(parsed);
    const body = getBody(parsed);
    const refs = enumerateBlockRefs(body);
    const anchorIndex = buildAnchorIndex(body);

    // Resolve every locator (index or anchor) upfront. `directBody` records
    // which targets are DIRECT-body paragraphs (anchor scope, v1); only those
    // auto-seed. An anchor target is always direct-body; an index may resolve
    // into a top-level w:sdt, which is out of anchor scope.
    const directBody = new Set<XNode>();
    const targets: XNode[] = items.map((item) => {
      const loc = locatorToLocation(refs, anchorIndex, item);
      if (loc.parent === body) directBody.add(loc.element);
      return loc.element;
    });

    const seeder = new AnchorSeeder(root, body);
    targets.forEach((element, i) => {
      applyHeadingLevel(element["w:p"] as XNode[], items[i].level);
      if (directBody.has(element)) seeder.seed(element);
    });

    serializeDocXml(handle, parsed);
    await saveDocx(handle);

    return `Set ${items.length} paragraph(s) to headings in ${path.basename(filePath)}.`;
  });
}

// ---------------------------------------------------------------------------
// 17. get_page_layout
// ---------------------------------------------------------------------------

export async function getPageLayout(filePath: string): Promise<string> {
  const handle = await openDocx(filePath);
  const parsed = await parseDocXml(handle);
  const body = getBody(parsed);
  const sectPr = getSectPr(body);

  if (!sectPr) {
    return `No section properties found in ${path.basename(filePath)}.`;
  }

  const sectChildren = sectPr["w:sectPr"] as XNode[];

  // Page size
  const pgSz = findOne(sectChildren, "w:pgSz");
  const w = parseInt(attr(pgSz, "w:w") ?? "11906");
  const h = parseInt(attr(pgSz, "w:h") ?? "16838");
  const orient = attr(pgSz, "w:orient") ?? (w > h ? "landscape" : "portrait");

  // Page margins
  const pgMar = findOne(sectChildren, "w:pgMar");
  const top = parseInt(attr(pgMar, "w:top") ?? "1440");
  const right = parseInt(attr(pgMar, "w:right") ?? "1440");
  const bottom = parseInt(attr(pgMar, "w:bottom") ?? "1440");
  const left = parseInt(attr(pgMar, "w:left") ?? "1440");
  const header = parseInt(attr(pgMar, "w:header") ?? "720");
  const footer = parseInt(attr(pgMar, "w:footer") ?? "720");
  const gutter = parseInt(attr(pgMar, "w:gutter") ?? "0");

  const preset = detectPageSizePreset(w, h);

  let output = `Page layout for ${path.basename(filePath)}:\n\n`;
  output += `Page size: ${twipsToMm(w)} × ${twipsToMm(h)} mm`;
  if (preset) output += `  (${PAGE_SIZE_PRESETS[preset].label})`;
  output += `\n`;
  output += `Orientation: ${orient}\n\n`;
  output += `Margins (mm):\n`;
  output += `  Top:    ${twipsToMm(top)}\n`;
  output += `  Right:  ${twipsToMm(right)}\n`;
  output += `  Bottom: ${twipsToMm(bottom)}\n`;
  output += `  Left:   ${twipsToMm(left)}\n`;
  output += `  Header: ${twipsToMm(header)}\n`;
  output += `  Footer: ${twipsToMm(footer)}\n`;
  output += `  Gutter: ${twipsToMm(gutter)}\n`;

  output += `\nAvailable page size presets: ${Object.keys(PAGE_SIZE_PRESETS).join(", ")}\n`;
  output += `Available margin presets: ${Object.keys(MARGIN_PRESETS).join(", ")}\n`;

  return output;
}

// ---------------------------------------------------------------------------
// 18. set_page_layout
// ---------------------------------------------------------------------------

// OOXML page geometry bounds (ECMA-376). w:w/w:h are unsigned (ST_TwipsMeasure)
// and Word caps a page at 22 in = 31680 twips ≈ 558.8 mm; margins are signed in
// the schema but a negative page margin is nonsensical here, so we require ≥ 0
// and within the same upper bound. mm bounds derived from the twip limits (N10).
const MAX_PAGE_TWIPS = 31680; // 22 inches
const MAX_PAGE_MM = MAX_PAGE_TWIPS / TWIPS_PER_MM; // ≈ 558.8

/** Throw INVALID_PARAMETER if a page dimension (width/height) is out of range. */
function validatePageDimMm(name: string, mm: number): void {
  if (!Number.isFinite(mm) || mm <= 0 || mm > MAX_PAGE_MM) {
    throw new EngineError(
      ErrorCode.INVALID_PARAMETER,
      `${name} must be > 0 and ≤ ${MAX_PAGE_MM.toFixed(1)} mm. Got: ${mm}.`,
    );
  }
}

/** Throw INVALID_PARAMETER if a margin is negative or out of range. */
function validateMarginMm(name: string, mm: number): void {
  if (!Number.isFinite(mm) || mm < 0 || mm > MAX_PAGE_MM) {
    throw new EngineError(
      ErrorCode.INVALID_PARAMETER,
      `${name} must be ≥ 0 and ≤ ${MAX_PAGE_MM.toFixed(1)} mm. Got: ${mm}.`,
    );
  }
}

export async function setPageLayout(
  filePath: string,
  options: PageLayoutOptions,
): Promise<string> {
  // Validate custom geometry BEFORE the file lock so a negative/absurd value is
  // rejected without writing schema-invalid (unsigned) w:w/w:h/margins (N10).
  if (options.widthMm !== undefined) validatePageDimMm("Page width", options.widthMm);
  if (options.heightMm !== undefined) validatePageDimMm("Page height", options.heightMm);
  if (options.topMm !== undefined) validateMarginMm("Top margin", options.topMm);
  if (options.rightMm !== undefined) validateMarginMm("Right margin", options.rightMm);
  if (options.bottomMm !== undefined) validateMarginMm("Bottom margin", options.bottomMm);
  if (options.leftMm !== undefined) validateMarginMm("Left margin", options.leftMm);
  if (options.headerMm !== undefined) validateMarginMm("Header distance", options.headerMm);
  if (options.footerMm !== undefined) validateMarginMm("Footer distance", options.footerMm);
  if (options.gutterMm !== undefined) validateMarginMm("Gutter margin", options.gutterMm);

  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);
    const parsed = await parseDocXml(handle);
    const body = getBody(parsed);

    let sectPr = getSectPr(body);
    if (!sectPr) {
      sectPr = el("w:sectPr", [
        el("w:pgSz", [], { "w:w": "11906", "w:h": "16838" }),
        el("w:pgMar", [], {
          "w:top": "1440", "w:right": "1440", "w:bottom": "1440", "w:left": "1440",
          "w:header": "720", "w:footer": "720", "w:gutter": "0",
        }),
      ]);
      body.push(sectPr);
    }

    const sectChildren = sectPr["w:sectPr"] as XNode[];
    const changes: string[] = [];

    // --- Page size ---
    let pgSz = findOne(sectChildren, "w:pgSz");
    if (!pgSz) {
      pgSz = el("w:pgSz", [], { "w:w": "11906", "w:h": "16838" });
      sectChildren.unshift(pgSz);
    }

    if (options.pageSizePreset) {
      const key = options.pageSizePreset.toUpperCase();
      const preset = PAGE_SIZE_PRESETS[key];
      if (!preset) {
        throw new EngineError(ErrorCode.INVALID_PARAMETER, `Unknown page size preset "${options.pageSizePreset}". Available: ${Object.keys(PAGE_SIZE_PRESETS).join(", ")}`);
      }
      const orient = options.orientation ?? "portrait";
      if (orient === "landscape") {
        setAttr(pgSz, "w:w", String(preset.h));
        setAttr(pgSz, "w:h", String(preset.w));
        setAttr(pgSz, "w:orient", "landscape");
      } else {
        setAttr(pgSz, "w:w", String(preset.w));
        setAttr(pgSz, "w:h", String(preset.h));
        if (pgSz[":@"]) delete pgSz[":@"]["@_w:orient"];
      }
      changes.push(`page size → ${preset.label}${orient === "landscape" ? " (landscape)" : ""}`);
    } else if (options.widthMm !== undefined || options.heightMm !== undefined) {
      if (options.widthMm !== undefined) {
        setAttr(pgSz, "w:w", String(mmToTwips(options.widthMm)));
      }
      if (options.heightMm !== undefined) {
        setAttr(pgSz, "w:h", String(mmToTwips(options.heightMm)));
      }
      const curW = parseInt(attr(pgSz, "w:w") ?? "0");
      const curH = parseInt(attr(pgSz, "w:h") ?? "0");
      if (curW > curH) {
        setAttr(pgSz, "w:orient", "landscape");
      } else {
        if (pgSz[":@"]) delete pgSz[":@"]["@_w:orient"];
      }
      changes.push(`page size → ${twipsToMm(curW)} × ${twipsToMm(curH)} mm`);
    } else if (options.orientation) {
      const curW = parseInt(attr(pgSz, "w:w") ?? "11906");
      const curH = parseInt(attr(pgSz, "w:h") ?? "16838");
      const isCurrentlyLandscape = curW > curH;
      const wantLandscape = options.orientation === "landscape";
      if (isCurrentlyLandscape !== wantLandscape) {
        setAttr(pgSz, "w:w", String(curH));
        setAttr(pgSz, "w:h", String(curW));
      }
      if (wantLandscape) {
        setAttr(pgSz, "w:orient", "landscape");
      } else {
        if (pgSz[":@"]) delete pgSz[":@"]["@_w:orient"];
      }
      changes.push(`orientation → ${options.orientation}`);
    }

    // --- Margins ---
    let pgMar = findOne(sectChildren, "w:pgMar");
    if (!pgMar) {
      pgMar = el("w:pgMar", [], {
        "w:top": "1440", "w:right": "1440", "w:bottom": "1440", "w:left": "1440",
        "w:header": "720", "w:footer": "720", "w:gutter": "0",
      });
      sectChildren.push(pgMar);
    }

    if (options.marginPreset) {
      const key = options.marginPreset.toUpperCase();
      const preset = MARGIN_PRESETS[key];
      if (!preset) {
        throw new EngineError(ErrorCode.INVALID_PARAMETER, `Unknown margin preset "${options.marginPreset}". Available: ${Object.keys(MARGIN_PRESETS).join(", ")}`);
      }
      setAttr(pgMar, "w:top", String(preset.top));
      setAttr(pgMar, "w:right", String(preset.right));
      setAttr(pgMar, "w:bottom", String(preset.bottom));
      setAttr(pgMar, "w:left", String(preset.left));
      changes.push(`margins → ${preset.label}`);
    }

    if (options.topMm !== undefined) { setAttr(pgMar, "w:top", String(mmToTwips(options.topMm))); changes.push(`top → ${options.topMm} mm`); }
    if (options.rightMm !== undefined) { setAttr(pgMar, "w:right", String(mmToTwips(options.rightMm))); changes.push(`right → ${options.rightMm} mm`); }
    if (options.bottomMm !== undefined) { setAttr(pgMar, "w:bottom", String(mmToTwips(options.bottomMm))); changes.push(`bottom → ${options.bottomMm} mm`); }
    if (options.leftMm !== undefined) { setAttr(pgMar, "w:left", String(mmToTwips(options.leftMm))); changes.push(`left → ${options.leftMm} mm`); }
    if (options.headerMm !== undefined) { setAttr(pgMar, "w:header", String(mmToTwips(options.headerMm))); changes.push(`header → ${options.headerMm} mm`); }
    if (options.footerMm !== undefined) { setAttr(pgMar, "w:footer", String(mmToTwips(options.footerMm))); changes.push(`footer → ${options.footerMm} mm`); }
    if (options.gutterMm !== undefined) { setAttr(pgMar, "w:gutter", String(mmToTwips(options.gutterMm))); changes.push(`gutter → ${options.gutterMm} mm`); }

    if (changes.length === 0) {
      return "No page layout changes specified.";
    }

    serializeDocXml(handle, parsed);
    await saveDocx(handle);

    return `Updated page layout in ${path.basename(filePath)}: ${changes.join(", ")}.`;
  });
}

// ---------------------------------------------------------------------------
// 19. accept_all_changes
// ---------------------------------------------------------------------------

export async function acceptAllChanges(filePath: string): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);
    const parsed = await parseDocXml(handle);
    const body = getBody(parsed);

    acceptChangesInNodes(body);

    // Process headers and footers
    await processHeaderFooterChanges(handle, acceptChangesInNodes);

    serializeDocXml(handle, parsed);
    await saveDocx(handle);

    return `Accepted all tracked changes in ${path.basename(filePath)}.`;
  });
}

// ---------------------------------------------------------------------------
// 20. reject_all_changes
// ---------------------------------------------------------------------------

export async function rejectAllChanges(filePath: string): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);
    const parsed = await parseDocXml(handle);
    const body = getBody(parsed);

    rejectChangesInNodes(body);

    // Process headers and footers
    await processHeaderFooterChanges(handle, rejectChangesInNodes);

    serializeDocXml(handle, parsed);
    await saveDocx(handle);

    return `Rejected all tracked changes in ${path.basename(filePath)}.`;
  });
}

/** Apply a change-processing function to all headers and footers in the document. */
async function processHeaderFooterChanges(
  handle: DocxHandle,
  processFn: (nodes: XNode[]) => void,
): Promise<void> {
  const hfFiles = getHeaderFooterFiles(handle);
  for (const hfFile of hfFiles) {
    const hfXml = await handle.zip.file(hfFile)?.async("string");
    if (!hfXml) continue;
    const hfParsed: XNode[] = parser.parse(hfXml);
    const rootEl = hfParsed.find((n: XNode) => n["w:hdr"] || n["w:ftr"]);
    if (!rootEl) continue;
    const hfChildren = rootEl["w:hdr"] ?? rootEl["w:ftr"];
    processFn(hfChildren);
    const updatedXml = builder.build(hfParsed);
    handle.zip.file(hfFile, updatedXml);
  }
}

// ---------------------------------------------------------------------------
// 21. read_header_footer
// ---------------------------------------------------------------------------

export async function readHeaderFooter(filePath: string): Promise<string> {
  const handle = await openDocx(filePath);
  const hfFiles = getHeaderFooterFiles(handle);

  if (hfFiles.length === 0) {
    return `No headers or footers found in ${path.basename(filePath)}.`;
  }

  const sections: string[] = [];
  for (const hfFile of hfFiles) {
    const xml = await handle.zip.file(hfFile)?.async("string");
    if (!xml) continue;
    const parsed: XNode[] = parser.parse(xml);
    const isHeader = hfFile.includes("header");
    const type = isHeader ? "Header" : "Footer";
    const text = extractTextFromHdrFtr(parsed);
    if (text.trim()) {
      sections.push(`[${type}: ${path.basename(hfFile)}]\n${text}`);
    }
  }

  if (sections.length === 0) {
    return `No headers or footers found in ${path.basename(filePath)}.`;
  }

  return sections.join("\n\n");
}

// ---------------------------------------------------------------------------
// Shared helper: navigate to a table cell
// ---------------------------------------------------------------------------

/**
 * Navigate to a specific cell in a table block and return its first paragraph
 * element together with the cell's children array (the paragraph's parent),
 * which an untracked multi-line replacement needs in order to splice in
 * additional sibling paragraphs.
 */
function getTableCellParagraph(
  refs: BlockRef[],
  blockIndex: number,
  rowIndex: number,
  colIndex: number,
): { paraEl: XNode; cellChildren: XNode[] } {
  const ref = resolveBlockRef(refs, blockIndex);

  if (!ref.element["w:tbl"]) {
    throw new EngineError(ErrorCode.NOT_A_TABLE, `Block ${blockIndex} is not a table.`);
  }

  const tblChildren = ref.element["w:tbl"] as XNode[];
  const rows = findAll(tblChildren, "w:tr");

  if (!Number.isInteger(rowIndex) || rowIndex < 0 || rowIndex >= rows.length) {
    throw new EngineError(ErrorCode.INDEX_OUT_OF_RANGE, `Row index ${rowIndex} out of range (0–${rows.length - 1}).`);
  }

  const row = rows[rowIndex];
  const cells = findAll(row["w:tr"], "w:tc");

  if (!Number.isInteger(colIndex) || colIndex < 0 || colIndex >= cells.length) {
    throw new EngineError(ErrorCode.INDEX_OUT_OF_RANGE, `Column index ${colIndex} out of range (0–${cells.length - 1}).`);
  }

  const cell = cells[colIndex];
  const cellChildren = cell["w:tc"] as XNode[];
  const paraEl = cellChildren.find((c: XNode) => c["w:p"]);
  if (!paraEl) {
    throw new EngineError(ErrorCode.NOT_A_PARAGRAPH, `Cell [${rowIndex},${colIndex}] has no paragraph.`);
  }

  return { paraEl, cellChildren };
}

// ---------------------------------------------------------------------------
// 22. edit_table_cells
// ---------------------------------------------------------------------------

export interface EditTableCellItem {
  blockIndex: number;
  rowIndex: number;
  colIndex: number;
  newText: string;
}

export async function editTableCells(
  filePath: string,
  edits: EditTableCellItem[],
  trackChanges: boolean = true,
  author: string = "Claude",
): Promise<string> {
  if (edits.length === 0) {
    return `No cell edits to apply.`;
  }

  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);
    const parsed = await parseDocXml(handle);
    const body = getBody(parsed);
    const refs = enumerateBlockRefs(body);

    // Validate all cells upfront and collect paragraph references
    const targets: { paraEl: XNode; cellChildren: XNode[]; newText: string; edit: EditTableCellItem }[] = [];
    for (const edit of edits) {
      const { paraEl, cellChildren } = getTableCellParagraph(refs, edit.blockIndex, edit.rowIndex, edit.colIndex);
      targets.push({ paraEl, cellChildren, newText: edit.newText, edit });
    }

    if (trackChanges) {
      for (const { paraEl, edit } of targets) {
        if (paragraphHasRevisions(paraEl["w:p"] as XNode[])) {
          throwPendingRevisions(`Table cell at block ${edit.blockIndex} row ${edit.rowIndex} col ${edit.colIndex}`);
        }
      }
    }

    const ctx = trackChanges
      ? newRevisionContext((await scanMaxIdAcrossParts(handle, parsed)) + 1, author)
      : null;

    for (const { paraEl, cellChildren, newText } of targets) {
      if (ctx) {
        // Tracked: diff-replace the first paragraph in place (del old + ins new,
        // "\n" → soft break). Tracked paragraph-mark insertion is avoided.
        replaceParagraphText(paraEl, newText, ctx);
      } else {
        // Untracked: replace the entire cell so multi-line input lands as real
        // paragraphs and re-edits leave no residue.
        replaceCellContentUntracked(cellChildren, newText);
      }
    }

    serializeDocXml(handle, parsed);
    await saveDocx(handle);

    const mode = trackChanges ? " (tracked)" : "";
    return `Updated ${edits.length} table cell(s) in ${path.basename(filePath)}${mode}.`;
  });
}

// ---------------------------------------------------------------------------
// 22b. Per-paragraph table cell editing (edit / delete / insert)
//
// These address an individual <w:p> inside a <w:tc> by a cell-local paragraph
// index (counting only the cell's w:p children, 0-based). Row/column indices
// are physical w:tc positions — the same convention edit_table_cells uses; a
// horizontally (gridSpan) or vertically (vMerge) merged cell is one physical
// position. Use read_table_cell to see a cell's paragraph indices.
// ---------------------------------------------------------------------------

/** Resolve a cell to its <w:tc> children array, validating every index. */
function getCellChildren(
  refs: BlockRef[],
  blockIndex: number,
  rowIndex: number,
  colIndex: number,
): XNode[] {
  const ref = resolveBlockRef(refs, blockIndex);
  if (!ref.element["w:tbl"]) {
    throw new EngineError(ErrorCode.NOT_A_TABLE, `Block ${blockIndex} is not a table.`);
  }
  const rows = findAll(ref.element["w:tbl"] as XNode[], "w:tr");
  if (!Number.isInteger(rowIndex) || rowIndex < 0 || rowIndex >= rows.length) {
    throw new EngineError(ErrorCode.INDEX_OUT_OF_RANGE, `Row index ${rowIndex} out of range (0–${rows.length - 1}).`);
  }
  const cells = findAll(rows[rowIndex]["w:tr"] as XNode[], "w:tc");
  if (!Number.isInteger(colIndex) || colIndex < 0 || colIndex >= cells.length) {
    throw new EngineError(ErrorCode.INDEX_OUT_OF_RANGE, `Column index ${colIndex} out of range (0–${cells.length - 1}).`);
  }
  return cells[colIndex]["w:tc"] as XNode[];
}

/** Resolve a cell's w:p element by cell-local paragraph index. */
function getCellParagraph(cellChildren: XNode[], paragraphIndex: number, where: string): XNode {
  const paras = findAll(cellChildren, "w:p");
  if (!Number.isInteger(paragraphIndex) || paragraphIndex < 0 || paragraphIndex >= paras.length) {
    throw new EngineError(ErrorCode.INDEX_OUT_OF_RANGE, `Paragraph index ${paragraphIndex} out of range (0–${paras.length - 1}) in ${where}.`);
  }
  return paras[paragraphIndex];
}

export interface EditTableParagraphItem {
  blockIndex: number;
  rowIndex: number;
  colIndex: number;
  paragraphIndex: number;
  newText: string;
}

export async function editTableParagraphs(
  filePath: string,
  edits: EditTableParagraphItem[],
  trackChanges: boolean = true,
  author: string = "Claude",
): Promise<string> {
  if (edits.length === 0) {
    return `No table paragraph edits to apply.`;
  }

  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);
    const parsed = await parseDocXml(handle);
    const body = getBody(parsed);
    const refs = enumerateBlockRefs(body);

    // Resolve and validate every target paragraph upfront.
    const targets: XNode[] = [];
    for (const e of edits) {
      const cellChildren = getCellChildren(refs, e.blockIndex, e.rowIndex, e.colIndex);
      const where = `cell [${e.rowIndex},${e.colIndex}] of block ${e.blockIndex}`;
      const paraEl = getCellParagraph(cellChildren, e.paragraphIndex, where);
      if (trackChanges && paragraphHasRevisions(paraEl["w:p"] as XNode[])) {
        throwPendingRevisions(`Paragraph ${e.paragraphIndex} in ${where}`);
      }
      targets.push(paraEl);
    }

    const ctx = trackChanges
      ? newRevisionContext((await scanMaxIdAcrossParts(handle, parsed)) + 1, author)
      : null;

    // Collapse duplicate edits to the same paragraph (last-write-wins).
    const finalText = new Map<XNode, string>();
    edits.forEach((e, i) => finalText.set(targets[i], e.newText));
    for (const [paraEl, newText] of finalText) {
      replaceParagraphText(paraEl, newText, ctx);
    }

    serializeDocXml(handle, parsed);
    await saveDocx(handle);

    const mode = trackChanges ? " (tracked)" : "";
    return `Updated ${finalText.size} table paragraph(s) in ${path.basename(filePath)}${mode}.`;
  });
}

export interface DeleteTableParagraphItem {
  blockIndex: number;
  rowIndex: number;
  colIndex: number;
  paragraphIndex: number;
}

export async function deleteTableParagraphs(
  filePath: string,
  targets: DeleteTableParagraphItem[],
  trackChanges: boolean = true,
  author: string = "Claude",
): Promise<string> {
  if (targets.length === 0) {
    return `No table paragraphs to delete.`;
  }

  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);
    const parsed = await parseDocXml(handle);
    const body = getBody(parsed);
    const refs = enumerateBlockRefs(body);

    // Resolve and validate upfront, deduplicating by paragraph identity.
    const resolved: { cellChildren: XNode[]; paraEl: XNode }[] = [];
    const seen = new Set<XNode>();
    for (const t of targets) {
      const cellChildren = getCellChildren(refs, t.blockIndex, t.rowIndex, t.colIndex);
      const where = `cell [${t.rowIndex},${t.colIndex}] of block ${t.blockIndex}`;
      const paraEl = getCellParagraph(cellChildren, t.paragraphIndex, where);
      if (trackChanges && paragraphHasRevisions(paraEl["w:p"] as XNode[])) {
        throwPendingRevisions(`Paragraph ${t.paragraphIndex} in ${where}`);
      }
      if (seen.has(paraEl)) continue;
      seen.add(paraEl);
      resolved.push({ cellChildren, paraEl });
    }

    if (trackChanges) {
      const ctx = newRevisionContext((await scanMaxIdAcrossParts(handle, parsed)) + 1, author);
      // Mark the paragraph's runs and its paragraph mark as deleted. The element
      // stays in place, so the cell keeps at least one paragraph structurally.
      for (const { paraEl } of resolved) {
        markParagraphRunsAsDeleted(paraEl["w:p"] as XNode[], ctx);
      }
    } else {
      // Hard delete by reference (order-independent).
      for (const { cellChildren, paraEl } of resolved) {
        const idx = cellChildren.indexOf(paraEl);
        if (idx !== -1) cellChildren.splice(idx, 1);
      }
      // A <w:tc> must END with a <w:p> (ECMA-376 CT_Tc) — stronger than "has ≥1
      // w:p". Restore a blank paragraph when the cell now has none OR ends in a
      // nested <w:tbl> (e.g. the deleted paragraph was the one terminating the
      // cell after a nested table).
      const affectedCells = new Set(resolved.map((r) => r.cellChildren));
      for (const cellChildren of affectedCells) {
        if (!cellChildren.some((c: XNode) => c["w:p"])) {
          cellChildren.push(el("w:p", []));
        } else {
          ensureCellEndsWithParagraph(cellChildren);
        }
      }
    }

    serializeDocXml(handle, parsed);
    await saveDocx(handle);

    const mode = trackChanges ? " (tracked)" : "";
    return `Deleted ${resolved.length} table paragraph(s) from ${path.basename(filePath)}${mode}.`;
  });
}

/** Resolve build options for a cell-local insert (copy_format_from indexes within the same cell). */
function resolveCellInsertOpts(
  paras: XNode[],
  numId?: number,
  numLevel?: number,
  copyFormatFrom?: number,
): BuildParagraphOptions | undefined {
  if (copyFormatFrom !== undefined) {
    if (!Number.isInteger(copyFormatFrom) || copyFormatFrom < 0 || copyFormatFrom >= paras.length) {
      throw new EngineError(
        ErrorCode.INDEX_OUT_OF_RANGE,
        `copy_format_from index ${copyFormatFrom} out of range (0–${paras.length - 1}) for the target cell.`,
      );
    }
    const pPr = findOne(paras[copyFormatFrom]["w:p"] as XNode[], "w:pPr");
    return pPr ? { sourcePPr: pPr } : undefined;
  }
  if (numId !== undefined) {
    return { numId, numLevel };
  }
  return undefined;
}

export interface InsertTableParagraphItem {
  blockIndex: number;
  rowIndex: number;
  colIndex: number;
  /** Cell-local paragraph index to insert before; out-of-range / -1 appends to the cell. */
  position: number;
  text: string;
  style?: string;
  numId?: number;
  numLevel?: number;
  /** Cell-local paragraph index whose pPr to deep-copy (overrides style/numId/numLevel). */
  copyFormatFrom?: number;
}

export async function insertTableParagraphs(
  filePath: string,
  inserts: InsertTableParagraphItem[],
  trackChanges: boolean = true,
  author: string = "Claude",
): Promise<string> {
  if (inserts.length === 0) {
    return `No table paragraphs to insert.`;
  }

  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);
    const parsed = await parseDocXml(handle);
    const body = getBody(parsed);
    const refs = enumerateBlockRefs(body);

    const ctx = trackChanges
      ? newRevisionContext((await scanMaxIdAcrossParts(handle, parsed)) + 1, author)
      : null;

    // Resolve every insert against the ORIGINAL cell paragraph list: the anchor
    // is the existing paragraph the new one goes before (null = append to cell).
    // Anchoring by element reference keeps multiple inserts into the same cell
    // correct and in array order, without the reverse-order quirk of body
    // insert_paragraphs.
    interface Resolved {
      cellChildren: XNode[];
      anchor: XNode | null;
      newParas: XNode[];
    }
    const resolved: Resolved[] = [];
    for (const ins of inserts) {
      const cellChildren = getCellChildren(refs, ins.blockIndex, ins.rowIndex, ins.colIndex);
      const paras = findAll(cellChildren, "w:p");
      const opts = resolveCellInsertOpts(paras, ins.numId, ins.numLevel, ins.copyFormatFrom);
      // buildNewParagraph returns one or more <w:p> (untracked multi-line input
      // splits on "\n"); splice them all in at the resolved spot.
      const newParas = buildNewParagraph(ins.text, ins.style, ctx, opts);
      // F4: a NON-INTEGER position is rejected (matching the M7 integer policy
      // and the spliceNewParagraph guard) rather than silently appending. A -1
      // or out-of-range INTEGER position still appends (anchor null), unchanged.
      if (!Number.isInteger(ins.position)) {
        throw new EngineError(
          ErrorCode.INDEX_OUT_OF_RANGE,
          `Insert position ${ins.position} must be an integer cell-paragraph index (use -1 to append).`,
        );
      }
      const anchor =
        ins.position >= 0 && ins.position < paras.length ? paras[ins.position] : null;
      resolved.push({ cellChildren, anchor, newParas });
    }

    for (const { cellChildren, anchor, newParas } of resolved) {
      if (anchor) {
        const idx = cellChildren.indexOf(anchor);
        cellChildren.splice(idx === -1 ? cellChildren.length : idx, 0, ...newParas);
      } else {
        // Append after the cell's last paragraph. If the cell has no paragraph
        // at all, append at the very end (never before w:tcPr, which must stay
        // first; a cell's content also has to end with a paragraph).
        let lastP = -1;
        for (let i = 0; i < cellChildren.length; i++) {
          if (cellChildren[i]["w:p"]) lastP = i;
        }
        const insertAt = lastP === -1 ? cellChildren.length : lastP + 1;
        cellChildren.splice(insertAt, 0, ...newParas);
      }
    }

    serializeDocXml(handle, parsed);
    await saveDocx(handle);

    const mode = trackChanges ? " (tracked)" : "";
    return `Inserted ${inserts.length} table paragraph(s) in ${path.basename(filePath)}${mode}.`;
  });
}

// ---------------------------------------------------------------------------
// 23. read_footnotes
// ---------------------------------------------------------------------------

export async function readFootnotes(filePath: string): Promise<string> {
  const handle = await openDocx(filePath);
  const xml = await handle.zip.file("word/footnotes.xml")?.async("string");

  if (!xml) {
    return `No footnotes found in ${path.basename(filePath)}.`;
  }

  const parsed: XNode[] = parser.parse(xml);
  const rootEl = parsed.find((n: XNode) => n["w:footnotes"]);
  if (!rootEl) {
    return `No footnotes found in ${path.basename(filePath)}.`;
  }

  const footnotes = findAll(rootEl["w:footnotes"], "w:footnote");
  const contentFootnotes = footnotes.filter((fn: XNode) => {
    const t = attr(fn, "w:type");
    return !t || (t !== "separator" && t !== "continuationSeparator");
  });

  if (contentFootnotes.length === 0) {
    return `No footnotes found in ${path.basename(filePath)}.`;
  }

  const lines: string[] = [];
  for (const fn of contentFootnotes) {
    const id = attr(fn, "w:id") ?? "?";
    const fnChildren = fn["w:footnote"] as XNode[];
    const texts: string[] = [];
    for (const child of fnChildren) {
      if (child["w:p"]) {
        texts.push(extractParagraphText(child["w:p"]));
      }
    }
    lines.push(`[Footnote ${id}] ${texts.join(" ").trim()}`);
  }

  return lines.join("\n");
}

// ---------------------------------------------------------------------------
// 24. list_images
// ---------------------------------------------------------------------------

export async function listImages(filePath: string): Promise<string> {
  const result = await listImagesStructured(filePath);

  if (result.totalImages === 0) {
    // Emit the structured block (zero total / empty array) so the contract holds
    // even when there are no images (L2).
    return `No images found in ${result.file}.\n\n` + jsonBlock(result);
  }

  let output = `Images in ${result.file} (${result.totalImages}):\n\n`;
  for (const img of result.images) {
    output += `[Block ${img.blockIndex}] ${img.name || img.filename}\n`;
    output += `  File: ${img.filename}\n`;
    output += `  Type: ${img.contentType}\n`;
    output += `  Size: ${img.sizeBytes} bytes\n`;
    if (img.widthEmu || img.heightEmu) {
      const wMm = Math.round((img.widthEmu / 914400) * 25.4 * 10) / 10;
      const hMm = Math.round((img.heightEmu / 914400) * 25.4 * 10) / 10;
      output += `  Dimensions: ${wMm} × ${hMm} mm (${img.widthEmu} × ${img.heightEmu} EMU)\n`;
    }
    if (img.altText) {
      output += `  Alt text: ${img.altText}\n`;
    }
    output += "\n";
  }

  output += jsonBlock(result);
  return output;
}

export async function listImagesStructured(
  filePath: string,
): Promise<ListImagesResult> {
  const handle = await openDocx(filePath);
  const images = await scanImages(handle);

  return {
    file: path.basename(filePath),
    totalImages: images.length,
    images,
  };
}
