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
  blockBodyIndices,
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
  replaceInParagraph,
  replaceInParagraphTracked,
  paragraphHasRevisions,
  scanMaxId,
  newRevisionContext,
  allocRevId,
  getRunRPr,
  makeTextRun,
  makeDelTextRun,
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

// ===========================================================================================
// PUBLIC API
// ===========================================================================================

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
    for (const m of xml.matchAll(/\bw:id="(\d+)"/g)) {
      const v = parseInt(m[1], 10);
      if (!isNaN(v) && v > max) max = v;
    }
  }
  return max;
}

// ---------------------------------------------------------------------------
// 1. read_document
// ---------------------------------------------------------------------------

export async function readDocument(
  filePath: string,
  startParagraph?: number,
  endParagraph?: number,
  showRevisions: boolean = false,
): Promise<string> {
  const handle = await openDocx(filePath);
  const parsed = await parseDocXml(handle);
  const body = getBody(parsed);
  const blocks = enumerateBlocks(body, showRevisions);

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

  output += "\n\n<json>\n" + JSON.stringify(info) + "\n</json>";
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
    return `No matches found for "${query}" in ${result.file}.`;
  }

  let output = `Found ${result.totalMatches} match(es) for "${query}":\n\n`;
  for (const m of result.matches) {
    output += `[Block ${m.blockIndex}] ${m.context}\n`;
  }

  output += "\n\n<json>\n" + JSON.stringify(result) + "\n</json>";
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
}

/** searchTextStructured の戻り値型 */
export interface SearchTextResult {
  file: string;
  query: string;
  totalMatches: number;
  matches: SearchMatch[];
}

/**
 * searchText と同じロジックだが、テキストではなく構造化オブジェクトを返す。
 */
export async function searchTextStructured(
  filePath: string,
  query: string,
  caseSensitive: boolean = false,
): Promise<SearchTextResult> {
  const handle = await openDocx(filePath);
  const parsed = await parseDocXml(handle);
  const body = getBody(parsed);
  const blocks = enumerateBlocks(body);

  const searchStr = caseSensitive ? query : query.toLowerCase();
  const matches: SearchMatch[] = [];
  let totalMatches = 0;

  for (const b of blocks) {
    const compare = caseSensitive ? b.text : b.text.toLowerCase();
    if (compare.includes(searchStr)) {
      // ブロック内の出現回数をカウント
      let occurrences = 0;
      let searchFrom = 0;
      while (true) {
        const idx = compare.indexOf(searchStr, searchFrom);
        if (idx === -1) break;
        occurrences++;
        searchFrom = idx + searchStr.length;
      }
      totalMatches += occurrences;

      // 最初のマッチ周辺のコンテキストを抽出
      const matchIdx = compare.indexOf(searchStr);
      const ctxStart = Math.max(0, matchIdx - 40);
      const ctxEnd = Math.min(b.text.length, matchIdx + query.length + 40);
      const ctx =
        (ctxStart > 0 ? "..." : "") +
        b.text.substring(ctxStart, ctxEnd) +
        (ctxEnd < b.text.length ? "..." : "");
      matches.push({
        blockIndex: b.index,
        occurrences,
        context: ctx,
        fullText: b.text,
      });
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

    if (trackChanges) {
      const conflict = findRevisionConflict(body, (pChildren) => {
        if (!paragraphHasRevisions(pChildren)) return false;
        const text = extractParagraphText(pChildren);
        return items.some((it) => {
          const cs = it.caseSensitive ?? false;
          const haystack = cs ? text : text.toLowerCase();
          const needle = cs ? it.search : it.search.toLowerCase();
          return haystack.includes(needle);
        });
      });
      if (conflict) throwPendingRevisions(conflict.description);
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
      const hfFiles = getHeaderFooterFiles(handle);
      let hfMaxId = ctx ? ctx.nextId - 1 : 0;

      for (const hfFile of hfFiles) {
        const hfXml = await handle.zip.file(hfFile)?.async("string");
        if (!hfXml) continue;
        const hfParsed: XNode[] = parser.parse(hfXml);
        const rootEl = hfParsed.find((n: XNode) => n["w:hdr"] || n["w:ftr"]);
        if (!rootEl) continue;
        const hfChildren = rootEl["w:hdr"] ?? rootEl["w:ftr"];

        const ctx2 = trackChanges ? newRevisionContext(hfMaxId + 1, author) : null;
        applyToChildren(hfChildren, ctx2);
        if (ctx2) hfMaxId = ctx2.nextId - 1;
        handle.zip.file(hfFile, builder.build(hfParsed));
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

/** Replace the text content of a w:p element, preserving pPr, structural elements, and optionally tracking changes. */
function replaceParagraphText(
  element: XNode,
  newText: string,
  ctx: RevisionContext | null,
): void {
  const pChildren = element["w:p"] as XNode[];
  const pPr = findOne(pChildren, "w:pPr");

  // Collect structural elements that must be preserved (bookmarks, comment ranges, drawings)
  const structuralElements: XNode[] = [];
  for (const child of pChildren) {
    if (
      child["w:bookmarkStart"] !== undefined ||
      child["w:bookmarkEnd"] !== undefined ||
      child["w:commentRangeStart"] !== undefined ||
      child["w:commentRangeEnd"] !== undefined ||
      child["w:commentReference"] !== undefined
    ) {
      structuralElements.push(child);
    } else if (child["w:r"]) {
      const runC = child["w:r"] as XNode[];
      const hasDrawing = runC.some((rc) => rc["w:drawing"] !== undefined);
      if (hasDrawing) structuralElements.push(child);
    }
  }

  if (ctx) {
    // Extract current full text and first rPr from existing runs
    let oldText = "";
    let firstRPr: XNode | null = null;
    for (const child of pChildren) {
      if (child["w:r"]) {
        const runC = child["w:r"] as XNode[];
        const hasDrawing = runC.some((rc) => rc["w:drawing"] !== undefined);
        if (hasDrawing) continue;
        const rPr = getRunRPr(runC);
        if (!firstRPr && rPr) firstRPr = rPr;
        oldText += extractRunText(runC);
      } else if (child["w:ins"]) {
        for (const insChild of child["w:ins"]) {
          if (insChild["w:r"]) {
            const runC = insChild["w:r"] as XNode[];
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
    if (diff.prefix) newChildren.push(makeTextRun(diff.prefix, firstRPr));
    if (diff.oldMiddle) newChildren.push(wrapInDel([makeDelTextRun(diff.oldMiddle, firstRPr)], ctx));
    if (diff.newMiddle) newChildren.push(wrapInIns([makeTextRun(diff.newMiddle, firstRPr)], ctx));
    if (diff.suffix) newChildren.push(makeTextRun(diff.suffix, firstRPr));
    element["w:p"] = newChildren;
  } else {
    const newRun = el("w:r", [
      el("w:t", [textNode(newText)], { "xml:space": "preserve" }),
    ]);
    const newChildren: XNode[] = [];
    if (pPr) newChildren.push(pPr);
    newChildren.push(...structuralElements);
    newChildren.push(newRun);
    element["w:p"] = newChildren;
  }
}

// ---------------------------------------------------------------------------
// 5. edit_paragraphs
// ---------------------------------------------------------------------------

export interface EditParagraphItem {
  paragraphIndex: number;
  newText: string;
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
    const body = getBody(parsed);
    const bodyIdxs = blockBodyIndices(body);

    // Validate all indices upfront
    for (const edit of edits) {
      if (edit.paragraphIndex < 0 || edit.paragraphIndex >= bodyIdxs.length) {
        throw new EngineError(ErrorCode.INDEX_OUT_OF_RANGE, `Paragraph index ${edit.paragraphIndex} out of range (0–${bodyIdxs.length - 1}).`);
      }
      const element = body[bodyIdxs[edit.paragraphIndex]];
      if (!element["w:p"]) {
        throw new EngineError(ErrorCode.NOT_A_PARAGRAPH, `Block ${edit.paragraphIndex} is not a paragraph (it may be a table).`);
      }
      if (trackChanges && paragraphHasRevisions(element["w:p"] as XNode[])) {
        throwPendingRevisions(`Paragraph at index ${edit.paragraphIndex}`);
      }
    }

    const ctx = trackChanges
      ? newRevisionContext((await scanMaxIdAcrossParts(handle, parsed)) + 1, author)
      : null;

    for (const edit of edits) {
      replaceParagraphText(body[bodyIdxs[edit.paragraphIndex]], edit.newText, ctx);
    }

    serializeDocXml(handle, parsed);
    await saveDocx(handle);

    const mode = trackChanges ? " (tracked)" : "";
    return `Updated ${edits.length} paragraph(s) in ${path.basename(filePath)}${mode}.`;
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

/** Build a w:p element with optional style, numbering, copied format, and tracked insertion context. */
function buildNewParagraph(
  text: string,
  style: string | undefined,
  ctx: RevisionContext | null,
  opts?: BuildParagraphOptions,
): XNode {
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
  } else {
    for (let i = 0; i < lines.length; i++) {
      if (i > 0) pChildren.push(el("w:r", [el("w:br")]));
      if (lines[i]) {
        pChildren.push(
          el("w:r", [
            el("w:t", [textNode(lines[i])], { "xml:space": "preserve" }),
          ]),
        );
      }
    }
  }

  return el("w:p", pChildren);
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

    // Always strip stale revision markers from the cloned rPr —
    // they belong to the source paragraph, not the new one.
    const existingRPr = findOne(pPrChildren, "w:rPr");
    if (existingRPr) {
      const rPrChildren = existingRPr["w:rPr"] as XNode[];
      for (let i = rPrChildren.length - 1; i >= 0; i--) {
        if (rPrChildren[i]["w:ins"] !== undefined || rPrChildren[i]["w:del"] !== undefined) {
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

/** Insert a paragraph element into body at the given position, handling sectPr and append. */
function spliceNewParagraph(
  body: XNode[],
  bodyIdxs: number[],
  position: number,
  newPara: XNode,
): void {
  if (position < 0 || position >= bodyIdxs.length) {
    const sectPrIdx = body.findIndex((n: XNode) => n["w:sectPr"]);
    if (sectPrIdx !== -1) {
      body.splice(sectPrIdx, 0, newPara);
    } else {
      body.push(newPara);
    }
  } else {
    body.splice(bodyIdxs[position], 0, newPara);
  }
}

// ---------------------------------------------------------------------------
// 6. insert_paragraphs
// ---------------------------------------------------------------------------

/** Resolve BuildParagraphOptions from user-facing parameters. */
function resolveInsertOpts(
  body: XNode[],
  bodyIdxs: number[],
  numId?: number,
  numLevel?: number,
  copyFormatFrom?: number,
): BuildParagraphOptions | undefined {
  if (copyFormatFrom !== undefined) {
    if (copyFormatFrom < 0 || copyFormatFrom >= bodyIdxs.length) {
      throw new EngineError(
        ErrorCode.INDEX_OUT_OF_RANGE,
        `copy_format_from index ${copyFormatFrom} out of range (0–${bodyIdxs.length - 1}).`,
      );
    }
    const srcEl = body[bodyIdxs[copyFormatFrom]];
    if (!srcEl["w:p"]) {
      throw new EngineError(
        ErrorCode.NOT_A_PARAGRAPH,
        `Block ${copyFormatFrom} is not a paragraph (copy_format_from must reference a paragraph).`,
      );
    }
    const pPr = findOne(srcEl["w:p"] as XNode[], "w:pPr");
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
  position: number;
  style?: string;
  numId?: number;
  numLevel?: number;
  copyFormatFrom?: number;
}

export async function insertParagraphs(
  filePath: string,
  items: InsertParagraphItem[],
  trackChanges: boolean = true,
  author: string = "Claude",
): Promise<string> {
  if (items.length === 0) {
    return `No paragraphs to insert.`;
  }

  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);
    const parsed = await parseDocXml(handle);
    const body = getBody(parsed);

    const ctx = trackChanges
      ? newRevisionContext((await scanMaxIdAcrossParts(handle, parsed)) + 1, author)
      : null;

    // Sort by position descending so higher-index inserts don't shift lower ones.
    // Append-position items (position < 0 or out of range) sort to the front
    // so they are processed before specific-position items.
    const bodyIdxsForSort = blockBodyIndices(body);
    const sorted = [...items].sort((a, b) => {
      const posA = a.position < 0 || a.position >= bodyIdxsForSort.length ? Infinity : a.position;
      const posB = b.position < 0 || b.position >= bodyIdxsForSort.length ? Infinity : b.position;
      if (posA === posB) return 0;
      return posB - posA;
    });

    // Resolve copy_format_from before any inserts (indices refer to original document)
    const resolvedOpts = new Map<InsertParagraphItem, BuildParagraphOptions | undefined>();
    for (const item of sorted) {
      const opts = resolveInsertOpts(body, bodyIdxsForSort, item.numId, item.numLevel, item.copyFormatFrom);
      resolvedOpts.set(item, opts);
    }

    for (const item of sorted) {
      const newPara = buildNewParagraph(item.text, item.style, ctx, resolvedOpts.get(item));
      // Recalculate indices after each splice to avoid stale mapping
      const currentBodyIdxs = blockBodyIndices(body);
      spliceNewParagraph(body, currentBodyIdxs, item.position, newPara);
    }

    serializeDocXml(handle, parsed);
    await saveDocx(handle);

    const mode = trackChanges ? " (tracked)" : "";
    return `Inserted ${items.length} paragraph(s) in ${path.basename(filePath)}${mode}.`;
  });
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

  // Convert all runs to w:del with w:delText
  const delRuns: XNode[] = [];
  const indicesToRemove: number[] = [];
  for (let i = 0; i < pChildren.length; i++) {
    const child = pChildren[i];
    if (child["w:r"]) {
      const runC = child["w:r"] as XNode[];
      const rPr = getRunRPr(runC);
      const runText = extractRunText(runC);
      if (runText) {
        delRuns.push(makeDelTextRun(runText, rPr));
      }
      indicesToRemove.push(i);
    }
  }

  // Remove original runs (reverse order to preserve indices)
  for (let i = indicesToRemove.length - 1; i >= 0; i--) {
    pChildren.splice(indicesToRemove[i], 1);
  }

  // Append the w:del element with all deleted runs
  if (delRuns.length > 0) {
    pChildren.push(wrapInDel(delRuns, ctx));
  }
}

/** Mark a body-level block element (paragraph or table) as tracked deletion. */
function markBlockAsDeleted(element: XNode, ctx: RevisionContext): void {
  if (element["w:p"]) {
    markParagraphRunsAsDeleted(element["w:p"] as XNode[], ctx);
  } else if (element["w:tbl"]) {
    const rows = findAll(element["w:tbl"], "w:tr");
    for (const row of rows) {
      const cells = findAll(row["w:tr"], "w:tc");
      for (const cell of cells) {
        const paras = findAll(cell["w:tc"], "w:p");
        for (const p of paras) {
          markParagraphRunsAsDeleted(p["w:p"] as XNode[], ctx);
        }
      }
    }
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
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);
    const parsed = await parseDocXml(handle);
    const body = getBody(parsed);
    const bodyIdxs = blockBodyIndices(body);

    // Deduplicate indices upfront
    const unique = [...new Set(paragraphIndices)];

    // Validate all indices
    for (const idx of unique) {
      if (idx < 0 || idx >= bodyIdxs.length) {
        throw new EngineError(ErrorCode.INDEX_OUT_OF_RANGE, `Paragraph index ${idx} out of range (0–${bodyIdxs.length - 1}).`);
      }
    }

    if (trackChanges) {
      // Reject if any targeted paragraph already has tracked-change markup.
      // Tables are checked cell-by-cell. Paragraph-mark deletion via
      // pPr > rPr > w:del also counts.
      for (const idx of unique) {
        const element = body[bodyIdxs[idx]];
        if (element["w:p"] && paragraphHasRevisions(element["w:p"] as XNode[])) {
          throwPendingRevisions(`Block at index ${idx}`);
        } else if (element["w:tbl"]) {
          let conflict = false;
          forEachParagraphInTable(element["w:tbl"], (pChildren) => {
            if (paragraphHasRevisions(pChildren)) conflict = true;
          });
          if (conflict) throwPendingRevisions(`Table at index ${idx}`);
        }
      }

      const maxId = await scanMaxIdAcrossParts(handle, parsed);
      const ctx = newRevisionContext(maxId + 1, author);

      for (const idx of unique) {
        markBlockAsDeleted(body[bodyIdxs[idx]], ctx);
      }
    } else {
      // Hard delete: sort descending to avoid index shifting
      const sorted = unique.sort((a, b) => b - a);
      for (const idx of sorted) {
        body.splice(bodyIdxs[idx], 1);
      }
    }

    serializeDocXml(handle, parsed);
    await saveDocx(handle);

    const mode = trackChanges ? " (tracked)" : "";
    return `Deleted ${unique.length} block(s) from ${path.basename(filePath)}${mode}.`;
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
  groups: Array<{ indices: number[]; format: ParagraphFormat }>,
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);
    const parsed = await parseDocXml(handle);
    const body = getBody(parsed);
    const bodyIdxs = blockBodyIndices(body);

    // Validate all indices upfront
    let totalCount = 0;
    for (const group of groups) {
      for (const idx of group.indices) {
        if (idx < 0 || idx >= bodyIdxs.length) {
          throw new EngineError(ErrorCode.INDEX_OUT_OF_RANGE, `Paragraph index ${idx} out of range (0–${bodyIdxs.length - 1}).`);
        }
        const element = body[bodyIdxs[idx]];
        if (!element["w:p"]) {
          throw new EngineError(ErrorCode.NOT_A_PARAGRAPH, `Block ${idx} is not a paragraph.`);
        }
        totalCount++;
      }
    }

    // Apply formatting
    for (const group of groups) {
      for (const idx of group.indices) {
        const element = body[bodyIdxs[idx]];
        applyParagraphFormat(element["w:p"], group.format);
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

    // Add comment to comments.xml
    const lines = commentText.split("\n");
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

      const lines = c.comment_text.split("\n");
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
    return `No comments found in ${result.file}.`;
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

  output += "\n<json>\n" + JSON.stringify(result) + "\n</json>";
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

    const lines = commentText.split("\n");
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
      bodyXml += `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>${escapeXml(title)}</w:t></w:r></w:p>\n`;
    }

    if (content) {
      const lines = content.split("\n");
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

export async function insertTable(
  filePath: string,
  position: number,
  rows: number,
  cols: number,
  data?: string[][],
): Promise<string> {
  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);
    const parsed = await parseDocXml(handle);
    const body = getBody(parsed);
    const bodyIdxs = blockBodyIndices(body);

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

    // Table rows
    for (let r = 0; r < rows; r++) {
      const cellNodes: XNode[] = [];
      for (let c = 0; c < cols; c++) {
        const cellText = data?.[r]?.[c] ?? "";
        cellNodes.push(
          el("w:tc", [
            el("w:tcPr", [el("w:tcW", [], { "w:w": "0", "w:type": "auto" })]),
            el("w:p", [
              el("w:r", [
                el("w:t", [textNode(cellText)], { "xml:space": "preserve" }),
              ]),
            ]),
          ]),
        );
      }
      tblChildren.push(el("w:tr", cellNodes));
    }

    const tblNode = el("w:tbl", tblChildren);

    // Insert at position
    if (position < 0 || position >= bodyIdxs.length) {
      const sectPrIdx = body.findIndex((n: XNode) => n["w:sectPr"]);
      if (sectPrIdx !== -1) {
        body.splice(sectPrIdx, 0, tblNode);
      } else {
        body.push(tblNode);
      }
    } else {
      const bodyIdx = bodyIdxs[position];
      body.splice(bodyIdx, 0, tblNode);
    }

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
  paragraphIndex: number;
  level: number;
}

export async function setHeadings(
  filePath: string,
  items: SetHeadingItem[],
): Promise<string> {
  if (items.length === 0) {
    return `No headings to set.`;
  }

  // Validate levels upfront (before acquiring file lock)
  for (const item of items) {
    if (item.level < 1 || item.level > 9) {
      throw new EngineError(ErrorCode.INVALID_PARAMETER, `Heading level must be between 1 and 9. Got: ${item.level}.`);
    }
  }

  return withFileLock(filePath, async () => {
    const handle = await openDocx(filePath);
    const parsed = await parseDocXml(handle);
    const body = getBody(parsed);
    const bodyIdxs = blockBodyIndices(body);

    // Validate all indices upfront
    for (const item of items) {
      if (item.paragraphIndex < 0 || item.paragraphIndex >= bodyIdxs.length) {
        throw new EngineError(ErrorCode.INDEX_OUT_OF_RANGE, `Paragraph index ${item.paragraphIndex} out of range (0–${bodyIdxs.length - 1}).`);
      }
      const element = body[bodyIdxs[item.paragraphIndex]];
      if (!element["w:p"]) {
        throw new EngineError(ErrorCode.NOT_A_PARAGRAPH, `Block ${item.paragraphIndex} is not a paragraph.`);
      }
    }

    for (const item of items) {
      applyHeadingLevel(body[bodyIdxs[item.paragraphIndex]]["w:p"] as XNode[], item.level);
    }

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

export async function setPageLayout(
  filePath: string,
  options: PageLayoutOptions,
): Promise<string> {
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
// Shared helper: replace cell text content in-place
// ---------------------------------------------------------------------------

/** Replace the text content of a cell's first paragraph, preserving pPr and structural elements. */
function replaceCellText(
  paraEl: XNode,
  newText: string,
  ctx: RevisionContext | null,
): void {
  // Delegate to replaceParagraphText which handles structural element preservation
  replaceParagraphText(paraEl, newText, ctx);
}

/** Navigate to a specific cell in a table block and return its first paragraph element. */
function getTableCellParagraph(
  body: XNode[],
  bodyIdxs: number[],
  blockIndex: number,
  rowIndex: number,
  colIndex: number,
): XNode {
  if (blockIndex < 0 || blockIndex >= bodyIdxs.length) {
    throw new EngineError(ErrorCode.INDEX_OUT_OF_RANGE, `Block index ${blockIndex} out of range (0–${bodyIdxs.length - 1}).`);
  }

  const element = body[bodyIdxs[blockIndex]];

  if (!element["w:tbl"]) {
    throw new EngineError(ErrorCode.NOT_A_TABLE, `Block ${blockIndex} is not a table.`);
  }

  const tblChildren = element["w:tbl"] as XNode[];
  const rows = findAll(tblChildren, "w:tr");

  if (rowIndex < 0 || rowIndex >= rows.length) {
    throw new EngineError(ErrorCode.INDEX_OUT_OF_RANGE, `Row index ${rowIndex} out of range (0–${rows.length - 1}).`);
  }

  const row = rows[rowIndex];
  const cells = findAll(row["w:tr"], "w:tc");

  if (colIndex < 0 || colIndex >= cells.length) {
    throw new EngineError(ErrorCode.INDEX_OUT_OF_RANGE, `Column index ${colIndex} out of range (0–${cells.length - 1}).`);
  }

  const cell = cells[colIndex];
  const cellChildren = cell["w:tc"] as XNode[];
  const paraEl = cellChildren.find((c: XNode) => c["w:p"]);
  if (!paraEl) {
    throw new EngineError(ErrorCode.NOT_A_PARAGRAPH, `Cell [${rowIndex},${colIndex}] has no paragraph.`);
  }

  return paraEl;
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
    const bodyIdxs = blockBodyIndices(body);

    // Validate all cells upfront and collect paragraph references
    const targets: { paraEl: XNode; newText: string; edit: EditTableCellItem }[] = [];
    for (const edit of edits) {
      const paraEl = getTableCellParagraph(body, bodyIdxs, edit.blockIndex, edit.rowIndex, edit.colIndex);
      targets.push({ paraEl, newText: edit.newText, edit });
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

    for (const { paraEl, newText } of targets) {
      replaceCellText(paraEl, newText, ctx);
    }

    serializeDocXml(handle, parsed);
    await saveDocx(handle);

    const mode = trackChanges ? " (tracked)" : "";
    return `Updated ${edits.length} table cell(s) in ${path.basename(filePath)}${mode}.`;
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
    return `No images found in ${result.file}.`;
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

  output += "<json>\n" + JSON.stringify(result) + "\n</json>";
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
