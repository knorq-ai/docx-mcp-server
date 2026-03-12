/**
 * Comment-related helpers: parsing, infrastructure, anchor matching, markers.
 */

import * as crypto from "crypto";
import {
  type XNode,
  parser,
  builder,
  attr,
  setAttr,
  findAll,
  findOne,
  el,
} from "./xml-helpers.js";
import type { DocxHandle } from "./docx-io.js";
import { extractParagraphText } from "./text.js";
import { forEachParagraphInTable } from "./docx-io.js";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface CommentInfo {
  id: string;
  author: string;
  date: string;
  text: string;
}

export interface BatchCommentInput {
  anchor_text: string;
  comment_text: string;
  author?: string;
}

export interface BatchCommentResult {
  anchor_text: string;
  status: "added" | "failed";
  comment_id?: number;
  error?: string;
  fuzzy_matched?: boolean;
}

// ---------------------------------------------------------------------------
// comments.xml parsing
// ---------------------------------------------------------------------------

export async function parseCommentsXml(handle: DocxHandle): Promise<XNode[]> {
  const xml = await handle.zip
    .file("word/comments.xml")
    ?.async("string");
  if (!xml) return [];
  return parser.parse(xml);
}

export function getCommentsArray(parsed: XNode[]): XNode[] {
  const root = parsed.find((n: XNode) => n["w:comments"]);
  if (!root) return [];
  return root["w:comments"];
}

export function getNextCommentId(commentsChildren: XNode[]): number {
  let maxId = -1;
  for (const c of findAll(commentsChildren, "w:comment")) {
    const id = parseInt(attr(c, "w:id") ?? "-1");
    if (id > maxId) maxId = id;
  }
  return maxId + 1;
}

// ---------------------------------------------------------------------------
// Infrastructure helpers (parsed XML, not string manipulation)
// ---------------------------------------------------------------------------

/** Add an Override entry to [Content_Types].xml if not already present. */
async function ensureContentTypeOverride(
  handle: DocxHandle,
  partName: string,
  contentType: string,
): Promise<void> {
  const ctXml = await handle.zip.file("[Content_Types].xml")?.async("string");
  if (!ctXml) return;

  const parsed: XNode[] = parser.parse(ctXml);
  const typesRoot = parsed.find((n: XNode) => n["Types"] !== undefined);
  if (!typesRoot) return;

  const children = typesRoot["Types"] as XNode[];

  // Check if Override for this partName already exists
  for (const child of children) {
    if (child["Override"] !== undefined && attr(child, "PartName") === partName) {
      return; // Already present
    }
  }

  // Add Override entry
  children.push(
    el("Override", [], { PartName: partName, ContentType: contentType }),
  );

  handle.zip.file("[Content_Types].xml", builder.build(parsed));
}

/** Add a Relationship entry to word/_rels/document.xml.rels if not already present. */
async function ensureDocRelationship(
  handle: DocxHandle,
  type: string,
  target: string,
): Promise<void> {
  const relsXml = await handle.zip
    .file("word/_rels/document.xml.rels")
    ?.async("string");
  if (!relsXml) return;

  const parsed: XNode[] = parser.parse(relsXml);
  const root = parsed.find((n: XNode) => n["Relationships"] !== undefined);
  if (!root) return;

  const children = root["Relationships"] as XNode[];

  // Check if relationship with this target already exists
  for (const child of children) {
    if (child["Relationship"] !== undefined && attr(child, "Target") === target) {
      return; // Already present
    }
  }

  // Find max rId
  let maxRId = 0;
  for (const child of children) {
    if (child["Relationship"] !== undefined) {
      const id = attr(child, "Id") ?? "";
      const m = id.match(/^rId(\d+)$/);
      if (m) {
        const num = parseInt(m[1]);
        if (num > maxRId) maxRId = num;
      }
    }
  }

  // Add Relationship entry
  children.push(
    el("Relationship", [], {
      Id: `rId${maxRId + 1}`,
      Type: type,
      Target: target,
    }),
  );

  handle.zip.file("word/_rels/document.xml.rels", builder.build(parsed));
}

export async function ensureCommentsInfrastructure(
  handle: DocxHandle,
): Promise<void> {
  // Ensure [Content_Types].xml has comments entry
  await ensureContentTypeOverride(
    handle,
    "/word/comments.xml",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
  );

  // Ensure word/_rels/document.xml.rels has comments relationship
  await ensureDocRelationship(
    handle,
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    "comments.xml",
  );
}

// ---------------------------------------------------------------------------
// commentsExtended.xml helpers (w15 namespace for reply threading)
// ---------------------------------------------------------------------------

export async function parseCommentsExtendedXml(handle: DocxHandle): Promise<XNode[]> {
  const xml = await handle.zip
    .file("word/commentsExtended.xml")
    ?.async("string");
  if (!xml) return [];
  return parser.parse(xml);
}

export function getCommentsExtendedArray(parsed: XNode[]): XNode[] {
  const root = parsed.find((n: XNode) => n["w15:commentsEx"]);
  if (!root) return [];
  return root["w15:commentsEx"];
}

export async function ensureCommentsExtendedInfrastructure(
  handle: DocxHandle,
): Promise<void> {
  // Ensure [Content_Types].xml has commentsExtended entry
  await ensureContentTypeOverride(
    handle,
    "/word/commentsExtended.xml",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml",
  );

  // Ensure word/_rels/document.xml.rels has commentsExtended relationship
  await ensureDocRelationship(
    handle,
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
    "commentsExtended.xml",
  );
}

/** Generate a random 8-char uppercase hex string for w14:paraId */
export function generateParaId(): string {
  return crypto.randomBytes(4).toString("hex").toUpperCase();
}

// ---------------------------------------------------------------------------
// Fuzzy anchor matching helpers
// ---------------------------------------------------------------------------

/** Normalize text for fuzzy matching: full-width→half-width, collapse whitespace, lowercase */
export function normalizeForMatching(text: string): string {
  // Full-width ASCII (U+FF01-FF5E) → half-width (U+0021-007E)
  let result = text.replace(/[\uFF01-\uFF5E]/g, (ch) =>
    String.fromCharCode(ch.charCodeAt(0) - 0xFEE0),
  );
  // Collapse all whitespace (including U+3000 ideographic space, NBSP) to single space
  result = result.replace(/[\s\u3000\u00A0]+/g, " ");
  // Lowercase
  return result.toLowerCase();
}

export interface AnchorMatch {
  pChildren: XNode[];
  matchedText: string;
  strategy: "exact" | "fuzzy";
}

/** Find a paragraph containing the anchor text, trying exact then fuzzy match. */
export function findAnchorParagraph(
  body: XNode[],
  anchorText: string,
): AnchorMatch | null {
  // Strategy 1: exact match
  for (const child of body) {
    if (child["w:p"]) {
      const text = extractParagraphText(child["w:p"]);
      if (text.includes(anchorText)) {
        return { pChildren: child["w:p"], matchedText: anchorText, strategy: "exact" };
      }
    } else if (child["w:tbl"]) {
      let result: AnchorMatch | null = null;
      forEachParagraphInTable(child["w:tbl"], (pChildren) => {
        if (!result) {
          const text = extractParagraphText(pChildren);
          if (text.includes(anchorText)) {
            result = { pChildren, matchedText: anchorText, strategy: "exact" };
          }
        }
      });
      if (result) return result;
    }
  }

  // Strategy 2: fuzzy match (normalized comparison)
  const normalizedAnchor = normalizeForMatching(anchorText);
  for (const child of body) {
    if (child["w:p"]) {
      const text = extractParagraphText(child["w:p"]);
      if (normalizeForMatching(text).includes(normalizedAnchor)) {
        return { pChildren: child["w:p"], matchedText: anchorText, strategy: "fuzzy" };
      }
    } else if (child["w:tbl"]) {
      let result: AnchorMatch | null = null;
      forEachParagraphInTable(child["w:tbl"], (pChildren) => {
        if (!result) {
          const text = extractParagraphText(pChildren);
          if (normalizeForMatching(text).includes(normalizedAnchor)) {
            result = { pChildren, matchedText: anchorText, strategy: "fuzzy" };
          }
        }
      });
      if (result) return result;
    }
  }

  return null;
}

/** Insert commentRangeStart/End and commentReference markers into paragraph children. */
export function insertCommentRangeMarkers(
  pChildren: XNode[],
  commentId: number,
  anchorText: string,
): void {
  const rangeStart = el("w:commentRangeStart", [], { "w:id": String(commentId) });
  const rangeEnd = el("w:commentRangeEnd", [], { "w:id": String(commentId) });
  const refRun = el("w:r", [
    el("w:rPr", [el("w:rStyle", [], { "w:val": "CommentReference" })]),
    el("w:commentReference", [], { "w:id": String(commentId) }),
  ]);

  // Try to find a run containing the anchor text and wrap it
  let inserted = false;
  for (let i = 0; i < pChildren.length; i++) {
    const child = pChildren[i];
    if (child["w:r"]) {
      let runText = "";
      for (const rc of child["w:r"]) {
        if (rc["w:t"]) {
          for (const tn of rc["w:t"]) {
            if (tn["#text"] !== undefined) runText += String(tn["#text"]);
          }
        }
      }
      // Try exact match on individual run, then fuzzy
      if (
        runText.includes(anchorText) ||
        normalizeForMatching(runText).includes(normalizeForMatching(anchorText))
      ) {
        pChildren.splice(i, 0, rangeStart);
        pChildren.splice(i + 2, 0, rangeEnd, refRun);
        inserted = true;
        break;
      }
    }
  }

  if (!inserted) {
    // Fallback: put markers at the beginning and end of paragraph content
    const pPr = findOne(pChildren, "w:pPr");
    const insertIdx = pPr ? 1 : 0;
    pChildren.splice(insertIdx, 0, rangeStart);
    pChildren.push(rangeEnd, refRun);
  }
}
