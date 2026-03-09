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

export async function ensureCommentsInfrastructure(
  handle: DocxHandle,
): Promise<void> {
  // Ensure [Content_Types].xml has comments entry
  const ctXml = await handle.zip
    .file("[Content_Types].xml")
    ?.async("string");
  if (ctXml && !ctXml.includes("comments.xml")) {
    const insertion =
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>';
    const newCtXml = ctXml.replace("</Types>", insertion + "</Types>");
    handle.zip.file("[Content_Types].xml", newCtXml);
  }

  // Ensure word/_rels/document.xml.rels has comments relationship
  const relsXml = await handle.zip
    .file("word/_rels/document.xml.rels")
    ?.async("string");
  if (relsXml && !relsXml.includes("comments.xml")) {
    // Find max rId
    const rIdMatches = [...relsXml.matchAll(/rId(\d+)/g)];
    const maxRId = rIdMatches.reduce(
      (max, m) => Math.max(max, parseInt(m[1])),
      0,
    );
    const newRId = `rId${maxRId + 1}`;
    const insertion = `<Relationship Id="${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>`;
    const newRelsXml = relsXml.replace(
      "</Relationships>",
      insertion + "</Relationships>",
    );
    handle.zip.file("word/_rels/document.xml.rels", newRelsXml);
  }
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
  const ctXml = await handle.zip
    .file("[Content_Types].xml")
    ?.async("string");
  if (ctXml && !ctXml.includes("commentsExtended.xml")) {
    const insertion =
      '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>';
    const newCtXml = ctXml.replace("</Types>", insertion + "</Types>");
    handle.zip.file("[Content_Types].xml", newCtXml);
  }

  // Ensure word/_rels/document.xml.rels has commentsExtended relationship
  const relsXml = await handle.zip
    .file("word/_rels/document.xml.rels")
    ?.async("string");
  if (relsXml && !relsXml.includes("commentsExtended.xml")) {
    const rIdMatches = [...relsXml.matchAll(/rId(\d+)/g)];
    const maxRId = rIdMatches.reduce(
      (max, m) => Math.max(max, parseInt(m[1])),
      0,
    );
    const newRId = `rId${maxRId + 1}`;
    const insertion = `<Relationship Id="${newRId}" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>`;
    const newRelsXml = relsXml.replace(
      "</Relationships>",
      insertion + "</Relationships>",
    );
    handle.zip.file("word/_rels/document.xml.rels", newRelsXml);
  }
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
