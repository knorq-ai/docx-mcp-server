/**
 * Stable paragraph anchors (issue #7).
 *
 * An anchor is a paragraph's `w14:paraId` — a Word-native, index-independent id
 * that survives the index shifts caused by insert/delete. This module holds the
 * low-level machinery: id generation, part-wide id collection, root-namespace
 * seeding, and resolving an anchor (or a paragraph index) to a concrete
 * location in the document body.
 *
 * v1 anchors cover only DIRECT-body paragraphs — the same set
 * `blockBodyIndices()` enumerates and every mutation tool resolves against.
 * Paragraphs inside tables or top-level `w:sdt` are not anchored (their ids, if
 * any, still count toward part-wide uniqueness).
 */

import * as crypto from "crypto";
import { type XNode, attr, setAttr, findOne } from "./xml-helpers.js";
import { ErrorCode, EngineError } from "./docx-io.js";

export const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";
export const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/** A resolved paragraph location, enough for edit / delete / anchor-relative insert. */
export interface ParagraphLocation {
  /** The <w:p> wrapper node. */
  element: XNode;
  /** The array that directly holds `element` (the document body in v1). */
  parent: XNode[];
  /** `element`'s index within `parent`. */
  bodyIndex: number;
  /** 0-based blockBodyIndices position (paragraphs + tables), for messages. */
  blockIndex: number;
}

/**
 * Whether `id` is a canonical, MS-DOCX-valid `w14:paraId`: 8 hex chars whose
 * value is nonzero and < 0x80000000. (Case-insensitive on the hex digits; the
 * generator always emits uppercase.)
 */
export function isValidParaId(id: string | undefined): boolean {
  if (!id || !/^[0-9A-Fa-f]{8}$/.test(id)) return false;
  const n = parseInt(id, 16);
  return n > 0 && n < 0x80000000;
}

/**
 * Generate a valid, unique `w14:paraId`: a 32-bit value in 0x00000001–0x7FFFFFFF
 * (MS-DOCX requires nonzero and < 0x80000000, unique within the part), formatted
 * as 8 uppercase hex chars. `used` is mutated to include the returned id.
 */
export function generateDocParaId(used: Set<string>): string {
  for (let attempt = 0; attempt < 10000; attempt++) {
    const n = crypto.randomBytes(4).readUInt32BE(0) & 0x7fffffff;
    if (n === 0) continue;
    const id = n.toString(16).toUpperCase().padStart(8, "0");
    if (!used.has(id)) {
      used.add(id);
      return id;
    }
  }
  throw new EngineError(ErrorCode.INVALID_DOCX, "Unable to allocate a unique paragraph id.");
}

/**
 * Collect the `w14:paraId` of every `<w:p>` in the given nodes, recursing into
 * tables and `w:sdt`. Returns the full set (for part-wide uniqueness) and the
 * subset that appears more than once (duplicate ids).
 */
export function collectAllParaIds(nodes: XNode[]): { all: Set<string>; duplicates: Set<string> } {
  const all = new Set<string>();
  const duplicates = new Set<string>();
  const visit = (arr: XNode[]): void => {
    for (const node of arr) {
      if (node["w:p"]) {
        const id = attr(node, "w14:paraId");
        if (id) {
          if (all.has(id)) duplicates.add(id);
          else all.add(id);
        }
      } else if (node["w:tbl"]) {
        visit(node["w:tbl"] as XNode[]);
      } else if (node["w:tr"]) {
        visit(node["w:tr"] as XNode[]);
      } else if (node["w:tc"]) {
        visit(node["w:tc"] as XNode[]);
      } else if (node["w:sdt"]) {
        const content = findOne(node["w:sdt"] as XNode[], "w:sdtContent");
        if (content) visit(content["w:sdtContent"] as XNode[]);
      }
    }
  };
  visit(nodes);
  return { all, duplicates };
}

/** Find the `<w:document>` root node within the parsed top level. */
export function getDocumentRoot(parsed: XNode[]): XNode {
  const root = parsed.find((n: XNode) => n["w:document"]);
  if (!root) {
    throw new EngineError(ErrorCode.INVALID_DOCX, "document.xml has no <w:document> root.");
  }
  return root;
}

/**
 * Ensure the `<w:document>` root declares the `w14` and `mc` namespaces and lists
 * `w14` in `mc:Ignorable`. Idempotent. Refuses (`INVALID_DOCX`) if either prefix
 * is already bound to a different URI, rather than emit a wrong-namespace
 * attribute. Returns true if anything was added/changed.
 */
export function ensureW14Namespace(root: XNode): boolean {
  let changed = false;

  const ensureBinding = (prefixAttr: string, expected: string): void => {
    const current = attr(root, prefixAttr);
    if (current === undefined) {
      setAttr(root, prefixAttr, expected);
      changed = true;
    } else if (current !== expected) {
      throw new EngineError(
        ErrorCode.INVALID_DOCX,
        `${prefixAttr} is bound to an unexpected namespace URI; cannot safely add w14:paraId.`,
      );
    }
  };

  ensureBinding("xmlns:w14", W14_NS);
  ensureBinding("xmlns:mc", MC_NS);

  // mc:Ignorable as a whitespace-delimited token set.
  const ignorable = attr(root, "mc:Ignorable");
  const tokens = (ignorable ?? "").split(/\s+/).filter(Boolean);
  if (!tokens.includes("w14")) {
    tokens.push("w14");
    setAttr(root, "mc:Ignorable", tokens.join(" "));
    changed = true;
  }

  return changed;
}

/**
 * List the locations of every DIRECT-body paragraph. `blockIndex` matches the
 * blockBodyIndices numbering (paragraphs and tables both advance it; tables are
 * not anchorable). Paragraphs inside tables / `w:sdt` are intentionally skipped.
 */
export function directBodyParagraphLocations(body: XNode[]): ParagraphLocation[] {
  const locations: ParagraphLocation[] = [];
  let blockIndex = 0;
  for (let i = 0; i < body.length; i++) {
    const node = body[i];
    if (node["w:p"]) {
      locations.push({ element: node, parent: body, bodyIndex: i, blockIndex });
      blockIndex++;
    } else if (node["w:tbl"] || node["w:sdt"]) {
      // Tables and content controls each advance the block index by one but are
      // not anchorable; this mirrors blockBodyIndices()/enumerateBlocks().
      blockIndex++;
    }
  }
  return locations;
}

/** Map paraId → direct-body locations (an id may appear more than once in a malformed doc). */
export function buildAnchorIndex(body: XNode[]): Map<string, ParagraphLocation[]> {
  const index = new Map<string, ParagraphLocation[]>();
  for (const loc of directBodyParagraphLocations(body)) {
    const id = attr(loc.element, "w14:paraId");
    if (!id) continue;
    const list = index.get(id);
    if (list) list.push(loc);
    else index.set(id, [loc]);
  }
  return index;
}

/**
 * Resolve a single anchor to a direct-body paragraph location.
 * Throws ANCHOR_NOT_FOUND (no match) or AMBIGUOUS_ANCHOR (>1 match).
 */
export function resolveAnchor(
  anchorIndex: Map<string, ParagraphLocation[]>,
  anchor: string,
): ParagraphLocation {
  const matches = anchorIndex.get(anchor);
  if (!matches || matches.length === 0) {
    throw new EngineError(ErrorCode.ANCHOR_NOT_FOUND, `No paragraph found for anchor "${anchor}".`);
  }
  if (matches.length > 1) {
    throw new EngineError(
      ErrorCode.AMBIGUOUS_ANCHOR,
      `Anchor "${anchor}" matches ${matches.length} paragraphs; run ensure_anchors to repair duplicate ids.`,
    );
  }
  return matches[0];
}

/**
 * Read this paragraph's existing anchor, or assign and return a fresh one. The
 * caller must have ensured the w14 namespace on the document root (see
 * ensureW14Namespace) and must pass the part-wide `used` set so new ids stay
 * unique. Used for auto-seeding paragraphs a write touches or creates.
 */
export function seedParagraphAnchor(element: XNode, used: Set<string>): string {
  const existing = attr(element, "w14:paraId");
  if (existing) return existing;
  const id = generateDocParaId(used);
  setAttr(element, "w14:paraId", id);
  return id;
}
