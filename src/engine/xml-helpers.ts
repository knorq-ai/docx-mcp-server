/**
 * XML helper types and functions for preserveOrder mode of fast-xml-parser.
 */

import { XMLParser, XMLBuilder } from "fast-xml-parser";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export type XNode = any;

// ---------------------------------------------------------------------------
// Parser / Builder instances (shared, stateless)
// ---------------------------------------------------------------------------

/** Whether `cp` is a code point allowed in XML 1.0 character content. */
function isValidXmlCodePoint(cp: number): boolean {
  return (
    cp === 0x9 ||
    cp === 0xa ||
    cp === 0xd ||
    (cp >= 0x20 && cp <= 0xd7ff) ||
    (cp >= 0xe000 && cp <= 0xfffd) ||
    (cp >= 0x10000 && cp <= 0x10ffff)
  );
}

/**
 * Decode numeric character references (&#NNN; and &#xHHH;) that processEntities
 * doesn't handle. References to code points that are not valid XML characters
 * (e.g. NUL, control chars, out-of-range) are left as the original entity text
 * rather than decoded, so they cannot inject raw control characters.
 */
function decodeNumericRefs(_name: string, val: unknown): unknown {
  if (typeof val !== "string") return val;
  const decode = (whole: string, cp: number): string =>
    Number.isFinite(cp) && isValidXmlCodePoint(cp)
      ? String.fromCodePoint(cp)
      : whole;
  return val
    .replace(/&#x([0-9a-fA-F]+);/g, (m, hex: string) => decode(m, parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (m, dec: string) => decode(m, parseInt(dec, 10)));
}

const parserOpts = {
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  preserveOrder: true,
  trimValues: false,
  // Decode standard XML entities (&amp; → &, &lt; → <, etc.) so text nodes
  // contain human-readable text.  The builder re-encodes them on output.
  processEntities: true,
  // Never convert text content to numbers — "1." and ".0" must stay as strings
  parseTagValue: false,
  // commentPropName keeps XML comments (<!-- ... -->) instead of dropping them
  commentPropName: "#comment",
  // Decode numeric character references (&#160; → NBSP, &#x20AC; → €, etc.)
  // that processEntities alone does not handle.
  tagValueProcessor: decodeNumericRefs,
};

const builderOpts = {
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  preserveOrder: true,
  suppressEmptyNode: true,
  // Re-encode &, <, > in text nodes so the output is always valid XML.
  processEntities: true,
  commentPropName: "#comment",
};

export const parser = new XMLParser(parserOpts);
export const builder = new XMLBuilder(builderOpts);

// ---------------------------------------------------------------------------
// DOM-like helpers for preserveOrder nodes
// ---------------------------------------------------------------------------

export function tagName(el: XNode): string | null {
  for (const k of Object.keys(el)) {
    if (k !== ":@" && k !== "#text" && k !== "#comment") return k;
  }
  return null;
}

export function children(el: XNode): XNode[] {
  const t = tagName(el);
  return t ? el[t] ?? [] : [];
}

export function attr(el: XNode, name: string): string | undefined {
  return el[":@"]?.["@_" + name];
}

export function setAttr(el: XNode, name: string, value: string): void {
  if (!el[":@"]) el[":@"] = {};
  el[":@"]["@_" + name] = value;
}

export function findAll(nodes: XNode[], tag: string): XNode[] {
  return nodes.filter((n) => n[tag] !== undefined);
}

export function findOne(nodes: XNode[], tag: string): XNode | undefined {
  return nodes.find((n) => n[tag] !== undefined);
}

/** Create a new element node */
export function el(
  tag: string,
  childArr: XNode[] = [],
  attrs?: Record<string, string>,
): XNode {
  const node: XNode = { [tag]: childArr };
  if (attrs) {
    node[":@"] = {};
    for (const [k, v] of Object.entries(attrs)) {
      node[":@"]["@_" + k] = v;
    }
  }
  return node;
}

/** Create a text node */
export function textNode(text: string): XNode {
  return { "#text": text };
}

export function cloneNode(node: XNode): XNode {
  return structuredClone(node);
}
