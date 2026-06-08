/**
 * XML helper types and functions for preserveOrder mode of fast-xml-parser.
 */

import { XMLParser, XMLBuilder } from "fast-xml-parser";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export type XNode = any;

// ---------------------------------------------------------------------------
// Parser / Builder instances (shared, stateless)
// ---------------------------------------------------------------------------

/**
 * A Unicode scalar value that `String.fromCodePoint` accepts AND that is a legal
 * XML character: in 0x0–0x10FFFF and not a UTF-16 surrogate (0xD800–0xDFFF).
 * A reference outside this range (e.g. `&#x110000;` or a 14-digit decimal) would
 * make `String.fromCodePoint` throw `RangeError: Invalid code point`; we instead
 * leave such a (malformed) reference as literal text so the document still parses.
 */
function isDecodableCodePoint(cp: number): boolean {
  return Number.isInteger(cp) && cp >= 0 && cp <= 0x10ffff && !(cp >= 0xd800 && cp <= 0xdfff);
}

/**
 * Decode numeric character references (&#NNN; and &#xHHH;) that processEntities
 * doesn't handle. Lenient: a reference whose code point is out of the valid XML
 * range is left UNCHANGED (returned as the original matched text) rather than
 * crashing the read with a RangeError — document.xml is untrusted input.
 */
function decodeNumericRefs(_name: string, val: unknown): unknown {
  if (typeof val !== "string") return val;
  return val
    .replace(/&#x([0-9a-fA-F]+);/g, (match, hex: string) => {
      const cp = parseInt(hex, 16);
      return isDecodableCodePoint(cp) ? String.fromCodePoint(cp) : match;
    })
    .replace(/&#(\d+);/g, (match, dec: string) => {
      const cp = parseInt(dec, 10);
      return isDecodableCodePoint(cp) ? String.fromCodePoint(cp) : match;
    });
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

/**
 * XML-1.0 forbids these characters in character data: the C0 controls
 * U+0000–U+0008, U+000B (vertical tab), U+000C (form feed), U+000E–U+001F (incl.
 * U+001B ESC), and the noncharacters U+FFFE / U+FFFF. Only \t (U+0009), \n
 * (U+000A) and \r (U+000D) are legal whitespace and are preserved. Such chars
 * routinely arrive in text pasted from PDFs/terminals; writing them verbatim
 * produces a document.xml that Word and python-docx reject. We strip them
 * (Word's own behavior) so no run-text path can corrupt the file.
 */
// eslint-disable-next-line no-control-regex
const ILLEGAL_XML_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g;

/**
 * Drop UNPAIRED UTF-16 surrogates (also XML-1.0-illegal): a high surrogate
 * U+D800–DBFF not immediately followed by a low U+DC00–DFFF, or a low surrogate
 * not immediately preceded by a high one. VALID surrogate pairs (astral chars
 * such as emoji) are kept intact — the pair is consumed together and re-emitted.
 */
function stripLoneSurrogates(text: string): string {
  let out = "";
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    if (code >= 0xd800 && code <= 0xdbff) {
      // High surrogate: keep only if the next unit is a low surrogate (valid pair).
      const next = text.charCodeAt(i + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        out += text[i] + text[i + 1];
        i++; // consume the low surrogate too
      }
      // else: lone high surrogate → drop.
    } else if (code >= 0xdc00 && code <= 0xdfff) {
      // Lone low surrogate (a valid pair would have been consumed above) → drop.
    } else {
      out += text[i];
    }
  }
  return out;
}

/**
 * Remove XML-1.0-illegal characters from run text: C0 controls (keeps \t \n \r),
 * the noncharacters U+FFFE/U+FFFF, and unpaired surrogates. Valid surrogate
 * pairs (emoji / astral chars) survive.
 */
export function sanitizeXmlText(text: string): string {
  return stripLoneSurrogates(text.replace(ILLEGAL_XML_CHARS, ""));
}

/**
 * Normalize line endings to a lone "\n": collapse CRLF ("\r\n") and a bare
 * carriage return ("\r") to "\n". Applied at the start of every edit/insert
 * text pipeline BEFORE splitting on "\n" or building soft-break runs. Without
 * it, a stray "\r" survives inside a <w:t>; a conformant reader (Word /
 * python-docx, per XML 1.0 §2.11 line-end normalization) then turns it into a
 * literal newline character in the run — rendered as whitespace, not the
 * intended line/paragraph break. "\r" is XML-legal (sanitizeXmlText keeps it),
 * so this normalization is specific to the edit-text pipeline and does not
 * touch verbatim document content elsewhere.
 */
export function normalizeNewlines(text: string): string {
  return text.replace(/\r\n?/g, "\n");
}

/** Create a text node (illegal XML control chars are stripped here). */
export function textNode(text: string): XNode {
  return { "#text": sanitizeXmlText(text) };
}

export function cloneNode(node: XNode): XNode {
  return structuredClone(node);
}
