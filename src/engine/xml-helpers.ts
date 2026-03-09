/**
 * XML helper types and functions for preserveOrder mode of fast-xml-parser.
 */

import { XMLParser, XMLBuilder } from "fast-xml-parser";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export type XNode = any;

// ---------------------------------------------------------------------------
// Parser / Builder instances (shared, stateless)
// ---------------------------------------------------------------------------

const parserOpts = {
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  preserveOrder: true,
  trimValues: false,
  processEntities: false,
  // Never convert text content to numbers — "1." and ".0" must stay as strings
  parseTagValue: false,
  // commentPropName keeps XML comments (<!-- ... -->) instead of dropping them
  commentPropName: "#comment",
};

const builderOpts = {
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  preserveOrder: true,
  suppressEmptyNode: true,
  processEntities: false,
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
  return JSON.parse(JSON.stringify(node));
}
