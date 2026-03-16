/**
 * Character and paragraph formatting helpers.
 */

import {
  type XNode,
  findOne,
  el,
} from "./xml-helpers.js";
import {
  collectRunsWithIndices,
  makeTextRun,
  type RunWithIndex,
} from "./text.js";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface TextFormatting {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  highlightColor?: string;
  fontName?: string;
  fontSize?: number;
  fontColor?: string;
}

export interface ParagraphFormat {
  alignment?: string; // left, center, right, justify
  spaceBefore?: number; // in points
  spaceAfter?: number; // in points
  lineSpacing?: number; // in points (line height)
  indentLeft?: number; // in twips
  indentRight?: number; // in twips
  firstLineIndent?: number; // in twips
  hangingIndent?: number; // in twips
}

// ---------------------------------------------------------------------------
// Run-level formatting
// ---------------------------------------------------------------------------

export function ensureRunProperties(runChildren: XNode[]): XNode {
  let rPr = findOne(runChildren, "w:rPr");
  if (!rPr) {
    rPr = el("w:rPr");
    // Insert at beginning of run children (before w:t)
    runChildren.unshift(rPr);
  }
  return rPr;
}

export function setRunFormatting(
  runChildren: XNode[],
  fmt: TextFormatting,
): void {
  const rPr = ensureRunProperties(runChildren);
  const props = rPr["w:rPr"] as XNode[];

  const setFlag = (tag: string, value: boolean | undefined) => {
    if (value === undefined) return;
    const existing = props.findIndex((n: XNode) => n[tag] !== undefined);
    if (value) {
      if (existing === -1) props.push(el(tag));
    } else {
      if (existing !== -1) props.splice(existing, 1);
    }
  };

  setFlag("w:b", fmt.bold);
  setFlag("w:i", fmt.italic);
  setFlag("w:strike", fmt.strikethrough);

  // Underline
  if (fmt.underline !== undefined) {
    const idx = props.findIndex((n: XNode) => n["w:u"] !== undefined);
    if (fmt.underline) {
      const uEl = el("w:u", [], { "w:val": "single" });
      if (idx !== -1) props[idx] = uEl;
      else props.push(uEl);
    } else {
      if (idx !== -1) props.splice(idx, 1);
    }
  }

  // Highlight
  if (fmt.highlightColor) {
    const color = mapHighlightColor(fmt.highlightColor);
    const idx = props.findIndex((n: XNode) => n["w:highlight"] !== undefined);
    const hEl = el("w:highlight", [], { "w:val": color });
    if (idx !== -1) props[idx] = hEl;
    else props.push(hEl);
  }

  // Font name
  if (fmt.fontName) {
    const idx = props.findIndex((n: XNode) => n["w:rFonts"] !== undefined);
    const fEl = el("w:rFonts", [], {
      "w:ascii": fmt.fontName,
      "w:eastAsia": fmt.fontName,
      "w:hAnsi": fmt.fontName,
      "w:cs": fmt.fontName,
    });
    if (idx !== -1) props[idx] = fEl;
    else props.push(fEl);
  }

  // Font size (points → half-points)
  if (fmt.fontSize) {
    const halfPts = String(Math.round(fmt.fontSize * 2));
    const idxSz = props.findIndex((n: XNode) => n["w:sz"] !== undefined);
    const szEl = el("w:sz", [], { "w:val": halfPts });
    if (idxSz !== -1) props[idxSz] = szEl;
    else props.push(szEl);

    const idxSzCs = props.findIndex((n: XNode) => n["w:szCs"] !== undefined);
    const szCsEl = el("w:szCs", [], { "w:val": halfPts });
    if (idxSzCs !== -1) props[idxSzCs] = szCsEl;
    else props.push(szCsEl);
  }

  // Font color (hex)
  if (fmt.fontColor) {
    const hex = fmt.fontColor.replace("#", "").toUpperCase();
    const idx = props.findIndex((n: XNode) => n["w:color"] !== undefined);
    const cEl = el("w:color", [], { "w:val": hex });
    if (idx !== -1) props[idx] = cEl;
    else props.push(cEl);
  }
}

export function mapHighlightColor(color: string): string {
  const c = color.toLowerCase().replace("#", "");
  const map: Record<string, string> = {
    yellow: "yellow",
    ffff00: "yellow",
    green: "green",
    "00ff00": "green",
    cyan: "cyan",
    "00ffff": "cyan",
    magenta: "magenta",
    ff00ff: "magenta",
    blue: "blue",
    "0000ff": "blue",
    red: "red",
    ff0000: "red",
    darkblue: "darkBlue",
    "000080": "darkBlue",
    darkcyan: "darkCyan",
    "008080": "darkCyan",
    darkgreen: "darkGreen",
    "008000": "darkGreen",
    darkmagenta: "darkMagenta",
    "800080": "darkMagenta",
    darkred: "darkRed",
    "800000": "darkRed",
    darkyellow: "darkYellow",
    "808000": "darkYellow",
    darkgray: "darkGray",
    "808080": "darkGray",
    lightgray: "lightGray",
    c0c0c0: "lightGray",
    black: "black",
    "000000": "black",
  };
  return map[c] ?? "yellow";
}

/**
 * Apply formatting to all occurrences of `search` text in a paragraph.
 * Handles cross-run matches and formats only the matched portion of runs.
 * Returns count of occurrences formatted.
 */
export function formatInParagraph(
  pChildren: XNode[],
  search: string,
  fmt: TextFormatting,
  caseSensitive: boolean,
): number {
  const runs = collectRunsWithIndices(pChildren);
  if (runs.length === 0) return 0;

  const fullText = runs.map((r) => r.text).join("");
  const searchStr = caseSensitive ? search : search.toLowerCase();
  const compareText = caseSensitive ? fullText : fullText.toLowerCase();

  // Find all match positions
  const matches: number[] = [];
  let pos = 0;
  while (true) {
    const idx = compareText.indexOf(searchStr, pos);
    if (idx === -1) break;
    matches.push(idx);
    pos = idx + searchStr.length;
  }
  if (matches.length === 0) return 0;

  // Process matches in reverse order to preserve splice indices
  for (let mi = matches.length - 1; mi >= 0; mi--) {
    const matchStart = matches[mi];
    const matchEnd = matchStart + search.length;

    // Find which runs are affected
    let firstRunIdx = -1;
    let lastRunIdx = -1;
    for (let ri = 0; ri < runs.length; ri++) {
      const rStart = runs[ri].startOffset;
      const rEnd = rStart + runs[ri].text.length;
      if (rStart < matchEnd && rEnd > matchStart) {
        if (firstRunIdx === -1) firstRunIdx = ri;
        lastRunIdx = ri;
      }
    }
    if (firstRunIdx === -1) continue;

    if (firstRunIdx === lastRunIdx) {
      // Single-run match — split into [before] + [matched+formatted] + [after]
      const run = runs[firstRunIdx];
      const localStart = matchStart - run.startOffset;
      const beforeText = run.text.substring(0, localStart);
      const matchedText = run.text.substring(localStart, localStart + search.length);
      const afterText = run.text.substring(localStart + search.length);

      const newNodes: XNode[] = [];
      if (beforeText) newNodes.push(makeTextRun(beforeText, run.rPr));
      const fmtRun = makeTextRun(matchedText, run.rPr);
      setRunFormatting(fmtRun["w:r"], fmt);
      newNodes.push(fmtRun);
      if (afterText) newNodes.push(makeTextRun(afterText, run.rPr));

      pChildren.splice(run.pIdx, 1, ...newNodes);

      // Update indices for subsequent (earlier) matches
      const delta = newNodes.length - 1;
      for (let ri = firstRunIdx + 1; ri < runs.length; ri++) {
        runs[ri].pIdx += delta;
      }
    } else {
      // Cross-run match
      const newNodes: XNode[] = [];

      for (let ri = firstRunIdx; ri <= lastRunIdx; ri++) {
        const run = runs[ri];
        if (ri === firstRunIdx) {
          const localStart = matchStart - run.startOffset;
          const beforeText = run.text.substring(0, localStart);
          const matchedText = run.text.substring(localStart);
          if (beforeText) newNodes.push(makeTextRun(beforeText, run.rPr));
          const fmtRun = makeTextRun(matchedText, run.rPr);
          setRunFormatting(fmtRun["w:r"], fmt);
          newNodes.push(fmtRun);
        } else if (ri === lastRunIdx) {
          const localEnd = matchEnd - run.startOffset;
          const matchedText = run.text.substring(0, localEnd);
          const afterText = run.text.substring(localEnd);
          const fmtRun = makeTextRun(matchedText, run.rPr);
          setRunFormatting(fmtRun["w:r"], fmt);
          newNodes.push(fmtRun);
          if (afterText) newNodes.push(makeTextRun(afterText, run.rPr));
        } else {
          // Middle run — entirely formatted
          const fmtRun = makeTextRun(run.text, run.rPr);
          setRunFormatting(fmtRun["w:r"], fmt);
          newNodes.push(fmtRun);
        }
      }

      const pIdxStart = runs[firstRunIdx].pIdx;
      const pIdxEnd = runs[lastRunIdx].pIdx;
      pChildren.splice(pIdxStart, pIdxEnd - pIdxStart + 1, ...newNodes);

      // Update indices for subsequent (earlier) matches
      const removedCount = pIdxEnd - pIdxStart + 1;
      const delta = newNodes.length - removedCount;
      for (let ri = lastRunIdx + 1; ri < runs.length; ri++) {
        runs[ri].pIdx += delta;
      }
    }
  }

  return matches.length;
}

// ---------------------------------------------------------------------------
// Paragraph-level formatting
// ---------------------------------------------------------------------------

export function applyParagraphFormat(
  pChildren: XNode[],
  fmt: ParagraphFormat,
): void {
  let pPr = findOne(pChildren, "w:pPr");
  if (!pPr) {
    pPr = el("w:pPr");
    pChildren.unshift(pPr);
  }
  const props = pPr["w:pPr"] as XNode[];

  // Alignment
  if (fmt.alignment) {
    const valMap: Record<string, string> = {
      left: "left",
      center: "center",
      right: "right",
      justify: "both",
      both: "both",
    };
    const val = valMap[fmt.alignment] ?? "left";
    const idx = props.findIndex((n: XNode) => n["w:jc"] !== undefined);
    const jcEl = el("w:jc", [], { "w:val": val });
    if (idx !== -1) props[idx] = jcEl;
    else props.push(jcEl);
  }

  // Spacing — merge with existing attributes to preserve values not being overridden
  if (
    fmt.spaceBefore !== undefined ||
    fmt.spaceAfter !== undefined ||
    fmt.lineSpacing !== undefined
  ) {
    const idx = props.findIndex((n: XNode) => n["w:spacing"] !== undefined);
    const existing = idx !== -1 ? (props[idx][":@"] ?? {}) : {};
    const attrs: Record<string, string> = {};
    // Preserve existing attributes
    for (const [k, v] of Object.entries(existing)) {
      if (k.startsWith("@_")) attrs[k.slice(2)] = v as string;
    }
    // Override with new values
    if (fmt.spaceBefore !== undefined) {
      attrs["w:before"] = String(Math.round(fmt.spaceBefore * 20)); // pts to twips
    }
    if (fmt.spaceAfter !== undefined) {
      attrs["w:after"] = String(Math.round(fmt.spaceAfter * 20));
    }
    if (fmt.lineSpacing !== undefined) {
      attrs["w:line"] = String(Math.round(fmt.lineSpacing * 20));
      attrs["w:lineRule"] = "exact";
    }
    const spEl = el("w:spacing", [], attrs);
    if (idx !== -1) props[idx] = spEl;
    else props.push(spEl);
  }

  // Indentation — merge with existing attributes to preserve values not being overridden
  if (
    fmt.indentLeft !== undefined ||
    fmt.indentRight !== undefined ||
    fmt.firstLineIndent !== undefined ||
    fmt.hangingIndent !== undefined
  ) {
    const idx = props.findIndex((n: XNode) => n["w:ind"] !== undefined);
    const existing = idx !== -1 ? (props[idx][":@"] ?? {}) : {};
    const attrs: Record<string, string> = {};
    // Preserve existing attributes
    for (const [k, v] of Object.entries(existing)) {
      if (k.startsWith("@_")) attrs[k.slice(2)] = v as string;
    }
    // Override with new values
    if (fmt.indentLeft !== undefined) attrs["w:left"] = String(fmt.indentLeft);
    if (fmt.indentRight !== undefined) attrs["w:right"] = String(fmt.indentRight);
    if (fmt.firstLineIndent !== undefined)
      attrs["w:firstLine"] = String(fmt.firstLineIndent);
    if (fmt.hangingIndent !== undefined)
      attrs["w:hanging"] = String(fmt.hangingIndent);
    const indEl = el("w:ind", [], attrs);
    if (idx !== -1) props[idx] = indEl;
    else props.push(indEl);
  }
}
