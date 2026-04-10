/**
 * Text extraction, block enumeration, cross-run replacement, and track changes helpers.
 */

import {
  type XNode,
  tagName,
  attr,
  findAll,
  findOne,
  el,
  textNode,
  cloneNode,
} from "./xml-helpers.js";

// ---------------------------------------------------------------------------
// Text extraction
// ---------------------------------------------------------------------------

/** Extract text from a single run's children (w:t, w:tab, w:br). */
export function extractRunText(runChildren: XNode[]): string {
  let text = "";
  for (const rc of runChildren) {
    if (rc["w:t"]) {
      for (const tn of rc["w:t"]) {
        if (tn["#text"] !== undefined) text += String(tn["#text"]);
      }
    } else if (rc["w:tab"]) {
      text += "\t";
    } else if (rc["w:br"]) {
      text += "\n";
    }
  }
  return text;
}

/** Extract w:delText from a run (deleted text in track changes). */
function extractDelRunText(runChildren: XNode[]): string {
  let text = "";
  for (const rc of runChildren) {
    if (rc["w:delText"]) {
      for (const tn of rc["w:delText"]) {
        if (tn["#text"] !== undefined) text += String(tn["#text"]);
      }
    }
  }
  return text;
}

/**
 * Extract paragraph text.
 * - Default (showRevisions=false): "accepted" view — includes original + inserted text, skips deleted.
 * - showRevisions=true: annotates changes like [-deleted-][+inserted+].
 */
export function extractParagraphText(
  pChildren: XNode[],
  showRevisions: boolean = false,
): string {
  let text = "";
  for (const child of pChildren) {
    if (child["w:r"]) {
      text += extractRunText(child["w:r"]);
    } else if (child["w:hyperlink"]) {
      for (const hlChild of child["w:hyperlink"]) {
        if (hlChild["w:r"]) {
          text += extractRunText(hlChild["w:r"]);
        }
      }
    } else if (child["w:ins"]) {
      // Tracked insertions
      let insText = "";
      for (const insChild of child["w:ins"]) {
        if (insChild["w:r"]) {
          insText += extractRunText(insChild["w:r"]);
        }
      }
      text += showRevisions && insText ? `[+${insText}+]` : insText;
    } else if (child["w:del"]) {
      // Tracked deletions — only show if revisions requested
      if (showRevisions) {
        let delText = "";
        for (const delChild of child["w:del"]) {
          if (delChild["w:r"]) {
            delText += extractDelRunText(delChild["w:r"]);
          }
        }
        if (delText) text += `[-${delText}-]`;
      }
      // In default mode: skip deleted text entirely (accepted view)
    } else if (child["w:sdt"]) {
      // Structured document tags (common in Google Docs exports) — recurse into sdtContent
      const sdtContent = findOne(child["w:sdt"] as XNode[], "w:sdtContent");
      if (sdtContent) {
        text += extractParagraphText(sdtContent["w:sdtContent"] as XNode[], showRevisions);
      }
    }
  }
  return text;
}

export function extractCellText(cellChildren: XNode[], showRevisions: boolean): string {
  const parts: string[] = [];
  for (const child of cellChildren) {
    if (child["w:p"]) {
      parts.push(extractParagraphText(child["w:p"], showRevisions));
    } else if (child["w:tbl"]) {
      // Nested table — include its text too
      parts.push(extractTableText(child["w:tbl"], showRevisions));
    }
  }
  return parts.join("\\n");
}

export function extractTableText(
  tblChildren: XNode[],
  showRevisions: boolean = false,
): string {
  const rows = findAll(tblChildren, "w:tr");
  if (rows.length === 0) return "[TABLE (empty)]";

  let out = "[TABLE]\n";
  for (const row of rows) {
    const cells = findAll(row["w:tr"], "w:tc");
    const cellTexts = cells.map((cell) =>
      extractCellText(cell["w:tc"], showRevisions)
    );
    out += "| " + cellTexts.join(" | ") + " |\n";
  }
  out += "[/TABLE]";
  return out;
}

// ---------------------------------------------------------------------------
// Style / heading detection
// ---------------------------------------------------------------------------

export function getParagraphStyle(pChildren: XNode[]): string | undefined {
  const pPr = findOne(pChildren, "w:pPr");
  if (!pPr) return undefined;
  const pStyle = findOne(pPr["w:pPr"], "w:pStyle");
  if (!pStyle) return undefined;
  return attr(pStyle, "w:val");
}

export function getHeadingLevel(style: string | undefined): number | undefined {
  if (!style) return undefined;
  const m = style.match(/^[Hh]eading(\d)$/);
  return m ? parseInt(m[1]) : undefined;
}

function getParagraphAlignment(pChildren: XNode[]): string | undefined {
  const pPr = findOne(pChildren, "w:pPr");
  if (!pPr) return undefined;
  const jc = findOne(pPr["w:pPr"], "w:jc");
  if (!jc) return undefined;
  return attr(jc, "w:val");
}

// ---------------------------------------------------------------------------
// Paragraph enumeration (including tables as blocks)
// ---------------------------------------------------------------------------

export interface BlockInfo {
  index: number;
  type: "paragraph" | "table";
  text: string;
  style?: string;
  alignment?: string;
  headingLevel?: number;
}

export function enumerateBlocks(
  body: XNode[],
  showRevisions: boolean = false,
): BlockInfo[] {
  const blocks: BlockInfo[] = [];
  let idx = 0;
  for (const child of body) {
    if (child["w:p"]) {
      const pChildren = child["w:p"];
      const text = extractParagraphText(pChildren, showRevisions);
      const style = getParagraphStyle(pChildren);
      const alignment = getParagraphAlignment(pChildren);
      const hl = getHeadingLevel(style);
      blocks.push({
        index: idx,
        type: "paragraph",
        text,
        style,
        alignment,
        headingLevel: hl,
      });
      idx++;
    } else if (child["w:tbl"]) {
      blocks.push({
        index: idx,
        type: "table",
        text: extractTableText(child["w:tbl"], showRevisions),
        style: "Table",
      });
      idx++;
    } else if (child["w:sdt"]) {
      // Content controls — extract paragraphs from w:sdtContent
      const sdtChildren = child["w:sdt"] as XNode[];
      const sdtContent = findOne(sdtChildren, "w:sdtContent");
      if (sdtContent) {
        const contentChildren = sdtContent["w:sdtContent"] as XNode[];
        for (const contentChild of contentChildren) {
          if (contentChild["w:p"]) {
            const pChildren = contentChild["w:p"];
            const text = extractParagraphText(pChildren, showRevisions);
            const style = getParagraphStyle(pChildren);
            const alignment = getParagraphAlignment(pChildren);
            const hl = getHeadingLevel(style);
            blocks.push({
              index: idx,
              type: "paragraph",
              text,
              style,
              alignment,
              headingLevel: hl,
            });
            idx++;
          }
        }
      }
    }
    // Skip sectPr and other non-content elements for block counting
  }
  return blocks;
}

// ---------------------------------------------------------------------------
// Cross-run text replacement
// ---------------------------------------------------------------------------

interface RunInfo {
  node: XNode; // The w:r element in the parent array
  textNodes: XNode[]; // The w:t > #text nodes
  text: string;
  startOffset: number; // Character offset within the paragraph
}

function collectRuns(pChildren: XNode[]): RunInfo[] {
  const runs: RunInfo[] = [];
  let offset = 0;

  function collectFromChildren(children: XNode[]): void {
    for (const child of children) {
      if (child["w:r"]) {
        const runC = child["w:r"];
        const textNodes: XNode[] = [];
        let runText = "";
        for (const rc of runC) {
          if (rc["w:t"]) {
            for (const tn of rc["w:t"]) {
              if (tn["#text"] !== undefined) {
                textNodes.push(tn);
                runText += String(tn["#text"]);
              }
            }
          }
        }
        if (textNodes.length > 0) {
          runs.push({ node: child, textNodes, text: runText, startOffset: offset });
        }
        offset += runText.length;
      } else if (child["w:ins"]) {
        // Include runs inside w:ins tracked insertions
        const insChildren = child["w:ins"] as XNode[];
        for (const insChild of insChildren) {
          if (insChild["w:r"]) {
            const runC = insChild["w:r"] as XNode[];
            const textNodes: XNode[] = [];
            let runText = "";
            for (const rc of runC) {
              if (rc["w:t"]) {
                for (const tn of rc["w:t"]) {
                  if (tn["#text"] !== undefined) {
                    textNodes.push(tn);
                    runText += String(tn["#text"]);
                  }
                }
              }
            }
            if (textNodes.length > 0) {
              runs.push({ node: insChild, textNodes, text: runText, startOffset: offset });
            }
            offset += runText.length;
          }
        }
      } else if (child["w:sdt"]) {
        // Recurse into structured document tag content
        const sdtContent = findOne(child["w:sdt"] as XNode[], "w:sdtContent");
        if (sdtContent) {
          collectFromChildren(sdtContent["w:sdtContent"] as XNode[]);
        }
      }
    }
  }

  collectFromChildren(pChildren);
  return runs;
}

/**
 * Replace `search` with `replace` inside the runs of a single paragraph.
 * Handles cross-run matches. Returns number of replacements made.
 */
export function replaceInParagraph(
  pChildren: XNode[],
  search: string,
  replace: string,
  caseSensitive: boolean,
): number {
  const runs = collectRuns(pChildren);
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

  // Process matches in reverse order to maintain offsets
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
      // Single run — simple replacement
      const run = runs[firstRunIdx];
      const localStart = matchStart - run.startOffset;
      const newText =
        run.text.substring(0, localStart) +
        replace +
        run.text.substring(localStart + search.length);
      // Update the text node(s) — put all text in first text node
      run.textNodes[0]["#text"] = newText;
      for (let ti = 1; ti < run.textNodes.length; ti++) {
        run.textNodes[ti]["#text"] = "";
      }
      run.text = newText;
    } else {
      // Cross-run replacement
      // Put replacement in first run, clear matched portion from others
      const firstRun = runs[firstRunIdx];
      const localStart = matchStart - firstRun.startOffset;
      const firstRunNewText =
        firstRun.text.substring(0, localStart) + replace;
      firstRun.textNodes[0]["#text"] = firstRunNewText;
      for (let ti = 1; ti < firstRun.textNodes.length; ti++) {
        firstRun.textNodes[ti]["#text"] = "";
      }
      firstRun.text = firstRunNewText;

      for (let ri = firstRunIdx + 1; ri <= lastRunIdx; ri++) {
        const run = runs[ri];
        if (ri === lastRunIdx) {
          // Last run — keep text after the match
          const localEnd = matchEnd - run.startOffset;
          const kept = run.text.substring(localEnd);
          run.textNodes[0]["#text"] = kept;
          for (let ti = 1; ti < run.textNodes.length; ti++) {
            run.textNodes[ti]["#text"] = "";
          }
          run.text = kept;
        } else {
          // Middle runs — clear completely
          for (const tn of run.textNodes) {
            tn["#text"] = "";
          }
          run.text = "";
        }
      }
    }
  }

  return matches.length;
}

// ---------------------------------------------------------------------------
// Track changes helpers
// ---------------------------------------------------------------------------

export interface RevisionContext {
  nextId: number;
  author: string;
  date: string;
}

export function newRevisionContext(startId: number, author: string): RevisionContext {
  return { nextId: startId, author, date: new Date().toISOString() };
}

export function allocRevId(ctx: RevisionContext): number {
  return ctx.nextId++;
}

/** Recursively scan parsed XML tree for the maximum w:id attribute value */
export function scanMaxId(nodes: XNode[]): number {
  let max = 0;
  for (const node of nodes) {
    const id = attr(node, "w:id");
    if (id) {
      const v = parseInt(id);
      if (!isNaN(v) && v > max) max = v;
    }
    const t = tagName(node);
    if (t && Array.isArray(node[t])) {
      const childMax = scanMaxId(node[t]);
      if (childMax > max) max = childMax;
    }
  }
  return max;
}

export function getRunRPr(runChildren: XNode[]): XNode | null {
  const rPr = findOne(runChildren, "w:rPr");
  return rPr ? cloneNode(rPr) : null;
}

/** Create a normal w:r with w:t */
export function makeTextRun(text: string, rPr: XNode | null): XNode {
  const runC: XNode[] = [];
  if (rPr) runC.push(cloneNode(rPr));
  runC.push(el("w:t", [textNode(text)], { "xml:space": "preserve" }));
  return el("w:r", runC);
}

/** Create a w:r with w:delText (for use inside w:del) */
export function makeDelTextRun(text: string, rPr: XNode | null): XNode {
  const runC: XNode[] = [];
  if (rPr) runC.push(cloneNode(rPr));
  runC.push(
    el("w:delText", [textNode(text)], { "xml:space": "preserve" }),
  );
  return el("w:r", runC);
}

/** Wrap runs in a w:del revision element */
export function wrapInDel(runs: XNode[], ctx: RevisionContext): XNode {
  return el("w:del", runs, {
    "w:id": String(allocRevId(ctx)),
    "w:author": ctx.author,
    "w:date": ctx.date,
  });
}

/** Wrap runs in a w:ins revision element */
export function wrapInIns(runs: XNode[], ctx: RevisionContext): XNode {
  return el("w:ins", runs, {
    "w:id": String(allocRevId(ctx)),
    "w:author": ctx.author,
    "w:date": ctx.date,
  });
}

/**
 * Find the common prefix and suffix between two strings, returning the minimal
 * changed region. Used to produce word-level tracked-change diffs rather than
 * whole-paragraph deletions.
 */
export function computeMinimalDiff(
  oldText: string,
  newText: string,
): { prefix: string; oldMiddle: string; newMiddle: string; suffix: string } {
  const oldLen = oldText.length;
  const newLen = newText.length;
  const minLen = Math.min(oldLen, newLen);

  let prefixLen = 0;
  while (prefixLen < minLen && oldText[prefixLen] === newText[prefixLen]) {
    prefixLen++;
  }

  let suffixLen = 0;
  const maxSuffix = minLen - prefixLen;
  while (
    suffixLen < maxSuffix &&
    oldText[oldLen - 1 - suffixLen] === newText[newLen - 1 - suffixLen]
  ) {
    suffixLen++;
  }

  return {
    prefix: oldText.slice(0, prefixLen),
    oldMiddle: oldText.slice(prefixLen, suffixLen > 0 ? oldLen - suffixLen : oldLen),
    newMiddle: newText.slice(prefixLen, suffixLen > 0 ? newLen - suffixLen : newLen),
    suffix: suffixLen > 0 ? oldText.slice(oldLen - suffixLen) : "",
  };
}

// ---------------------------------------------------------------------------
// Tracked replacement (cross-run aware)
// ---------------------------------------------------------------------------

export interface RunWithIndex {
  pIdx: number; // index in pChildren
  runChildren: XNode[];
  rPr: XNode | null;
  text: string;
  startOffset: number;
}

export function collectRunsWithIndices(pChildren: XNode[]): RunWithIndex[] {
  const runs: RunWithIndex[] = [];
  let offset = 0;
  // NOTE: w:sdt is NOT traversed here because splice operations use pIdx
  // directly on the children array. SDT content is handled separately by
  // replaceInParagraphTracked which recurses into each SDT's sdtContent.
  for (let i = 0; i < pChildren.length; i++) {
    const child = pChildren[i];
    if (child["w:r"]) {
      const runC = child["w:r"] as XNode[];
      let runText = "";
      for (const rc of runC) {
        if (rc["w:t"]) {
          for (const tn of rc["w:t"]) {
            if (tn["#text"] !== undefined) runText += String(tn["#text"]);
          }
        }
      }
      const rPr = getRunRPr(runC);
      runs.push({
        pIdx: i,
        runChildren: runC,
        rPr,
        text: runText,
        startOffset: offset,
      });
      offset += runText.length;
    } else if (child["w:ins"]) {
      // Include runs inside w:ins tracked insertions
      const insChildren = child["w:ins"] as XNode[];
      for (const insChild of insChildren) {
        if (insChild["w:r"]) {
          const runC = insChild["w:r"] as XNode[];
          let runText = "";
          for (const rc of runC) {
            if (rc["w:t"]) {
              for (const tn of rc["w:t"]) {
                if (tn["#text"] !== undefined) runText += String(tn["#text"]);
              }
            }
          }
          const rPr = getRunRPr(runC);
          // Use pIdx = i (the w:ins element index) so splice operations target the right element
          runs.push({
            pIdx: i,
            runChildren: runC,
            rPr,
            text: runText,
            startOffset: offset,
          });
          offset += runText.length;
        }
      }
    }
  }
  return runs;
}

/**
 * Replace `search` with `replace` using tracked changes markup.
 * Old text is wrapped in <w:del> with <w:delText>, new text in <w:ins>.
 * Handles cross-run matches. Returns number of replacements made.
 */
export function replaceInParagraphTracked(
  pChildren: XNode[],
  search: string,
  replace: string,
  caseSensitive: boolean,
  ctx: RevisionContext,
): number {
  let count = replaceInChildrenTracked(pChildren, search, replace, caseSensitive, ctx);

  // Recurse into w:sdt elements — each SDT's sdtContent is an independent splice scope
  for (const child of pChildren) {
    if (child["w:sdt"]) {
      const sdtContent = findOne(child["w:sdt"] as XNode[], "w:sdtContent");
      if (sdtContent) {
        count += replaceInChildrenTracked(
          sdtContent["w:sdtContent"] as XNode[], search, replace, caseSensitive, ctx,
        );
      }
    }
  }

  return count;
}

/** Core tracked replacement on a flat children array (no SDT traversal). */
function replaceInChildrenTracked(
  children: XNode[],
  search: string,
  replace: string,
  caseSensitive: boolean,
  ctx: RevisionContext,
): number {
  const runs = collectRunsWithIndices(children);
  if (runs.length === 0) return 0;

  const fullText = runs.map((r) => r.text).join("");
  const searchStr = caseSensitive ? search : search.toLowerCase();
  const compareText = caseSensitive ? fullText : fullText.toLowerCase();

  const matches: number[] = [];
  let pos = 0;
  while (true) {
    const idx = compareText.indexOf(searchStr, pos);
    if (idx === -1) break;
    matches.push(idx);
    pos = idx + searchStr.length;
  }
  if (matches.length === 0) return 0;

  // Pre-compute minimal diff so only the changed characters appear in del/ins
  const diff = computeMinimalDiff(search, replace);
  const prefixLen = diff.prefix.length;
  const suffixLen = diff.suffix.length;
  const effectiveReplace = diff.newMiddle;

  // Process in reverse order to preserve children indices
  for (let mi = matches.length - 1; mi >= 0; mi--) {
    const matchStart = matches[mi];
    const matchEnd = matchStart + search.length;

    // Shrink del/ins boundaries to only the changed middle
    const effectiveStart = matchStart + prefixLen;
    const effectiveEnd = matchEnd - suffixLen;

    // Nothing actually changed (identical strings) — skip
    if (effectiveStart === effectiveEnd && !effectiveReplace) continue;

    let firstRunIdx = -1;
    let lastRunIdx = -1;
    for (let ri = 0; ri < runs.length; ri++) {
      const rStart = runs[ri].startOffset;
      const rEnd = rStart + runs[ri].text.length;
      if (rStart < effectiveEnd && rEnd > effectiveStart) {
        if (firstRunIdx === -1) firstRunIdx = ri;
        lastRunIdx = ri;
      }
    }
    if (firstRunIdx === -1) continue;

    const newNodes: XNode[] = [];
    const delRuns: XNode[] = [];
    const delLen = effectiveEnd - effectiveStart;

    if (firstRunIdx === lastRunIdx) {
      // Single-run match
      const run = runs[firstRunIdx];
      const localStart = effectiveStart - run.startOffset;
      const beforeText = run.text.substring(0, localStart);
      const matchedText = run.text.substring(localStart, localStart + delLen);
      const afterText = run.text.substring(localStart + delLen);

      if (beforeText) newNodes.push(makeTextRun(beforeText, run.rPr));
      if (matchedText) {
        delRuns.push(makeDelTextRun(matchedText, run.rPr));
        newNodes.push(wrapInDel(delRuns, ctx));
      }
      if (effectiveReplace) {
        newNodes.push(wrapInIns([makeTextRun(effectiveReplace, run.rPr)], ctx));
      }
      if (afterText) newNodes.push(makeTextRun(afterText, run.rPr));

      children.splice(run.pIdx, 1, ...newNodes);
    } else {
      // Cross-run match
      for (let ri = firstRunIdx; ri <= lastRunIdx; ri++) {
        const run = runs[ri];
        if (ri === firstRunIdx) {
          const localStart = effectiveStart - run.startOffset;
          const beforeText = run.text.substring(0, localStart);
          const matchedText = run.text.substring(localStart);
          if (beforeText) newNodes.push(makeTextRun(beforeText, run.rPr));
          delRuns.push(makeDelTextRun(matchedText, run.rPr));
        } else if (ri === lastRunIdx) {
          const localEnd = effectiveEnd - run.startOffset;
          const matchedText = run.text.substring(0, localEnd);
          delRuns.push(makeDelTextRun(matchedText, run.rPr));
        } else {
          delRuns.push(makeDelTextRun(run.text, run.rPr));
        }
      }

      newNodes.push(wrapInDel(delRuns, ctx));
      const firstRun = runs[firstRunIdx];
      if (effectiveReplace) {
        newNodes.push(
          wrapInIns([makeTextRun(effectiveReplace, firstRun.rPr)], ctx),
        );
      }

      const lastRun = runs[lastRunIdx];
      const localEnd = effectiveEnd - lastRun.startOffset;
      const afterText = lastRun.text.substring(localEnd);
      if (afterText) newNodes.push(makeTextRun(afterText, lastRun.rPr));

      const pIdxStart = runs[firstRunIdx].pIdx;
      const pIdxEnd = runs[lastRunIdx].pIdx;
      children.splice(pIdxStart, pIdxEnd - pIdxStart + 1, ...newNodes);
    }
  }

  return matches.length;
}

// ---------------------------------------------------------------------------
// Accept / reject changes helpers
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Shared helpers for revision property cleanup
// ---------------------------------------------------------------------------

/** Tags that are move/revision range markers — should be removed on accept/reject. */
const RANGE_MARKER_TAGS = new Set([
  "w:moveFromRangeStart",
  "w:moveFromRangeEnd",
  "w:moveToRangeStart",
  "w:moveToRangeEnd",
]);

function isRangeMarker(node: XNode): boolean {
  for (const tag of RANGE_MARKER_TAGS) {
    if (node[tag] !== undefined) return true;
  }
  return false;
}

/** Remove a *Change element from a property container (accept mode — keep current props). */
function stripChangeElement(propChildren: XNode[], changeTag: string): void {
  for (let i = propChildren.length - 1; i >= 0; i--) {
    if (propChildren[i][changeTag] !== undefined) {
      propChildren.splice(i, 1);
    }
  }
}

/** Replace current properties with old ones stored in a *Change element (reject mode). */
function restoreFromChangeElement(
  propChildren: XNode[],
  changeTag: string,
  innerTag: string,
): void {
  const changeNode = findOne(propChildren, changeTag);
  if (!changeNode) return;
  const changeChildren = changeNode[changeTag] as XNode[];
  const oldProp = findOne(changeChildren, innerTag);
  // Remove the change element
  const idx = propChildren.indexOf(changeNode);
  if (idx !== -1) propChildren.splice(idx, 1);
  if (oldProp) {
    const oldChildren = oldProp[innerTag] as XNode[];
    propChildren.length = 0;
    propChildren.push(...oldChildren);
  }
}

/** Strip w:rPrChange from every w:r > w:rPr (accept mode). */
function stripRunPropertyChanges(pChildren: XNode[]): void {
  for (const child of pChildren) {
    if (child["w:r"]) {
      const rPr = findOne(child["w:r"] as XNode[], "w:rPr");
      if (rPr) stripChangeElement(rPr["w:rPr"] as XNode[], "w:rPrChange");
    }
  }
}

/** Restore w:rPr from w:rPrChange for every run (reject mode). */
function restoreRunPropertyChanges(pChildren: XNode[]): void {
  for (const child of pChildren) {
    if (child["w:r"]) {
      const rPr = findOne(child["w:r"] as XNode[], "w:rPr");
      if (rPr) restoreFromChangeElement(rPr["w:rPr"] as XNode[], "w:rPrChange", "w:rPr");
    }
  }
}

/** Remove w:ins and w:del markers from pPr > rPr (paragraph break markers). */
function cleanParagraphRevisionMarkers(pChildren: XNode[]): void {
  const pPr = findOne(pChildren, "w:pPr");
  if (!pPr) return;
  const pPrChildren = pPr["w:pPr"] as XNode[];
  const rPr = findOne(pPrChildren, "w:rPr");
  if (!rPr) return;
  const rPrChildren = rPr["w:rPr"] as XNode[];
  for (let i = rPrChildren.length - 1; i >= 0; i--) {
    if (rPrChildren[i]["w:ins"] !== undefined || rPrChildren[i]["w:del"] !== undefined) {
      rPrChildren.splice(i, 1);
    }
  }
  if (rPrChildren.length === 0) {
    const rPrIdx = pPrChildren.indexOf(rPr);
    if (rPrIdx !== -1) pPrChildren.splice(rPrIdx, 1);
  }
}

/** Post-process a paragraph after the main accept pass. */
function postProcessParagraphAccept(pChildren: XNode[]): void {
  cleanParagraphRevisionMarkers(pChildren);
  const pPr = findOne(pChildren, "w:pPr");
  if (pPr) stripChangeElement(pPr["w:pPr"] as XNode[], "w:pPrChange");
  stripRunPropertyChanges(pChildren);
}

/** Post-process a paragraph after the main reject pass. */
function postProcessParagraphReject(pChildren: XNode[]): void {
  cleanParagraphRevisionMarkers(pChildren);
  const pPr = findOne(pChildren, "w:pPr");
  if (pPr) restoreFromChangeElement(pPr["w:pPr"] as XNode[], "w:pPrChange", "w:pPr");
  restoreRunPropertyChanges(pChildren);
}

/** Accept tracked changes within a table (properties, rows, cells, paragraphs). */
function acceptChangesInTable(tblChildren: XNode[]): void {
  const tblPr = findOne(tblChildren, "w:tblPr");
  if (tblPr) stripChangeElement(tblPr["w:tblPr"] as XNode[], "w:tblPrChange");
  const rows = findAll(tblChildren, "w:tr");
  for (const row of rows) {
    const trChildren = row["w:tr"] as XNode[];
    const trPr = findOne(trChildren, "w:trPr");
    if (trPr) {
      stripChangeElement(trPr["w:trPr"] as XNode[], "w:trPrChange");
      stripChangeElement(trPr["w:trPr"] as XNode[], "w:ins");
      stripChangeElement(trPr["w:trPr"] as XNode[], "w:del");
    }
    const cells = findAll(trChildren, "w:tc");
    for (const cell of cells) {
      const tcChildren = cell["w:tc"] as XNode[];
      const tcPr = findOne(tcChildren, "w:tcPr");
      if (tcPr) stripChangeElement(tcPr["w:tcPr"] as XNode[], "w:tcPrChange");
      acceptChangesInNodes(tcChildren);
    }
  }
}

/** Reject tracked changes within a table. */
function rejectChangesInTable(tblChildren: XNode[]): void {
  const tblPr = findOne(tblChildren, "w:tblPr");
  if (tblPr) restoreFromChangeElement(tblPr["w:tblPr"] as XNode[], "w:tblPrChange", "w:tblPr");
  const rows = findAll(tblChildren, "w:tr");
  for (const row of rows) {
    const trChildren = row["w:tr"] as XNode[];
    const trPr = findOne(trChildren, "w:trPr");
    if (trPr) {
      restoreFromChangeElement(trPr["w:trPr"] as XNode[], "w:trPrChange", "w:trPr");
      stripChangeElement(trPr["w:trPr"] as XNode[], "w:ins");
      stripChangeElement(trPr["w:trPr"] as XNode[], "w:del");
    }
    const cells = findAll(trChildren, "w:tc");
    for (const cell of cells) {
      const tcChildren = cell["w:tc"] as XNode[];
      const tcPr = findOne(tcChildren, "w:tcPr");
      if (tcPr) restoreFromChangeElement(tcPr["w:tcPr"] as XNode[], "w:tcPrChange", "w:tcPr");
      rejectChangesInNodes(tcChildren);
    }
  }
}

// ---------------------------------------------------------------------------
// Accept / reject — main traversal
// ---------------------------------------------------------------------------

/**
 * Accept all tracked changes: remove w:del/w:moveFrom elements entirely,
 * unwrap w:ins/w:moveTo so their children become normal content,
 * and strip all *Change revision properties.
 */
export function acceptChangesInNodes(nodes: XNode[]): void {
  for (let i = nodes.length - 1; i >= 0; i--) {
    const node = nodes[i];
    if (node["w:del"] || node["w:moveFrom"]) {
      nodes.splice(i, 1);
    } else if (node["w:ins"] || node["w:moveTo"]) {
      const tag = node["w:ins"] !== undefined ? "w:ins" : "w:moveTo";
      const children = node[tag] as XNode[];
      nodes.splice(i, 1, ...children);
    } else if (isRangeMarker(node)) {
      nodes.splice(i, 1);
    } else if (node["w:p"]) {
      const pChildren = node["w:p"] as XNode[];
      acceptChangesInNodes(pChildren);
      postProcessParagraphAccept(pChildren);
    } else if (node["w:tbl"]) {
      acceptChangesInTable(node["w:tbl"]);
    } else if (node["w:sdt"]) {
      const sdtContent = findOne(node["w:sdt"] as XNode[], "w:sdtContent");
      if (sdtContent) acceptChangesInNodes(sdtContent["w:sdtContent"]);
    } else if (node["w:sectPr"]) {
      stripChangeElement(node["w:sectPr"] as XNode[], "w:sectPrChange");
    }
  }
}

/**
 * Reject all tracked changes: remove w:ins/w:moveTo elements entirely,
 * unwrap w:del/w:moveFrom so their runs become normal (converting w:delText → w:t),
 * and restore old properties from *Change elements.
 */
export function rejectChangesInNodes(nodes: XNode[]): void {
  for (let i = nodes.length - 1; i >= 0; i--) {
    const node = nodes[i];
    if (node["w:ins"] || node["w:moveTo"]) {
      nodes.splice(i, 1);
    } else if (node["w:del"] || node["w:moveFrom"]) {
      const tag = node["w:del"] !== undefined ? "w:del" : "w:moveFrom";
      const delChildren = node[tag] as XNode[];
      const runs: XNode[] = [];
      for (const dc of delChildren) {
        if (dc["w:r"]) {
          const runC = dc["w:r"] as XNode[];
          for (const rc of runC) {
            if (rc["w:delText"]) {
              const text = rc["w:delText"];
              delete rc["w:delText"];
              rc["w:t"] = text;
            }
          }
          runs.push(dc);
        }
      }
      nodes.splice(i, 1, ...runs);
    } else if (isRangeMarker(node)) {
      nodes.splice(i, 1);
    } else if (node["w:p"]) {
      const pChildren = node["w:p"] as XNode[];
      rejectChangesInNodes(pChildren);
      postProcessParagraphReject(pChildren);
    } else if (node["w:tbl"]) {
      rejectChangesInTable(node["w:tbl"]);
    } else if (node["w:sdt"]) {
      const sdtContent = findOne(node["w:sdt"] as XNode[], "w:sdtContent");
      if (sdtContent) rejectChangesInNodes(sdtContent["w:sdtContent"]);
    } else if (node["w:sectPr"]) {
      restoreFromChangeElement(node["w:sectPr"] as XNode[], "w:sectPrChange", "w:sectPr");
    }
  }
}

/** Extract text from all paragraphs in a parsed header/footer XML */
export function extractTextFromHdrFtr(parsed: XNode[]): string {
  const rootEl = parsed.find((n: XNode) => n["w:hdr"] || n["w:ftr"]);
  if (!rootEl) return "";
  const children = rootEl["w:hdr"] ?? rootEl["w:ftr"];
  const lines: string[] = [];
  for (const child of children) {
    if (child["w:p"]) {
      const text = extractParagraphText(child["w:p"]);
      if (text.trim()) lines.push(text);
    }
  }
  return lines.join("\n");
}
