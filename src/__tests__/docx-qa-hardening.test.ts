/**
 * QA hardening regression tests (safety/corruption cluster).
 *
 * Findings covered:
 *   C2 — XML-illegal control chars silently corrupt the .docx
 *   C3 — empty search query infinite-loops / bricks the server
 *   H4 — format_text / highlight_text empty search → RangeError spin
 *   H5 — insert_table omits required <w:tblGrid>
 *   H2 — whole-cell replace / delete-trailing leaves a <w:tc> ending in <w:tbl>
 *   L1 — insert_table clamp/pad behavior documented
 *
 * Ground truth is asserted against raw XML (readRawDocXml), xmllint --noout, and
 * python-docx — the same validators the QA repros used — never against
 * readDocument output (which embeds the temp path and would false-fail).
 */
import { describe, it, expect, afterEach } from "vitest";
import { execFileSync } from "child_process";
import * as fs from "fs/promises";
import { XMLParser } from "fast-xml-parser";
import {
  createTmpDoc,
  createDocWithSdt,
  cleanupTmpFiles,
  readRawDocXml,
  readRawCommentsXml,
  tmpDocxPath,
  trackTmpFile,
  writeMinimalDocx,
} from "./helpers.js";
import {
  createDocument,
  editParagraphs,
  setHeadings,
  setParagraphFormats,
  insertParagraphs,
  insertTableParagraphs,
  replaceTexts,
  editTableCells,
  editTableParagraphs,
  deleteTableParagraphs,
  insertTable,
  searchText,
  searchTextStructured,
  readTableStructure,
  readTableCell,
  readTableCellStructured,
  getParagraphFormat,
  getParagraphFormatStructured,
  getDocumentInfoStructured,
  ensureAnchors,
  ensureAnchorsStructured,
  rejectAllChanges,
  listImages,
  readComments,
  addComment,
  addComments,
  EngineError,
  ErrorCode,
  formatText,
  highlightText,
  setPageLayout,
} from "../docx-engine.js";

afterEach(cleanupTmpFiles);

// =========================================================================
// External-validator capability detection (CI portability)
//
// This suite cross-checks engine output against two EXTERNAL validators —
// `xmllint` (XML well-formedness) and `python-docx` (Word can open the file).
// Neither is guaranteed on every CI runner. When a validator is ABSENT we
// degrade gracefully: the external check becomes a pass-through, and every
// assertion that consumed a python-docx VALUE instead derives the same value
// from the raw `word/document.xml` (parsed with fast-xml-parser, the engine's
// own parser config) so the test still asserts its real behavior — it never
// becomes a silent no-op. When the tools ARE present (e.g. CI installs them,
// or a dev box has them) the FULL external validation still runs unchanged.
// =========================================================================

/** True if `xmllint --version` runs (the validator is installed + on PATH). */
const HAS_XMLLINT: boolean = (() => {
  try {
    execFileSync("xmllint", ["--version"], { stdio: "ignore" });
    return true;
  } catch {
    return false;
  }
})();

/**
 * Resolve a python interpreter that can `import docx`. Probes the hardcoded CI
 * venv path first, then the PATH interpreters. `null` if none can import it.
 */
const PYTHON: string | null = (() => {
  for (const cand of ["/tmp/docx-venv/bin/python", "python3", "python"]) {
    try {
      execFileSync(cand, ["-c", "import docx"], { stdio: "ignore" });
      return cand;
    } catch {
      // try the next candidate
    }
  }
  return null;
})();

/** True if some interpreter with python-docx was found. */
const HAS_PYDOCX: boolean = PYTHON !== null;

// ---------------------------------------------------------------------------
// Self-contained fallbacks (used when python-docx is unavailable). These parse
// word/document.xml with fast-xml-parser in the SAME preserveOrder mode the
// engine uses, then replicate the exact python-docx accessor each helper stood
// in for. Verified against python-docx 1.1.2 ground truth:
//   • paragraph.text  = run texts concatenated; <w:br/>/<w:cr/> → "\n",
//                       <w:tab/> → "\t"; <w:del>/<w:delText> excluded.
//   • cell.text       = the cell's paragraphs' .text joined by "\n".
//   • table.columns   = one per <w:gridCol> in the table's <w:tblGrid>.
//   • cell.tables     = nested <w:tbl> direct children of the <w:tc>.
//   • document.paragraphs = <w:p> that are DIRECT children of <w:body>.
// ---------------------------------------------------------------------------

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type RawNode = any;

const rawParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  preserveOrder: true,
  trimValues: false,
  processEntities: true,
  parseTagValue: false,
});

/** preserveOrder tag name of a node (the single non-meta key). */
function rawTag(node: RawNode): string | null {
  for (const k of Object.keys(node)) {
    if (k !== ":@" && k !== "#text" && k !== "#comment") return k;
  }
  return null;
}

/** preserveOrder child array of a node. */
function rawChildren(node: RawNode): RawNode[] {
  const t = rawTag(node);
  return t ? (node[t] ?? []) : [];
}

/** First descendant (BFS over the supplied roots) with the given tag. */
function firstByTag(roots: RawNode[], tag: string): RawNode | undefined {
  const queue = [...roots];
  while (queue.length) {
    const n = queue.shift();
    if (rawTag(n) === tag) return n;
    queue.push(...rawChildren(n));
  }
  return undefined;
}

/** Parse word/document.xml and return the <w:body>'s child node array. */
async function rawBodyChildren(filePath: string): Promise<RawNode[]> {
  const xml = await readRawDocXml(filePath);
  const tree: RawNode[] = rawParser.parse(xml);
  const doc = tree.find((n) => rawTag(n) === "w:document");
  const body = doc ? rawChildren(doc).find((n) => rawTag(n) === "w:body") : undefined;
  return body ? rawChildren(body) : [];
}

/**
 * python-docx run.text / paragraph.text semantics for a single <w:p> node:
 * concatenate <w:t> text across runs (descending into <w:ins>), mapping
 * <w:br/> and <w:cr/> to "\n" and <w:tab/> to "\t"; <w:del>/<w:delText> are
 * excluded (rejected/inserted text is what python-docx surfaces).
 */
function paragraphText(pNode: RawNode): string {
  let out = "";
  const walk = (nodes: RawNode[]) => {
    for (const n of nodes) {
      const t = rawTag(n);
      if (t === "w:del" || t === "w:delText") continue; // deletion markup → not in .text
      if (t === "w:t") {
        for (const c of rawChildren(n)) if ("#text" in c) out += String(c["#text"]);
      } else if (t === "w:br" || t === "w:cr") {
        out += "\n";
      } else if (t === "w:tab") {
        out += "\t";
      } else if (t === "w:r" || t === "w:ins" || t === "w:smartTag" || t === "w:hyperlink") {
        walk(rawChildren(n)); // descend into run containers
      }
    }
  };
  walk(rawChildren(pNode));
  return out;
}

/** Fallback for python-docx `d.paragraphs` text list (direct-body <w:p> only). */
async function rawBodyParagraphTexts(filePath: string): Promise<string[]> {
  const body = await rawBodyChildren(filePath);
  return body.filter((n) => rawTag(n) === "w:p").map(paragraphText);
}

/** Fallback for python-docx `tables[ti].rows[ri].cells[ci]` <w:tc> node. */
async function rawCellNode(
  filePath: string,
  tableIndex: number,
  rowIndex: number,
  colIndex: number,
): Promise<RawNode | undefined> {
  const body = await rawBodyChildren(filePath);
  const tbl = body.filter((n) => rawTag(n) === "w:tbl")[tableIndex];
  if (!tbl) return undefined;
  const row = rawChildren(tbl).filter((n) => rawTag(n) === "w:tr")[rowIndex];
  if (!row) return undefined;
  return rawChildren(row).filter((n) => rawTag(n) === "w:tc")[colIndex];
}

/** Fallback for python-docx `cell.text`: cell paragraphs' text joined by "\n". */
async function rawFirstCellText(filePath: string): Promise<string> {
  const tc = await rawCellNode(filePath, 0, 0, 0);
  if (!tc) throw new Error("no cell[0,0] in first table");
  return rawChildren(tc)
    .filter((n) => rawTag(n) === "w:p")
    .map(paragraphText)
    .join("\n");
}

/** Fallback for python-docx `len(tables[0].columns)`: count <w:gridCol> in grid. */
async function rawFirstTableColumnCount(filePath: string): Promise<number> {
  const body = await rawBodyChildren(filePath);
  const tbl = body.find((n) => rawTag(n) === "w:tbl");
  if (!tbl) throw new Error("no table in body");
  const grid = rawChildren(tbl).find((n) => rawTag(n) === "w:tblGrid");
  return grid ? rawChildren(grid).filter((n) => rawTag(n) === "w:gridCol").length : 0;
}

/** Fallback for python-docx `len(cell[0,0].tables)`: nested <w:tbl> in the cell. */
async function rawFirstCellNestedTableCount(filePath: string): Promise<number> {
  const tc = await rawCellNode(filePath, 0, 0, 0);
  if (!tc) throw new Error("no cell[0,0] in first table");
  return rawChildren(tc).filter((n) => rawTag(n) === "w:tbl").length;
}

/** True if `word/document.xml` is well-formed XML (per xmllint). */
function xmlIsWellFormed(filePath: string): boolean {
  // No validator on this host → pass through. Behavior is still pinned by the
  // engine round-trip plus the raw-XML structural assertions in each test.
  if (!HAS_XMLLINT) return true;
  // Extract the part to a temp path xmllint can read, then validate.
  try {
    execFileSync("bash", [
      "-c",
      `unzip -p "${filePath}" word/document.xml | xmllint --noout -`,
    ]);
    return true;
  } catch {
    return false;
  }
}

/** True if python-docx can open the file without raising. */
function pythonDocxOpens(filePath: string): boolean {
  // No python-docx on this host → pass through (the well-formedness + raw-XML
  // structure assertions remain the load-bearing checks).
  if (!HAS_PYDOCX) return true;
  try {
    execFileSync(PYTHON as string, ["-c", `import docx; docx.Document(${JSON.stringify(filePath)})`]);
    return true;
  } catch {
    return false;
  }
}

/** Run an arbitrary python-docx snippet, returning stdout (throws on python error). */
function runPython(snippet: string): string {
  if (!HAS_PYDOCX) throw new Error("python-docx unavailable: caller must guard with HAS_PYDOCX");
  return execFileSync(PYTHON as string, ["-c", snippet], { encoding: "utf8" });
}

/** Assert comments.xml is well-formed via xmllint (no-op when xmllint absent). */
function expectCommentsXmlWellFormed(filePath: string): void {
  if (!HAS_XMLLINT) return;
  execFileSync("bash", ["-c", `unzip -p "${filePath}" word/comments.xml | xmllint --noout -`]);
}

/**
 * Build a DOCX whose outer cell[0,0] children are
 *   [w:p "BEFORE", nested w:tbl (with tblGrid), w:p "AFTER"]
 * — the canonical Word layout where a paragraph terminates the cell after a
 * nested table. The outer table carries a tblGrid too.
 */
async function createDocWithNestedTrailingPara(): Promise<string> {
  const p = tmpDocxPath();
  trackTmpFile(p);
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:tbl>
  <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
  <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
  <w:tr>
    <w:tc>
      <w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>
      <w:p><w:r><w:t>BEFORE</w:t></w:r></w:p>
      <w:tbl>
        <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
        <w:tblGrid><w:gridCol w:w="2500"/></w:tblGrid>
        <w:tr>
          <w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p><w:r><w:t>NESTED</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>
      <w:p><w:r><w:t>AFTER</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
</w:body>
</w:document>`;
  await writeMinimalDocx(p, documentXml);
  return p;
}

/**
 * Return the tag-name sequence (e.g. ["p","tbl","p"]) of the block-level
 * children of the OUTER table's cell[0,0]. When python-docx is present we read
 * via its lxml; otherwise we derive the identical sequence directly from the
 * raw word/document.xml (both reflect the on-disk structure, not an engine
 * re-parse — that's the point of this helper).
 */
async function outerCellBlockChildren(filePath: string): Promise<string[]> {
  if (HAS_PYDOCX) {
    const py = `
import docx, json
from docx.oxml.ns import qn
d = docx.Document(${JSON.stringify(filePath)})
tc = d.tables[0].rows[0].cells[0]._tc
seq = []
for ch in tc:
    tag = ch.tag.split('}')[-1]
    if tag in ('p', 'tbl'):
        seq.append(tag)
print(json.dumps(seq))
`;
    return JSON.parse(runPython(py));
  }
  const tc = await rawCellNode(filePath, 0, 0, 0);
  if (!tc) throw new Error("no cell[0,0] in first table");
  return rawChildren(tc)
    .map((n) => rawTag(n))
    .filter((t): t is string => t === "w:p" || t === "w:tbl")
    .map((t) => (t === "w:p" ? "p" : "tbl"));
}

// =========================================================================
// C2 — illegal XML control chars must be stripped, not written raw
// =========================================================================

describe("C2: illegal XML control characters are sanitized", () => {
  // The XML-1.0-illegal C0 set the spec forbids in text (everything except
  // \t \n \r). Each must be stripped before reaching <w:t>.
  const ILLEGAL = [0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x0b, 0x0c, 0x0e, 0x1b, 0x1f];

  it("createDocument strips control chars and produces a well-formed, openable file", async () => {
    const dirty = "Page1" + String.fromCharCode(0x0c) + "Page2" + String.fromCharCode(0x00) + "End";
    const p = await createTmpDoc(dirty, "title");

    const xml = await readRawDocXml(p);
    // No raw illegal codepoint survives anywhere in the XML.
    for (const cp of [0x0c, 0x00]) {
      expect(xml).not.toContain(String.fromCharCode(cp));
    }
    // Surrounding text is intact and contiguous (the char is removed, not replaced).
    expect(xml).toContain("Page1Page2End");
    expect(xmlIsWellFormed(p)).toBe(true);
    expect(pythonDocxOpens(p)).toBe(true);
  });

  it("editParagraphs (untracked) strips every illegal C0 char", async () => {
    const p = await createTmpDoc("seed");
    const dirty = "a" + ILLEGAL.map((c) => String.fromCharCode(c)).join("") + "b";
    await editParagraphs(p, [{ paragraphIndex: 0, newText: dirty }], false);

    const xml = await readRawDocXml(p);
    for (const cp of ILLEGAL) {
      expect(xml).not.toContain(String.fromCharCode(cp));
    }
    expect(xml).toContain("<w:t xml:space=\"preserve\">ab</w:t>");
    expect(xmlIsWellFormed(p)).toBe(true);
    expect(pythonDocxOpens(p)).toBe(true);
  });

  it("editParagraphs (tracked) strips illegal chars in the inserted run", async () => {
    const p = await createTmpDoc("hello world");
    await editParagraphs(p, [{ paragraphIndex: 0, newText: "good" + String.fromCharCode(0x0b) + "bye" }], true);

    const xml = await readRawDocXml(p);
    expect(xml).not.toContain(String.fromCharCode(0x0b));
    expect(xml).toContain("goodbye");
    expect(xmlIsWellFormed(p)).toBe(true);
    expect(pythonDocxOpens(p)).toBe(true);
  });

  it("insertParagraphs strips illegal chars", async () => {
    const p = await createTmpDoc("anchor");
    await insertParagraphs(p, [{ text: "x" + String.fromCharCode(0x0c) + "y", position: 0 }], false);

    const xml = await readRawDocXml(p);
    expect(xml).not.toContain(String.fromCharCode(0x0c));
    expect(xml).toContain("xy");
    expect(xmlIsWellFormed(p)).toBe(true);
    expect(pythonDocxOpens(p)).toBe(true);
  });

  it("replaceTexts (untracked) strips illegal chars in the replacement", async () => {
    const p = await createTmpDoc("replace me please");
    await replaceTexts(p, [{ search: "me", replace: "ME" + String.fromCharCode(0x1b) }], false);

    const xml = await readRawDocXml(p);
    expect(xml).not.toContain(String.fromCharCode(0x1b));
    expect(xml).toContain("ME");
    expect(xmlIsWellFormed(p)).toBe(true);
    expect(pythonDocxOpens(p)).toBe(true);
  });

  it("editTableCells strips illegal chars (tracked + untracked)", async () => {
    for (const track of [true, false]) {
      const p = await createTmpDoc("intro");
      await insertTable(p, -1, 1, 1, [["cell"]]);
      await editTableCells(
        p,
        [{ blockIndex: 1, rowIndex: 0, colIndex: 0, newText: "C" + String.fromCharCode(0x07) + "D" }],
        track,
      );
      const xml = await readRawDocXml(p);
      expect(xml).not.toContain(String.fromCharCode(0x07));
      expect(xml).toContain("CD");
      expect(xmlIsWellFormed(p)).toBe(true);
      expect(pythonDocxOpens(p)).toBe(true);
    }
  });

  it("preserves the LEGAL whitespace char \\t verbatim (tabs are not line-ends)", async () => {
    // \t is legal XML data and not a line-ending, so it must survive verbatim.
    // \r and \n ARE line-ends: the edit pipeline normalizes them to breaks
    // (see the H1 suite below), so they are not asserted to survive here. This
    // is distinct from C2's job, which is stripping the *illegal* control set.
    const p = await createTmpDoc("seed");
    await editParagraphs(p, [{ paragraphIndex: 0, newText: "tab\there and more" }], false);
    const xml = await readRawDocXml(p);
    expect(xml).toContain("\t");
    expect(xml).toContain("tab\there and more");
    expect(xmlIsWellFormed(p)).toBe(true);
    expect(pythonDocxOpens(p)).toBe(true);
  });
});

// =========================================================================
// C3 — empty search query must not infinite-loop
// =========================================================================

describe("C3: empty search query is bounded", () => {
  it("searchText('') returns fast with zero matches (no hang)", async () => {
    const p = await createTmpDoc("alpha beta gamma");
    // If the engine guard is missing this never resolves; the 4s race makes the
    // failure a clean timeout instead of a hung worker.
    const result = await Promise.race([
      searchText(p, ""),
      new Promise<string>((_, rej) => setTimeout(() => rej(new Error("HANG: searchText('') did not return")), 4000)),
    ]);
    expect(typeof result).toBe("string");
    // Either a clean "no matches" or a 0-count summary; must not claim matches.
    expect(result.toLowerCase()).not.toMatch(/found\s+[1-9]/);
  });

  it("replaceTexts still rejects an empty search (existing guard, untracked)", async () => {
    const p = await createTmpDoc("alpha beta");
    await expect(
      Promise.race([
        replaceTexts(p, [{ search: "", replace: "x" }], false),
        new Promise((_, rej) => setTimeout(() => rej(new Error("HANG")), 4000)),
      ]),
    ).rejects.toThrow();
  });

  it("replaceTexts with empty search is bounded under track_changes too", async () => {
    const p = await createTmpDoc("alpha beta");
    await expect(
      Promise.race([
        replaceTexts(p, [{ search: "", replace: "x" }], true),
        new Promise((_, rej) => setTimeout(() => rej(new Error("HANG")), 4000)),
      ]),
    ).rejects.toThrow();
  });
});

// =========================================================================
// H4 — format_text / highlight_text empty search must not spin to RangeError
// =========================================================================

describe("H4: empty search in format/highlight is bounded", () => {
  it("formatText('') returns fast without RangeError / spin", async () => {
    const p = await createTmpDoc("alpha beta gamma");
    const result = await Promise.race([
      formatText(p, "", { bold: true }),
      new Promise<string>((_, rej) => setTimeout(() => rej(new Error("HANG: formatText('') did not return")), 4000)),
    ]);
    expect(typeof result).toBe("string");
    // Must not surface the cryptic RangeError.
    expect(result).not.toContain("Invalid array length");
  });

  it("highlightText('') returns fast without RangeError / spin", async () => {
    const p = await createTmpDoc("alpha beta gamma");
    const result = await Promise.race([
      highlightText(p, "", "yellow"),
      new Promise<string>((_, rej) => setTimeout(() => rej(new Error("HANG: highlightText('') did not return")), 4000)),
    ]);
    expect(typeof result).toBe("string");
    expect(result).not.toContain("Invalid array length");
  });
});

// =========================================================================
// H5 — insert_table must emit <w:tblGrid> with `cols` <w:gridCol/>
// =========================================================================

describe("H5: insert_table emits a conformant tblGrid", () => {
  it("emits one <w:gridCol/> per column, after <w:tblPr>", async () => {
    const p = await createTmpDoc("intro");
    await insertTable(p, -1, 2, 3, [["a", "b", "c"], ["d", "e", "f"]]);

    const xml = await readRawDocXml(p);
    expect(xml).toContain("<w:tblGrid>");
    // Exactly cols (3) gridCol children.
    const gridColCount = (xml.match(/<w:gridCol[ /]/g) ?? []).length;
    expect(gridColCount).toBe(3);
    // tblGrid must come after tblPr and before the first row.
    const gridIdx = xml.indexOf("<w:tblGrid>");
    const tblPrIdx = xml.indexOf("<w:tblPr>");
    const firstTrIdx = xml.indexOf("<w:tr>");
    expect(tblPrIdx).toBeGreaterThan(-1);
    expect(gridIdx).toBeGreaterThan(tblPrIdx);
    expect(firstTrIdx).toBeGreaterThan(gridIdx);
  });

  it("python-docx opens it and reports the right column count (no InvalidXmlError)", async () => {
    const p = await createTmpDoc("intro");
    await insertTable(p, -1, 2, 4);

    if (HAS_PYDOCX) {
      const out = runPython(`
import docx
d = docx.Document(${JSON.stringify(p)})
print(len(d.tables[0].columns))
`).trim();
      expect(out).toBe("4");
    } else {
      // Fallback: python-docx derives table.columns from <w:gridCol> count; we
      // assert the same value straight from the grid so this is no no-op.
      expect(await rawFirstTableColumnCount(p)).toBe(4);
    }
  });
});

// =========================================================================
// H2 — a <w:tc> must end with <w:p>, never a nested <w:tbl>
// =========================================================================

describe("H2: cell content must end with a paragraph", () => {
  it("whole-cell untracked replace keeps a trailing <w:p> after a nested table", async () => {
    const p = await createDocWithNestedTrailingPara();
    // sanity: starts as [p, tbl, p]
    expect(await outerCellBlockChildren(p)).toEqual(["p", "tbl", "p"]);

    // editTableCells(untracked) replaces the whole cell's paragraph content.
    await editTableCells(p, [{ blockIndex: 0, rowIndex: 0, colIndex: 0, newText: "NEW" }], false);

    const seq = await outerCellBlockChildren(p);
    expect(seq[seq.length - 1]).toBe("p"); // last block child must be a paragraph
    // The nested table must survive.
    expect(seq).toContain("tbl");
    expect(pythonDocxOpens(p)).toBe(true);
    // Nested table still present (python-docx cell.tables, or its raw-XML
    // equivalent — the nested <w:tbl> direct child of the cell).
    if (HAS_PYDOCX) {
      const nested = runPython(`
import docx
d = docx.Document(${JSON.stringify(p)})
print(len(d.tables[0].rows[0].cells[0].tables))
`).trim();
      expect(nested).toBe("1");
    } else {
      expect(await rawFirstCellNestedTableCount(p)).toBe(1);
    }
  });

  it("deleting the trailing paragraph after a nested table re-adds a blank <w:p>", async () => {
    const p = await createDocWithNestedTrailingPara();
    expect(await outerCellBlockChildren(p)).toEqual(["p", "tbl", "p"]);

    // Cell-local paragraph indices count only w:p children: [0]=BEFORE, [1]=AFTER.
    // Delete the trailing AFTER paragraph (untracked).
    await deleteTableParagraphs(p, [{ blockIndex: 0, rowIndex: 0, colIndex: 0, paragraphIndex: 1 }], false);

    const seq = await outerCellBlockChildren(p);
    expect(seq[seq.length - 1]).toBe("p"); // must NOT end in tbl
    expect(seq).toContain("tbl");
    expect(pythonDocxOpens(p)).toBe(true);
  });

  it("deleting the FIRST paragraph (before the nested table) is unaffected and stays valid", async () => {
    const p = await createDocWithNestedTrailingPara();
    await deleteTableParagraphs(p, [{ blockIndex: 0, rowIndex: 0, colIndex: 0, paragraphIndex: 0 }], false);
    const seq = await outerCellBlockChildren(p);
    // [tbl, p] — already ends in p; the trailing AFTER paragraph remains.
    expect(seq[seq.length - 1]).toBe("p");
    expect(pythonDocxOpens(p)).toBe(true);
  });
});

// =========================================================================
// L1 — insert_table clamp/pad behavior is documented in the tool description
// =========================================================================

describe("L1: insert_table clamp/pad behavior is documented", () => {
  it("the data field description mentions the clamp & pad rules", async () => {
    // Read the MCP layer source and assert the documentation sentence exists.
    const src = await fs.readFile(new URL("../index.ts", import.meta.url), "utf8");
    // Find the insert_table data description.
    const idx = src.indexOf('"insert_table"');
    expect(idx).toBeGreaterThan(-1);
    const region = src.slice(idx, idx + 1200);
    expect(region).toMatch(/beyond\s+rows\s*[x×*]\s*cols.*ignored/i);
    expect(region).toMatch(/short rows.*padded/i);
  });

  it("behavior is unchanged: overflow data is silently clamped (regression lock)", async () => {
    const p = await createTmpDoc("intro");
    await insertTable(p, -1, 2, 2, [["a", "b", "EXTRA"], ["d", "e"], ["f", "g", "ROW3"]]);
    const xml = await readRawDocXml(p);
    expect(xml).toContain("a");
    expect(xml).toContain("e");
    // Overflow cell + overflow row are dropped (documented clamp).
    expect(xml).not.toContain("EXTRA");
    expect(xml).not.toContain("ROW3");
  });
});

// =========================================================================
// Newline fidelity + integer guards cluster (H1, M1, M2, M3, M7).
//
// H1 — CRLF / lone CR must normalize to "\n" before splitting so no raw \r
//      (0x0D) survives inside a run, across every #4 edit/insert tool.
// M1 — replace_texts must emit "\n" as a <w:br/> soft break (not a literal LF).
// M2 — untracked editTableParagraphs must emit "\n" as a <w:br/> soft break,
//      matching its tracked path.
// M3/M7 — the #5/#6 navigation helpers must reject non-integer / wrong-type
//      indices with EngineError(INDEX_OUT_OF_RANGE), never a raw TypeError and
//      never a silent string-coercion mutation.
// =========================================================================

/** Count raw carriage-return (0x0D) bytes in the saved word/document.xml. */
function rawCRCount(xml: string): number {
  let n = 0;
  for (let i = 0; i < xml.length; i++) if (xml.charCodeAt(i) === 0x0d) n++;
  return n;
}

/**
 * python-docx body paragraph texts (after XML §2.11 line-end normalization).
 * Falls back to a raw-XML derivation of the same value when python-docx is
 * absent (direct-body <w:p> texts; see rawBodyParagraphTexts).
 */
async function pyBodyParagraphs(filePath: string): Promise<string[]> {
  if (HAS_PYDOCX) {
    const py = `
import docx, json
d = docx.Document(${JSON.stringify(filePath)})
print(json.dumps([p.text for p in d.paragraphs]))
`;
    return JSON.parse(runPython(py));
  }
  return rawBodyParagraphTexts(filePath);
}

/**
 * python-docx text of the first cell of the first table. Falls back to a
 * raw-XML derivation of the same value when python-docx is absent (the cell's
 * paragraph texts joined by "\n"; see rawFirstCellText).
 */
async function pyFirstCellText(filePath: string): Promise<string> {
  if (HAS_PYDOCX) {
    const py = `
import docx, json
d = docx.Document(${JSON.stringify(filePath)})
print(json.dumps(d.tables[0].rows[0].cells[0].text))
`;
    return JSON.parse(runPython(py));
  }
  return rawFirstCellText(filePath);
}

/**
 * Build a doc whose block 0 is a plain paragraph and block 1 is a 1×1 table
 * holding "ORIG". Lets a test target the table at a non-zero block index and
 * detect silent mutation by a wrong-type ('1') index.
 */
async function createDocParaThenTable(): Promise<string> {
  const p = tmpDocxPath();
  trackTmpFile(p);
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:p><w:r><w:t>lead paragraph</w:t></w:r></w:p>
<w:tbl>
  <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
  <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p><w:r><w:t>ORIG</w:t></w:r></w:p></w:tc>
  </w:tr>
</w:tbl>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
</w:body>
</w:document>`;
  await writeMinimalDocx(p, documentXml);
  return p;
}

describe("H1: CRLF / lone CR is normalized so no raw carriage return reaches a run", () => {
  it("editParagraphs (untracked) splits CRLF into real <w:p> with zero raw \\r", async () => {
    const p = await createTmpDoc("orig");
    await editParagraphs(p, [{ paragraphIndex: 0, newText: "a\r\nb\r\nc" }], false);

    const xml = await readRawDocXml(p);
    expect(rawCRCount(xml)).toBe(0);
    // python-docx sees three clean lines (no embedded newline char in any run).
    expect(await pyBodyParagraphs(p)).toEqual(["a", "b", "c"]);
    expect(xmlIsWellFormed(p)).toBe(true);
  });

  it("editParagraphs (tracked) emits <w:br/> soft breaks with zero raw \\r", async () => {
    const p = await createTmpDoc("orig");
    await editParagraphs(p, [{ paragraphIndex: 0, newText: "a\r\nb" }], true);

    const xml = await readRawDocXml(p);
    expect(rawCRCount(xml)).toBe(0);
    expect(xml).toContain("<w:br/>");
    // The inserted run text holds no embedded newline character.
    expect(xml).not.toMatch(/<w:t[^>]*>[^<]*\n[^<]*<\/w:t>/);
    expect(xmlIsWellFormed(p)).toBe(true);
  });

  it("editParagraphs normalizes a lone \\r (no \\n) into a paragraph break", async () => {
    const p = await createTmpDoc("orig");
    await editParagraphs(p, [{ paragraphIndex: 0, newText: "a\rb" }], false);

    const xml = await readRawDocXml(p);
    expect(rawCRCount(xml)).toBe(0);
    expect(await pyBodyParagraphs(p)).toEqual(["a", "b"]);
  });

  it("insertParagraphs (untracked) splits CRLF with zero raw \\r", async () => {
    const p = await createTmpDoc("anchor");
    await insertParagraphs(p, [{ text: "ins1\r\nins2", position: -1 }], false);

    const xml = await readRawDocXml(p);
    expect(rawCRCount(xml)).toBe(0);
    expect(await pyBodyParagraphs(p)).toContain("ins1");
    expect(await pyBodyParagraphs(p)).toContain("ins2");
    expect(xmlIsWellFormed(p)).toBe(true);
  });

  it("editTableCells (untracked) splits CRLF cell text with zero raw \\r", async () => {
    const p = await createTmpDoc("seed");
    await insertTable(p, -1, 1, 1);
    await editTableCells(p, [{ blockIndex: 1, rowIndex: 0, colIndex: 0, newText: "r1\r\nr2" }], false);

    const xml = await readRawDocXml(p);
    expect(rawCRCount(xml)).toBe(0);
    // Two real paragraphs in the cell → python-docx joins them with one "\n".
    expect(await pyFirstCellText(p)).toBe("r1\nr2");
    expect(xmlIsWellFormed(p)).toBe(true);
  });

  it("editTableCells (tracked) emits <w:br/> for CRLF with zero raw \\r", async () => {
    const p = await createTmpDoc("seed");
    await insertTable(p, -1, 1, 1);
    await editTableCells(p, [{ blockIndex: 1, rowIndex: 0, colIndex: 0, newText: "r1\r\nr2" }], true);

    const xml = await readRawDocXml(p);
    expect(rawCRCount(xml)).toBe(0);
    expect(xml).toContain("<w:br/>");
    expect(xmlIsWellFormed(p)).toBe(true);
  });
});

describe("M1: replace_texts renders a replacement \\n as a soft break, not a literal LF", () => {
  it("untracked replacement emits <w:br/> and no literal newline in <w:t>", async () => {
    const p = await createTmpDoc("hello world");
    await replaceTexts(p, [{ search: "world", replace: "big\nplanet" }], false);

    const xml = await readRawDocXml(p);
    expect(xml).toContain("<w:br/>");
    // No <w:t> node contains an embedded newline character.
    expect(xml).not.toMatch(/<w:t[^>]*>[^<]*\n[^<]*<\/w:t>/);
    // python-docx still sees the break as a newline in the paragraph text.
    expect((await pyBodyParagraphs(p))[0]).toBe("hello big\nplanet");
    expect(xmlIsWellFormed(p)).toBe(true);
  });

  it("tracked replacement emits <w:br/> inside the <w:ins> run, no literal LF", async () => {
    const p = await createTmpDoc("hello world");
    await replaceTexts(p, [{ search: "world", replace: "big\nplanet" }], true, "QA");

    const xml = await readRawDocXml(p);
    expect(xml).toContain("<w:ins");
    expect(xml).toContain("<w:br/>");
    expect(xml).not.toMatch(/<w:t[^>]*>[^<]*\n[^<]*<\/w:t>/);
    expect(xmlIsWellFormed(p)).toBe(true);
  });

  it("normalizes CRLF in the replacement (zero raw \\r)", async () => {
    const p = await createTmpDoc("hello world");
    await replaceTexts(p, [{ search: "world", replace: "big\r\nplanet" }], false);

    const xml = await readRawDocXml(p);
    expect(rawCRCount(xml)).toBe(0);
    expect(xml).toContain("<w:br/>");
  });
});

describe("M2: untracked editTableParagraphs renders \\n as a <w:br/> soft break", () => {
  it("untracked produces one <w:p> with a <w:br/> between lines (no literal LF)", async () => {
    const p = await createTmpDoc("seed");
    await insertTable(p, -1, 1, 1);
    await editTableParagraphs(
      p,
      [{ blockIndex: 1, rowIndex: 0, colIndex: 0, paragraphIndex: 0, newText: "AA\nBB" }],
      false,
    );

    const xml = await readRawDocXml(p);
    expect(xml).toContain("<w:br/>");
    expect(xml).not.toMatch(/<w:t[^>]*>[^<]*\n[^<]*<\/w:t>/);
    // Still a single cell paragraph (a soft break, not a paragraph split).
    expect(await pyFirstCellText(p)).toBe("AA\nBB");
    expect(xmlIsWellFormed(p)).toBe(true);
  });

  it("untracked matches the tracked path (both emit a <w:br/>)", async () => {
    const pUn = await createTmpDoc("seed");
    await insertTable(pUn, -1, 1, 1);
    await editTableParagraphs(
      pUn,
      [{ blockIndex: 1, rowIndex: 0, colIndex: 0, paragraphIndex: 0, newText: "AA\nBB" }],
      false,
    );
    const pTr = await createTmpDoc("seed");
    await insertTable(pTr, -1, 1, 1);
    await editTableParagraphs(
      pTr,
      [{ blockIndex: 1, rowIndex: 0, colIndex: 0, paragraphIndex: 0, newText: "AA\nBB" }],
      true,
    );
    expect((await readRawDocXml(pUn)).includes("<w:br/>")).toBe(true);
    expect((await readRawDocXml(pTr)).includes("<w:br/>")).toBe(true);
  });

  it("normalizes CRLF in untracked editTableParagraphs (zero raw \\r)", async () => {
    const p = await createTmpDoc("seed");
    await insertTable(p, -1, 1, 1);
    await editTableParagraphs(
      p,
      [{ blockIndex: 1, rowIndex: 0, colIndex: 0, paragraphIndex: 0, newText: "AA\r\nBB" }],
      false,
    );
    const xml = await readRawDocXml(p);
    expect(rawCRCount(xml)).toBe(0);
    expect(xml).toContain("<w:br/>");
  });

  it("the edit_table_paragraphs tool docs mention \\n behavior", async () => {
    const src = await fs.readFile(new URL("../index.ts", import.meta.url), "utf8");
    const idx = src.indexOf('"edit_table_paragraphs"');
    expect(idx).toBeGreaterThan(-1);
    const region = src.slice(idx, idx + 1600);
    expect(region).toMatch(/\\n/);
    expect(region).toMatch(/soft|break/i);
  });
});

describe("M3/M7: non-integer / wrong-type indices are rejected with INDEX_OUT_OF_RANGE", () => {
  const BAD: [string, unknown][] = [
    ["fractional 1.5", 1.5],
    ["NaN", NaN],
    ["null", null],
    ["string '1'", "1"],
  ];

  async function expectIndexError(fn: () => Promise<unknown>): Promise<void> {
    let err: unknown;
    try {
      await fn();
    } catch (e) {
      err = e;
    }
    expect(err, "expected a throw").toBeInstanceOf(EngineError);
    expect((err as EngineError).code).toBe(ErrorCode.INDEX_OUT_OF_RANGE);
  }

  describe("editTableCells", () => {
    for (const [label, val] of BAD) {
      it(`rejects blockIndex=${label}`, async () => {
        const p = await createDocParaThenTable();
        await expectIndexError(() =>
          editTableCells(p, [{ blockIndex: val as number, rowIndex: 0, colIndex: 0, newText: "x" }], false),
        );
      });
      it(`rejects rowIndex=${label}`, async () => {
        const p = await createDocParaThenTable();
        await expectIndexError(() =>
          editTableCells(p, [{ blockIndex: 1, rowIndex: val as number, colIndex: 0, newText: "x" }], false),
        );
      });
      it(`rejects colIndex=${label}`, async () => {
        const p = await createDocParaThenTable();
        await expectIndexError(() =>
          editTableCells(p, [{ blockIndex: 1, rowIndex: 0, colIndex: val as number, newText: "x" }], false),
        );
      });
    }

    it("a wrong-type string blockIndex does NOT silently mutate the cell", async () => {
      const p = await createDocParaThenTable();
      // The table lives at block index 1; '1' must be rejected, not coerced.
      await expect(
        editTableCells(p, [{ blockIndex: "1" as unknown as number, rowIndex: "0" as unknown as number, colIndex: "0" as unknown as number, newText: "WROTE-VIA-STRING-INDEX" }], false),
      ).rejects.toBeInstanceOf(EngineError);
      // The on-disk cell text is untouched.
      expect(await pyFirstCellText(p)).toBe("ORIG");
      const xml = await readRawDocXml(p);
      expect(xml).not.toContain("WROTE-VIA-STRING-INDEX");
    });
  });

  describe("readTableStructure", () => {
    for (const [label, val] of BAD) {
      it(`rejects blockIndex=${label}`, async () => {
        const p = await createDocParaThenTable();
        await expectIndexError(() => readTableStructure(p, val as number));
      });
    }
  });

  describe("readTableCell", () => {
    for (const [label, val] of BAD) {
      it(`rejects rowIndex=${label}`, async () => {
        const p = await createDocParaThenTable();
        await expectIndexError(() => readTableCell(p, 1, val as number, 0));
      });
      it(`rejects colIndex=${label}`, async () => {
        const p = await createDocParaThenTable();
        await expectIndexError(() => readTableCell(p, 1, 0, val as number));
      });
    }
  });

  describe("getParagraphFormat", () => {
    for (const [label, val] of BAD) {
      it(`rejects paragraphIndex=${label}`, async () => {
        const p = await createTmpDoc("only paragraph");
        await expectIndexError(() => getParagraphFormat(p, val as number));
      });
    }
  });

  describe("insertParagraphs copy_format_from", () => {
    for (const [label, val] of BAD) {
      it(`rejects copyFormatFrom=${label}`, async () => {
        const p = await createTmpDoc("base paragraph");
        await expectIndexError(() =>
          insertParagraphs(p, [{ text: "new", position: 0, copyFormatFrom: val as number }], false),
        );
      });
    }
  });
});

// =========================================================================
// H3 — reject_all_changes must REMOVE a tracked paragraph-insertion, not
//      leave an empty <w:p> shell (body + table); guard the last paragraph.
// =========================================================================

describe("H3: reject removes a tracked paragraph insertion (no residual empty <w:p>)", () => {
  it("body: a single tracked insert + reject restores the pre-insert block count", async () => {
    const p = await createTmpDoc("Existing only");
    const before = (await getDocumentInfoStructured(p)).totalBlocks;
    expect(before).toBe(1);

    await insertParagraphs(p, [{ text: "Should disappear", position: -1 }], true);
    await rejectAllChanges(p);

    // Structured block count is back to the pre-insert value (NOT 2).
    expect((await getDocumentInfoStructured(p)).totalBlocks).toBe(before);
    // python-docx sees exactly one paragraph, with no trailing empty.
    expect(await pyBodyParagraphs(p)).toEqual(["Existing only"]);
  });

  it("body: two mid-document tracked inserts + reject leave no interleaved empties", async () => {
    const p = await createTmpDoc("BASE0\nBASE1\nBASE2");
    expect((await getDocumentInfoStructured(p)).totalBlocks).toBe(3);

    await insertParagraphs(
      p,
      [
        { text: "NEW-A", position: 1 },
        { text: "NEW-B", position: 2 },
      ],
      true,
      "R",
    );
    await rejectAllChanges(p);

    expect((await getDocumentInfoStructured(p)).totalBlocks).toBe(3);
    expect(await pyBodyParagraphs(p)).toEqual(["BASE0", "BASE1", "BASE2"]);
  });

  it("table: a tracked cell insert + reject restores the original cell paragraphs", async () => {
    const p = await createTmpDoc("lead");
    await insertTable(p, -1, 2, 2, [["c1", "c2"], ["c3", "c4"]]);

    // Body layout after append: block 0 = "lead" paragraph, block 1 = table.
    const tableBlock = 1;
    const cellBefore = await readTableCellStructured(p, tableBlock, 1, 1);
    expect(cellBefore.paragraphs.map((x) => x.text)).toEqual(["c4"]);

    await insertTableParagraphs(
      p,
      [{ blockIndex: tableBlock, rowIndex: 1, colIndex: 1, position: 9999, text: "T1\nT2" }],
      true,
      "R",
    );
    await rejectAllChanges(p);

    const cellAfter = await readTableCellStructured(p, tableBlock, 1, 1);
    expect(cellAfter.paragraphs.map((x) => x.text)).toEqual(["c4"]);
  });

  it("CONTROL: a single-line tracked text EDIT + reject still reverts cleanly", async () => {
    const p = await createTmpDoc("HELLO");
    await editParagraphs(p, [{ paragraphIndex: 0, newText: "WORLD" }], true);
    await rejectAllChanges(p);

    expect((await getDocumentInfoStructured(p)).totalBlocks).toBe(1);
    expect(await pyBodyParagraphs(p)).toEqual(["HELLO"]);
  });

  it("GUARD: rejecting a tracked insert that is the cell's sole paragraph leaves a blank <w:p>", async () => {
    // A 1×1 table whose only cell paragraph is itself a tracked insertion. The
    // body's lead paragraph keeps a valid block 0; rejecting must not empty the
    // <w:tc> of its required terminating <w:p>.
    const p = tmpDocxPath();
    trackTmpFile(p);
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:p><w:r><w:t>lead</w:t></w:r></w:p>
<w:tbl>
  <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
  <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
  <w:tr>
    <w:tc>
      <w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>
      <w:p>
        <w:pPr><w:rPr><w:ins w:id="900" w:author="R" w:date="2024-01-01T00:00:00Z"/></w:rPr></w:pPr>
        <w:ins w:id="901" w:author="R" w:date="2024-01-01T00:00:00Z">
          <w:r><w:t xml:space="preserve">sole inserted</w:t></w:r>
        </w:ins>
      </w:p>
    </w:tc>
  </w:tr>
</w:tbl>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
</w:body>
</w:document>`;
    await writeMinimalDocx(p, documentXml);

    await rejectAllChanges(p);

    // The <w:tc> must still END with a <w:p> (OOXML CT_Tc requirement); python-docx
    // opens it and the cell reads as empty (sole inserted text gone, blank para kept).
    expect(pythonDocxOpens(p)).toBe(true);
    expect(xmlIsWellFormed(p)).toBe(true);
    expect(await pyFirstCellText(p)).toBe("");
    // No revision markers remain.
    const xml = await readRawDocXml(p);
    expect(xml).not.toContain("w:ins");
  });
});

// =========================================================================
// M6 — untracked multi-line split must carry the original w14:paraId onto
//      the FIRST produced line so a captured/seeded anchor stays resolvable.
// =========================================================================

/** All w14:paraId values present in word/document.xml, in document order. */
function paraIdsInXml(xml: string): string[] {
  return [...xml.matchAll(/w14:paraId="([0-9A-Fa-f]+)"/g)].map((m) => m[1]);
}

describe("M6: untracked multi-line split keeps the captured anchor on the first line", () => {
  it("first resulting line keeps the anchor; re-edit by it succeeds; no duplicate paraIds", async () => {
    const p = await createTmpDoc("One\nTwo\nThree");
    await ensureAnchors(p);
    const seeded = await ensureAnchorsStructured(p);
    const two = seeded.blocks.find((b) => b.textPreview === "Two");
    expect(two?.anchor).toBeTruthy();
    const anchorA = two!.anchor as string;

    // Untracked multi-line edit on the anchored paragraph.
    await editParagraphs(p, [{ anchor: anchorA, newText: "a\nb\nc" }], false);

    const xml = await readRawDocXml(p);
    const ids = paraIdsInXml(xml);
    // The captured anchor survives on the first produced line.
    expect(ids).toContain(anchorA);
    // No duplicate paraIds were introduced by the split.
    expect(new Set(ids).size).toBe(ids.length);

    // Re-editing by the captured anchor must resolve (not ANCHOR_NOT_FOUND).
    await editParagraphs(p, [{ anchor: anchorA, newText: "again" }], false);
    const after = await readRawDocXml(p);
    expect(after).toContain("again");
    expect(paraIdsInXml(after)).toContain(anchorA);
  });
});

// =========================================================================
// M4 — the <json> structured block must be delimiter-safe: document text
//      containing the literal "</json>" must not break extraction.
// =========================================================================

/** Extract the payload between the LAST "<json>" and the LAST "</json>". */
function lastJsonPayload(output: string): string {
  const open = output.lastIndexOf("<json>");
  const close = output.lastIndexOf("</json>");
  expect(open).toBeGreaterThanOrEqual(0);
  expect(close).toBeGreaterThan(open);
  return output.slice(open + "<json>".length, close);
}

describe("M4: <json> structured block is delimiter-safe against document text", () => {
  it("searchText payload contains no literal </json> and round-trips the sentinel text", async () => {
    // Build a doc whose paragraph AND cell text contain the literal "</json>".
    const p = tmpDocxPath();
    trackTmpFile(p);
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:p><w:r><w:t xml:space="preserve">para with &lt;/json&gt; needle</w:t></w:r></w:p>
<w:tbl>
  <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
  <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>
      <w:p><w:r><w:t xml:space="preserve">cell with &lt;/json&gt; needle</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
</w:body>
</w:document>`;
    await writeMinimalDocx(p, documentXml);

    const out = await searchText(p, "needle");
    const payload = lastJsonPayload(out);

    // The payload itself must not contain the delimiter literal.
    expect(payload).not.toContain("</json>");
    // It parses, and the decoded values round-trip the literal "</json>" text.
    const parsed = JSON.parse(payload);
    expect(parsed.totalMatches).toBeGreaterThanOrEqual(2);
    const contexts = parsed.matches.map((m: { context: string }) => m.context).join(" ");
    expect(contexts).toContain("</json>");
  });
});

// =========================================================================
// L2 — zero-result read tools must still emit a parseable <json> block.
// =========================================================================

describe("L2: zero-result read tools emit an empty <json> block", () => {
  it("searchText with no match still emits parseable JSON with an empty array", async () => {
    const p = await createTmpDoc("nothing to find here");
    const out = await searchText(p, "zzz-no-match-zzz");
    expect(out).toContain("No matches found");
    const parsed = JSON.parse(lastJsonPayload(out));
    expect(parsed.totalMatches).toBe(0);
    expect(parsed.matches).toEqual([]);
  });

  it("listImages with no images still emits parseable JSON with an empty array", async () => {
    const p = await createTmpDoc("no pictures");
    const out = await listImages(p);
    expect(out).toContain("No images found");
    const parsed = JSON.parse(lastJsonPayload(out));
    expect(parsed.totalImages).toBe(0);
    expect(parsed.images).toEqual([]);
  });

  it("readComments with no comments still emits parseable JSON with an empty array", async () => {
    const p = await createTmpDoc("uncommented");
    const out = await readComments(p);
    expect(out).toContain("No comments found");
    const parsed = JSON.parse(lastJsonPayload(out));
    expect(parsed.totalComments).toBe(0);
    expect(parsed.comments).toEqual([]);
  });
});

// =========================================================================
// L3 — get_paragraph_format JSON must include resolved defaults (alignment /
//      style / headingLevel) for a plain paragraph so it matches the text.
// =========================================================================

describe("L3: get_paragraph_format JSON carries explicit defaults for plain paragraphs", () => {
  it("structured result has explicit alignment plus style/headingLevel keys", async () => {
    const p = await createTmpDoc("plain unstyled paragraph");
    const r = await getParagraphFormatStructured(p, 0);

    // Explicit, not undefined — the JSON must agree with the "alignment left
    // (default)" / "style default" the human text prints.
    expect(r.alignment).toBe("left");
    expect("style" in r).toBe(true);
    expect("headingLevel" in r).toBe(true);
    expect(r.style).toBeNull();
    expect(r.headingLevel).toBeNull();

    // And the serialized <json> the tool emits keeps those keys (not dropped).
    const out = await getParagraphFormat(p, 0);
    const parsed = JSON.parse(lastJsonPayload(out));
    expect(parsed.alignment).toBe("left");
    expect(Object.prototype.hasOwnProperty.call(parsed, "style")).toBe(true);
    expect(Object.prototype.hasOwnProperty.call(parsed, "headingLevel")).toBe(true);
  });
});

// M7 (residual) — the integer guard missed two INSERT sites: spliceNewParagraph
// (insert_paragraphs / insert_table position) and resolveCellInsertOpts
// (insert_table_paragraphs copy_format_from). A fractional/NaN value fell through
// the `< 0 || >= len` bounds check and threw a raw TypeError surfaced as
// INTERNAL_ERROR. Found by the re-QA regression sweep.
describe("M7 residual: non-integer insert position / copy_format_from rejected cleanly", () => {
  async function expectIndexError(fn: () => Promise<unknown>): Promise<void> {
    let err: unknown;
    try {
      await fn();
    } catch (e) {
      err = e;
    }
    expect(err, "expected a throw").toBeInstanceOf(EngineError);
    expect((err as EngineError).code).toBe(ErrorCode.INDEX_OUT_OF_RANGE);
  }

  for (const [label, val] of [["fractional 1.5", 1.5], ["NaN", NaN]] as [string, number][]) {
    it(`insert_paragraphs rejects position=${label}`, async () => {
      const p = await createTmpDoc("a");
      await expectIndexError(() => insertParagraphs(p, [{ text: "x", position: val }], false));
    });
    it(`insert_table rejects position=${label}`, async () => {
      const p = await createTmpDoc("a");
      await expectIndexError(() => insertTable(p, val, 1, 1, [["c"]]));
    });
    it(`insert_table_paragraphs rejects copy_format_from=${label}`, async () => {
      const p = await createDocParaThenTable(); // table at block 1, cell has 1 paragraph
      await expectIndexError(() =>
        insertTableParagraphs(
          p,
          [{ blockIndex: 1, rowIndex: 0, colIndex: 0, position: -1, text: "x", copyFormatFrom: val }],
          false,
        ),
      );
    });
  }

  it("still appends with an out-of-range INTEGER position (-1 and large), unchanged", async () => {
    const p = await createTmpDoc("base");
    await insertParagraphs(p, [{ text: "appended-neg1", position: -1 }], false);
    await insertParagraphs(p, [{ text: "appended-large", position: 999 }], false);
    const xml = await readRawDocXml(p);
    expect(xml).toContain("appended-neg1");
    expect(xml).toContain("appended-large");
  });
});

// =========================================================================
// MEDIUM (three-stage review): editing an SDT-inner paragraph by its unified
// block index must NOT seed a w14:paraId onto the SDT-contained <w:p>. Anchor
// scope (v1) is direct-body paragraphs only; the SDT <w:p> is unreachable by
// anchor (buildAnchorIndex skips it), so a seeded id there is scope-violating
// AND useless. Normal direct-body edits must STILL auto-seed (no over-correct).
//
// Fixture: createDocWithSdt → block 0 = "Normal paragraph" (direct body),
//          block 1 = the SDT-inner paragraph.
// =========================================================================
describe("SDT-inner paragraph edits do not seed a scope-violating anchor (v1)", () => {
  /** The portion of document.xml inside the top-level <w:sdt>. */
  function sdtPart(xml: string): string {
    return xml.slice(xml.indexOf("<w:sdt"));
  }

  it("edit_paragraphs at an SDT-inner block index leaves the SDT <w:p> WITHOUT a paraId", async () => {
    const p = await createDocWithSdt("inside the content control");
    // Block 1 is the SDT-inner paragraph (block 0 is the direct-body paragraph).
    await editParagraphs(p, [{ paragraphIndex: 1, newText: "edited inside sdt" }], false);
    const xml = await readRawDocXml(p);
    expect(xml).toContain("edited inside sdt"); // the edit DID land
    expect(sdtPart(xml)).not.toContain("w14:paraId"); // …without seeding the SDT <w:p>
  });

  it("set_headings at an SDT-inner block index leaves the SDT <w:p> WITHOUT a paraId", async () => {
    const p = await createDocWithSdt("inside the content control");
    await setHeadings(p, [{ paragraphIndex: 1, level: 1 }]);
    const xml = await readRawDocXml(p);
    expect(sdtPart(xml)).not.toContain("w14:paraId");
  });

  it("set_paragraph_formats at an SDT-inner block index leaves the SDT <w:p> WITHOUT a paraId", async () => {
    const p = await createDocWithSdt("inside the content control");
    await setParagraphFormats(p, [{ indices: [1], format: { alignment: "center" } }]);
    const xml = await readRawDocXml(p);
    expect(sdtPart(xml)).not.toContain("w14:paraId");
  });

  it("a NORMAL direct-body edit STILL seeds an anchor (guard against over-correction)", async () => {
    const p = await createDocWithSdt("inside the content control");
    // Block 0 is the direct-body "Normal paragraph".
    await editParagraphs(p, [{ paragraphIndex: 0, newText: "Normal paragraph edited" }], false);
    const xml = await readRawDocXml(p);
    // The direct-body paragraph (everything before the SDT) MUST carry a paraId…
    const beforeSdt = xml.slice(0, xml.indexOf("<w:sdt"));
    expect(beforeSdt).toContain("w14:paraId");
    expect(xml).toContain("xmlns:w14");
    // …and the SDT paragraph still must NOT.
    expect(sdtPart(xml)).not.toContain("w14:paraId");
  });

  it("set_headings / set_paragraph_formats on a direct-body block still seed (over-correction guard)", async () => {
    const ph = await createDocWithSdt("inside the content control");
    await setHeadings(ph, [{ paragraphIndex: 0, level: 1 }]);
    const xmlH = await readRawDocXml(ph);
    expect(xmlH.slice(0, xmlH.indexOf("<w:sdt"))).toContain("w14:paraId");

    const pf = await createDocWithSdt("inside the content control");
    await setParagraphFormats(pf, [{ indices: [0], format: { alignment: "center" } }]);
    const xmlF = await readRawDocXml(pf);
    expect(xmlF.slice(0, xmlF.indexOf("<w:sdt"))).toContain("w14:paraId");
  });
});

// =========================================================================
// LOW (three-stage review): a malformed numeric character reference in the
// untrusted document.xml (code point > U+10FFFF, or a surrogate) must NOT
// crash the read path with a RangeError surfaced as [INTERNAL_ERROR]. The
// malformed ref is left as literal text so the document still parses; a VALID
// numeric ref still decodes.
// =========================================================================
describe("malformed numeric character references read without throwing", () => {
  /** A minimal doc whose body text is exactly `bodyText` (inserted verbatim). */
  async function docWithBodyText(bodyText: string): Promise<string> {
    const p = tmpDocxPath();
    trackTmpFile(p);
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:r><w:t xml:space="preserve">${bodyText}</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`;
    await writeMinimalDocx(p, documentXml);
    return p;
  }

  it("a > U+10FFFF hex ref and a huge decimal ref read without throwing (left as literal)", async () => {
    const p = await docWithBodyText("before &#x110000; mid &#99999999999999; after");
    // The read/parse path must not throw (previously: RangeError → INTERNAL_ERROR).
    await expect(getDocumentInfoStructured(p)).resolves.toBeDefined();
    // The paragraph text round-trips with the malformed refs left as literal text
    // (lenient: not decoded, not dropped). Asserted on the structured fullText.
    const r = await searchTextStructured(p, "before");
    expect(r.matches).toHaveLength(1);
    const text = r.matches[0].fullText;
    expect(text).toContain("before");
    expect(text).toContain("after");
    expect(text).toContain("&#x110000;");
    expect(text).toContain("&#99999999999999;");
  });

  it("a surrogate-range numeric ref (&#xD800;) reads without throwing (left as literal)", async () => {
    const p = await docWithBodyText("xsurr &#xD800; ysurr");
    await expect(getDocumentInfoStructured(p)).resolves.toBeDefined();
    const r = await searchTextStructured(p, "xsurr");
    expect(r.matches[0].fullText).toContain("&#xD800;");
  });

  it("a VALID numeric ref still decodes (&#65; and &#x41; → A)", async () => {
    const p = await docWithBodyText("decval=&#65; hexval=&#x41;");
    const r = await searchTextStructured(p, "decval");
    expect(r.matches).toHaveLength(1);
    expect(r.matches[0].fullText).toContain("decval=A");
    expect(r.matches[0].fullText).toContain("hexval=A");
  });
});

// =========================================================================
// LOW (three-stage review): sanitizeXmlText must also drop U+FFFE/U+FFFF and
// unpaired UTF-16 surrogates (all XML-1.0-illegal), while keeping valid
// surrogate PAIRS (astral chars / emoji) intact. Verified end-to-end: the
// written document.xml is well-formed (xmllint) and python-docx opens it.
// =========================================================================
describe("sanitizer strips noncharacters and lone surrogates, keeps astral pairs", () => {
  it("a valid emoji (astral surrogate pair) SURVIVES intact and the file is well-formed", async () => {
    const p = tmpDocxPath();
    trackTmpFile(p);
    await createDocument(p, "smile \u{1F600} end");
    const xml = await readRawDocXml(p);
    expect(/\u{1F600}/u.test(xml)).toBe(true); // emoji preserved
    expect(xmlIsWellFormed(p)).toBe(true);
    expect(pythonDocxOpens(p)).toBe(true);
  });

  it("lone surrogates (\\uD800, \\uDFFF) and U+FFFE/U+FFFF are stripped; file stays well-formed", async () => {
    const p = tmpDocxPath();
    trackTmpFile(p);
    // A\uD800B\uDFFFC￾D￿E — only the ASCII letters may survive.
    await createDocument(p, "A\uD800B\uDFFFC￾D￿E");
    const xml = await readRawDocXml(p);
    const m = xml.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/);
    const inner = m ? m[1] : "";
    expect(inner).toBe("ABCDE"); // every illegal char removed, letters intact & in order
    expect(xmlIsWellFormed(p)).toBe(true);
    expect(pythonDocxOpens(p)).toBe(true);
  });

  it("editParagraphs with the same illegal mix also yields a well-formed file", async () => {
    const p = await createTmpDoc("seed");
    await editParagraphs(p, [{ paragraphIndex: 0, newText: "X\uD800Y\uDFFFZ￿W" }], false);
    expect(xmlIsWellFormed(p)).toBe(true);
    expect(pythonDocxOpens(p)).toBe(true);
    const xml = await readRawDocXml(p);
    const m = xml.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/);
    expect(m ? m[1] : "").toBe("XYZW");
  });

  it("a valid pair followed by a lone surrogate keeps the pair, drops the lone one", async () => {
    const p = tmpDocxPath();
    trackTmpFile(p);
    // emoji (D83D DE00) then a lone high D800 then 'Z'.
    await createDocument(p, "\u{1F600}\uD800Z");
    const xml = await readRawDocXml(p);
    const m = xml.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/);
    const inner = m ? m[1] : "";
    expect(/\u{1F600}/u.test(inner)).toBe(true);
    expect(inner).toContain("Z");
    expect(/[\uD800-\uDFFF]/.test(inner.replace(/\u{1F600}/u, ""))).toBe(false);
    expect(xmlIsWellFormed(p)).toBe(true);
  });
});

// =========================================================================
// CODEX FINAL-GATE FINDINGS (F1–F4)
// =========================================================================

// -------------------------------------------------------------------------
// F1 [HIGH] — user-controlled strings that flow into XML ATTRIBUTE values
// (w:author, w:pStyle/@w:val, font names, colors, comment author) must be
// sanitized too. C2 only covered TEXT nodes; a control char or lone surrogate
// in an attribute value was written raw → malformed document.xml/comments.xml.
// Fix: sanitize STRING attribute values at the setAttr()/el() chokepoints.
// -------------------------------------------------------------------------
describe("F1: XML attribute values are sanitized (control chars + lone surrogates)", () => {
  const NUL = String.fromCharCode(0x00); // C0 NUL — XML-illegal
  const FF = String.fromCharCode(0x0c); // U+000C form feed — XML-illegal
  const HI = "\uD800"; // lone high surrogate — XML-illegal

  it("editParagraphs author with a control char + lone surrogate → well-formed, openable", async () => {
    const p = await createTmpDoc("seed");
    await editParagraphs(
      p,
      [{ paragraphIndex: 0, newText: "world" }],
      true,
      "auth" + NUL + "o" + FF + "r" + HI,
    );
    const xml = await readRawDocXml(p);
    // No raw illegal codepoint survives in the w:author attribute (or anywhere).
    expect(xml).not.toContain(NUL);
    expect(xml).not.toContain(FF);
    expect(/[\uD800-\uDFFF]/.test(xml)).toBe(false);
    // The clean part of the author name is intact.
    expect(xml).toContain('w:author="author"');
    expect(xmlIsWellFormed(p)).toBe(true);
    expect(pythonDocxOpens(p)).toBe(true);
  });

  it("addComment author with a control char + lone surrogate → comments.xml well-formed", async () => {
    const p = await createTmpDoc("anchor here please");
    await addComment(p, "anchor", "a note", "rev" + NUL + "ie" + FF + "wer" + HI);
    const cxml = await readRawCommentsXml(p);
    expect(cxml).not.toContain(NUL);
    expect(cxml).not.toContain(FF);
    expect(/[\uD800-\uDFFF]/.test(cxml)).toBe(false);
    expect(cxml).toContain('w:author="reviewer"');
    // Both document.xml and comments.xml must be well-formed and the file opens.
    expect(xmlIsWellFormed(p)).toBe(true);
    // comments.xml well-formed too (skipped when xmllint absent; the raw
    // no-illegal-char + author assertions above remain load-bearing).
    expectCommentsXmlWellFormed(p);
    expect(pythonDocxOpens(p)).toBe(true);
  });

  it("formatText font_name with a control char + lone surrogate → well-formed, openable", async () => {
    const p = await createTmpDoc("style this word");
    await formatText(p, "word", { fontName: "Ari" + NUL + "a" + FF + "l" + HI });
    const xml = await readRawDocXml(p);
    expect(xml).not.toContain(NUL);
    expect(xml).not.toContain(FF);
    expect(/[\uD800-\uDFFF]/.test(xml)).toBe(false);
    // The rFonts attributes carry the cleaned font name.
    expect(xml).toContain('w:ascii="Arial"');
    expect(xmlIsWellFormed(p)).toBe(true);
    expect(pythonDocxOpens(p)).toBe(true);
  });

  it("a normal author with an emoji (astral surrogate PAIR) survives intact", async () => {
    const p = await createTmpDoc("seed para");
    await editParagraphs(p, [{ paragraphIndex: 0, newText: "edited" }], true, "Rev \u{1F600} iewer");
    const xml = await readRawDocXml(p);
    // The emoji (valid pair) is preserved in the attribute value.
    expect(/\u{1F600}/u.test(xml)).toBe(true);
    expect(xml).toContain("Rev \u{1F600} iewer");
    expect(xmlIsWellFormed(p)).toBe(true);
    expect(pythonDocxOpens(p)).toBe(true);
  });

  it("a normal font name with an emoji survives intact in rFonts", async () => {
    const p = await createTmpDoc("font me");
    await formatText(p, "font", { fontName: "Emoji\u{1F600}Font" });
    const xml = await readRawDocXml(p);
    expect(/\u{1F600}/u.test(xml)).toBe(true);
    expect(xml).toContain('w:ascii="Emoji\u{1F600}Font"');
    expect(xmlIsWellFormed(p)).toBe(true);
    expect(pythonDocxOpens(p)).toBe(true);
  });
});

// -------------------------------------------------------------------------
// F2 [MEDIUM] — an UNTRACKED MULTILINE edit of an SDT-inner paragraph must
// REFUSE with INVALID_LOCATOR (a paragraph inside a w:sdt cannot be split into
// sibling paragraphs; the old code silently degraded to a single <w:p> with
// soft <w:br/> breaks). A single-line untracked SDT-inner edit still works in
// place; a direct-body multiline untracked edit still splits.
// -------------------------------------------------------------------------
describe("F2: untracked multiline edit of an SDT-inner paragraph is refused", () => {
  /** The portion of document.xml inside the top-level <w:sdt>. */
  function sdtPart(xml: string): string {
    return xml.slice(xml.indexOf("<w:sdt"), xml.indexOf("</w:sdt>") + "</w:sdt>".length);
  }

  it("throws INVALID_LOCATOR and leaves the SDT <w:p> unchanged", async () => {
    const p = await createDocWithSdt("inside the content control");
    const before = await readRawDocXml(p);
    const sdtBefore = sdtPart(before);

    let err: unknown;
    try {
      await editParagraphs(p, [{ paragraphIndex: 1, newText: "a\nb\nc" }], false);
    } catch (e) {
      err = e;
    }
    expect(err, "expected a throw").toBeInstanceOf(EngineError);
    expect((err as EngineError).code).toBe(ErrorCode.INVALID_LOCATOR);

    // The SDT-inner paragraph is untouched (the refusal happens before any write).
    const after = await readRawDocXml(p);
    expect(sdtPart(after)).toBe(sdtBefore);
  });

  it("a SINGLE-LINE untracked edit of an SDT-inner paragraph still edits in place", async () => {
    const p = await createDocWithSdt("inside the content control");
    await editParagraphs(p, [{ paragraphIndex: 1, newText: "edited inside sdt" }], false);
    const xml = await readRawDocXml(p);
    expect(xml).toContain("edited inside sdt");
    // Still exactly one <w:p> inside the SDT (no split, no extra paragraphs).
    const pInSdt = (sdtPart(xml).match(/<w:p[ >]/g) ?? []).length;
    expect(pInSdt).toBe(1);
    expect(xmlIsWellFormed(p)).toBe(true);
    expect(pythonDocxOpens(p)).toBe(true);
  });

  it("a direct-body multiline untracked edit STILL splits into separate paragraphs", async () => {
    const p = await createTmpDoc("orig");
    await editParagraphs(p, [{ paragraphIndex: 0, newText: "a\nb\nc" }], false);
    expect(await pyBodyParagraphs(p)).toEqual(["a", "b", "c"]);
    expect(xmlIsWellFormed(p)).toBe(true);
  });

  it("a TRACKED multiline edit of an SDT-inner paragraph is unaffected (soft breaks, no refusal)", async () => {
    const p = await createDocWithSdt("inside the content control");
    await editParagraphs(p, [{ paragraphIndex: 1, newText: "a\nb\nc" }], true);
    const xml = await readRawDocXml(p);
    // Tracked path uses <w:br/> soft breaks inside a <w:ins>, no paragraph split.
    expect(sdtPart(xml)).toContain("<w:br/>");
    expect(xmlIsWellFormed(p)).toBe(true);
    expect(pythonDocxOpens(p)).toBe(true);
  });
});

// -------------------------------------------------------------------------
// F3 [MEDIUM] — CRLF / "\n" handling missed three <w:t> writers:
//   - createDocument(content): a preceding "\r" survived in <w:t>.
//   - addComment / addComments: a "\r" survived in comments.xml.
//   - insertTable cell data: a literal "\n" was written into <w:t>.
// Fix: normalizeNewlines before the "\n" split on each path; insertTable cell
// data additionally renders "\n" as a <w:br/> soft break.
// -------------------------------------------------------------------------
describe("F3: CRLF / newline normalization on create / table / comment write paths", () => {
  it("createDocument with CRLF content → zero raw \\r, two clean paragraphs", async () => {
    const p = tmpDocxPath();
    trackTmpFile(p);
    await createDocument(p, undefined, "p1\r\np2");
    const xml = await readRawDocXml(p);
    expect(rawCRCount(xml)).toBe(0);
    expect(await pyBodyParagraphs(p)).toEqual(["p1", "p2"]);
    expect(xmlIsWellFormed(p)).toBe(true);
    expect(pythonDocxOpens(p)).toBe(true);
  });

  it("addComment with CRLF comment text → zero raw \\r in comments.xml", async () => {
    const p = await createTmpDoc("comment anchor word");
    await addComment(p, "anchor", "a\r\nb", "QA");
    const cxml = await readRawCommentsXml(p);
    expect(rawCRCount(cxml)).toBe(0);
    // Two comment paragraphs (the "\n" became a paragraph split, no stray "\r").
    expect(cxml).toContain("a");
    expect(cxml).toContain("b");
    expectCommentsXmlWellFormed(p);
    expect(pythonDocxOpens(p)).toBe(true);
  });

  it("addComments (batch) with CRLF comment text → zero raw \\r in comments.xml", async () => {
    const p = await createTmpDoc("batch anchor word");
    await addComments(p, [{ anchor_text: "anchor", comment_text: "x\r\ny" }], "QA");
    const cxml = await readRawCommentsXml(p);
    expect(rawCRCount(cxml)).toBe(0);
    expectCommentsXmlWellFormed(p);
  });

  it("insertTable cell data with \\n → a <w:br/> soft break, no literal LF in <w:t>", async () => {
    const p = await createTmpDoc("intro");
    await insertTable(p, -1, 1, 1, [["c1\nc2"]]);
    const xml = await readRawDocXml(p);
    // No <w:t> node contains an embedded newline character.
    expect(xml).not.toMatch(/<w:t[^>]*>[^<]*\n[^<]*<\/w:t>/);
    // The newline rendered as a soft break inside the cell run.
    expect(xml).toContain("<w:br/>");
    // python-docx reads the cell as the two lines joined by a newline.
    expect(await pyFirstCellText(p)).toBe("c1\nc2");
    expect(xmlIsWellFormed(p)).toBe(true);
    expect(pythonDocxOpens(p)).toBe(true);
  });

  it("insertTable cell data with CRLF → zero raw \\r and a <w:br/> soft break", async () => {
    const p = await createTmpDoc("intro");
    await insertTable(p, -1, 1, 1, [["r1\r\nr2"]]);
    const xml = await readRawDocXml(p);
    expect(rawCRCount(xml)).toBe(0);
    expect(xml).toContain("<w:br/>");
    expect(await pyFirstCellText(p)).toBe("r1\nr2");
    expect(xmlIsWellFormed(p)).toBe(true);
  });

  it("insertTable cell data without a newline still writes a single plain run", async () => {
    const p = await createTmpDoc("intro");
    await insertTable(p, -1, 1, 1, [["plain"]]);
    const xml = await readRawDocXml(p);
    expect(xml).toContain("plain");
    expect(await pyFirstCellText(p)).toBe("plain");
    expect(xmlIsWellFormed(p)).toBe(true);
  });
});

// -------------------------------------------------------------------------
// F4 [LOW] — insertTableParagraphs must REJECT a non-integer `position` with
// INDEX_OUT_OF_RANGE (it previously treated 0.5 / NaN as append), matching the
// M7 integer-validation policy. A -1 or out-of-range INTEGER position still
// appends, unchanged.
// -------------------------------------------------------------------------
describe("F4: insertTableParagraphs rejects a non-integer position", () => {
  async function expectIndexError(fn: () => Promise<unknown>): Promise<void> {
    let err: unknown;
    try {
      await fn();
    } catch (e) {
      err = e;
    }
    expect(err, "expected a throw").toBeInstanceOf(EngineError);
    expect((err as EngineError).code).toBe(ErrorCode.INDEX_OUT_OF_RANGE);
  }

  for (const [label, val] of [["fractional 0.5", 0.5], ["NaN", NaN]] as [string, number][]) {
    it(`rejects position=${label}`, async () => {
      const p = await createDocParaThenTable(); // table at block 1, cell has 1 paragraph
      await expectIndexError(() =>
        insertTableParagraphs(
          p,
          [{ blockIndex: 1, rowIndex: 0, colIndex: 0, position: val, text: "frac" }],
          false,
        ),
      );
      // The cell was not mutated by the rejected insert.
      expect(await pyFirstCellText(p)).toBe("ORIG");
    });
  }

  it("a -1 position still APPENDS (unchanged)", async () => {
    const p = await createDocParaThenTable();
    await insertTableParagraphs(
      p,
      [{ blockIndex: 1, rowIndex: 0, colIndex: 0, position: -1, text: "appended" }],
      false,
    );
    expect(await pyFirstCellText(p)).toBe("ORIG\nappended");
    expect(xmlIsWellFormed(p)).toBe(true);
  });

  it("a large out-of-range INTEGER position still APPENDS (unchanged)", async () => {
    const p = await createDocParaThenTable();
    await insertTableParagraphs(
      p,
      [{ blockIndex: 1, rowIndex: 0, colIndex: 0, position: 9999, text: "appended-large" }],
      false,
    );
    expect(await pyFirstCellText(p)).toBe("ORIG\nappended-large");
    expect(xmlIsWellFormed(p)).toBe(true);
  });
});

// F3 (Codex re-review): createDocument TITLE path was the one <w:t> writer that
// still didn't normalize CRLF, so a "\r\n" in the title left a raw CR in <w:t>.
describe("F3 (Codex re-review): createDocument title normalizes CRLF", () => {
  it("a CRLF title leaves no raw carriage return and renders a newline as a soft break", async () => {
    const p = await createTmpDoc("body", "Title\r\nLine2");
    const xml = await readRawDocXml(p);
    expect(xml).not.toMatch(/\r/);
    expect(xml).toContain("<w:br");
    expect(xmlIsWellFormed(p)).toBe(true);
  });
  it("a single-line title is unchanged and well-formed", async () => {
    const p = await createTmpDoc("body", "Single Title");
    const xml = await readRawDocXml(p);
    expect(xml).toContain("Single Title");
    expect(xml).not.toContain("<w:br");
    expect(xmlIsWellFormed(p)).toBe(true);
  });
});

// =========================================================================
// Round-4 QA findings: input validation (N2/N9, N5, N10, N11) + content
// preservation (N4). Each finding had a confirmed repro in the QA sweep.
// =========================================================================

// -------------------------------------------------------------------------
// N2 / N9 — insertTable must bound rows/cols so a single bad call cannot OOM
// the long-lived server, and rows<1/cols<1 must NOT emit an empty invalid
// <w:tbl> with a misleading "0x0" success.
// -------------------------------------------------------------------------
describe("N2/N9: insertTable bounds rows/cols (no OOM, no empty invalid table)", () => {
  const badDims: Array<[string, number, number]> = [
    ["rows = 0", 0, 2],
    ["cols = 0", 2, 0],
    ["rows = -3", -3, 2],
    ["cols = -3", 2, -3],
    ["rows = 1.5 (non-integer)", 1.5, 2],
    ["cols = 1.5 (non-integer)", 2, 1.5],
    ["rows = NaN", NaN, 2],
    ["cols = NaN", 2, NaN],
    ["rows over cap (3000)", 3000, 2],
    ["cols over cap (3000)", 2, 3000],
    ["product over cap (200x200=40000)", 200, 200],
  ];

  for (const [label, rows, cols] of badDims) {
    it(`rejects ${label} with INVALID_PARAMETER quickly (no OOM/hang)`, async () => {
      const p = await createTmpDoc("body");
      const start = Date.now();
      let err: unknown;
      try {
        await insertTable(p, -1, rows, cols);
      } catch (e) {
        err = e;
      }
      expect(err).toBeInstanceOf(EngineError);
      expect((err as EngineError).code).toBe(ErrorCode.INVALID_PARAMETER);
      // Must fail fast — a real OOM/giant-build would take seconds.
      expect(Date.now() - start).toBeLessThan(2000);
    });
  }

  it("a normal small 2x3 table still inserts and round-trips", async () => {
    const p = await createTmpDoc("body");
    const msg = await insertTable(p, -1, 2, 3, [["a", "b", "c"], ["d", "e", "f"]]);
    expect(msg).toContain("2x3");
    const xml = await readRawDocXml(p);
    expect(xml).toContain("<w:tbl>");
    // exactly 3 gridCols and 2 rows
    expect((xml.match(/<w:gridCol/g) ?? []).length).toBe(3);
    expect((xml.match(/<w:tr>/g) ?? []).length).toBe(2);
    expect(xmlIsWellFormed(p)).toBe(true);
    expect(pythonDocxOpens(p)).toBe(true);
  });

  it("never emits a 0x0 empty <w:tbl> nor a 0/negative success message", async () => {
    const p = await createTmpDoc("body");
    await expect(insertTable(p, -1, 0, 0)).rejects.toThrow();
    const xml = await readRawDocXml(p);
    expect(xml).not.toContain("<w:tbl>");
  });
});

// -------------------------------------------------------------------------
// N5 — setHeadings must reject non-integer levels (1.5 → "Heading1.5" style +
// outlineLvl "0.5", both schema-invalid; python-docx then raises on .style).
// -------------------------------------------------------------------------
describe("N5: setHeadings rejects non-integer level; integer level round-trips", () => {
  for (const lvl of [1.5, NaN]) {
    it(`level=${lvl} → INVALID_PARAMETER`, async () => {
      const p = await createTmpDoc("Heading me");
      let err: unknown;
      try {
        await setHeadings(p, [{ paragraphIndex: 0, level: lvl }]);
      } catch (e) {
        err = e;
      }
      expect(err).toBeInstanceOf(EngineError);
      expect((err as EngineError).code).toBe(ErrorCode.INVALID_PARAMETER);
      // No invalid style leaked into the file.
      const xml = await readRawDocXml(p);
      expect(xml).not.toMatch(/Heading\d+\.\d+/);
      expect(xml).not.toMatch(/w:outlineLvl w:val="\d+\.\d+"/);
    });
  }

  it("an integer level=2 emits Heading2 / outlineLvl 1 and round-trips", async () => {
    const p = await createTmpDoc("Heading me");
    await setHeadings(p, [{ paragraphIndex: 0, level: 2 }]);
    const xml = await readRawDocXml(p);
    expect(xml).toContain('<w:pStyle w:val="Heading2"');
    expect(xml).toContain('<w:outlineLvl w:val="1"');
    expect(xmlIsWellFormed(p)).toBe(true);
    expect(pythonDocxOpens(p)).toBe(true);
  });
});

// -------------------------------------------------------------------------
// N10 — setPageLayout must reject negative / absurd page geometry (OOXML w:w /
// w:h are unsigned, ~1..31680 twips).
// -------------------------------------------------------------------------
describe("N10: setPageLayout rejects negative/absurd geometry; A4 round-trips", () => {
  it("a negative width → INVALID_PARAMETER (no negative w:w written)", async () => {
    const p = await createTmpDoc("body");
    let err: unknown;
    try {
      await setPageLayout(p, { widthMm: -50 });
    } catch (e) {
      err = e;
    }
    expect(err).toBeInstanceOf(EngineError);
    expect((err as EngineError).code).toBe(ErrorCode.INVALID_PARAMETER);
    const xml = await readRawDocXml(p);
    expect(xml).not.toMatch(/w:w="-\d+"/);
  });

  it("an absurd width (1e9 mm) → INVALID_PARAMETER", async () => {
    const p = await createTmpDoc("body");
    let err: unknown;
    try {
      await setPageLayout(p, { widthMm: 1e9 });
    } catch (e) {
      err = e;
    }
    expect(err).toBeInstanceOf(EngineError);
    expect((err as EngineError).code).toBe(ErrorCode.INVALID_PARAMETER);
  });

  it("a negative margin → INVALID_PARAMETER", async () => {
    const p = await createTmpDoc("body");
    let err: unknown;
    try {
      await setPageLayout(p, { topMm: -10 });
    } catch (e) {
      err = e;
    }
    expect(err).toBeInstanceOf(EngineError);
    expect((err as EngineError).code).toBe(ErrorCode.INVALID_PARAMETER);
  });

  it("a normal A4 portrait layout still applies and round-trips", async () => {
    const p = await createTmpDoc("body");
    const msg = await setPageLayout(p, { widthMm: 210, heightMm: 297, topMm: 25.4 });
    expect(msg).toContain("page size");
    const xml = await readRawDocXml(p);
    expect(xml).toMatch(/<w:pgSz[^>]*w:w="11906"/);
    expect(xmlIsWellFormed(p)).toBe(true);
    expect(pythonDocxOpens(p)).toBe(true);
  });
});

// -------------------------------------------------------------------------
// N11 — formatText font_size must reject negative/absurd (ST_HpsMeasure is
// unsigned half-points); a normal size applies <w:sz w:val="24"/>.
// -------------------------------------------------------------------------
describe("N11: formatText font_size rejects negative/absurd; 12 → sz 24", () => {
  it("fontSize=-8 → INVALID_PARAMETER (no negative w:sz written)", async () => {
    const p = await createTmpDoc("color me");
    let err: unknown;
    try {
      await formatText(p, "color me", { fontSize: -8 });
    } catch (e) {
      err = e;
    }
    expect(err).toBeInstanceOf(EngineError);
    expect((err as EngineError).code).toBe(ErrorCode.INVALID_PARAMETER);
    const xml = await readRawDocXml(p);
    expect(xml).not.toMatch(/w:sz w:val="-\d+"/);
  });

  it("fontSize=1e9 → INVALID_PARAMETER", async () => {
    const p = await createTmpDoc("color me");
    let err: unknown;
    try {
      await formatText(p, "color me", { fontSize: 1e9 });
    } catch (e) {
      err = e;
    }
    expect(err).toBeInstanceOf(EngineError);
    expect((err as EngineError).code).toBe(ErrorCode.INVALID_PARAMETER);
  });

  it("fontSize=12 applies <w:sz w:val=\"24\"/> and round-trips", async () => {
    const p = await createTmpDoc("color me");
    await formatText(p, "color me", { fontSize: 12 });
    const xml = await readRawDocXml(p);
    expect(xml).toContain('<w:sz w:val="24"');
    expect(xml).toContain('<w:szCs w:val="24"');
    expect(xmlIsWellFormed(p)).toBe(true);
    expect(pythonDocxOpens(p)).toBe(true);
  });
});

// -------------------------------------------------------------------------
// N4 — editParagraphs must PRESERVE footnote/endnote reference runs and field
// codes (fldChar / instrText / fldSimple); previously they were dropped,
// orphaning footnote content and breaking cross-references / page numbers.
// -------------------------------------------------------------------------

/**
 * Build a DOCX with:
 *  - para 0: "Lead" + a footnoteReference run + " tail"
 *  - para 1: a REF field (begin fldChar / instrText / end fldChar) + an
 *    endnoteReference run
 * plus the minimal footnotes.xml / endnotes.xml parts so python-docx opens it.
 */
async function createDocWithNotesAndFields(): Promise<string> {
  const p = tmpDocxPath();
  trackTmpFile(p);
  const JSZip = (await import("jszip")).default;
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:p><w:r><w:t>Lead</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r><w:r><w:t> tail</w:t></w:r></w:p>
<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF _Ref1 \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>see</w:t></w:r><w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="1"/></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
</w:body>
</w:document>`;
  const footnotesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
<w:footnote w:id="1"><w:p><w:r><w:t>The footnote body.</w:t></w:r></w:p></w:footnote>
</w:footnotes>`;
  const endnotesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>
<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>
<w:endnote w:id="1"><w:p><w:r><w:t>The endnote body.</w:t></w:r></w:p></w:endnote>
</w:endnotes>`;
  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/></w:style>
<w:style w:type="character" w:styleId="EndnoteReference"><w:name w:val="endnote reference"/></w:style>
</w:styles>`;
  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
</Types>`;
  const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
  const docRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
</Relationships>`;
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypesXml);
  zip.file("_rels/.rels", relsXml);
  zip.file("word/document.xml", documentXml);
  zip.file("word/styles.xml", stylesXml);
  zip.file("word/footnotes.xml", footnotesXml);
  zip.file("word/endnotes.xml", endnotesXml);
  zip.file("word/_rels/document.xml.rels", docRelsXml);
  const buf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  await fs.writeFile(p, buf);
  return p;
}

function countMatches(xml: string, re: RegExp): number {
  return (xml.match(re) ?? []).length;
}

describe("N4: editParagraphs preserves footnote/endnote refs and field codes", () => {
  for (const tracked of [true, false]) {
    it(`(${tracked ? "tracked" : "untracked"}) keeps footnoteRef/endnoteRef/fldChar/instrText and inserts new text`, async () => {
      const p = await createDocWithNotesAndFields();
      const before = await readRawDocXml(p);
      const fnBefore = countMatches(before, /<w:footnoteReference\b/g);
      const enBefore = countMatches(before, /<w:endnoteReference\b/g);
      const fldBefore = countMatches(before, /<w:fldChar\b/g);
      const instrBefore = countMatches(before, /<w:instrText\b/g);
      expect(fnBefore).toBe(1);
      expect(enBefore).toBe(1);
      expect(fldBefore).toBe(2);
      expect(instrBefore).toBe(1);

      await editParagraphs(
        p,
        [
          { paragraphIndex: 0, newText: "Lead REWRITTEN tail" },
          { paragraphIndex: 1, newText: "see REWRITTEN" },
        ],
        tracked,
      );

      const after = await readRawDocXml(p);
      // Structural runs preserved (count unchanged).
      expect(countMatches(after, /<w:footnoteReference\b/g)).toBe(fnBefore);
      expect(countMatches(after, /<w:endnoteReference\b/g)).toBe(enBefore);
      expect(countMatches(after, /<w:fldChar\b/g)).toBe(fldBefore);
      expect(countMatches(after, /<w:instrText\b/g)).toBe(instrBefore);
      // New text present.
      expect(after).toContain("REWRITTEN");
      // File well-formed + openable.
      expect(xmlIsWellFormed(p)).toBe(true);
      expect(pythonDocxOpens(p)).toBe(true);
    });
  }

  it("(untracked) preserves a fldSimple field run", async () => {
    const pth = tmpDocxPath();
    trackTmpFile(pth);
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>Page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:t> of N</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
</w:body>
</w:document>`;
    await writeMinimalDocx(pth, documentXml);
    await editParagraphs(pth, [{ paragraphIndex: 0, newText: "Page X of Y" }], false);
    const after = await readRawDocXml(pth);
    expect(countMatches(after, /<w:fldSimple\b/g)).toBe(1);
    expect(after).toContain("Page X of Y");
    expect(xmlIsWellFormed(pth)).toBe(true);
  });
});
