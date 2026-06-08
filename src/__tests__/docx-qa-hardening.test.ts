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
import {
  createTmpDoc,
  cleanupTmpFiles,
  readRawDocXml,
  tmpDocxPath,
  trackTmpFile,
  writeMinimalDocx,
} from "./helpers.js";
import {
  createDocument,
  editParagraphs,
  insertParagraphs,
  insertTableParagraphs,
  replaceTexts,
  editTableCells,
  editTableParagraphs,
  deleteTableParagraphs,
  insertTable,
  searchText,
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
  EngineError,
  ErrorCode,
  formatText,
  highlightText,
} from "../docx-engine.js";

afterEach(cleanupTmpFiles);

const PYTHON = "/tmp/docx-venv/bin/python";

/** True if `word/document.xml` is well-formed XML (per xmllint). */
function xmlIsWellFormed(filePath: string): boolean {
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
  try {
    execFileSync(PYTHON, ["-c", `import docx; docx.Document(${JSON.stringify(filePath)})`]);
    return true;
  } catch {
    return false;
  }
}

/** Run an arbitrary python-docx snippet, returning stdout (throws on python error). */
function runPython(snippet: string): string {
  return execFileSync(PYTHON, ["-c", snippet], { encoding: "utf8" });
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
 * children of the OUTER table's cell[0,0], read straight from raw XML via
 * python-docx's lxml (so it reflects on-disk structure, not engine re-parse).
 */
function outerCellBlockChildren(filePath: string): string[] {
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

    const out = runPython(`
import docx
d = docx.Document(${JSON.stringify(p)})
print(len(d.tables[0].columns))
`).trim();
    expect(out).toBe("4");
  });
});

// =========================================================================
// H2 — a <w:tc> must end with <w:p>, never a nested <w:tbl>
// =========================================================================

describe("H2: cell content must end with a paragraph", () => {
  it("whole-cell untracked replace keeps a trailing <w:p> after a nested table", async () => {
    const p = await createDocWithNestedTrailingPara();
    // sanity: starts as [p, tbl, p]
    expect(outerCellBlockChildren(p)).toEqual(["p", "tbl", "p"]);

    // editTableCells(untracked) replaces the whole cell's paragraph content.
    await editTableCells(p, [{ blockIndex: 0, rowIndex: 0, colIndex: 0, newText: "NEW" }], false);

    const seq = outerCellBlockChildren(p);
    expect(seq[seq.length - 1]).toBe("p"); // last block child must be a paragraph
    // The nested table must survive.
    expect(seq).toContain("tbl");
    expect(pythonDocxOpens(p)).toBe(true);
    // Nested table still present per python-docx.
    const nested = runPython(`
import docx
d = docx.Document(${JSON.stringify(p)})
print(len(d.tables[0].rows[0].cells[0].tables))
`).trim();
    expect(nested).toBe("1");
  });

  it("deleting the trailing paragraph after a nested table re-adds a blank <w:p>", async () => {
    const p = await createDocWithNestedTrailingPara();
    expect(outerCellBlockChildren(p)).toEqual(["p", "tbl", "p"]);

    // Cell-local paragraph indices count only w:p children: [0]=BEFORE, [1]=AFTER.
    // Delete the trailing AFTER paragraph (untracked).
    await deleteTableParagraphs(p, [{ blockIndex: 0, rowIndex: 0, colIndex: 0, paragraphIndex: 1 }], false);

    const seq = outerCellBlockChildren(p);
    expect(seq[seq.length - 1]).toBe("p"); // must NOT end in tbl
    expect(seq).toContain("tbl");
    expect(pythonDocxOpens(p)).toBe(true);
  });

  it("deleting the FIRST paragraph (before the nested table) is unaffected and stays valid", async () => {
    const p = await createDocWithNestedTrailingPara();
    await deleteTableParagraphs(p, [{ blockIndex: 0, rowIndex: 0, colIndex: 0, paragraphIndex: 0 }], false);
    const seq = outerCellBlockChildren(p);
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

/** python-docx body paragraph texts (after XML §2.11 line-end normalization). */
function pyBodyParagraphs(filePath: string): string[] {
  const py = `
import docx, json
d = docx.Document(${JSON.stringify(filePath)})
print(json.dumps([p.text for p in d.paragraphs]))
`;
  return JSON.parse(runPython(py));
}

/** python-docx text of the first cell of the first table. */
function pyFirstCellText(filePath: string): string {
  const py = `
import docx, json
d = docx.Document(${JSON.stringify(filePath)})
print(json.dumps(d.tables[0].rows[0].cells[0].text))
`;
  return JSON.parse(runPython(py));
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
    expect(pyBodyParagraphs(p)).toEqual(["a", "b", "c"]);
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
    expect(pyBodyParagraphs(p)).toEqual(["a", "b"]);
  });

  it("insertParagraphs (untracked) splits CRLF with zero raw \\r", async () => {
    const p = await createTmpDoc("anchor");
    await insertParagraphs(p, [{ text: "ins1\r\nins2", position: -1 }], false);

    const xml = await readRawDocXml(p);
    expect(rawCRCount(xml)).toBe(0);
    expect(pyBodyParagraphs(p)).toContain("ins1");
    expect(pyBodyParagraphs(p)).toContain("ins2");
    expect(xmlIsWellFormed(p)).toBe(true);
  });

  it("editTableCells (untracked) splits CRLF cell text with zero raw \\r", async () => {
    const p = await createTmpDoc("seed");
    await insertTable(p, -1, 1, 1);
    await editTableCells(p, [{ blockIndex: 1, rowIndex: 0, colIndex: 0, newText: "r1\r\nr2" }], false);

    const xml = await readRawDocXml(p);
    expect(rawCRCount(xml)).toBe(0);
    // Two real paragraphs in the cell → python-docx joins them with one "\n".
    expect(pyFirstCellText(p)).toBe("r1\nr2");
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
    expect(pyBodyParagraphs(p)[0]).toBe("hello big\nplanet");
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
    expect(pyFirstCellText(p)).toBe("AA\nBB");
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
      expect(pyFirstCellText(p)).toBe("ORIG");
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
    expect(pyBodyParagraphs(p)).toEqual(["Existing only"]);
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
    expect(pyBodyParagraphs(p)).toEqual(["BASE0", "BASE1", "BASE2"]);
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
    expect(pyBodyParagraphs(p)).toEqual(["HELLO"]);
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
    expect(pyFirstCellText(p)).toBe("");
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
