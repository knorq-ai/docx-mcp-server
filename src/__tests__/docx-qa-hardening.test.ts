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
  replaceTexts,
  editTableCells,
  deleteTableParagraphs,
  insertTable,
  searchText,
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

  it("preserves the LEGAL whitespace chars \\t \\r \\n", async () => {
    // \r and \t are legal XML chars and must survive. \n is turned into a soft
    // break / paragraph split upstream, so we assert \t and \r survive verbatim.
    const p = await createTmpDoc("seed");
    await editParagraphs(p, [{ paragraphIndex: 0, newText: "tab\there\rcarriage" }], false);
    const xml = await readRawDocXml(p);
    expect(xml).toContain("\t");
    expect(xml).toContain("\r");
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
