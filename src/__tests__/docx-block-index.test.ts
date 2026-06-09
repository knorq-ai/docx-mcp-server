/**
 * Block-index space unification (QA findings C1 critical + M5 medium).
 *
 * Two block-enumeration schemes used to disagree whenever a top-level `w:sdt`
 * (containing paragraphs) preceded a block:
 *   - read_document / get_document_info / search_text / ensure_anchors-by-index
 *     DESCEND into `w:sdt > w:sdtContent` and count each inner `w:p` as a block.
 *   - the index-consuming edit/table/format tools resolved a block index via a
 *     plain top-level body scan, which SKIPPED `w:sdt` entirely.
 * The off-by-N mismatch made the documented `search_text → edit_table_cells`
 * workflow silently edit the WRONG table (data corruption), or throw
 * NOT_A_TABLE / INDEX_OUT_OF_RANGE for indices the read tools call valid.
 *
 * These tests pin the unified behaviour: every index-consuming tool now agrees
 * with the read/search numbering (the numbering users actually see). For docs
 * with NO top-level w:sdt the behaviour is unchanged (the common case).
 */

import { describe, it, expect, afterEach } from "vitest";
import {
  cleanupTmpFiles,
  writeMinimalDocx,
  readRawDocXml,
  tmpDocxPath,
  trackTmpFile,
} from "./helpers.js";
import {
  searchTextStructured,
  editTableCells,
  readTableStructureStructured,
  readTableCellStructured,
  getParagraphFormatStructured,
  ensureAnchorsStructured,
  insertParagraphs,
  deleteParagraphs,
  editTableParagraphs,
  readDocument,
} from "../docx-engine.js";
import { EngineError } from "../docx-engine.js";

afterEach(cleanupTmpFiles);

const DOC_OPEN = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>`;
const DOC_CLOSE = `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
</w:body>
</w:document>`;

function sdtWithParagraph(text: string): string {
  return `<w:sdt><w:sdtPr><w:tag w:val="Field"/></w:sdtPr><w:sdtContent>
    <w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>
  </w:sdtContent></w:sdt>`;
}

function oneCellTable(cellText: string): string {
  return `<w:tbl>
    <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
    <w:tblGrid><w:gridCol w:w="0"/></w:tblGrid>
    <w:tr><w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>
      <w:p><w:r><w:t xml:space="preserve">${cellText}</w:t></w:r></w:p>
    </w:tc></w:tr>
  </w:tbl>`;
}

function paragraph(text: string): string {
  return `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

async function makeDoc(...bodyParts: string[]): Promise<string> {
  const p = tmpDocxPath();
  trackTmpFile(p);
  await writeMinimalDocx(p, DOC_OPEN + bodyParts.join("\n") + DOC_CLOSE);
  return p;
}

// =========================================================================
// 1. The critical repro (C1): search_text -> edit_table_cells must hit the
//    table that actually contains the matched cell, never a different one.
// =========================================================================

describe("C1 critical: SDT-before-table does not cause a silent wrong-table edit", () => {
  it("edit_table_cells changes the table that search_text matched (CELL_ONE), never the other", async () => {
    const p = await makeDoc(
      sdtWithParagraph("SdtHeader"),
      oneCellTable("CELL_ONE"),
      oneCellTable("CELL_TWO"),
    );

    const search = await searchTextStructured(p, "CELL_ONE");
    expect(search.matches).toHaveLength(1);
    const B = search.matches[0].blockIndex;

    await editTableCells(
      p,
      [{ blockIndex: B, rowIndex: 0, colIndex: 0, newText: "XXX" }],
      false,
    );

    const xml = await readRawDocXml(p);
    // The table that held CELL_ONE must now hold XXX; CELL_TWO must be untouched.
    expect(xml).toContain("XXX");
    expect(xml).not.toContain("CELL_ONE");
    expect(xml).toContain("CELL_TWO");
    // And the edit must have landed in the FIRST table (the one with CELL_ONE),
    // i.e. XXX appears before CELL_TWO in document order.
    expect(xml.indexOf("XXX")).toBeLessThan(xml.indexOf("CELL_TWO"));
  });

  it("read_table_structure at the search-reported blockIndex describes the CELL_ONE table", async () => {
    const p = await makeDoc(
      sdtWithParagraph("SdtHeader"),
      oneCellTable("CELL_ONE"),
      oneCellTable("CELL_TWO"),
    );
    const search = await searchTextStructured(p, "CELL_ONE");
    const B = search.matches[0].blockIndex;

    const struct = await readTableStructureStructured(p, B);
    expect(struct.rows).toBe(1);
    const cellRead = await readTableCellStructured(p, B, 0, 0);
    expect(cellRead.paragraphs.map((x) => x.text)).toContain("CELL_ONE");
  });
});

// =========================================================================
// 2. Loud variant (C1): indices the read tools report must resolve, not throw.
// =========================================================================

describe("C1: read-tool block indices resolve in the index-consuming tools (no spurious NOT_A_TABLE / OOR)", () => {
  it("a table after an SDT is reachable by read_table_structure; the trailing paragraph by get_paragraph_format", async () => {
    const p = await makeDoc(
      paragraph("Alpha"),
      sdtWithParagraph("InsideSDT"),
      oneCellTable("TableCell"),
      paragraph("Omega"),
    );

    // read_document numbering: [0]Alpha [1]InsideSDT [2]table [3]Omega
    const tableMatch = await searchTextStructured(p, "TableCell");
    expect(tableMatch.matches).toHaveLength(1);
    const tableBlock = tableMatch.matches[0].blockIndex;
    expect(tableBlock).toBe(2);

    // Must resolve as a table (previously threw NOT_A_TABLE).
    const struct = await readTableStructureStructured(p, tableBlock);
    expect(struct.rows).toBe(1);

    const omegaMatch = await searchTextStructured(p, "Omega");
    const omegaBlock = omegaMatch.matches[0].blockIndex;
    expect(omegaBlock).toBe(3);

    // Must resolve as a paragraph (previously threw INDEX_OUT_OF_RANGE 0-2).
    const fmt = await getParagraphFormatStructured(p, omegaBlock);
    expect(fmt.paragraphIndex).toBe(3);
  });

  it("preserves NOT_A_PARAGRAPH / NOT_A_TABLE semantics against the unified list", async () => {
    const p = await makeDoc(
      paragraph("Alpha"),
      sdtWithParagraph("InsideSDT"),
      oneCellTable("TableCell"),
      paragraph("Omega"),
    );
    // block 2 is the table -> get_paragraph_format must reject it as not a paragraph.
    await expect(getParagraphFormatStructured(p, 2)).rejects.toMatchObject({
      code: "NOT_A_PARAGRAPH",
    });
    // block 3 is Omega (paragraph) -> read_table_structure must reject as not a table.
    await expect(readTableStructureStructured(p, 3)).rejects.toMatchObject({
      code: "NOT_A_TABLE",
    });
    // out of range against the unified count (4 blocks: 0-3).
    await expect(getParagraphFormatStructured(p, 4)).rejects.toMatchObject({
      code: "INDEX_OUT_OF_RANGE",
    });
  });
});

// =========================================================================
// 3. M5: ensure_anchors index column matches read_document show_anchors.
// =========================================================================

describe("M5: ensure_anchors index column matches read_document/search_text for SDT docs", () => {
  it("numbers blocks the same way as read_document show_anchors", async () => {
    const p = await makeDoc(
      paragraph("Alpha"),
      sdtWithParagraph("InsideSDT"),
      paragraph("Bravo"),
      paragraph("Charlie"),
    );

    const ensured = await ensureAnchorsStructured(p);
    // Unified numbering: [0]Alpha [1]InsideSDT(no anchor) [2]Bravo [3]Charlie.
    const byIndex = new Map(ensured.blocks.map((b) => [b.index, b]));
    expect(byIndex.get(0)?.textPreview).toBe("Alpha");
    expect(byIndex.get(2)?.textPreview).toBe("Bravo");
    expect(byIndex.get(3)?.textPreview).toBe("Charlie");

    // search_text must report Bravo at the SAME index ensure_anchors prints, with
    // the SAME anchor.
    const bravo = await searchTextStructured(p, "Bravo");
    expect(bravo.matches[0].blockIndex).toBe(2);
    expect(bravo.matches[0].anchor).toBe(byIndex.get(2)?.anchor ?? undefined);

    // read_document show_anchors prints Bravo at [2] with that anchor too.
    const read = await readDocument(p, undefined, undefined, false, true);
    const anchor2 = byIndex.get(2)?.anchor;
    expect(read).toContain(`[2] @${anchor2} Bravo`);
  });
});

// =========================================================================
// 4. No-SDT sanity: every tool agrees exactly (the common case is unchanged).
// =========================================================================

describe("no-SDT docs: read/search/edit numbering agree exactly (no behaviour change)", () => {
  it("paragraph + table + paragraph resolve identically across tools", async () => {
    const p = await makeDoc(
      paragraph("First"),
      oneCellTable("InTable"),
      paragraph("Last"),
    );

    const inTable = await searchTextStructured(p, "InTable");
    expect(inTable.matches[0].blockIndex).toBe(1);
    const struct = await readTableStructureStructured(p, 1);
    expect(struct.rows).toBe(1);

    const last = await searchTextStructured(p, "Last");
    expect(last.matches[0].blockIndex).toBe(2);
    const fmt = await getParagraphFormatStructured(p, 2);
    expect(fmt.paragraphIndex).toBe(2);

    // edit_table_cells on block 1 edits the table.
    await editTableCells(
      p,
      [{ blockIndex: 1, rowIndex: 0, colIndex: 0, newText: "EDITED" }],
      false,
    );
    const xml = await readRawDocXml(p);
    expect(xml).toContain("EDITED");
    expect(xml).not.toContain("InTable");
  });
});

// =========================================================================
// 5. Insert / delete around and within an SDT.
// =========================================================================

describe("insert/delete with a leading SDT: body-level correct; SDT-inner handled safely", () => {
  it("body-level insert lands at the unified index (after the SDT), shifting later blocks", async () => {
    const p = await makeDoc(
      sdtWithParagraph("SdtHeader"),
      paragraph("Body0"),
      paragraph("Body1"),
    );
    // Unified: [0]SdtHeader [1]Body0 [2]Body1. Insert before block 2 (Body1).
    await insertParagraphs(p, [{ text: "INSERTED", position: 2 }], false);

    const read = await readDocument(p);
    // New order: [0]SdtHeader [1]Body0 [2]INSERTED [3]Body1.
    expect(read).toContain("[1] Body0");
    expect(read).toContain("[2] INSERTED");
    expect(read).toContain("[3] Body1");

    const xml = await readRawDocXml(p);
    // INSERTED must sit between Body0 and Body1 in document order, and must NOT
    // be inside the SDT.
    expect(xml.indexOf("Body0")).toBeLessThan(xml.indexOf("INSERTED"));
    expect(xml.indexOf("INSERTED")).toBeLessThan(xml.indexOf("Body1"));
    expect(xml.indexOf("</w:sdt>")).toBeLessThan(xml.indexOf("INSERTED"));
  });

  it("body-level delete removes the unified-index block (a body paragraph), leaving the SDT intact", async () => {
    const p = await makeDoc(
      sdtWithParagraph("SdtHeader"),
      paragraph("Body0"),
      paragraph("Body1"),
    );
    // Delete block 1 (Body0) untracked.
    await deleteParagraphs(p, [1], false);

    const read = await readDocument(p);
    expect(read).toContain("SdtHeader");
    expect(read).not.toContain("Body0");
    expect(read).toContain("Body1");
    const xml = await readRawDocXml(p);
    expect(xml).toContain("SdtHeader"); // SDT untouched
  });

  it("insert before an SDT-inner block throws a clear error (does not silently target a body block)", async () => {
    const p = await makeDoc(
      sdtWithParagraph("SdtHeader"),
      paragraph("Body0"),
    );
    // Block 0 is the SDT-inner paragraph. A position insert there must be a LOUD
    // failure, never a silent body splice.
    await expect(
      insertParagraphs(p, [{ text: "NOPE", position: 0 }], false),
    ).rejects.toBeInstanceOf(EngineError);

    // The document must be unchanged: NOPE never written.
    const xml = await readRawDocXml(p);
    expect(xml).not.toContain("NOPE");
  });

  it("delete of an SDT-inner block throws a clear error (does not silently delete a body block)", async () => {
    const p = await makeDoc(
      sdtWithParagraph("SdtHeader"),
      paragraph("Body0"),
    );
    await expect(deleteParagraphs(p, [0], false)).rejects.toBeInstanceOf(
      EngineError,
    );
    const xml = await readRawDocXml(p);
    expect(xml).toContain("SdtHeader"); // untouched
    expect(xml).toContain("Body0");
  });
});

// =========================================================================
// 6. Table-paragraph ops (edit_table_paragraphs etc.) resolve the table by the
//    unified index too.
// =========================================================================

describe("table-paragraph ops resolve the table by unified index past an SDT", () => {
  it("edit_table_paragraphs edits the correct table after an SDT", async () => {
    const p = await makeDoc(
      sdtWithParagraph("SdtHeader"),
      oneCellTable("OLD"),
    );
    const search = await searchTextStructured(p, "OLD");
    const B = search.matches[0].blockIndex;
    expect(B).toBe(1);

    await editTableParagraphs(
      p,
      [{ blockIndex: B, rowIndex: 0, colIndex: 0, paragraphIndex: 0, newText: "NEWVAL" }],
      false,
    );
    const xml = await readRawDocXml(p);
    expect(xml).toContain("NEWVAL");
    expect(xml).not.toContain("OLD");
  });
});
