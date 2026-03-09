import { describe, it, expect, afterEach } from "vitest";
import {
  createTmpDoc,
  cleanupTmpFiles,
  readRawDocXml,
} from "./helpers.js";
import {
  readDocument,
  editParagraphs,
  insertParagraphs,
  setHeadingBulk,
  editTableCells,
  insertTable,
  acceptAllChanges,
} from "../docx-engine.js";

afterEach(cleanupTmpFiles);

// =========================================================================
// editParagraphs (bulk)
// =========================================================================

describe("editParagraphs", () => {
  it("edits multiple paragraphs in one call", async () => {
    const p = await createTmpDoc("Line one\nLine two\nLine three");
    await editParagraphs(
      p,
      [
        { paragraphIndex: 0, newText: "Updated one" },
        { paragraphIndex: 2, newText: "Updated three" },
      ],
      false,
    );
    const doc = await readDocument(p);
    expect(doc).toContain("Updated one");
    expect(doc).toContain("Line two");
    expect(doc).toContain("Updated three");
  });

  it("produces tracked changes with del/ins markup", async () => {
    const p = await createTmpDoc("Original text\nSecond paragraph");
    await editParagraphs(
      p,
      [
        { paragraphIndex: 0, newText: "Modified text" },
        { paragraphIndex: 1, newText: "Changed paragraph" },
      ],
      true,
      "TestBot",
    );
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:del");
    expect(xml).toContain("w:ins");
    expect(xml).toContain("TestBot");

    // Accept all and verify final text
    await acceptAllChanges(p);
    const doc = await readDocument(p);
    expect(doc).toContain("Modified text");
    expect(doc).toContain("Changed paragraph");
  });

  it("throws INDEX_OUT_OF_RANGE for invalid index", async () => {
    const p = await createTmpDoc("Only one");
    await expect(
      editParagraphs(p, [{ paragraphIndex: 99, newText: "x" }], false),
    ).rejects.toThrow("out of range");
  });

  it("throws NOT_A_PARAGRAPH for table block", async () => {
    const p = await createTmpDoc("Before table");
    await insertTable(p, -1, 2, 2);
    // Block 1 is the table
    await expect(
      editParagraphs(p, [{ paragraphIndex: 1, newText: "x" }], false),
    ).rejects.toThrow("not a paragraph");
  });

  it("returns correct summary message", async () => {
    const p = await createTmpDoc("A\nB\nC");
    const result = await editParagraphs(
      p,
      [
        { paragraphIndex: 0, newText: "X" },
        { paragraphIndex: 1, newText: "Y" },
      ],
      false,
    );
    expect(result).toContain("2 paragraph(s)");
  });

  it("handles empty edits array without file I/O", async () => {
    const p = await createTmpDoc("Unchanged");
    const result = await editParagraphs(p, [], false);
    expect(result).toContain("No edits to apply");
    const doc = await readDocument(p);
    expect(doc).toContain("Unchanged");
  });
});

// =========================================================================
// insertParagraphs (bulk)
// =========================================================================

describe("insertParagraphs", () => {
  it("inserts multiple paragraphs in one call", async () => {
    const p = await createTmpDoc("First\nLast");
    await insertParagraphs(
      p,
      [
        { text: "Inserted A", position: 1 },
        { text: "Inserted B", position: 1 },
      ],
      false,
    );
    const doc = await readDocument(p);
    expect(doc).toContain("Inserted A");
    expect(doc).toContain("Inserted B");
  });

  it("handles position -1 for append", async () => {
    const p = await createTmpDoc("Start");
    await insertParagraphs(
      p,
      [
        { text: "End A", position: -1 },
        { text: "End B", position: -1 },
      ],
      false,
    );
    const doc = await readDocument(p);
    expect(doc).toContain("End A");
    expect(doc).toContain("End B");
  });

  it("applies paragraph styles", async () => {
    const p = await createTmpDoc("Existing");
    await insertParagraphs(
      p,
      [{ text: "New Heading", position: 0, style: "Heading1" }],
      false,
    );
    const doc = await readDocument(p);
    expect(doc).toContain("(H1)");
    expect(doc).toContain("New Heading");
  });

  it("produces tracked changes", async () => {
    const p = await createTmpDoc("Existing");
    await insertParagraphs(
      p,
      [{ text: "Tracked insert", position: 0 }],
      true,
      "BulkBot",
    );
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:ins");
    expect(xml).toContain("BulkBot");
  });

  it("returns correct summary", async () => {
    const p = await createTmpDoc("A");
    const result = await insertParagraphs(
      p,
      [
        { text: "B", position: 0 },
        { text: "C", position: -1 },
        { text: "D", position: -1 },
      ],
      false,
    );
    expect(result).toContain("3 paragraph(s)");
  });

  it("inserts at different specific positions correctly", async () => {
    const p = await createTmpDoc("P0\nP1\nP2");
    await insertParagraphs(
      p,
      [
        { text: "Before P1", position: 1 },
        { text: "Before P2", position: 2 },
      ],
      false,
    );
    const doc = await readDocument(p);
    const posP0 = doc.indexOf("P0");
    const posBeforeP1 = doc.indexOf("Before P1");
    const posP1 = doc.indexOf("[1]") > -1 ? doc.indexOf("P1", doc.indexOf("[1]")) : doc.indexOf("P1");
    const posBeforeP2 = doc.indexOf("Before P2");
    expect(posP0).toBeLessThan(posBeforeP1);
    expect(posBeforeP1).toBeLessThan(posBeforeP2);
  });

  it("handles empty items array without file I/O", async () => {
    const p = await createTmpDoc("Unchanged");
    const result = await insertParagraphs(p, [], false);
    expect(result).toContain("No paragraphs to insert");
    const doc = await readDocument(p);
    expect(doc).toContain("Unchanged");
  });
});

// =========================================================================
// setHeadingBulk
// =========================================================================

describe("setHeadingBulk", () => {
  it("sets multiple headings in one call", async () => {
    const p = await createTmpDoc("Title\nSubtitle\nBody\nAnother");
    await setHeadingBulk(p, [
      { paragraphIndex: 0, level: 1 },
      { paragraphIndex: 1, level: 2 },
      { paragraphIndex: 3, level: 3 },
    ]);
    const doc = await readDocument(p);
    expect(doc).toContain("(H1)");
    expect(doc).toContain("(H2)");
    expect(doc).toContain("(H3)");
  });

  it("validates heading level range", async () => {
    const p = await createTmpDoc("Test");
    await expect(
      setHeadingBulk(p, [{ paragraphIndex: 0, level: 0 }]),
    ).rejects.toThrow("between 1 and 9");
    await expect(
      setHeadingBulk(p, [{ paragraphIndex: 0, level: 10 }]),
    ).rejects.toThrow("between 1 and 9");
  });

  it("throws INDEX_OUT_OF_RANGE for invalid index", async () => {
    const p = await createTmpDoc("One");
    await expect(
      setHeadingBulk(p, [{ paragraphIndex: 99, level: 1 }]),
    ).rejects.toThrow("out of range");
  });

  it("throws NOT_A_PARAGRAPH for table block", async () => {
    const p = await createTmpDoc("Before");
    await insertTable(p, -1, 2, 2);
    await expect(
      setHeadingBulk(p, [{ paragraphIndex: 1, level: 1 }]),
    ).rejects.toThrow("not a paragraph");
  });

  it("returns correct summary", async () => {
    const p = await createTmpDoc("A\nB");
    const result = await setHeadingBulk(p, [
      { paragraphIndex: 0, level: 1 },
      { paragraphIndex: 1, level: 2 },
    ]);
    expect(result).toContain("2 paragraph(s)");
  });

  it("sets outline level correctly in XML", async () => {
    const p = await createTmpDoc("Test");
    await setHeadingBulk(p, [{ paragraphIndex: 0, level: 3 }]);
    const xml = await readRawDocXml(p);
    expect(xml).toContain('w:val="Heading3"');
    expect(xml).toContain('w:val="2"'); // outlineLvl = level - 1
  });

  it("handles empty items array without file I/O", async () => {
    const p = await createTmpDoc("Unchanged");
    const result = await setHeadingBulk(p, []);
    expect(result).toContain("No headings to set");
  });
});

// =========================================================================
// editTableCells (bulk)
// =========================================================================

describe("editTableCells", () => {
  it("edits multiple cells in one call", async () => {
    const p = await createTmpDoc("");
    await insertTable(p, 0, 2, 2, [
      ["A1", "B1"],
      ["A2", "B2"],
    ]);
    await editTableCells(
      p,
      [
        { blockIndex: 0, rowIndex: 0, colIndex: 0, newText: "X1" },
        { blockIndex: 0, rowIndex: 1, colIndex: 1, newText: "Y2" },
      ],
      false,
    );
    const doc = await readDocument(p);
    expect(doc).toContain("X1");
    expect(doc).toContain("B1");
    expect(doc).toContain("A2");
    expect(doc).toContain("Y2");
  });

  it("produces tracked changes", async () => {
    const p = await createTmpDoc("");
    await insertTable(p, 0, 1, 2, [["old1", "old2"]]);
    await editTableCells(
      p,
      [
        { blockIndex: 0, rowIndex: 0, colIndex: 0, newText: "new1" },
        { blockIndex: 0, rowIndex: 0, colIndex: 1, newText: "new2" },
      ],
      true,
      "CellBot",
    );
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:del");
    expect(xml).toContain("w:ins");
    expect(xml).toContain("CellBot");

    await acceptAllChanges(p);
    const doc = await readDocument(p);
    expect(doc).toContain("new1");
    expect(doc).toContain("new2");
  });

  it("throws NOT_A_TABLE for paragraph block", async () => {
    const p = await createTmpDoc("Paragraph");
    await expect(
      editTableCells(
        p,
        [{ blockIndex: 0, rowIndex: 0, colIndex: 0, newText: "x" }],
        false,
      ),
    ).rejects.toThrow("not a table");
  });

  it("throws INDEX_OUT_OF_RANGE for invalid row", async () => {
    const p = await createTmpDoc("");
    await insertTable(p, 0, 2, 2);
    await expect(
      editTableCells(
        p,
        [{ blockIndex: 0, rowIndex: 99, colIndex: 0, newText: "x" }],
        false,
      ),
    ).rejects.toThrow("out of range");
  });

  it("throws INDEX_OUT_OF_RANGE for invalid column", async () => {
    const p = await createTmpDoc("");
    await insertTable(p, 0, 2, 2);
    await expect(
      editTableCells(
        p,
        [{ blockIndex: 0, rowIndex: 0, colIndex: 99, newText: "x" }],
        false,
      ),
    ).rejects.toThrow("out of range");
  });

  it("returns correct summary", async () => {
    const p = await createTmpDoc("");
    await insertTable(p, 0, 2, 2);
    const result = await editTableCells(
      p,
      [
        { blockIndex: 0, rowIndex: 0, colIndex: 0, newText: "a" },
        { blockIndex: 0, rowIndex: 0, colIndex: 1, newText: "b" },
        { blockIndex: 0, rowIndex: 1, colIndex: 0, newText: "c" },
      ],
      false,
    );
    expect(result).toContain("3 table cell(s)");
  });

  it("handles empty edits array without file I/O", async () => {
    const result = await editTableCells("/tmp/nonexistent.docx", [], false);
    expect(result).toContain("No cell edits to apply");
  });
});
