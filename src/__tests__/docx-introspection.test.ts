import { describe, it, expect, afterEach } from "vitest";
import {
  createTmpDoc,
  cleanupTmpFiles,
  createDocWithNumberedParagraph,
} from "./helpers.js";
import {
  insertTable,
  setParagraphFormats,
  readTableStructure,
  readTableStructureStructured,
  readTableCell,
  readTableCellStructured,
  getParagraphFormat,
  getParagraphFormatStructured,
  searchTextStructured,
} from "../docx-engine.js";

afterEach(cleanupTmpFiles);

// =========================================================================
// Issue #6.2 — read_table_structure
// =========================================================================

describe("read_table_structure", () => {
  it("reports dimensions and per-cell previews", async () => {
    const p = await createTmpDoc("before");
    await insertTable(p, -1, 2, 3, [
      ["A1", "B1", "C1"],
      ["A2", "B2", "C2"],
    ]);
    const r = await readTableStructureStructured(p, 1);
    expect(r.rows).toBe(2);
    expect(r.cols).toBe(3);
    expect(r.cells).toHaveLength(6);
    const c12 = r.cells.find((c) => c.rowIndex === 1 && c.colIndex === 2);
    expect(c12?.preview).toBe("C2");
    expect(c12?.gridSpan).toBe(1);
    expect(c12?.vMerge).toBeNull();
  });

  it("text form includes a <json> block", async () => {
    const p = await createTmpDoc("before");
    await insertTable(p, -1, 1, 1, [["only"]]);
    const text = await readTableStructure(p, 1);
    expect(text).toContain("1 row(s)");
    expect(text).toContain("<json>");
  });

  it("throws NOT_A_TABLE for a paragraph block", async () => {
    const p = await createTmpDoc("just a paragraph");
    await expect(readTableStructureStructured(p, 0)).rejects.toMatchObject({
      code: "NOT_A_TABLE",
    });
  });

  it("throws INDEX_OUT_OF_RANGE for a bad block index", async () => {
    const p = await createTmpDoc("x");
    await expect(readTableStructureStructured(p, 99)).rejects.toMatchObject({
      code: "INDEX_OUT_OF_RANGE",
    });
  });
});

// =========================================================================
// Issue #6.2 — read_table_cell
// =========================================================================

describe("read_table_cell", () => {
  it("returns the cell's paragraphs", async () => {
    const p = await createTmpDoc("before");
    await insertTable(p, -1, 2, 2, [
      ["A1", "B1"],
      ["A2", "B2"],
    ]);
    const r = await readTableCellStructured(p, 1, 1, 0);
    expect(r.paragraphs).toHaveLength(1);
    expect(r.paragraphs[0].text).toBe("A2");
  });

  it("text form lists the cell paragraphs and a <json> block", async () => {
    const p = await createTmpDoc("before");
    await insertTable(p, -1, 1, 1, [["only cell"]]);
    const text = await readTableCell(p, 1, 0, 0);
    expect(text).toContain("Cell [0,0]");
    expect(text).toContain("only cell");
    expect(text).toContain("1 paragraph(s)");
    expect(text).toContain("<json>");
  });

  it("throws INDEX_OUT_OF_RANGE for a bad column index", async () => {
    const p = await createTmpDoc("before");
    await insertTable(p, -1, 1, 2, [["A", "B"]]);
    await expect(readTableCellStructured(p, 1, 0, 9)).rejects.toMatchObject({
      code: "INDEX_OUT_OF_RANGE",
    });
  });
});

// =========================================================================
// Issue #6.5 — get_paragraph_format
// =========================================================================

describe("get_paragraph_format", () => {
  it("reports style, alignment, numbering and indentation", async () => {
    const p = await createDocWithNumberedParagraph("clause", 14, 0, {
      alignment: "center",
      indentLeft: 720,
    });
    const r = await getParagraphFormatStructured(p, 0);
    expect(r.numId).toBe(14);
    expect(r.numLevel).toBe(0);
    expect(r.alignment).toBe("center");
    expect(r.indentLeft).toBe(720);
  });

  it("normalizes justify alignment and reports spacing in points", async () => {
    const p = await createTmpDoc("body");
    await setParagraphFormats(p, [
      { indices: [0], format: { alignment: "justify", spaceBefore: 6, spaceAfter: 12 } },
    ]);
    const r = await getParagraphFormatStructured(p, 0);
    expect(r.alignment).toBe("justify");
    expect(r.spaceBefore).toBe(6);
    expect(r.spaceAfter).toBe(12);
  });

  it("text form includes a <json> block", async () => {
    const p = await createTmpDoc("body");
    const text = await getParagraphFormat(p, 0);
    expect(text).toContain("Format of paragraph 0");
    expect(text).toContain("<json>");
  });

  it("throws NOT_A_PARAGRAPH for a table block", async () => {
    const p = await createTmpDoc("before");
    await insertTable(p, -1, 1, 1, [["cell"]]);
    await expect(getParagraphFormatStructured(p, 1)).rejects.toMatchObject({
      code: "NOT_A_PARAGRAPH",
    });
  });
});

// =========================================================================
// Issue #6.3 — search_text reports row/col for table matches
// =========================================================================

describe("search_text table location", () => {
  it("reports rowIndex/colIndex for a match inside a table cell", async () => {
    const p = await createTmpDoc("intro paragraph");
    await insertTable(p, -1, 2, 2, [
      ["alpha", "beta"],
      ["gamma", "delta"],
    ]);
    const r = await searchTextStructured(p, "delta");
    expect(r.matches).toHaveLength(1);
    expect(r.matches[0].blockIndex).toBe(1);
    expect(r.matches[0].rowIndex).toBe(1);
    expect(r.matches[0].colIndex).toBe(1);
  });

  it("leaves rowIndex/colIndex undefined for a plain paragraph match", async () => {
    const p = await createTmpDoc("the keyword is here");
    const r = await searchTextStructured(p, "keyword");
    expect(r.matches).toHaveLength(1);
    expect(r.matches[0].blockIndex).toBe(0);
    expect(r.matches[0].rowIndex).toBeUndefined();
    expect(r.matches[0].colIndex).toBeUndefined();
  });

  it("finds matches in both a paragraph and a table cell", async () => {
    const p = await createTmpDoc("shared term in paragraph");
    await insertTable(p, -1, 1, 1, [["shared term in cell"]]);
    const r = await searchTextStructured(p, "shared term");
    expect(r.totalMatches).toBe(2);
    const para = r.matches.find((m) => m.rowIndex === undefined);
    const cell = r.matches.find((m) => m.rowIndex !== undefined);
    expect(para?.blockIndex).toBe(0);
    expect(cell?.blockIndex).toBe(1);
    expect(cell?.colIndex).toBe(0);
  });
});
