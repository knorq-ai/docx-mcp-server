import { describe, it, expect, afterEach } from "vitest";
import {
  createTmpDoc,
  cleanupTmpFiles,
  readRawDocXml,
  createDocWithPPrChange,
} from "./helpers.js";
import {
  insertTable,
  insertTableParagraphs,
  insertParagraphs,
  editTableParagraphs,
  deleteTableParagraphs,
  readDocument,
  acceptAllChanges,
  rejectAllChanges,
} from "../docx-engine.js";

afterEach(cleanupTmpFiles);

/** Count <w:p> elements in raw document.xml (covers `<w:p>`, `<w:p ...>`, `<w:p/>`). */
function countWp(xml: string): number {
  return (xml.match(/<w:p[ >/]/g) ?? []).length;
}

/** Build a 1×1 table whose only cell holds the given paragraphs (untracked). */
async function makeCellWithParagraphs(lines: string[]): Promise<string> {
  const p = await createTmpDoc("before table");
  await insertTable(p, -1, 1, 1, [[lines[0] ?? ""]]);
  for (let i = 1; i < lines.length; i++) {
    await insertTableParagraphs(
      p,
      [{ blockIndex: 1, rowIndex: 0, colIndex: 0, position: -1, text: lines[i] }],
      false,
    );
  }
  return p;
}

// =========================================================================
// insert_table_paragraphs
// =========================================================================

describe("multi-paragraph cell text extraction", () => {
  it("joins cell paragraphs with a real newline, not a literal backslash-n", async () => {
    const p = await makeCellWithParagraphs(["line one", "line two"]);
    const doc = await readDocument(p);
    // Regression: extractCellText used join("\\n") (literal backslash + n).
    expect(doc).not.toContain("line one\\nline two");
    expect(doc).toContain("line one\nline two");
  });
});

describe("insert_table_paragraphs", () => {
  it("appends a paragraph to a cell (untracked)", async () => {
    const p = await createTmpDoc("before table");
    await insertTable(p, -1, 1, 1, [["alpha"]]);
    // body paragraph + 1 cell paragraph = 2
    expect(countWp(await readRawDocXml(p))).toBe(2);

    await insertTableParagraphs(
      p,
      [{ blockIndex: 1, rowIndex: 0, colIndex: 0, position: -1, text: "beta" }],
      false,
    );
    const xml = await readRawDocXml(p);
    expect(countWp(xml)).toBe(3);
    expect(xml).toContain("alpha");
    expect(xml).toContain("beta");
  });

  it("inserts before a given position", async () => {
    const p = await makeCellWithParagraphs(["alpha", "gamma"]);
    await insertTableParagraphs(
      p,
      [{ blockIndex: 1, rowIndex: 0, colIndex: 0, position: 1, text: "beta" }],
      false,
    );
    const xml = await readRawDocXml(p);
    // Order in XML: alpha, beta, gamma
    expect(xml.indexOf("alpha")).toBeLessThan(xml.indexOf("beta"));
    expect(xml.indexOf("beta")).toBeLessThan(xml.indexOf("gamma"));
  });

  it("multiple inserts into the same cell keep array order", async () => {
    const p = await createTmpDoc("before table");
    await insertTable(p, -1, 1, 1, [["anchor"]]);
    await insertTableParagraphs(
      p,
      [
        { blockIndex: 1, rowIndex: 0, colIndex: 0, position: 0, text: "one" },
        { blockIndex: 1, rowIndex: 0, colIndex: 0, position: 0, text: "two" },
      ],
      false,
    );
    const xml = await readRawDocXml(p);
    expect(xml.indexOf("one")).toBeLessThan(xml.indexOf("two"));
    expect(xml.indexOf("two")).toBeLessThan(xml.indexOf("anchor"));
  });

  it("copy_format_from carries numbering from a sibling cell paragraph", async () => {
    const p = await createTmpDoc("before table");
    await insertTable(p, -1, 1, 1, [["first"]]);
    // Give paragraph 0 numbering, then insert a sibling copying its format.
    await insertTableParagraphs(
      p,
      [{ blockIndex: 1, rowIndex: 0, colIndex: 0, position: -1, text: "numbered", numId: 7 }],
      false,
    );
    // Now copy the numbered paragraph's format (index 1) onto a new one.
    await insertTableParagraphs(
      p,
      [{ blockIndex: 1, rowIndex: 0, colIndex: 0, position: -1, text: "copied", copyFormatFrom: 1 }],
      false,
    );
    const xml = await readRawDocXml(p);
    expect((xml.match(/w:numId w:val="7"/g) ?? []).length).toBe(2);
  });

  it("tracked insert wraps the new paragraph as an insertion", async () => {
    const p = await createTmpDoc("before table");
    await insertTable(p, -1, 1, 1, [["alpha"]]);
    await insertTableParagraphs(
      p,
      [{ blockIndex: 1, rowIndex: 0, colIndex: 0, position: -1, text: "beta" }],
      true,
    );
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:ins");
    await rejectAllChanges(p);
    expect(await readRawDocXml(p)).not.toContain("beta");
  });

  it("throws INDEX_OUT_OF_RANGE for a non-table block", async () => {
    const p = await createTmpDoc("just a paragraph");
    await expect(
      insertTableParagraphs(
        p,
        [{ blockIndex: 1, rowIndex: 0, colIndex: 0, position: 0, text: "x" }],
        false,
      ),
    ).rejects.toMatchObject({ code: "INDEX_OUT_OF_RANGE" });
  });
});

// =========================================================================
// edit_table_paragraphs
// =========================================================================

describe("edit_table_paragraphs", () => {
  it("edits one paragraph and leaves the others (untracked)", async () => {
    const p = await makeCellWithParagraphs(["alpha", "beta", "gamma"]);
    await editTableParagraphs(
      p,
      [{ blockIndex: 1, rowIndex: 0, colIndex: 0, paragraphIndex: 1, newText: "BETA" }],
      false,
    );
    const xml = await readRawDocXml(p);
    expect(xml).toContain("alpha");
    expect(xml).toContain("BETA");
    expect(xml).toContain("gamma");
    expect(xml).not.toMatch(/>beta</);
  });

  it("records a tracked edit as del/ins", async () => {
    const p = await makeCellWithParagraphs(["alpha", "beta"]);
    await editTableParagraphs(
      p,
      [{ blockIndex: 1, rowIndex: 0, colIndex: 0, paragraphIndex: 1, newText: "changed" }],
      true,
    );
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:delText");
    expect(xml).toContain("w:ins");
  });

  it("throws INDEX_OUT_OF_RANGE for a bad cell paragraph index", async () => {
    const p = await makeCellWithParagraphs(["only"]);
    await expect(
      editTableParagraphs(
        p,
        [{ blockIndex: 1, rowIndex: 0, colIndex: 0, paragraphIndex: 5, newText: "x" }],
        false,
      ),
    ).rejects.toMatchObject({ code: "INDEX_OUT_OF_RANGE" });
  });

  it("rejects a fractional paragraph index instead of silently no-op'ing", async () => {
    const p = await makeCellWithParagraphs(["alpha", "beta"]);
    await expect(
      editTableParagraphs(
        p,
        [{ blockIndex: 1, rowIndex: 0, colIndex: 0, paragraphIndex: 0.5, newText: "x" }],
        false,
      ),
    ).rejects.toMatchObject({ code: "INDEX_OUT_OF_RANGE" });
  });
});

// =========================================================================
// copy_format_from must not carry the source paragraph's tracked-change markup
// (shared buildPPrForNewParagraph fix; exercised here via body insert)
// =========================================================================

describe("copy_format_from strips stale revision metadata", () => {
  it("does not copy a source paragraph's w:pPrChange onto the new paragraph", async () => {
    const p = await createDocWithPPrChange("source");
    const before = await readRawDocXml(p);
    expect(before).toContain("w:pPrChange"); // sanity: source has one

    await insertParagraphs(
      p,
      [{ text: "copied", position: 1, copyFormatFrom: 0 }],
      false,
    );
    const after = await readRawDocXml(p);
    // Still exactly one w:pPrChange — the new paragraph did not inherit it.
    expect((after.match(/<w:pPrChange/g) ?? []).length).toBe(1);
  });
});

// =========================================================================
// delete_table_paragraphs
// =========================================================================

describe("delete_table_paragraphs", () => {
  it("deletes one paragraph and keeps the others (untracked)", async () => {
    const p = await makeCellWithParagraphs(["alpha", "beta", "gamma"]);
    const before = countWp(await readRawDocXml(p));
    await deleteTableParagraphs(
      p,
      [{ blockIndex: 1, rowIndex: 0, colIndex: 0, paragraphIndex: 1 }],
      false,
    );
    const xml = await readRawDocXml(p);
    expect(countWp(xml)).toBe(before - 1);
    expect(xml).toContain("alpha");
    expect(xml).toContain("gamma");
    expect(xml).not.toMatch(/>beta</);
  });

  it("keeps a blank paragraph when the cell's last paragraph is deleted", async () => {
    const p = await createTmpDoc("before table");
    await insertTable(p, -1, 1, 1, [["only"]]);
    await deleteTableParagraphs(
      p,
      [{ blockIndex: 1, rowIndex: 0, colIndex: 0, paragraphIndex: 0 }],
      false,
    );
    const xml = await readRawDocXml(p);
    expect(xml).not.toContain("only");
    // Cell still has a paragraph: body paragraph + one (blank) cell paragraph.
    expect(countWp(xml)).toBe(2);
    // Document still reads without error.
    expect(await readDocument(p)).toContain("[TABLE]");
  });

  it("tracked delete marks the paragraph deleted and round-trips", async () => {
    const p = await makeCellWithParagraphs(["alpha", "beta"]);
    await deleteTableParagraphs(
      p,
      [{ blockIndex: 1, rowIndex: 0, colIndex: 0, paragraphIndex: 1 }],
      true,
    );
    const tracked = await readRawDocXml(p);
    expect(tracked).toContain("w:del");
    // Accepted view hides the deleted text; rejected view restores it.
    const accepted = await readDocument(p);
    expect(accepted).not.toContain("beta");

    const p2 = await makeCellWithParagraphs(["alpha", "beta"]);
    await deleteTableParagraphs(
      p2,
      [{ blockIndex: 1, rowIndex: 0, colIndex: 0, paragraphIndex: 1 }],
      true,
    );
    await rejectAllChanges(p2);
    expect(await readDocument(p2)).toContain("beta");
  });

  it("accepting a tracked delete removes the text", async () => {
    const p = await makeCellWithParagraphs(["alpha", "beta"]);
    await deleteTableParagraphs(
      p,
      [{ blockIndex: 1, rowIndex: 0, colIndex: 0, paragraphIndex: 1 }],
      true,
    );
    await acceptAllChanges(p);
    expect(await readDocument(p)).not.toContain("beta");
  });
});
