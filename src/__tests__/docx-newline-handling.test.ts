import { describe, it, expect, afterEach } from "vitest";
import {
  createTmpDoc,
  cleanupTmpFiles,
  readRawDocXml,
  createDocWithNumberedParagraph,
} from "./helpers.js";
import {
  readDocument,
  editParagraphs,
  editTableCells,
  insertParagraphs,
  insertTable,
  getDocumentInfo,
  acceptAllChanges,
  rejectAllChanges,
} from "../docx-engine.js";

afterEach(cleanupTmpFiles);

/**
 * Count occurrences of a w:p opening tag in raw document.xml.
 * fast-xml-parser's builder emits self-closing `<w:p/>` for empty paragraphs,
 * so match both `<w:p>` and `<w:p ...>` and `<w:p/>`.
 */
function countParagraphs(xml: string): number {
  return (xml.match(/<w:p[ >/]/g) ?? []).length;
}

// =========================================================================
// Issue #4 — "\n" in edit text must become paragraph breaks, not literal text
// =========================================================================

describe("edit_paragraphs newline handling (untracked)", () => {
  it("splits a multi-line edit into separate paragraphs", async () => {
    const p = await createTmpDoc("original");
    await editParagraphs(
      p,
      [{ paragraphIndex: 0, newText: "line one\nline two\nline three" }],
      false,
    );

    const xml = await readRawDocXml(p);
    // No literal newline survives inside a single run's text.
    expect(xml).not.toContain("line one\nline two");
    // Each line is now its own paragraph.
    expect(countParagraphs(xml)).toBe(3);

    const doc = await readDocument(p);
    expect(doc).toContain("[0] line one");
    expect(doc).toContain("[1] line two");
    expect(doc).toContain("[2] line three");
  });

  it("preserves the source paragraph's numbering on every produced line", async () => {
    const p = await createDocWithNumberedParagraph("first item", 14, 0);
    await editParagraphs(
      p,
      [{ paragraphIndex: 0, newText: "item A\nitem B\nitem C" }],
      false,
    );

    const xml = await readRawDocXml(p);
    // numId 14 is carried onto each of the three split paragraphs.
    expect((xml.match(/w:numId w:val="14"/g) ?? []).length).toBe(3);

    const doc = await readDocument(p);
    expect(doc).toContain("item A");
    expect(doc).toContain("item B");
    expect(doc).toContain("item C");
  });

  it("renders an empty line as an empty paragraph", async () => {
    const p = await createTmpDoc("original");
    await editParagraphs(
      p,
      [{ paragraphIndex: 0, newText: "above\n\nbelow" }],
      false,
    );
    const info = await getDocumentInfo(p);
    // above / (empty) / below
    expect(info).toContain("Total blocks: 3");
  });

  it("leaves a single-line edit as one paragraph", async () => {
    const p = await createTmpDoc("original");
    await editParagraphs(p, [{ paragraphIndex: 0, newText: "just one line" }], false);
    const xml = await readRawDocXml(p);
    expect(countParagraphs(xml)).toBe(1);
    expect(await readDocument(p)).toContain("just one line");
  });

  it("collapses duplicate edits to the same paragraph (last-write-wins) when the first splits", async () => {
    const p = await createTmpDoc("original");
    await editParagraphs(
      p,
      [
        { paragraphIndex: 0, newText: "a\nb" },
        { paragraphIndex: 0, newText: "final" },
      ],
      false,
    );
    const doc = await readDocument(p);
    expect(doc).toContain("final");
    expect(doc).not.toContain("[1] b");
    const info = await getDocumentInfo(p);
    expect(info).toContain("Total blocks: 1");
  });

  it("keeps later edit targets correct after an earlier split shifts indices", async () => {
    const p = await createTmpDoc("alpha");
    await insertParagraphs(
      p,
      [
        { text: "beta", position: 1 },
        { text: "gamma", position: 2 },
      ],
      false,
    );
    // Edit block 0 (splits into 2) and block 2 (gamma) in the same call.
    await editParagraphs(
      p,
      [
        { paragraphIndex: 0, newText: "alpha-1\nalpha-2" },
        { paragraphIndex: 2, newText: "gamma-edited" },
      ],
      false,
    );
    const doc = await readDocument(p);
    expect(doc).toContain("alpha-1");
    expect(doc).toContain("alpha-2");
    expect(doc).toContain("beta");
    expect(doc).toContain("gamma-edited");
    expect(doc).not.toContain("gamma\n");
  });
});

describe("edit_paragraphs newline handling (tracked)", () => {
  it("uses soft breaks inside a single tracked paragraph (no literal newline, no paragraph split)", async () => {
    const p = await createTmpDoc("original");
    await editParagraphs(
      p,
      [{ paragraphIndex: 0, newText: "line one\nline two" }],
      true,
    );
    const xml = await readRawDocXml(p);
    expect(countParagraphs(xml)).toBe(1);
    expect(xml).toContain("<w:br");
    expect(xml).toContain("w:ins");
    expect(xml).not.toContain("line one\nline two");
  });

  it("round-trips through accept (keeps new multi-line text)", async () => {
    const p = await createTmpDoc("original");
    await editParagraphs(p, [{ paragraphIndex: 0, newText: "line one\nline two" }], true);
    await acceptAllChanges(p);
    const doc = await readDocument(p);
    expect(doc).toContain("line one");
    expect(doc).toContain("line two");
    expect(doc).not.toContain("original");
  });

  it("round-trips through reject (restores original)", async () => {
    const p = await createTmpDoc("original");
    await editParagraphs(p, [{ paragraphIndex: 0, newText: "line one\nline two" }], true);
    await rejectAllChanges(p);
    const doc = await readDocument(p);
    expect(doc).toContain("original");
    expect(doc).not.toContain("line one");
  });

  it("deleting across a previously-inserted soft break does not leave a literal newline on reject", async () => {
    // Repro: tracked multi-line edit -> accept (paragraph now holds a w:br) ->
    // tracked replacement whose deleted middle spans the break -> reject.
    // The deleted run must round-trip as a soft break, not a literal "\n".
    const p = await createTmpDoc("ab");
    await editParagraphs(p, [{ paragraphIndex: 0, newText: "a\nb" }], true);
    await acceptAllChanges(p);
    await editParagraphs(p, [{ paragraphIndex: 0, newText: "Z" }], true);
    await rejectAllChanges(p);

    const xml = await readRawDocXml(p);
    expect(xml).toContain("<w:br");
    // No w:t carries an embedded newline.
    expect(xml).not.toContain("a\nb");
    expect(xml).not.toMatch(/<w:t[^>]*>[^<]*\n[^<]*<\/w:t>/);
  });
});

describe("edit_table_cells newline handling", () => {
  it("untracked: splits cell content into separate paragraphs", async () => {
    const p = await createTmpDoc("before table");
    await insertTable(p, -1, 1, 1, [["old content"]]);
    await editTableCells(
      p,
      [{ blockIndex: 1, rowIndex: 0, colIndex: 0, newText: "1. a\n2. b\n3. c" }],
      false,
    );
    const xml = await readRawDocXml(p);
    expect(xml).not.toContain("1. a\n2. b");
    expect(xml).not.toContain("old content");

    const doc = await readDocument(p);
    expect(doc).toContain("1. a");
    expect(doc).toContain("2. b");
    expect(doc).toContain("3. c");
  });

  it("untracked: re-editing a multi-line cell fully replaces previous content", async () => {
    const p = await createTmpDoc("before table");
    await insertTable(p, -1, 1, 1, [["old"]]);
    await editTableCells(
      p,
      [{ blockIndex: 1, rowIndex: 0, colIndex: 0, newText: "x1\nx2\nx3" }],
      false,
    );
    await editTableCells(
      p,
      [{ blockIndex: 1, rowIndex: 0, colIndex: 0, newText: "y1\ny2" }],
      false,
    );
    // Assert on the raw cell XML (ground truth), not readDocument output:
    // readDocument embeds the absolute file Path, so a temp dir whose name
    // happens to contain "x1"/"x3" (e.g. macOS /var/folders/.../h73hykpx0x1bd.../)
    // would make a substring-absence check false-fail on some machines.
    const xml = await readRawDocXml(p);
    expect(xml).toContain("y1");
    expect(xml).toContain("y2");
    expect(xml).not.toContain("x1");
    expect(xml).not.toContain("x2");
    expect(xml).not.toContain("x3");
  });

  it("tracked: uses soft breaks within the cell paragraph", async () => {
    const p = await createTmpDoc("before table");
    await insertTable(p, -1, 1, 1, [["old"]]);
    await editTableCells(
      p,
      [{ blockIndex: 1, rowIndex: 0, colIndex: 0, newText: "n1\nn2" }],
      true,
    );
    const xml = await readRawDocXml(p);
    expect(xml).toContain("<w:br");
    expect(xml).toContain("w:ins");
    expect(xml).not.toContain("n1\nn2");
  });
});

describe("insert_paragraphs newline handling", () => {
  it("untracked: a multi-line insert becomes multiple paragraphs", async () => {
    const p = await createTmpDoc("anchor");
    await insertParagraphs(
      p,
      [{ text: "x\ny\nz", position: -1 }],
      false,
    );
    const xml = await readRawDocXml(p);
    expect(xml).not.toContain("x\ny\nz");
    const doc = await readDocument(p);
    // anchor + x + y + z
    expect(doc).toContain("Total blocks: 4");
    expect(doc).toContain("x");
    expect(doc).toContain("y");
    expect(doc).toContain("z");
  });

  it("tracked: a multi-line insert stays a single paragraph with soft breaks", async () => {
    const p = await createTmpDoc("anchor");
    await insertParagraphs(p, [{ text: "x\ny\nz", position: -1 }], true);
    const xml = await readRawDocXml(p);
    expect(xml).toContain("<w:br");
    expect(xml).toContain("w:ins");
    expect(xml).not.toContain("x\ny\nz");
  });
});
