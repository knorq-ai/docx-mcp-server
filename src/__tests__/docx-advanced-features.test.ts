import { describe, it, expect, afterEach } from "vitest";
import {
  createTmpDoc,
  cleanupTmpFiles,
  readRawDocXml,
  createDocWithHeaderFooter,
  createDocWithFootnote,
} from "./helpers.js";
import {
  readHeaderFooter,
  readFootnotes,
  editTableCell,
  editParagraph,
  replaceText,
  readDocument,
  insertTable,
} from "../docx-engine.js";

afterEach(cleanupTmpFiles);

// =========================================================================
// Fix 3: Headers and footers
// =========================================================================

describe("readHeaderFooter", () => {
  it("returns header content", async () => {
    const p = await createDocWithHeaderFooter("Body text", "CONFIDENTIAL", "Page 1");
    const result = await readHeaderFooter(p);
    expect(result).toContain("CONFIDENTIAL");
  });

  it("returns footer content", async () => {
    const p = await createDocWithHeaderFooter("Body text", "Header", "Footer text here");
    const result = await readHeaderFooter(p);
    expect(result).toContain("Footer text here");
  });

  it("returns message when no header/footer exists", async () => {
    const p = await createTmpDoc("Just body text");
    const result = await readHeaderFooter(p);
    expect(result).toContain("No headers or footers");
  });
});

describe("replaceText with includeHeadersFooters option", () => {
  it("replaces text in header when includeHeadersFooters=true", async () => {
    const p = await createDocWithHeaderFooter("Body", "DRAFT Contract", "Page 1");
    const result = await replaceText(p, "DRAFT", "FINAL", false, false, "Claude", true);
    expect(result).toContain("1 occurrence");
    const hf = await readHeaderFooter(p);
    expect(hf).toContain("FINAL Contract");
    expect(hf).not.toContain("DRAFT Contract");
  });

  it("does NOT replace in header when includeHeadersFooters=false (default)", async () => {
    const p = await createDocWithHeaderFooter("Body", "DRAFT Header", "Footer");
    await replaceText(p, "DRAFT", "FINAL", false, false);
    const hf = await readHeaderFooter(p);
    expect(hf).toContain("DRAFT Header"); // unchanged
  });
});

// =========================================================================
// Fix 6: editTableCell
// =========================================================================

describe("editTableCell", () => {
  it("edits a specific cell by row and col index", async () => {
    const p = await createTmpDoc("Before table");
    await insertTable(p, -1, 2, 2, [
      ["A1", "B1"],
      ["A2", "B2"],
    ]);
    await editTableCell(p, 1, 0, 1, "Updated B1", false);
    const doc = await readDocument(p);
    expect(doc).toContain("Updated B1");
  });

  it("editTableCell with track changes wraps in del/ins", async () => {
    const p = await createTmpDoc("Paragraph");
    await insertTable(p, -1, 1, 1, [["Original"]]);
    await editTableCell(p, 1, 0, 0, "Changed", true);
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:delText");
    expect(xml).toContain("w:ins");
  });

  it("editTableCell tracked changes survive a subsequent editParagraph round-trip", async () => {
    const p = await createTmpDoc("Paragraph text");
    await insertTable(p, -1, 1, 1, [["Original cell"]]);
    // Step 1: edit table cell with tracked changes
    await editTableCell(p, 1, 0, 0, "Changed cell", true);
    // Verify tracked changes are present after step 1
    const xmlAfterTableEdit = await readRawDocXml(p);
    expect(xmlAfterTableEdit).toContain("w:delText");
    expect(xmlAfterTableEdit).toContain("w:ins");

    // Step 2: edit paragraph (triggers parse -> serialize round-trip)
    await editParagraph(p, 0, "Updated paragraph", true);

    // Verify table cell tracked changes survive the round-trip
    const xmlAfterBothEdits = await readRawDocXml(p);
    // The table should still contain w:del with "Original" (delText) and w:ins with "Changed"
    expect(xmlAfterBothEdits).toContain("w:delText");
    // Check revision view shows both tracked changes
    const rev = await readDocument(p, undefined, undefined, true);
    expect(rev).toContain("[-Original-]");
    expect(rev).toContain("[+Changed+]");
    expect(rev).toContain("[-Paragraph text-]");
    expect(rev).toContain("[+Updated paragraph+]");
  });

  it("throws INDEX_OUT_OF_RANGE for bad block index", async () => {
    const p = await createTmpDoc("Paragraph");
    await expect(editTableCell(p, 99, 0, 0, "text", false)).rejects.toMatchObject({
      code: "INDEX_OUT_OF_RANGE",
    });
  });

  it("throws NOT_A_TABLE error when block is not a table", async () => {
    const p = await createTmpDoc("Just a paragraph");
    await expect(editTableCell(p, 0, 0, 0, "text", false)).rejects.toMatchObject({
      code: "NOT_A_TABLE",
    });
  });

  it("throws INDEX_OUT_OF_RANGE for bad row index", async () => {
    const p = await createTmpDoc("Before");
    await insertTable(p, -1, 2, 2);
    await expect(editTableCell(p, 1, 99, 0, "text", false)).rejects.toMatchObject({
      code: "INDEX_OUT_OF_RANGE",
    });
  });

  it("throws INDEX_OUT_OF_RANGE for bad col index", async () => {
    const p = await createTmpDoc("Before");
    await insertTable(p, -1, 2, 2);
    await expect(editTableCell(p, 1, 0, 99, "text", false)).rejects.toMatchObject({
      code: "INDEX_OUT_OF_RANGE",
    });
  });
});

// =========================================================================
// Fix 7: readFootnotes
// =========================================================================

describe("readFootnotes", () => {
  it("returns footnote content", async () => {
    const p = await createDocWithFootnote("See note below", "This is the footnote text");
    const result = await readFootnotes(p);
    expect(result).toContain("This is the footnote text");
  });

  it("returns message when no footnotes exist", async () => {
    const p = await createTmpDoc("No footnotes here");
    const result = await readFootnotes(p);
    expect(result).toContain("No footnotes");
  });
});
