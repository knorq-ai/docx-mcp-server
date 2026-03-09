import { describe, it, expect, afterEach } from "vitest";
import {
  createTmpDoc,
  cleanupTmpFiles,
  readRawDocXml,
  createDocWithBookmark,
  createDocWithInsertedRun,
  createDocWithCommentRange,
  createDocWithDrawing,
  createDocWithSdt,
  createDocWithNestedTable,
} from "./helpers.js";
import {
  editParagraph,
  readDocument,
  formatText,
  replaceText,
} from "../docx-engine.js";

afterEach(cleanupTmpFiles);

// =========================================================================
// Minimal diff: editParagraph only marks the changed portion
// =========================================================================

describe("editParagraph produces minimal tracked-change diff", () => {
  it("changing one number marks only that character in del/ins", async () => {
    const p = await createTmpDoc("Pay per Article 5 of this agreement.");
    await editParagraph(p, 0, "Pay per Article 6 of this agreement.", true);
    const xml = await readRawDocXml(p);
    // Del should contain only "5", not the whole sentence
    expect(xml).toMatch(/<w:delText[^>]*>5<\/w:delText>/);
    // Ins should contain only "6", not the whole sentence
    expect(xml).toMatch(/<w:t[^>]*>6<\/w:t>/);
    // Common text should NOT be inside del/ins tags
    expect(xml).toContain("Pay per Article");
    // Verify the del tag does NOT contain the full sentence
    const delMatch = xml.match(/<w:del\b[^>]*>([\s\S]*?)<\/w:del>/);
    expect(delMatch?.[0]).not.toContain("Pay per Article");
  });

  it("unchanged paragraph produces no del/ins", async () => {
    const p = await createTmpDoc("Same text");
    await editParagraph(p, 0, "Same text", true);
    const xml = await readRawDocXml(p);
    expect(xml).not.toContain("w:del");
    expect(xml).not.toContain("w:ins");
  });

  it("prefix and suffix are preserved as plain runs", async () => {
    const p = await createTmpDoc("From 14 to 21 days.");
    await editParagraph(p, 0, "From 21 to 21 days.", true);
    const xml = await readRawDocXml(p);
    // "14" deleted, "21" inserted; "From " prefix and " to 21 days." suffix are plain
    expect(xml).toMatch(/<w:delText[^>]*>14<\/w:delText>/);
    // Verify the del block contains only "14", not "From"
    const delBlock = xml.match(/<w:del\b[^>]*>[\s\S]*?<\/w:del>/)?.[0] ?? "";
    expect(delBlock).not.toContain("From");
    expect(delBlock).toContain("14");
  });
});

// =========================================================================
// Fix 1: editParagraph preserves structural elements
// =========================================================================

describe("editParagraph preserves structural elements", () => {
  it("preserves w:bookmarkStart and w:bookmarkEnd", async () => {
    const p = await createDocWithBookmark("Original text", "clause1");
    await editParagraph(p, 0, "Updated text", false);
    const xml = await readRawDocXml(p);
    expect(xml).toContain('w:name="clause1"');
    expect(xml).toContain("w:bookmarkStart");
    expect(xml).toContain("w:bookmarkEnd");
  });

  it("preserves w:bookmarkStart/End with track changes", async () => {
    const p = await createDocWithBookmark("Original text", "sigBlock");
    await editParagraph(p, 0, "New text", true);
    const xml = await readRawDocXml(p);
    expect(xml).toContain('w:name="sigBlock"');
    expect(xml).toContain("w:bookmarkStart");
    expect(xml).toContain("w:bookmarkEnd");
    expect(xml).toContain("w:delText");
    expect(xml).toContain("w:ins");
  });

  it("preserves w:commentRangeStart and w:commentRangeEnd", async () => {
    const p = await createDocWithCommentRange("Contract text");
    await editParagraph(p, 0, "Revised text", false);
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:commentRangeStart");
    expect(xml).toContain("w:commentRangeEnd");
  });

  it("preserves w:drawing (inline images) when editing text", async () => {
    const p = await createDocWithDrawing();
    await editParagraph(p, 0, "New text with image preserved", false);
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:drawing");
    expect(xml).toContain("wp:inline");
  });
});

// =========================================================================
// Fix 2: formatText and replaceText see text inside w:ins runs
// =========================================================================

describe("formatText finds text inside w:ins tracked insertions", () => {
  it("formats text that lives inside a w:ins run", async () => {
    // Doc: "Hello " + w:ins("world") + " end"
    const p = await createDocWithInsertedRun("Hello ", "world", " end");
    const result = await formatText(p, "world", { bold: true });
    expect(result).toContain("1 occurrence");
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:b");
  });

  it("formats text spanning a normal run and a w:ins run", async () => {
    // "Hello wo" + w:ins("rld") — "world" spans the boundary
    const p = await createDocWithInsertedRun("Hello wo", "rld", " end");
    const result = await formatText(p, "world", { italic: true });
    expect(result).toContain("1 occurrence");
  });

  it("full paragraph text includes w:ins text", async () => {
    const p = await createDocWithInsertedRun("Foo ", "Bar", " Baz");
    const doc = await readDocument(p);
    expect(doc).toContain("Foo Bar Baz");
  });
});

describe("replaceText finds and replaces text inside w:ins runs", () => {
  it("replaces text that lives inside a w:ins run (no track changes)", async () => {
    const p = await createDocWithInsertedRun("Hello ", "old value", " end");
    const result = await replaceText(p, "old value", "new value", false, false);
    expect(result).toContain("1 occurrence");
    const doc = await readDocument(p);
    expect(doc).toContain("new value");
  });
});

// =========================================================================
// Fix 4: w:sdt (content controls) visible in block enumeration
// =========================================================================

describe("enumerateBlocks includes w:sdt content controls", () => {
  it("reads text inside a w:sdt content control", async () => {
    const p = await createDocWithSdt("Counterparty Name Ltd", "partyName");
    const doc = await readDocument(p);
    expect(doc).toContain("Counterparty Name Ltd");
  });

  it("includes sdt block in total block count", async () => {
    const p = await createDocWithSdt("SDT Content");
    const doc = await readDocument(p);
    // Should have at least 2 blocks: normal paragraph + sdt
    expect(doc).toMatch(/Total blocks: [2-9]/);
  });
});

// =========================================================================
// Fix 8: nested tables scanned by replaceText / formatText
// =========================================================================

describe("replaceText and formatText scan nested tables", () => {
  it("replaceText finds text inside nested table cell", async () => {
    const p = await createDocWithNestedTable();
    const result = await replaceText(p, "nested text", "replaced text", false, false);
    expect(result).toContain("1 occurrence");
    const doc = await readDocument(p);
    expect(doc).toContain("replaced text");
  });

  it("formatText finds text inside nested table cell", async () => {
    const p = await createDocWithNestedTable();
    const result = await formatText(p, "nested text", { bold: true });
    expect(result).toContain("1 occurrence");
  });
});
