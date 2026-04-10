import { describe, it, expect, afterEach } from "vitest";
import * as fs from "fs/promises";
import {
  createTmpDoc,
  cleanupTmpFiles,
  readRawDocXml,
  readRawHeaderXml,
  tmpDocxPath,
  trackTmpFile,
  createDocWithMoveTracking,
  createDocWithRPrChange,
  createDocWithPPrChange,
  createDocWithSectPrChange,
  createDocWithTrackedSdt,
  createDocWithTrackedHeader,
} from "./helpers.js";
import {
  createDocument,
  insertTable,
  getPageLayout,
  setPageLayout,
  acceptAllChanges,
  rejectAllChanges,
  readDocument,
  replaceText,
  editParagraph,
  insertParagraph,
  deleteParagraph,
  getDocumentInfo,
} from "../docx-engine.js";

afterEach(cleanupTmpFiles);

// =========================================================================
// createDocument
// =========================================================================

describe("createDocument", () => {
  it("creates a valid docx file", async () => {
    const p = tmpDocxPath();
    trackTmpFile(p);
    await createDocument(p);
    const stat = await fs.stat(p);
    expect(stat.size).toBeGreaterThan(0);
    // Should be readable
    const result = await readDocument(p);
    expect(result).toContain("Total blocks:");
  });

  it("creates document with title", async () => {
    const p = tmpDocxPath();
    trackTmpFile(p);
    await createDocument(p, undefined, "My Document");
    const doc = await readDocument(p);
    expect(doc).toContain("My Document");
    expect(doc).toContain("(H1)");
  });

  it("creates document with content", async () => {
    const p = tmpDocxPath();
    trackTmpFile(p);
    await createDocument(p, "Line 1\nLine 2\nLine 3");
    const doc = await readDocument(p);
    expect(doc).toContain("Line 1");
    expect(doc).toContain("Line 2");
    expect(doc).toContain("Line 3");
  });

  it("creates document with title and content", async () => {
    const p = tmpDocxPath();
    trackTmpFile(p);
    await createDocument(p, "Body text", "Title Text");
    const doc = await readDocument(p);
    expect(doc).toContain("Title Text");
    expect(doc).toContain("Body text");
  });

  it("creates parent directories if needed", async () => {
    const p = tmpDocxPath().replace(".docx", "/sub/dir/test.docx");
    trackTmpFile(p);
    await createDocument(p, "Nested");
    const stat = await fs.stat(p);
    expect(stat.size).toBeGreaterThan(0);
    // Clean up nested dirs
    await fs.rm(p.replace("/sub/dir/test.docx", ""), {
      recursive: true,
      force: true,
    });
  });

  it("produced docx has proper XML structure", async () => {
    const p = tmpDocxPath();
    trackTmpFile(p);
    await createDocument(p, "Test");
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:document");
    expect(xml).toContain("w:body");
    expect(xml).toContain("w:sectPr");
    expect(xml).toContain("w:pgSz");
    expect(xml).toContain("w:pgMar");
  });

  it("created document can be edited by other functions", async () => {
    const p = await createTmpDoc("Editable content");
    await replaceText(p, "content", "text", false, false);
    const doc = await readDocument(p);
    expect(doc).toContain("Editable text");
  });
});

// =========================================================================
// insertTable
// =========================================================================

describe("insertTable", () => {
  it("inserts a table at end of document", async () => {
    const p = await createTmpDoc("Before table");
    const result = await insertTable(p, -1, 2, 3);
    expect(result).toContain("2x3 table");
    const info = await getDocumentInfo(p);
    expect(info).toContain("Tables: 1");
  });

  it("inserts a table with data", async () => {
    const p = await createTmpDoc("Content");
    await insertTable(p, -1, 2, 2, [
      ["A1", "B1"],
      ["A2", "B2"],
    ]);
    const doc = await readDocument(p);
    expect(doc).toContain("A1");
    expect(doc).toContain("B1");
    expect(doc).toContain("A2");
    expect(doc).toContain("B2");
    expect(doc).toContain("[TABLE]");
  });

  it("inserts a table at specific position", async () => {
    const p = await createTmpDoc("Para 0\nPara 1");
    await insertTable(p, 1, 1, 2);
    const doc = await readDocument(p);
    // Table should be at index 1, pushing Para 1 to index 2
    expect(doc).toMatch(/\[0\].*Para 0/);
    expect(doc).toMatch(/\[1\].*\[TABLE\]/);
    expect(doc).toMatch(/\[2\].*Para 1/);
  });

  it("creates table with proper XML structure", async () => {
    const p = await createTmpDoc("Text");
    await insertTable(p, -1, 2, 3);
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:tbl");
    expect(xml).toContain("w:tr");
    expect(xml).toContain("w:tc");
    expect(xml).toContain("w:tblBorders");
    expect(xml).toContain("w:tblPr");
  });

  it("inserts empty table when no data provided", async () => {
    const p = await createTmpDoc("Text");
    await insertTable(p, -1, 3, 4);
    const xml = await readRawDocXml(p);
    // Should have 3 rows
    const trCount = (xml.match(/w:tr/g) || []).length;
    // Each w:tr appears at least twice (open/close or self-closing usage)
    expect(trCount).toBeGreaterThanOrEqual(3);
  });
});

// =========================================================================
// getPageLayout
// =========================================================================

describe("getPageLayout", () => {
  it("returns page size and margins for default document", async () => {
    const p = await createTmpDoc("Test");
    const result = await getPageLayout(p);
    expect(result).toContain("Page size:");
    expect(result).toContain("Margins (mm):");
    expect(result).toContain("Orientation:");
    expect(result).toContain("portrait");
  });

  it("detects A4 page size preset", async () => {
    const p = await createTmpDoc("A4 doc");
    const result = await getPageLayout(p);
    // Default createDocument uses A4 dimensions (11906 x 16838 twips)
    expect(result).toContain("A4");
  });

  it("shows available presets", async () => {
    const p = await createTmpDoc("Test");
    const result = await getPageLayout(p);
    expect(result).toContain("Available page size presets:");
    expect(result).toContain("Available margin presets:");
  });
});

// =========================================================================
// setPageLayout
// =========================================================================

describe("setPageLayout", () => {
  it("sets page size to LETTER", async () => {
    const p = await createTmpDoc("Test");
    const result = await setPageLayout(p, { pageSizePreset: "LETTER" });
    expect(result).toContain("Letter");
    const xml = await readRawDocXml(p);
    expect(xml).toContain('w:w="12240"');
    expect(xml).toContain('w:h="15840"');
  });

  it("sets page orientation to landscape", async () => {
    const p = await createTmpDoc("Test");
    await setPageLayout(p, { orientation: "landscape" });
    const xml = await readRawDocXml(p);
    expect(xml).toContain('w:orient="landscape"');
  });

  it("sets custom page dimensions in mm", async () => {
    const p = await createTmpDoc("Test");
    await setPageLayout(p, { widthMm: 200, heightMm: 300 });
    const layout = await getPageLayout(p);
    expect(layout).toContain("200");
    expect(layout).toContain("300");
  });

  it("sets margin preset to NARROW", async () => {
    const p = await createTmpDoc("Test");
    const result = await setPageLayout(p, { marginPreset: "NARROW" });
    expect(result).toContain("Narrow");
    const xml = await readRawDocXml(p);
    expect(xml).toContain('w:top="720"');
    expect(xml).toContain('w:right="720"');
  });

  it("sets custom margins in mm", async () => {
    const p = await createTmpDoc("Test");
    await setPageLayout(p, { topMm: 30, bottomMm: 30 });
    const layout = await getPageLayout(p);
    expect(layout).toContain("30");
  });

  it("throws INVALID_PARAMETER for unknown page size preset", async () => {
    const p = await createTmpDoc("Test");
    await expect(setPageLayout(p, { pageSizePreset: "UNKNOWN" })).rejects.toMatchObject({
      code: "INVALID_PARAMETER",
    });
  });

  it("throws INVALID_PARAMETER for unknown margin preset", async () => {
    const p = await createTmpDoc("Test");
    await expect(setPageLayout(p, { marginPreset: "NOPE" })).rejects.toMatchObject({
      code: "INVALID_PARAMETER",
    });
  });

  it("sets header and footer distances", async () => {
    const p = await createTmpDoc("Test");
    await setPageLayout(p, { headerMm: 15, footerMm: 15 });
    const layout = await getPageLayout(p);
    expect(layout).toContain("15");
  });

  it("returns message when no changes specified", async () => {
    const p = await createTmpDoc("Test");
    const result = await setPageLayout(p, {});
    expect(result).toContain("No page layout changes");
  });

  it("sets gutter margin", async () => {
    const p = await createTmpDoc("Test");
    await setPageLayout(p, { gutterMm: 10 });
    const layout = await getPageLayout(p);
    expect(layout).toContain("10");
  });

  it("LETTER landscape swaps dimensions", async () => {
    const p = await createTmpDoc("Test");
    await setPageLayout(p, {
      pageSizePreset: "LETTER",
      orientation: "landscape",
    });
    const xml = await readRawDocXml(p);
    // Landscape swaps: w becomes h, h becomes w
    expect(xml).toContain('w:w="15840"');
    expect(xml).toContain('w:h="12240"');
    expect(xml).toContain('w:orient="landscape"');
  });
});

// =========================================================================
// acceptAllChanges
// =========================================================================

describe("acceptAllChanges", () => {
  it("removes w:del elements and unwraps w:ins", async () => {
    const p = await createTmpDoc("Hello old world");
    await replaceText(p, "old", "new", false, true);

    // Before accept: should have del/ins
    let xml = await readRawDocXml(p);
    expect(xml).toContain("w:del");
    expect(xml).toContain("w:ins");

    await acceptAllChanges(p);

    // After accept: no more del/ins
    xml = await readRawDocXml(p);
    expect(xml).not.toContain("w:del");
    expect(xml).not.toContain("w:ins");
    expect(xml).not.toContain("w:delText");
  });

  it("accepted text becomes permanent", async () => {
    const p = await createTmpDoc("Replace me");
    await replaceText(p, "me", "you", false, true);
    await acceptAllChanges(p);
    const doc = await readDocument(p);
    expect(doc).toContain("Replace you");
    // "Replace me" should no longer appear (the old text is gone)
    expect(doc).not.toContain("Replace me");
  });

  it("accepts tracked paragraph edit", async () => {
    const p = await createTmpDoc("Old text");
    await editParagraph(p, 0, "New text", true);
    await acceptAllChanges(p);
    const doc = await readDocument(p);
    expect(doc).toContain("New text");
    const xml = await readRawDocXml(p);
    expect(xml).not.toContain("w:del");
    expect(xml).not.toContain("w:ins");
  });

  it("accepts tracked paragraph insert", async () => {
    const p = await createTmpDoc("Existing");
    await insertParagraph(p, "Inserted", -1, undefined, true);
    await acceptAllChanges(p);
    const doc = await readDocument(p);
    expect(doc).toContain("Inserted");
    const xml = await readRawDocXml(p);
    expect(xml).not.toContain("w:ins");
  });

  it("accepts tracked paragraph deletion (removes deleted content)", async () => {
    const p = await createTmpDoc("Keep this\nDelete this");
    await deleteParagraph(p, 1, true);
    await acceptAllChanges(p);
    const doc = await readDocument(p);
    expect(doc).toContain("Keep this");
    // The deleted paragraph's text should be gone
    const xml = await readRawDocXml(p);
    expect(xml).not.toContain("w:del");
    expect(xml).not.toContain("w:delText");
  });

  it("cleans pPr > rPr > w:ins/w:del markers", async () => {
    const p = await createTmpDoc("To delete");
    await deleteParagraph(p, 0, true);
    let xml = await readRawDocXml(p);
    // Before: should have pPr > rPr > w:del marker
    expect(xml).toContain("w:del");

    await acceptAllChanges(p);
    xml = await readRawDocXml(p);
    expect(xml).not.toContain("w:del");
    expect(xml).not.toContain("w:ins");
  });
});

// =========================================================================
// rejectAllChanges
// =========================================================================

describe("rejectAllChanges", () => {
  it("removes w:ins elements and unwraps w:del (restores original)", async () => {
    const p = await createTmpDoc("Hello old world");
    await replaceText(p, "old", "new", false, true);

    await rejectAllChanges(p);

    const xml = await readRawDocXml(p);
    expect(xml).not.toContain("w:del");
    expect(xml).not.toContain("w:ins");
    expect(xml).not.toContain("w:delText");
  });

  it("restores original text after reject", async () => {
    const p = await createTmpDoc("Original text");
    await replaceText(p, "Original", "Changed", false, true);
    await rejectAllChanges(p);
    const doc = await readDocument(p);
    expect(doc).toContain("Original");
    expect(doc).not.toContain("Changed");
  });

  it("converts w:delText back to w:t", async () => {
    const p = await createTmpDoc("Delete this word");
    await replaceText(p, "this", "that", false, true);

    // Before reject: has w:delText
    let xml = await readRawDocXml(p);
    expect(xml).toContain("w:delText");

    await rejectAllChanges(p);

    // After reject: w:delText converted to w:t
    xml = await readRawDocXml(p);
    expect(xml).not.toContain("w:delText");
    expect(xml).toContain("w:t");
  });

  it("rejects tracked paragraph edit (restores old content)", async () => {
    const p = await createTmpDoc("Original paragraph");
    await editParagraph(p, 0, "Modified paragraph", true);
    await rejectAllChanges(p);
    const doc = await readDocument(p);
    expect(doc).toContain("Original paragraph");
    expect(doc).not.toContain("Modified paragraph");
  });

  it("rejects tracked paragraph insert (removes it)", async () => {
    const p = await createTmpDoc("Existing only");
    await insertParagraph(p, "Should disappear", -1, undefined, true);
    await rejectAllChanges(p);
    const doc = await readDocument(p);
    expect(doc).toContain("Existing only");
    expect(doc).not.toContain("Should disappear");
  });

  it("rejects tracked paragraph deletion (restores deleted text)", async () => {
    const p = await createTmpDoc("Keep me\nRestore me");
    await deleteParagraph(p, 1, true);

    // Before reject: deleted text is hidden in default view
    let doc = await readDocument(p, undefined, undefined, false);
    // After reject: text is restored
    await rejectAllChanges(p);
    doc = await readDocument(p);
    expect(doc).toContain("Restore me");
  });

  it("cleans pPr > rPr > w:ins/w:del markers", async () => {
    const p = await createTmpDoc("Tracked insert");
    await insertParagraph(p, "New para", -1, undefined, true);
    await rejectAllChanges(p);
    const xml = await readRawDocXml(p);
    expect(xml).not.toContain("w:ins");
    expect(xml).not.toContain("w:del");
  });
});

// =========================================================================
// acceptAllChanges — move tracking (w:moveFrom / w:moveTo)
// =========================================================================

describe("acceptAllChanges — move tracking", () => {
  it("removes w:moveFrom and unwraps w:moveTo", async () => {
    const p = await createDocWithMoveTracking("Hello ", "world", " end");
    let xml = await readRawDocXml(p);
    expect(xml).toContain("w:moveFrom");
    expect(xml).toContain("w:moveTo");

    await acceptAllChanges(p);

    xml = await readRawDocXml(p);
    expect(xml).not.toContain("w:moveFrom");
    expect(xml).not.toContain("w:moveTo");
    expect(xml).not.toContain("w:moveFromRangeStart");
    expect(xml).not.toContain("w:moveFromRangeEnd");
    expect(xml).not.toContain("w:moveToRangeStart");
    expect(xml).not.toContain("w:moveToRangeEnd");
  });

  it("keeps moved text at destination after accept", async () => {
    const p = await createDocWithMoveTracking("Hello ", "world", " end");
    await acceptAllChanges(p);
    const doc = await readDocument(p);
    // The moved text should appear at the destination (second paragraph)
    expect(doc).toContain("world");
  });
});

// =========================================================================
// rejectAllChanges — move tracking
// =========================================================================

describe("rejectAllChanges — move tracking", () => {
  it("removes w:moveTo and unwraps w:moveFrom", async () => {
    const p = await createDocWithMoveTracking("Hello ", "world", " end");
    await rejectAllChanges(p);

    const xml = await readRawDocXml(p);
    expect(xml).not.toContain("w:moveFrom");
    expect(xml).not.toContain("w:moveTo");
    expect(xml).not.toContain("w:moveFromRangeStart");
    expect(xml).not.toContain("w:moveToRangeStart");
  });

  it("restores text at original position after reject", async () => {
    const p = await createDocWithMoveTracking("Hello ", "world", " end");
    await rejectAllChanges(p);
    const doc = await readDocument(p);
    // The moved text should stay at original location
    expect(doc).toContain("world");
  });
});

// =========================================================================
// acceptAllChanges — w:rPrChange (formatting change tracking)
// =========================================================================

describe("acceptAllChanges — rPrChange", () => {
  it("strips w:rPrChange from run properties", async () => {
    const p = await createDocWithRPrChange("Bold text");
    let xml = await readRawDocXml(p);
    expect(xml).toContain("w:rPrChange");

    await acceptAllChanges(p);

    xml = await readRawDocXml(p);
    expect(xml).not.toContain("w:rPrChange");
    // The current formatting (bold) should remain
    expect(xml).toContain("w:b");
  });
});

// =========================================================================
// rejectAllChanges — w:rPrChange
// =========================================================================

describe("rejectAllChanges — rPrChange", () => {
  it("strips w:rPrChange and restores old formatting", async () => {
    const p = await createDocWithRPrChange("Was plain text");
    await rejectAllChanges(p);

    const xml = await readRawDocXml(p);
    expect(xml).not.toContain("w:rPrChange");
    // The bold formatting should be reverted (old rPr was empty)
    expect(xml).not.toContain("<w:b/>");
    expect(xml).not.toContain("<w:b>");
  });
});

// =========================================================================
// acceptAllChanges — w:pPrChange (paragraph property change tracking)
// =========================================================================

describe("acceptAllChanges — pPrChange", () => {
  it("strips w:pPrChange from paragraph properties", async () => {
    const p = await createDocWithPPrChange("Centered text");
    let xml = await readRawDocXml(p);
    expect(xml).toContain("w:pPrChange");

    await acceptAllChanges(p);

    xml = await readRawDocXml(p);
    expect(xml).not.toContain("w:pPrChange");
    // Current alignment (center) should remain
    expect(xml).toContain("center");
  });
});

// =========================================================================
// rejectAllChanges — w:pPrChange
// =========================================================================

describe("rejectAllChanges — pPrChange", () => {
  it("restores old paragraph properties from pPrChange", async () => {
    const p = await createDocWithPPrChange("Was left-aligned");
    await rejectAllChanges(p);

    const xml = await readRawDocXml(p);
    expect(xml).not.toContain("w:pPrChange");
    // Alignment should be reverted to left
    expect(xml).toContain("left");
    expect(xml).not.toContain("center");
  });
});

// =========================================================================
// acceptAllChanges — w:sectPrChange
// =========================================================================

describe("acceptAllChanges — sectPrChange", () => {
  it("strips w:sectPrChange from section properties", async () => {
    const p = await createDocWithSectPrChange("Some text");
    let xml = await readRawDocXml(p);
    expect(xml).toContain("w:sectPrChange");

    await acceptAllChanges(p);

    xml = await readRawDocXml(p);
    expect(xml).not.toContain("w:sectPrChange");
  });
});

// =========================================================================
// acceptAllChanges — tracked changes inside w:sdt
// =========================================================================

describe("acceptAllChanges — SDT container", () => {
  it("accepts tracked changes inside structured document tags", async () => {
    const p = await createDocWithTrackedSdt("old text", "new text");
    let xml = await readRawDocXml(p);
    expect(xml).toContain("w:del");
    expect(xml).toContain("w:ins");

    await acceptAllChanges(p);

    xml = await readRawDocXml(p);
    expect(xml).not.toContain("w:del");
    expect(xml).not.toContain("w:ins");
    const doc = await readDocument(p);
    expect(doc).toContain("new text");
    expect(doc).not.toContain("old text");
  });
});

// =========================================================================
// acceptAllChanges — tracked changes in headers/footers
// =========================================================================

describe("acceptAllChanges — headers/footers", () => {
  it("accepts tracked changes in headers", async () => {
    const p = await createDocWithTrackedHeader("Body text", "old header", "new header");
    let hdrXml = await readRawHeaderXml(p);
    expect(hdrXml).toContain("w:del");
    expect(hdrXml).toContain("w:ins");

    await acceptAllChanges(p);

    hdrXml = await readRawHeaderXml(p);
    expect(hdrXml).not.toContain("w:del");
    expect(hdrXml).not.toContain("w:ins");
  });
});
