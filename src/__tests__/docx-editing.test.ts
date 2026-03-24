import { describe, it, expect, afterEach } from "vitest";
import {
  createTmpDoc,
  cleanupTmpFiles,
  readRawDocXml,
  createCrossRunDoc,
  createDocWithInlineSdt,
  createDocWithNumberedParagraph,
  createDocWithTrackedPPr,
} from "./helpers.js";
import {
  readDocument,
  replaceText,
  editParagraph,
  insertParagraph,
  deleteParagraph,
  deleteParagraphs,
  acceptAllChanges,
  rejectAllChanges,
} from "../docx-engine.js";

afterEach(cleanupTmpFiles);

// =========================================================================
// replaceText — untracked
// =========================================================================

describe("replaceText (untracked)", () => {
  it("replaces text in a single paragraph", async () => {
    const p = await createTmpDoc("Hello old world");
    const result = await replaceText(p, "old", "new", false, false);
    expect(result).toContain("1 occurrence");
    const doc = await readDocument(p);
    expect(doc).toContain("Hello new world");
  });

  it("replaces multiple occurrences", async () => {
    const p = await createTmpDoc("foo bar foo baz foo");
    const result = await replaceText(p, "foo", "qux", false, false);
    expect(result).toContain("3 occurrence");
    const doc = await readDocument(p);
    expect(doc).toContain("qux bar qux baz qux");
  });

  it("performs case-insensitive replacement by default", async () => {
    const p = await createTmpDoc("Hello HELLO hello");
    const result = await replaceText(p, "hello", "Hi", false, false);
    expect(result).toContain("3 occurrence");
  });

  it("performs case-sensitive replacement when requested", async () => {
    const p = await createTmpDoc("Hello HELLO hello");
    const result = await replaceText(p, "Hello", "Hi", true, false);
    expect(result).toContain("1 occurrence");
    const doc = await readDocument(p);
    expect(doc).toContain("Hi");
    expect(doc).toContain("HELLO");
    expect(doc).toContain("hello");
  });

  it("reports no occurrences when text is absent", async () => {
    const p = await createTmpDoc("Hello world");
    const result = await replaceText(p, "xyz", "abc", false, false);
    expect(result).toContain("No occurrences");
  });

  it("handles cross-run text replacement", async () => {
    const p = await createCrossRunDoc(["Hel", "lo Wor", "ld"]);
    const result = await replaceText(p, "Hello World", "Hi", false, false);
    expect(result).toContain("1 occurrence");
    const doc = await readDocument(p);
    expect(doc).toContain("Hi");
    expect(doc).not.toContain("Hello World");
  });
});

// =========================================================================
// replaceText — tracked
// =========================================================================

describe("replaceText (tracked)", () => {
  it("creates w:del and w:ins markup", async () => {
    const p = await createTmpDoc("Hello old world");
    await replaceText(p, "old", "new", false, true, "TestAuthor");
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:del");
    expect(xml).toContain("w:ins");
    expect(xml).toContain("w:delText");
    expect(xml).toContain("TestAuthor");
  });

  it("preserves original text in w:delText", async () => {
    const p = await createTmpDoc("Replace me please");
    await replaceText(p, "me", "you", false, true);
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:delText");
    // The accepted view should show the new text
    const doc = await readDocument(p);
    expect(doc).toContain("Replace you please");
  });

  it("sets author and date attributes on revisions", async () => {
    const p = await createTmpDoc("Some text");
    await replaceText(p, "text", "content", false, true, "Alice");
    const xml = await readRawDocXml(p);
    expect(xml).toContain("Alice");
    // Date should be an ISO string
    expect(xml).toMatch(/w:date="[^"]+T[^"]+"/);
  });

  it("handles cross-run tracked replacement", async () => {
    const p = await createCrossRunDoc(["Hel", "lo Wor", "ld"]);
    await replaceText(p, "Hello World", "Hi Earth", false, true);
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:del");
    expect(xml).toContain("w:ins");
    const doc = await readDocument(p);
    expect(doc).toContain("Hi Earth");
  });

  it("default author is Claude", async () => {
    const p = await createTmpDoc("Test text");
    await replaceText(p, "text", "data", false, true);
    const xml = await readRawDocXml(p);
    expect(xml).toContain("Claude");
  });

  it("shows revision annotations in show_revisions mode", async () => {
    const p = await createTmpDoc("Before change");
    await replaceText(p, "change", "modification", false, true);
    const result = await readDocument(p, undefined, undefined, true);
    expect(result).toContain("[-change-]");
    expect(result).toContain("[+modification+]");
  });
});

// =========================================================================
// replaceText — inline SDT (Google Docs export pattern)
// =========================================================================

describe("replaceText with inline w:sdt", () => {
  it("reads text from inline SDT", async () => {
    const p = await createDocWithInlineSdt("Hello SDT world");
    const doc = await readDocument(p);
    expect(doc).toContain("Hello SDT world");
  });

  it("replaces text inside inline SDT (untracked)", async () => {
    const p = await createDocWithInlineSdt("Hello SDT world");
    const result = await replaceText(p, "SDT", "replaced", false, false);
    expect(result).toContain("1 occurrence");
    const doc = await readDocument(p);
    expect(doc).toContain("Hello replaced world");
  });

  it("replaces text inside inline SDT (tracked)", async () => {
    const p = await createDocWithInlineSdt("Hello SDT world");
    await replaceText(p, "SDT", "tracked", false, true, "TestAuthor");
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:del");
    expect(xml).toContain("w:ins");
    // SDT wrapper should still be present
    expect(xml).toContain("w:sdtContent");
    const doc = await readDocument(p);
    expect(doc).toContain("Hello tracked world");
  });

  it("tracked replacement preserves SDT structure", async () => {
    const p = await createDocWithInlineSdt("Replace me here");
    await replaceText(p, "me", "you", false, true);
    const xml = await readRawDocXml(p);
    // sdtPr with tag should still exist
    expect(xml).toContain("goog_rdk_0");
    expect(xml).toContain("w:sdtPr");
  });
});

// =========================================================================
// editParagraph
// =========================================================================

describe("editParagraph", () => {
  it("replaces paragraph content (untracked)", async () => {
    const p = await createTmpDoc("Original text\nSecond para");
    const result = await editParagraph(p, 0, "New text", false);
    expect(result).toContain("Updated paragraph 0");
    const doc = await readDocument(p);
    expect(doc).toContain("New text");
    expect(doc).not.toContain("Original text");
  });

  it("replaces paragraph content (tracked) with w:del and w:ins", async () => {
    const p = await createTmpDoc("Old content here");
    await editParagraph(p, 0, "New content here", true, "Editor");
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:del");
    expect(xml).toContain("w:ins");
    expect(xml).toContain("Editor");
    const doc = await readDocument(p);
    expect(doc).toContain("New content here");
  });

  it("preserves paragraph style after edit", async () => {
    const p = await createTmpDoc("Body", "Title");
    // Edit the title (index 0 which has Heading1 style)
    await editParagraph(p, 0, "Updated Title", false);
    const xml = await readRawDocXml(p);
    expect(xml).toContain("Heading1");
    const doc = await readDocument(p);
    expect(doc).toContain("Updated Title");
  });

  it("throws INDEX_OUT_OF_RANGE for out-of-range index", async () => {
    const p = await createTmpDoc("Solo paragraph");
    await expect(editParagraph(p, 99, "New text", false)).rejects.toMatchObject({
      code: "INDEX_OUT_OF_RANGE",
    });
  });

  it("tracked edit shows revisions correctly", async () => {
    const p = await createTmpDoc("Before editing");
    await editParagraph(p, 0, "After editing", true);
    const rev = await readDocument(p, undefined, undefined, true);
    // Minimal diff: "Before" → "After", common suffix " editing" is plain text
    expect(rev).toContain("[-Before-]");
    expect(rev).toContain("[+After+]");
    expect(rev).toContain("editing");
  });
});

// =========================================================================
// insertParagraph
// =========================================================================

describe("insertParagraph", () => {
  it("inserts at position 0 (untracked)", async () => {
    const p = await createTmpDoc("Existing paragraph");
    await insertParagraph(p, "Inserted first", 0, undefined, false);
    const doc = await readDocument(p);
    expect(doc).toMatch(/\[0\].*Inserted first/);
    expect(doc).toMatch(/\[1\].*Existing paragraph/);
  });

  it("inserts at end with position=-1 (untracked)", async () => {
    const p = await createTmpDoc("First paragraph");
    await insertParagraph(p, "Last paragraph", -1, undefined, false);
    const doc = await readDocument(p);
    expect(doc).toContain("First paragraph");
    expect(doc).toContain("Last paragraph");
  });

  it("inserts with style", async () => {
    const p = await createTmpDoc("Body text");
    await insertParagraph(p, "Section Title", 0, "Heading2", false);
    const xml = await readRawDocXml(p);
    expect(xml).toContain("Heading2");
    const doc = await readDocument(p);
    expect(doc).toContain("(H2)");
  });

  it("inserts tracked paragraph with w:ins wrapper", async () => {
    const p = await createTmpDoc("Existing content");
    await insertParagraph(p, "New tracked para", -1, undefined, true, "Bob");
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:ins");
    expect(xml).toContain("Bob");
    const doc = await readDocument(p);
    expect(doc).toContain("New tracked para");
  });

  it("tracked insert has pPr rPr ins marker", async () => {
    const p = await createTmpDoc("Content");
    await insertParagraph(p, "Tracked", -1, undefined, true);
    const xml = await readRawDocXml(p);
    // Should have w:ins inside w:rPr inside w:pPr (paragraph break marker)
    expect(xml).toContain("w:pPr");
    expect(xml).toContain("w:rPr");
  });

  it("inserts with numId and numLevel (untracked)", async () => {
    const p = await createTmpDoc("Existing");
    await insertParagraph(p, "Numbered item", -1, undefined, false, "Claude", 14, 0);
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:numPr");
    expect(xml).toContain("w:numId");
    expect(xml).toMatch(/w:numId[^>]*w:val="14"/);
    expect(xml).toMatch(/w:ilvl[^>]*w:val="0"/);
  });

  it("inserts with numId and custom numLevel", async () => {
    const p = await createTmpDoc("Existing");
    await insertParagraph(p, "Sub-item", -1, undefined, false, "Claude", 5, 2);
    const xml = await readRawDocXml(p);
    expect(xml).toMatch(/w:numId[^>]*w:val="5"/);
    expect(xml).toMatch(/w:ilvl[^>]*w:val="2"/);
  });

  it("inserts with numId combined with style", async () => {
    const p = await createTmpDoc("Existing");
    await insertParagraph(p, "Heading item", -1, "Heading1", false, "Claude", 14, 0);
    const xml = await readRawDocXml(p);
    expect(xml).toContain("Heading1");
    expect(xml).toContain("w:numPr");
  });

  it("inserts with numId tracked", async () => {
    const p = await createTmpDoc("Existing");
    await insertParagraph(p, "Tracked numbered", -1, undefined, true, "Claude", 14, 0);
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:numPr");
    expect(xml).toContain("w:ins");
  });

  it("inserts with copy_format_from (untracked)", async () => {
    const p = await createTmpDoc("Existing");
    // First, insert a paragraph with numbering to use as source
    await insertParagraph(p, "Source heading", 0, "Heading1", false, "Claude", 14, 0);
    // Now insert a new paragraph copying format from index 0
    await insertParagraph(p, "Copied format", -1, undefined, false, "Claude", undefined, undefined, 0);
    const xml = await readRawDocXml(p);
    // The last paragraph should have the same numPr as the source
    // Count occurrences of numId=14 — should appear twice
    const matches = xml.match(/w:numId/g);
    expect(matches).not.toBeNull();
    expect(matches!.length).toBeGreaterThanOrEqual(2);
  });

  it("inserts with copy_format_from (tracked)", async () => {
    const p = await createTmpDoc("Existing");
    await insertParagraph(p, "Source", 0, "Heading2", false);
    // Copy format from the Heading2 paragraph with tracking
    await insertParagraph(p, "Copied tracked", -1, undefined, true, "Bob", undefined, undefined, 0);
    const xml = await readRawDocXml(p);
    // Should have Heading2 style in the copied pPr and w:ins for tracking
    expect(xml).toContain("Heading2");
    expect(xml).toContain("w:ins");
    expect(xml).toContain("Bob");
  });

  it("copy_format_from throws on invalid index", async () => {
    const p = await createTmpDoc("Only one paragraph");
    await expect(
      insertParagraph(p, "Bad ref", -1, undefined, false, "Claude", undefined, undefined, 999),
    ).rejects.toThrow(/out of range/);
  });

  it("copy_format_from a paragraph with no pPr produces plain paragraph", async () => {
    const p = await createTmpDoc("No formatting here");
    // Block 0 has no pPr — copy_format_from should fall back to no-format insert
    await insertParagraph(p, "Plain copy", -1, undefined, false, "Claude", undefined, undefined, 0);
    const xml = await readRawDocXml(p);
    // The new paragraph should exist and have text, but no extra pPr
    expect(xml).toContain("Plain copy");
  });

  it("copy_format_from a table block throws NOT_A_PARAGRAPH", async () => {
    const { insertTable } = await import("../docx-engine.js");
    const p = await createTmpDoc("Before table");
    await insertTable(p, -1, 2, 2);
    // Block 1 is now a table
    await expect(
      insertParagraph(p, "Bad", -1, undefined, false, "Claude", undefined, undefined, 1),
    ).rejects.toThrow(/not a paragraph/);
  });

  it("copy_format_from strips stale revision markers from source pPr (tracked)", async () => {
    const p = await createDocWithTrackedPPr("Source with tracked rPr");
    // Block 0 has w:ins from OldAuthor in its pPr > rPr
    await insertParagraph(p, "Fresh copy", -1, undefined, true, "NewAuthor", undefined, undefined, 0);
    const xml = await readRawDocXml(p);
    // Should NOT contain OldAuthor in the new paragraph's pPr
    // The last w:ins in the doc should be from NewAuthor
    const insMatches = [...xml.matchAll(/w:author="([^"]+)"/g)];
    const authors = insMatches.map(m => m[1]);
    // OldAuthor should appear only once (the original paragraph), NewAuthor for the copy
    expect(authors.filter(a => a === "OldAuthor").length).toBe(1);
    expect(authors.filter(a => a === "NewAuthor").length).toBeGreaterThanOrEqual(1);
  });

  it("copy_format_from + track_changes=false strips stale revision markers", async () => {
    const p = await createDocWithTrackedPPr("Source with tracked rPr");
    // Block 0 has w:ins from OldAuthor in its pPr > rPr
    // Insert with track_changes=false — new paragraph must NOT carry OldAuthor's marker
    await insertParagraph(p, "Untracked copy", -1, undefined, false, "Claude", undefined, undefined, 0);
    const xml = await readRawDocXml(p);
    // Count w:author="OldAuthor" — should appear only once (the original paragraph)
    const oldAuthorMatches = [...xml.matchAll(/w:author="OldAuthor"/g)];
    expect(oldAuthorMatches.length).toBe(1);
  });

  it("copy_format_from preserves alignment, indentation, and numPr", async () => {
    const p = await createDocWithNumberedParagraph("第1条 定義", 14, 0, {
      style: "Heading1",
      alignment: "center",
      indentLeft: 720,
    });
    // Copy format from block 0 to a new paragraph
    await insertParagraph(p, "第2条 遡及適用", -1, undefined, false, "Claude", undefined, undefined, 0);
    const xml = await readRawDocXml(p);
    // Should have two paragraphs with Heading1, numId=14, center alignment, indent
    expect((xml.match(/Heading1/g) || []).length).toBeGreaterThanOrEqual(2);
    expect((xml.match(/w:numId/g) || []).length).toBeGreaterThanOrEqual(2);
    expect((xml.match(/w:jc/g) || []).length).toBeGreaterThanOrEqual(2);
    expect((xml.match(/w:ind/g) || []).length).toBeGreaterThanOrEqual(2);
  });

  it("copy_format_from overrides style and num_id when both provided", async () => {
    const p = await createDocWithNumberedParagraph("Source", 14, 0, { style: "Heading1" });
    // Provide style=Heading3 and num_id=99, but copy_format_from=0 should win
    await insertParagraph(p, "Should get Heading1", -1, "Heading3", false, "Claude", 99, 1, 0);
    const xml = await readRawDocXml(p);
    // Heading3 and numId=99 should NOT appear; Heading1 and numId=14 should appear twice
    expect(xml).not.toContain("Heading3");
    expect(xml).not.toMatch(/w:numId[^>]*w:val="99"/);
    expect((xml.match(/Heading1/g) || []).length).toBeGreaterThanOrEqual(2);
    expect((xml.match(/w:val="14"/g) || []).length).toBeGreaterThanOrEqual(2);
  });

  it("num_level without num_id does not emit numPr", async () => {
    const p = await createTmpDoc("Existing");
    await insertParagraph(p, "No numbering", -1, undefined, false, "Claude", undefined, 2);
    const xml = await readRawDocXml(p);
    expect(xml).not.toContain("w:numPr");
  });
});

// =========================================================================
// deleteParagraph
// =========================================================================

describe("deleteParagraph", () => {
  it("removes paragraph physically (untracked)", async () => {
    const p = await createTmpDoc("Para A\nPara B\nPara C");
    await deleteParagraph(p, 1, false);
    const doc = await readDocument(p);
    expect(doc).toContain("Para A");
    expect(doc).not.toContain("Para B");
    expect(doc).toContain("Para C");
    expect(doc).toContain("Total blocks: 2");
  });

  it("marks paragraph as deleted (tracked)", async () => {
    const p = await createTmpDoc("Para A\nPara B");
    await deleteParagraph(p, 0, true, "Deleter");
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:del");
    expect(xml).toContain("w:delText");
    expect(xml).toContain("Deleter");
    // The accepted view should not show the deleted text
    const doc = await readDocument(p, undefined, undefined, false);
    expect(doc).not.toMatch(/\[0\].*Para A\b/);
  });

  it("throws INDEX_OUT_OF_RANGE for out-of-range index", async () => {
    const p = await createTmpDoc("Only one");
    await expect(deleteParagraph(p, 5, false)).rejects.toMatchObject({
      code: "INDEX_OUT_OF_RANGE",
    });
  });

  it("tracked deletion converts runs to w:delText", async () => {
    const p = await createTmpDoc("Delete me tracked");
    await deleteParagraph(p, 0, true);
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:delText");
    // Also adds pPr > rPr > w:del marker
    expect(xml).toMatch(/w:pPr/);
  });

  it("tracked deletion shows in revision view", async () => {
    const p = await createTmpDoc("Will be deleted");
    await deleteParagraph(p, 0, true);
    const rev = await readDocument(p, undefined, undefined, true);
    expect(rev).toContain("[-Will be deleted-]");
  });
});

// =========================================================================
// deleteParagraphs (bulk)
// =========================================================================

describe("deleteParagraphs", () => {
  it("hard-deletes multiple paragraphs at once", async () => {
    const p = await createTmpDoc("Keep A\nDelete B\nKeep C\nDelete D\nKeep E");
    const result = await deleteParagraphs(p, [1, 3], false);
    expect(result).toContain("2 block(s)");
    const doc = await readDocument(p);
    expect(doc).toContain("Keep A");
    expect(doc).not.toContain("Delete B");
    expect(doc).toContain("Keep C");
    expect(doc).not.toContain("Delete D");
    expect(doc).toContain("Keep E");
  });

  it("tracked-deletes multiple paragraphs with w:delText", async () => {
    const p = await createTmpDoc("First\nSecond\nThird");
    const result = await deleteParagraphs(p, [0, 2], true, "BulkAuthor");
    expect(result).toContain("2 block(s)");
    expect(result).toContain("(tracked)");
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:delText");
    expect(xml).toContain("BulkAuthor");
  });

  it("tracked deletion shows in revision view", async () => {
    const p = await createTmpDoc("Alpha\nBeta\nGamma");
    await deleteParagraphs(p, [0, 2], true);
    const rev = await readDocument(p, undefined, undefined, true);
    expect(rev).toContain("[-Alpha-]");
    expect(rev).toContain("[-Gamma-]");
    expect(rev).not.toContain("[-Beta-]");
  });

  it("throws on out-of-range index", async () => {
    const p = await createTmpDoc("One\nTwo");
    await expect(deleteParagraphs(p, [0, 10], false)).rejects.toMatchObject({
      code: "INDEX_OUT_OF_RANGE",
    });
  });

  it("handles deleting all paragraphs (hard delete)", async () => {
    const p = await createTmpDoc("A\nB\nC");
    const result = await deleteParagraphs(p, [0, 1, 2], false);
    expect(result).toContain("3 block(s)");
  });
});

// =========================================================================
// XML entity handling — special characters in text
// =========================================================================

describe("XML entity handling", () => {
  it("round-trips text with ampersand via editParagraph", async () => {
    const p = await createTmpDoc("Hello world");
    await editParagraph(p, 0, "AT&T and <tags> work", false);
    const result = await readDocument(p);
    expect(result).toContain("AT&T and <tags> work");
  });

  it("round-trips text with ampersand via replaceText (untracked)", async () => {
    const p = await createTmpDoc("Hello world");
    await replaceText(p, "world", "R&D dept", false, false);
    const result = await readDocument(p);
    expect(result).toContain("Hello R&D dept");
  });

  it("round-trips text with ampersand via replaceText (tracked)", async () => {
    const p = await createTmpDoc("Hello world");
    await replaceText(p, "world", "R&D dept", false, true);
    const result = await readDocument(p);
    expect(result).toContain("Hello R&D dept");
  });

  it("round-trips text with ampersand via insertParagraph", async () => {
    const p = await createTmpDoc("First");
    await insertParagraph(p, "A<B & C>D", -1, undefined, false);
    const result = await readDocument(p);
    expect(result).toContain("A<B & C>D");
  });

  it("produces valid XML when text contains special chars", async () => {
    const p = await createTmpDoc("Hello world");
    await editParagraph(p, 0, "AT&T <Corp> \"quoted\"", false);
    const raw = await readRawDocXml(p);
    // The raw XML should contain proper entity encoding
    expect(raw).toContain("&amp;");
    expect(raw).toContain("&lt;");
    expect(raw).toContain("&gt;");
    // And should NOT contain unescaped & or < in text content
    expect(raw).not.toMatch(/AT&T/);
    expect(raw).not.toMatch(/<Corp>/);
  });
});
