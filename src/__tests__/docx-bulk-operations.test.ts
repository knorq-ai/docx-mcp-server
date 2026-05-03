import { describe, it, expect, afterEach } from "vitest";
import {
  createTmpDoc,
  cleanupTmpFiles,
  readRawDocXml,
  createDocWithNumberedParagraph,
  createDocWithHeaderFooter,
} from "./helpers.js";
import {
  readDocument,
  editParagraphs,
  insertParagraphs,
  setHeadings,
  editTableCells,
  insertTable,
  acceptAllChanges,
  rejectAllChanges,
  replaceTexts,
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
// replaceTexts (bulk)
// =========================================================================

describe("replaceTexts", () => {
  it("applies multiple replacements in one call", async () => {
    const p = await createTmpDoc("alpha beta gamma\ndelta alpha epsilon");
    const result = await replaceTexts(
      p,
      [
        { search: "alpha", replace: "ALPHA" },
        { search: "delta", replace: "DELTA" },
      ],
      false,
    );
    const doc = await readDocument(p);
    expect(doc).toContain("ALPHA");
    expect(doc).toContain("DELTA");
    expect(result).toContain("Replaced 3 occurrence(s) across 2 item(s)");
  });

  it("returns no-op message when nothing matches", async () => {
    const p = await createTmpDoc("hello world");
    const result = await replaceTexts(
      p,
      [
        { search: "missing-1", replace: "x" },
        { search: "missing-2", replace: "y" },
      ],
      false,
    );
    expect(result).toContain("No occurrences");
    const doc = await readDocument(p);
    expect(doc).toContain("hello world");
  });

  it("handles empty items array without file I/O", async () => {
    const p = await createTmpDoc("Unchanged");
    const result = await replaceTexts(p, [], false);
    expect(result).toContain("No replacements");
    const doc = await readDocument(p);
    expect(doc).toContain("Unchanged");
  });

  it("produces tracked changes with del/ins markup and unique revision IDs", async () => {
    const p = await createTmpDoc("foo bar baz\nqux foo");
    await replaceTexts(
      p,
      [
        { search: "foo", replace: "FOO" },
        { search: "baz", replace: "BAZ" },
      ],
      true,
      "TestBot",
    );
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:del");
    expect(xml).toContain("w:ins");
    expect(xml).toContain("TestBot");

    // Revision IDs across all w:ins/w:del must be unique
    const idMatches = [...xml.matchAll(/w:(?:ins|del)[^>]*\bw:id="(\d+)"/g)].map((m) => m[1]);
    expect(idMatches.length).toBeGreaterThan(0);
    expect(new Set(idMatches).size).toBe(idMatches.length);

    await acceptAllChanges(p);
    const doc = await readDocument(p);
    expect(doc).toContain("FOO");
    expect(doc).toContain("BAZ");
  });

  it("respects per-item case_sensitive flag", async () => {
    const p = await createTmpDoc("Hello hello HELLO");
    await replaceTexts(
      p,
      [{ search: "Hello", replace: "X", caseSensitive: true }],
      false,
    );
    const doc = await readDocument(p);
    // Only the exact-case match is replaced
    expect(doc).toContain("X hello HELLO");
  });

  it("rejects an empty search string instead of looping forever", async () => {
    const p = await createTmpDoc("anything");
    await expect(
      replaceTexts(p, [{ search: "", replace: "x" }], false),
    ).rejects.toThrow(/empty/);
  });

  it("supports mixed per-item case_sensitive flags", async () => {
    const p = await createTmpDoc("Hello hello WORLD world");
    await replaceTexts(
      p,
      [
        { search: "Hello", replace: "X", caseSensitive: true },
        { search: "world", replace: "Y", caseSensitive: false },
      ],
      false,
    );
    const doc = await readDocument(p);
    // Only "Hello" replaced (case-sensitive); both "WORLD" and "world" replaced (case-insensitive)
    expect(doc).toContain("X hello Y Y");
  });

  it("replaces inside headers and footers when include_headers_footers is true", async () => {
    const p = await createDocWithHeaderFooter("body alpha", "header alpha", "footer alpha");
    await replaceTexts(
      p,
      [{ search: "alpha", replace: "ALPHA" }],
      true,
      "TestBot",
      true,
    );

    const fs = await import("fs/promises");
    const JSZip = (await import("jszip")).default;
    const after = await JSZip.loadAsync(await fs.readFile(p));
    const bodyAfter = await after.file("word/document.xml")!.async("string");
    // Header/footer file names follow whatever createDocWithHeaderFooter sets up.
    const hfFiles = Object.keys(after.files).filter((name) =>
      /^word\/(header|footer)\d*\.xml$/.test(name),
    );
    expect(hfFiles.length).toBeGreaterThan(0);

    expect(bodyAfter).toContain("ALPHA");
    for (const hfFile of hfFiles) {
      const xml = await after.file(hfFile)!.async("string");
      expect(xml).toContain("ALPHA");
    }

    // Cross-part revision IDs should be unique (the body's scanMaxId seed
    // forwards a running counter into each HF file).
    const allXml = [bodyAfter, ...(await Promise.all(
      hfFiles.map((f) => after.file(f)!.async("string")),
    ))];
    const allIds = allXml.flatMap((xml) =>
      [...xml.matchAll(/w:(?:ins|del)[^>]*\bw:id="(\d+)"/g)].map((m) => m[1]),
    );
    expect(allIds.length).toBeGreaterThan(0);
    expect(new Set(allIds).size).toBe(allIds.length);
  });

  it("applies items sequentially (later items can match earlier replacements)", async () => {
    const p = await createTmpDoc("alpha");
    await replaceTexts(
      p,
      [
        { search: "alpha", replace: "beta" },
        { search: "beta", replace: "gamma" },
      ],
      false,
    );
    const doc = await readDocument(p);
    expect(doc).toContain("gamma");
  });

  it("rejects overlapping items in tracked mode (later search ⊇ earlier replace)", async () => {
    const p = await createTmpDoc("alpha");
    // earlier.replace = "beta", later.search = "beta" — exact match
    await expect(
      replaceTexts(
        p,
        [
          { search: "alpha", replace: "beta" },
          { search: "beta", replace: "gamma" },
        ],
        true,
      ),
    ).rejects.toMatchObject({ code: "INVALID_PARAMETER" });
  });

  it("allows tracked delete (empty replace) followed by unrelated tracked replace", async () => {
    const p = await createTmpDoc("obsolete\nfoo bar");
    // Item 1 deletes "obsolete" (replace=""). Item 2 replaces "foo" with "bar".
    // Without the empty-replace skip, the bidirectional check would falsely
    // reject this because needle.includes("") is always true.
    await replaceTexts(
      p,
      [
        { search: "obsolete", replace: "" },
        { search: "foo", replace: "FOO" },
      ],
      true,
    );
    const xml = await readRawDocXml(p);
    expect(xml).toContain("FOO");
  });

  it("rejects overlapping items in tracked mode (boundary overlap, no full containment)", async () => {
    const p = await createTmpDoc("abcZ");
    // earlier.replace = "aXc", later.search = "XcZ".
    // Neither contains the other, but they share "Xc" at the boundary:
    // "aXc".endsWith("Xc") and "XcZ".startsWith("Xc"). After tracked
    // item 1 the doc has w:ins("aXc") followed by "Z" in normal text,
    // so item 2's "XcZ" search would span the ins boundary — corrupting
    // reject. The guard refuses upfront.
    await expect(
      replaceTexts(
        p,
        [
          { search: "abc", replace: "aXc" },
          { search: "XcZ", replace: "Y" },
        ],
        true,
      ),
    ).rejects.toMatchObject({ code: "INVALID_PARAMETER" });
  });

  it("rejects overlapping items in tracked mode (later search wraps earlier replace)", async () => {
    const p = await createTmpDoc("ax alpha y");
    // earlier.replace = "beta", later.search = "x beta y" — wraps the earlier replace.
    // Without bidirectional check this would slip through.
    await expect(
      replaceTexts(
        p,
        [
          { search: "alpha", replace: "beta" },
          { search: "x beta y", replace: "z" },
        ],
        true,
      ),
    ).rejects.toMatchObject({ code: "INVALID_PARAMETER" });
  });

  it("allows the same overlap in untracked mode (no nesting risk)", async () => {
    const p = await createTmpDoc("alpha");
    await replaceTexts(
      p,
      [
        { search: "alpha", replace: "beta" },
        { search: "beta", replace: "gamma" },
      ],
      false,
    );
    const doc = await readDocument(p);
    expect(doc).toContain("gamma");
  });

  it("respects per-item case_sensitive when detecting tracked overlap", async () => {
    const p = await createTmpDoc("Alpha");
    // Both items use caseSensitive=true. Item 1 produces "BETA" (uppercase),
    // item 2 searches for "beta" (lowercase) — case-sensitive comparison
    // means no overlap, so this should be allowed.
    await replaceTexts(
      p,
      [
        { search: "Alpha", replace: "BETA", caseSensitive: true },
        { search: "beta", replace: "gamma", caseSensitive: true },
      ],
      true,
    );
    // The document should be successfully edited (no throw) and contain BETA.
    const xml = await readRawDocXml(p);
    expect(xml).toContain("BETA");
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

  it("inserts with numId and numLevel", async () => {
    const p = await createTmpDoc("Existing");
    await insertParagraphs(
      p,
      [
        { text: "Item 1", position: -1, numId: 14, numLevel: 0 },
        { text: "Item 2", position: -1, numId: 14, numLevel: 0 },
      ],
      false,
    );
    const xml = await readRawDocXml(p);
    const numIdMatches = xml.match(/w:numId/g);
    expect(numIdMatches).not.toBeNull();
    expect(numIdMatches!.length).toBe(2);
  });

  it("inserts with copy_format_from", async () => {
    const p = await createTmpDoc("Existing");
    // Insert a heading to serve as format source
    await insertParagraphs(
      p,
      [{ text: "Source heading", position: 0, style: "Heading1", numId: 14, numLevel: 0 }],
      false,
    );
    // Now copy format from index 0
    await insertParagraphs(
      p,
      [{ text: "Copied para", position: -1, copyFormatFrom: 0 }],
      false,
    );
    const xml = await readRawDocXml(p);
    // Should have Heading1 and numId in both original and copied
    const headingMatches = xml.match(/Heading1/g);
    expect(headingMatches).not.toBeNull();
    expect(headingMatches!.length).toBeGreaterThanOrEqual(2);
  });

  it("bulk copy_format_from references original indices (not shifted)", async () => {
    // Block 0 = numbered heading, Block 1 = plain
    const p = await createDocWithNumberedParagraph("第1条 定義", 14, 0, { style: "Heading1" });
    // Insert two paragraphs both copying from original block 0.
    // Even though positions shift, copy_format_from should resolve against the
    // original document state before any inserts.
    await insertParagraphs(
      p,
      [
        { text: "第2条 適用", position: 1, copyFormatFrom: 0 },
        { text: "第3条 委託", position: 2, copyFormatFrom: 0 },
      ],
      false,
    );
    const xml = await readRawDocXml(p);
    // Should have three Heading1 paragraphs with numId=14
    expect((xml.match(/Heading1/g) || []).length).toBeGreaterThanOrEqual(3);
    expect((xml.match(/w:numId/g) || []).length).toBeGreaterThanOrEqual(3);
  });

  it("bulk mixed numId and copy_format_from items", async () => {
    const p = await createDocWithNumberedParagraph("Source", 14, 0, { style: "Heading1" });
    await insertParagraphs(
      p,
      [
        { text: "Explicit num", position: -1, numId: 5, numLevel: 1 },
        { text: "Copied format", position: -1, copyFormatFrom: 0 },
      ],
      false,
    );
    const xml = await readRawDocXml(p);
    // numId=5 for explicit, numId=14 for copy (and original)
    expect(xml).toMatch(/w:numId[^>]*w:val="5"/);
    expect((xml.match(/w:val="14"/g) || []).length).toBeGreaterThanOrEqual(2);
  });
});

// =========================================================================
// setHeadings
// =========================================================================

describe("setHeadings", () => {
  it("sets multiple headings in one call", async () => {
    const p = await createTmpDoc("Title\nSubtitle\nBody\nAnother");
    await setHeadings(p, [
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
      setHeadings(p, [{ paragraphIndex: 0, level: 0 }]),
    ).rejects.toThrow("between 1 and 9");
    await expect(
      setHeadings(p, [{ paragraphIndex: 0, level: 10 }]),
    ).rejects.toThrow("between 1 and 9");
  });

  it("throws INDEX_OUT_OF_RANGE for invalid index", async () => {
    const p = await createTmpDoc("One");
    await expect(
      setHeadings(p, [{ paragraphIndex: 99, level: 1 }]),
    ).rejects.toThrow("out of range");
  });

  it("throws NOT_A_PARAGRAPH for table block", async () => {
    const p = await createTmpDoc("Before");
    await insertTable(p, -1, 2, 2);
    await expect(
      setHeadings(p, [{ paragraphIndex: 1, level: 1 }]),
    ).rejects.toThrow("not a paragraph");
  });

  it("returns correct summary", async () => {
    const p = await createTmpDoc("A\nB");
    const result = await setHeadings(p, [
      { paragraphIndex: 0, level: 1 },
      { paragraphIndex: 1, level: 2 },
    ]);
    expect(result).toContain("2 paragraph(s)");
  });

  it("sets outline level correctly in XML", async () => {
    const p = await createTmpDoc("Test");
    await setHeadings(p, [{ paragraphIndex: 0, level: 3 }]);
    const xml = await readRawDocXml(p);
    expect(xml).toContain('w:val="Heading3"');
    expect(xml).toContain('w:val="2"'); // outlineLvl = level - 1
  });

  it("handles empty items array without file I/O", async () => {
    const p = await createTmpDoc("Unchanged");
    const result = await setHeadings(p, []);
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
