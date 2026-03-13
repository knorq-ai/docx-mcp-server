import { describe, it, expect, afterEach } from "vitest";
import {
  createTmpDoc,
  cleanupTmpFiles,
  readRawDocXml,
  createCrossRunDoc,
} from "./helpers.js";
import {
  formatText,
  setParagraphFormat,
  setParagraphFormats,
  highlightText,
  setHeading,
  readDocument,
} from "../docx-engine.js";

afterEach(cleanupTmpFiles);

// =========================================================================
// formatText
// =========================================================================

describe("formatText", () => {
  it("applies bold formatting", async () => {
    const p = await createTmpDoc("Make this bold");
    const result = await formatText(p, "bold", { bold: true });
    expect(result).toContain("1 occurrence");
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:b");
  });

  it("applies italic formatting", async () => {
    const p = await createTmpDoc("Make this italic");
    await formatText(p, "italic", { italic: true });
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:i");
  });

  it("applies underline formatting", async () => {
    const p = await createTmpDoc("Underline me");
    await formatText(p, "Underline", { underline: true });
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:u");
    expect(xml).toContain('w:val="single"');
  });

  it("applies strikethrough formatting", async () => {
    const p = await createTmpDoc("Strike this out");
    await formatText(p, "Strike", { strikethrough: true });
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:strike");
  });

  it("applies font name", async () => {
    const p = await createTmpDoc("Change my font");
    await formatText(p, "font", { fontName: "Arial" });
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:rFonts");
    expect(xml).toContain("Arial");
  });

  it("applies font size (points to half-points)", async () => {
    const p = await createTmpDoc("Resize me");
    await formatText(p, "Resize", { fontSize: 14 });
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:sz");
    expect(xml).toContain('w:val="28"'); // 14 * 2 = 28 half-points
  });

  it("applies font color", async () => {
    const p = await createTmpDoc("Color me red");
    await formatText(p, "Color", { fontColor: "FF0000" });
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:color");
    expect(xml).toContain("FF0000");
  });

  it("applies multiple formatting options at once", async () => {
    const p = await createTmpDoc("Format everything");
    await formatText(p, "everything", {
      bold: true,
      italic: true,
      fontSize: 16,
    });
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:b");
    expect(xml).toContain("w:i");
    expect(xml).toContain('w:val="32"'); // 16 * 2
  });

  it("case-insensitive matching by default", async () => {
    const p = await createTmpDoc("BOLD text");
    const result = await formatText(p, "bold", { bold: true }, false);
    expect(result).toContain("1 occurrence");
  });

  it("case-sensitive matching when requested", async () => {
    const p = await createTmpDoc("BOLD text bold");
    const result = await formatText(p, "BOLD", { bold: true }, true);
    // Only "BOLD" matches case-sensitively
    expect(result).toContain("1 occurrence");
  });

  it("reports no matches when text is absent", async () => {
    const p = await createTmpDoc("Hello world");
    const result = await formatText(p, "nonexistent", { bold: true });
    expect(result).toContain("No occurrences");
  });

  it("can remove bold formatting", async () => {
    const p = await createTmpDoc("Already bold text");
    await formatText(p, "bold", { bold: true });
    let xml = await readRawDocXml(p);
    expect(xml).toContain("w:b");

    await formatText(p, "bold", { bold: false });
    xml = await readRawDocXml(p);
    // w:b tag should be removed (check for the element, not substring in w:body etc.)
    expect(xml).not.toMatch(/<w:b[\s/>]/);
  });

  it("formats text in multiple runs across paragraphs", async () => {
    const p = await createTmpDoc("Match here\nMatch there");
    const result = await formatText(p, "Match", { bold: true });
    expect(result).toContain("2 occurrence");
  });
});

// =========================================================================
// setParagraphFormat
// =========================================================================

describe("setParagraphFormat", () => {
  it("sets alignment to center", async () => {
    const p = await createTmpDoc("Center this");
    await setParagraphFormat(p, 0, { alignment: "center" });
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:jc");
    expect(xml).toContain('w:val="center"');
  });

  it("sets alignment to justify (maps to both)", async () => {
    const p = await createTmpDoc("Justify this");
    await setParagraphFormat(p, 0, { alignment: "justify" });
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:jc");
    expect(xml).toContain('w:val="both"');
  });

  it("sets space before and after (points to twips)", async () => {
    const p = await createTmpDoc("Spaced paragraph");
    await setParagraphFormat(p, 0, { spaceBefore: 12, spaceAfter: 6 });
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:spacing");
    expect(xml).toContain('w:before="240"'); // 12 * 20
    expect(xml).toContain('w:after="120"'); // 6 * 20
  });

  it("sets line spacing", async () => {
    const p = await createTmpDoc("Line spaced");
    await setParagraphFormat(p, 0, { lineSpacing: 24 });
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:spacing");
    expect(xml).toContain('w:line="480"'); // 24 * 20
    expect(xml).toContain('w:lineRule="exact"');
  });

  it("sets left and right indentation", async () => {
    const p = await createTmpDoc("Indented paragraph");
    await setParagraphFormat(p, 0, { indentLeft: 720, indentRight: 360 });
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:ind");
    expect(xml).toContain('w:left="720"');
    expect(xml).toContain('w:right="360"');
  });

  it("sets first line indent", async () => {
    const p = await createTmpDoc("First line indented");
    await setParagraphFormat(p, 0, { firstLineIndent: 360 });
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:ind");
    expect(xml).toContain('w:firstLine="360"');
  });

  it("sets hanging indent", async () => {
    const p = await createTmpDoc("Hanging indent paragraph");
    await setParagraphFormat(p, 0, { hangingIndent: 720 });
    const xml = await readRawDocXml(p);
    expect(xml).toContain('w:hanging="720"');
  });

  it("throws INDEX_OUT_OF_RANGE for out-of-range index", async () => {
    const p = await createTmpDoc("Only one");
    await expect(setParagraphFormat(p, 10, { alignment: "center" })).rejects.toMatchObject({
      code: "INDEX_OUT_OF_RANGE",
    });
  });
});

// =========================================================================
// highlightText
// =========================================================================

describe("highlightText", () => {
  it("highlights text with default yellow color", async () => {
    const p = await createTmpDoc("Highlight me");
    const result = await highlightText(p, "Highlight");
    expect(result).toContain("1 occurrence");
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:highlight");
    expect(xml).toContain('w:val="yellow"');
  });

  it("highlights text with custom color", async () => {
    const p = await createTmpDoc("Green highlight");
    await highlightText(p, "Green", "green");
    const xml = await readRawDocXml(p);
    expect(xml).toContain('w:val="green"');
  });

  it("highlights with case-insensitive matching", async () => {
    const p = await createTmpDoc("HIGHLIGHT this");
    const result = await highlightText(p, "highlight", "cyan", false);
    expect(result).toContain("1 occurrence");
  });
});

// =========================================================================
// setHeading
// =========================================================================

describe("setHeading", () => {
  it("converts paragraph to Heading 1", async () => {
    const p = await createTmpDoc("Section Title");
    await setHeading(p, 0, 1);
    const xml = await readRawDocXml(p);
    expect(xml).toContain("Heading1");
    expect(xml).toContain('w:val="0"'); // outlineLvl = level - 1
    const doc = await readDocument(p);
    expect(doc).toContain("(H1)");
  });

  it("converts paragraph to Heading 3", async () => {
    const p = await createTmpDoc("Subsection");
    await setHeading(p, 0, 3);
    const xml = await readRawDocXml(p);
    expect(xml).toContain("Heading3");
    expect(xml).toContain('w:val="2"'); // outlineLvl
  });

  it("throws INVALID_PARAMETER for invalid heading level", async () => {
    const p = await createTmpDoc("Text");
    await expect(setHeading(p, 0, 0)).rejects.toMatchObject({
      code: "INVALID_PARAMETER",
    });
  });

  it("throws INDEX_OUT_OF_RANGE for out-of-range paragraph index", async () => {
    const p = await createTmpDoc("Text");
    await expect(setHeading(p, 99, 1)).rejects.toMatchObject({
      code: "INDEX_OUT_OF_RANGE",
    });
  });

  it("replaces existing heading style", async () => {
    const p = await createTmpDoc("Text", "Title");
    // Title is H1, change it to H2
    await setHeading(p, 0, 2);
    const xml = await readRawDocXml(p);
    expect(xml).toContain("Heading2");
    expect(xml).not.toContain("Heading1");
  });
});

// =========================================================================
// Cross-run formatting
// =========================================================================

describe("formatText cross-run", () => {
  it("formats text spanning multiple runs", async () => {
    // "Hello World" split across two runs: "Hello Wo" + "rld"
    const p = await createCrossRunDoc(["Hello Wo", "rld"]);
    const result = await formatText(p, "World", { bold: true }, false);
    expect(result).toContain("1 occurrence(s)");
    const xml = await readRawDocXml(p);
    // The matched portion should have bold formatting
    expect(xml).toContain("w:b");
  });

  it("formats only the matched portion of a single run", async () => {
    const p = await createCrossRunDoc(["Hello World Goodbye"]);
    const result = await formatText(p, "World", { italic: true }, false);
    expect(result).toContain("1 occurrence(s)");
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:i");
    // The text should be split: "Hello " (no italic), "World" (italic), " Goodbye" (no italic)
    const doc = await readDocument(p);
    expect(doc).toContain("Hello World Goodbye");
  });

  it("handles multiple cross-run matches in one paragraph", async () => {
    // "abc abc" split as "ab" + "c ab" + "c"
    const p = await createCrossRunDoc(["ab", "c ab", "c"]);
    const result = await formatText(p, "abc", { underline: true }, false);
    expect(result).toContain("2 occurrence(s)");
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:u");
  });

  it("cross-run highlight delegates to formatText", async () => {
    const p = await createCrossRunDoc(["High", "light me"]);
    const result = await highlightText(p, "Highlight", "green", false);
    expect(result).toContain("1 occurrence(s)");
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:highlight");
    expect(xml).toContain('w:val="green"');
  });
});

// =========================================================================
// setParagraphFormats
// =========================================================================

describe("setParagraphFormats", () => {
  it("applies different formats to different groups", async () => {
    const p = await createTmpDoc("Line one\nLine two\nLine three\nLine four");
    const result = await setParagraphFormats(p, [
      { indices: [0, 1], format: { alignment: "center", spaceAfter: 6 } },
      { indices: [2, 3], format: { alignment: "right", spaceBefore: 12 } },
    ]);
    expect(result).toContain("4 paragraph(s)");
    const xml = await readRawDocXml(p);
    expect(xml).toContain('w:val="center"');
    expect(xml).toContain('w:val="right"');
    // spaceAfter=6 → 6*20=120 twips
    expect(xml).toContain('w:after="120"');
    // spaceBefore=12 → 12*20=240 twips
    expect(xml).toContain('w:before="240"');
  });

  it("applies indentation via bulk", async () => {
    const p = await createTmpDoc("Indent me\nAnd me too");
    await setParagraphFormats(p, [
      { indices: [0, 1], format: { indentLeft: 720 } },
    ]);
    const xml = await readRawDocXml(p);
    expect(xml).toContain('w:left="720"');
  });

  it("throws on out-of-range index", async () => {
    const p = await createTmpDoc("Only one paragraph");
    await expect(
      setParagraphFormats(p, [
        { indices: [0, 5], format: { alignment: "center" } },
      ]),
    ).rejects.toMatchObject({ code: "INDEX_OUT_OF_RANGE" });
  });

  it("throws on non-paragraph block", async () => {
    const { insertTable } = await import("../docx-engine.js");
    const p = await createTmpDoc("Before table");
    await insertTable(p, -1, 1, 1);
    // Block 1 is the table
    await expect(
      setParagraphFormats(p, [
        { indices: [1], format: { alignment: "center" } },
      ]),
    ).rejects.toMatchObject({ code: "NOT_A_PARAGRAPH" });
  });
});
