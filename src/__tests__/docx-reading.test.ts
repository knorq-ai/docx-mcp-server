import { describe, it, expect, afterEach } from "vitest";
import {
  createTmpDoc,
  cleanupTmpFiles,
  createCrossRunDoc,
  createDocWithEmbeddedImage,
} from "./helpers.js";
import {
  readDocument,
  getDocumentInfo,
  searchText,
  createDocument,
  insertParagraph,
  replaceText,
  insertTable,
  addComment,
  setHeading,
  listImages,
  listImagesStructured,
} from "../docx-engine.js";

afterEach(cleanupTmpFiles);

// =========================================================================
// readDocument
// =========================================================================

describe("readDocument", () => {
  it("reads a simple document with content", async () => {
    const p = await createTmpDoc("Hello World");
    const result = await readDocument(p);
    expect(result).toContain("Hello World");
    expect(result).toContain("Total blocks:");
  });

  it("reads a document with title and content", async () => {
    const p = await createTmpDoc("Body text here", "My Title");
    const result = await readDocument(p);
    expect(result).toContain("My Title");
    expect(result).toContain("Body text here");
    expect(result).toContain("(H1)");
  });

  it("reads a document with multiple paragraphs", async () => {
    const p = await createTmpDoc("Line one\nLine two\nLine three");
    const result = await readDocument(p);
    expect(result).toContain("Line one");
    expect(result).toContain("Line two");
    expect(result).toContain("Line three");
  });

  it("supports startParagraph / endParagraph slicing", async () => {
    const p = await createTmpDoc("Para 0\nPara 1\nPara 2\nPara 3");
    const result = await readDocument(p, 1, 3);
    expect(result).toContain("Para 1");
    expect(result).toContain("Para 2");
    expect(result).not.toContain("[0]");
    expect(result).toContain("Showing blocks");
  });

  it("returns block indices in output", async () => {
    const p = await createTmpDoc("First\nSecond");
    const result = await readDocument(p);
    expect(result).toContain("[0]");
    expect(result).toContain("[1]");
  });

  it("shows default (accepted) view — skips deleted text", async () => {
    const p = await createTmpDoc("Hello old world");
    await replaceText(p, "old", "new", false, true);
    const result = await readDocument(p, undefined, undefined, false);
    expect(result).toContain("new");
    expect(result).not.toContain("[-old-]");
  });

  it("shows revisions view with [-deleted-] and [+inserted+]", async () => {
    const p = await createTmpDoc("Hello old world");
    await replaceText(p, "old", "new", false, true);
    const result = await readDocument(p, undefined, undefined, true);
    expect(result).toContain("[-old-]");
    expect(result).toContain("[+new+]");
  });

  it("reads an empty document (no content, no title)", async () => {
    const p = await createTmpDoc();
    const result = await readDocument(p);
    expect(result).toContain("Total blocks:");
  });
});

// =========================================================================
// getDocumentInfo
// =========================================================================

describe("getDocumentInfo", () => {
  it("counts paragraphs correctly", async () => {
    const p = await createTmpDoc("One\nTwo\nThree");
    const result = await getDocumentInfo(p);
    expect(result).toContain("Paragraphs: 3");
  });

  it("counts headings and shows document outline", async () => {
    const p = await createTmpDoc("Body", "My Heading");
    const result = await getDocumentInfo(p);
    expect(result).toContain("Headings: 1");
    expect(result).toContain("H1: My Heading");
    expect(result).toContain("Document outline:");
  });

  it("counts tables", async () => {
    const p = await createTmpDoc("Before table");
    await insertTable(p, -1, 2, 3);
    const result = await getDocumentInfo(p);
    expect(result).toContain("Tables: 1");
  });

  it("reports comment presence", async () => {
    const p = await createTmpDoc("Some text here");
    // Before adding comments
    let result = await getDocumentInfo(p);
    expect(result).toContain("Has comments: false");

    // After adding a comment
    await addComment(p, "Some text", "A comment", "Tester");
    result = await getDocumentInfo(p);
    expect(result).toContain("Has comments: true");
  });
});

// =========================================================================
// searchText
// =========================================================================

describe("searchText", () => {
  it("finds text (case-insensitive by default)", async () => {
    const p = await createTmpDoc("Hello World\nAnother line");
    const result = await searchText(p, "hello");
    expect(result).toContain("1 match");
    expect(result).toContain("Hello World");
  });

  it("finds text case-sensitively", async () => {
    const p = await createTmpDoc("Hello World\nhello again");
    const sensitive = await searchText(p, "Hello", true);
    expect(sensitive).toContain("1 match");

    const insensitive = await searchText(p, "Hello", false);
    expect(insensitive).toContain("2 match");
  });

  it("returns no matches when text is absent", async () => {
    const p = await createTmpDoc("Hello World");
    const result = await searchText(p, "nonexistent");
    expect(result).toContain("No matches found");
  });

  it("finds multiple matches across paragraphs", async () => {
    const p = await createTmpDoc("Foo bar\nFoo baz\nQux foo");
    const result = await searchText(p, "foo", false);
    expect(result).toContain("3 match");
  });

  it("shows context around matches", async () => {
    const p = await createTmpDoc("This is a test sentence with keyword inside");
    const result = await searchText(p, "keyword");
    expect(result).toContain("keyword");
    expect(result).toContain("Block 0");
  });
});

// =========================================================================
// listImages
// =========================================================================

describe("listImages", () => {
  it("lists embedded images with metadata", async () => {
    const p = await createDocWithEmbeddedImage("My alt text");
    const result = await listImages(p);
    expect(result).toContain("Images in");
    expect(result).toContain("(1)");
    expect(result).toContain("image1.png");
    expect(result).toContain("image/png");
    expect(result).toContain("My alt text");
  });

  it("returns 'no images' for document without images", async () => {
    const p = await createTmpDoc("Just text, no images");
    const result = await listImages(p);
    expect(result).toContain("No images");
  });

  it("returns structured data with correct fields", async () => {
    const p = await createDocWithEmbeddedImage("Structured test");
    const result = await listImagesStructured(p);
    expect(result.totalImages).toBe(1);
    expect(result.images).toHaveLength(1);
    const img = result.images[0];
    expect(img.filename).toBe("media/image1.png");
    expect(img.contentType).toBe("image/png");
    expect(img.altText).toBe("Structured test");
    expect(img.name).toBe("Picture 1");
    expect(img.widthEmu).toBe(914400);
    expect(img.heightEmu).toBe(914400);
    expect(img.sizeBytes).toBeGreaterThan(0);
    expect(img.blockIndex).toBe(1); // second block (after "Paragraph before image")
  });

  it("returns empty structured result for no images", async () => {
    const p = await createTmpDoc("No images here");
    const result = await listImagesStructured(p);
    expect(result.totalImages).toBe(0);
    expect(result.images).toEqual([]);
  });
});
