import { describe, it, expect, afterEach } from "vitest";
import {
  createTmpDoc,
  createDocWithNestedTable,
  cleanupTmpFiles,
  readRawDocXml,
  readRawCommentsXml,
  readRawContentTypes,
  readRawDocRels,
} from "./helpers.js";
import {
  addComment,
  addComments,
  readComments,
  replyToComment,
  deleteComment,
} from "../docx-engine.js";

afterEach(cleanupTmpFiles);

// =========================================================================
// addComment
// =========================================================================

describe("addComment", () => {
  it("adds a comment and creates comments.xml", async () => {
    const p = await createTmpDoc("Annotate this text");
    const result = await addComment(p, "this text", "My comment", "Reviewer");
    expect(result).toContain("Added comment");
    expect(result).toContain("Reviewer");

    const commentsXml = await readRawCommentsXml(p);
    expect(commentsXml).toContain("w:comment");
    expect(commentsXml).toContain("My comment");
    expect(commentsXml).toContain("Reviewer");
  });

  it("inserts comment range markers in document.xml", async () => {
    const p = await createTmpDoc("Mark this text please");
    await addComment(p, "this text", "Note here", "Author");
    const xml = await readRawDocXml(p);
    expect(xml).toContain("w:commentRangeStart");
    expect(xml).toContain("w:commentRangeEnd");
    expect(xml).toContain("w:commentReference");
  });

  it("sets up Content_Types and rels infrastructure", async () => {
    const p = await createTmpDoc("Some content");
    await addComment(p, "content", "Check this", "Claude");

    const ct = await readRawContentTypes(p);
    expect(ct).toContain("comments.xml");

    const rels = await readRawDocRels(p);
    expect(rels).toContain("comments.xml");
  });

  it("adds multiple comments with incremented IDs", async () => {
    const p = await createTmpDoc("First target and second target");
    await addComment(p, "First target", "Comment 1", "A");
    await addComment(p, "second target", "Comment 2", "B");

    const commentsXml = await readRawCommentsXml(p);
    expect(commentsXml).toContain("Comment 1");
    expect(commentsXml).toContain("Comment 2");

    const result = await readComments(p);
    expect(result).toContain("Comment 1");
    expect(result).toContain("Comment 2");
  });

  it("throws ANCHOR_NOT_FOUND when anchor text is not found", async () => {
    const p = await createTmpDoc("Hello world");
    await expect(addComment(p, "nonexistent", "Comment")).rejects.toMatchObject({
      code: "ANCHOR_NOT_FOUND",
    });
  });

  it("uses default author Claude", async () => {
    const p = await createTmpDoc("Default author test");
    await addComment(p, "author", "Test comment");
    const commentsXml = await readRawCommentsXml(p);
    expect(commentsXml).toContain("Claude");
  });

  it("escapes XML special characters in comment text", async () => {
    const p = await createTmpDoc("Check this section");
    await addComment(p, "this section", "Review M&A terms & conditions <important>", "Author");
    const commentsXml = await readRawCommentsXml(p);
    // & must be escaped as &amp; in valid XML
    expect(commentsXml).toContain("M&amp;A terms");
    expect(commentsXml).toContain("&amp; conditions");
    expect(commentsXml).toContain("&lt;important&gt;");
  });

  it("splits multi-line comment text into separate paragraphs", async () => {
    const p = await createTmpDoc("Multi line test");
    await addComment(p, "line test", "Line one\nLine two\nLine three", "Author");
    const commentsXml = await readRawCommentsXml(p);
    // Each line should be in its own <w:p>
    expect(commentsXml).toContain("Line one");
    expect(commentsXml).toContain("Line two");
    expect(commentsXml).toContain("Line three");
    // Count w:p elements in the comment (should be 3)
    const pCount = (commentsXml.match(/<w:p>/g) || []).length;
    expect(pCount).toBe(3);
  });

  it("adds a comment on text inside a table cell", async () => {
    const p = await createDocWithNestedTable();
    const result = await addComment(p, "outer cell", "Table comment", "Reviewer");
    expect(result).toContain("Added comment");

    const commentsXml = await readRawCommentsXml(p);
    expect(commentsXml).toContain("Table comment");
    expect(commentsXml).toContain("Reviewer");

    const docXml = await readRawDocXml(p);
    expect(docXml).toContain("w:commentRangeStart");
    expect(docXml).toContain("w:commentRangeEnd");
  });

  it("adds a comment on text inside a nested table", async () => {
    const p = await createDocWithNestedTable();
    const result = await addComment(p, "nested text", "Nested comment", "Author");
    expect(result).toContain("Added comment");

    const commentsXml = await readRawCommentsXml(p);
    expect(commentsXml).toContain("Nested comment");
  });
});

// =========================================================================
// addComment — fuzzy matching
// =========================================================================

describe("addComment fuzzy matching", () => {
  it("exact match still works and does not report fuzzy", async () => {
    const p = await createTmpDoc("Exact match target");
    const result = await addComment(p, "match target", "Exact comment");
    expect(result).toContain("Added comment");
    expect(result).not.toContain("fuzzy-matched");
  });

  it("matches with whitespace normalization", async () => {
    // Document has double spaces; anchor has single space
    const p = await createTmpDoc("Text  with   extra  spaces");
    const result = await addComment(p, "with extra spaces", "Whitespace comment");
    expect(result).toContain("fuzzy-matched");

    const commentsXml = await readRawCommentsXml(p);
    expect(commentsXml).toContain("Whitespace comment");
  });

  it("matches full-width digits to half-width (１０ → 10)", async () => {
    const p = await createTmpDoc("増加率は１０倍になった");
    const result = await addComment(p, "10倍", "Full-width match");
    expect(result).toContain("fuzzy-matched");

    const commentsXml = await readRawCommentsXml(p);
    expect(commentsXml).toContain("Full-width match");
  });

  it("matches case-insensitively", async () => {
    const p = await createTmpDoc("The Quick Brown Fox");
    const result = await addComment(p, "quick brown fox", "Case comment");
    expect(result).toContain("fuzzy-matched");

    const commentsXml = await readRawCommentsXml(p);
    expect(commentsXml).toContain("Case comment");
  });

  it("fuzzy matches text inside a table cell", async () => {
    const p = await createDocWithNestedTable();
    // Document has "outer cell", search with different case
    const result = await addComment(p, "OUTER CELL", "Fuzzy table comment");
    expect(result).toContain("fuzzy-matched");

    const commentsXml = await readRawCommentsXml(p);
    expect(commentsXml).toContain("Fuzzy table comment");
  });

  it("still throws ANCHOR_NOT_FOUND when fuzzy match also fails", async () => {
    const p = await createTmpDoc("Hello world");
    await expect(
      addComment(p, "completely different text", "No match"),
    ).rejects.toMatchObject({ code: "ANCHOR_NOT_FOUND" });
  });

  it("matches full-width ASCII letters (Ａ → A)", async () => {
    const p = await createTmpDoc("ＡＢＣのテスト");
    const result = await addComment(p, "ABCのテスト", "Full-width letter match");
    expect(result).toContain("fuzzy-matched");

    const commentsXml = await readRawCommentsXml(p);
    expect(commentsXml).toContain("Full-width letter match");
  });
});

// =========================================================================
// addComments
// =========================================================================

describe("addComments", () => {
  it("adds a single comment in batch", async () => {
    const p = await createTmpDoc("Batch test content");
    const result = await addComments(p, [
      { anchor_text: "test content", comment_text: "Single batch" },
    ]);
    expect(result).toContain("1 added, 0 failed");

    const commentsXml = await readRawCommentsXml(p);
    expect(commentsXml).toContain("Single batch");
  });

  it("adds multiple comments to different paragraphs", async () => {
    const p = await createTmpDoc("First paragraph\nSecond paragraph\nThird paragraph");
    const result = await addComments(p, [
      { anchor_text: "First paragraph", comment_text: "Comment on first" },
      { anchor_text: "Second paragraph", comment_text: "Comment on second" },
      { anchor_text: "Third paragraph", comment_text: "Comment on third" },
    ]);
    expect(result).toContain("3 added, 0 failed");

    const comments = await readComments(p);
    expect(comments).toContain("Comment on first");
    expect(comments).toContain("Comment on second");
    expect(comments).toContain("Comment on third");
  });

  it("handles partial failure (some anchors not found)", async () => {
    const p = await createTmpDoc("Only this exists");
    const result = await addComments(p, [
      { anchor_text: "this exists", comment_text: "Found it" },
      { anchor_text: "nonexistent text", comment_text: "Won't be added" },
    ]);
    expect(result).toContain("1 added, 1 failed");
    expect(result).toContain("[OK]");
    expect(result).toContain("[FAIL]");

    const commentsXml = await readRawCommentsXml(p);
    expect(commentsXml).toContain("Found it");
    expect(commentsXml).not.toContain("Won't be added");
  });

  it("throws on empty comments array", async () => {
    const p = await createTmpDoc("Some text");
    await expect(addComments(p, [])).rejects.toMatchObject({
      code: "INVALID_PARAMETER",
    });
  });

  it("uses per-comment author when provided", async () => {
    const p = await createTmpDoc("Author test text");
    const result = await addComments(
      p,
      [{ anchor_text: "Author test", comment_text: "Custom author", author: "Alice" }],
      "DefaultAuthor",
    );
    expect(result).toContain("1 added");

    const commentsXml = await readRawCommentsXml(p);
    expect(commentsXml).toContain("Alice");
    expect(commentsXml).not.toContain("DefaultAuthor");
  });

  it("uses default author when per-comment author is not provided", async () => {
    const p = await createTmpDoc("Default author batch");
    await addComments(
      p,
      [{ anchor_text: "author batch", comment_text: "Default author test" }],
      "BatchReviewer",
    );

    const commentsXml = await readRawCommentsXml(p);
    expect(commentsXml).toContain("BatchReviewer");
  });

  it("uses fuzzy matching in batch mode", async () => {
    const p = await createTmpDoc("価格は１０００円です");
    const result = await addComments(p, [
      { anchor_text: "1000円", comment_text: "Fuzzy batch" },
    ]);
    expect(result).toContain("1 added");
    expect(result).toContain("fuzzy-matched");
  });
});

// =========================================================================
// readComments
// =========================================================================

describe("readComments", () => {
  it("reads comments from a document", async () => {
    const p = await createTmpDoc("Read my comments");
    await addComment(p, "comments", "First comment", "Alice");
    const result = await readComments(p);
    expect(result).toContain("1");
    expect(result).toContain("Alice");
    expect(result).toContain("First comment");
  });

  it("returns no-comments message for document without comments", async () => {
    const p = await createTmpDoc("No comments here");
    const result = await readComments(p);
    expect(result).toContain("No comments");
  });

  it("reads multiple comments with authors and dates", async () => {
    const p = await createTmpDoc("Text one and text two");
    await addComment(p, "Text one", "Comment A", "Author1");
    await addComment(p, "text two", "Comment B", "Author2");
    const result = await readComments(p);
    expect(result).toContain("Author1");
    expect(result).toContain("Author2");
    expect(result).toContain("Comment A");
    expect(result).toContain("Comment B");
  });

  it("shows threaded replies indented under parent", async () => {
    const p = await createTmpDoc("Threaded test content");
    await addComment(p, "test content", "Parent comment", "Alice");
    await replyToComment(p, 0, "Reply to parent", "Bob");

    const result = await readComments(p);
    expect(result).toContain("Parent comment");
    expect(result).toContain("Reply to parent");
    expect(result).toContain("Alice");
    expect(result).toContain("Bob");
  });

  it("falls back to flat display without commentsExtended.xml", async () => {
    const p = await createTmpDoc("Flat display test");
    await addComment(p, "display test", "Comment A", "Author1");
    await addComment(p, "Flat", "Comment B", "Author2");

    // No replies, so no commentsExtended.xml — all displayed flat
    const result = await readComments(p);
    expect(result).toContain("Comment A");
    expect(result).toContain("Comment B");
  });
});

// =========================================================================
// replyToComment
// =========================================================================

describe("replyToComment", () => {
  it("adds a basic reply to a comment", async () => {
    const p = await createTmpDoc("Reply test content");
    await addComment(p, "test content", "Original comment", "Alice");
    const result = await replyToComment(p, 0, "My reply", "Bob");

    expect(result).toContain("Added reply");
    expect(result).toContain("comment 0");
    expect(result).toContain("Bob");

    const commentsXml = await readRawCommentsXml(p);
    expect(commentsXml).toContain("Original comment");
    expect(commentsXml).toContain("My reply");
  });

  it("throws error when replying to nonexistent comment", async () => {
    const p = await createTmpDoc("No comments here");
    await expect(
      replyToComment(p, 999, "Reply to nothing"),
    ).rejects.toMatchObject({ code: "INVALID_PARAMETER" });
  });

  it("supports reply chains (reply to a reply)", async () => {
    const p = await createTmpDoc("Chain test content");
    await addComment(p, "test content", "Root comment", "Alice");
    await replyToComment(p, 0, "First reply", "Bob");
    await replyToComment(p, 1, "Reply to reply", "Charlie");

    const result = await readComments(p);
    expect(result).toContain("Root comment");
    expect(result).toContain("First reply");
    expect(result).toContain("Reply to reply");
  });

  it("creates commentsExtended.xml infrastructure", async () => {
    const p = await createTmpDoc("Infrastructure test");
    await addComment(p, "test", "Parent", "Author");
    await replyToComment(p, 0, "Reply", "Author");

    const ct = await readRawContentTypes(p);
    expect(ct).toContain("commentsExtended.xml");

    const rels = await readRawDocRels(p);
    expect(rels).toContain("commentsExtended.xml");
  });

  it("uses default author Claude", async () => {
    const p = await createTmpDoc("Default author reply");
    await addComment(p, "author reply", "Parent comment", "Alice");
    await replyToComment(p, 0, "Default author test");

    const commentsXml = await readRawCommentsXml(p);
    // Reply should use default "Claude" author
    expect(commentsXml).toContain("Claude");
  });
});

// =========================================================================
// deleteComment
// =========================================================================

describe("deleteComment", () => {
  it("removes comment from comments.xml and document markers", async () => {
    const p = await createTmpDoc("Delete this comment");
    await addComment(p, "this comment", "Will be deleted", "User");

    // Verify comment exists
    let comments = await readComments(p);
    expect(comments).toContain("Will be deleted");

    // Delete it
    await deleteComment(p, 0);

    // Verify comment is gone from comments.xml
    const commentsXml = await readRawCommentsXml(p);
    expect(commentsXml).not.toContain("Will be deleted");

    // Verify markers are removed from document.xml
    const xml = await readRawDocXml(p);
    expect(xml).not.toContain("w:commentRangeStart");
    expect(xml).not.toContain("w:commentRangeEnd");
    expect(xml).not.toContain("w:commentReference");
  });

  it("deletes only the specified comment, leaving others", async () => {
    const p = await createTmpDoc("First part and second part");
    await addComment(p, "First part", "Keep me", "A");
    await addComment(p, "second part", "Delete me", "B");

    // Delete the second comment (ID 1)
    await deleteComment(p, 1);

    const result = await readComments(p);
    expect(result).toContain("Keep me");
    expect(result).not.toContain("Delete me");
  });
});

// =========================================================================
// Entity handling in comments
// =========================================================================

describe("comment entity handling", () => {
  it("preserves special characters in comment text", async () => {
    const p = await createTmpDoc("Hello world");
    await addComment(p, "Hello", "Use AT&T <network> for \"enterprise\"");
    const result = await readComments(p);
    expect(result).toContain("AT&T <network>");
  });

  it("preserves special characters in reply text", async () => {
    const p = await createTmpDoc("Hello world");
    await addComment(p, "Hello", "First comment");
    await replyToComment(p, 0, "Reply with & and <angle>");
    const result = await readComments(p);
    expect(result).toContain("Reply with & and <angle>");
  });
});
